(() => {
  "use strict";

  const PAGE_SIZE = 50;

  // 실제 raw 컬럼명이 플랫폼/광고상품마다 다르므로, 리포트 필드 <-> raw 컬럼명 매핑표.
  // 값은 문자열 하나 또는 후보 문자열 배열(여러 광고상품이 서로 다른 이름을 쓸 때
  // 앞에서부터 순서대로 시도). null이면 아래 FIELD_AUTO_PATTERNS로 자동 추정하고,
  // NOT_MAPPED이면 추정도 하지 않고 "-"로 비워둔다.
  //
  // 쿠팡: PA(키워드), BPA(브랜드 광고그룹), NCA(성과형) raw 샘플을 직접 확인해서 채움.
  //   같은 개념이라도 상품별로 "광고비" / "집행 광고비" / "광고비(원)"처럼 표기가 다르고,
  //   NCA는 "총 전환 매출액 (1일)(원)"처럼 띄어쓰기와 단위가 더 붙는 식이라
  //   resolveField에서 공백을 지우고 한 번 더 비교(접두 일치)한다.
  // 네이버: 쇼핑검색광고(카탈로그)/애드부스트/카탈로그 raw 샘플을 직접 확인해서 채움.
  //   쇼핑검색광고(카탈로그)는 전환/매출 컬럼 자체가 없는 리포트라 그 파일에서는
  //   전환수/전환매출이 자동으로 "-"가 된다(정상 동작).
  const NOT_MAPPED = Symbol("not-mapped");
  const FIELD_MAP = {
    coupang: {
      date: "날짜",
      campaign: ["캠페인명", "캠페인 이름"],
      group: "광고그룹",
      cost: ["광고비", "집행 광고비", "광고비(원)"],
      impressions: "노출수",
      clicks: "클릭수",
      conversions: "직접 판매수량(1일)",
      revenue: "직접 전환매출액(1일)",
      // INDEX/ 품목 분류 파일과 매칭하기 위한 키. NCA는 둘 다 없어서 항상 "미분류"가 된다.
      campaignId: "캠페인 ID",
      viid: "광고집행 옵션ID",
      // 주간RAW 다운로드(CE_RAW/MX_RAW 양식)용 그대로 옮기는 원본 컬럼들.
      // "총 전환매출액(14일)" 같은 후보 하나만 줘도, 공백을 지운 뒤 접두 일치로
      // NCA의 "총 전환 매출액 (14일)(원)" 같은 표기도 자동으로 잡힌다.
      execProductName: "광고집행 상품명",
      keyword: "키워드",
      // NCA 세부품목 보정(의류케어 기획전 캠페인의 세탁기/건조기 구분)용.
      adName: "광고명",
      totalConv1d: "총 판매수량(1일)",
      totalRevenue1d: "총 전환매출액(1일)",
      totalConv14d: "총 판매수량(14일)",
      directConv14d: "직접 판매수량(14일)",
      totalRevenue14d: "총 전환매출액(14일)",
      directRevenue14d: "직접 전환매출액(14일)",
    },
    naver: {
      date: ["기간", "일별"],
      campaign: ["캠페인 이름", "캠페인"],
      group: ["광고 그룹 이름", "광고그룹", "상품명"],
      cost: "총비용",
      impressions: "노출수",
      clicks: "클릭수",
      conversions: "구매완료 수",
      revenue: "구매완료 전환매출액",
      // GFA리포트RAW 다운로드(네이버_GFA_RAW_양식.xlsb)용. 쇼핑검색광고(카탈로그)는
      // 캠페인ID/상품ID 자체가 없어서 둘 다 안 잡히고, 카탈로그는 캠페인ID만 있다.
      campaignId: "캠페인 ID",
      shoppingProductId: "쇼핑몰 상품 ID",
      // 쇼핑검색광고(카탈로그형) RAW 다운로드용. PC/모바일 매체·시간대별·
      // 평균노출순위는 쇼핑검색광고(카탈로그) 일별 리포트에만 있고, 검색어는
      // "쇼핑검색_검색어_RAW" 리포트에만 있어서 이 컬럼의 존재 여부로 두
      // 리포트를 구분한다.
      pcMobile: "PC/모바일 매체",
      hourly: "시간대별",
      avgRank: "평균노출순위",
      keyword: "검색어",
    },
  };

  // 명시적 매핑이 하나도 안 맞을 때(새로운 광고상품 등) 컬럼명을 추정하는 패턴.
  // "클릭률"이 "클릭수"로, "캠페인 ID"가 캠페인명으로 오인되지 않도록 주의해서 만듦.
  const FIELD_AUTO_PATTERNS = {
    date: /^(날짜|일자|기간|일별)$/i,
    campaign: /^캠페인(명|\s?이름)?$/i,
    group: /그룹(명|\s?이름)?$/i,
    cost: /광고비|총비용|^비용$|spend|cost/i,
    impressions: /노출수|^노출$|impression/i,
    clicks: /^클릭수$|^클릭$/i,
  };

  function normalizeColName(s) {
    return s.replace(/\s+/g, "");
  }

  // 표4(우수 판매 TOP10)에서 "상품" 하나를 무엇으로 볼지 - 상품구분별로 다르다.
  // "__campaign"/"__group"은 이미 계산해둔 캠페인명/그룹명을 그대로 재사용하라는 뜻.
  const PRODUCT_LABEL_FIELD = {
    coupang: {
      PA: "광고전환매출발생 상품명",
      BPA: "__campaign",
      // NCA는 캠페인명이 "[NCA] 10월 브랜드위크_모바일(상품소재)"처럼 여러
      // 상품을 묶는 기획전 이름이라 상품을 구분 못 한다. "광고명"은 실제
      // 광고 소재 단위라 거기 들어간 상품명으로 구분한다.
      NCA: "광고명",
    },
    naver: {
      "쇼핑검색광고(카탈로그)": "__campaign",
      애드부스트: "상품명",
      카탈로그: "__group",
    },
  };

  function resolveProductLabel(row, filePath, platform, productType, campaign, group) {
    const spec = PRODUCT_LABEL_FIELD[platform]?.[productType];
    if (!spec || spec === "__campaign") return campaign;
    if (spec === "__group") return group;
    const cols = state.colsByFile.get(filePath);
    if (cols && cols.has(spec)) return row[spec] ?? campaign;
    return campaign;
  }

  const METRIC_COLUMNS = [
    { key: "cost", label: "광고비", format: "money" },
    { key: "impressions", label: "노출", format: "count" },
    { key: "clicks", label: "클릭", format: "count" },
    { key: "ctr", label: "CTR", format: "pct" },
    { key: "cpc", label: "CPC", format: "money" },
    { key: "conversions", label: "전환수", format: "count" },
    { key: "revenue", label: "전환매출", format: "money" },
    { key: "cvr", label: "CVR", format: "pct" },
    { key: "aov", label: "객단가", format: "money" },
    { key: "roas", label: "ROAS", format: "pct" },
  ];

  // 표1: 쿠팡은 INDEX/ 품목 분류 파일로 캠페인&그룹, 상품구분 대신 품목
  // 기준으로만 묶는다(상품구분이 달라도 같은 품목이면 하나로 합산).
  // 네이버는 품목 분류 데이터가 없으니 기존처럼 상품구분+캠페인명/그룹명 기준.
  const TABLE1_COLS_COUPANG = [{ key: "itemCategory", label: "품목" }, ...METRIC_COLUMNS];
  const TABLE1_COLS_NAVER = [
    { key: "productType", label: "상품구분" },
    { key: "campaign", label: "캠페인명" },
    ...METRIC_COLUMNS,
  ];
  function table1ColsFor(platform) {
    return platform === "coupang" ? TABLE1_COLS_COUPANG : TABLE1_COLS_NAVER;
  }
  const TABLE2_COLS = [{ key: "productType", label: "상품구분" }, ...METRIC_COLUMNS];
  const TABLE3_COLS = [
    { key: "weekLabel", label: "주차" },
    ...METRIC_COLUMNS,
  ];
  const TABLE4_COLS = [
    { key: "rank", label: "순위" },
    { key: "productType", label: "상품구분" },
    { key: "productName", label: "상품명" },
    ...METRIC_COLUMNS,
  ];

  const state = {
    manifest: { files: [] },
    selectedPaths: new Set(),
    // 표1~4/슬라이서가 쓰는 행. 쿠팡 PA 키워드 리포트처럼 "키워드" 컬럼이
    // 있는 파일은 캠페인당 키워드별로 수십만 행까지 쪼개져 있는데, 표1~4는
    // 품목/상품구분/캠페인+상품 단위로만 묶어서 보여줘서 키워드 단위가
    // 전혀 필요 없다. 그래서 그런 파일은 읽은 직후 바로 키워드 단위를
    // 합쳐서(fetchAllDataRows/aggregateKeywordRows) 원본 키워드별 행 자체를
    // 메모리에 남기지 않는다 - 모바일처럼 메모리가 적은 기기에서 수십만
    // 행짜리 원본을 계속 들고 있으면 무한로딩/탭 다운으로 이어질 수
    // 있어서다. "리포트RAW 다운받기"처럼 키워드 단위가 꼭 필요한 다운로드는
    // 클릭했을 때 그 파일들만 따로 다시 받아서(fetchFullCoupangRows)
    // 처리한다.
    rows: [],
    colsByFile: new Map(), // filePath -> Set(원본 컬럼명)
    indexMaps: null, // INDEX/ 품목 분류 파일에서 만든 조회용 맵 (쿠팡 전용)
    gfaIndexMaps: null, // 리포트양식/ GFA 양식 파일의 품목 조회용 맵 (네이버 전용)
    activePlatform: "coupang",
    table1Page: 1,
    // 슬라이서 선택 상태 - 플랫폼(탭)마다 따로 기억한다(탭을 바꿔도 유지됨).
    filters: {
      coupang: { years: new Set(), months: new Set(), productTypes: new Set(), itemCategories: new Set() },
      naver: { years: new Set(), months: new Set(), productTypes: new Set(), itemCategories: new Set() },
    },
  };

  const el = (id) => document.getElementById(id);

  // ---------- 인코딩/파싱 유틸 ----------

  async function fetchArrayBuffer(url) {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`${url} 요청 실패 (${res.status})`);
    return res.arrayBuffer();
  }

  function decodeCsvBuffer(buf) {
    // 쿠팡/네이버 raw csv는 UTF-8이 아닌 EUC-KR(CP949)로 내려받아지는 경우가 많다.
    // UTF-8로 먼저 시도하고, 깨지면(fatal) EUC-KR로 다시 디코딩한다.
    try {
      return new TextDecoder("utf-8", { fatal: true }).decode(buf);
    } catch (e) {
      return new TextDecoder("euc-kr").decode(buf);
    }
  }

  function trimKeys(row) {
    const out = {};
    for (const k of Object.keys(row)) out[k.trim()] = row[k];
    return out;
  }

  // csv/xlsx 파일 하나를 읽어서 {header: value} 형태의 행 배열로 반환한다.
  async function parseDataFile(path) {
    const buf = await fetchArrayBuffer(path);
    if (/\.(xlsx?|xlsb)$/i.test(path)) {
      // 엑셀은 날짜 셀이 내부적으로 숫자(일련번호, 예: 46266)로 저장된다.
      // cellDates로 최대한 JS Date로 바꾸되, 셀 서식이 인식 안 되는 경우를 대비해
      // raw:true로 원본 값(숫자 or Date)을 그대로 받아 parseDateLoose에서 이중으로 처리한다.
      const workbook = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { raw: true, defval: "" });
      return rows.map((row) => trimKeys(normalizeXlsxRow(row)));
    }
    const text = stripLeadingTitleLine(decodeCsvBuffer(buf).replace(/^﻿/, ""));
    const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
    return parsed.data.map(trimKeys);
  }

  function splitCsvFields(line) {
    const fields = [];
    let cur = "";
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') inQuotes = !inQuotes;
      else if (ch === "," && !inQuotes) {
        fields.push(cur);
        cur = "";
      } else {
        cur += ch;
      }
    }
    fields.push(cur);
    return fields;
  }

  // 실제 값이 들어있는 필드 수만 센다(따옴표 안 콤마는 필드를 안 나눔).
  // "제목,ID" 뒤에 열 맞추기용 빈 콤마가 더 붙는 리포트(예: 검색어_RAW)가
  // 있어서, 그냥 필드 개수가 아니라 "비어있지 않은" 필드 개수로 비교해야
  // 제목 줄과 진짜 헤더 줄을 구분할 수 있다.
  function countNonEmptyCsvFields(line) {
    return splitCsvFields(line).filter((f) => f.trim() !== "").length;
  }

  // 네이버 리포트 중 일부(예: 쇼핑검색광고(카탈로그) 다운로드)는 실제 헤더 위에
  // "리포트 제목,ID" 같은 한 줄짜리 메타데이터 행이 붙어 있어서, 그걸 헤더로
  // 잘못 읽으면 노출수/클릭수/총비용은 물론 캠페인/그룹/날짜까지 다 깨진다.
  // 첫 줄의 (비어있지 않은) 필드 수가 둘째 줄보다 훨씬 적으면 첫 줄을
  // 제목으로 보고 건너뛴다.
  function stripLeadingTitleLine(text) {
    const nl = text.indexOf("\n");
    if (nl === -1) return text;
    const firstLine = text.slice(0, nl).replace(/\r$/, "");
    const nl2 = text.indexOf("\n", nl + 1);
    const secondLine = (nl2 === -1 ? text.slice(nl + 1) : text.slice(nl + 1, nl2)).replace(/\r$/, "");
    const firstCount = countNonEmptyCsvFields(firstLine);
    const secondCount = countNonEmptyCsvFields(secondLine);
    if (firstCount <= 2 && secondCount >= firstCount + 3) {
      return text.slice(nl + 1);
    }
    return text;
  }

  // sheet_to_json이 돌려준 JS Date 객체를 문자열로 바꿔서, 이후 파싱 로직이
  // csv에서 온 문자열 값과 동일하게 다룰 수 있게 한다. 숫자/문자열은 그대로 둔다.
  function normalizeXlsxRow(row) {
    const out = {};
    for (const k of Object.keys(row)) {
      const v = row[k];
      out[k] = v instanceof Date ? formatDateUTC(v) : v;
    }
    return out;
  }

  function formatDateUTC(d) {
    const y = d.getUTCFullYear();
    const mo = String(d.getUTCMonth() + 1).padStart(2, "0");
    const da = String(d.getUTCDate()).padStart(2, "0");
    return `${y}-${mo}-${da}`;
  }

  // 엑셀 날짜 일련번호(1899-12-30을 0으로 보는 방식) -> YYYY-MM-DD 문자열.
  function excelSerialToDateStr(serial) {
    const ms = Math.round((serial - 25569) * 86400 * 1000);
    return formatDateUTC(new Date(ms));
  }

  // 캠페인ID/VIID 조회 키를 만들기 위한 정규화. 엑셀에서 숫자로 읽히면(92628741335.0
  // 처럼 float) 소수부를 떼고, 문자열이면 앞뒤 공백만 지운다.
  function normalizeId(v) {
    if (v == null) return "";
    if (typeof v === "number") {
      return Number.isFinite(v) ? String(Math.round(v)) : "";
    }
    const s = String(v).trim();
    if (/^\d+(\.0+)?$/.test(s)) return String(Math.round(Number(s)));
    return s;
  }

  function parseNumberLoose(v) {
    if (v == null) return NaN;
    if (typeof v === "number") return Number.isFinite(v) ? v : NaN;
    const s = String(v).trim().replace(/[,₩%원\s]/g, "");
    if (s === "" || s === "-") return NaN;
    const n = Number(s);
    return Number.isFinite(n) ? n : NaN;
  }

  function parseDateLoose(v) {
    if (v == null) return null;
    if (typeof v === "number") {
      // 쿠팡 상품별로 날짜가 YYYYMMDD 형태의 숫자(예: 20260914)로 오거나,
      // cellDates가 못 알아본 엑셀 날짜 일련번호(예: 46266)로 그대로 남기도 한다.
      if (Number.isInteger(v) && v >= 19000101 && v <= 21001231) {
        v = String(v);
      } else if (v > 20000 && v < 80000) {
        return excelSerialToDateStr(v);
      } else {
        return null;
      }
    }
    let s = String(v).trim();
    if (!s) return null;
    // 문자열이지만 순수 숫자(엑셀이 raw:true에서도 텍스트로 내보낸 일련번호)인 경우
    if (/^\d{4,6}(\.\d+)?$/.test(s)) {
      const n = Number(s);
      if (n > 20000 && n < 80000) return excelSerialToDateStr(n);
    }
    if (/^\d{8}$/.test(s)) {
      s = `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}`;
    } else {
      s = s.replace(/[./]/g, "-");
    }
    const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!m) return null;
    const mm = m[2].padStart(2, "0");
    const dd = m[3].padStart(2, "0");
    if (+mm < 1 || +mm > 12 || +dd < 1 || +dd > 31) return null;
    return `${m[1]}-${mm}-${dd}`;
  }

  // "YYYY-MM-DD" -> 그 날짜가 속한 주(월요일~일요일)의 정렬용 키와 표시용 라벨.
  function weekOf(dateStr) {
    const d = new Date(`${dateStr}T00:00:00Z`);
    const day = d.getUTCDay(); // 0=일 ~ 6=토
    const diffToMonday = day === 0 ? -6 : 1 - day;
    const monday = new Date(d);
    monday.setUTCDate(d.getUTCDate() + diffToMonday);
    const sunday = new Date(monday);
    sunday.setUTCDate(monday.getUTCDate() + 6);
    const fmt = (x) => `${String(x.getUTCMonth() + 1).padStart(2, "0")}.${String(x.getUTCDate()).padStart(2, "0")}`;
    return { key: formatDateUTC(monday), label: `${fmt(monday)}~${fmt(sunday)}` };
  }

  function yearOf(dateStr) {
    return dateStr.slice(0, 4);
  }

  function monthOf(dateStr) {
    return dateStr.slice(5, 7);
  }

  // "쇼핑검색_검색어_RAW" 리포트처럼 행에 날짜 컬럼이 없고 파일명에만 날짜
  // 범위가 있는 경우("..._261027~261027.csv", "..._20260914~20260916.csv")
  // 쓰는 보조 함수. 범위의 시작일을 대표 날짜로 쓴다.
  function parseDateFromFileName(fileName) {
    const m = String(fileName || "").match(/(\d{6,8})\s*[~-]/);
    if (!m) return null;
    let s = m[1];
    if (s.length === 6) s = `20${s}`;
    return parseDateLoose(s);
  }

  // "YYYY-MM-DD" -> CE/MX 주간RAW 양식이 쓰는 "N월 M주차" (그 달 1일부터
  // 7일 단위로 끊은 주차. 주간RAW 표1~4의 월요일~일요일 주차(weekOf)와는
  // 다른 기준이라 별도로 둔다).
  function weekOfMonthLabel(dateStr) {
    const month = Number(monthOf(dateStr));
    const day = Number(dateStr.slice(8, 10));
    return `${month}월 ${Math.ceil(day / 7)}주차`;
  }

  // 년도/월/광고상품유형/품목 슬라이서 선택 상태에 이 행(날짜+상품구분+품목)이
  // 걸리는지. 어떤 슬라이서도 선택 안 했으면(빈 Set) 그 기준은 통과시킨다
  // (=전체 보기). 품목은 쿠팡 전용이라 네이버에서는 itemCategory가 항상
  // null이고, 네이버 쪽 itemCategories 슬라이서도 항상 비어있어 걸리지 않는다.
  function passesFilters(platform, date, productType, itemCategory) {
    const f = state.filters[platform];
    if (f.years.size && !f.years.has(yearOf(date))) return false;
    if (f.months.size && !f.months.has(monthOf(date))) return false;
    if (f.productTypes.size && !f.productTypes.has(productType)) return false;
    if (f.itemCategories.size && !f.itemCategories.has(itemCategory)) return false;
    return true;
  }

  // 현재 플랫폼에 실제로 존재하는 년도/월/상품구분/품목 목록 (슬라이서 버튼
  // 목록用). 지금 선택된 필터와 무관하게 전체 목록을 보여준다.
  function computeFilterOptions(platform) {
    const years = new Set();
    const months = new Set();
    const productTypes = new Set();
    const itemCategories = new Set();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const date = row.__date;
      if (date) {
        years.add(yearOf(date));
        months.add(monthOf(date));
      }
      productTypes.add(row.__productType);
      if (platform === "coupang") {
        itemCategories.add(row.__item);
      }
    }
    return {
      years: [...years].sort(),
      months: [...months].sort(),
      productTypes: [...productTypes].sort(),
      itemCategories: [...itemCategories].sort((a, b) => a.localeCompare(b, "ko")),
    };
  }

  function formatMetric(value, format) {
    if (value == null || !Number.isFinite(value)) return "-";
    if (format === "pct") return `${value.toLocaleString("ko-KR", { maximumFractionDigits: 2 })}%`;
    if (format === "money") return `${Math.round(value).toLocaleString("ko-KR")}원`;
    return value.toLocaleString("ko-KR", { maximumFractionDigits: 0 });
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;",
    }[c]));
  }

  // ---------- 컬럼 매핑 ----------

  // resolveField 결과는 filePath+field 조합에만 좌우되고(행 내용과 무관),
  // raw가 수십만 행이면 행마다 다시 계산하는 게 비싸서 캐시해둔다.
  // fetchAllDataRows()/fetchFullCoupangRows()가 state.colsByFile을 다시
  // 만들 때 함께 비운다.
  let fieldResolutionCache = new Map();

  // 쿠팡 raw 파일을 읽는 즉시(loadAndTagFile 안에서) 품목을 분류하는데,
  // 그 분류는 INDEX 조회표(state.indexMaps)가 준비돼야 한다. raw 파일
  // fetch는 INDEX 조회표 fetch와 병렬로 시작하므로, 분류하기 직전에 이
  // promise를 기다려서 INDEX가 아직 안 끝났을 때 생기는 미분류 오류를
  // 막는다. loadManifest()가 매번 새로 할당한다.
  let indexMapsReadyPromise = null;

  // platform의 field(날짜/캠페인/그룹/광고비 등)가 filePath 안에서 어떤 실제 컬럼명인지 찾는다.
  function resolveField(filePath, platform, field) {
    const cacheKey = `${filePath}\u0001${field}`;
    if (fieldResolutionCache.has(cacheKey)) return fieldResolutionCache.get(cacheKey);
    const result = resolveFieldUncached(filePath, platform, field);
    fieldResolutionCache.set(cacheKey, result);
    return result;
  }

  function resolveFieldUncached(filePath, platform, field) {
    const cols = state.colsByFile.get(filePath);
    if (!cols) return null;
    return resolveFieldForCols(cols, platform, field);
  }

  // cols(Set)만 있으면 되는 순수 버전. state.colsByFile이 아직 안 만들어진
  // 시점(파일 하나를 막 읽어서 키워드 단위를 즉시 합칠 때 등)에도 쓴다.
  function resolveFieldForCols(cols, platform, field) {
    const mapped = FIELD_MAP[platform]?.[field];
    if (mapped === NOT_MAPPED) return null;
    if (mapped) {
      const candidates = Array.isArray(mapped) ? mapped : [mapped];
      // 1) 정확히 일치하는 컬럼명부터 시도
      for (const c of candidates) if (cols.has(c)) return c;
      // 2) 광고상품마다 붙는 공백/단위가 달라서("총 전환 매출액 (1일)(원)" 등)
      //    공백을 지우고 다시 비교 (완전 일치 또는 접두 일치)
      for (const col of cols) {
        const normCol = normalizeColName(col);
        for (const cand of candidates) {
          const normCand = normalizeColName(cand);
          if (normCol === normCand || normCol.startsWith(normCand)) return col;
        }
      }
    }
    const pattern = FIELD_AUTO_PATTERNS[field];
    if (pattern) {
      for (const c of cols) if (pattern.test(c)) return c;
    }
    return null;
  }

  // 파일명으로 광고상품 구분(쿠팡 PA/BPA/NCA, 네이버 쇼핑검색/애드부스트/카탈로그)을 추정한다.
  // 서로 다른 상품인데 캠페인명/그룹명이 우연히 같아도 집계가 섞이지 않도록
  // 이 값을 집계 키에 포함시킨다.
  function detectProductType(platform, fileName) {
    if (platform === "coupang") {
      if (/BPA/i.test(fileName)) return "BPA";
      if (/NCA/i.test(fileName)) return "NCA";
      if (/PA/i.test(fileName)) return "PA";
      return "기타";
    }
    if (platform === "naver") {
      if (fileName.includes("쇼핑검색")) return "쇼핑검색광고(카탈로그)";
      if (fileName.includes("애드부스트")) return "애드부스트";
      if (fileName.includes("카탈로그")) return "카탈로그";
      return "기타";
    }
    return "기타";
  }

  // ---------- INDEX/ 품목 분류 ----------

  // INDEX 엑셀 한 시트에 캠페인ID/캠페인명/그룹명→품목(B~F열), VIID→품목(H~L열)
  // 두 개의 독립된 표가 나란히 들어있다(행이 서로 대응되는 게 아님).
  // 주의: 이 파일은 A열이 완전히 비어있어서, header:1 배열모드로 읽으면
  // 시트의 실제 사용범위(!ref)가 B열부터 시작해 배열 인덱스가 한 칸씩
  // 밀리는 문제가 있었다(그 결과 모든 행이 "미분류"로 나옴). 그래서
  // 배열 인덱스 대신 "C2", "E2"처럼 셀 주소를 직접 읽어서 이 문제를 피한다.
  function sheetCellText(sheet, colLetter, rowNum) {
    const cell = sheet[`${colLetter}${rowNum}`];
    if (!cell || cell.v == null) return "";
    return String(cell.v).trim();
  }

  // 캠페인ID/캠페인명/그룹명/VIID 각각을 단독 키로 조회하는 맵 4개를 만든다
  // (여러 조건을 동시에 만족해야 하는 조합 키는 쓰지 않는다). 값은 {item,subItem}.
  function buildIndexMaps(sheet) {
    const byCampaignId = new Map();
    const byCampaignName = new Map();
    const byGroupName = new Map();
    const byViid = new Map();
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    const lastRow = range.e.r + 1; // 1행이 헤더이므로 2행부터 lastRow까지 순회
    for (let r = 2; r <= lastRow; r++) {
      const item = sheetCellText(sheet, "E", r); // E열 품목
      if (item) {
        const campaignName = sheetCellText(sheet, "B", r); // B열 캠페인명
        const cid = normalizeId(sheetCellText(sheet, "C", r)); // C열 캠페인ID
        const groupName = sheetCellText(sheet, "D", r); // D열 그룹명
        const subItem = sheetCellText(sheet, "F", r) || item; // F열 세부품목
        const entry = { item, subItem };
        if (cid) byCampaignId.set(cid, entry);
        if (campaignName) byCampaignName.set(campaignName, entry);
        if (groupName) byGroupName.set(groupName, entry);
      }
      const vid = normalizeId(sheetCellText(sheet, "I", r)); // I열 VIID
      const rightItem = sheetCellText(sheet, "K", r); // K열 품목
      if (vid && rightItem) {
        const rightSubItem = sheetCellText(sheet, "L", r) || rightItem; // L열 세부품목
        byViid.set(vid, { item: rightItem, subItem: rightSubItem });
      }
    }
    return { byCampaignId, byCampaignName, byGroupName, byViid };
  }

  // localeCompare는 로캘에 따라 "_"와 "."의 순서를 다르게 취급해서
  // (예: "INDEX_260929_수정.xlsb"가 "INDEX_260929.xlsb"보다 최신인데도
  // 거꾸로 뽑히는 문제가 있었다), 코드 포인트 그대로 비교해서 파일명 기준
  // 최신 파일을 고른다.
  function pickLatestByName(files) {
    return [...files].sort((a, b) => (a.name > b.name ? -1 : a.name < b.name ? 1 : 0))[0];
  }

  async function loadIndexMaps() {
    const files = Array.isArray(state.manifest.indexFiles) ? state.manifest.indexFiles : [];
    if (files.length === 0) {
      state.indexMaps = null;
      return;
    }
    const latest = pickLatestByName(files);
    try {
      const buf = await fetchArrayBuffer(latest.path);
      const workbook = XLSX.read(new Uint8Array(buf), { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      state.indexMaps = buildIndexMaps(sheet);
    } catch (e) {
      console.error("품목 분류 파일을 불러오지 못했습니다:", latest.path, e);
      state.indexMaps = null;
    }
  }

  // 네이버 GFA 양식(리포트양식/ 폴더) 안에 있는 캠페인ID/상품코드 -> 품목구분
  // 조회표를 읽는다. GFA_캠페인RAW 시트 O~S열(월/SOP/캠페인종류/캠페인ID/품목구분),
  // GFA_상품RAW 시트 P~U열(월/SOP/캠페인종류/상품코드/품목구분/상품명)이 그
  // 조회표다. 같은 캠페인ID/상품코드라도 달이 바뀌면 다른 품목으로 재사용되는
  // 경우가 있어서(실제 양식에서 확인됨) 월까지 키에 포함한다.
  function parseLookupMonth(v) {
    const m = String(v ?? "").match(/\d+/);
    return m ? String(Number(m[0])) : "";
  }

  function buildGfaIndexMaps(campaignSheet, productSheet) {
    const byCampaign = new Map();
    const byProduct = new Map();
    if (campaignSheet) {
      const range = XLSX.utils.decode_range(campaignSheet["!ref"] || "A1:A1");
      for (let r = 2; r <= range.e.r + 1; r++) {
        const month = parseLookupMonth(sheetCellText(campaignSheet, "O", r));
        const campaignId = normalizeId(sheetCellText(campaignSheet, "R", r));
        const item = sheetCellText(campaignSheet, "S", r);
        if (month && campaignId && item) byCampaign.set(`${month}\u0001${campaignId}`, item);
      }
    }
    if (productSheet) {
      const range = XLSX.utils.decode_range(productSheet["!ref"] || "A1:A1");
      for (let r = 2; r <= range.e.r + 1; r++) {
        const month = parseLookupMonth(sheetCellText(productSheet, "P", r));
        const productCode = normalizeId(sheetCellText(productSheet, "S", r));
        const item = sheetCellText(productSheet, "T", r);
        const modelName = sheetCellText(productSheet, "U", r);
        if (month && productCode && item) byProduct.set(`${month}\u0001${productCode}`, { item, modelName });
      }
    }
    return { byCampaign, byProduct };
  }

  async function loadGfaIndexMaps() {
    const files = (Array.isArray(state.manifest.reportFormFiles) ? state.manifest.reportFormFiles : []).filter(
      (f) => f.name.includes("GFA")
    );
    if (files.length === 0) {
      state.gfaIndexMaps = null;
      return;
    }
    const latest = pickLatestByName(files);
    try {
      const buf = await fetchArrayBuffer(latest.path);
      const workbook = XLSX.read(new Uint8Array(buf), { type: "array" });
      state.gfaIndexMaps = buildGfaIndexMaps(
        workbook.Sheets["GFA_캠페인RAW"],
        workbook.Sheets["GFA_상품RAW"]
      );
    } catch (e) {
      console.error("GFA 품목 분류 파일을 불러오지 못했습니다:", latest.path, e);
      state.gfaIndexMaps = null;
    }
  }

  // 쿠팡 상품구분별로 품목을 조회할 때 시도할 순서. 각 단계는 단독 키로만
  // 조회하고(조합 키 아님), 앞 단계에서 못 찾으면 다음 단계로 넘어간다.
  const ITEM_CATEGORY_LOOKUP_ORDER = {
    PA: ["campaignId", "campaignName", "groupName", "viid"],
    BPA: ["campaignId", "groupName"],
    NCA: ["campaignName"],
    __default__: ["campaignId", "campaignName", "groupName", "viid"],
  };

  // NCA의 "[NCA] 10월 브랜드위크_의류케어(...)" 캠페인은 세탁기/건조기 광고를
  // 하나의 캠페인으로 묶어서 캠페인명만으로는 세부품목을 못 가른다. 실제
  // 광고 소재 단위인 "광고명"에 상품이 그대로 적혀 있어서(예: "...AI 콤보
  // 세탁기...", "...건조기 9kg...") 그 텍스트로 세탁기/건조기를 가른다.
  function refineNcaLaundrySubItem(adName) {
    const s = String(adName ?? "");
    if (!s) return null;
    if (/건조기/.test(s)) return "건조기";
    if (/콤보/.test(s)) return "세탁기";
    return null;
  }

  // 쿠팡 raw 한 행의 품목/세부품목을 찾는다. 상품구분(PA/BPA/NCA)에 따라 정해진
  // 순서로 캠페인ID/캠페인명/그룹명/VIID를 단독 키로 조회하고, 그래도 없으면
  // {item:"미분류", subItem:"미분류"}. 네이버는 품목 분류 데이터가 없으므로 null.
  function resolveItemClassification(row, fp, platform, productType) {
    if (platform !== "coupang") return null;
    const maps = state.indexMaps;
    const unclassified = { item: "미분류", subItem: "미분류" };
    const cidCol = resolveField(fp, platform, "campaignId");
    const campaignCol = resolveField(fp, platform, "campaign");
    const groupCol = resolveField(fp, platform, "group");
    const viidCol = resolveField(fp, platform, "viid");
    const values = {
      campaignId: cidCol ? normalizeId(row[cidCol]) : "",
      campaignName: campaignCol ? String(row[campaignCol] ?? "").trim() : "",
      groupName: groupCol ? String(row[groupCol] ?? "").trim() : "",
      viid: viidCol ? normalizeId(row[viidCol]) : "",
    };
    // 캠페인명에 "탭"/"태블릿"이 있으면 INDEX 조회와 상관없이 항상 태블릿으로
    // 분류한다(새 캠페인이 계속 생겨도 INDEX 파일을 매번 안 고쳐도 되게).
    if (/탭|태블릿/.test(values.campaignName)) {
      return { item: "태블릿", subItem: "태블릿" };
    }
    if (!maps) return unclassified;
    const mapFor = {
      campaignId: maps.byCampaignId,
      campaignName: maps.byCampaignName,
      groupName: maps.byGroupName,
      viid: maps.byViid,
    };
    const order = ITEM_CATEGORY_LOOKUP_ORDER[productType] || ITEM_CATEGORY_LOOKUP_ORDER.__default__;
    for (const field of order) {
      const v = values[field];
      if (v && mapFor[field].has(v)) {
        const result = mapFor[field].get(v);
        if (productType === "NCA" && result.item === "의류케어") {
          const adNameCol = resolveField(fp, platform, "adName");
          const refined = refineNcaLaundrySubItem(adNameCol ? row[adNameCol] : null);
          if (refined) return { item: result.item, subItem: refined };
        }
        return result;
      }
    }
    return unclassified;
  }

  // 품목 -> 사업부(CE/MX). 웨어러블/모바일은 MX, 그 외 알려진 품목은 CE,
  // "미분류"는 사업부를 특정할 수 없으므로 null.
  const MX_ITEMS = new Set(["웨어러블", "모바일"]);
  function classifyDivision(item) {
    if (!item || item === "미분류") return null;
    return MX_ITEMS.has(item) ? "MX" : "CE";
  }

  // ---------- 데이터 로드 ----------

  async function loadManifest() {
    try {
      const res = await fetch("data/manifest.json", { cache: "no-store" });
      state.manifest = res.ok ? await res.json() : { files: [] };
    } catch (e) {
      state.manifest = { files: [] };
    }
    if (!Array.isArray(state.manifest.files)) state.manifest.files = [];
    state.selectedPaths = new Set(state.manifest.files.map((f) => f.path));
    // INDEX 조회표 fetch와 raw 데이터 파일 fetch는 서로 의존 관계가 없어서
    // 순서대로 기다리지 않고 한꺼번에 병렬로 시작한다. 다만 각 raw 파일의
    // 품목 분류(loadAndTagFile 안에서 바로 계산)는 INDEX 조회표가 준비돼야
    // 하므로, indexMapsReadyPromise를 공유해서 "분류하는 순간"만 기다리게
    // 한다(파일 fetch 자체는 그대로 병렬로 진행됨).
    indexMapsReadyPromise = loadIndexMaps();
    const [, , allRows] = await Promise.all([indexMapsReadyPromise, loadGfaIndexMaps(), fetchAllDataRows()]);
    state.rows = allRows;
    renderAll();
    const banner = el("loadingBanner");
    if (banner) banner.hidden = true;
  }

  // 쿠팡 PA 키워드 리포트처럼 "키워드" 컬럼이 있는 파일만, 날짜+캠페인ID+
  // 캠페인명+그룹명+VIID+광고집행상품명+광고전환매출발생상품명 단위로(즉
  // 표1~4가 보는 모든 차원은 그대로 두고 키워드만) 묶어서 노출수/클릭수/
  // 광고비/전환수/전환매출만 합산한 가벼운 행을 만든다. 품목/구분은 이
  // 그룹 안에서 전부 같은 값이라(캠페인ID/캠페인명/그룹명/VIID가 품목
  // 조회 키라서) 대표로 하나만 가져오면 된다.
  const DASHBOARD_METRIC_FIELDS = ["cost", "impressions", "clicks", "conversions", "revenue"];

  function aggregateKeywordRows(rows, cols) {
    const platform = "coupang";
    const keyCols = {
      campaignId: resolveFieldForCols(cols, platform, "campaignId"),
      campaign: resolveFieldForCols(cols, platform, "campaign"),
      group: resolveFieldForCols(cols, platform, "group"),
      viid: resolveFieldForCols(cols, platform, "viid"),
      execProductName: resolveFieldForCols(cols, platform, "execProductName"),
    };
    const soldProductCol = cols.has("광고전환매출발생 상품명") ? "광고전환매출발생 상품명" : null;
    const metricCols = {};
    for (const f of DASHBOARD_METRIC_FIELDS) metricCols[f] = resolveFieldForCols(cols, platform, f);
    const groups = new Map();
    for (const row of rows) {
      const key = [
        row.__date,
        ...Object.values(keyCols).map((c) => (c ? row[c] : "")),
        soldProductCol ? row[soldProductCol] : "",
      ].join("\u0001");
      let g = groups.get(key);
      if (!g) {
        g = {
          __platform: row.__platform,
          __file: row.__file,
          __filePath: row.__filePath,
          __productType: row.__productType,
          __date: row.__date,
          __item: row.__item,
          __subItem: row.__subItem,
        };
        for (const name of Object.keys(keyCols)) {
          const col = keyCols[name];
          if (col) g[col] = row[col];
        }
        if (soldProductCol) g[soldProductCol] = row[soldProductCol];
        groups.set(key, g);
      }
      for (const f of DASHBOARD_METRIC_FIELDS) {
        const col = metricCols[f];
        if (!col) continue;
        const n = parseNumberLoose(row[col]);
        if (Number.isFinite(n)) g[col] = (g[col] || 0) + n;
      }
    }
    return [...groups.values()];
  }

  // raw 파일 하나를 읽어서 __platform/__productType/__date/__item/__subItem까지
  // 채운 행 배열로 만든다. aggregateKeywords가 true(기본값)이고 "키워드"
  // 컬럼이 있는 파일(쿠팡 PA)이면, 원본 키워드별 행은 반환하지 않고 바로
  // 키워드 단위를 합친 가벼운 행만 돌려준다 - 수십만 행짜리 원본을 한
  // 순간이라도 더 오래 메모리에 들고 있지 않기 위해서다.
  async function loadAndTagFile(f, aggregateKeywords = true) {
    const rows = await parseDataFile(f.path);
    if (rows.length === 0) return [];
    const productType = detectProductType(f.platform, f.name);
    const cols = new Set(Object.keys(rows[0]));
    state.colsByFile.set(f.path, cols);
    const dateCol = resolveFieldForCols(cols, f.platform, "date");
    // 품목 분류는 INDEX 조회표가 준비돼야 한다. raw 파일 fetch가 INDEX
    // fetch보다 먼저 끝나는 경우(특히 fetchFullCoupangRows처럼 이미 로드된
    // 뒤에 다시 부를 때는 즉시 반환됨)를 대비해 분류 직전에 기다린다.
    if (f.platform === "coupang" && indexMapsReadyPromise) await indexMapsReadyPromise;
    const tagged = rows.map((row) => {
      const out = { __platform: f.platform, __file: f.name, __filePath: f.path, __productType: productType, ...row };
      out.__date = dateCol ? parseDateLoose(row[dateCol]) : null;
      if (f.platform === "coupang") {
        const cls = resolveItemClassification(out, f.path, f.platform, productType);
        out.__item = cls.item;
        out.__subItem = cls.subItem;
      } else {
        out.__item = null;
        out.__subItem = null;
      }
      return out;
    });
    if (aggregateKeywords && f.platform === "coupang") {
      const keywordCol = resolveFieldForCols(cols, "coupang", "keyword");
      if (keywordCol) return aggregateKeywordRows(tagged, cols);
    }
    return tagged;
  }

  // raw 파일을 하나씩 순서대로 fetch+parse하면 느려서(raw가 많으면 파일도
  // 여러 개, 각각 수만~수십만 행), 서로 독립적인 파일들을 병렬로 불러온다.
  async function fetchAllDataRows() {
    state.colsByFile = new Map();
    fieldResolutionCache = new Map();
    const perFileRows = await Promise.all(
      state.manifest.files.map(async (f) => {
        try {
          return await loadAndTagFile(f);
        } catch (e) {
          console.error("파일을 불러오지 못했습니다:", f.path, e);
          return [];
        }
      })
    );
    return perFileRows.flat();
  }

  // "리포트RAW 다운받기"의 CE_RAW/MX_RAW(키워드별 1:1)와 쿠팡 키워드_RAW는
  // 키워드 단위가 꼭 필요해서, 평소 표1~4가 쓰는 가벼운 state.rows가 아니라
  // 클릭했을 때만 쿠팡 파일들을 다시 받아서 키워드를 합치지 않고 그대로 쓴다.
  async function fetchFullCoupangRows() {
    // colsByFile에 남아있는(키워드가 빠진) 가벼운 버전 캐시를 비워야, 이번에
    // 다시 받는 전체 행에서 "키워드" 컬럼이 제대로 잡힌다.
    fieldResolutionCache = new Map();
    const files = state.manifest.files.filter((f) => f.platform === "coupang");
    const perFileRows = await Promise.all(
      files.map(async (f) => {
        try {
          return await loadAndTagFile(f, false);
        } catch (e) {
          console.error("파일을 불러오지 못했습니다:", f.path, e);
          return [];
        }
      })
    );
    return perFileRows.flat();
  }

  // ---------- 집계 ----------

  // 네이버 raw의 "총비용"은 매체에서 추출할 때부터 VAT 포함 금액으로 내려와서,
  // VAT 제외 금액으로 보여주려면 1.1로 나눠야 한다(쿠팡 광고비는 원래 VAT
  // 제외라 그대로 둔다).
  function naverVatExcludedCost(cost) {
    return Number.isFinite(cost) ? cost / 1.1 : cost;
  }

  function accumulateRow(groups, key, base, row, fp, platform) {
    if (!groups.has(key)) {
      groups.set(key, { ...base, cost: 0, impressions: 0, clicks: 0, conversions: 0, revenue: 0 });
    }
    const g = groups.get(key);
    const costCol = resolveField(fp, platform, "cost");
    const imprCol = resolveField(fp, platform, "impressions");
    const clickCol = resolveField(fp, platform, "clicks");
    const convCol = resolveField(fp, platform, "conversions");
    const revCol = resolveField(fp, platform, "revenue");

    const costRaw = costCol ? parseNumberLoose(row[costCol]) : NaN;
    const cost = platform === "naver" ? naverVatExcludedCost(costRaw) : costRaw;
    const impr = imprCol ? parseNumberLoose(row[imprCol]) : NaN;
    const clicks = clickCol ? parseNumberLoose(row[clickCol]) : NaN;
    const conv = convCol ? parseNumberLoose(row[convCol]) : NaN;
    const rev = revCol ? parseNumberLoose(row[revCol]) : NaN;

    if (Number.isFinite(cost)) g.cost += cost;
    if (Number.isFinite(impr)) g.impressions += impr;
    if (Number.isFinite(clicks)) g.clicks += clicks;
    if (Number.isFinite(conv)) g.conversions += conv;
    if (Number.isFinite(rev)) g.revenue += rev;
  }

  function finalizeGroups(groups) {
    return [...groups.values()].map((g) => ({
      ...g,
      ctr: g.impressions > 0 ? (g.clicks / g.impressions) * 100 : null,
      cpc: g.clicks > 0 ? g.cost / g.clicks : null,
      cvr: g.clicks > 0 ? (g.conversions / g.clicks) * 100 : null,
      aov: g.conversions > 0 ? g.revenue / g.conversions : null,
      roas: g.cost > 0 ? (g.revenue / g.cost) * 100 : null,
    }));
  }

  // 표1. 전체기간 합산 (날짜 무시). 쿠팡은 INDEX 분류로 찾은 품목 기준으로만
  // 묶는다(상품구분이 달라도 같은 품목이면 하나로 합산). 네이버는 품목 분류
  // 데이터가 없어서 상품구분+캠페인명 기준(쿠팡 표1처럼 그룹 구분 없이
  // 한 가지 기준으로만 묶는다).
  function computeTable1(platform) {
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const date = row.__date;
      const productType = row.__productType;
      const itemCategory = row.__item;
      if (!date || !passesFilters(platform, date, productType, itemCategory)) continue;
      if (platform === "coupang") {
        accumulateRow(groups, itemCategory, { itemCategory }, row, fp, platform);
      } else {
        const campaignCol = resolveField(fp, platform, "campaign");
        const campaign = campaignCol ? row[campaignCol] ?? "" : "";
        const key = `${productType}\u0001${campaign}`;
        accumulateRow(groups, key, { productType, campaign }, row, fp, platform);
      }
    }
    const result = finalizeGroups(groups);
    if (platform === "coupang") {
      return result.sort((a, b) => a.itemCategory.localeCompare(b.itemCategory, "ko"));
    }
    return result.sort(
      (a, b) => a.productType.localeCompare(b.productType) || a.campaign.localeCompare(b.campaign, "ko")
    );
  }

  // 표2. 상품구분 기준 전체기간 합산
  function computeTable2(platform) {
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const date = row.__date;
      const productType = row.__productType;
      const itemCategory = row.__item;
      if (!date || !passesFilters(platform, date, productType, itemCategory)) continue;
      accumulateRow(groups, productType, { productType }, row, fp, platform);
    }
    return finalizeGroups(groups).sort((a, b) => (b.revenue || 0) - (a.revenue || 0));
  }

  // 표3. 주차 기준 합산(상품구분 구분 없이 전체 합산)
  function computeTable3(platform) {
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const date = row.__date;
      const productType = row.__productType;
      const itemCategory = row.__item;
      if (!date || !passesFilters(platform, date, productType, itemCategory)) continue;
      const week = weekOf(date);
      accumulateRow(groups, week.key, { weekKey: week.key, weekLabel: week.label }, row, fp, platform);
    }
    return finalizeGroups(groups).sort((a, b) => a.weekKey.localeCompare(b.weekKey));
  }

  // 표4. 상품구분+상품명 기준 합산 후 전환매출 상위 10개
  function computeTable4(platform) {
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const date = row.__date;
      const productType = row.__productType;
      const itemCategory = row.__item;
      if (!date || !passesFilters(platform, date, productType, itemCategory)) continue;
      const campaignCol = resolveField(fp, platform, "campaign");
      const groupCol = resolveField(fp, platform, "group");
      const campaign = campaignCol ? row[campaignCol] ?? "" : "";
      const group = groupCol ? row[groupCol] ?? "" : "";
      const productName = resolveProductLabel(row, fp, platform, productType, campaign, group);
      const key = `${productType}\u0001${productName}`;
      accumulateRow(groups, key, { productType, productName }, row, fp, platform);
    }
    return finalizeGroups(groups)
      .sort((a, b) => (b.revenue || 0) - (a.revenue || 0))
      .slice(0, 10)
      .map((r, i) => ({ ...r, rank: i + 1 }));
  }

  // ---------- 렌더링 ----------

  function renderAll() {
    renderTabs();
    renderFileNotice();
    renderSlicers();
    renderTable1();
    renderTable2();
    renderTable3();
    renderTable4();
  }

  function renderTabs() {
    for (const btn of document.querySelectorAll(".tab-btn")) {
      btn.classList.toggle("active", btn.dataset.platform === state.activePlatform);
    }
    // 플랫폼별 전용 버튼(쿠팡 "리포트RAW 다운받기", 네이버 "GFA리포트RAW 다운받기")은
    // 현재 탭에 맞는 것만 보여준다.
    for (const btn of document.querySelectorAll(".platform-only")) {
      btn.style.display = btn.dataset.platform === state.activePlatform ? "" : "none";
    }
  }

  function renderSlicerGroup(containerId, values, selectedSet, formatLabel) {
    const area = el(containerId);
    area.innerHTML = "";
    if (values.length === 0) {
      area.innerHTML = `<span class="empty-state">-</span>`;
      return;
    }
    for (const v of values) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "slicer-btn" + (selectedSet.has(v) ? " active" : "");
      btn.textContent = formatLabel ? formatLabel(v) : v;
      btn.addEventListener("click", () => {
        if (selectedSet.has(v)) selectedSet.delete(v);
        else selectedSet.add(v);
        state.table1Page = 1;
        renderAll();
      });
      area.appendChild(btn);
    }
  }

  function renderSlicers() {
    const platform = state.activePlatform;
    const options = computeFilterOptions(platform);
    const f = state.filters[platform];
    renderSlicerGroup("yearSlicer", options.years, f.years, (y) => `${y}년`);
    renderSlicerGroup("monthSlicer", options.months, f.months, (m) => `${Number(m)}월`);
    renderSlicerGroup("productTypeSlicer", options.productTypes, f.productTypes);
    renderSlicerGroup("itemCategorySlicer", options.itemCategories, f.itemCategories);
  }

  function formatSize(bytes) {
    if (!Number.isFinite(bytes)) return "";
    if (bytes < 1024) return `${bytes}B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)}KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)}MB`;
  }

  function filesForActivePlatform() {
    return state.manifest.files.filter((f) => f.platform === state.activePlatform);
  }

  // 파일 선택(체크박스) 없이, 지금 어떤 파일이 올라와 있는지만 안내하는 공지창.
  function renderFileNotice() {
    const area = el("fileNoticeArea");
    const files = filesForActivePlatform();
    if (files.length === 0) {
      area.innerHTML = `<div class="empty-state">아직 업로드된 데이터가 없습니다. <code>data/${state.activePlatform}</code> 폴더에 raw csv/xlsx 파일을 올려주세요.</div>`;
      return;
    }
    area.innerHTML = files
      .map((f) => `<div class="file-row"><span class="file-name">${escapeHtml(f.name)} (${formatSize(f.size)})</span></div>`)
      .join("");
  }

  function cellText(row, col) {
    if (col.key === "rank" || col.key === "productType" || col.key === "productName" || col.key === "campaign" || col.key === "group" || col.key === "weekLabel" || col.key === "itemCategory") {
      return row[col.key] ?? "";
    }
    return formatMetric(row[col.key], col.format);
  }

  function renderMetricTable(headEl, bodyEl, columns, rows, emptyMessage) {
    headEl.innerHTML = "";
    const tr = document.createElement("tr");
    for (const col of columns) {
      const th = document.createElement("th");
      th.textContent = col.label;
      tr.appendChild(th);
    }
    headEl.appendChild(tr);

    bodyEl.innerHTML = "";
    if (rows.length === 0) {
      const tr2 = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = columns.length;
      td.className = "empty-state";
      td.textContent = emptyMessage;
      tr2.appendChild(td);
      bodyEl.appendChild(tr2);
      return;
    }
    for (const row of rows) {
      const tr2 = document.createElement("tr");
      for (const col of columns) {
        const td = document.createElement("td");
        td.textContent = cellText(row, col);
        tr2.appendChild(td);
      }
      bodyEl.appendChild(tr2);
    }
  }

  function renderTable1() {
    const rows = computeTable1(state.activePlatform);
    const totalPages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    state.table1Page = Math.min(Math.max(1, state.table1Page), totalPages);
    const pageRows = rows.slice((state.table1Page - 1) * PAGE_SIZE, state.table1Page * PAGE_SIZE);
    renderMetricTable(el("table1Head"), el("table1Body"), table1ColsFor(state.activePlatform), pageRows, "데이터가 없습니다.");
    el("table1PageInfo").textContent = `총 ${rows.length.toLocaleString("ko-KR")}행 · ${state.table1Page} / ${totalPages} 페이지`;
    el("table1PrevBtn").disabled = state.table1Page <= 1;
    el("table1NextBtn").disabled = state.table1Page >= totalPages;
  }

  function renderTable2() {
    const rows = computeTable2(state.activePlatform);
    renderMetricTable(el("table2Head"), el("table2Body"), TABLE2_COLS, rows, "데이터가 없습니다.");
  }

  // 표3 맨 아래에 최근 주차 2개의 증감율(%)을 한 줄 추가한다. 공식 예:
  // 2주차/1주차-1 (최신 주 값 / 그 전 주 값 - 1). 주차가 2개 미만이면 생략.
  function appendWeekCompareRow(bodyEl, rows) {
    if (rows.length < 2) return;
    const prev = rows[rows.length - 2];
    const latest = rows[rows.length - 1];
    const tr = document.createElement("tr");
    tr.style.fontWeight = "700";
    const labelTd = document.createElement("td");
    labelTd.textContent = `${prev.weekLabel} → ${latest.weekLabel} 증감율`;
    tr.appendChild(labelTd);
    for (const col of METRIC_COLUMNS) {
      const td = document.createElement("td");
      const a = prev[col.key];
      const b = latest[col.key];
      if (!Number.isFinite(a) || a === 0 || !Number.isFinite(b)) {
        td.textContent = "-";
      } else {
        const pct = (b / a - 1) * 100;
        td.textContent = `${pct.toLocaleString("ko-KR", { maximumFractionDigits: 2, signDisplay: "exceptZero" })}%`;
      }
      tr.appendChild(td);
    }
    bodyEl.appendChild(tr);
  }

  function renderTable3() {
    const rows = computeTable3(state.activePlatform);
    renderMetricTable(el("table3Head"), el("table3Body"), TABLE3_COLS, rows, "데이터가 없습니다.");
    appendWeekCompareRow(el("table3Body"), rows);
  }

  function renderTable4() {
    const rows = computeTable4(state.activePlatform);
    renderMetricTable(el("table4Head"), el("table4Body"), TABLE4_COLS, rows, "판매 데이터가 없습니다.");
  }

  // ---------- 이벤트 ----------

  for (const btn of document.querySelectorAll(".tab-btn")) {
    btn.addEventListener("click", () => {
      state.activePlatform = btn.dataset.platform;
      state.table1Page = 1;
      renderAll();
    });
  }

  el("table1PrevBtn").addEventListener("click", () => {
    state.table1Page = Math.max(1, state.table1Page - 1);
    renderTable1();
  });
  el("table1NextBtn").addEventListener("click", () => {
    state.table1Page += 1;
    renderTable1();
  });

  // 리포트RAW 다운로드(쿠팡_전환상품_RAW/쿠팡 키워드_RAW 등)에서 원본 컬럼
  // 값을 그대로 옮길 때 쓴다. 컬럼이 없거나 비어있으면 "-".
  function weeklyRawText(row, fp, platform, field) {
    const col = resolveField(fp, platform, field);
    if (!col) return "-";
    const v = row[col];
    return v == null || v === "" ? "-" : v;
  }

  function weeklyRawNum(row, fp, platform, field) {
    const col = resolveField(fp, platform, field);
    if (!col) return "-";
    const n = parseNumberLoose(row[col]);
    return Number.isFinite(n) ? Math.round(n * 100) / 100 : "-";
  }

  // "리포트양식/CE_MX_주간RAW_양식.xlsb"의 CE_RAW/MX_RAW 시트와 동일한 25개
  // 컬럼 구성. 쿠팡 raw 한 행을 CE_RAW/MX_RAW 양식의 한 행(배열)으로 바꾼다.
  // 품목으로 사업부(CE/MX)를 특정할 수 없으면(미분류) null을 돌려줘서 두
  // 시트 어디에도 넣지 않는다 - 어느 사업부 매출인지 모르는 금액을 섞어
  // 넣지 않기 위함.
  const WEEKLY_RAW_HEADER = [
    "집행 년", "집행 월", "날짜", "주차", "구분", "광고유형", "캠페인 ID", "품목",
    "캠페인명", "그룹명", "광고집행상품 VIID", "광고집행 상품명", "키워드",
    "광고비", "노출수", "클릭수",
    "판매수량(1일)", "직접 판매수량(1일)", "전환매출액(1일)", "직접 전환매출액(1일)",
    "판매수량(14일)", "직접 판매수량(14일)", "전환매출액(14일)", "직접 전환매출액(14일)",
    "세부 품목",
  ];

  function buildWeeklyRawRow(row, fp, productType) {
    const platform = "coupang";
    const date = row.__date;
    if (!date) return null;
    const item = row.__item;
    const subItem = row.__subItem;
    const division = classifyDivision(item);
    if (!division || !passesFilters(platform, date, productType, item)) return null;
    return {
      division,
      values: [
        `${yearOf(date)}년`,
        `${Number(monthOf(date))}월`,
        Number(date.replace(/-/g, "")),
        weekOfMonthLabel(date),
        division,
        productType,
        weeklyRawText(row, fp, platform, "campaignId"),
        item,
        weeklyRawText(row, fp, platform, "campaign"),
        weeklyRawText(row, fp, platform, "group"),
        weeklyRawText(row, fp, platform, "viid"),
        weeklyRawText(row, fp, platform, "execProductName"),
        weeklyRawText(row, fp, platform, "keyword"),
        weeklyRawNum(row, fp, platform, "cost"),
        weeklyRawNum(row, fp, platform, "impressions"),
        weeklyRawNum(row, fp, platform, "clicks"),
        weeklyRawNum(row, fp, platform, "totalConv1d"),
        weeklyRawNum(row, fp, platform, "conversions"),
        weeklyRawNum(row, fp, platform, "totalRevenue1d"),
        weeklyRawNum(row, fp, platform, "revenue"),
        weeklyRawNum(row, fp, platform, "totalConv14d"),
        weeklyRawNum(row, fp, platform, "directConv14d"),
        weeklyRawNum(row, fp, platform, "totalRevenue14d"),
        weeklyRawNum(row, fp, platform, "directRevenue14d"),
        subItem,
      ],
    };
  }

  // 쿠팡 탭에 지금 적용된 필터(년도/월/광고상품유형/품목) 기준으로 CE_RAW,
  // MX_RAW, 그리고 그 둘을 합친 ce,mx 전체_raw 행을 만든다.
  function buildWeeklyRawSheets(rows) {
    const ceRows = [];
    const mxRows = [];
    for (const row of rows) {
      if (row.__platform !== "coupang") continue;
      const built = buildWeeklyRawRow(row, row.__filePath, row.__productType);
      if (!built) continue;
      (built.division === "CE" ? ceRows : mxRows).push(built.values);
    }
    return {
      CE_RAW: XLSX.utils.aoa_to_sheet([WEEKLY_RAW_HEADER, ...ceRows]),
      MX_RAW: XLSX.utils.aoa_to_sheet([WEEKLY_RAW_HEADER, ...mxRows]),
      "ce,mx 전체_raw": XLSX.utils.aoa_to_sheet([WEEKLY_RAW_HEADER, ...ceRows, ...mxRows]),
    };
  }


  // ---------- 리포트 다운로드 (쿠팡_RAW/쿠팡_전환상품_RAW/쿠팡 키워드_RAW 양식) ----------
  // "리포트양식/쿠팡상품광고_월간리포트양식RAW.xlsb"의 세 시트와 같은 기준으로
  // 직접 합산해서 만든다. 피벗 자체가 아니라 그 피벗이 보여주는 숫자만
  // 재현하는 것이라 원본 피벗/바이너리는 건드리지 않는다.
  const WEEKLY_REPORT_METRIC_FIELDS = [
    "cost", "impressions", "clicks",
    "totalConv1d", "conversions", "totalRevenue1d", "revenue",
    "totalConv14d", "directConv14d", "totalRevenue14d", "directRevenue14d",
  ];

  function sumWeeklyMetrics(row, fp, platform, acc) {
    for (const field of WEEKLY_REPORT_METRIC_FIELDS) {
      const col = resolveField(fp, platform, field);
      const n = col ? parseNumberLoose(row[col]) : NaN;
      if (Number.isFinite(n)) acc[field] = (acc[field] || 0) + n;
    }
  }

  function weeklyMetricValues(metrics) {
    return WEEKLY_REPORT_METRIC_FIELDS.map((f) =>
      Number.isFinite(metrics[f]) ? Math.round(metrics[f] * 100) / 100 : "-"
    );
  }

  // 쿠팡_RAW: 월간 집계라 날짜/캠페인ID/그룹/VIID/상품명/키워드를 다 버리고
  // 집행년+집행월+구분+품목명+광고유형+캠페인명+세부품목 단위로 합산한다.
  // "예산처" 칼럼은 우리 쪽에 원본 데이터가 없어서 값은 채우지 않고 빈 칸으로
  // 둔다(칼럼 자체는 양식과 자리를 맞추기 위해 남겨둔다).
  const MONTHLY_RAW_HEADER = [
    "집행 년", "집행 월", "구분", "품목명", "예산처", "광고유형", "캠페인명",
    "광고비", "노출수", "클릭수",
    "판매수량(1일)", "판매수량(1일,직접)", "전환매출액(1일)", "전환매출액(1일,직접)",
    "판매수량(14일)", "판매수량(14일,직접)", "전환매출액(14일)", "전환매출액(14일,직접)",
    "세부 품목",
  ];

  function buildMonthlyRawRows(rows) {
    const platform = "coupang";
    const groups = new Map();
    for (const row of rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const productType = row.__productType;
      const date = row.__date;
      if (!date) continue;
      const item = row.__item;
      const subItem = row.__subItem;
      const division = classifyDivision(item);
      if (!division || !passesFilters(platform, date, productType, item)) continue;
      const campaignRaw = weeklyRawText(row, fp, platform, "campaign");
      const month = String(Number(monthOf(date)));
      const key = [yearOf(date), month, productType, item, campaignRaw, subItem].join("\u0001");
      if (!groups.has(key)) {
        groups.set(key, { year: yearOf(date), month, division, item, subItem, productType, campaignRaw, metrics: {} });
      }
      sumWeeklyMetrics(row, fp, platform, groups.get(key).metrics);
    }
    return [...groups.values()].map((g) => [
      `${g.year}년`,
      `${g.month}월`,
      g.division,
      g.item,
      "",
      g.productType,
      g.campaignRaw,
      ...weeklyMetricValues(g.metrics),
      g.subItem,
    ]);
  }

  // 쿠팡_전환상품_RAW: 월간 집계라 주차/날짜/그룹명을 버리고 집행년+집행월+
  // 상품구분+캠페인ID+캠페인명+VIID+상품명 단위로(키워드 구분 없이) 합산.
  const CONVERSION_PRODUCT_HEADER = [
    "집행 년", "집행 월", "구분", "품목명", "광고유형", "캠페인 ID", "캠페인명",
    "광고집행상품 VIID", "광고집행 상품명",
    "광고비", "노출수", "클릭수",
    "판매수량(1일)", "직접 판매수량(1일)", "전환매출액(1일)", "직접 전환매출액(1일)",
    "판매수량(14일)", "직접 판매수량(14일)", "전환매출액(14일)", "직접 전환매출액(14일)",
    "세부 품목",
  ];

  function buildConversionProductRows(rows) {
    const platform = "coupang";
    const groups = new Map();
    for (const row of rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const productType = row.__productType;
      const date = row.__date;
      if (!date) continue;
      const item = row.__item;
      const subItem = row.__subItem;
      const division = classifyDivision(item);
      if (!division || !passesFilters(platform, date, productType, item)) continue;
      const campaignIdRaw = weeklyRawText(row, fp, platform, "campaignId");
      const campaignRaw = weeklyRawText(row, fp, platform, "campaign");
      const viidRaw = weeklyRawText(row, fp, platform, "viid");
      const execProductNameRaw = weeklyRawText(row, fp, platform, "execProductName");
      const month = String(Number(monthOf(date)));
      const key = [yearOf(date), month, productType, campaignIdRaw, campaignRaw, viidRaw, execProductNameRaw].join("\u0001");
      if (!groups.has(key)) {
        groups.set(key, {
          year: yearOf(date), month, productType, division, item, subItem,
          campaignIdRaw, campaignRaw, viidRaw, execProductNameRaw,
          metrics: {},
        });
      }
      sumWeeklyMetrics(row, fp, platform, groups.get(key).metrics);
    }
    return [...groups.values()].map((g) => [
      `${g.year}년`,
      `${g.month}월`,
      g.division,
      g.item,
      g.productType,
      g.campaignIdRaw,
      g.campaignRaw,
      g.viidRaw,
      g.execProductNameRaw,
      ...weeklyMetricValues(g.metrics),
      g.subItem,
    ]);
  }

  // 쿠팡 키워드_RAW: PA만 대상으로, 날짜+캠페인ID+캠페인명+VIID+키워드
  // 단위로(그룹 구분 없이) 합산. 날짜는 년/월/일로 나눠 담는다.
  const KEYWORD_HEADER = [
    "년", "월", "일", "품목", "광고유형", "캠페인ID", "캠페인명", "광고집행상품 VIID", "키워드",
    "노출수", "클릭수", "광고비",
    "총 판매수량(1일)", "직접 판매수량(1일)", "총 전환매출액(1일)", "직접 전환매출액(1일)",
    "총 판매수량(14일)", "직접 판매수량(14일)", "총 전환매출액(14일)", "직접 전환매출액(14일)",
    "세부 품목",
  ];

  function buildKeywordRows(rows) {
    const platform = "coupang";
    const groups = new Map();
    for (const row of rows) {
      if (row.__platform !== platform || row.__productType !== "PA") continue;
      const fp = row.__filePath;
      const productType = row.__productType;
      const date = row.__date;
      if (!date) continue;
      const item = row.__item;
      const subItem = row.__subItem;
      const division = classifyDivision(item);
      if (!division || !passesFilters(platform, date, productType, item)) continue;
      const campaignIdRaw = weeklyRawText(row, fp, platform, "campaignId");
      const campaignRaw = weeklyRawText(row, fp, platform, "campaign");
      const viidRaw = weeklyRawText(row, fp, platform, "viid");
      const keywordRaw = weeklyRawText(row, fp, platform, "keyword");
      const key = [date, campaignIdRaw, campaignRaw, viidRaw, keywordRaw].join("\u0001");
      if (!groups.has(key)) {
        groups.set(key, { date, productType, item, subItem, campaignIdRaw, campaignRaw, viidRaw, keywordRaw, metrics: {} });
      }
      sumWeeklyMetrics(row, fp, platform, groups.get(key).metrics);
    }
    return [...groups.values()].map((g) => [
      Number(yearOf(g.date)),
      Number(monthOf(g.date)),
      Number(g.date.slice(8, 10)),
      g.item,
      g.productType,
      g.campaignIdRaw,
      g.campaignRaw,
      g.viidRaw,
      g.keywordRaw,
      ...weeklyMetricValues(g.metrics),
      g.subItem,
    ]);
  }

  el("downloadReportBtn").addEventListener("click", async () => {
    const btn = el("downloadReportBtn");
    const originalText = btn.textContent;
    btn.disabled = true;
    btn.textContent = "리포트 준비 중...";
    try {
      // CE_RAW/MX_RAW/쿠팡 키워드_RAW는 키워드 단위가 꼭 필요해서, 평소
      // 표1~4가 쓰는 가벼운 state.rows 대신 지금 클릭한 시점에 쿠팡 파일을
      // 다시 받아서(키워드를 합치지 않은 채로) 쓴다.
      const fullRows = await fetchFullCoupangRows();
      const workbook = XLSX.utils.book_new();
      const rawSheets = buildWeeklyRawSheets(fullRows);
      for (const name of Object.keys(rawSheets)) {
        XLSX.utils.book_append_sheet(workbook, rawSheets[name], name);
      }
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet([MONTHLY_RAW_HEADER, ...buildMonthlyRawRows(fullRows)]),
        "쿠팡_RAW"
      );
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet([CONVERSION_PRODUCT_HEADER, ...buildConversionProductRows(fullRows)]),
        "쿠팡_전환상품_RAW"
      );
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet([KEYWORD_HEADER, ...buildKeywordRows(fullRows)]),
        "쿠팡 키워드_RAW"
      );
      XLSX.writeFile(workbook, `쿠팡_상품광고_주간리포트_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (e) {
      console.error("리포트RAW를 만들지 못했습니다:", e);
      alert("리포트RAW를 만드는 중 문제가 생겼습니다. 다시 시도해주세요.");
    } finally {
      btn.disabled = false;
      btn.textContent = originalText;
    }
  });

  // ---------- GFA리포트RAW 다운로드 (네이버 GFA_캠페인RAW/GFA_상품RAW 양식) ----------
  // "리포트양식/네이버_GFA_RAW_양식.xlsb"의 두 시트와 같은 컬럼 구성. 품목명은
  // 그 양식 안의 캠페인ID/상품코드 -> 품목구분 조회표(월 단위로 재사용되는
  // 경우가 있어 월+ID로 조회, loadGfaIndexMaps/buildGfaIndexMaps 참고)로
  // 채운다. 쇼핑검색광고(카탈로그)는 캠페인ID/상품ID 자체가 없어서 두 시트
  // 모두에서 빠진다. "광고비(-VAT)"는 "총비용(VAT포함,원)"을 1.1로 나눈
  // 값이다(네이버 raw 추출 시 총비용에 VAT가 포함되어 내려온다).
  const GFA_CAMPAIGN_HEADER = [
    "집행 년", "집행 월", "캠페인 ID", "품목명", "광고유형", "캠페인명",
    "총비용(VAT포함,원)", "노출수", "클릭수", "구매완료(1일)", "구매완료 전환매출액(1일)", "광고비(-VAT)",
  ];

  const GFA_PRODUCT_HEADER = [
    "집행 년", "집행 월", "쇼핑몰 상품 ID", "품목명", "모델명", "광고유형", "캠페인명",
    "총비용(VAT포함,원)", "노출수", "클릭수", "구매완료(1일)", "구매완료 전환매출액(1일)", "광고비(-VAT)",
  ];

  function resolveGfaCampaignItem(month, campaignIdRaw) {
    const maps = state.gfaIndexMaps;
    if (!maps) return "미분류";
    const key = `${month}\u0001${normalizeId(campaignIdRaw)}`;
    return maps.byCampaign.has(key) ? maps.byCampaign.get(key) : "미분류";
  }

  function resolveGfaProductItem(month, productCodeRaw) {
    const maps = state.gfaIndexMaps;
    if (!maps) return { item: "미분류", modelName: "-" };
    const key = `${month}\u0001${normalizeId(productCodeRaw)}`;
    return maps.byProduct.has(key) ? maps.byProduct.get(key) : { item: "미분류", modelName: "-" };
  }

  function sumGfaMetrics(row, fp, platform, acc) {
    for (const field of ["cost", "impressions", "clicks", "conversions", "revenue"]) {
      const col = resolveField(fp, platform, field);
      const n = col ? parseNumberLoose(row[col]) : NaN;
      if (Number.isFinite(n)) acc[field] = (acc[field] || 0) + n;
    }
  }

  function gfaMetricValues(metrics) {
    return ["cost", "impressions", "clicks", "conversions", "revenue"].map((f) =>
      Number.isFinite(metrics[f]) ? Math.round(metrics[f] * 100) / 100 : "-"
    );
  }

  // 캠페인ID가 있는 상품구분(애드부스트/카탈로그)만 대상으로, 집행년+집행월+
  // 캠페인ID+캠페인명 단위로(그룹/상품 구분 없이) 합산.
  function buildGfaCampaignRows() {
    const platform = "naver";
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const productType = row.__productType;
      const campaignIdCol = resolveField(fp, platform, "campaignId");
      if (!campaignIdCol) continue; // 쇼핑검색광고(카탈로그)는 캠페인ID가 없음
      const date = row.__date;
      if (!date) continue;
      const month = String(Number(monthOf(date)));
      const campaignIdRaw = row[campaignIdCol];
      const campaignRaw = weeklyRawText(row, fp, platform, "campaign");
      const key = [yearOf(date), month, productType, normalizeId(campaignIdRaw), campaignRaw].join("\u0001");
      if (!groups.has(key)) {
        groups.set(key, { year: yearOf(date), month, productType, campaignIdRaw, campaignRaw, metrics: {} });
      }
      sumGfaMetrics(row, fp, platform, groups.get(key).metrics);
    }
    return [...groups.values()].map((g) => [
      `${g.year}년`,
      `${g.month}월`,
      g.campaignIdRaw,
      resolveGfaCampaignItem(g.month, g.campaignIdRaw),
      g.productType,
      g.campaignRaw,
      ...gfaMetricValues(g.metrics),
      Number.isFinite(g.metrics.cost) ? Math.round(naverVatExcludedCost(g.metrics.cost) * 100) / 100 : "-",
    ]);
  }

  // 쇼핑몰 상품 ID가 있는 애드부스트만 대상으로, 집행년+집행월+캠페인명+
  // 상품ID 단위로 합산.
  function buildGfaProductRows() {
    const platform = "naver";
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform || row.__productType !== "애드부스트") continue;
      const fp = row.__filePath;
      const productType = row.__productType;
      const productIdCol = resolveField(fp, platform, "shoppingProductId");
      if (!productIdCol) continue;
      const date = row.__date;
      if (!date) continue;
      const month = String(Number(monthOf(date)));
      const productIdRaw = row[productIdCol];
      const campaignRaw = weeklyRawText(row, fp, platform, "campaign");
      const key = [yearOf(date), month, productType, campaignRaw, normalizeId(productIdRaw)].join("\u0001");
      if (!groups.has(key)) {
        groups.set(key, { year: yearOf(date), month, productType, campaignRaw, productIdRaw, metrics: {} });
      }
      sumGfaMetrics(row, fp, platform, groups.get(key).metrics);
    }
    return [...groups.values()].map((g) => {
      const { item, modelName } = resolveGfaProductItem(g.month, g.productIdRaw);
      return [
        `${g.year}년`,
        `${g.month}월`,
        g.productIdRaw,
        item,
        modelName,
        g.productType,
        g.campaignRaw,
        ...gfaMetricValues(g.metrics),
        Number.isFinite(g.metrics.cost) ? Math.round(naverVatExcludedCost(g.metrics.cost) * 100) / 100 : "-",
      ];
    });
  }

  el("downloadGfaReportBtn").addEventListener("click", () => {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([GFA_CAMPAIGN_HEADER, ...buildGfaCampaignRows()]),
      "GFA_캠페인RAW"
    );
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([GFA_PRODUCT_HEADER, ...buildGfaProductRows()]),
      "GFA_상품RAW"
    );
    XLSX.writeFile(workbook, `네이버_GFA리포트RAW_${new Date().toISOString().slice(0, 10)}.xlsx`);
  });

  // ---------- 쇼핑검색광고(카탈로그형) RAW 다운로드 ----------
  // "리포트양식/네이버_쇼핑검색광고_월간리포트양식RAW.xlsb"의 네이버_RAW/네이버
  // 키워드_RAW 시트와 같은 모양으로 만든다.
  //   - 네이버_RAW의 "품목명"/"모델명"은 쇼핑검색광고(카탈로그) raw에 상품/모델
  //     컬럼이 없고, 캠페인/광고그룹명 표기 방식이 캠페인마다 달라서(자동
  //     생성된 "CN~/FF~" 코드형, 수동으로 "쇼핑검색_<품목>"처럼 지은 이름형)
  //     추측으로 채우면 사업부 집계가 틀릴 위험이 있어 일단 비워두고, 대신
  //     대시보드 표1(캠페인명/그룹명 기준)과 맞춰볼 수 있게 "참고_캠페인"/
  //     "참고_광고그룹" 2개 컬럼을 뒤에 추가해서 수동으로 채울 수 있게 한다.
  //   - 네이버 키워드_RAW는 별도로 올라온 "...쇼핑검색광고_검색어_RAW_..."
  //     리포트(캠페인유형/캠페인/광고그룹/PC-모바일매체/검색어/노출수/클릭수/
  //     총비용, 날짜 컬럼은 없고 파일명에만 날짜 범위가 있음)에서 가져온다.
  //     "모델명"은 이 리포트의 광고그룹 값이 이미 상품 모델코드라 그대로
  //     옮기고, "품목"은 역시 분류 판단이 필요해 비워둔다(모델명으로 대시보드
  //     와 맞춰보고 수동으로 채울 수 있다).
  const NAVER_SHOPPING_RAW_HEADER = [
    "년도", "집행월", "품목명", "모델명", "PC/MO", "시간대",
    "노출수", "클릭수", "총비용(VAT포함,원)", "평균노출순위",
    "참고_캠페인", "참고_광고그룹",
  ];
  const NAVER_SHOPPING_KEYWORD_HEADER = [
    "년도", "월", "품목", "모델명", "검색어", "노출수", "클릭수", "총비용(VAT포함,원)",
  ];

  // 일별 쇼핑검색광고(카탈로그) 리포트(PC/모바일 매체+시간대별까지 쪼개진 것)만
  // 대상으로, 집행년+집행월+캠페인+광고그룹+PC/MO+시간대 단위로 합산한다.
  // 평균노출순위는 노출수로 가중평균한다.
  function buildNaverShoppingRawRows() {
    const platform = "naver";
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const hourlyCol = resolveField(fp, platform, "hourly");
      const pcMobileCol = resolveField(fp, platform, "pcMobile");
      if (!hourlyCol || !pcMobileCol) continue; // 쇼핑검색 일별 리포트만
      const date = row.__date;
      if (!date) continue;
      const campaignCol = resolveField(fp, platform, "campaign");
      const groupCol = resolveField(fp, platform, "group");
      const impressionsCol = resolveField(fp, platform, "impressions");
      const clicksCol = resolveField(fp, platform, "clicks");
      const costCol = resolveField(fp, platform, "cost");
      const rankCol = resolveField(fp, platform, "avgRank");
      const campaignRaw = campaignCol ? (row[campaignCol] ?? "-") : "-";
      const groupRaw = groupCol ? (row[groupCol] ?? "-") : "-";
      const pcMobileRaw = row[pcMobileCol] ?? "-";
      const hourlyRaw = row[hourlyCol] ?? "-";
      const month = String(Number(monthOf(date)));
      const key = [yearOf(date), month, campaignRaw, groupRaw, pcMobileRaw, hourlyRaw].join("\u0001");
      if (!groups.has(key)) {
        groups.set(key, {
          year: yearOf(date), month, campaignRaw, groupRaw, pcMobileRaw, hourlyRaw,
          impressions: 0, clicks: 0, cost: 0, rankWeighted: 0, rankWeight: 0,
        });
      }
      const g = groups.get(key);
      const impressions = parseNumberLoose(row[impressionsCol]) || 0;
      g.impressions += impressions;
      g.clicks += parseNumberLoose(row[clicksCol]) || 0;
      g.cost += parseNumberLoose(row[costCol]) || 0;
      const rank = rankCol ? parseNumberLoose(row[rankCol]) : NaN;
      if (Number.isFinite(rank)) {
        g.rankWeighted += rank * impressions;
        g.rankWeight += impressions;
      }
    }
    return [...groups.values()].map((g) => [
      `${g.year}년`,
      `${g.month}월`,
      "",
      "",
      g.pcMobileRaw,
      g.hourlyRaw,
      g.impressions,
      g.clicks,
      g.cost,
      g.rankWeight > 0 ? Math.round((g.rankWeighted / g.rankWeight) * 10) / 10 : "-",
      g.campaignRaw,
      g.groupRaw,
    ]);
  }

  // "검색어" 컬럼이 있는 리포트만 대상(= 쇼핑검색_검색어_RAW). 이 리포트는
  // 행에 날짜가 없어서 파일명의 날짜 범위로 집행년/집행월을 구한다.
  function buildNaverShoppingKeywordRows() {
    const platform = "naver";
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const keywordCol = resolveField(fp, platform, "keyword");
      if (!keywordCol) continue;
      const date = parseDateFromFileName(row.__file);
      if (!date) continue;
      const groupCol = resolveField(fp, platform, "group");
      const impressionsCol = resolveField(fp, platform, "impressions");
      const clicksCol = resolveField(fp, platform, "clicks");
      const costCol = resolveField(fp, platform, "cost");
      const keywordRaw = row[keywordCol] ?? "-";
      const modelRaw = groupCol ? (row[groupCol] ?? "-") : "-";
      const month = String(Number(monthOf(date)));
      const key = [yearOf(date), month, modelRaw, keywordRaw].join("\u0001");
      if (!groups.has(key)) {
        groups.set(key, { year: yearOf(date), month, modelRaw, keywordRaw, impressions: 0, clicks: 0, cost: 0 });
      }
      const g = groups.get(key);
      g.impressions += parseNumberLoose(row[impressionsCol]) || 0;
      g.clicks += parseNumberLoose(row[clicksCol]) || 0;
      g.cost += parseNumberLoose(row[costCol]) || 0;
    }
    return [...groups.values()].map((g) => [
      `${g.year}년`,
      `${g.month}월`,
      "",
      g.modelRaw,
      g.keywordRaw,
      g.impressions,
      g.clicks,
      g.cost,
    ]);
  }

  el("downloadShoppingSearchBtn").addEventListener("click", () => {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([NAVER_SHOPPING_RAW_HEADER, ...buildNaverShoppingRawRows()]),
      "네이버_RAW"
    );
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([NAVER_SHOPPING_KEYWORD_HEADER, ...buildNaverShoppingKeywordRows()]),
      "네이버 키워드_RAW"
    );
    XLSX.writeFile(workbook, `네이버_쇼핑검색광고_월간리포트RAW_${new Date().toISOString().slice(0, 10)}.xlsx`);
  });

  const themeToggle = el("themeToggle");
  function applyStoredTheme() {
    try {
      const saved = localStorage.getItem("theme");
      if (saved) document.documentElement.setAttribute("data-theme", saved);
    } catch (e) {
      /* 저장소를 못 쓰는 환경(프라이빗 모드 등)이면 그냥 기본값 사용 */
    }
  }
  themeToggle.addEventListener("click", () => {
    const current = document.documentElement.getAttribute("data-theme");
    const next = current === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    try {
      localStorage.setItem("theme", next);
    } catch (e) {
      /* 저장 실패해도 화면 전환 자체는 계속 동작 */
    }
  });
  applyStoredTheme();

  loadManifest();
})();
