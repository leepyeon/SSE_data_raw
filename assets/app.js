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
      // 품목/상세품목이 광고 기준(캠페인/VIID/광고집행 상품명)으로 정해지므로
      // 상품도 광고한 상품 기준으로 맞춘다. "광고전환매출발생 상품명"(광고를
      // 보고 실제로 산 상품)을 쓰면 에어드레서 광고로 세탁기가 팔린 경우처럼
      // 필터와 다른 상품이 섞여 나온다.
      PA: "광고집행 상품명",
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
    { key: "roas", label: "ROAS", format: "pct0" },
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
  // 상품명은 30자까지만 보여주고(넘으면 "…", 마우스를 올리면 전체 이름),
  // 쿠팡은 상품의 VIID(옵션ID)도 함께 보여준다.
  const PRODUCT_NAME_MAX = 30;
  function table4ColsFor(platform) {
    return [
      { key: "rank", label: "순위" },
      { key: "productType", label: "상품구분" },
      { key: "productName", label: "상품명", maxLen: PRODUCT_NAME_MAX },
      ...(platform === "coupang" ? [{ key: "viid", label: "VIID" }] : []),
      ...METRIC_COLUMNS,
    ];
  }

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
    // 클릭했을 때 그 파일들만 따로 다시 받아서(withFullCoupangRows)
    // 처리한다.
    rows: [],
    colsByFile: new Map(), // filePath -> Set(원본 컬럼명)
    indexMaps: null, // INDEX/ 품목 분류 파일에서 만든 조회용 맵 (쿠팡 전용)
    gfaIndexMaps: null, // 리포트양식/ GFA 양식 파일의 품목 조회용 맵 (네이버 전용)
    activePlatform: "coupang",
    table1Page: 1,
    // 슬라이서 선택 상태 - 플랫폼(탭)마다 따로 기억한다(탭을 바꿔도 유지됨).
    filters: {
      coupang: { years: new Set(), months: new Set(), weeks: new Set(), divisions: new Set(), productTypes: new Set(), itemCategories: new Set(), subItems: new Set() },
      naver: { years: new Set(), months: new Set(), weeks: new Set(), divisions: new Set(), productTypes: new Set(), itemCategories: new Set(), subItems: new Set() },
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
      // cellText:false - 셀마다 화면표시용 문자열(.w)을 따로 만들지 않게 해서
      // 큰 파일을 읽을 때 메모리를 아낀다(아래 sheet_to_json은 raw:true라 .v만 씀).
      const workbook = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true, cellText: false });
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
    // 몇 월 몇 주차인지는 그 주의 목요일이 속한 달로 정한다(주의 과반이 있는
    // 달). 예: 09.28~10.04는 목요일이 10.01이라 10월 1주차, 10.05~10.11은
    // 10월 2주차, 10.26~11.01은 10월 5주차.
    const thursday = new Date(monday);
    thursday.setUTCDate(monday.getUTCDate() + 3);
    return {
      key: formatDateUTC(monday),
      label: `${fmt(monday)}~${fmt(sunday)}`,
      year: thursday.getUTCFullYear(),
      month: thursday.getUTCMonth() + 1,
      n: Math.ceil(thursday.getUTCDate() / 7),
    };
  }

  // 주차 키 목록이 모두 같은 년/월이면 "2주차", 여러 달에 걸치면 "10월 2주차"
  // 처럼 달을 붙여서 구분되게 한다.
  function weekNameFormatter(keys) {
    const months = new Set(keys.map((k) => {
      const w = weekOf(k);
      return `${w.year}-${w.month}`;
    }));
    const withMonth = months.size > 1;
    return (key) => {
      const w = weekOf(key);
      return withMonth ? `${w.month}월 ${w.n}주차` : `${w.n}주차`;
    };
  }

  // 필터/집계에서 같은 날짜로 수십만 번 부르므로 날짜별로 기억해 둔다.
  const weekOfCache = new Map();
  function weekKeyOf(dateStr) {
    let w = weekOfCache.get(dateStr);
    if (!w) {
      w = weekOf(dateStr);
      weekOfCache.set(dateStr, w);
    }
    return w.key;
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

  // 년도/월/주차/구분(CE·MX)/광고상품유형/품목/상세 품목 슬라이서 선택 상태에
  // 이 행이 걸리는지. 어떤 슬라이서도 선택 안 했으면(빈 Set) 그 기준은
  // 통과시킨다(=전체 보기). 구분/품목/상세 품목은 쿠팡 전용이라 네이버에서는
  // 값이 항상 null이고, 네이버 쪽 슬라이서도 항상 비어있어 걸리지 않는다.
  // skip에 기준 이름을 주면 그 기준만 빼고 검사한다(필터 버튼 목록 계산용).
  function passesFilters(platform, date, productType, itemCategory, subItem, skip) {
    const f = state.filters[platform];
    if (skip !== "years" && f.years.size && !f.years.has(yearOf(date))) return false;
    if (skip !== "months" && f.months.size && !f.months.has(monthOf(date))) return false;
    if (skip !== "weeks" && f.weeks.size && !f.weeks.has(weekKeyOf(date))) return false;
    if (skip !== "divisions" && f.divisions.size && !f.divisions.has(classifyDivision(itemCategory))) return false;
    if (skip !== "productTypes" && f.productTypes.size && !f.productTypes.has(productType)) return false;
    if (skip !== "itemCategories" && f.itemCategories.size && !f.itemCategories.has(itemCategory)) return false;
    if (skip !== "subItems" && f.subItems.size && !f.subItems.has(subItem)) return false;
    return true;
  }

  // 필터 버튼 목록. 버튼 위치가 클릭할 때마다 움직이지 않도록 버튼은 항상
  // 전부 보여주고(all), "나머지 기준들"을 모두 통과하는 데이터가 있는 값만
  // 누를 수 있게 한다(available - 나머지는 흐리게, 예: PA를 고르면 PA 데이터가
  // 없는 품목은 흐려짐). 상세 품목은 고른 품목의 상세 품목만 보여준다.
  const FILTER_DIMENSIONS = ["years", "months", "weeks", "divisions", "productTypes", "itemCategories", "subItems"];

  function computeFilterOptions(platform) {
    const f = state.filters[platform];
    // 품목 선택이 바뀌어 더 이상 해당 없는 상세 품목 선택은 지운다(안 보이는
    // 필터가 계속 걸려 있지 않도록).
    if (!f.itemCategories.size) f.subItems.clear();
    const opts = {};
    const all = {};
    for (const dim of FILTER_DIMENSIONS) {
      opts[dim] = new Set(f[dim]);
      all[dim] = new Set(f[dim]);
    }
    const subItemsOfSelected = new Set();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const date = row.__date;
      if (!date) continue;
      const item = row.__item;
      const subItem = row.__subItem;
      const division = classifyDivision(item);
      const values = {
        years: yearOf(date),
        months: monthOf(date),
        weeks: weekKeyOf(date),
        divisions: division,
        productTypes: row.__productType,
        itemCategories: platform === "coupang" ? item : null,
        subItems: platform === "coupang" ? subItem : null,
      };
      if (platform === "coupang" && f.itemCategories.has(item) && subItem) subItemsOfSelected.add(subItem);
      for (const dim of FILTER_DIMENSIONS) {
        const v = values[dim];
        if (v == null) continue;
        if (dim === "subItems" && !f.itemCategories.has(item)) continue;
        all[dim].add(v);
        if (opts[dim].has(v)) continue;
        if (passesFilters(platform, date, row.__productType, item, subItem, dim)) opts[dim].add(v);
      }
    }
    for (const v of [...f.subItems]) {
      if (!subItemsOfSelected.has(v)) {
        f.subItems.delete(v);
        opts.subItems.delete(v);
        all.subItems.delete(v);
      }
    }
    const ko = (a, b) => a.localeCompare(b, "ko");
    const sorted = (dim) => [...all[dim]].sort(dim === "itemCategories" || dim === "subItems" ? ko : undefined);
    const result = { available: opts };
    for (const dim of FILTER_DIMENSIONS) result[dim] = sorted(dim);
    return result;
  }

  function formatMetric(value, format) {
    if (value == null || !Number.isFinite(value)) return "-";
    if (format === "pct") return `${value.toLocaleString("ko-KR", { maximumFractionDigits: 2 })}%`;
    if (format === "pct0") return `${Math.round(value).toLocaleString("ko-KR")}%`;
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
  // fetchAllDataRows()/withFullCoupangRows()가 state.colsByFile을 다시
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
  // 단, 그룹명이 인덱스에서 서로 다른 품목에 걸쳐 있으면(예: 쿠팡 기본값
  // "새 광고 그룹"은 TV/청소기/냉장고 등 여러 캠페인에 공통) 그 그룹명만으로는
  // 품목을 알 수 없으므로 조회 키에서 뺀다. 그래야 인덱스에 없는 새 캠페인이
  // 마지막 행의 엉뚱한 품목으로 가지 않고, 다음 기준(VIID)으로 찾거나 그래도
  // 없으면 "미분류"로 보인다. 품목은 같고 세부품목만 다른 그룹명은 그대로 쓴다.
  function buildIndexMaps(sheet) {
    const byCampaignId = new Map();
    const byCampaignName = new Map();
    const byGroupName = new Map();
    const byViid = new Map();
    const groupNameItems = new Map(); // 그룹명 -> 그 그룹명이 쓰인 품목들
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
        if (groupName) {
          byGroupName.set(groupName, entry);
          if (!groupNameItems.has(groupName)) groupNameItems.set(groupName, new Set());
          groupNameItems.get(groupName).add(item);
        }
      }
      const vid = normalizeId(sheetCellText(sheet, "I", r)); // I열 VIID
      const rightItem = sheetCellText(sheet, "K", r); // K열 품목
      if (vid && rightItem) {
        const rightSubItem = sheetCellText(sheet, "L", r) || rightItem; // L열 세부품목
        byViid.set(vid, { item: rightItem, subItem: rightSubItem });
      }
    }
    for (const [groupName, items] of groupNameItems) {
      if (items.size > 1) byGroupName.delete(groupName);
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
  // 의류케어 품목 중 아래 광고집행 상품은 세부품목을 상품 단위로 다시 나눈다
  // (캠페인/VIID 기준 INDEX만으로는 세탁기/건조기/세건패키지를 못 가른다).
  // 상품명 전체가 같거나, 상품명 뒤 칸의 모델코드가 같으면 해당 세부품목.
  const LAUNDRY_SUBITEM_PRODUCTS = {
    세탁기: [
      "삼성전자 그랑데 드럼세탁기 21kg 방문설치,WF21T6000KV,WF21T6000KV,블랙 케비어,블랙 케비어",
      "삼성전자 그랑데 드럼 세탁기 19kg 방문설치,WF19T6000KW,WF19T6000KW,화이트,화이트",
      "삼성전자 그랑데 통버블 세탁기 16kg 방문설치,WA16CG6441BY,WA16CG6441BY,라벤더그레이,라벤더그레이",
      "삼성전자 그랑데 통버블 세탁기 10kg 방문설치,WA10CG5441BW,WA10CG5441BW,화이트,화이트",
      "삼성전자 그랑데 통버블 세탁기 21kg 방문설치,WA21A8376KV,WA21A8376KV,블랙 케비어,블랙 케비어",
      "삼성전자 AI 세탁기 25kg 방문설치,WF25DG8250BV,WF25DG8250BV,블랙캐비어,블랙캐비어",
      "삼성전자 AI 세탁기 25kg 방문설치,WF25DG8250BW,WF25DG8250BW,화이트,화이트",
      "삼성전자 AI 통버블 일반 세탁기 19kg 방문설치,WA80F19E8W,WA80F19E8W,화이트,화이트",
      "삼성전자 그랑데 AI 세탁기 12kg 방문설치,WW12T504DTW,WW12T504DTW,화이트,화이트",
    ],
    건조기: [
      "삼성전자 AI 건조기 21kg 방문설치,DV21DG8200BV,DV21DG8200BV,블랙캐비어,블랙캐비어",
      "삼성전자 AI 건조기 21kg 방문설치,DV21DG8200BW,DV21DG8200BW,화이트,화이트",
      "삼성전자 건조기 9kg 방문설치,DV90TA040TE,DV90TA040TE,화이트,화이트",
      "삼성전자 인버터 건조기 9kg 방문설치,DV90T5540BV,DV90T5540BV,블랙케비어,블랙케비어",
    ],
    세건패키지: [
      "삼성전자 BESPOKE 그랑데 AI 원바디 Top Fit 세탁기 + 건조기 WF2420HCWWC 24kg + 20kg 방문설치,WF2420HCWWC,WF2420HCWWC,화이트,화이트",
      "삼성전자 BESPOKE 그랑데 AI 원바디 Top Fit 세탁기 + 건조기 WF2420HCVVC 24kg+20kg 방문설치,WF2420HCVVC,WF2420HCVVC,블랙캐비어,블랙캐비어",
      "삼성전자 비스포크 그랑데 AI 세탁기 WF21CB6650BW + 건조기 DV17CB6600BW 21kg + 17kg 방문설치,WF21CB6650BW(세탁기) + DV17CB6600BW(건조기),WF21CB6650BW(세탁기) + DV17CB6600BW(건조기),화이트,화이트",
      "삼성전자 AI 세탁기 WF25DG8250BW + 건조기 DV21DG8200BW 25kg + 21kg + 직렬설치키트 SKK-ATW 세트 방문설치,WF25DG8250BW2T,WF25DG8250BW2T,화이트,화이트",
      "삼성전자 AI 세탁기 WF25DG8250BV + 건조기 DV21DG8200BV 25kg + 21kg + 직렬설치키트 SKK-ATV 세트 방문설치,WF25DG8250BV2T,WF25DG8250BV2T,블랙캐비어,블랙캐비어",
      "삼성전자 Bespoke AI 콤보 LCD 드럼 세탁건조기 25kg / 18kg 71.1mm 방문설치,화이트,화이트,WD80H25BHW,WD80H25BHW",
      "삼성전자 Bespoke AI 콤보 LCD 드럼 세탁건조기 25kg / 20kg 177.8mm 방문설치,그레이지,그레이지,WD90H25BHY,WD90H25BHY",
      "삼성전자 Bespoke AI 콤보 LCD 드럼 세탁건조기 25kg / 20kg 177.8mm 방문설치,화이트,화이트,WD90H25BHW,WD90H25BHW",
    ],
    에어드레서: [
      "삼성전자 Bespoke 에어드레서 일반용량 상의 3~5벌 + 하의 1벌 DF18CB8600ER 방문설치,DF18CB8600ER,DF18CB8600ER,새틴 베이지,새틴 베이지",
    ],
  };
  const normalizeProductText = (v) => String(v ?? "").replace(/\s+/g, "").toUpperCase();
  // 상품명 뒤 칸들 중 모델코드처럼 생긴 값(영문+숫자, 한글 없음)만 모델 기준
  // 매칭에 쓴다. 콤보 세탁건조기처럼 상품명 다음 칸이 색상("화이트")인
  // 상품도 있어서, 색상끼리 잘못 매칭되지 않게 하기 위함.
  function productModelCodes(name) {
    return String(name ?? "")
      .split(",")
      .slice(1)
      .map(normalizeProductText)
      .filter((c) => /[A-Z]/.test(c) && /\d/.test(c) && !/[가-힣]/.test(c));
  }
  const LAUNDRY_SUBITEM_BY_NAME = new Map();
  const LAUNDRY_SUBITEM_BY_MODEL = new Map();
  for (const [subItem, names] of Object.entries(LAUNDRY_SUBITEM_PRODUCTS)) {
    for (const name of names) {
      LAUNDRY_SUBITEM_BY_NAME.set(normalizeProductText(name), subItem);
      for (const code of productModelCodes(name)) LAUNDRY_SUBITEM_BY_MODEL.set(code, subItem);
    }
  }
  function laundrySubItemByProduct(execProductName) {
    const s = String(execProductName ?? "");
    if (!s) return null;
    const byName = LAUNDRY_SUBITEM_BY_NAME.get(normalizeProductText(s));
    if (byName) return byName;
    for (const code of productModelCodes(s)) {
      const byModel = LAUNDRY_SUBITEM_BY_MODEL.get(code);
      if (byModel) return byModel;
    }
    return null;
  }

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
        if (result.item === "의류케어") {
          const execCol = resolveField(fp, platform, "execProductName");
          const refined = laundrySubItemByProduct(execCol ? row[execCol] : null);
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
    const banner = el("loadingBanner");
    try {
      indexMapsReadyPromise = loadIndexMaps();
      const [, , allRows] = await Promise.all([indexMapsReadyPromise, loadGfaIndexMaps(), fetchAllDataRows()]);
      state.rows = allRows;
      renderAll();
      if (banner) banner.hidden = true;
    } catch (e) {
      // 여기서 예외가 나면 배너가 계속 "불러오는 중"으로 남아 무한로딩처럼
      // 보이므로, 실패했다는 걸 화면에 분명히 보여준다.
      console.error("대시보드를 불러오지 못했습니다:", e);
      if (banner) banner.textContent = "데이터를 불러오는 중 문제가 생겼습니다. 새로고침해 주세요.";
    }
  }

  // 쿠팡 PA 키워드 리포트처럼 "키워드" 컬럼이 있는 파일만, 날짜+캠페인ID+
  // 캠페인명+그룹명+VIID+광고집행상품명 단위로(즉
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
    const metricCols = {};
    for (const f of DASHBOARD_METRIC_FIELDS) metricCols[f] = resolveFieldForCols(cols, platform, f);
    const groups = new Map();
    for (const row of rows) {
      const key = [
        row.__date,
        ...Object.values(keyCols).map((c) => (c ? row[c] : "")),
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
  // scripts/build_dashboard_cache.py가 만든 캐시 JSON(f.cache = 키워드를
  // 미리 합친 대시보드용, f.fullCache = 행은 원본 그대로에 다운로드용
  // 컬럼만 남긴 것)을 행 객체 배열로 푼다. 원본 파일이 바뀌었는데 캐시가
  // 아직 안 갱신된 경우(크기가 다름)는 null을 돌려서 원본을 읽게 한다.
  async function loadCacheRows(url, f) {
    const res = await fetch(url);
    if (!res.ok) return null;
    const cache = await res.json();
    if (cache.v !== 2 || cache.sourceSize !== f.size) return null;
    const { groupColumns, metricColumns, dicts } = cache;
    const g = groupColumns.length;
    return cache.rows.map((r) => {
      const row = {};
      for (let j = 0; j < g; j++) row[groupColumns[j]] = dicts[j][r[j]];
      for (let j = 0; j < metricColumns.length; j++) row[metricColumns[j]] = r[g + j];
      return row;
    });
  }

  async function loadAndTagFile(f, aggregateKeywords = true) {
    let rows = null;
    // 쿠팡 키워드 리포트는 브라우저에서 원본 xlsb를 직접 해석하지 않고
    // Actions가 미리 만들어 둔 캐시를 쓴다. 수십만 행짜리 xlsb는 모바일에선
    // 메모리 부족으로, PC에서도 SheetJS가 수 분씩 멈추거나 데이터를 놓쳐서
    // 제대로 읽히지 않았다. 대시보드는 키워드를 합친 캐시, 다운로드는
    // 행을 그대로 둔 캐시를 쓴다.
    const cacheUrl = aggregateKeywords ? f.cache : f.fullCache;
    if (cacheUrl) {
      try {
        rows = await loadCacheRows(cacheUrl, f);
      } catch (e) {
        console.error("캐시를 불러오지 못해 원본을 읽습니다:", cacheUrl, e);
      }
    }
    if (!rows) rows = await parseDataFile(f.path);
    if (rows.length === 0) return [];
    const productType = detectProductType(f.platform, f.name);
    const cols = new Set(Object.keys(rows[0]));
    state.colsByFile.set(f.path, cols);
    const dateCol = resolveFieldForCols(cols, f.platform, "date");
    // 품목 분류는 INDEX 조회표가 준비돼야 한다. raw 파일 fetch가 INDEX
    // fetch보다 먼저 끝나는 경우(특히 withFullCoupangRows처럼 이미 로드된
    // 뒤에 다시 부를 때는 즉시 반환됨)를 대비해 분류 직전에 기다린다.
    if (f.platform === "coupang" && indexMapsReadyPromise) await indexMapsReadyPromise;
    // 방금 읽은 행 객체를 복사하지 않고 그대로 태그를 붙인다(수십만 행을
    // 복사하면 메모리가 두 배로 든다).
    const tagged = rows;
    for (const out of tagged) {
      out.__platform = f.platform;
      out.__file = f.name;
      out.__filePath = f.path;
      out.__productType = productType;
      out.__date = dateCol ? parseDateLoose(out[dateCol]) : null;
      if (f.platform === "coupang") {
        const cls = resolveItemClassification(out, f.path, f.platform, productType);
        out.__item = cls.item;
        out.__subItem = cls.subItem;
      } else {
        out.__item = null;
        out.__subItem = null;
      }
    }
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
  // 클릭했을 때만 쿠팡 파일들의 다운로드용 캐시(키워드를 합치지 않은 것)를
  // 받아서 build(rows)에 넘긴다. 그동안 state.colsByFile은 다운로드용 행의
  // 컬럼으로 바뀌고(그래야 "키워드" 등이 잡힘), 끝나면 대시보드용으로
  // 되돌린다.
  async function withFullCoupangRows(build) {
    const savedCols = new Map(state.colsByFile);
    fieldResolutionCache = new Map();
    try {
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
      return build(perFileRows.flat());
    } finally {
      state.colsByFile = savedCols;
      fieldResolutionCache = new Map();
    }
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
      if (!date || !passesFilters(platform, date, productType, itemCategory, row.__subItem)) continue;
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
      if (!date || !passesFilters(platform, date, productType, itemCategory, row.__subItem)) continue;
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
      if (!date || !passesFilters(platform, date, productType, itemCategory, row.__subItem)) continue;
      const week = weekOf(date);
      accumulateRow(groups, week.key, { weekKey: week.key, weekLabel: week.label }, row, fp, platform);
    }
    // 필터에 걸린 주차를 모두 보여준다(최근 2개 주차의 증감율은 맨 아래 줄).
    const rows = finalizeGroups(groups).sort((a, b) => a.weekKey.localeCompare(b.weekKey));
    const name = weekNameFormatter(rows.map((r) => r.weekKey));
    for (const r of rows) r.weekLabel = `${name(r.weekKey)} (${r.weekLabel})`;
    return rows;
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
      if (!date || !passesFilters(platform, date, productType, itemCategory, row.__subItem)) continue;
      const campaignCol = resolveField(fp, platform, "campaign");
      const groupCol = resolveField(fp, platform, "group");
      const campaign = campaignCol ? row[campaignCol] ?? "" : "";
      const group = groupCol ? row[groupCol] ?? "" : "";
      const productName = resolveProductLabel(row, fp, platform, productType, campaign, group);
      const key = `${productType}\u0001${productName}`;
      accumulateRow(groups, key, { productType, productName, viids: new Set() }, row, fp, platform);
      if (platform === "coupang") {
        const viid = productViid(row, fp);
        if (viid) groups.get(key).viids.add(viid);
      }
    }
    const all = finalizeGroups(groups).sort((a, b) => (b.revenue || 0) - (a.revenue || 0));
    // 원그래프의 "기타"와 비중 계산용 - 필터에 걸린 전체 상품의 매출 합계
    const totalRevenue = all.reduce((sum, r) => sum + (r.revenue || 0), 0);
    const rows = all.slice(0, 10).map((r, i) => {
      const ids = [...r.viids];
      return { ...r, rank: i + 1, viid: ids.length ? ids[0] + (ids.length > 1 ? ` 외 ${ids.length - 1}` : "") : "-" };
    });
    return { rows, totalRevenue };
  }

  // TOP10 상품의 VIID: 광고집행 옵션ID(PA/BPA). NCA처럼 없으면 null.
  function productViid(row, fp) {
    const viidCol = resolveField(fp, "coupang", "viid");
    return viidCol ? normalizeId(row[viidCol]) || null : null;
  }

  // ---------- 렌더링 ----------

  function renderAll() {
    renderTabs();
    renderFileNotice();
    renderSlicers();
    renderDailyChart();
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

  // available(Set)에 없는 값은 지금 다른 필터 조건에서 데이터가 없다는 뜻이라
  // 흐리게 하고 못 누르게 한다(이미 선택한 값은 끌 수 있게 그대로 둠).
  function renderSlicerGroup(containerId, values, selectedSet, formatLabel, available, emptyText = "-") {
    const area = el(containerId);
    area.innerHTML = "";
    if (values.length === 0) {
      area.innerHTML = `<span class="empty-state">${escapeHtml(emptyText)}</span>`;
      return;
    }
    for (const v of values) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "slicer-btn" + (selectedSet.has(v) ? " active" : "");
      if (available && !available.has(v) && !selectedSet.has(v)) {
        btn.disabled = true;
        btn.classList.add("unavailable");
        btn.title = "지금 고른 다른 필터 조건에는 데이터가 없습니다";
      }
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
    const av = options.available;
    const f = state.filters[platform];
    renderSlicerGroup("yearSlicer", options.years, f.years, (y) => `${y}년`, av.years);
    renderSlicerGroup("monthSlicer", options.months, f.months, (m) => `${Number(m)}월`, av.months);
    renderSlicerGroup("weekSlicer", options.weeks, f.weeks, weekNameFormatter(options.weeks), av.weeks);
    // 버튼에 마우스를 올리면 실제 기간(월~일)이 보이게
    for (const btn of el("weekSlicer").querySelectorAll(".slicer-btn")) {
      const i = [...btn.parentNode.children].indexOf(btn);
      if (!btn.disabled) btn.title = weekOf(options.weeks[i]).label;
    }
    renderSlicerGroup("divisionSlicer", options.divisions, f.divisions, null, av.divisions);
    renderSlicerGroup("productTypeSlicer", options.productTypes, f.productTypes, null, av.productTypes);
    renderSlicerGroup("itemCategorySlicer", options.itemCategories, f.itemCategories, null, av.itemCategories);
    renderSlicerGroup("subItemSlicer", options.subItems, f.subItems, null, av.subItems, "품목을 고르면 상세 품목이 나옵니다");
    // 구분/품목/상세 품목은 쿠팡 전용이라 네이버 탭에서는 줄 자체를 숨긴다.
    for (const id of ["divisionFilterGroup", "itemFilterGroup", "subItemFilterGroup"]) {
      el(id).hidden = platform !== "coupang";
    }
    const selectedCount = FILTER_DIMENSIONS.reduce((n, dim) => n + f[dim].size, 0);
    const reset = el("filterResetBtn");
    reset.disabled = selectedCount === 0;
    reset.textContent = selectedCount ? `필터 초기화 (${selectedCount})` : "필터 초기화";
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

  const TEXT_COLUMN_KEYS = new Set(["rank", "productType", "productName", "viid", "campaign", "group", "weekLabel", "itemCategory"]);

  function cellText(row, col) {
    if (TEXT_COLUMN_KEYS.has(col.key)) {
      return row[col.key] ?? "";
    }
    return formatMetric(row[col.key], col.format);
  }

  // 표에 보이는 행들의 합계(광고비/노출/클릭/전환수/전환매출은 더하고,
  // CTR/CPC/CVR/객단가/ROAS는 그 합계로 다시 계산).
  function totalOf(rows) {
    const t = { cost: 0, impressions: 0, clicks: 0, conversions: 0, revenue: 0 };
    for (const r of rows) {
      for (const k of Object.keys(t)) if (Number.isFinite(r[k])) t[k] += r[k];
    }
    return finalizeGroups(new Map([["total", t]]))[0];
  }

  // 표 맨 위(컬럼 이름 바로 아래)에 붙는 합계 행
  function appendTotalRow(bodyEl, columns, rows) {
    const total = totalOf(rows);
    const tr = document.createElement("tr");
    tr.className = "total-row";
    columns.forEach((col, i) => {
      const td = document.createElement("td");
      if (i === 0) td.textContent = "합계";
      else if (!TEXT_COLUMN_KEYS.has(col.key)) td.textContent = formatMetric(total[col.key], col.format);
      tr.appendChild(td);
    });
    bodyEl.appendChild(tr);
  }

  // options.totalRows: 합계에 쓸 행(없으면 rows). options.decorate(td, row, col):
  // 칸을 꾸밀 때(예: 순위 옆 색 표시).
  function renderMetricTable(headEl, bodyEl, columns, rows, emptyMessage, options = {}) {
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
    appendTotalRow(bodyEl, columns, options.totalRows || rows);
    for (const row of rows) {
      const tr2 = document.createElement("tr");
      for (const col of columns) {
        const td = document.createElement("td");
        const text = String(cellText(row, col));
        if (col.maxLen && text.length > col.maxLen) {
          td.textContent = `${text.slice(0, col.maxLen)}…`;
          td.title = text;
        } else {
          td.textContent = text;
        }
        if (options.decorate) options.decorate(td, row, col);
        tr2.appendChild(td);
      }
      bodyEl.appendChild(tr2);
    }
  }

  function renderTable1() {
    const rows = computeTable1(state.activePlatform);
    el("table1Title").textContent = state.activePlatform === "coupang" ? "품목별 데이터" : "캠페인별 데이터";
    const totalPages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    state.table1Page = Math.min(Math.max(1, state.table1Page), totalPages);
    const pageRows = rows.slice((state.table1Page - 1) * PAGE_SIZE, state.table1Page * PAGE_SIZE);
    renderMetricTable(el("table1Head"), el("table1Body"), table1ColsFor(state.activePlatform), pageRows, "데이터가 없습니다.", { totalRows: rows });
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
  // 증가는 빨간색(▲), 하락은 파란색(▼).
  function appendWeekCompareRow(bodyEl, rows) {
    if (rows.length < 2) return;
    const prev = rows[rows.length - 2];
    const latest = rows[rows.length - 1];
    const tr = document.createElement("tr");
    tr.className = "week-compare-row";
    const labelTd = document.createElement("td");
    labelTd.textContent = "최근 2개 주차 증감율";
    tr.appendChild(labelTd);
    for (const col of METRIC_COLUMNS) {
      const td = document.createElement("td");
      const a = prev[col.key];
      const b = latest[col.key];
      if (!Number.isFinite(a) || a === 0 || !Number.isFinite(b)) {
        td.textContent = "-";
      } else {
        const pct = (b / a - 1) * 100;
        const text = `${Math.abs(pct).toLocaleString("ko-KR", { maximumFractionDigits: 1 })}%`;
        if (pct > 0) {
          td.textContent = `▲ ${text}`;
          td.className = "delta-up";
        } else if (pct < 0) {
          td.textContent = `▼ ${text}`;
          td.className = "delta-down";
        } else {
          td.textContent = "0%";
        }
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

  // 원그래프 조각 색: 순위가 높을수록 진한 한 가지 색(보라) 계열, 기타는 회색.
  function rankColor(i) {
    return `color-mix(in oklab, var(--chart-bar) ${100 - i * 6}%, var(--surface-1))`;
  }

  function renderTable4() {
    const platform = state.activePlatform;
    const { rows, totalRevenue } = computeTable4(platform);
    renderMetricTable(el("table4Head"), el("table4Body"), table4ColsFor(platform), rows, "판매 데이터가 없습니다.", {
      decorate: (td, row, col) => {
        if (col.key !== "rank") return;
        const sw = document.createElement("span");
        sw.className = "rank-swatch";
        sw.style.background = rankColor(row.rank - 1);
        td.prepend(sw);
      },
    });
    renderTop10Pie(rows, totalRevenue);
  }

  // 판매 우수 상품 TOP 10의 매출 비중 원그래프(도넛). 비중 = 각 상품 매출 /
  // 필터에 걸린 전체 상품 매출. TOP 10 밖의 매출은 "기타"(회색)로 묶는다.
  function renderTop10Pie(rows, totalRevenue) {
    const area = el("top10Pie");
    const legend = el("top10PieLegend");
    const tooltip = el("top10PieTooltip");
    area.innerHTML = "";
    legend.innerHTML = "";
    tooltip.hidden = true;
    if (!rows.length || !(totalRevenue > 0)) {
      area.innerHTML = `<div class="empty-state">데이터가 없습니다.</div>`;
      return;
    }
    const top = rows.reduce((sum, r) => sum + (r.revenue || 0), 0);
    const slices = rows.map((r, i) => ({ label: `${r.rank}위 ${r.productName}`, value: r.revenue || 0, color: rankColor(i), rank: r.rank }));
    if (totalRevenue - top > 0) slices.push({ label: "기타 (TOP 10 외)", value: totalRevenue - top, color: "var(--pie-other)", rank: null });
    const size = 220;
    const R = size / 2 - 4;
    const r0 = R * 0.6;
    const c = size / 2;
    const svg = svgEl("svg", { width: size, height: size, viewBox: `0 0 ${size} ${size}`, role: "img", "aria-label": "TOP 10 상품 매출 비중" });
    let angle = -Math.PI / 2;
    const pt = (rad, a) => `${c + rad * Math.cos(a)},${c + rad * Math.sin(a)}`;
    const pct = (v) => `${((v / totalRevenue) * 100).toLocaleString("ko-KR", { maximumFractionDigits: 1 })}%`;
    for (const sl of slices) {
      const sweep = (sl.value / totalRevenue) * Math.PI * 2;
      if (sweep <= 0) continue;
      const a1 = angle + Math.min(sweep, Math.PI * 2 - 1e-4);
      const large = sweep > Math.PI ? 1 : 0;
      const d = `M${pt(R, angle)}A${R},${R} 0 ${large} 1 ${pt(R, a1)}L${pt(r0, a1)}A${r0},${r0} 0 ${large} 0 ${pt(r0, angle)}Z`;
      const path = svgEl("path", { d, class: "pie-slice" });
      path.style.fill = sl.color;
      const show = (e) => {
        tooltip.innerHTML = `<div class="tt-title">${escapeHtml(sl.label.length > 34 ? sl.label.slice(0, 34) + "…" : sl.label)}</div>` +
          `<div>매출 <b>${formatMetric(sl.value, "money")}</b></div><div class="tt-muted">비중 ${pct(sl.value)}</div>`;
        tooltip.hidden = false;
        const box = area.getBoundingClientRect();
        tooltip.style.left = `${Math.min(e.clientX - box.left + 12, box.width - tooltip.offsetWidth)}px`;
        tooltip.style.top = `${e.clientY - box.top + 12}px`;
      };
      path.addEventListener("pointermove", show);
      path.addEventListener("pointerdown", show);
      path.addEventListener("pointerleave", () => (tooltip.hidden = true));
      svg.appendChild(path);
      angle += sweep;
    }
    const t1 = svgEl("text", { x: c, y: c - 6, "text-anchor": "middle", class: "pie-center-label" });
    t1.textContent = "TOP 10 비중";
    const t2 = svgEl("text", { x: c, y: c + 18, "text-anchor": "middle", class: "pie-center-value" });
    t2.textContent = pct(top);
    svg.appendChild(t1);
    svg.appendChild(t2);
    area.appendChild(svg);
    legend.innerHTML = slices
      .map((sl) => `<li><span class="rank-swatch" style="background:${sl.color}"></span><span class="legend-name">${sl.rank ? `${sl.rank}위` : "기타"}</span><span class="legend-pct">${pct(sl.value)}</span></li>`)
      .join("");
  }

  // ---------- 일자별 광고비·ROAS 그래프 ----------
  // 지금 필터 기준으로 일자별 광고비(막대, 왼쪽 축 = 금액)와 ROAS(선, 오른쪽
  // 축 = %)를 그린다. 단위가 달라 축을 나누고, 두 축 모두 4칸으로 맞춰
  // 눈금선을 공유한다. 라이브러리 없이 SVG로 직접 그린다.
  function computeDaily(platform) {
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const date = row.__date;
      if (!date || !passesFilters(platform, date, row.__productType, row.__item, row.__subItem)) continue;
      accumulateRow(groups, date, { date }, row, row.__filePath, platform);
    }
    return finalizeGroups(groups).sort((a, b) => a.date.localeCompare(b.date));
  }

  // 축 눈금용 짧은 금액 표기 (예: 1.2억, 3,000만)
  function shortMoney(v) {
    if (v === 0) return "0";
    if (Math.abs(v) >= 1e8) return `${(v / 1e8).toLocaleString("ko-KR", { maximumFractionDigits: 1 })}억`;
    if (Math.abs(v) >= 1e4) return `${Math.round(v / 1e4).toLocaleString("ko-KR")}만`;
    return Math.round(v).toLocaleString("ko-KR");
  }

  function niceStep(max, count) {
    const raw = max / count;
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const n = raw / mag;
    return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * mag;
  }

  const SVG_NS = "http://www.w3.org/2000/svg";
  function svgEl(tag, attrs) {
    const node = document.createElementNS(SVG_NS, tag);
    for (const k of Object.keys(attrs)) node.setAttribute(k, attrs[k]);
    return node;
  }

  // 위쪽 두 모서리만 둥근 막대 (바닥은 기준선에 딱 붙임)
  function barPath(x, y, w, h, r) {
    const rr = Math.min(r, w / 2, h);
    return `M${x},${y + h}V${y + rr}Q${x},${y} ${x + rr},${y}H${x + w - rr}Q${x + w},${y} ${x + w},${y + rr}V${y + h}Z`;
  }

  function renderDailyChart() {
    const area = el("dailyChart");
    const tooltip = el("chartTooltip");
    if (!area) return;
    const rows = computeDaily(state.activePlatform);
    area.innerHTML = "";
    tooltip.hidden = true;
    if (rows.length === 0) {
      area.innerHTML = `<div class="empty-state">데이터가 없습니다.</div>`;
      return;
    }
    const width = Math.max(300, area.clientWidth || 800);
    const height = 280;
    const TICKS = 4;
    const costStep = niceStep(Math.max(1, ...rows.map((r) => r.cost || 0)), TICKS);
    const roasStep = niceStep(Math.max(1, ...rows.map((r) => (Number.isFinite(r.roas) ? r.roas : 0))), TICKS);
    const roasLabel = (v) => `${Math.round(v).toLocaleString("ko-KR")}%`;
    // 축 숫자가 잘리지 않게, 가장 긴 눈금 글자 길이만큼 좌우 여백을 잡는다.
    const labelW = (text) => 12 + String(text).length * 7;
    const m = {
      top: 44,
      right: Math.max(44, labelW(roasLabel(roasStep * TICKS))),
      bottom: 28,
      left: Math.max(44, labelW(shortMoney(costStep * TICKS))),
    };
    const plotW = width - m.left - m.right;
    const plotH = height - m.top - m.bottom;
    const yCost = (v) => m.top + plotH - (v / (costStep * TICKS)) * plotH;
    const yRoas = (v) => m.top + plotH - (v / (roasStep * TICKS)) * plotH;
    const band = plotW / rows.length;
    const barW = Math.max(2, Math.min(24, band * 0.6));
    const cx = (i) => m.left + band * i + band / 2;

    const svg = svgEl("svg", { width, height, viewBox: `0 0 ${width} ${height}`, role: "img", "aria-label": "일자별 광고비와 ROAS 추이" });

    // 막대는 위에서 아래로 옅어지는 보라, ROAS 선 아래는 은은한 민트 음영
    const defs = svgEl("defs", {});
    const grad = (id, color, from, to) => {
      const g = svgEl("linearGradient", { id, x1: 0, y1: 0, x2: 0, y2: 1 });
      const s1 = svgEl("stop", { offset: "0%" });
      s1.style.stopColor = `var(${color})`;
      s1.style.stopOpacity = from;
      const s2 = svgEl("stop", { offset: "100%" });
      s2.style.stopColor = `var(${color})`;
      s2.style.stopOpacity = to;
      g.appendChild(s1);
      g.appendChild(s2);
      defs.appendChild(g);
    };
    grad("dailyBarGrad", "--chart-bar", 1, 0.55);
    grad("dailyAreaGrad", "--chart-line", 0.22, 0);
    svg.appendChild(defs);

    // 눈금선 + 왼쪽(광고비 금액)/오른쪽(ROAS %)
    for (let k = 0; k <= TICKS; k++) {
      const yy = yCost(costStep * k);
      svg.appendChild(svgEl("line", { x1: m.left, x2: width - m.right, y1: yy, y2: yy, class: k === 0 ? "chart-baseline" : "chart-grid" }));
      const left = svgEl("text", { x: m.left - 8, y: yy + 4, "text-anchor": "end", class: "chart-axis-label" });
      left.textContent = shortMoney(costStep * k);
      svg.appendChild(left);
      const right = svgEl("text", { x: width - m.right + 8, y: yy + 4, "text-anchor": "start", class: "chart-axis-label" });
      right.textContent = roasLabel(roasStep * k);
      svg.appendChild(right);
    }
    const leftTitle = svgEl("text", { x: m.left - 8, y: 12, "text-anchor": "end", class: "chart-axis-title" });
    leftTitle.textContent = "광고비";
    svg.appendChild(leftTitle);
    const rightTitle = svgEl("text", { x: width - m.right + 8, y: 12, "text-anchor": "start", class: "chart-axis-title" });
    rightTitle.textContent = "ROAS";
    svg.appendChild(rightTitle);

    // x축 날짜 (겹치지 않게 일정 간격으로만 표시)
    const every = Math.max(1, Math.ceil(44 / band));
    rows.forEach((r, i) => {
      if (i % every !== 0) return;
      const t = svgEl("text", { x: cx(i), y: height - 8, "text-anchor": "middle", class: "chart-axis-label" });
      t.textContent = `${r.date.slice(5, 7)}.${r.date.slice(8, 10)}`;
      svg.appendChild(t);
    });

    // 주차 구분 점선 + 주차 이름 (월요일이 시작되는 곳마다)
    const weekKeys = [...new Set(rows.map((r) => weekKeyOf(r.date)))];
    const weekName = weekNameFormatter(weekKeys);
    let segStart = 0;
    const weekLabel = (from, to) => {
      const x1 = m.left + band * from;
      const x2 = m.left + band * to;
      if (x2 - x1 < 32) return; // 이름이 안 들어갈 만큼 짧은 주(데이터 1~2일)는 생략
      const t = svgEl("text", { x: (x1 + x2) / 2, y: m.top - 10, "text-anchor": "middle", class: "chart-week-label" });
      t.textContent = weekName(weekKeyOf(rows[from].date));
      svg.appendChild(t);
    };
    rows.forEach((r, i) => {
      if (i === 0 || weekKeyOf(r.date) === weekKeyOf(rows[i - 1].date)) return;
      const x = m.left + band * i;
      svg.appendChild(svgEl("line", { x1: x, x2: x, y1: m.top - 22, y2: m.top + plotH, class: "chart-week-sep" }));
      weekLabel(segStart, i);
      segStart = i;
    });
    weekLabel(segStart, rows.length);

    // 호버 시 해당 일자 강조 배경
    const hoverBand = svgEl("rect", { x: 0, y: m.top, width: band, height: plotH, class: "chart-hover-band", visibility: "hidden" });
    svg.appendChild(hoverBand);

    // 광고비 막대
    rows.forEach((r, i) => {
      const v = r.cost || 0;
      if (v <= 0) return;
      const top = yCost(v);
      svg.appendChild(svgEl("path", { d: barPath(cx(i) - barW / 2, top, barW, m.top + plotH - top, 4), class: "chart-bar", fill: "url(#dailyBarGrad)" }));
    });

    // ROAS 선 + 아래 음영 + 점 (광고비가 0인 날은 ROAS가 없어서 선을 끊는다)
    const baseY = m.top + plotH;
    let run = [];
    const flush = () => {
      if (run.length > 1) {
        const firstX = run[0].split(",")[0];
        const lastX = run[run.length - 1].split(",")[0];
        svg.appendChild(svgEl("path", { d: `M${firstX},${baseY}L${run.join("L")}L${lastX},${baseY}Z`, fill: "url(#dailyAreaGrad)" }));
        svg.appendChild(svgEl("polyline", { points: run.join(" "), class: "chart-line" }));
      }
      run = [];
    };
    rows.forEach((r, i) => {
      if (Number.isFinite(r.roas)) run.push(`${cx(i)},${yRoas(r.roas)}`);
      else flush();
    });
    flush();
    if (band >= 10) {
      rows.forEach((r, i) => {
        if (Number.isFinite(r.roas)) svg.appendChild(svgEl("circle", { cx: cx(i), cy: yRoas(r.roas), r: 4, class: "chart-dot" }));
      });
    }

    // 일자별로 넓은 투명 영역을 깔아서 마우스/터치로 값 확인
    const hit = svgEl("rect", { x: m.left, y: m.top, width: plotW, height: plotH, fill: "transparent" });
    const showAt = (clientX) => {
      const box = svg.getBoundingClientRect();
      const i = Math.min(rows.length - 1, Math.max(0, Math.floor((clientX - box.left - m.left) / band)));
      const r = rows[i];
      hoverBand.setAttribute("x", m.left + band * i);
      hoverBand.setAttribute("visibility", "visible");
      tooltip.innerHTML =
        `<div class="tt-title">${r.date.replace(/-/g, ".")} · ${weekName(weekKeyOf(r.date))}</div>` +
        `<div><span class="tt-swatch tt-bar"></span>광고비 <b>${formatMetric(r.cost, "money")}</b></div>` +
        `<div><span class="tt-swatch tt-line"></span>ROAS <b>${formatMetric(r.roas, "pct0")}</b></div>`;
      tooltip.hidden = false;
      const areaBox = area.getBoundingClientRect();
      const x = cx(i);
      const tw = tooltip.offsetWidth;
      const left = x + 12 + tw > width ? x - 12 - tw : x + 12;
      tooltip.style.left = `${Math.max(0, left) + (box.left - areaBox.left)}px`;
      tooltip.style.top = `${m.top}px`;
    };
    hit.addEventListener("pointermove", (e) => showAt(e.clientX));
    hit.addEventListener("pointerdown", (e) => showAt(e.clientX));
    hit.addEventListener("pointerleave", () => {
      hoverBand.setAttribute("visibility", "hidden");
      tooltip.hidden = true;
    });
    svg.appendChild(hit);
    area.appendChild(svg);
  }

  let chartResizeTimer = null;
  window.addEventListener("resize", () => {
    clearTimeout(chartResizeTimer);
    chartResizeTimer = setTimeout(() => {
      if (state.rows.length) renderDailyChart();
    }, 150);
  });

  // ---------- 이벤트 ----------

  for (const btn of document.querySelectorAll(".tab-btn")) {
    btn.addEventListener("click", () => {
      state.activePlatform = btn.dataset.platform;
      state.table1Page = 1;
      renderAll();
    });
  }

  el("filterResetBtn").addEventListener("click", () => {
    const f = state.filters[state.activePlatform];
    for (const dim of FILTER_DIMENSIONS) f[dim].clear();
    state.table1Page = 1;
    renderAll();
  });

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
    if (!division || !passesFilters(platform, date, productType, item, row.__subItem)) return null;
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
  // MX_RAW 행(배열)을 만든다.
  function buildWeeklyRawRows(rows) {
    const ceRows = [];
    const mxRows = [];
    for (const row of rows) {
      if (row.__platform !== "coupang") continue;
      const built = buildWeeklyRawRow(row, row.__filePath, row.__productType);
      if (!built) continue;
      (built.division === "CE" ? ceRows : mxRows).push(built.values);
    }
    return { ceRows, mxRows };
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
      if (!division || !passesFilters(platform, date, productType, item, row.__subItem)) continue;
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
      if (!division || !passesFilters(platform, date, productType, item, row.__subItem)) continue;
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

  // 쿠팡 키워드_RAW: PA만 대상으로, 년+월+캠페인ID+캠페인명+VIID+키워드
  // 단위로(그룹·일자 구분 없이) 월별 합산.
  const KEYWORD_HEADER = [
    "년", "월", "품목", "광고유형", "캠페인ID", "캠페인명", "광고집행상품 VIID", "키워드",
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
      if (!division || !passesFilters(platform, date, productType, item, row.__subItem)) continue;
      const campaignIdRaw = weeklyRawText(row, fp, platform, "campaignId");
      const campaignRaw = weeklyRawText(row, fp, platform, "campaign");
      const viidRaw = weeklyRawText(row, fp, platform, "viid");
      const keywordRaw = weeklyRawText(row, fp, platform, "keyword");
      const year = yearOf(date);
      const month = monthOf(date);
      const key = [year, month, campaignIdRaw, campaignRaw, viidRaw, keywordRaw].join("\u0001");
      if (!groups.has(key)) {
        groups.set(key, { year, month, productType, item, subItem, campaignIdRaw, campaignRaw, viidRaw, keywordRaw, metrics: {} });
      }
      sumWeeklyMetrics(row, fp, platform, groups.get(key).metrics);
    }
    return [...groups.values()].map((g) => [
      Number(g.year),
      Number(g.month),
      g.item,
      g.productType,
      g.campaignIdRaw,
      g.campaignRaw,
      g.viidRaw,
      g.keywordRaw,
      ...keywordMetricValues(g.metrics),
      g.subItem,
    ]);
  }

  // 쿠팡 키워드_RAW 양식은 지표 순서가 "노출수, 클릭수, 광고비, ..."라서
  // (다른 시트는 "광고비, 노출수, 클릭수, ...") 앞의 세 개 순서를 맞춘다.
  function keywordMetricValues(metrics) {
    const [cost, impressions, clicks, ...rest] = weeklyMetricValues(metrics);
    return [impressions, clicks, cost, ...rest];
  }

  // ---------- 대용량 xlsx 직접 쓰기 ----------
  // CE_RAW는 원본 키워드 리포트 행 그대로라 20만 행이 넘는다. SheetJS는 시트
  // 하나의 모든 칸(20만 x 25 = 500만 개)을 객체 하나에 담았다가 한 번에
  // 쓰기 때문에, 브라우저 메모리를 2.5GB 넘게 쓰고 행이 조금만 더 늘면
  // 브라우저 자체 한도("invalid table size")에 걸려 아예 실패한다. 그래서
  // 이 다운로드는 시트 XML을 몇 천 행씩 바로 압축하면서 xlsx(zip)로 직접
  // 쓴다 - 메모리를 거의 안 쓰고 행 수에 상관없이 동작한다.
  const CRC32_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32Update(crc, bytes) {
    let c = crc ^ 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC32_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  // produce(write)가 write(Uint8Array)로 넘겨주는 내용을 zip 항목 하나로
  // 만든다. 브라우저가 CompressionStream("deflate-raw")를 지원하면 압축하고,
  // 아니면 압축 없이 그대로 담는다(파일만 커질 뿐 똑같이 열린다).
  async function makeZipEntry(name, produce) {
    let crc = 0;
    let size = 0;
    let stream = null;
    try {
      if (typeof CompressionStream !== "undefined") stream = new CompressionStream("deflate-raw");
    } catch (e) {
      stream = null;
    }
    const parts = [];
    let compSize = 0;
    if (!stream) {
      await produce(async (bytes) => {
        crc = crc32Update(crc, bytes);
        size += bytes.length;
        parts.push(bytes);
      });
      return { name, parts, crc, size, compSize: size, method: 0 };
    }
    const writer = stream.writable.getWriter();
    const reader = stream.readable.getReader();
    const reading = (async () => {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        parts.push(value);
        compSize += value.length;
      }
    })();
    await produce(async (bytes) => {
      crc = crc32Update(crc, bytes);
      size += bytes.length;
      await writer.write(bytes);
    });
    await writer.close();
    await reading;
    return { name, parts, crc, size, compSize, method: 8 };
  }

  function zipToBlob(entries) {
    const enc = new TextEncoder();
    const out = [];
    const central = [];
    let offset = 0;
    for (const e of entries) {
      const nameBytes = enc.encode(e.name);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);
      local.setUint16(4, 20, true);
      local.setUint16(6, 0x0800, true); // 파일명 UTF-8
      local.setUint16(8, e.method, true);
      local.setUint16(10, 0, true);
      local.setUint16(12, 0x21, true); // 1980-01-01
      local.setUint32(14, e.crc, true);
      local.setUint32(18, e.compSize, true);
      local.setUint32(22, e.size, true);
      local.setUint16(26, nameBytes.length, true);
      local.setUint16(28, 0, true);
      out.push(local, nameBytes, ...e.parts);
      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true);
      cd.setUint16(4, 20, true);
      cd.setUint16(6, 20, true);
      cd.setUint16(8, 0x0800, true);
      cd.setUint16(10, e.method, true);
      cd.setUint16(12, 0, true);
      cd.setUint16(14, 0x21, true);
      cd.setUint32(16, e.crc, true);
      cd.setUint32(20, e.compSize, true);
      cd.setUint32(24, e.size, true);
      cd.setUint16(28, nameBytes.length, true);
      cd.setUint32(42, offset, true);
      central.push(cd, nameBytes);
      offset += 30 + nameBytes.length + e.compSize;
    }
    const cdSize = central.reduce((n, part) => n + part.byteLength, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, entries.length, true);
    end.setUint16(10, entries.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);
    return new Blob([...out, ...central, end], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
  }

  function xmlEscape(v) {
    return String(v)
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function columnLetter(i) {
    let s = "";
    for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
  }

  const XLSX_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const XLSX_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  // sheets: [{ name, header, rows }] (rows는 값 배열의 배열). 숫자는 숫자
  // 칸으로, 그 외는 글자 칸으로 쓴다.
  async function downloadLargeXlsx(fileName, sheets) {
    const enc = new TextEncoder();
    const textEntry = (name, text) => makeZipEntry(name, (write) => write(enc.encode(text)));
    const entries = [];
    for (let si = 0; si < sheets.length; si++) {
      const { header, rows } = sheets[si];
      const width = Math.max(header.length, ...rows.slice(0, 1).map((r) => r.length));
      const letters = Array.from({ length: width }, (_, i) => columnLetter(i));
      const rowXml = (values, r) => {
        let x = `<row r="${r}">`;
        for (let c = 0; c < values.length; c++) {
          const v = values[c];
          if (v == null || v === "") continue;
          const ref = letters[c] || columnLetter(c);
          if (typeof v === "number") {
            if (Number.isFinite(v)) x += `<c r="${ref}${r}"><v>${v}</v></c>`;
          } else {
            x += `<c r="${ref}${r}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(v)}</t></is></c>`;
          }
        }
        return x + "</row>";
      };
      entries.push(
        await makeZipEntry(`xl/worksheets/sheet${si + 1}.xml`, async (write) => {
          await write(enc.encode(`${XML_HEAD}<worksheet xmlns="${XLSX_MAIN_NS}"><sheetData>${rowXml(header, 1)}`));
          const CHUNK = 2000;
          for (let i = 0; i < rows.length; i += CHUNK) {
            let x = "";
            const stop = Math.min(rows.length, i + CHUNK);
            for (let j = i; j < stop; j++) x += rowXml(rows[j], j + 2);
            await write(enc.encode(x));
          }
          await write(enc.encode("</sheetData></worksheet>"));
        })
      );
    }
    const sheetList = sheets
      .map((sh, i) => `<sheet name="${xmlEscape(sh.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
      .join("");
    entries.push(
      await textEntry(
        "[Content_Types].xml",
        `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
          sheets
            .map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
            .join("") +
          "</Types>"
      ),
      await textEntry(
        "_rels/.rels",
        `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="${XLSX_REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
      ),
      await textEntry(
        "xl/workbook.xml",
        `${XML_HEAD}<workbook xmlns="${XLSX_MAIN_NS}" xmlns:r="${XLSX_REL_NS}"><sheets>${sheetList}</sheets></workbook>`
      ),
      await textEntry(
        "xl/_rels/workbook.xml.rels",
        `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          sheets
            .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${XLSX_REL_NS}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
            .join("") +
          `<Relationship Id="rId${sheets.length + 1}" Type="${XLSX_REL_NS}/styles" Target="styles.xml"/></Relationships>`
      ),
      await textEntry(
        "xl/styles.xml",
        `${XML_HEAD}<styleSheet xmlns="${XLSX_MAIN_NS}">` +
          '<fonts count="1"><font><sz val="11"/><name val="맑은 고딕"/></font></fonts>' +
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
          '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
          "</styleSheet>"
      )
    );
    const blob = zipToBlob(entries);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  el("downloadReportBtn").addEventListener("click", async () => {
    const btn = el("downloadReportBtn");
    const originalText = btn.textContent;
    btn.disabled = true;
    btn.textContent = "리포트 준비 중...";
    try {
      // CE_RAW/MX_RAW/쿠팡 키워드_RAW는 키워드 단위가 꼭 필요해서, 평소
      // 표1~4가 쓰는 가벼운 state.rows 대신 지금 클릭한 시점에 쿠팡 파일의
      // 다운로드용 캐시(키워드를 합치지 않은 것)를 받아서 쓴다.
      const sheets = await withFullCoupangRows((fullRows) => {
        const { ceRows, mxRows } = buildWeeklyRawRows(fullRows);
        return [
          { name: "CE_RAW", header: WEEKLY_RAW_HEADER, rows: ceRows },
          { name: "MX_RAW", header: WEEKLY_RAW_HEADER, rows: mxRows },
          { name: "쿠팡_RAW", header: MONTHLY_RAW_HEADER, rows: buildMonthlyRawRows(fullRows) },
          { name: "쿠팡_전환상품_RAW", header: CONVERSION_PRODUCT_HEADER, rows: buildConversionProductRows(fullRows) },
          { name: "쿠팡 키워드_RAW", header: KEYWORD_HEADER, rows: buildKeywordRows(fullRows) },
        ];
      });
      await downloadLargeXlsx(`쿠팡_상품광고_주간리포트_${new Date().toISOString().slice(0, 10)}.xlsx`, sheets);
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

  // ---------- 접속 비밀번호 ----------
  // 비밀번호 원문은 어디에도 두지 않고 SHA-256 해시값만 둔다. 입력값을 같은
  // 방식으로 해시해서 비교하고, 맞으면 그때 데이터를 불러온다. 한 번 맞히면
  // 같은 탭에서는(새로고침 포함) 다시 묻지 않는다.
  const PASSWORD_SHA256 = "dd0a3d18a5b7de33fb11906a12a10b1d3dca8fe42e99a2b76ec6e56d926c3513";
  const UNLOCK_KEY = "dashboardUnlocked";

  async function sha256Hex(text) {
    const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
    return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
  }

  function unlockDashboard() {
    document.body.classList.remove("locked");
    el("lockScreen").hidden = true;
    loadManifest();
  }

  let alreadyUnlocked = false;
  try {
    alreadyUnlocked = sessionStorage.getItem(UNLOCK_KEY) === PASSWORD_SHA256;
  } catch (e) {
    /* 저장소를 못 쓰면 매번 묻는다 */
  }
  if (alreadyUnlocked) {
    unlockDashboard();
  } else {
    const input = el("lockPassword");
    const error = el("lockError");
    el("lockForm").addEventListener("submit", async (e) => {
      e.preventDefault();
      error.hidden = true;
      let hash = "";
      try {
        hash = await sha256Hex(input.value);
      } catch (err) {
        error.textContent = "이 브라우저에서는 확인할 수 없습니다. https 주소로 접속해 주세요.";
        error.hidden = false;
        return;
      }
      if (hash !== PASSWORD_SHA256) {
        error.textContent = "비밀번호가 맞지 않습니다.";
        error.hidden = false;
        input.value = "";
        input.focus();
        return;
      }
      try {
        sessionStorage.setItem(UNLOCK_KEY, PASSWORD_SHA256);
      } catch (err) {
        /* 저장 실패해도 이번 접속은 그대로 열린다 */
      }
      unlockDashboard();
    });
    input.focus();
  }
})();
