(() => {
  "use strict";

  const PLATFORM_LABEL = { coupang: "쿠팡", naver: "네이버" };
  const PLATFORMS = ["coupang", "naver"];
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
      NCA: "__campaign",
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

  const TABLE1_COLS = [
    { key: "productType", label: "상품구분" },
    { key: "campaign", label: "캠페인명" },
    { key: "group", label: "그룹명" },
    ...METRIC_COLUMNS,
  ];
  const TABLE2_COLS = [{ key: "productType", label: "상품구분" }, ...METRIC_COLUMNS];
  const TABLE3_COLS = [
    { key: "weekLabel", label: "주차" },
    { key: "productType", label: "상품구분" },
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
    rows: [], // 원본 raw 행 (플랫폼/파일 태그 포함)
    colsByFile: new Map(), // filePath -> Set(원본 컬럼명)
    activePlatform: "coupang",
    table1Page: 1,
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
    if (/\.xlsx?$/i.test(path)) {
      // 엑셀은 날짜 셀이 내부적으로 숫자(일련번호, 예: 46266)로 저장된다.
      // cellDates로 최대한 JS Date로 바꾸되, 셀 서식이 인식 안 되는 경우를 대비해
      // raw:true로 원본 값(숫자 or Date)을 그대로 받아 parseDateLoose에서 이중으로 처리한다.
      const workbook = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { raw: true, defval: "" });
      return rows.map((row) => trimKeys(normalizeXlsxRow(row)));
    }
    const text = decodeCsvBuffer(buf).replace(/^﻿/, "");
    const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
    return parsed.data.map(trimKeys);
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

  // platform의 field(날짜/캠페인/그룹/광고비 등)가 filePath 안에서 어떤 실제 컬럼명인지 찾는다.
  function resolveField(filePath, platform, field) {
    const cols = state.colsByFile.get(filePath);
    if (!cols) return null;
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
    await reloadData();
  }

  async function reloadData() {
    const allRows = [];
    for (const f of state.manifest.files) {
      try {
        const rows = await parseDataFile(f.path);
        const productType = detectProductType(f.platform, f.name);
        for (const row of rows) {
          allRows.push({ __platform: f.platform, __file: f.name, __filePath: f.path, __productType: productType, ...row });
        }
      } catch (e) {
        console.error("파일을 불러오지 못했습니다:", f.path, e);
      }
    }
    state.rows = allRows;
    computeColsByFile();
    renderAll();
  }

  function computeColsByFile() {
    const map = new Map();
    for (const row of state.rows) {
      if (!map.has(row.__filePath)) map.set(row.__filePath, new Set());
      const set = map.get(row.__filePath);
      for (const k of Object.keys(row)) {
        if (k.startsWith("__")) continue;
        set.add(k);
      }
    }
    state.colsByFile = map;
  }

  // ---------- 집계 ----------

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

    const cost = costCol ? parseNumberLoose(row[costCol]) : NaN;
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

  // 표1. 상품구분+캠페인명+그룹명 기준 전체기간 합산 (날짜 무시)
  function computeTable1(platform) {
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const dateCol = resolveField(fp, platform, "date");
      if (!dateCol || !parseDateLoose(row[dateCol])) continue;
      const campaignCol = resolveField(fp, platform, "campaign");
      const groupCol = resolveField(fp, platform, "group");
      const campaign = campaignCol ? row[campaignCol] ?? "" : "";
      const group = groupCol ? row[groupCol] ?? "" : "";
      const productType = row.__productType;
      const key = `${productType}\u0001${campaign}\u0001${group}`;
      accumulateRow(groups, key, { productType, campaign, group }, row, fp, platform);
    }
    return finalizeGroups(groups).sort(
      (a, b) => a.productType.localeCompare(b.productType) || a.campaign.localeCompare(b.campaign, "ko") || a.group.localeCompare(b.group, "ko")
    );
  }

  // 표2. 상품구분 기준 전체기간 합산
  function computeTable2(platform) {
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const dateCol = resolveField(fp, platform, "date");
      if (!dateCol || !parseDateLoose(row[dateCol])) continue;
      const productType = row.__productType;
      accumulateRow(groups, productType, { productType }, row, fp, platform);
    }
    return finalizeGroups(groups).sort((a, b) => (b.revenue || 0) - (a.revenue || 0));
  }

  // 표3. 주차+상품구분 기준 합산
  function computeTable3(platform) {
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const dateCol = resolveField(fp, platform, "date");
      const date = dateCol ? parseDateLoose(row[dateCol]) : null;
      if (!date) continue;
      const week = weekOf(date);
      const productType = row.__productType;
      const key = `${week.key}\u0001${productType}`;
      accumulateRow(groups, key, { weekKey: week.key, weekLabel: week.label, productType }, row, fp, platform);
    }
    return finalizeGroups(groups).sort((a, b) => a.weekKey.localeCompare(b.weekKey) || a.productType.localeCompare(b.productType));
  }

  // 표4. 상품구분+상품명 기준 합산 후 전환매출 상위 10개
  function computeTable4(platform) {
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      const fp = row.__filePath;
      const dateCol = resolveField(fp, platform, "date");
      if (!dateCol || !parseDateLoose(row[dateCol])) continue;
      const productType = row.__productType;
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
    renderRawDownload();
    renderTable1();
    renderTable2();
    renderTable3();
    renderTable4();
  }

  function renderTabs() {
    for (const btn of document.querySelectorAll(".tab-btn")) {
      btn.classList.toggle("active", btn.dataset.platform === state.activePlatform);
    }
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

  // 원본 다운로드는 화면 맨 아래에서만 제공한다.
  function renderRawDownload() {
    const area = el("rawDownloadArea");
    const files = filesForActivePlatform();
    area.innerHTML = "";
    if (files.length === 0) {
      area.innerHTML = `<div class="empty-state">다운로드할 파일이 없습니다.</div>`;
      return;
    }
    for (const f of files) {
      const row = document.createElement("div");
      row.className = "file-row";
      const span = document.createElement("span");
      span.className = "file-name";
      span.textContent = `${f.name} (${formatSize(f.size)})`;
      row.appendChild(span);
      const link = document.createElement("a");
      link.href = f.path;
      link.download = f.name;
      link.textContent = "원본 다운로드 ↓";
      row.appendChild(link);
      area.appendChild(row);
    }
  }

  function cellText(row, col) {
    if (col.key === "rank" || col.key === "productType" || col.key === "productName" || col.key === "campaign" || col.key === "group" || col.key === "weekLabel") {
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
    renderMetricTable(el("table1Head"), el("table1Body"), TABLE1_COLS, pageRows, "데이터가 없습니다.");
    el("table1PageInfo").textContent = `총 ${rows.length.toLocaleString("ko-KR")}행 · ${state.table1Page} / ${totalPages} 페이지`;
    el("table1PrevBtn").disabled = state.table1Page <= 1;
    el("table1NextBtn").disabled = state.table1Page >= totalPages;
  }

  function renderTable2() {
    const rows = computeTable2(state.activePlatform);
    renderMetricTable(el("table2Head"), el("table2Body"), TABLE2_COLS, rows, "데이터가 없습니다.");
  }

  function renderTable3() {
    const rows = computeTable3(state.activePlatform);
    renderMetricTable(el("table3Head"), el("table3Body"), TABLE3_COLS, rows, "데이터가 없습니다.");
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

  // ---------- 화면을 엑셀로 다운로드 ----------

  function exportCellValue(row, col) {
    if (col.key === "rank" || col.key === "productType" || col.key === "productName" || col.key === "campaign" || col.key === "group" || col.key === "weekLabel") {
      return row[col.key] ?? "";
    }
    const v = row[col.key];
    if (v == null || !Number.isFinite(v)) return "";
    return Math.round(v * 100) / 100;
  }

  function exportColumnLabel(col) {
    if (col.format === "pct") return `${col.label}(%)`;
    if (col.format === "money") return `${col.label}(원)`;
    return col.label;
  }

  function addTableToAOA(aoa, title, columns, rows) {
    aoa.push([title]);
    aoa.push(columns.map(exportColumnLabel));
    for (const row of rows) aoa.push(columns.map((col) => exportCellValue(row, col)));
    aoa.push([]);
  }

  function buildPlatformSheet(platform) {
    const aoa = [];
    addTableToAOA(aoa, "표1. 전체기간의 지표", TABLE1_COLS, computeTable1(platform));
    addTableToAOA(aoa, "표2. 광고상품별 지표 (전체기간)", TABLE2_COLS, computeTable2(platform));
    addTableToAOA(aoa, "표3. 주차별, 광고상품별 지표", TABLE3_COLS, computeTable3(platform));
    addTableToAOA(aoa, "표4. 우수 판매 TOP10 상품", TABLE4_COLS, computeTable4(platform));
    return XLSX.utils.aoa_to_sheet(aoa);
  }

  el("downloadScreenXlsxBtn").addEventListener("click", () => {
    const workbook = XLSX.utils.book_new();
    for (const platform of PLATFORMS) {
      XLSX.utils.book_append_sheet(workbook, buildPlatformSheet(platform), PLATFORM_LABEL[platform]);
    }
    XLSX.writeFile(workbook, `ad_report_${new Date().toISOString().slice(0, 10)}.xlsx`);
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
