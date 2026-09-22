(() => {
  "use strict";

  const PLATFORM_LABEL = { coupang: "쿠팡", naver: "네이버" };
  const PLATFORMS = ["coupang", "naver"];
  const PAGE_SIZE = 50;

  // 실제 raw 컬럼명이 플랫폼마다 다르므로, 리포트 필드 <-> raw 컬럼명 매핑표.
  // null이면 아래 FIELD_AUTO_PATTERNS로 자동 추정하고, NOT_MAPPED이면 추정도 하지 않고 "-"로 비워둔다.
  const NOT_MAPPED = Symbol("not-mapped");
  const FIELD_MAP = {
    coupang: {
      date: "날짜",
      campaign: "캠페인명",
      group: "광고그룹",
      cost: "광고비",
      impressions: "노출수",
      clicks: "클릭수",
      conversions: "직접 판매수량(1일)",
      revenue: "직접 전환매출액(1일)",
    },
    naver: {
      date: null,
      campaign: null,
      group: null,
      cost: null,
      impressions: null,
      clicks: null,
      // 네이버 raw 샘플을 아직 못 봐서 전환수/전환매출 컬럼명을 확정하지 못했다.
      // 실제 파일을 받으면 위 coupang처럼 정확한 컬럼명으로 채워 넣으면 된다.
      conversions: NOT_MAPPED,
      revenue: NOT_MAPPED,
    },
  };

  // 명시적 매핑이 없을 때(null) 컬럼명을 추정하는 패턴. "클릭률"이 "클릭수"로 오인되지 않도록 주의.
  const FIELD_AUTO_PATTERNS = {
    date: /날짜|일자|date/i,
    campaign: /^캠페인명$|^캠페인$/i,
    group: /그룹명|광고그룹|^그룹$/i,
    cost: /광고비|총비용|^비용$|spend|cost/i,
    impressions: /노출수|^노출$|impression/i,
    clicks: /클릭수|^클릭$/i,
  };

  const METRIC_COLUMNS = [
    { key: "cost", label: "광고비", deps: ["cost"], format: "money" },
    { key: "impressions", label: "노출", deps: ["impressions"], format: "count" },
    { key: "clicks", label: "클릭", deps: ["clicks"], format: "count" },
    { key: "ctr", label: "CTR", deps: ["clicks", "impressions"], format: "pct" },
    { key: "cpc", label: "CPC", deps: ["cost", "clicks"], format: "money" },
    { key: "conversions", label: "전환수", deps: ["conversions"], format: "count" },
    { key: "revenue", label: "전환매출", deps: ["revenue"], format: "money" },
    { key: "cvr", label: "CVR", deps: ["conversions", "clicks"], format: "pct" },
    { key: "aov", label: "객단가", deps: ["revenue", "conversions"], format: "money" },
    { key: "roas", label: "ROAS", deps: ["revenue", "cost"], format: "pct" },
  ];
  const METRIC_BY_KEY = Object.fromEntries(METRIC_COLUMNS.map((m) => [m.key, m]));

  const state = {
    manifest: { files: [] },
    selectedPaths: new Set(),
    rows: [], // 원본 raw 행 (플랫폼/파일 태그 포함)
    colsByFile: new Map(), // filePath -> Set(원본 컬럼명)
    activePlatform: "coupang",
    visibleMetrics: new Set(), // 기본은 비어있음 = 날짜/캠페인명/그룹명만 표시
    aggregatedRows: [], // 현재 플랫폼 기준, (날짜,캠페인명,그룹명)으로 합산된 행
    search: "",
    sortColumn: null,
    sortDir: 1,
    page: 1,
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
      // 엑셀은 날짜 셀이 숫자(일련번호)로 저장되므로 cellDates로 JS Date로 바꾸고,
      // sheet_to_json에서 raw:false + dateNF로 화면에 보이는 형태의 문자열로 뽑는다.
      const workbook = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { raw: false, dateNF: "yyyy-mm-dd", defval: "" });
      return rows.map(trimKeys);
    }
    const text = decodeCsvBuffer(buf).replace(/^﻿/, "");
    const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
    return parsed.data.map(trimKeys);
  }

  function parseNumberLoose(v) {
    if (v == null) return NaN;
    const s = String(v).trim().replace(/[,₩%원\s]/g, "");
    if (s === "" || s === "-") return NaN;
    const n = Number(s);
    return Number.isFinite(n) ? n : NaN;
  }

  function parseDateLoose(v) {
    if (v == null) return null;
    let s = String(v).trim();
    if (!s) return null;
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
    if (mapped) return cols.has(mapped) ? mapped : null;
    const pattern = FIELD_AUTO_PATTERNS[field];
    if (pattern) {
      for (const c of cols) if (pattern.test(c)) return c;
    }
    return null;
  }

  function filesForPlatform(platform) {
    return state.manifest.files.filter((f) => f.platform === platform && state.selectedPaths.has(f.path)).map((f) => f.path);
  }

  function isMetricAvailable(platform, metricKey) {
    const deps = METRIC_BY_KEY[metricKey].deps;
    const paths = filesForPlatform(platform);
    return paths.some((fp) => deps.every((d) => resolveField(fp, platform, d)));
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
    await reloadSelectedData();
  }

  async function reloadSelectedData() {
    const files = state.manifest.files.filter((f) => state.selectedPaths.has(f.path));
    const allRows = [];
    for (const f of files) {
      try {
        const rows = await parseDataFile(f.path);
        for (const row of rows) {
          allRows.push({ __platform: f.platform, __file: f.name, __filePath: f.path, ...row });
        }
      } catch (e) {
        console.error("파일을 불러오지 못했습니다:", f.path, e);
      }
    }
    state.rows = allRows;
    computeColsByFile();
    state.page = 1;
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

  function computeAggregatedRows(platform) {
    const groups = new Map();
    for (const row of state.rows) {
      if (row.__platform !== platform) continue;
      if (!state.selectedPaths.has(row.__filePath)) continue;
      const fp = row.__filePath;

      const dateCol = resolveField(fp, platform, "date");
      const date = dateCol ? parseDateLoose(row[dateCol]) : null;
      if (!date) continue;

      const campaignCol = resolveField(fp, platform, "campaign");
      const groupCol = resolveField(fp, platform, "group");
      const campaign = campaignCol ? row[campaignCol] ?? "" : "";
      const group = groupCol ? row[groupCol] ?? "" : "";

      const key = `${date}\u0001${campaign}\u0001${group}`;
      if (!groups.has(key)) {
        groups.set(key, { date, campaign, group, cost: 0, impressions: 0, clicks: 0, conversions: 0, revenue: 0 });
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

    return [...groups.values()].map((g) => ({
      ...g,
      ctr: g.impressions > 0 ? (g.clicks / g.impressions) * 100 : null,
      cpc: g.clicks > 0 ? g.cost / g.clicks : null,
      cvr: g.clicks > 0 ? (g.conversions / g.clicks) * 100 : null,
      aov: g.conversions > 0 ? g.revenue / g.conversions : null,
      roas: g.cost > 0 ? (g.revenue / g.cost) * 100 : null,
    }));
  }

  function recompute() {
    state.aggregatedRows = computeAggregatedRows(state.activePlatform);
    // 현재 플랫폼에서 더 이상 계산 불가능한 지표는 선택에서 제거
    for (const key of [...state.visibleMetrics]) {
      if (!isMetricAvailable(state.activePlatform, key)) state.visibleMetrics.delete(key);
    }
  }

  // ---------- 렌더링 ----------

  function renderAll() {
    recompute();
    renderTabs();
    renderFileList();
    renderColumnControls();
    renderSummary();
    renderTable();
  }

  function renderTabs() {
    for (const btn of document.querySelectorAll(".tab-btn")) {
      btn.classList.toggle("active", btn.dataset.platform === state.activePlatform);
    }
  }

  function renderFileList() {
    const area = el("fileListArea");
    const files = state.manifest.files.filter((f) => f.platform === state.activePlatform);
    if (files.length === 0) {
      area.innerHTML = `<div class="empty-state">아직 업로드된 데이터가 없습니다. <code>data/${state.activePlatform}</code> 폴더에 raw csv/xlsx 파일을 올려주세요.</div>`;
      return;
    }
    area.innerHTML = "";
    for (const f of files) {
      const row = document.createElement("div");
      row.className = "file-row";
      const label = document.createElement("label");
      label.className = "file-name";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = state.selectedPaths.has(f.path);
      checkbox.addEventListener("change", () => {
        if (checkbox.checked) state.selectedPaths.add(f.path);
        else state.selectedPaths.delete(f.path);
        renderAll();
      });
      label.appendChild(checkbox);
      const nameSpan = document.createElement("span");
      nameSpan.textContent = ` ${f.name} (${formatSize(f.size)})`;
      label.appendChild(nameSpan);
      row.appendChild(label);

      const link = document.createElement("a");
      link.href = f.path;
      link.download = f.name;
      link.textContent = "원본 다운로드 ↓";
      row.appendChild(link);

      area.appendChild(row);
    }
  }

  function formatSize(bytes) {
    if (!Number.isFinite(bytes)) return "";
    if (bytes < 1024) return `${bytes}B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)}KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)}MB`;
  }

  function renderColumnControls() {
    const hasData = state.aggregatedRows.length > 0;
    el("columnCard").style.display = hasData ? "" : "none";
    if (!hasData) return;

    const chipList = el("metricChipList");
    chipList.innerHTML = "";
    const unavailable = [];
    for (const metric of METRIC_COLUMNS) {
      const available = isMetricAvailable(state.activePlatform, metric.key);
      if (!available) unavailable.push(metric.label);

      const chip = document.createElement("label");
      chip.className = "chip";
      if (!available) chip.style.opacity = "0.4";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.disabled = !available;
      cb.checked = state.visibleMetrics.has(metric.key);
      cb.addEventListener("change", () => {
        if (cb.checked) state.visibleMetrics.add(metric.key);
        else state.visibleMetrics.delete(metric.key);
        renderSummary();
        renderTable();
      });
      chip.appendChild(cb);
      chip.appendChild(document.createTextNode(metric.label));
      chipList.appendChild(chip);
    }

    const hint = el("unavailableHint");
    hint.textContent = unavailable.length
      ? `${PLATFORM_LABEL[state.activePlatform]}에서 아직 인식되지 않는 지표: ${unavailable.join(", ")} (raw 컬럼 확인 필요)`
      : "";

    el("searchInput").oninput = (e) => {
      state.search = e.target.value.trim().toLowerCase();
      state.page = 1;
      renderTable();
    };
  }

  function renderSummary() {
    const hasMetrics = state.aggregatedRows.length > 0 && state.visibleMetrics.size > 0;
    el("summaryCard").style.display = hasMetrics ? "" : "none";
    if (!hasMetrics) return;

    const totals = { cost: 0, impressions: 0, clicks: 0, conversions: 0, revenue: 0 };
    for (const row of state.aggregatedRows) {
      totals.cost += row.cost;
      totals.impressions += row.impressions;
      totals.clicks += row.clicks;
      totals.conversions += row.conversions;
      totals.revenue += row.revenue;
    }
    const derived = {
      ctr: totals.impressions > 0 ? (totals.clicks / totals.impressions) * 100 : null,
      cpc: totals.clicks > 0 ? totals.cost / totals.clicks : null,
      cvr: totals.clicks > 0 ? (totals.conversions / totals.clicks) * 100 : null,
      aov: totals.conversions > 0 ? totals.revenue / totals.conversions : null,
      roas: totals.cost > 0 ? (totals.revenue / totals.cost) * 100 : null,
    };
    const combined = { ...totals, ...derived };

    const grid = el("statGrid");
    grid.innerHTML = "";
    for (const key of state.visibleMetrics) {
      const metric = METRIC_BY_KEY[key];
      const tile = document.createElement("div");
      tile.className = "stat-tile";
      tile.innerHTML = `
        <div class="stat-label">${escapeHtml(metric.label)} 합계</div>
        <div class="stat-value">${formatMetric(combined[key], metric.format)}</div>`;
      grid.appendChild(tile);
    }
  }

  function visibleColumns() {
    return ["date", "campaign", "group", ...METRIC_COLUMNS.filter((m) => state.visibleMetrics.has(m.key)).map((m) => m.key)];
  }

  function columnLabel(col) {
    if (col === "date") return "날짜";
    if (col === "campaign") return "캠페인명";
    if (col === "group") return "그룹명";
    return METRIC_BY_KEY[col]?.label ?? col;
  }

  function getFilteredSortedRows() {
    let rows = state.aggregatedRows;
    if (state.search) {
      const q = state.search;
      rows = rows.filter((r) => `${r.date} ${r.campaign} ${r.group}`.toLowerCase().includes(q));
    }
    if (state.sortColumn) {
      const col = state.sortColumn;
      rows = [...rows].sort((a, b) => {
        const av = a[col];
        const bv = b[col];
        let cmp;
        if (typeof av === "number" || typeof bv === "number") {
          const an = Number.isFinite(av) ? av : -Infinity;
          const bn = Number.isFinite(bv) ? bv : -Infinity;
          cmp = an - bn;
        } else {
          cmp = String(av ?? "").localeCompare(String(bv ?? ""), "ko");
        }
        return cmp * state.sortDir;
      });
    }
    return rows;
  }

  function renderTable() {
    const hasData = state.aggregatedRows.length > 0;
    el("tableCard").style.display = hasData ? "" : "none";
    if (!hasData) return;

    const rows = getFilteredSortedRows();
    const cols = visibleColumns();

    const head = el("tableHead");
    head.innerHTML = "";
    const tr = document.createElement("tr");
    for (const col of cols) {
      const th = document.createElement("th");
      th.textContent = columnLabel(col);
      if (col === state.sortColumn) th.classList.add("sorted");
      th.addEventListener("click", () => {
        if (state.sortColumn === col) state.sortDir *= -1;
        else {
          state.sortColumn = col;
          state.sortDir = 1;
        }
        renderTable();
      });
      tr.appendChild(th);
    }
    head.appendChild(tr);

    const totalPages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    state.page = Math.min(state.page, totalPages);
    const pageRows = rows.slice((state.page - 1) * PAGE_SIZE, state.page * PAGE_SIZE);

    const body = el("tableBody");
    body.innerHTML = "";
    for (const row of pageRows) {
      const tr2 = document.createElement("tr");
      for (const col of cols) {
        const td = document.createElement("td");
        if (col === "date" || col === "campaign" || col === "group") {
          td.textContent = row[col] ?? "";
        } else {
          td.textContent = formatMetric(row[col], METRIC_BY_KEY[col].format);
        }
        tr2.appendChild(td);
      }
      body.appendChild(tr2);
    }

    el("pageInfo").textContent = `총 ${rows.length.toLocaleString("ko-KR")}행 · ${state.page} / ${totalPages} 페이지`;
    el("prevPageBtn").disabled = state.page <= 1;
    el("nextPageBtn").disabled = state.page >= totalPages;
  }

  // ---------- 이벤트 ----------

  for (const btn of document.querySelectorAll(".tab-btn")) {
    btn.addEventListener("click", () => {
      state.activePlatform = btn.dataset.platform;
      state.page = 1;
      state.sortColumn = null;
      renderAll();
    });
  }

  el("selectAllMetricsBtn").addEventListener("click", () => {
    for (const metric of METRIC_COLUMNS) {
      if (isMetricAvailable(state.activePlatform, metric.key)) state.visibleMetrics.add(metric.key);
    }
    renderColumnControls();
    renderSummary();
    renderTable();
  });

  el("clearMetricsBtn").addEventListener("click", () => {
    state.visibleMetrics.clear();
    renderColumnControls();
    renderSummary();
    renderTable();
  });

  el("prevPageBtn").addEventListener("click", () => {
    state.page = Math.max(1, state.page - 1);
    renderTable();
  });
  el("nextPageBtn").addEventListener("click", () => {
    state.page += 1;
    renderTable();
  });

  function buildDownloadRows() {
    const rows = getFilteredSortedRows();
    const cols = visibleColumns();
    return rows.map((row) => {
      const out = {};
      for (const col of cols) {
        out[columnLabel(col)] = col === "date" || col === "campaign" || col === "group" ? row[col] ?? "" : row[col];
      }
      return out;
    });
  }

  function triggerBlobDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  el("downloadFilteredBtn").addEventListener("click", () => {
    const csv = Papa.unparse(buildDownloadRows());
    const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
    triggerBlobDownload(blob, `${state.activePlatform}_report_${new Date().toISOString().slice(0, 10)}.csv`);
  });

  el("downloadFilteredXlsxBtn").addEventListener("click", () => {
    const worksheet = XLSX.utils.json_to_sheet(buildDownloadRows());
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, PLATFORM_LABEL[state.activePlatform]);
    XLSX.writeFile(workbook, `${state.activePlatform}_report_${new Date().toISOString().slice(0, 10)}.xlsx`);
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
