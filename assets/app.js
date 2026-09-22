(() => {
  "use strict";

  const PLATFORM_LABEL = { coupang: "쿠팡", naver: "네이버" };
  const PAGE_SIZE = 50;

  const PLATFORMS = ["coupang", "naver"];

  const state = {
    manifest: { files: [] },
    selectedPaths: new Set(),
    rows: [],
    columns: [],
    colsByFile: new Map(), // filePath -> Set(컬럼명) — 같은 플랫폼이어도 광고상품별로 raw 컬럼 구성이 다를 수 있어 파일 단위로 관리
    dateColumnByFile: new Map(), // filePath -> 자동 감지된 날짜 컬럼명
    manualDateColumn: null, // 사용자가 직접 고른 날짜 컬럼명 (null이면 자동 감지 사용)
    metricColumns: new Set(),
    numericColumns: [],
    search: "",
    sortColumn: null,
    sortDir: 1,
    page: 1,
  };

  function platformsInRows() {
    return PLATFORMS.filter((p) => state.rows.some((r) => r.__platform === p));
  }

  function filesInRows() {
    return [...new Set(state.rows.map((r) => r.__filePath))];
  }

  // 같은 플랫폼이라도 광고상품(리포트 종류)마다 컬럼 구성이 다를 수 있어
  // (예: 쿠팡 "날짜" vs 네이버 "일자", 혹은 같은 쿠팡이라도 상품별로 컬럼이 다름),
  // 파일 단위로 실제 존재하는 컬럼 중에서만 날짜 컬럼을 찾는다.
  function resolveDateColumn(filePath) {
    if (state.manualDateColumn && state.colsByFile.get(filePath)?.has(state.manualDateColumn)) {
      return state.manualDateColumn;
    }
    return state.dateColumnByFile.get(filePath) || null;
  }

  function applyDates() {
    for (const row of state.rows) {
      const col = resolveDateColumn(row.__filePath);
      row.__date = col ? parseDateLoose(row[col]) : null;
    }
  }

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

  function formatNumber(n) {
    if (!Number.isFinite(n)) return "-";
    return n.toLocaleString("ko-KR", { maximumFractionDigits: 2 });
  }

  function cssVar(name) {
    return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
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
    renderFileList();
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
    computeColumns();
    state.page = 1;
    renderAll();
  }

  function computeColumns() {
    const cols = new Set();
    const colsByFile = new Map();
    for (const row of state.rows) {
      if (!colsByFile.has(row.__filePath)) colsByFile.set(row.__filePath, new Set());
      const fileCols = colsByFile.get(row.__filePath);
      for (const k of Object.keys(row)) {
        if (k === "__platform" || k === "__file" || k === "__filePath" || k === "__date") continue;
        cols.add(k);
        fileCols.add(k);
      }
    }
    state.columns = [...cols];
    state.colsByFile = colsByFile;

    state.dateColumnByFile = new Map();
    for (const filePath of filesInRows()) {
      state.dateColumnByFile.set(filePath, detectDateColumnForFile(filePath));
    }
    applyDates();

    state.numericColumns = detectNumericColumns();

    if (state.metricColumns.size === 0) {
      const preferred = state.numericColumns.filter((c) =>
        /비용|cost|spend|광고비|노출|impression|클릭|click|전환|conversion/i.test(c)
      );
      const defaults = (preferred.length ? preferred : state.numericColumns).slice(0, 2);
      state.metricColumns = new Set(defaults);
    } else {
      // 파일이 바뀌어 더 이상 존재하지 않는 컬럼은 선택에서 제거
      for (const c of [...state.metricColumns]) {
        if (!state.numericColumns.includes(c)) state.metricColumns.delete(c);
      }
    }
  }

  // 해당 파일에 실제로 존재하는 컬럼들 중에서만 날짜 컬럼 후보를 찾는다.
  function detectDateColumnForFile(filePath) {
    let best = null;
    let bestScore = 0;
    const rows = state.rows.filter((r) => r.__filePath === filePath);
    for (const col of state.colsByFile.get(filePath)) {
      let hit = 0;
      let total = 0;
      for (const row of rows) {
        const v = row[col];
        if (v == null || v === "") continue;
        total++;
        if (parseDateLoose(v)) hit++;
      }
      if (total === 0) continue;
      let score = hit / total;
      if (/날짜|date|일자/i.test(col)) score += 0.05;
      if (score > bestScore) {
        bestScore = score;
        best = col;
      }
    }
    return best;
  }

  function detectNumericColumns() {
    const dateCols = new Set([...state.dateColumnByFile.values()].filter(Boolean));
    const result = [];
    for (const col of state.columns) {
      if (dateCols.has(col)) continue;
      let hit = 0;
      let total = 0;
      for (const row of state.rows) {
        const v = row[col];
        if (v == null || v === "") continue;
        total++;
        if (!Number.isNaN(parseNumberLoose(v))) hit++;
      }
      if (total > 0 && hit / total >= 0.6) result.push(col);
    }
    return result;
  }

  // ---------- 렌더링 ----------

  function renderAll() {
    renderColumnControls();
    renderSummary();
    renderCharts();
    renderTable();
  }

  function renderFileList() {
    const area = el("fileListArea");
    if (state.manifest.files.length === 0) {
      area.innerHTML = `<div class="empty-state">아직 업로드된 데이터가 없습니다. <code>data/coupang</code> 또는 <code>data/naver</code> 폴더에 raw csv 파일을 올려주세요.</div>`;
      return;
    }
    area.innerHTML = "";
    for (const platform of ["coupang", "naver"]) {
      const files = state.manifest.files.filter((f) => f.platform === platform);
      if (files.length === 0) continue;
      const heading = document.createElement("div");
      heading.style.margin = "10px 0 4px";
      heading.style.fontSize = "12px";
      heading.style.color = "var(--text-muted)";
      heading.textContent = `${PLATFORM_LABEL[platform]} (${files.length}개 파일)`;
      area.appendChild(heading);

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
          reloadSelectedData();
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
  }

  function formatSize(bytes) {
    if (!Number.isFinite(bytes)) return "";
    if (bytes < 1024) return `${bytes}B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)}KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)}MB`;
  }

  function renderColumnControls() {
    const hasData = state.rows.length > 0;
    el("columnCard").style.display = hasData ? "" : "none";
    if (!hasData) return;

    const dateSelect = el("dateColumnSelect");
    dateSelect.innerHTML = "";
    const autoOpt = document.createElement("option");
    autoOpt.value = "";
    autoOpt.textContent = "자동 감지";
    dateSelect.appendChild(autoOpt);
    for (const col of state.columns) {
      const opt = document.createElement("option");
      opt.value = col;
      opt.textContent = col;
      dateSelect.appendChild(opt);
    }
    dateSelect.value = state.manualDateColumn || "";
    dateSelect.onchange = () => {
      state.manualDateColumn = dateSelect.value || null;
      applyDates();
      renderDateHint();
      renderCharts();
    };
    renderDateHint();

    const chipList = el("metricChipList");
    chipList.innerHTML = "";
    for (const col of state.numericColumns) {
      const chip = document.createElement("label");
      chip.className = "chip";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = state.metricColumns.has(col);
      cb.addEventListener("change", () => {
        if (cb.checked) state.metricColumns.add(col);
        else state.metricColumns.delete(col);
        renderSummary();
        renderCharts();
      });
      chip.appendChild(cb);
      chip.appendChild(document.createTextNode(col));
      chipList.appendChild(chip);
    }

    el("searchInput").oninput = (e) => {
      state.search = e.target.value.trim().toLowerCase();
      state.page = 1;
      renderTable();
    };
  }

  // 같은 플랫폼이라도 광고상품(리포트 종류)마다 컬럼명이 달라질 수 있어서
  // (예: 쿠팡 "날짜" / 네이버 "일자", 혹은 같은 쿠팡이라도 상품별로 다름),
  // 파일별로 실제로 어떤 컬럼이 날짜로 쓰이고 있는지 보여준다.
  function renderDateHint() {
    const hint = el("dateColumnHint");
    if (!hint) return;
    const files = state.manifest.files.filter((f) => state.selectedPaths.has(f.path));
    const parts = files.map((f) => {
      const col = resolveDateColumn(f.path);
      return `${f.name} → ${col ? col : "인식 실패"}`;
    });
    const MAX_SHOWN = 4;
    const shown = parts.slice(0, MAX_SHOWN).join(" · ");
    const extra = parts.length > MAX_SHOWN ? ` 외 ${parts.length - MAX_SHOWN}개` : "";
    hint.textContent = parts.length ? `현재 매칭: ${shown}${extra}` : "";
  }

  function renderSummary() {
    const hasMetrics = state.rows.length > 0 && state.metricColumns.size > 0;
    el("summaryCard").style.display = hasMetrics ? "" : "none";
    if (!hasMetrics) return;

    const grid = el("statGrid");
    grid.innerHTML = "";
    for (const metric of state.metricColumns) {
      const totals = { coupang: 0, naver: 0 };
      for (const row of state.rows) {
        const n = parseNumberLoose(row[metric]);
        if (Number.isFinite(n) && totals[row.__platform] != null) totals[row.__platform] += n;
      }
      const total = totals.coupang + totals.naver;

      const tile = document.createElement("div");
      tile.className = "stat-tile";
      tile.innerHTML = `
        <div class="stat-label">${escapeHtml(metric)} 합계</div>
        <div class="stat-value">${formatNumber(total)}</div>
        <div class="stat-breakdown">
          <span><span class="legend-dot dot-coupang"></span>쿠팡 ${formatNumber(totals.coupang)}</span>
          <span><span class="legend-dot dot-naver"></span>네이버 ${formatNumber(totals.naver)}</span>
        </div>`;
      grid.appendChild(tile);
    }
  }

  let charts = [];

  function renderCharts() {
    const hasAnyDate = filesInRows().some((fp) => resolveDateColumn(fp));
    const hasChart = state.rows.length > 0 && state.metricColumns.size > 0 && hasAnyDate;
    el("chartCard").style.display = hasChart ? "" : "none";
    for (const c of charts) c.destroy();
    charts = [];
    if (!hasChart) return;

    const area = el("chartsArea");
    area.innerHTML = "";

    for (const metric of state.metricColumns) {
      // 날짜 x 플랫폼 별 합계 집계 (row.__date는 플랫폼별로 감지된 날짜 컬럼 기준으로 이미 정규화되어 있음)
      const byDate = new Map(); // date -> {coupang, naver}
      for (const row of state.rows) {
        const d = row.__date;
        if (!d) continue;
        const n = parseNumberLoose(row[metric]);
        if (!Number.isFinite(n)) continue;
        if (!byDate.has(d)) byDate.set(d, { coupang: 0, naver: 0 });
        byDate.get(d)[row.__platform] += n;
      }
      const dates = [...byDate.keys()].sort();

      const wrap = document.createElement("div");
      wrap.style.marginBottom = "18px";
      const title = document.createElement("div");
      title.style.fontSize = "12.5px";
      title.style.color = "var(--text-secondary)";
      title.style.marginBottom = "6px";
      title.textContent = metric;
      wrap.appendChild(title);
      const chartWrap = document.createElement("div");
      chartWrap.className = "chart-wrap";
      const canvas = document.createElement("canvas");
      chartWrap.appendChild(canvas);
      wrap.appendChild(chartWrap);
      area.appendChild(wrap);

      if (dates.length === 0) {
        chartWrap.innerHTML = `<div class="empty-state">날짜를 인식하지 못했습니다. 위 "날짜로 쓸 컬럼"에서 다른 컬럼을 선택해 보세요.</div>`;
        continue;
      }

      const platformsPresent = platformsInRows();
      const colorFor = { coupang: cssVar("--series-1"), naver: cssVar("--series-2") };
      const datasets = platformsPresent.map((p) => ({
        label: PLATFORM_LABEL[p],
        data: dates.map((d) => byDate.get(d)[p]),
        borderColor: colorFor[p],
        backgroundColor: colorFor[p],
        borderWidth: 2,
        pointRadius: 2,
        pointHoverRadius: 5,
        tension: 0.15,
      }));

      const chart = new Chart(canvas.getContext("2d"), {
        type: "line",
        data: { labels: dates, datasets },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: "index", intersect: false },
          plugins: {
            legend: { display: datasets.length > 1, labels: { color: cssVar("--text-secondary") } },
            tooltip: { mode: "index", intersect: false },
          },
          scales: {
            x: {
              grid: { color: cssVar("--gridline") },
              ticks: { color: cssVar("--text-muted") },
            },
            y: {
              beginAtZero: true,
              grid: { color: cssVar("--gridline") },
              ticks: { color: cssVar("--text-muted") },
            },
          },
        },
      });
      charts.push(chart);
    }
  }

  function getFilteredSortedRows() {
    let rows = state.rows;
    if (state.search) {
      rows = rows.filter((row) =>
        Object.values(row).some((v) => String(v ?? "").toLowerCase().includes(state.search))
      );
    }
    if (state.sortColumn) {
      const col = state.sortColumn;
      rows = [...rows].sort((a, b) => {
        const av = a[col];
        const bv = b[col];
        const an = parseNumberLoose(av);
        const bn = parseNumberLoose(bv);
        let cmp;
        if (!Number.isNaN(an) && !Number.isNaN(bn)) cmp = an - bn;
        else cmp = String(av ?? "").localeCompare(String(bv ?? ""), "ko");
        return cmp * state.sortDir;
      });
    }
    return rows;
  }

  function renderTable() {
    const hasData = state.rows.length > 0;
    el("tableCard").style.display = hasData ? "" : "none";
    if (!hasData) return;

    const rows = getFilteredSortedRows();
    const displayCols = ["__platform", "__file", ...state.columns];

    const head = el("tableHead");
    head.innerHTML = "";
    const tr = document.createElement("tr");
    for (const col of displayCols) {
      const th = document.createElement("th");
      th.textContent = col === "__platform" ? "플랫폼" : col === "__file" ? "파일" : col;
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
      for (const col of displayCols) {
        const td = document.createElement("td");
        if (col === "__platform") {
          td.innerHTML = `<span class="platform-badge ${row.__platform}">${PLATFORM_LABEL[row.__platform] || row.__platform}</span>`;
        } else if (col === "__file") {
          td.textContent = row.__file;
        } else {
          td.textContent = row[col] ?? "";
        }
        tr2.appendChild(td);
      }
      body.appendChild(tr2);
    }

    el("pageInfo").textContent = `총 ${rows.length.toLocaleString("ko-KR")}행 · ${state.page} / ${totalPages} 페이지`;
    el("prevPageBtn").disabled = state.page <= 1;
    el("nextPageBtn").disabled = state.page >= totalPages;
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

  // ---------- 이벤트 ----------

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
    return rows.map((row) => {
      const out = { 플랫폼: PLATFORM_LABEL[row.__platform] || row.__platform, 파일: row.__file };
      for (const col of state.columns) out[col] = row[col] ?? "";
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
    triggerBlobDownload(blob, `ad_report_${new Date().toISOString().slice(0, 10)}.csv`);
  });

  el("downloadFilteredXlsxBtn").addEventListener("click", () => {
    const worksheet = XLSX.utils.json_to_sheet(buildDownloadRows());
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "리포트");
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
    renderCharts();
  });
  applyStoredTheme();

  loadManifest();
})();
