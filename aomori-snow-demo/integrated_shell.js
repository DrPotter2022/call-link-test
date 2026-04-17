(function () {
  'use strict';

  const shellState = {
    activeTab: 'panelFundSearch',
    frameReady: { app1: false, app2: false },
    portfolioItems: [],
    portfolioSource: '',
    drag: null,
    localMode: window.location.protocol === 'file:',
    buyOnlyMode: true,
    frontier: [],
    scatterDragActive: false,
    drawdownTolerance: -20,
    hiddenFunds: new Set(),
    scatterHidden: new Set(),
    goalPlannerMode: 'manual',
    manualTargetReturnChoice: 'mid',
    investmentMode: 'lump',
    savingsExtraCashEdited: false,
    savingsYearsEdited: false,
    lumpYearsBase: 10,
    syncMarks: {
      app1: null,
      app2: null,
      app3: null
    }
  };

  const els = {};

  function byId(id) {
    return document.getElementById(id);
  }

  function decodeBase64Utf8(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return new TextDecoder('utf-8').decode(bytes);
  }

  function resizeIframe(iframe) {
    if (!iframe || !iframe.contentDocument) return;
    const doc = iframe.contentDocument;
    const body = doc.body;
    const root = doc.documentElement;
    if (!body || !root) return;
    const nextHeight = Math.max(
      body.scrollHeight || 0,
      body.offsetHeight || 0,
      root.scrollHeight || 0,
      root.offsetHeight || 0,
      960
    );
    iframe.style.height = `${nextHeight + 12}px`;
  }

  function preprocessFrameHtml(html) {
    let next = String(html || '');
    const insertBeforeClosingTag = (source, tagName, snippet) => {
      const lower = source.toLowerCase();
      const closeTag = `</${tagName.toLowerCase()}>`;
      const idx = lower.lastIndexOf(closeTag);
      if (idx >= 0) return source.slice(0, idx) + snippet + source.slice(idx);
      return source + snippet;
    };
    if (!/<base\s/i.test(next)) {
      next = next.replace(/<head([^>]*)>/i, '<head$1><base href="about:srcdoc" target="_self">');
    }
    if (!/codexUnifiedFrameTheme/.test(next)) {
      const unifiedTheme = [
        '<style id="codexUnifiedFrameTheme">',
        '  :root {',
        '    --codex-bg: #eef3f9;',
        '    --codex-card: #ffffff;',
        '    --codex-border: #b8ccdf;',
        '    --codex-text: #0f2f4f;',
        '    --codex-muted: #5d7288;',
        '    --codex-primary: #2f6d9d;',
        '    --codex-primary-2: #2a88ac;',
        '  }',
        '  html, body { background: var(--codex-bg) !important; color: var(--codex-text) !important; }',
        '  body { font-family: "Segoe UI", "Yu Gothic UI", "Hiragino Kaku Gothic ProN", sans-serif !important; }',
        '  .app-shell, .app-root, .appRoot, .workspace, .main, .main-shell, .container, .tool-shell {',
        '    background: transparent !important;',
        '    color: var(--codex-text) !important;',
        '  }',
        '  .page-shell, .app-layout, .main-stack, .left-stack { background: transparent !important; }',
        '  header, .app-header, .top-header, .hero-header {',
        '    background: linear-gradient(135deg, #2f6d9d, #356e9f) !important;',
        '    color: #ffffff !important;',
        '    box-shadow: none !important;',
        '  }',
        '  .header-sticky-mount .sticky-main-card, .header-sticky-mount .sticky-main-card.section-collapsed {',
        '    background: linear-gradient(135deg, #2f6d9d, #356e9f) !important;',
        '    color: #ffffff !important;',
        '    border-color: #2f6d9d !important;',
        '  }',
        '  .header-sticky-mount .sticky-main-card .section-header {',
        '    background: rgba(255,255,255,0.14) !important;',
        '    color: #ffffff !important;',
        '    border-color: rgba(255,255,255,0.25) !important;',
        '  }',
        '  .header-sticky-mount .sticky-main-card h1, .header-sticky-mount .sticky-main-card h2, .header-sticky-mount .sticky-main-card h3, .header-sticky-mount .sticky-main-card .muted, .header-sticky-mount .sticky-main-card p, .header-sticky-mount .sticky-main-card .card-title-text, .header-sticky-mount .sticky-main-card .section-title, .header-sticky-mount .sticky-main-card .sub {',
        '    color: #ffffff !important;',
        '  }',
        '  .section, .panel, .card, .box, .block, .content-card, .sticky-main-card, .left-panel, .right-panel {',
        '    background: var(--codex-card) !important;',
        '    border: 1px solid var(--codex-border) !important;',
        '    border-radius: 14px !important;',
        '    box-shadow: none !important;',
        '  }',
        '  .left-panel, #leftPanel {',
        '    box-sizing: border-box !important;',
        '    overflow: hidden !important;',
        '  }',
        '  .left-panel > *, #leftPanel > * {',
        '    max-width: 100% !important;',
        '    box-sizing: border-box !important;',
        '  }',
        '  .left-panel .card, .left-panel .section, #leftPanel .card, #leftPanel .section {',
        '    width: 100% !important;',
        '    max-width: 100% !important;',
        '    overflow: hidden !important;',
        '  }',
        '  .header-sticky-mount, .sticky-main-card, .dashboard-shell, .main-wrapper {',
        '    background: #f3f7fc !important;',
        '    border-color: var(--codex-border) !important;',
        '  }',
        '  .section-header, .panel-title, .header, .toolbar, .title-row, .app-titlebar, .panel-head {',
        '    background: #f7fbff !important;',
        '    color: var(--codex-text) !important;',
        '    border-color: var(--codex-border) !important;',
        '  }',
        '  h1, h2, h3, h4, h5, h6, .title, .heading { color: var(--codex-text) !important; }',
        '  p, small, .muted, .caption, .sub, .note { color: var(--codex-muted) !important; }',
        '  header h1, header h2, header h3, header p, header .muted, .app-header h1, .app-header p { color: #ffffff !important; }',
        '  button, .btn, .small-btn, input[type="button"], input[type="submit"] {',
        '    background: #f7fbff !important;',
        '    color: var(--codex-text) !important;',
        '    border: 1px solid var(--codex-border) !important;',
        '    border-radius: 12px !important;',
        '    box-shadow: none !important;',
        '  }',
        '  button:hover, .btn:hover, .small-btn:hover { background: #eaf3fc !important; }',
        '  input, select, textarea {',
        '    background: #ffffff !important;',
        '    color: var(--codex-text) !important;',
        '    border: 1px solid var(--codex-border) !important;',
        '    border-radius: 10px !important;',
        '  }',
        '  input[type="date"] { color-scheme: light !important; }',
        '  input[type="date"]::-webkit-calendar-picker-indicator { opacity: 1 !important; filter: none !important; }',
        '  table, .table-wrap, .grid-table { background: #ffffff !important; color: var(--codex-text) !important; }',
        '  th, td { border-color: #d8e5f2 !important; }',
        '  .plotly-graph-div, .chart, .chart-card, .plot-wrap { background: #ffffff !important; }',
        '</style>'
      ].join('');
      if (/<\/head>/i.test(next)) next = insertBeforeClosingTag(next, 'head', unifiedTheme);
      else next = `${unifiedTheme}${next}`;
    }
    // NOTE: bridge script injection into srcdoc can break parsing on some environments.
    // Keep disabled and use parent-side direct state bridge instead.
    if (false && !/codexLocalFrameGuard/.test(next)) {
      const guard = [
        '<script>',
        '(function(){',
        '  function blocked(href){ return !href || href === "#" || href.startsWith("#") || href.startsWith("file:"); }',
        '  window.addEventListener("click", function(event){',
        '    const link = event.target && event.target.closest ? event.target.closest("a[href]") : null;',
        '    if (!link) return;',
        '    const href = link.getAttribute("href") || "";',
        '    if (!blocked(href)) return;',
        '    event.preventDefault();',
        '    event.stopPropagation();',
        '    const id = href.replace(/^#/, "");',
        '    if (!id) return;',
        '    const target = document.getElementById(id);',
        '    if (target && target.scrollIntoView) target.scrollIntoView({ behavior: "smooth", block: "start" });',
        '  }, true);',
        '  document.addEventListener("submit", function(event){',
        '    const form = event.target;',
        '    const rawAction = form && form.getAttribute ? (form.getAttribute("action") || "") : "";',
        '    const resolvedAction = (function(){',
        '      try { return form && form.action ? String(form.action) : ""; } catch (e) { return ""; }',
        '    })();',
        '    if (!rawAction || rawAction.startsWith("file:") || resolvedAction.startsWith("file:")) {',
        '      event.preventDefault();',
        '      event.stopPropagation();',
        '    }',
        '  }, true);',
        '  function getScopedState(){',
        '    try { return (typeof state !== "undefined") ? state : null; } catch (e) { return null; }',
        '  }',
        '  if (!window.__codexBridgeInstalled) {',
        '    window.__codexGetState = function(){ return getScopedState(); };',
        '    window.__codexGetWorkbookRows = function(){',
        '      const s = getScopedState();',
        '      return (s && Array.isArray(s.workbookRows)) ? s.workbookRows : [];',
        '    };',
        '    window.__codexGetFundSeriesAll = function(){',
        '      const s = getScopedState();',
        '      return (s && Array.isArray(s.fundSeriesAll)) ? s.fundSeriesAll : [];',
        '    };',
        '    window.__codexGetSeriesColors = function(){',
        '      const s = getScopedState();',
        '      return (s && s.seriesColors) ? s.seriesColors : {};',
        '    };',
        '    window.__codexSetWorkbookRows = function(rows){',
        '      const s = getScopedState();',
        '      if (!s) return false;',
        '      s.workbookRows = Array.isArray(rows) ? rows : [];',
        '      s.workbookAnalysis = null;',
        '      s.workbookLastError = "";',
        '      return true;',
        '    };',
        '    window.__codexSetFundSeriesAll = function(series){',
        '      const s = getScopedState();',
        '      if (!s) return false;',
        '      s.fundSeriesAll = Array.isArray(series) ? series : [];',
        '      return true;',
        '    };',
        '    window.__codexSetBenchmarkRows = function(rows){',
        '      const s = getScopedState();',
        '      if (!s) return false;',
        '      const list = Array.isArray(rows) ? rows : [];',
        '      s.benchmarkRows = list;',
        '      s.hasBenchmark = list.length > 0;',
        '      if (s.bm && typeof s.bm === "object") {',
        '        s.bm.rows = list.map(function(r){',
        '          const dateISO = String((r && r.date) || "").trim();',
        '          const d = new Date(dateISO);',
        '          const nav = Number(r && (r.price != null ? r.price : r.nav));',
        '          if (!dateISO || !Number.isFinite(nav) || Number.isNaN(d.getTime())) return null;',
        '          return { dateISO: dateISO, date: d, nav: nav };',
        '        }).filter(Boolean);',
        '        if (list.length && !s.bm.name) s.bm.name = "Fund Search BM";',
        '      }',
        '      return true;',
        '    };',
        '    window.__codexMarkFundSearchImport = function(summary){',
        '      const fundCount = Number((summary && summary.fundCount) || 0);',
        '      const bmCount = Number((summary && summary.benchmarkCount) || 0);',
        '      const ts = String((summary && summary.timestamp) || "");',
        '      const id = "codex-fund-search-import-status";',
        '      let el = document.getElementById(id);',
        '      if (!el) {',
        '        el = document.createElement("div");',
        '        el.id = id;',
        '        el.style.margin = "8px 12px";',
        '        el.style.padding = "8px 10px";',
        '        el.style.border = "1px solid #b6d4fe";',
        '        el.style.background = "#eef6ff";',
        '        el.style.color = "#0b3b75";',
        '        el.style.borderRadius = "8px";',
        '        el.style.fontSize = "12px";',
        '        el.style.lineHeight = "1.5";',
        '        const root = document.querySelector(".app-shell, .container, .main, body");',
        '        if (root && root.firstChild) root.insertBefore(el, root.firstChild);',
        '        else if (root) root.appendChild(el);',
        '      }',
        '      const stamp = ts ? ` / ${ts}` : "";',
        '      el.textContent = `Fund Search 受信済み: ファンドCSV=${fundCount}件 / ベンチマークCSV=${bmCount}件${stamp}`;',
        '      return true;',
        '    };',
        '    window.__codexNormalizeFundHistoryRows = function(payload){',
        '      const code = String((payload && payload.code) || "").trim();',
        '      const name = String((payload && payload.name) || code || "").trim();',
        '      const company = String((payload && payload.company) || "").trim();',
        '      const history = Array.isArray(payload && payload.history) ? payload.history : [];',
        '      const rows = history.map(function(h){',
        '        const dateISO = String((h && h.date) || "").trim();',
        '        const dateObj = new Date(dateISO);',
        '        const nav = Number(h && h.nav);',
        '        if (!dateISO || !Number.isFinite(nav) || Number.isNaN(dateObj.getTime())) return null;',
        '        return { dateISO: dateISO, date: dateObj, nav: nav };',
        '      }).filter(Boolean).sort(function(a,b){ return a.date - b.date; });',
        '      return { kind: "fund", code: code, name: name, company: company, rows: rows };',
        '    };',
        '    window.__codexFetchFundDetail = async function(code){',
        '      const c = String(code || "").trim();',
        '      if (!c) throw new Error("code is required");',
        '      const baseHref = (window.parent && window.parent.location && window.parent.location.href) ? window.parent.location.href : window.location.href;',
        '      const url = new URL("project/public/fund_detail/" + encodeURIComponent(c) + ".json", baseHref).href;',
        '      const resp = await fetch(url, { cache: "no-store" });',
        '      if (resp.status === 404) throw new Error("detail file not found: " + c);',
        '      if (!resp.ok) throw new Error("detail fetch failed: HTTP " + resp.status);',
        '      const payload = await resp.json();',
        '      return window.__codexNormalizeFundHistoryRows(payload);',
        '    };',
        '    window.__codexFetchFundDetails = async function(codes){',
        '      const list = Array.isArray(codes) ? codes : [];',
        '      const out = [];',
        '      for (const code of list) {',
        '        try { out.push(await window.__codexFetchFundDetail(code)); } catch (e) {',
        '          out.push({ kind: "fund", code: String(code || ""), name: String(code || ""), company: "", rows: [], error: String((e && e.message) || e || "error") });',
        '        }',
        '      }',
        '      return out;',
        '    };',
        '    window.__codexToWorkbookRowsFromApiFunds = function(apiFunds){',
        '      const list = Array.isArray(apiFunds) ? apiFunds : [];',
        '      return list.map(function(f){',
        '        const rows = Array.isArray(f && f.rows) ? f.rows : [];',
        '        const first = rows.length ? rows[0].dateISO : "";',
        '        const last = rows.length ? rows[rows.length - 1].dateISO : "";',
        '        const latest = rows.length ? rows[rows.length - 1].nav : 0;',
        '        return {',
        '          fund: String((f && f.name) || (f && f.code) || ""),',
        '          code: String((f && f.code) || ""),',
        '          company: String((f && f.company) || ""),',
        '          periodStart: first,',
        '          periodEnd: last,',
        '          totalInvestment: 0,',
        '          recurringInvestment: 0,',
        '          spotInvestment: 0,',
        '          valuation: Number.isFinite(latest) ? latest : 0,',
        '          annualVol: NaN,',
        '          maxDrawdown: NaN,',
        '          sharpe: NaN,',
        '          beta: NaN,',
        '          trackingError: NaN,',
        '          historyCount: rows.length',
        '        };',
        '      });',
        '    };',
        '    window.__codexImportApiFundsToState = async function(codes){',
        '      const s = getScopedState();',
        '      if (!s) return { ok: false, reason: "state missing" };',
        '      const apiFunds = await window.__codexFetchFundDetails(codes);',
        '      if (Array.isArray(s.funds)) {',
        '        if (!s.api || typeof s.api !== "object") s.api = {};',
        '        if (!(s.api.selectedCodes instanceof Set)) s.api.selectedCodes = new Set();',
        '        if (!(s.api.selectedFunds instanceof Map)) s.api.selectedFunds = new Map();',
        '        if (!(s.api.loadingDetails instanceof Set)) s.api.loadingDetails = new Set();',
        '        if (!(s.api.detailErrors instanceof Map)) s.api.detailErrors = new Map();',
        '        apiFunds.forEach(function(f){',
        '          const code = String((f && f.code) || "");',
        '          if (!code) return;',
        '          const rows = (Array.isArray(f && f.rows) ? f.rows : []).map(function(r){',
        '            const dateISO = String((r && r.dateISO) || (r && r.date) || "").trim();',
        '            const d = new Date(dateISO);',
        '            const nav = Number(r && r.nav);',
        '            if (!dateISO || !Number.isFinite(nav) || Number.isNaN(d.getTime())) return null;',
        '            return { dateISO: dateISO, date: d, nav: nav };',
        '          }).filter(Boolean);',
        '          const model = {',
        '            kind: "fund",',
        '            code: code,',
        '            name: String((f && f.name) || code),',
        '            company: String((f && f.company) || ""),',
        '            rows: rows',
        '          };',
        '          s.api.selectedCodes.add(code);',
        '          s.api.selectedFunds.set(code, model);',
        '          const exists = s.funds.some(function(x){ return x && x.source === "api" && x.sourceCode === code; });',
        '          if (!exists) {',
        '            const first = rows[0];',
        '            const last = rows.length ? rows[rows.length - 1] : null;',
        '            const slot = s.funds.find(function(x){',
        '              const nameEmpty = !String((x && x.name) || "").trim() || String((x && x.name) || "").trim() === "（名称未設定）";',
        '              const rowsEmpty = !Array.isArray(x && x.rows) || (x.rows || []).length === 0;',
        '              const untouched = !x || !x.source || x.source === "manual";',
        '              return rowsEmpty && (nameEmpty || untouched);',
        '            });',
        '            const payload = {',
        '              id: (slot && slot.id) ? slot.id : ((window.crypto && window.crypto.randomUUID) ? window.crypto.randomUUID() : ("api-" + code + "-" + Date.now())),',
        '              name: model.name,',
        '              rows: rows,',
        '              fileName: "API:" + code,',
        '              fileStatus: rows.length ? ("API:" + code + "：" + rows.length + "件（" + (first ? first.dateISO : "") + " ～ " + (last ? last.dateISO : "") + "）") : ("API:" + code + "：0件"),',
        '              investAmount: Number((slot && slot.investAmount) || 0),',
        '              lumpBuyDate: String((slot && slot.lumpBuyDate) || ""),',
        '              sipAmount: Number((slot && slot.sipAmount) || 10000),',
        '              sipDay: Number((slot && slot.sipDay) || 1),',
        '              sipStartDate: String((slot && slot.sipStartDate) || ""),',
        '              visible: true,',
        '              lastEncoding: "api",',
        '              source: "api",',
        '              sourceCode: code,',
        '              sourceKind: "fund"',
        '            };',
        '            if (slot) Object.assign(slot, payload);',
        '            else s.funds.push(payload);',
        '          }',
        '        });',
        '        return { ok: true, mode: "app1", count: apiFunds.length, withHistory: apiFunds.filter(function(f){ return (f.rows || []).length > 0; }).length };',
        '      }',
        '      const fundSeries = apiFunds.map(function(f){ return {',
        '        name: f.name || f.code,',
        '        key: "api_" + String(f.code || ""),',
        '        kind: f.kind || "fund",',
        '        code: f.code || "",',
        '        company: f.company || "",',
        '        rows: Array.isArray(f.rows) ? f.rows : [],',
        '        visibleRows: Array.isArray(f.rows) ? f.rows.map(function(r){ return { date: r.dateISO, nav: r.nav }; }) : []',
        '      }; });',
        '      const workbookRows = window.__codexToWorkbookRowsFromApiFunds(apiFunds);',
        '      s.fundSeriesAll = fundSeries;',
        '      s.workbookRows = workbookRows;',
        '      s.workbookAnalysis = null;',
        '      s.workbookLastError = "";',
        '      return { ok: true, count: apiFunds.length, withHistory: apiFunds.filter(function(f){ return (f.rows || []).length > 0; }).length };',
        '    };',
        '    try {',
        '      const origAssign = window.location.assign.bind(window.location);',
        '      const origReplace = window.location.replace.bind(window.location);',
        '      window.location.assign = function(url){ if (String(url || "").startsWith("file:")) return; origAssign(url); };',
        '      window.location.replace = function(url){ if (String(url || "").startsWith("file:")) return; origReplace(url); };',
        '    } catch (e) {}',
        '    try {',
        '      const origOpen = window.open;',
        '      window.open = function(url){ if (String(url || "").startsWith("file:")) return null; return origOpen.apply(window, arguments); };',
        '    } catch (e) {}',
        '    try {',
        '      const patchPlotly = function(){',
        '        if (!window.Plotly || window.__codexPatchedPlotly) return !!window.__codexPatchedPlotly;',
        '        const origNewPlot = window.Plotly.newPlot.bind(window.Plotly);',
        '        window.Plotly.newPlot = function(gd, data, layout, config){',
        '          try {',
        '            const id = (gd && gd.id) ? String(gd.id) : "";',
        '            const traces = Array.isArray(data) ? data : [];',
        '            const hasMarkerTrace = traces.some(function(tr){ return /markers/i.test(String((tr && tr.mode) || "")); });',
        '            if (/scatter|risk|rolling/i.test(id) || hasMarkerTrace) {',
        '              let minY = Infinity;',
        '              let maxY = -Infinity;',
        '              traces.forEach(function(tr){',
        '                (Array.isArray(tr && tr.y) ? tr.y : []).forEach(function(v){',
        '                  const n = Number(v);',
        '                  if (Number.isFinite(n)) { if (n < minY) minY = n; if (n > maxY) maxY = n; }',
        '                });',
        '              });',
        '              if (Number.isFinite(minY) && Number.isFinite(maxY)) {',
        '                const span = Math.max(1, maxY - minY);',
        '                const low = minY - span * 0.12;',
        '                const high = maxY + span * 0.12;',
        '                layout = layout || {};',
        '                layout.yaxis = layout.yaxis || {};',
        '                layout.xaxis = layout.xaxis || {};',
        '                layout.yaxis.automargin = false;',
        '                layout.xaxis.automargin = false;',
        '                layout.yaxis.range = [low, high];',
        '                if (/rollingchart/i.test(id)) {',
        '                  layout.margin = Object.assign({ l: 54, r: 24, t: 44, b: 44 }, layout.margin || {});',
        '                }',
        '              }',
        '            }',
        '          } catch (e) {}',
        '          if (window.Plotly.react && gd && typeof gd === "string" && document.getElementById(gd)) {',
        '            return window.Plotly.react(gd, data, layout, config);',
        '          }',
        '          if (window.Plotly.react && gd && gd.id && document.getElementById(gd.id)) {',
        '            return window.Plotly.react(gd, data, layout, config);',
        '          }',
        '          return origNewPlot(gd, data, layout, config);',
        '        };',
        '        window.__codexPatchedPlotly = true;',
        '        return true;',
        '      };',
        '      if (!patchPlotly()) {',
        '        let tries = 0;',
        '        const timer = setInterval(function(){',
        '          tries += 1;',
        '          if (patchPlotly() || tries > 40) clearInterval(timer);',
        '        }, 250);',
        '      }',
        '    } catch (e) {}',
        '    try {',
        '      const remapInvestorNames = function(){',
        '        const s = getScopedState();',
        '        if (!s || !Array.isArray(s.workbookRows) || !Array.isArray(s.fundSeriesAll)) return;',
        '        const byIdx = [];',
        '        const bmName = String((s && (s.benchmarkName || s.bmName)) || "").trim();',
        '        s.workbookRows.forEach(function(r){',
        '          const n = (r && (r.fund || r.name)) ? String(r.fund || r.name) : "";',
        '          if (!n) return;',
        '          byIdx.push(n);',
        '        });',
        '        Array.from(document.querySelectorAll(".additional-fund-name")).forEach(function(inp){',
        '          const n = String((inp && inp.value) || "").trim();',
        '          if (n) byIdx.push(n);',
        '        });',
        '        s.fundSeriesAll.forEach(function(series){',
        '          const n = String((series && series.name) || "").trim();',
        '          if (!n) return;',
        '          if (bmName && n.toLowerCase() === bmName.toLowerCase()) return;',
        '          if (/^(?:投資家|investor)\\s*\\d+$/i.test(n)) return;',
        '          if (!byIdx.includes(n)) byIdx.push(n);',
        '        });',
        '        s.fundSeriesAll.forEach(function(series){',
        '          if (!series || !series.name) return;',
        '          const m = String(series.name).match(/^(?:投資家|investor)\\s*(\\d+)$/i);',
        '          if (!m) return;',
        '          const idx = Math.max(0, Number(m[1]) - 1);',
        '          const mapped = byIdx[idx];',
        '          if (mapped) series.name = mapped;',
        '        });',
        '        if (s.seriesColors && typeof s.seriesColors === "object") {',
        '          Object.keys(s.seriesColors).forEach(function(k){',
        '            const m = String(k).match(/^(?:投資家|investor)\\s*(\\d+)$/i);',
        '            if (!m) return;',
        '            const idx = Math.max(0, Number(m[1]) - 1);',
        '            const mapped = byIdx[idx];',
        '            if (mapped && !s.seriesColors[mapped]) s.seriesColors[mapped] = s.seriesColors[k];',
        '          });',
        '        }',
        '      };',
        '      setTimeout(remapInvestorNames, 100);',
        '      setTimeout(remapInvestorNames, 800);',
        '      setInterval(remapInvestorNames, 2000);',
        '      const refreshPaletteLabels = function(){',
        '        try {',
        '          const s = getScopedState();',
        '          if (!s) return;',
        '          const names = [];',
        '          (Array.isArray(s.workbookRows) ? s.workbookRows : []).forEach(function(r){',
        '            const n = String((r && (r.fund || r.name)) || "").trim();',
        '            if (n && names.indexOf(n) < 0) names.push(n);',
        '          });',
        '          Array.from(document.querySelectorAll(".additional-fund-name")).forEach(function(inp){',
        '            const n = String((inp && inp.value) || "").trim();',
        '            if (n && names.indexOf(n) < 0) names.push(n);',
        '          });',
        '          const labels = Array.from(document.querySelectorAll(".seriesColorName, .legend-name, .series-name"));',
        '          labels.forEach(function(el){',
        '            const t = String((el && el.textContent) || "").trim();',
        '            const m = t.match(/^(?:投資家|investor)\\s*(\\d+)$/i);',
        '            if (!m) return;',
        '            const idx = Math.max(0, Number(m[1]) - 1);',
        '            if (names[idx]) el.textContent = names[idx];',
        '          });',
        '        } catch (e) {}',
        '      };',
        '      const syncPaletteKeyNames = function(){',
        '        try {',
        '          const s = getScopedState();',
        '          if (!s) return;',
        '          const workbookNames = (Array.isArray(s.workbookRows) ? s.workbookRows : [])',
        '            .map(function(r){ return String((r && (r.fund || r.name)) || "").trim(); })',
        '            .filter(Boolean);',
        '          if (!workbookNames.length) return;',
        '          if (s.seriesColors && typeof s.seriesColors === "object") {',
        '            Object.keys(s.seriesColors).forEach(function(k){',
        '              const m = String(k).match(/^(?:投資家|investor)\\s*(\\d+)$/i);',
        '              if (!m) return;',
        '              const idx = Math.max(0, Number(m[1]) - 1);',
        '              const mapped = workbookNames[idx];',
        '              if (mapped && !s.seriesColors[mapped]) s.seriesColors[mapped] = s.seriesColors[k];',
        '            });',
        '          }',
        '          const controls = document.querySelectorAll("[data-series-color]");',
        '          controls.forEach(function(inp){',
        '            const key = String(inp.getAttribute("data-series-color") || "").trim();',
        '            const m = key.match(/^(?:投資家|investor)\\s*(\\d+)$/i);',
        '            if (!m) return;',
        '            const idx = Math.max(0, Number(m[1]) - 1);',
        '            const mapped = workbookNames[idx];',
        '            if (!mapped) return;',
        '            inp.setAttribute("data-series-color", mapped);',
        '            const row = inp.closest(".color-row");',
        '            if (row) {',
        '              const nm = row.querySelector(".color-row-name");',
        '              if (nm) nm.textContent = mapped;',
        '              const txt = row.querySelector("[data-series-color-text]");',
        '              if (txt) txt.setAttribute("data-series-color-text", mapped);',
        '              const rst = row.querySelector("[data-series-reset]");',
        '              if (rst) rst.setAttribute("data-series-reset", mapped);',
        '            }',
        '          });',
        '          // avoid forcing chart redraw on every palette sync tick',
        '        } catch (e) {}',
        '      };',
        '      setTimeout(refreshPaletteLabels, 250);',
        '      setTimeout(refreshPaletteLabels, 1200);',
        '      setTimeout(syncPaletteKeyNames, 300);',
        '      setTimeout(syncPaletteKeyNames, 1400);',
        '      const normalizeApp2Layout = function(){',
        '        try {',
        '          const root = document.documentElement;',
        '          const left = document.getElementById("leftPanel") || document.querySelector(".left-panel");',
        '          const main = document.getElementById("mainPanel") || document.querySelector(".main-panel");',
        '          const vw = Math.max(320, window.innerWidth || 1280);',
        '          if (vw < 1500) {',
        '            root.style.setProperty("--left-panel-width", "480px");',
        '            if (left) { left.style.width = "480px"; left.style.flexBasis = "480px"; left.style.maxWidth = "480px"; }',
        '          } else {',
        '            root.style.setProperty("--left-panel-width", "560px");',
        '            if (left) { left.style.width = "560px"; left.style.flexBasis = "560px"; left.style.maxWidth = "560px"; }',
        '          }',
        '          if (vw <= 900 && left) {',
        '            root.style.setProperty("--left-panel-width", "100%");',
        '            left.style.width = "100%";',
        '            left.style.flexBasis = "100%";',
        '            left.style.maxWidth = "100%";',
        '          }',
        '          if (main) { main.style.minWidth = "0"; main.style.width = "auto"; }',
        '          document.body.style.overflowX = "hidden";',
        '          const charts = document.querySelectorAll("#drawdownChart, #rollingChart, [id^=rollingChart_], [id*=RiskReturn][id*=Chart], .chart-box");',
        '          charts.forEach(function(c){ c.style.width = "100%"; c.style.minWidth = "0"; });',
        '          const settingsSection = document.getElementById("settingsSection");',
        '          const findInputSettingsCard = function(){',
        '            const cards = Array.from(document.querySelectorAll(".card, .section, .panel, .sticky-main-card"));',
        '            for (const card of cards) {',
        '              const t = (card.querySelector(".section-title, .card-title-text, .chart-title, h2, h3") || {}).textContent || "";',
        '              if (String(t).indexOf("入力・設定") >= 0) return card;',
        '            }',
        '            return null;',
        '          };',
        '          const inputSettingsCard = findInputSettingsCard();',
        '          if (settingsSection) {',
        '            settingsSection.classList.remove("section-collapsed");',
        '            settingsSection.style.boxSizing = "border-box";',
        '            settingsSection.style.width = "100%";',
        '            settingsSection.style.maxWidth = "100%";',
        '            settingsSection.style.overflow = "hidden";',
        '            const panel = settingsSection.querySelector(".section-body, .card-body, .settings-body") || settingsSection;',
        '            if (panel) {',
        '              panel.style.boxSizing = "border-box";',
        '              panel.style.width = "100%";',
        '              panel.style.maxWidth = "100%";',
        '            }',
        '            const controls = settingsSection.querySelectorAll("input, select, textarea, button, .file-row, .file-input-row, .upload-row, .input-group, .form-row, .field-row");',
        '            controls.forEach(function(el){',
        '              if (!el || !el.style) return;',
        '              el.style.boxSizing = "border-box";',
        '              if (el.tagName === "INPUT" || el.tagName === "SELECT" || el.tagName === "TEXTAREA") {',
        '                el.style.maxWidth = "100%";',
        '              }',
        '            });',
        '          }',
        '          if (inputSettingsCard) {',
        '            inputSettingsCard.style.boxSizing = "border-box";',
        '            inputSettingsCard.style.width = "100%";',
        '            inputSettingsCard.style.maxWidth = "100%";',
        '            inputSettingsCard.style.overflow = "hidden";',
        '            inputSettingsCard.style.borderRadius = "14px";',
        '            const body = inputSettingsCard.querySelector(".body, .section-body, .card-body");',
        '            if (body) {',
        '              body.style.boxSizing = "border-box";',
        '              body.style.width = "100%";',
        '              body.style.maxWidth = "100%";',
        '              body.style.overflow = "hidden";',
        '            }',
        '          }',
        '          if (left) {',
        '            left.style.overflow = "hidden";',
        '            left.style.paddingRight = "0px";',
        '            if (vw <= 900) left.style.resize = "none";',
        '          }',
        '          const workbookChartCard = document.getElementById("workbookChartCard");',
        '          if (workbookChartCard) workbookChartCard.classList.remove("section-collapsed");',
        '          if (typeof window.syncLeftPanelWidth === "function") { try { window.syncLeftPanelWidth(); } catch (e) {} }',
        '          if (typeof window.syncPanelTopAlignment === "function") { try { window.syncPanelTopAlignment(); } catch (e) {} }',
        '        } catch (e) {}',
        '      };',
        '      setTimeout(normalizeApp2Layout, 120);',
        '      setTimeout(normalizeApp2Layout, 700);',
        '      window.addEventListener("resize", normalizeApp2Layout);',
        '      const patchApp2Scatter = function(){',
        '        if (window.__codexPatchedApp2Scatter) return true;',
        '        if (typeof window.renderAdditionalFundSections !== "function") return false;',
        '        const origRender = window.renderAdditionalFundSections;',
        '        window.renderAdditionalFundSections = function(){',
        '          const out = origRender.apply(this, arguments);',
        '          try {',
        '            const gd = document.getElementById("workbookRiskReturnChart") || document.getElementById("workbookScatterChart") || document.querySelector("[id*=RiskReturn][id*=Chart]");',
        '            if (gd && window.Plotly && typeof window.Plotly.relayout === "function") {',
        '              const s = getScopedState();',
        '              const rows = (s && Array.isArray(s.workbookRows)) ? s.workbookRows : [];',
        '              const riskTolInput = document.getElementById("riskTolerance");',
        '              const riskTol = riskTolInput ? Number(riskTolInput.value || 0) : 0;',
        '              const xs = rows.map(function(r){ return Number(r && (r.annualVol || r.vol || r.expectedRisk)); }).filter(Number.isFinite).map(function(v){ return Math.abs(v) > 1 ? v : v * 100; });',
        '              const ys = rows.map(function(r){ return Number(r && (r.returnPct || r.return)); }).filter(Number.isFinite).map(function(v){ return Math.abs(v) > 1 ? v : v * 100; });',
        '              if (ys.length) {',
        '                const minY = Math.min.apply(null, ys);',
        '                const maxY = Math.max.apply(null, ys);',
        '                const span = Math.max(5, maxY - minY);',
        '                const minX = xs.length ? Math.min.apply(null, xs) : 0;',
        '                const maxX = xs.length ? Math.max.apply(null, xs) : 30;',
        '                const xSpan = Math.max(5, maxX - minX);',
        '                window.Plotly.relayout(gd, {',
        '                  "xaxis.range":[Math.max(0, minX - xSpan * 0.15), maxX + xSpan * 0.15],',
        '                  "yaxis.range":[minY - span * 0.2, maxY + span * 0.2],',
        '                  "yaxis.automargin": false,',
        '                  "xaxis.automargin": false,',
        '                  "shapes":[{type:"line", x0:riskTol, x1:riskTol, y0:minY - span * 0.2, y1:maxY + span * 0.2, line:{color:"#ef4444", width:1.5, dash:"dot"}}]',
        '                });',
        '              }',
        '            }',
        '          } catch (e) {}',
        '          try {',
        '            const cards = Array.from(document.querySelectorAll("input, .seriesColorName, .color-name, .legend-name, .series-name"));',
        '            const s = getScopedState();',
        '            const names = (s && Array.isArray(s.workbookRows)) ? s.workbookRows.map(function(r){ return (r && (r.fund || r.name)) ? String(r.fund || r.name) : ""; }) : [];',
        '            cards.forEach(function(el){',
        '              const txt = (el.value || el.textContent || "").trim();',
        '              const m = txt.match(/^(?:投資家|investor)\\s*(\\d+)$/i);',
        '              if (!m) return;',
        '              const idx = Math.max(0, Number(m[1]) - 1);',
        '              if (!names[idx]) return;',
        '              if ("value" in el) el.value = names[idx];',
        '              else el.textContent = names[idx];',
        '            });',
        '            if (s && s.seriesColors && typeof s.seriesColors === "object") {',
        '              Object.keys(s.seriesColors).forEach(function(k){',
        '                const m = String(k).match(/^(?:投資家|investor)\\s*(\\d+)$/i);',
        '                if (!m) return;',
        '                const idx = Math.max(0, Number(m[1]) - 1);',
        '                const mapped = names[idx];',
        '                if (mapped && !s.seriesColors[mapped]) s.seriesColors[mapped] = s.seriesColors[k];',
        '              });',
        '              if (typeof window.renderSeriesColorControls === "function") { try { window.renderSeriesColorControls(); } catch (e) {} }',
        '            }',
        '          } catch (e) {}',
        '          return out;',
        '        };',
        '        window.__codexPatchedApp2Scatter = true;',
        '        return true;',
        '      };',
        '      if (!patchApp2Scatter()) {',
        '        let tries3 = 0;',
        '        const timer3 = setInterval(function(){',
        '          tries3 += 1;',
        '          if (patchApp2Scatter() || tries3 > 40) clearInterval(timer3);',
        '        }, 300);',
        '      }',
        '    } catch (e) {}',
        '    try {',
        '      const patchRollingRenderer = function(){',
        '        if (window.__codexPatchedRollingRenderer) return true;',
        '        if (typeof window.renderAdditionalFundSections !== "function") return false;',
        '        const orig = window.renderAdditionalFundSections;',
        '        window.renderAdditionalFundSections = function(){',
          '          try {',
          '            const s = getScopedState();',
          '            if (s && Array.isArray(s.fundSeriesAll)) {',
          '              let hasBenchLike = !!s.hasBenchmark;',
          '              s.fundSeriesAll.forEach(function(series){',
          '                const rows = (series && Array.isArray(series.visibleRows)) ? series.visibleRows : [];',
          '                rows.forEach(function(r){',
          '                  const rr = Number(r && r.rollingExcess);',
          '                  const er = Number(r && r.excessReturn);',
          '                  const fr = Number(r && r.fundReturn);',
          '                  if (!Number.isFinite(rr) && Number.isFinite(er)) r.rollingExcess = er;',
          '                  else if (!Number.isFinite(rr) && Number.isFinite(fr)) r.rollingExcess = fr;',
          '                  if (Number.isFinite(Number(r && r.benchReturn)) || Number.isFinite(Number(r && r.benchIndex)) || Number.isFinite(Number(r && r.excessReturn)) || Number.isFinite(Number(r && r.rollingExcess))) hasBenchLike = true;',
          '                });',
          '              });',
          '              if (hasBenchLike) s.hasBenchmark = true;',
          '            }',
          '          } catch (e) {}',
        '          return orig.apply(this, arguments);',
        '        };',
        '        window.__codexPatchedRollingRenderer = true;',
        '        return true;',
        '      };',
        '      if (!patchRollingRenderer()) {',
        '        let tries = 0;',
        '        const timer2 = setInterval(function(){',
        '          tries += 1;',
        '          if (patchRollingRenderer() || tries > 40) clearInterval(timer2);',
        '        }, 300);',
        '      }',
        '    } catch (e) {}',
        '    window.__codexBridgeInstalled = true;',
        '  }',
        '  window.codexLocalFrameGuard = true;',
        '})();',
        '<\/script>'
      ].join('');
      next = insertBeforeClosingTag(next, 'body', guard);
    }
    if (false && !/codexMiniStateBridge/.test(next)) {
      const miniBridge = [
        '<script id="codexMiniStateBridge">',
        '(function(){',
        '  try {',
        '    if (typeof window.__codexGetState !== "function") {',
        '      window.__codexGetState = function(){ try { return state; } catch (e) { return null; } };',
        '    }',
        '    if (typeof window.__codexSetWorkbookRows !== "function") {',
        '      window.__codexSetWorkbookRows = function(rows){',
        '        var s = window.__codexGetState();',
        '        if (!s) return false;',
        '        s.workbookRows = Array.isArray(rows) ? rows : [];',
        '        s.workbookAnalysis = null;',
        '        s.workbookLastError = "";',
        '        return true;',
        '      };',
        '    }',
        '    if (typeof window.__codexSetFundSeriesAll !== "function") {',
        '      window.__codexSetFundSeriesAll = function(series){',
        '        var s = window.__codexGetState();',
        '        if (!s) return false;',
        '        s.fundSeriesAll = Array.isArray(series) ? series : [];',
        '        return true;',
        '      };',
        '    }',
        '    if (typeof window.__codexSetBenchmarkRows !== "function") {',
        '      window.__codexSetBenchmarkRows = function(rows){',
        '        var s = window.__codexGetState();',
        '        if (!s) return false;',
        '        s.benchmarkRows = Array.isArray(rows) ? rows : [];',
        '        s.hasBenchmark = s.benchmarkRows.length > 0;',
        '        return true;',
        '      };',
        '    }',
        '    // Rolling表示補正は親側で実施（srcdoc内の描画関数パッチは無効化）',
        '  } catch (e) {}',
        '})();',
        '<\/script>'
      ].join('');
      next = insertBeforeClosingTag(next, 'body', miniBridge);
    }
    return next;
  }

  function setShellStatus(text) {
    if (els.shellStatusText) els.shellStatusText.textContent = text;
  }

  function setSyncMark(appKey, status) {
    if (!shellState.syncMarks) shellState.syncMarks = { app1: null, app2: null, app3: null };
    shellState.syncMarks[appKey] = (status === 'ok' || status === 'ng') ? status : null;
    renderTabSyncMarks();
  }

  function toggleSendPrefix(buttons, on) {
    (buttons || []).forEach(btn => {
      if (!btn) return;
      const raw = (btn.getAttribute('data-base-label') || btn.textContent || '').trim();
      if (!btn.getAttribute('data-base-label')) btn.setAttribute('data-base-label', raw.replace(/^📤\s*/u, ''));
      const base = (btn.getAttribute('data-base-label') || '').trim();
      btn.textContent = on ? `📤 ${base}` : base;
    });
  }

  function renderTabSyncMarks() {
    const map = {
      panelFundSearch: null,
      panelApp1: shellState.syncMarks && shellState.syncMarks.app1,
      panelApp2: shellState.syncMarks && shellState.syncMarks.app2,
      panelApp3: shellState.syncMarks && shellState.syncMarks.app3
    };
    document.querySelectorAll('.tab-btn').forEach(btn => {
      const tab = btn.getAttribute('data-tab-target');
      const st = map[tab];
      const dot = st === 'ok' ? '✅' : st === 'ng' ? '❌' : '';
      const cls = st === 'ok' ? 'sync-dot ok' : st === 'ng' ? 'sync-dot ng' : 'sync-dot';
      const label = btn.textContent.replace(/^[✅❌○×]\s*/u, '').trim();
      btn.innerHTML = dot ? `<span class="${cls}">${dot}</span><span>${label}</span>` : `<span>${label}</span>`;
    });
  }

  function installApp1TransferButton() {
    if (!els.frameApp1 || !els.frameApp1.contentDocument) return;
    const doc = els.frameApp1.contentDocument;
    if (!doc || doc.getElementById('codexApp1ToApp2FromPdfBtn')) return;
    const pdfBtn = doc.getElementById('btnExportPdf45');
    if (!pdfBtn || !pdfBtn.parentElement) return;
    const transferBtn = doc.createElement('button');
    transferBtn.type = 'button';
    transferBtn.id = 'codexApp1ToApp2FromPdfBtn';
    transferBtn.className = pdfBtn.className || 'ghost';
    transferBtn.textContent = '📤 App1⑥のデータをApp2「積立・スポット投資実績.XLSX」へ転送';
    transferBtn.style.marginTop = '8px';
    transferBtn.addEventListener('click', async () => {
      try {
        await syncApp1ToApp2();
        setSyncMark('app1', 'ok');
        setSyncMark('app2', 'ok');
      } catch (error) {
        setSyncMark('app1', 'ng');
        setSyncMark('app2', 'ng');
        alert(error && error.message ? error.message : 'App2への転送に失敗しました。');
      }
    });
    pdfBtn.parentElement.appendChild(transferBtn);
  }

  function ensureApp1DrawModeCheckboxes() {
    if (!els.frameApp1 || !els.frameApp1.contentDocument) return;
    const doc = els.frameApp1.contentDocument;
    if (!doc) return;
    const calcBtn = Array.from(doc.querySelectorAll('button, input[type="button"], input[type="submit"]'))
      .find(el => String((el.value || el.textContent || '')).includes('計算・描画'));
    if (!calcBtn || !calcBtn.parentElement) return;
    if (doc.getElementById('codexApp1DrawModeBox')) return;

    const box = doc.createElement('span');
    box.id = 'codexApp1DrawModeBox';
    box.style.display = 'inline-flex';
    box.style.alignItems = 'center';
    box.style.flexWrap = 'wrap';
    box.style.gap = '10px';
    box.style.marginLeft = '10px';
    box.style.verticalAlign = 'middle';
    box.innerHTML = [
      '<label style="display:inline-flex;align-items:center;gap:4px;cursor:pointer;"><input type="checkbox" id="codexDrawTsumitate" checked>積立</label>',
      '<label style="display:inline-flex;align-items:center;gap:4px;cursor:pointer;"><input type="checkbox" id="codexDrawLump">任意一括</label>',
      '<label style="display:inline-flex;align-items:center;gap:4px;cursor:pointer;"><input type="checkbox" id="codexDrawSameLump">積立同額一括</label>'
    ].join('');
    calcBtn.insertAdjacentElement('afterend', box);

    const rerender = () => {
      const app1 = safeGetAppWindow('app1');
      if (!app1) return;
      if (typeof app1.__codexEnforceDrawModeVisibility === 'function') {
        try { app1.__codexEnforceDrawModeVisibility(); } catch (e) { /* no-op */ }
      }
      if (typeof app1.renderAllFromVisibleRange === 'function') {
        try { app1.renderAllFromVisibleRange(); } catch (e) { /* no-op */ }
      } else if (typeof app1.renderResults === 'function') {
        try { app1.renderResults(); } catch (e) { /* no-op */ }
      }
    };
    Array.from(box.querySelectorAll('input[type="checkbox"]')).forEach(chk => {
      chk.addEventListener('change', rerender);
    });
  }

  function applyApp1ChartDrawModeFilter() {
    try {
      if (!els.frameApp1 || !els.frameApp1.contentWindow) return;
      const app1 = els.frameApp1.contentWindow;
      if (app1.__codexApp1ChartFilterPatched || typeof app1.Chart !== 'function') return;
      const OriginalChart = app1.Chart;
      const shouldKeepDataset = label => {
        const doc = els.frameApp1 && els.frameApp1.contentDocument;
        if (!doc) return true;
        const sumi = !!(doc.getElementById('codexDrawTsumitate') && doc.getElementById('codexDrawTsumitate').checked);
        const lump = !!(doc.getElementById('codexDrawLump') && doc.getElementById('codexDrawLump').checked);
        const same = !!(doc.getElementById('codexDrawSameLump') && doc.getElementById('codexDrawSameLump').checked);
        if (!sumi && !lump && !same) return true;
        const txt = String(label || '');
        if (/BM|ベンチ|benchmark/i.test(txt)) return true;
        const isSame = /同額一括|積立総額→同額一括|積立同額一括/.test(txt);
        const isLump = /任意一括/.test(txt);
        const isSumi = /積立/.test(txt) && !isSame;
        if (isSame) return same;
        if (isLump) return lump;
        if (isSumi) return sumi;
        return true;
      };
      const patchConfig = cfg => {
        if (!cfg || !cfg.data || !Array.isArray(cfg.data.datasets)) return cfg;
        const datasets = cfg.data.datasets;
        const hasTarget = datasets.some(ds => /積立|任意一括|同額一括|積立総額→同額一括/.test(String(ds && ds.label || '')));
        if (!hasTarget) return cfg;
        cfg.data.datasets = datasets.filter(ds => shouldKeepDataset(ds && ds.label));
        return cfg;
      };
      const WrappedChart = function wrappedChart(item, config) {
        return new OriginalChart(item, patchConfig(config));
      };
      Object.setPrototypeOf(WrappedChart, OriginalChart);
      WrappedChart.prototype = OriginalChart.prototype;
      Object.keys(OriginalChart).forEach(k => {
        try { WrappedChart[k] = OriginalChart[k]; } catch (e) { /* no-op */ }
      });
      app1.Chart = WrappedChart;
      app1.__codexApp1ChartFilterPatched = true;
    } catch (e) {
      // no-op
    }
  }

  function installParentFileGuard() {
    window.addEventListener('click', event => {
      const link = event.target && event.target.closest ? event.target.closest('a[href]') : null;
      if (!link) return;
      const href = link.getAttribute('href') || '';
      if (!(href === '#' || href.startsWith('#') || href.startsWith('file:'))) return;
      event.preventDefault();
      event.stopPropagation();
      const id = href.replace(/^#/, '');
      if (!id) return;
      const target = document.getElementById(id);
      if (target && target.scrollIntoView) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, true);

    document.addEventListener('submit', event => {
      const form = event.target;
      const action = form && form.getAttribute ? (form.getAttribute('action') || '') : '';
      if (!action.startsWith('file:')) return;
      event.preventDefault();
      event.stopPropagation();
    }, true);
  }

  function showBanner(text, type) {
    if (!els.portfolioBanner) return;
    if (!text) {
      els.portfolioBanner.className = 'banner info';
      els.portfolioBanner.textContent = '';
      return;
    }
    els.portfolioBanner.className = `banner ${type || 'info'} show`;
    els.portfolioBanner.textContent = text;
  }

  function switchTab(tabId) {
    shellState.activeTab = tabId;
    document.querySelectorAll('.panel').forEach(panel => panel.classList.toggle('active', panel.id === tabId));
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.toggle('active', btn.getAttribute('data-tab-target') === tabId));
    if (tabId === 'panelFundSearch') setTimeout(() => resizeIframe(els.frameFundSearch), 80);
    if (tabId === 'panelApp1') setTimeout(() => resizeIframe(els.frameApp1), 80);
    if (tabId === 'panelApp2') setTimeout(() => resizeIframe(els.frameApp2), 80);
    if (tabId === 'panelApp3' && shellState.frameReady.app2 && !shellState.portfolioItems.length) {
      setTimeout(() => {
        syncApp2ToPortfolio().catch(() => { /* keep manual flow if no data */ });
      }, 120);
    }
  }

  function safeGetAppWindow(kind) {
    const frame = kind === 'app1' ? els.frameApp1 : els.frameApp2;
    return frame && frame.contentWindow ? frame.contentWindow : null;
  }

  function safeGetAppDocument(kind) {
    const app = safeGetAppWindow(kind);
    try {
      return app && app.document ? app.document : null;
    } catch (error) {
      return null;
    }
  }

  function safeGetAppState(appWindow) {
    if (!appWindow) return null;
    if (typeof appWindow.__codexGetState === 'function') {
      try { return appWindow.__codexGetState(); } catch (error) { return null; }
    }
    if (appWindow.state) return appWindow.state;
    try {
      if (typeof appWindow.eval === 'function') {
        const v = appWindow.eval('(typeof state !== "undefined") ? state : null');
        if (v && typeof v === 'object') return v;
      }
    } catch (error) {
      // no-op
    }
    return null;
  }

  function num(value, fallback) {
    const normalized = String(value ?? '')
      .replace(/[,\s\u3000]/g, '')
      .replace(/[％%]/g, '')
      .replace(/[−―ー]/g, '-')
      .replace(/[^0-9.\-]/g, '');
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : (fallback == null ? NaN : fallback);
  }

  function normalizeDateKey(value) {
    if (value == null) return '';
    if (value instanceof Date) {
      if (Number.isNaN(value.getTime())) return '';
      const y = value.getFullYear();
      const m = String(value.getMonth() + 1).padStart(2, '0');
      const d = String(value.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    const raw = String(value).trim();
    if (!raw) return '';
    const m1 = raw.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
    if (m1) return `${m1[1]}-${String(m1[2]).padStart(2, '0')}-${String(m1[3]).padStart(2, '0')}`;
    const dt = new Date(raw);
    if (Number.isNaN(dt.getTime())) return '';
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }

  function slugify(text) {
    const ascii = String(text || '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '');
    if (ascii) return ascii.slice(0, 40);
    let hash = 0;
    for (const ch of String(text || 'item')) hash = ((hash << 5) - hash + ch.charCodeAt(0)) >>> 0;
    return `item-${hash.toString(16)}`;
  }

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function fmtMoney(value) {
    const n = Number(value);
    return Number.isFinite(n) ? n.toLocaleString('ja-JP', { maximumFractionDigits: 0 }) : '-';
  }

  function fmtPct(value, digits = 2) {
    const n = Number(value);
    return Number.isFinite(n)
      ? `${(n * 100).toLocaleString('ja-JP', { minimumFractionDigits: digits, maximumFractionDigits: digits })}%`
      : '-';
  }

  function fmtInput(value, digits = 2) {
    const n = Number(value);
    return Number.isFinite(n) ? n.toFixed(digits) : '';
  }

  function formatMoneyInputValue(value) {
    const n = Math.max(0, Math.round(num(value, 0)));
    return n.toLocaleString('ja-JP');
  }

  function inferColor(index, name) {
    const fixed = ['#36a2eb', '#ff6384', '#ffcd56', '#4bc0c0', '#9966ff', '#ff9f40', '#8dd17e', '#b07aa1'];
    const seed = Array.from(String(name || '')).reduce((sum, ch) => sum + ch.charCodeAt(0), 0);
    return fixed[(seed + index) % fixed.length];
  }

  function getApp2SeriesColors() {
    const app2 = safeGetAppWindow('app2');
    if (!app2) return {};
    const out = {};
    if (typeof app2.__codexGetSeriesColors === 'function') {
      try { Object.assign(out, app2.__codexGetSeriesColors() || {}); } catch (error) {}
    } else {
      const state = safeGetAppState(app2);
      if (state && state.seriesColors) Object.assign(out, state.seriesColors);
    }
    const names = getWorkbookFundNamesForColorMapping();
    if (typeof app2.getSeriesColor === 'function') {
      names.forEach((name, idx) => {
        if (!name || out[name]) return;
        try {
          const c = app2.getSeriesColor(name, idx);
          if (c) out[name] = c;
        } catch (error) {}
      });
    } else if (typeof app2.getDefaultSeriesColor === 'function') {
      names.forEach((name, idx) => {
        if (!name || out[name]) return;
        try {
          const c = app2.getDefaultSeriesColor(name, idx);
          if (c) out[name] = c;
        } catch (error) {}
      });
    }
    return out;
  }

  function getApp2FundColorMap() {
    const app2 = safeGetAppWindow('app2');
    let map = {};
    try {
      const state = safeGetAppState(app2);
      if (state && state.workbookAnalysis && state.workbookAnalysis.fundColorMap) {
        map = Object.assign(map, state.workbookAnalysis.fundColorMap);
      }
    } catch (error) {}
    try {
      const ls = (app2 && app2.localStorage) ? app2.localStorage : window.localStorage;
      const raw = ls.getItem('app2FundColorMapForApp3');
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') map = Object.assign(map, parsed);
      }
    } catch (error) {}
    return map;
  }

  function getWorkbookFundNamesForColorMapping() {
    const app2 = safeGetAppWindow('app2');
    const state = safeGetAppState(app2) || {};
    const bmName = String(state.benchmarkName || state.bmName || '').trim();
    const names = [];
    const push = (v) => {
      const s = String(v || '').trim();
      if (!s) return;
      if (bmName && slugify(s) === slugify(bmName)) return;
      if (names.some(n => slugify(n) === slugify(s))) return;
      names.push(s);
    };
    let rows = [];
    try { rows = getApp2WorkbookRows(); } catch (error) { rows = []; }
    rows.forEach(row => push(row && (row.fund || row.name)));
    const analysisRows = (state && state.workbookAnalysis && Array.isArray(state.workbookAnalysis.rows))
      ? state.workbookAnalysis.rows
      : [];
    analysisRows.forEach(row => push(row && (row.fund || row.name)));
    const seriesAll = (state && Array.isArray(state.fundSeriesAll)) ? state.fundSeriesAll : [];
    seriesAll.forEach(series => push(series && series.name));
    return Array.from(new Set(
      names
    ));
  }

  function buildResolvedApp2ColorMap(colorMap) {
    const resolved = Object.assign({}, colorMap || {});
    const names = getWorkbookFundNamesForColorMapping();
    Object.keys(resolved).forEach(key => {
      const m = String(key).match(/^(?:投資家|investor)\s*(\d+)$/i);
      if (!m) return;
      const ix = Math.max(0, Number(m[1]) - 1);
      const mappedName = names[ix];
      if (mappedName && !resolved[mappedName]) resolved[mappedName] = resolved[key];
    });
    return resolved;
  }

  function pickColorByName(colorMap, name) {
    if (!colorMap || !name) return '';
    if (colorMap[name]) return colorMap[name];
    const key = slugify(name);
    const pair = Object.entries(colorMap).find(([k]) => slugify(k) === key);
    return pair ? pair[1] : '';
  }

  function normalizeItem(raw, index) {
    const name = raw.name || raw.fund || `Fund ${index + 1}`;
    const currentValue = Math.max(0, num(raw.currentValue ?? raw.valuation, 0));
    let units = Math.max(0, num(raw.units, 0));
    const nav = Math.max(1, num(raw.nav, units > 0 && currentValue > 0 ? currentValue / units : 10000));
    if (units <= 0 && currentValue > 0 && nav > 0) {
      units = currentValue / nav;
    }
    const totalInvestment = Math.max(0, num(raw.totalInvestment, 0));
    const expectedReturn = num(raw.expectedReturn, num(raw.returnPct, 0));
    const expectedRisk = Math.max(0, num(raw.expectedRisk, num(raw.annualVol, 0)));
    const maxDrawdown = num(raw.maxDrawdown, NaN);
    const app2Colors = buildResolvedApp2ColorMap(getApp2SeriesColors());
    const color = raw.color || pickColorByName(app2Colors, name) || inferColor(index, name);
    return {
      id: raw.id || `${slugify(name)}-${index}`,
      name,
      source: raw.source || 'imported',
      currentValue,
      currentWeight: 0,
      targetWeight: num(raw.targetWeight, NaN),
      targetValue: 0,
      diffValue: 0,
      units,
      unitDiff: 0,
      nav,
      expectedReturn: Number.isFinite(expectedReturn) ? expectedReturn : 0,
      expectedRisk,
      maxDrawdown,
      totalInvestment,
      recurringInvestment: Math.max(0, num(raw.recurringInvestment, 0)),
      spotInvestment: Math.max(0, num(raw.spotInvestment, 0)),
      sharpe: num(raw.sharpe, NaN),
      beta: num(raw.beta, NaN),
      trackingError: num(raw.trackingError, NaN),
      color,
      notes: raw.notes || ''
    };
  }

  function normalizePortfolio() {
    const items = shellState.portfolioItems.map(normalizeItem);
    const totalCurrent = items.reduce((sum, item) => sum + item.currentValue, 0);
    items.forEach(item => {
      item.currentWeight = totalCurrent > 0 ? item.currentValue / totalCurrent : 0;
    });
    const targetTotal = items.reduce((sum, item) => sum + (Number.isFinite(item.targetWeight) ? Math.max(0, item.targetWeight) : 0), 0);
    if (items.length) {
      if (targetTotal > 0) {
        items.forEach(item => {
          item.targetWeight = (Number.isFinite(item.targetWeight) ? Math.max(0, item.targetWeight) : 0) / targetTotal;
        });
      } else if (totalCurrent > 0) {
        items.forEach(item => { item.targetWeight = item.currentWeight; });
      } else {
        const equal = 1 / items.length;
        items.forEach(item => { item.targetWeight = equal; });
      }
    }
    shellState.portfolioItems = items;
  }

  function computePortfolioMetrics(items, weights, correlation) {
    if (!items.length) return { expectedReturn: 0, expectedRisk: 0, maxDrawdown: 0 };
    let expectedReturn = 0;
    let variance = 0;
    let maxDrawdown = 0;
    for (let i = 0; i < items.length; i += 1) {
      const wi = weights[i] || 0;
      expectedReturn += wi * (items[i].expectedReturn || 0);
      variance += wi * wi * Math.pow(items[i].expectedRisk || 0, 2);
      if (Number.isFinite(items[i].maxDrawdown)) maxDrawdown += wi * items[i].maxDrawdown;
    }
    for (let i = 0; i < items.length; i += 1) {
      for (let j = i + 1; j < items.length; j += 1) {
        variance += 2 * (weights[i] || 0) * (weights[j] || 0) * (items[i].expectedRisk || 0) * (items[j].expectedRisk || 0) * correlation;
      }
    }
    return {
      expectedReturn,
      expectedRisk: Math.sqrt(Math.max(variance, 0)),
      maxDrawdown
    };
  }

  function computeMinimumExtraCashForTargetMix(items, currentTotal) {
    if (!Array.isArray(items) || !items.length) return 0;
    let requiredCapital = Math.max(0, Number(currentTotal) || 0);
    items.forEach(item => {
      const w = Math.max(0, Number(item && item.targetWeight) || 0);
      const currentValue = Math.max(0, Number(item && item.currentValue) || 0);
      if (currentValue > 0 && w > 0) {
        requiredCapital = Math.max(requiredCapital, currentValue / w);
      }
    });
    return Math.max(0, requiredCapital - Math.max(0, Number(currentTotal) || 0));
  }

  function getSelectedTargetReturnForExtraCash() {
    // Extra Cash定義: GoalをYears内で達成するために任意選択したパフォーマンス(=Target Return)を使用
    if (shellState.goalPlannerMode === 'manual') {
      const mid = num(els.goalTargetReturnMidPct && els.goalTargetReturnMidPct.value, NaN);
      const high = num(els.goalTargetReturnHighPct && els.goalTargetReturnHighPct.value, NaN);
      const choice = shellState.manualTargetReturnChoice || 'mid';
      if (choice === 'high' && Number.isFinite(high)) return high / 100;
      if (Number.isFinite(mid)) return mid / 100;
      if (Number.isFinite(high)) return high / 100;
      return 0.03;
    }
    const app2Targets = getApp2PlannerTargets();
    if (Number.isFinite(app2Targets.expectedRet)) return app2Targets.expectedRet;
    return 0.03;
  }

  function buildPortfolioPlan() {
    normalizePortfolio();
    let extraCash = Math.max(0, num(els.portfolioExtraCash && els.portfolioExtraCash.value, 0));
    const correlation = clamp(num(els.portfolioCorrelation && els.portfolioCorrelation.value, 0.25), 0, 1);
    const riskFree = num(els.portfolioRiskFree && els.portfolioRiskFree.value, 0) / 100;
    const currentTotal = shellState.portfolioItems.reduce((sum, item) => sum + item.currentValue, 0);
    const minimumExtraCash = computeMinimumExtraCashForTargetMix(shellState.portfolioItems, currentTotal);
    const goalCapital = Math.max(0, num(els.goalTargetAmount && els.goalTargetAmount.value, NaN));
    const yearsToGoal = Math.max(1, Math.round(num(els.goalYears && els.goalYears.value, 10)));
    const selectedReturn = getSelectedTargetReturnForExtraCash();
    const pvGoal = (goalCapital > 0 && selectedReturn > -1)
      ? (goalCapital / Math.pow(1 + selectedReturn, yearsToGoal))
      : goalCapital;
    const requiredPresentValueExtra = Math.max(0, pvGoal - currentTotal);
    if (shellState.investmentMode === 'lump') {
      extraCash = requiredPresentValueExtra;
    } else if (!shellState.savingsExtraCashEdited) {
      extraCash = requiredPresentValueExtra + (requiredPresentValueExtra * Math.SQRT2);
    }
    let targetCapital = Number.isFinite(goalCapital) && goalCapital > 0
      ? Math.max(goalCapital, currentTotal)
      : (currentTotal + extraCash);
    if (shellState.buyOnlyMode && shellState.portfolioItems.length) {
      let requiredCapital = currentTotal;
      shellState.portfolioItems.forEach(item => {
        const w = Math.max(0, item.targetWeight || 0);
        if (item.currentValue > 0 && w > 0) {
          requiredCapital = Math.max(requiredCapital, item.currentValue / w);
        }
      });
      // Buy-Only時は効率的フロンティア最小額条件も満たす
      const frontierExtra = Math.max(0, requiredCapital - currentTotal);
      extraCash = Math.max(extraCash, frontierExtra);
      targetCapital = currentTotal + extraCash;
    }
    if (els.portfolioExtraCash) els.portfolioExtraCash.value = formatMoneyInputValue(extraCash);
    shellState.portfolioItems.filter(item => !shellState.hiddenFunds.has(item.id)).forEach(item => {
      const rawTargetValue = targetCapital * item.targetWeight;
      item.targetValue = shellState.buyOnlyMode ? Math.max(item.currentValue, rawTargetValue) : rawTargetValue;
      item.diffValue = item.targetValue - item.currentValue;
      item.unitDiff = item.nav > 0 ? item.diffValue / item.nav : 0;
      item.targetUnits = item.nav > 0 ? (item.targetValue / item.nav) : 0;
    });
    if (shellState.buyOnlyMode) {
      const effectiveTarget = shellState.portfolioItems.reduce((sum, item) => sum + item.targetValue, 0);
      extraCash = Math.max(0, effectiveTarget - currentTotal);
      targetCapital = effectiveTarget;
      if (els.portfolioExtraCash) els.portfolioExtraCash.value = formatMoneyInputValue(extraCash);
    }
    const currentMetrics = computePortfolioMetrics(shellState.portfolioItems, shellState.portfolioItems.map(item => item.currentWeight), correlation);
    const targetMetrics = computePortfolioMetrics(shellState.portfolioItems, shellState.portfolioItems.map(item => item.targetWeight), correlation);
    return { currentTotal, targetCapital, extraCash, minimumExtraCash, correlation, riskFree, currentMetrics, targetMetrics, selectedReturn };
  }

  function renderSummary(plan) {
    const visibleCount = shellState.portfolioItems.filter(item => !shellState.hiddenFunds.has(item.id)).length;
    const cards = [
      { label: 'Current value', value: fmtMoney(plan.currentTotal), sub: 'JPY' },
      { label: 'Target value', value: fmtMoney(plan.targetCapital), sub: `期待値基準 Extra cash ${fmtMoney(plan.extraCash)} / 最小 ${fmtMoney(plan.minimumExtraCash || 0)}` },
      { label: 'Target return', value: fmtPct(plan.targetMetrics.expectedReturn), sub: `Current ${fmtPct(plan.currentMetrics.expectedReturn)}` },
      { label: 'Target risk', value: fmtPct(plan.targetMetrics.expectedRisk), sub: `Current ${fmtPct(plan.currentMetrics.expectedRisk)}` },
      { label: 'Risk free', value: fmtPct(plan.riskFree), sub: `Correlation ${plan.correlation.toFixed(2)}` },
      { label: 'Funds', value: String(visibleCount), sub: shellState.portfolioSource || 'No source' }
    ];
    els.portfolioSummaryGrid.innerHTML = cards.map(card => `
      <div class="metric-card">
        <div class="metric-label">${escapeHtml(card.label)}</div>
        <div class="metric-value">${escapeHtml(card.value)}</div>
        <div class="metric-sub">${escapeHtml(card.sub)}</div>
      </div>
    `).join('');
    if (els.portfolioSourceText) els.portfolioSourceText.textContent = shellState.portfolioSource || 'No source';
  }

  function renderGoalProjection(plan) {
    if (!els.goalProjectionSummary || !els.goalProjectionChart) return;
    const current = Math.max(0, plan.currentTotal || 0);
    const years = Math.max(1, Math.round(num(els.goalYears && els.goalYears.value, 10)));
    let goal = Math.max(0, num(els.goalTargetAmount && els.goalTargetAmount.value, current));
    let reqReturn = (current > 0 && goal > 0) ? (Math.pow(goal / current, 1 / years) - 1) : 0;
    const mean = Number.isFinite(plan.targetMetrics.expectedReturn) ? plan.targetMetrics.expectedReturn : 0;
    let lowerRet = -0.03;
    let midRet = 0.03;
    let upperRet = 0.06;
    let maxDD = Math.min(...shellState.portfolioItems.map(item => item.maxDrawdown).filter(Number.isFinite), plan.currentMetrics.maxDrawdown || 0);
    let maxRet = Math.max(...shellState.portfolioItems.map(item => item.expectedReturn || 0), mean);
    if (shellState.goalPlannerMode === 'manual') {
      const manualLowInput = num(els.goalTargetReturnLowPct && els.goalTargetReturnLowPct.value, -3) / 100;
      midRet = num(els.goalTargetReturnMidPct && els.goalTargetReturnMidPct.value, 3) / 100;
      upperRet = num(els.goalTargetReturnHighPct && els.goalTargetReturnHighPct.value, 6) / 100;
      // Manual Inputは最低値入力をそのまま採用
      lowerRet = manualLowInput;
      maxDD = lowerRet;
      maxRet = upperRet;
    } else {
      const app2Targets = getApp2PlannerTargets();
      if (Number.isFinite(app2Targets.expectedRet)) midRet = app2Targets.expectedRet;
      if (Number.isFinite(app2Targets.maxRet)) upperRet = app2Targets.maxRet;
      // 最低値は「期待値 - |平均最大DD|」
      if (Number.isFinite(app2Targets.minRet) && Number.isFinite(midRet)) {
        lowerRet = midRet - Math.abs(app2Targets.minRet);
      } else if (Number.isFinite(app2Targets.minRet)) {
        lowerRet = app2Targets.minRet;
      }
      maxDD = lowerRet;
      maxRet = Number.isFinite(app2Targets.maxRet) ? app2Targets.maxRet : maxRet;
    }
    const months = years * 12;
    const expectedSeries = [];
    const requiredSeries = [];
    const upperSeries = [];
    const lowerSeries = [];
    for (let m = 0; m <= months; m += 1) {
      const t = m / 12;
      const expected = current * Math.pow(1 + midRet, t);
      const required = current * Math.pow(1 + reqReturn, t);
      expectedSeries.push(expected);
      requiredSeries.push(required);
      upperSeries.push(current * Math.pow(1 + upperRet, t));
      lowerSeries.push(current * Math.pow(1 + lowerRet, t));
    }
    const allValues = expectedSeries.concat(requiredSeries, upperSeries, lowerSeries);
    const maxY = Math.max(1, ...allValues) * 1.1;
    const w = 920;
    const h = 260;
    const pad = { l: 62, r: 16, t: 30, b: 34 };
    const pw = w - pad.l - pad.r;
    const ph = h - pad.t - pad.b;
    const xPos = i => pad.l + (pw * (i / Math.max(1, months)));
    const yPos = v => pad.t + ph - (ph * (v / maxY));
    const pathFrom = arr => arr.map((v, i) => `${i === 0 ? 'M' : 'L'} ${xPos(i)} ${yPos(v)}`).join(' ');
    const grid = Array.from({ length: 6 }, (_, i) => {
      const y = pad.t + (ph * (i / 5));
      return `<line x1="${pad.l}" y1="${y}" x2="${pad.l + pw}" y2="${y}" stroke="#e2e8f0" stroke-width="1"/>`;
    }).join('');
    if (window.Plotly && typeof window.Plotly.newPlot === 'function') {
      const x = Array.from({ length: months + 1 }, (_, i) => i / 12);
      window.Plotly.newPlot(els.goalProjectionChart, [
        { x, y: lowerSeries, mode: 'lines', name: `最低値 (${fmtPct(maxDD)})`, line: { color: '#ef4444', width: 2 } },
        { x, y: upperSeries, mode: 'lines', name: `最大値 (${fmtPct(upperRet)})`, line: { color: '#22c55e', width: 2 } },
        { x, y: expectedSeries, mode: 'lines', name: `期待値 (${fmtPct(midRet)})`, line: { color: '#0f172a', width: 2.4 } },
        { x, y: requiredSeries, mode: 'lines', name: `必要到達線 (${fmtPct(reqReturn)})`, line: { color: '#3b82f6', width: 2.2, dash: 'dash' } }
      ], {
        margin: { l: 60, r: 16, t: 30, b: 38 },
        xaxis: {
          title: 'Years',
          automargin: true,
          range: [0, years],
          tickmode: 'linear',
          dtick: 1,
          tickformat: '.0f',
          hoverformat: '.0f'
        },
        yaxis: { title: 'JPY', automargin: true, tickformat: ',.0f' },
        paper_bgcolor: '#ffffff',
        plot_bgcolor: '#ffffff',
        legend: { orientation: 'h' }
      }, { responsive: true, displaylogo: false, modeBarButtonsToRemove: ['select2d', 'lasso2d'] });
      if (shellState.goalPlannerMode === 'chart') {
        els.goalProjectionChart.on('plotly_click', ev => {
          const p = ev && ev.points && ev.points[0];
          if (!p) return;
          const y = Math.max(0, Math.round(Number(p.y) || 0));
          const xYear = Math.max(1, Math.round(Number(p.x) || years));
          if (els.goalTargetAmount) els.goalTargetAmount.value = formatMoneyInputValue(y);
          if (els.goalYears) els.goalYears.value = String(xYear);
          if (els.goalYearsRange) els.goalYearsRange.value = String(xYear);
          renderPortfolio();
        });
      } else if (els.goalProjectionChart.removeAllListeners) {
        els.goalProjectionChart.removeAllListeners('plotly_click');
      }
    } else {
      els.goalProjectionChart.innerHTML = `<svg viewBox="0 0 ${w} ${h}" width="100%" height="100%" aria-label="target value projection">
      ${grid}
      <line x1="${pad.l}" y1="${pad.t + ph}" x2="${pad.l + pw}" y2="${pad.t + ph}" stroke="#64748b" stroke-width="1.2"/>
      <line x1="${pad.l}" y1="${pad.t}" x2="${pad.l}" y2="${pad.t + ph}" stroke="#64748b" stroke-width="1.2"/>
      <path d="${pathFrom(lowerSeries)}" fill="none" stroke="#ef4444" stroke-width="1.8"/>
      <path d="${pathFrom(upperSeries)}" fill="none" stroke="#22c55e" stroke-width="1.8"/>
      <path d="${pathFrom(expectedSeries)}" fill="none" stroke="#0f172a" stroke-width="2.2"/>
      <path d="${pathFrom(requiredSeries)}" fill="none" stroke="#3b82f6" stroke-width="2.2" stroke-dasharray="6 4"/>
    </svg>`;
    }
    els.goalProjectionSummary.innerHTML = `
      投資原資 <b>${fmtMoney(current)} JPY</b> |
      目標金額 ${fmtMoney(goal)} JPY / ${years}年 |
      必要年換算リターン <b>${fmtPct(reqReturn)}</b> |
      期待値リターン <b>${fmtPct(midRet)}</b> |
      最大値リターン <b>${fmtPct(upperRet)}</b> |
      最大DD <b>${fmtPct(maxDD)}</b> |
      リスク許容度 <b>${Math.abs(shellState.drawdownTolerance).toFixed(1)}%</b>
    `;
  }

  function renderDonut(plan) {
    const svg = els.portfolioDonut;
    if (!svg) return;
    const size = 360;
    const radius = 108;
    const circumference = 2 * Math.PI * radius;
    let offset = 0;
    const arcs = [];
    shellState.portfolioItems.forEach(item => {
      const share = item.targetWeight || 0;
      if (share <= 0) return;
      const dash = `${(share * circumference).toFixed(2)} ${(circumference - share * circumference).toFixed(2)}`;
      arcs.push(`<circle cx="180" cy="180" r="${radius}" fill="none" stroke="${item.color}" stroke-width="64" stroke-dasharray="${dash}" stroke-dashoffset="${(-offset).toFixed(2)}" transform="rotate(-90 180 180)"></circle>`);
      offset += share * circumference;
    });
    svg.innerHTML = `<circle cx="180" cy="180" r="${radius}" fill="none" stroke="#e5e7eb" stroke-width="64"></circle>${arcs.join('')}`;
    els.donutCenter.innerHTML = `
      <div class="donut-main">${fmtPct(plan.targetMetrics.expectedReturn)}</div>
      <div class="donut-sub">Target return</div>
      <div class="donut-mini">Risk ${fmtPct(plan.targetMetrics.expectedRisk)}</div>
    `;
    els.portfolioLegend.innerHTML = shellState.portfolioItems.filter(item => !shellState.hiddenFunds.has(item.id)).map(item => `
      <div class="legend-item">
        <span class="legend-swatch" style="background:${escapeHtml(item.color)}"></span>
        <span>${escapeHtml(item.name)}</span>
        <span style="display:flex;align-items:center;gap:8px;">
          <span class="legend-value">${fmtPct(item.targetWeight)}</span>
          <button type="button" class="small-btn" data-remove-item="${escapeHtml(item.id)}" style="padding:2px 8px;line-height:1.2;">×</button>
        </span>
      </div>
    `).join('');
  }

  function renderSingleDonut(svg, centerEl, weights, mainText, subLine, miniLine) {
    if (!svg || !centerEl) return;
    const radius = 108;
    const circumference = 2 * Math.PI * radius;
    let offset = 0;
    const arcs = [];
    shellState.portfolioItems.forEach((item, index) => {
      if (shellState.hiddenFunds.has(item.id)) return;
      const share = Math.max(0, Number(weights[index] || 0));
      if (share <= 0) return;
      const dash = `${(share * circumference).toFixed(2)} ${(circumference - share * circumference).toFixed(2)}`;
      arcs.push(`<circle cx="180" cy="180" r="${radius}" fill="none" stroke="${item.color}" stroke-width="64" stroke-dasharray="${dash}" stroke-dashoffset="${(-offset).toFixed(2)}" transform="rotate(-90 180 180)"></circle>`);
      offset += share * circumference;
    });
    svg.innerHTML = `<circle cx="180" cy="180" r="${radius}" fill="none" stroke="#e5e7eb" stroke-width="64"></circle>${arcs.join('')}`;
    centerEl.innerHTML = `
      <div class="donut-main">${escapeHtml(mainText)}</div>
      <div class="donut-sub">${escapeHtml(subLine)}</div>
      <div class="donut-mini">${escapeHtml(miniLine)}</div>
    `;
  }

  function renderEditableMix(plan) {
    if (!els.currentPortfolioDonut || !els.currentPortfolioCenter || !els.editableMixDonut || !els.editableMixCenter) return;
    const currentWeights = shellState.portfolioItems.map(item => item.currentWeight || 0);
    const targetWeights = shellState.portfolioItems.map(item => item.targetWeight || 0);
    renderSingleDonut(
      els.currentPortfolioDonut,
      els.currentPortfolioCenter,
      currentWeights,
      fmtPct(plan.currentMetrics.expectedReturn),
      'Current return',
      `Risk ${fmtPct(plan.currentMetrics.expectedRisk)}`
    );
    renderSingleDonut(
      els.editableMixDonut,
      els.editableMixCenter,
      targetWeights,
      fmtPct(plan.targetMetrics.expectedReturn),
      'Expected return',
      `Risk ${fmtPct(plan.targetMetrics.expectedRisk)}`
    );
    const handles = [];
    let cum = 0;
    targetWeights.forEach((w, idx) => {
      if (shellState.hiddenFunds.has(shellState.portfolioItems[idx].id)) return;
      const share = Math.max(0, Number(w) || 0);
      if (share <= 0) return;
      const mid = cum + (share / 2);
      const angle = (-Math.PI / 2) + (2 * Math.PI * mid);
      const cx = 180 + Math.cos(angle) * 138;
      const cy = 180 + Math.sin(angle) * 138;
      handles.push(`<circle cx="${cx.toFixed(2)}" cy="${cy.toFixed(2)}" r="8" fill="${escapeHtml(shellState.portfolioItems[idx].color)}" stroke="#fff" stroke-width="2" data-mix-handle="${escapeHtml(shellState.portfolioItems[idx].id)}" style="cursor:grab;"></circle>`);
      cum += share;
    });
    els.editableMixDonut.innerHTML += handles.join('');
  }

  function renderUnitBarControl(plan) {
    if (!els.unitBarRows) return;
    const visibleItems = shellState.portfolioItems.filter(item => !shellState.hiddenFunds.has(item.id));
    const currentUnitsTotal = visibleItems.reduce((sum, item) => sum + Math.max(0, item.units || 0), 0);
    const targetUnitsTotal = visibleItems.reduce((sum, item) => sum + Math.max(0, item.nav > 0 ? item.targetValue / item.nav : 0), 0);
    const frontierExtra = Math.max(0, plan && Number.isFinite(plan.minimumExtraCash) ? plan.minimumExtraCash : 0);
    const frontierUnitsNote = `Units ${fmtInput(currentUnitsTotal, 0)} -> ${fmtInput(targetUnitsTotal, 0)}`;
    const frontierCostNote = `必要追加費用 ${fmtMoney(frontierExtra)}`;
    const maxUnits = Math.max(1, ...shellState.portfolioItems.map(item => {
      const targetUnits = item.nav > 0 ? item.targetValue / item.nav : 0;
      return Math.max(item.units || 0, targetUnits || 0);
    })) * 1.15;
    els.unitBarRows.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px;">
        <label style="display:inline-flex;align-items:center;gap:8px;margin:0;color:#334155;">
          <input type="checkbox" id="buyOnlyToggleInline" ${shellState.buyOnlyMode ? 'checked' : ''}>
          Buy-Only Frontier Simulation
        </label>
        <div class="inline-note">${shellState.buyOnlyMode ? `${frontierUnitsNote} / ${frontierCostNote} を Extra cashへ反映中` : `${frontierUnitsNote} / ${frontierCostNote}（Buy-Only ONでExtra cashへ反映）`}</div>
      </div>
    ` + shellState.portfolioItems.filter(item => !shellState.hiddenFunds.has(item.id)).map(item => {
      const currentUnits = Math.max(0, item.units || 0);
      const targetUnits = Math.max(0, item.nav > 0 ? item.targetValue / item.nav : 0);
      const currentPct = clamp((currentUnits / maxUnits) * 100, 0, 100);
      const targetPct = clamp((targetUnits / maxUnits) * 100, 0, 100);
      return `
        <div style="display:grid;grid-template-columns:240px 1fr 220px;gap:10px;align-items:center;margin-bottom:10px;">
          <div style="font-size:12px;color:#334155;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;display:flex;align-items:center;gap:8px;">
            <span>${escapeHtml(item.name)}</span>
            <button type="button" class="small-btn" data-hide-item="${escapeHtml(item.id)}" style="padding:1px 7px;line-height:1;">×</button>
          </div>
          <div data-unit-bar="${escapeHtml(item.id)}" style="position:relative;height:28px;border:1px solid #d8e5f2;border-radius:8px;background:#f8fafc;overflow:hidden;cursor:ew-resize;">
            <div style="position:absolute;left:0;top:4px;bottom:4px;width:${currentPct}%;background:#bfd4ea;border-radius:6px;"></div>
            <div style="position:absolute;left:0;top:9px;bottom:9px;width:${targetPct}%;background:${escapeHtml(item.color)};opacity:.72;border-radius:6px;"></div>
          </div>
          <div style="display:flex;align-items:center;justify-content:flex-end;">
            <span style="font-size:12px;color:#334155;">Units ${fmtInput(currentUnits, 0)} -> ${fmtInput(targetUnits, 0)}</span>
          </div>
        </div>
      `;
    }).join('');
  }

  function renderTradePlan() {
    els.tradePlanBody.innerHTML = shellState.portfolioItems.map(item => {
      const direction = item.diffValue > 1 ? 'Buy' : (item.diffValue < -1 ? 'Sell' : 'Hold');
      return `
        <tr>
          <td>${escapeHtml(item.name)}</td>
          <td class="num">${fmtMoney(item.currentValue)}</td>
          <td class="num">${fmtMoney(item.targetValue)}</td>
          <td class="num">${fmtMoney(item.diffValue)}</td>
          <td class="num">${fmtInput(item.unitDiff, 2)}</td>
          <td>${direction}</td>
        </tr>
      `;
    }).join('');
  }

  function computeEfficientFrontier(items, correlation) {
    if (!Array.isArray(items) || items.length < 2) return [];
    const sampleCount = 2200;
    const samples = [];
    for (let i = 0; i < sampleCount; i += 1) {
      const seed = items.map(() => Math.max(0.0001, Math.random()));
      const total = seed.reduce((a, b) => a + b, 0) || 1;
      const weights = seed.map(v => v / total);
      const m = computePortfolioMetrics(items, weights, correlation);
      samples.push({ risk: (m.expectedRisk || 0) * 100, ret: (m.expectedReturn || 0) * 100, weights });
    }
    const single = items.map((item, idx) => {
      const w = items.map((_, i) => (i === idx ? 1 : 0));
      const m = computePortfolioMetrics(items, w, correlation);
      return { risk: (m.expectedRisk || 0) * 100, ret: (m.expectedReturn || 0) * 100, weights: w };
    });
    const all = samples.concat(single).filter(p => Number.isFinite(p.risk) && Number.isFinite(p.ret));
    all.sort((a, b) => a.risk - b.risk || b.ret - a.ret);
    const frontier = [];
    let bestRet = -Infinity;
    all.forEach(p => {
      if (p.ret > bestRet + 0.01) {
        frontier.push(p);
        bestRet = p.ret;
      }
    });
    return frontier;
  }

  function applyTargetWeights(weights) {
    if (!Array.isArray(weights) || !weights.length) return;
    const total = weights.reduce((sum, v) => sum + Math.max(0, Number(v) || 0), 0) || 1;
    shellState.portfolioItems.forEach((item, idx) => {
      item.targetWeight = Math.max(0, Number(weights[idx]) || 0) / total;
    });
  }

  function renderRiskReturnScatter(plan) {
    if (!els.riskReturnScatter || !els.riskReturnLegend) return;
    const w = 760;
    const h = 330;
    const pad = { l: 62, r: 20, t: 18, b: 42 };
    const plotW = w - pad.l - pad.r;
    const plotH = h - pad.t - pad.b;
    const funds = shellState.portfolioItems.filter(item => !shellState.hiddenFunds.has(item.id)).map(item => ({
      name: item.name,
      color: item.color,
      x: (item.expectedRisk || 0) * 100,
      y: (item.expectedReturn || 0) * 100
    }));
    const extras = [
      { name: 'Current Portfolio', color: '#fa8072', x: (plan.currentMetrics.expectedRisk || 0) * 100, y: (plan.currentMetrics.expectedReturn || 0) * 100 },
      { name: 'Editable Target Mix', color: '#00ffff', x: (plan.targetMetrics.expectedRisk || 0) * 100, y: (plan.targetMetrics.expectedReturn || 0) * 100 }
    ];
    const frontier = computeEfficientFrontier(shellState.portfolioItems, plan.correlation);
    shellState.frontier = frontier;
    const allPoints = funds.concat(extras, frontier.map(p => ({ x: p.risk, y: p.ret })));
    const maxX = Math.max(5, ...allPoints.map(p => p.x || 0)) * 1.15;
    const minY = Math.min(-2, ...allPoints.map(p => p.y || 0)) * 1.1;
    const maxY = Math.max(2, ...allPoints.map(p => p.y || 0)) * 1.15;
    const xPos = value => pad.l + (plotW * (value / maxX));
    const yPos = value => pad.t + plotH - (plotH * ((value - minY) / (maxY - minY || 1)));
    const gridX = Array.from({ length: 6 }, (_, i) => (maxX / 5) * i);
    const gridY = Array.from({ length: 6 }, (_, i) => minY + ((maxY - minY) / 5) * i);
    const grid = []
      .concat(gridX.map(v => `<line x1="${xPos(v)}" y1="${pad.t}" x2="${xPos(v)}" y2="${pad.t + plotH}" stroke="#d8e5f2" stroke-width="1"/>`))
      .concat(gridY.map(v => `<line x1="${pad.l}" y1="${yPos(v)}" x2="${pad.l + plotW}" y2="${yPos(v)}" stroke="#d8e5f2" stroke-width="1"/>`))
      .join('');
    const axes = `
      <line x1="${pad.l}" y1="${pad.t + plotH}" x2="${pad.l + plotW}" y2="${pad.t + plotH}" stroke="#4f6b87" stroke-width="1.2"/>
      <line x1="${pad.l}" y1="${pad.t}" x2="${pad.l}" y2="${pad.t + plotH}" stroke="#4f6b87" stroke-width="1.2"/>
      <text x="${pad.l + plotW / 2}" y="${h - 8}" text-anchor="middle" fill="#516b83" font-size="12">Risk (%)</text>
      <text x="15" y="${pad.t + plotH / 2}" text-anchor="middle" fill="#516b83" font-size="12" transform="rotate(-90 15 ${pad.t + plotH / 2})">Return (%)</text>
    `;
    const riskTol = Math.abs(shellState.drawdownTolerance || 0);
    const diagonal = `<line x1="${xPos(0)}" y1="${yPos(0)}" x2="${xPos(maxX)}" y2="${yPos(maxX)}" stroke="#94a3b8" stroke-width="1.2" stroke-dasharray="4 4"></line>`;
    const riskLine = `<line x1="${xPos(riskTol)}" y1="${yPos(minY)}" x2="${xPos(riskTol)}" y2="${yPos(maxY)}" stroke="#ef4444" stroke-width="1.2" stroke-dasharray="4 4"></line>`;
    const frontierPath = frontier.map((p, idx) => `${idx === 0 ? 'M' : 'L'} ${xPos(p.risk)} ${yPos(p.ret)}`).join(' ');
    const frontierLine = frontierPath
      ? `<path d="${frontierPath}" fill="none" stroke="#f59e0b" stroke-width="2.2" stroke-dasharray="6 4"></path>`
      : '';
    const points = funds.filter(p => !shellState.scatterHidden.has(p.name)).map(p => `<circle cx="${xPos(p.x)}" cy="${yPos(p.y)}" r="6.8" fill="${p.color}" stroke="#fff" stroke-width="1.5" data-point-name="${escapeHtml(p.name)}" data-point-risk="${p.x.toFixed(2)}" data-point-return="${p.y.toFixed(2)}"/>`).join('')
      + extras.filter(p => !shellState.scatterHidden.has(p.name)).map(p => `<circle cx="${xPos(p.x)}" cy="${yPos(p.y)}" r="9" fill="${p.color}" stroke="#1f2d3d" stroke-width="1.5" data-point-name="${escapeHtml(p.name)}" data-point-risk="${p.x.toFixed(2)}" data-point-return="${p.y.toFixed(2)}" ${p.name === 'Editable Target Mix' ? 'data-target-dot="1"' : ''}/>`).join('');
    els.riskReturnScatter.innerHTML = `<svg id="riskReturnScatterSvg" viewBox="0 0 ${w} ${h}" width="100%" height="100%" aria-label="risk return scatter">${grid}${axes}${diagonal}${riskLine}${frontierLine}${points}</svg>`;
    els.riskReturnLegend.innerHTML = funds.concat(extras).map(item => (
      `<button type="button" class="small-btn" data-toggle-series="${escapeHtml(item.name)}" style="margin:0 8px 8px 0;padding:4px 8px;opacity:${shellState.scatterHidden.has(item.name) ? 0.45 : 1};"><span style="width:10px;height:10px;border-radius:50%;background:${item.color};display:inline-block;margin-right:6px;"></span>${escapeHtml(item.name)}</button>`
    )).join('');
    const svg = byId('riskReturnScatterSvg');
    if (!svg) return;
    svg.onclick = event => {
      const p = event.target.closest('[data-point-name]');
      if (!p || !els.riskReturnSelected) return;
      els.riskReturnSelected.textContent = `${p.getAttribute('data-point-name')} | Risk ${p.getAttribute('data-point-risk')}% / Return ${p.getAttribute('data-point-return')}%`;
    };
    svg.onmousedown = event => {
      const target = event.target;
      if (target && target.getAttribute && target.getAttribute('data-target-dot') === '1') {
        shellState.scatterDragActive = true;
        const rect = svg.getBoundingClientRect();
        shellState.drag = {
          kind: 'frontier-drag',
          rect,
          maxX,
          minY,
          maxY,
          frontier: (shellState.frontier || []).slice()
        };
        event.preventDefault();
      }
    };
    svg.onmouseup = () => { shellState.scatterDragActive = false; if (shellState.drag && shellState.drag.kind === 'frontier-drag') shellState.drag = null; };
    svg.onmouseleave = () => { shellState.scatterDragActive = false; if (shellState.drag && shellState.drag.kind === 'frontier-drag') shellState.drag = null; };
  }

  function renderDrawdownBars(plan) {
    if (!els.drawdownBars) return;
    const tolerance = Number.isFinite(shellState.drawdownTolerance) ? shellState.drawdownTolerance : -20;
    els.drawdownBars.innerHTML = `<div class="inline-note">Drawdown details hidden（リスク許容度 ${Math.abs(tolerance).toFixed(1)}%）</div>`;
  }

  function renderPortfolioTable() {
    els.portfolioTableBody.innerHTML = shellState.portfolioItems.map(item => `
      <tr data-item-id="${escapeHtml(item.id)}">
        <td>${escapeHtml(item.name)}</td>
        <td>${escapeHtml(item.source)}</td>
        <td class="num"><input type="text" data-field="currentValue" value="${Number(item.currentValue || 0).toLocaleString('ja-JP')}" style="width:105px;"></td>
        <td class="num"><input type="number" step="0.01" data-field="targetUnits" value="${fmtInput(item.targetUnits, 2)}"></td>
        <td class="num"><input type="text" data-field="nav" value="${Number(item.nav || 0).toLocaleString('ja-JP', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}" style="width:92px;"></td>
        <td class="num">${fmtPct(item.currentWeight)}</td>
        <td class="num"><input type="number" step="0.01" data-field="targetWeightPct" value="${fmtInput(item.targetWeight * 100, 2)}"></td>
        <td class="num"><input type="number" step="0.01" data-field="expectedReturnPct" value="${fmtInput(item.expectedReturn * 100, 2)}"></td>
        <td class="num"><input type="number" step="0.01" data-field="expectedRiskPct" value="${fmtInput(item.expectedRisk * 100, 2)}"></td>
        <td class="num">${fmtMoney(item.targetValue)}</td>
        <td class="num">${fmtMoney(item.diffValue)}</td>
        <td><button type="button" data-remove-item="${escapeHtml(item.id)}">×</button></td>
      </tr>
    `).join('');
  }

  function renderPortfolio() {
    const plan = buildPortfolioPlan();
    if (els.extraCashBreakdown) {
      if (shellState.investmentMode === 'savings') {
        const years = Math.max(1, Math.round(num(els.goalYears && els.goalYears.value, 10)));
        const annual = Math.max(0, (plan.extraCash || 0) / years);
        const monthly = annual / 12;
        els.extraCashBreakdown.style.display = '';
        els.extraCashBreakdown.textContent = `年換算: ${fmtMoney(Math.round(annual))} / 月換算: ${fmtMoney(Math.round(monthly))}`;
      } else {
        els.extraCashBreakdown.style.display = 'none';
        els.extraCashBreakdown.textContent = '';
      }
    }
    renderSummary(plan);
    renderDonut(plan);
    renderGoalProjection(plan);
    renderEditableMix(plan);
    renderUnitBarControl(plan);
    renderTradePlan();
    renderPortfolioTable();
    renderRiskReturnScatter(plan);
    renderDrawdownBars(plan);
    if (!shellState.portfolioItems.length) {
      showBanner('App1 または App2 から送るか、App3 で CSV を読み込んでください。', 'info');
    }
  }

  function redistributeTargetWeights(changedId, changedShare, snapshot) {
    const items = shellState.portfolioItems;
    const others = items.filter(item => item.id !== changedId);
    const otherTotal = others.reduce((sum, item) => sum + (snapshot.get(item.id) || 0), 0);
    items.forEach(item => {
      if (item.id === changedId) {
        item.targetWeight = changedShare;
      } else if (otherTotal > 0) {
        item.targetWeight = ((snapshot.get(item.id) || 0) / otherTotal) * (1 - changedShare);
      } else {
        item.targetWeight = others.length ? (1 - changedShare) / others.length : 0;
      }
    });
  }

  function handlePortfolioTableInput(event) {
    const input = event.target.closest('input[data-field]');
    if (!input) return;
    const row = input.closest('tr[data-item-id]');
    const item = shellState.portfolioItems.find(entry => entry.id === row.getAttribute('data-item-id'));
    if (!item) return;
    const field = input.getAttribute('data-field');
    const value = Number(String(input.value || '').replace(/,/g, ''));
    if (field === 'currentValue') item.currentValue = Math.max(0, value);
    if (field === 'units') item.units = Math.max(0, value);
    if (field === 'nav') item.nav = Math.max(1, value);
    if (field === 'targetWeightPct') item.targetWeight = Math.max(0, value / 100);
    if (field === 'targetUnits') {
      const targetUnits = Math.max(0, value);
      const targetValue = Math.max(0, targetUnits * Math.max(1, item.nav || 1));
      const otherTotal = shellState.portfolioItems
        .filter(entry => entry.id !== item.id)
        .reduce((sum, entry) => sum + Math.max(0, entry.targetValue || 0), 0);
      const total = otherTotal + targetValue;
      if (total > 0) {
        const itemShare = clamp(targetValue / total, 0, 1);
        const snapshot = new Map(shellState.portfolioItems.map(entry => [entry.id, entry.targetWeight || 0]));
        redistributeTargetWeights(item.id, itemShare, snapshot);
      }
    }
    if (field === 'expectedReturnPct') item.expectedReturn = value / 100;
    if (field === 'expectedRiskPct') item.expectedRisk = Math.max(0, value / 100);
    renderPortfolio();
  }

  function applyUnitBarShare(id, changedShare) {
    const item = shellState.portfolioItems.find(entry => entry.id === id);
    if (!item) return;
    const maxUnits = Math.max(1, ...shellState.portfolioItems.map(entry => {
      const targetUnits = entry.nav > 0 ? entry.targetValue / entry.nav : 0;
      return Math.max(entry.units || 0, targetUnits || 0);
    })) * 1.15;
    const targetUnits = clamp(changedShare, 0, 1) * maxUnits;
    const targetValue = Math.max(0, targetUnits * Math.max(1, item.nav || 1));
    const otherTotal = shellState.portfolioItems
      .filter(entry => entry.id !== id)
      .reduce((sum, entry) => sum + Math.max(0, entry.targetValue || 0), 0);
    const total = otherTotal + targetValue;
    if (total > 0) {
      const itemShare = clamp(targetValue / total, 0, 1);
      const snapshot = new Map(shellState.portfolioItems.map(entry => [entry.id, entry.targetWeight || 0]));
      redistributeTargetWeights(id, itemShare, snapshot);
    } else {
      const snapshot = new Map(shellState.portfolioItems.map(entry => [entry.id, entry.targetWeight || 0]));
      redistributeTargetWeights(id, changedShare, snapshot);
    }
    renderPortfolio();
  }

  function startUnitBarDrag(event) {
    const bar = event.target.closest('[data-unit-bar]');
    if (!bar) return;
    const id = bar.getAttribute('data-unit-bar');
    if (!id) return;
    const rect = bar.getBoundingClientRect();
    const share = clamp((event.clientX - rect.left) / Math.max(1, rect.width), 0, 1);
    shellState.drag = {
      kind: 'unit-bar',
      id,
      barRect: rect
    };
    applyUnitBarShare(id, share);
    event.preventDefault();
  }

  function handleEditableMixMouseDown(event) {
    const handle = event.target.closest('[data-mix-handle]');
    if (!handle) return;
    const id = handle.getAttribute('data-mix-handle');
    const item = shellState.portfolioItems.find(entry => entry.id === id);
    if (!item) return;
    shellState.drag = {
      kind: 'mix-handle',
      id,
      startX: event.clientX,
      startWeight: item.targetWeight || 0,
      snapshot: new Map(shellState.portfolioItems.map(entry => [entry.id, entry.targetWeight || 0]))
    };
    event.preventDefault();
  }

  function handleGlobalPointerMove(event) {
    if (!shellState.drag) return;
    if (shellState.drag.kind === 'mix-handle') {
      const delta = (event.clientX - shellState.drag.startX) / 220;
      const changedShare = clamp(shellState.drag.startWeight + delta, 0, 1);
      redistributeTargetWeights(shellState.drag.id, changedShare, shellState.drag.snapshot);
      renderPortfolio();
      return;
    }
    if (shellState.drag.kind === 'unit-bar') {
      const rect = shellState.drag.barRect;
      const changedShare = clamp((event.clientX - rect.left) / Math.max(1, rect.width), 0, 1);
      applyUnitBarShare(shellState.drag.id, changedShare);
      return;
    }
    if (shellState.drag.kind === 'frontier-drag') {
      const d = shellState.drag;
      const xRatio = clamp((event.clientX - d.rect.left) / Math.max(1, d.rect.width), 0, 1);
      const yRatio = 1 - clamp((event.clientY - d.rect.top) / Math.max(1, d.rect.height), 0, 1);
      const risk = xRatio * d.maxX;
      const ret = d.minY + yRatio * (d.maxY - d.minY);
      let nearest = null;
      let best = Infinity;
      d.frontier.forEach(fp => {
        const dx = fp.risk - risk;
        const dy = fp.ret - ret;
        const score = (dx * dx) + (dy * dy);
        if (score < best) {
          best = score;
          nearest = fp;
        }
      });
      if (nearest) {
        applyTargetWeights(nearest.weights);
        renderPortfolio();
      }
    }
  }

  function handleGlobalPointerUp() {
    shellState.drag = null;
    shellState.scatterDragActive = false;
  }

  function exportPortfolioCsv() {
    const header = ['fund', 'source', 'currentValue', 'targetWeight', 'targetValue', 'diffValue', 'units', 'nav', 'expectedReturn', 'expectedRisk'];
    const rows = shellState.portfolioItems.map(item => [
      item.name,
      item.source,
      Math.round(item.currentValue),
      item.targetWeight.toFixed(6),
      Math.round(item.targetValue),
      Math.round(item.diffValue),
      item.units.toFixed(4),
      item.nav.toFixed(2),
      item.expectedReturn.toFixed(6),
      item.expectedRisk.toFixed(6)
    ]);
    const csv = [header].concat(rows).map(cols => cols.map(value => `"${String(value).replace(/"/g, '""')}"`).join(',')).join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = `portfolio_rebalance_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  function removePortfolioItem(id) {
    shellState.portfolioItems = shellState.portfolioItems.filter(item => item.id !== id);
    renderPortfolio();
  }

  function allocateWeightToNewItem(item, desiredWeight) {
    shellState.portfolioItems.push(item);
    const snapshot = new Map(shellState.portfolioItems.map(entry => [entry.id, entry.targetWeight || 0]));
    redistributeTargetWeights(item.id, clamp(desiredWeight, 0, 0.4), snapshot);
    renderPortfolio();
  }

  function summarizeSeriesRows(series, fallbackName) {
    const rows = Array.isArray(series) ? series.filter(row => row && Number.isFinite(Number(row.nav))) : [];
    if (rows.length < 2) {
      return { name: fallbackName, nav: 10000, annualReturn: 0, annualVol: 0 };
    }
    const sorted = rows.slice().sort((a, b) => String(a.date).localeCompare(String(b.date)));
    const first = sorted[0];
    const last = sorted[sorted.length - 1];
    const dailyReturns = [];
    for (let i = 1; i < sorted.length; i += 1) {
      const prev = Number(sorted[i - 1].nav);
      const next = Number(sorted[i].nav);
      if (prev > 0 && Number.isFinite(next)) dailyReturns.push((next / prev) - 1);
    }
    const days = Math.max(1, (new Date(last.date) - new Date(first.date)) / 86400000);
    const totalReturn = Number(first.nav) > 0 ? (Number(last.nav) / Number(first.nav)) - 1 : 0;
    const annualReturn = Math.pow(1 + totalReturn, 365 / days) - 1;
    const mean = dailyReturns.reduce((sum, value) => sum + value, 0) / Math.max(1, dailyReturns.length);
    const variance = dailyReturns.length > 1
      ? dailyReturns.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) / (dailyReturns.length - 1)
      : 0;
    return { name: fallbackName, nav: Number(last.nav) || 10000, annualReturn, annualVol: Math.sqrt(Math.max(variance, 0)) * Math.sqrt(252) };
  }

  function summarizeApp2SeriesRows(series, fallbackName) {
    const rows = Array.isArray(series) ? series.filter(row => row && Number.isFinite(Number(row.fundIndex))) : [];
    if (rows.length < 2) {
      return { name: fallbackName, nav: 10000, annualReturn: 0, annualVol: 0, maxDrawdown: NaN };
    }
    const sorted = rows.slice().sort((a, b) => String(a.date || '').localeCompare(String(b.date || '')));
    const first = sorted[0];
    const last = sorted[sorted.length - 1];
    const dailyReturns = [];
    for (let i = 1; i < sorted.length; i += 1) {
      const prev = Number(sorted[i - 1].fundIndex);
      const next = Number(sorted[i].fundIndex);
      if (prev > 0 && Number.isFinite(next)) dailyReturns.push((next / prev) - 1);
    }
    const days = Math.max(1, (new Date(last.date) - new Date(first.date)) / 86400000);
    const totalReturn = Number(first.fundIndex) > 0 ? (Number(last.fundIndex) / Number(first.fundIndex)) - 1 : 0;
    const annualReturn = Math.pow(1 + totalReturn, 365 / days) - 1;
    const mean = dailyReturns.reduce((sum, value) => sum + value, 0) / Math.max(1, dailyReturns.length);
    const variance = dailyReturns.length > 1
      ? dailyReturns.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) / (dailyReturns.length - 1)
      : 0;
    const maxDrawdown = Math.min(...sorted.map(row => num(row.fundDrawdown, NaN)).filter(Number.isFinite), NaN);
    return {
      name: fallbackName,
      nav: 10000,
      annualReturn,
      annualVol: Math.sqrt(Math.max(variance, 0)) * Math.sqrt(252),
      maxDrawdown
    };
  }

  function getApp2FundSeriesCandidates() {
    const app2 = safeGetAppWindow('app2');
    if (!app2) return [];
    const seriesList = typeof app2.__codexGetFundSeriesAll === 'function'
      ? app2.__codexGetFundSeriesAll()
      : ((safeGetAppState(app2) && safeGetAppState(app2).fundSeriesAll) || []);
    const rawColors = typeof app2.__codexGetSeriesColors === 'function'
      ? app2.__codexGetSeriesColors()
      : ((safeGetAppState(app2) && safeGetAppState(app2).seriesColors) || {});
    const colors = buildResolvedApp2ColorMap(rawColors);
    let workbookRows = [];
    try { workbookRows = getApp2WorkbookRows(); } catch (error) { workbookRows = []; }
    return (Array.isArray(seriesList) ? seriesList : []).map((series, index) => {
      const rawName = (series && series.name) || `App2 Fund ${index + 1}`;
      let name = rawName;
      const m = String(rawName).match(/^(?:投資家|investor)\s*(\d+)$/i);
      if (m) {
        const ix = Math.max(0, Number(m[1]) - 1);
        const rowName = workbookRows[ix] && (workbookRows[ix].fund || workbookRows[ix].name);
        if (rowName) name = String(rowName);
      }
      const rows = (series && Array.isArray(series.visibleRows)) ? series.visibleRows : [];
      const stats = summarizeApp2SeriesRows(rows, name);
      return {
        id: `app2-csv-${slugify(name)}`,
        name,
        source: 'app2-csv',
        currentValue: 0,
        targetWeight: 0,
        units: 0,
        nav: stats.nav,
        expectedReturn: stats.annualReturn,
        expectedRisk: stats.annualVol,
        maxDrawdown: stats.maxDrawdown,
        color: pickColorByName(colors, name) || pickColorByName(colors, rawName) || undefined
      };
    });
  }

  function getApp2SelectedFundFiles() {
    const doc = safeGetAppDocument('app2');
    if (!doc) return [];
    const input = doc.getElementById('fundFile');
    if (!input || !input.files) return [];
    return Array.from(input.files || []);
  }

  function getApp2SelectedWorkbookFiles() {
    const doc = safeGetAppDocument('app2');
    if (!doc) return [];
    const input = doc.getElementById('workbookFile');
    if (!input || !input.files) return [];
    return Array.from(input.files || []);
  }

  function getApp2AverageMaxDrawdown() {
    const app2 = safeGetAppWindow('app2');
    const state = safeGetAppState(app2);
    const fromState = state && state.workbookAnalysis && state.workbookAnalysis.summary
      ? num(state.workbookAnalysis.summary.avgDD, NaN)
      : NaN;
    if (Number.isFinite(fromState)) {
      return Math.abs(fromState) > 1 ? fromState / 100 : fromState;
    }
    const doc = safeGetAppDocument('app2');
    if (!doc) return NaN;
    const text = (doc.getElementById('workbookSummaryMetrics') || { textContent: '' }).textContent || '';
    const m = text.match(/平均最大DD[^-\d]*(-?\d+(?:\.\d+)?)\s*%/);
    return m ? (Number(m[1]) / 100) : NaN;
  }

  function getApp2PlannerTargets() {
    const app2 = safeGetAppWindow('app2');
    const state = safeGetAppState(app2);
    const summary = state && state.workbookAnalysis && state.workbookAnalysis.summary
      ? state.workbookAnalysis.summary
      : null;
    const normalizeRatio = value => {
      const v = num(value, NaN);
      if (!Number.isFinite(v)) return NaN;
      return Math.abs(v) > 1 ? (v / 100) : v;
    };
    const minRet = summary ? normalizeRatio(summary.avgDD) : NaN;
    const expectedRet = summary ? normalizeRatio(summary.weightedReturn) : NaN;
    const maxRet = summary ? normalizeRatio(summary.bestReturnValue) : NaN;
    return { minRet, expectedRet, maxRet };
  }

  function getApp2RiskTolerance() {
    const doc = safeGetAppDocument('app2');
    if (!doc) return NaN;
    const input = doc.getElementById('riskTolerance');
    const v = input ? num(input.value, NaN) : NaN;
    if (!Number.isFinite(v)) return NaN;
    return v > 0 ? -v : v;
  }

  function getApp2RiskFreeRate() {
    const doc = safeGetAppDocument('app2');
    if (!doc) return NaN;
    const input = doc.getElementById('riskFree');
    return input ? num(input.value, NaN) : NaN;
  }

  function parseWorkbookRowsFromXlsx(file, xlsxLib) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = event => {
        try {
          const wb = xlsxLib.read(event.target.result, { type: 'array' });
          const firstSheet = wb.SheetNames && wb.SheetNames[0] ? wb.SheetNames[0] : null;
          if (!firstSheet) throw new Error('XLSX のシートが見つかりません。');
          const rows = xlsxLib.utils.sheet_to_json(wb.Sheets[firstSheet], { defval: '', raw: false });
          resolve(rows);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = () => reject(new Error('XLSX 読み込みに失敗しました。'));
      reader.readAsArrayBuffer(file);
    });
  }

  async function readCsvTextSmart(file) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    try {
      const utf8 = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
      if (/[�]/.test(utf8)) throw new Error('utf8 replacement character detected');
      return utf8;
    } catch (_) {
      try {
        return new TextDecoder('shift_jis').decode(bytes);
      } catch (_) {
        return new TextDecoder('utf-8').decode(bytes);
      }
    }
  }

  function normalizeWorkbookObjectRows(rows, sourceName) {
    const pick = (row, keys) => {
      for (const key of keys) {
        if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== '') return row[key];
      }
      return '';
    };
    return (Array.isArray(rows) ? rows : []).map((row, index) => {
      const fund = String(pick(row, ['ファンド', 'Fund', 'fund', '銘柄', '銘柄名']) || `Fund ${index + 1}`).trim();
      const totalInvestment = Math.max(0, num(pick(row, ['投資額合計', '投資総計', 'totalInvestment']), 0));
      const recurringInvestment = Math.max(0, num(pick(row, ['積立投資額', '積立額', '積立合計', 'recurringInvestment']), 0));
      const spotInvestment = Math.max(0, num(pick(row, ['スポット合計', '一括投資額', 'spotInvestment']), 0));
      const units = Math.max(0, num(pick(row, ['保有口数', '口数', 'units']), 0));
      const valuation = Math.max(0, num(pick(row, ['評価額', '現在評価額', 'valuation']), 0));
      const annualVol = Math.max(0, num(pick(row, ['標準偏差_年率', '標準偏差（年率換算）', 'annualVol']), 0));
      const maxDrawdown = num(pick(row, ['最大ドローダウン', '最大DD', 'maxDrawdown']), NaN);
      const returnPct = totalInvestment > 0 ? (valuation / totalInvestment) - 1 : num(pick(row, ['リターン', 'returnPct']), 0);
      return {
        id: `${sourceName}-${slugify(fund)}-${index}`,
        fund,
        totalInvestment,
        recurringInvestment,
        spotInvestment,
        units,
        valuation,
        annualVol,
        maxDrawdown,
        returnPct
      };
    }).filter(row => row.fund);
  }

  async function parseAdditionalFundCsv(file) {
    const fallbackName = file.name.replace(/\.[^.]+$/, '');
    const app2 = safeGetAppWindow('app2');
    if (app2 && typeof app2.parseSingleFundCsvFile === 'function') {
      const series = await app2.parseSingleFundCsvFile(file);
      return summarizeSeriesRows(series, fallbackName);
    }
    const text = await readCsvTextSmart(file);
    const lines = text.replace(/^\uFEFF/, '').split(/\r?\n/).filter(Boolean);
    if (lines.length < 2) throw new Error(`${file.name} のCSVを読み取れませんでした。`);
    const headers = lines[0].split(',');
    const dateIndex = headers.findIndex(h => /date|日付/i.test(h));
    const navIndex = headers.findIndex(h => /nav|price|基準価額|value/i.test(h));
    if (dateIndex < 0 || navIndex < 0) throw new Error(`${file.name} から日付またはNAV列を見つけられません。`);
    const rows = lines.slice(1).map(line => {
      const cells = line.split(',');
      return { date: cells[dateIndex], nav: Number(String(cells[navIndex] || '').replace(/,/g, '')) };
    }).filter(row => row.date && Number.isFinite(row.nav));
    return summarizeSeriesRows(rows, fallbackName);
  }

  async function importPortfolioCsvFiles(files) {
    const list = Array.from(files || []);
    if (!list.length) return;
    for (const file of list) {
      const isXlsx = /\.(xlsx|xls)$/i.test(file.name || '');
      if (isXlsx) {
        const app2 = safeGetAppWindow('app2');
        const xlsxLib = window.XLSX || (app2 && app2.XLSX);
        if (!xlsxLib) throw new Error('XLSX ライブラリが読み込まれていません。');
        const rawRows = await parseWorkbookRowsFromXlsx(file, xlsxLib);
        const workbookRows = normalizeWorkbookObjectRows(rawRows, 'xlsx-import');
        if (!workbookRows.length) throw new Error(`${file.name} からファンド行を抽出できませんでした。`);
        workbookRows.forEach((row, index) => {
          allocateWeightToNewItem(normalizeItem({
            id: `xlsx-${slugify(row.fund)}-${index}`,
            name: row.fund,
            source: 'xlsx',
            currentValue: row.valuation,
            units: row.units,
            nav: row.units > 0 ? row.valuation / row.units : 10000,
            expectedReturn: row.returnPct,
            expectedRisk: row.annualVol,
            maxDrawdown: row.maxDrawdown
          }, shellState.portfolioItems.length), 0.05);
        });
      } else {
        const metrics = await parseAdditionalFundCsv(file);
        allocateWeightToNewItem(normalizeItem({
          id: `candidate-${slugify(metrics.name)}`,
          name: metrics.name,
          source: 'csv',
          currentValue: 0,
          units: 0,
          nav: metrics.nav,
          expectedReturn: metrics.annualReturn,
          expectedRisk: metrics.annualVol,
          notes: 'CSV import'
        }, shellState.portfolioItems.length), 0.05);
      }
    }
    shellState.portfolioSource = `${list.length} file(s) imported`;
    showBanner(`${list.length} 件のファイルを App3 に読み込みました。`, 'info');
  }

  function workbookRowToItem(row, index, source) {
    return normalizeItem({
      id: `${source}-${slugify(row.fund || row.name || `fund-${index + 1}`)}`,
      name: row.fund || row.name || `Fund ${index + 1}`,
      source,
      currentValue: row.valuation,
      units: row.units,
      nav: row.units > 0 && row.valuation > 0 ? row.valuation / row.units : 10000,
      expectedReturn: row.returnPct,
      expectedRisk: row.annualVol,
      maxDrawdown: row.maxDrawdown,
      totalInvestment: row.totalInvestment,
      recurringInvestment: row.recurringInvestment,
      spotInvestment: row.spotInvestment,
      sharpe: row.sharpe,
      beta: row.beta,
      trackingError: row.trackingError,
      color: row.color || row.fundColor
    }, index);
  }

  function getApp1WorkbookRows() {
    const app1 = safeGetAppWindow('app1');
    const state = safeGetAppState(app1);
    const values = state && state.recalc ? Object.values(state.recalc) : [];
    if (!values.length) throw new Error('App1 に再計算データがありません。');
    return values.map(row => {
      const fund = String(row && row.name ? row.name : 'Fund').trim() || 'Fund';
      const totalInvestment = Math.max(0, num(row && row.invested, 0));
      const rawSpot = Math.max(0, num(row && row.spotTotal, 0));
      const recurringInvestment = Math.max(0, totalInvestment - rawSpot);
      const spotInvestment = Math.max(0, totalInvestment > 0 ? Math.min(rawSpot, totalInvestment) : rawSpot);
      const valuation = Math.max(0, num(row && row.value, 0));
      const profit = valuation - totalInvestment;
      const returnPct = totalInvestment > 0 ? (valuation / totalInvestment) - 1 : 0;
      const recurringRatio = totalInvestment > 0 ? recurringInvestment / totalInvestment : 0;
      const spotRatio = totalInvestment > 0 ? spotInvestment / totalInvestment : 0;
      return {
        fund,
        name: fund,
        displayName: fund,
        periodStart: String((row && row.periodStart) || '').trim(),
        periodEnd: String((row && row.periodEnd) || '').trim(),
        totalInvestment,
        recurringInvestment,
        spotInvestment,
        units: Math.max(0, num(row && row.units, 0)),
        valuation,
        annualVol: Math.max(0, num(row && row.annualizedStd, 0)),
        maxDrawdown: num(row && row.maxDrawdown, NaN),
        sharpe: num(row && row.sharpe, NaN),
        beta: num(row && row.beta, NaN),
        trackingError: num(row && row.trackingError, NaN),
        returnPct,
        profit,
        recurringRatio,
        spotRatio
      };
    });
  }

  function getApp1WorkbookRowsFromFundsState() {
    const app1 = safeGetAppWindow('app1');
    const state = safeGetAppState(app1);
    const funds = state && Array.isArray(state.funds) ? state.funds : [];
    const rows = funds
      .map(f => {
        const hist = Array.isArray(f && f.rows) ? f.rows : [];
        if (!hist.length) return null;
        const first = hist[0];
        const last = hist[hist.length - 1];
        const valuation = Math.max(0, num(last && last.nav, 0));
        const totalInvestment = Math.max(0, num(f && (f.investAmount != null ? f.investAmount : 0), 0));
        const recurringInvestment = Math.max(0, num(f && (f.sipAmount != null ? f.sipAmount : 0), 0));
        const spotInvestment = totalInvestment;
        const units = Math.max(0, valuation > 0 ? totalInvestment / Math.max(1, valuation) : 0);
        const fundName = String((f && f.name) || '').trim() || 'Fund';
        return {
          fund: fundName,
          name: fundName,
          displayName: fundName,
          periodStart: String((first && (first.dateISO || first.date)) || '').trim(),
          periodEnd: String((last && (last.dateISO || last.date)) || '').trim(),
          totalInvestment,
          recurringInvestment,
          spotInvestment,
          units,
          valuation,
          annualVol: 0,
          maxDrawdown: NaN,
          sharpe: NaN,
          beta: NaN,
          trackingError: NaN,
          returnPct: totalInvestment > 0 ? (valuation / totalInvestment) - 1 : 0,
          profit: valuation - totalInvestment,
          recurringRatio: totalInvestment > 0 ? recurringInvestment / totalInvestment : 0,
          spotRatio: totalInvestment > 0 ? spotInvestment / totalInvestment : 0
        };
      })
      .filter(Boolean);
    return rows;
  }

  function getApp2WorkbookRows() {
    const tableRows = getApp2WorkbookRowsFromTable();
    if (tableRows.length && tableRows.some(row => num(row.valuation, 0) > 0)) {
      return tableRows.map(row => ({
        fund: row.fund || row.displayName || 'Fund',
        periodStart: row.periodStart || '',
        periodEnd: row.periodEnd || '',
        totalInvestment: Math.max(0, num(row.totalInvestment, 0)),
        recurringInvestment: Math.max(0, num(row.recurringInvestment, 0)),
        spotInvestment: Math.max(0, num(row.spotInvestment, 0)),
        units: Math.max(0, num(row.units ?? row.holdingUnits ?? row.unitCount, 0)),
        valuation: Math.max(0, num(row.valuation ?? row.value ?? row.currentValue, 0)),
        annualVol: Math.max(0, num(row.annualVol, 0)),
        maxDrawdown: num(row.maxDrawdown, NaN),
        sharpe: num(row.sharpe, NaN),
        beta: num(row.beta, NaN),
        trackingError: num(row.trackingError, NaN),
        returnPct: Number.isFinite(num(row.totalInvestment, NaN)) && num(row.totalInvestment, 0) > 0
          ? (num(row.valuation, 0) / num(row.totalInvestment, 1)) - 1
          : num(row.returnPct, 0)
      }));
    }
    const app2 = safeGetAppWindow('app2');
    let rows = [];
    if (app2 && typeof app2.__codexGetWorkbookRows === 'function') {
      rows = app2.__codexGetWorkbookRows();
    } else {
      const state = safeGetAppState(app2);
      rows = state && Array.isArray(state.workbookRows) ? state.workbookRows : [];
    }
    if (!rows.length) rows = tableRows;
    if (!rows.length) throw new Error('App2 に XLSX データがありません。');
    return rows.map(row => ({
      fund: row.fund || row.displayName || 'Fund',
      periodStart: row.periodStart || '',
      periodEnd: row.periodEnd || '',
      totalInvestment: Math.max(0, num(row.totalInvestment, 0)),
      recurringInvestment: Math.max(0, num(row.recurringInvestment, 0)),
      spotInvestment: Math.max(0, num(row.spotInvestment, 0)),
      units: Math.max(0, num(row.units ?? row.holdingUnits ?? row.unitCount, 0)),
      valuation: Math.max(0, num(row.valuation ?? row.value ?? row.currentValue, 0)),
      annualVol: Math.max(0, num(row.annualVol, 0)),
      maxDrawdown: num(row.maxDrawdown, NaN),
      sharpe: num(row.sharpe, NaN),
      beta: num(row.beta, NaN),
      trackingError: num(row.trackingError, NaN),
      returnPct: Number.isFinite(num(row.totalInvestment, NaN)) && num(row.totalInvestment, 0) > 0
        ? (num(row.valuation, 0) / num(row.totalInvestment, 1)) - 1
        : num(row.returnPct, 0)
    }));
  }

  function getApp2WorkbookRowsFromTable() {
    const doc = safeGetAppDocument('app2');
    if (!doc) return [];
    const tableRows = Array.from(doc.querySelectorAll('#workbookTable table tbody tr'));
    const parsed = tableRows.map(tr => {
      const tds = Array.from(tr.querySelectorAll('td')).map(td => td.textContent.trim());
      if (tds.length < 10) return null;
      const fund = tds[0] || '';
      const periodText = tds[1] || '';
      const [periodStart, periodEnd] = periodText.split(/～|~|-/).map(v => (v || '').trim());
      const totalInvestment = num(tds[2], NaN);
      const valuation = num(tds[3], NaN);
      const annualVolRaw = num(tds[8], NaN);
      const maxDrawdownRaw = num(tds[9], NaN);
      const annualVol = Number.isFinite(annualVolRaw) ? annualVolRaw / 100 : NaN;
      const maxDrawdown = Number.isFinite(maxDrawdownRaw) ? maxDrawdownRaw / 100 : NaN;
      const sharpe = num(tds[10], NaN);
      const beta = num(tds[11], NaN);
      const teRaw = num(tds[12], NaN);
      const trackingError = Number.isFinite(teRaw) ? teRaw / 100 : NaN;
      if (!fund || !Number.isFinite(totalInvestment) || !Number.isFinite(valuation)) return null;
      return {
        fund,
        periodStart: periodStart || '',
        periodEnd: periodEnd || '',
        totalInvestment,
        recurringInvestment: Math.max(0, num(tds[6], 0) / 100 * totalInvestment),
        spotInvestment: Math.max(0, num(tds[7], 0) / 100 * totalInvestment),
        units: Math.max(0, num(tds[3], 0) / 10000),
        valuation,
        annualVol,
        maxDrawdown,
        sharpe,
        beta,
        trackingError,
        returnPct: totalInvestment > 0 ? (valuation / totalInvestment) - 1 : 0
      };
    }).filter(Boolean);
    return parsed;
  }

  function syncApp1ToPortfolio() {
    const rows = getApp1WorkbookRows();
    shellState.portfolioItems = rows.map((row, index) => workbookRowToItem(row, index, 'app1'));
    shellState.portfolioSource = 'App1 recalculation';
    renderPortfolio();
    showBanner('App1 の再計算結果を App3 に反映しました。', 'success');
    switchTab('panelApp3');
  }

  async function importApiFundsToApp1(codes) {
    if (!shellState.frameReady.app1) {
      throw new Error('App1 がまだ読み込み中です。少し待ってから再実行してください。');
    }
    const app1 = safeGetAppWindow('app1');
    if (!app1 || typeof app1.__codexImportApiFundsToState !== 'function') {
      throw new Error('App1 API bridge が利用できません。');
    }

    const list = Array.isArray(codes) ? codes.map(v => String(v || '').trim()).filter(Boolean) : [];
    if (!list.length) {
      throw new Error('code の配列が必要です。');
    }

    const result = await app1.__codexImportApiFundsToState(list);
    if (typeof app1.renderAllFromVisibleRange === 'function') {
      try { app1.renderAllFromVisibleRange(); } catch (e) { /* no-op */ }
    }
    if (typeof app1.renderWorkbookArea === 'function') {
      try { app1.renderWorkbookArea(); } catch (e) { /* no-op */ }
    }
    resizeIframe(els.frameApp1);
    setShellStatus(`App1 API import: ${result && result.withHistory || 0}/${result && result.count || list.length}`);
    showBanner(`App1にAPIファンドを取込: ${result && result.withHistory || 0}/${result && result.count || list.length}`, 'success');
    return result;
  }

  function getFundSearchSharedSelection() {
    const key = 'fund_search_selected_for_apps_v1';
    try {
      const raw = window.localStorage.getItem(key);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      const funds = Array.isArray(parsed && parsed.funds) ? parsed.funds : [];
      return funds.filter(Boolean);
    } catch (error) {
      return [];
    }
  }

  function convertFundSearchToWorkbookRows(funds) {
    return (Array.isArray(funds) ? funds : []).map(f => {
      const rows = Array.isArray(f.rows) ? f.rows : [];
      const first = rows[0] && rows[0].dateISO ? rows[0].dateISO : '';
      const last = rows.length ? (rows[rows.length - 1].dateISO || '') : '';
      const latestNav = rows.length ? num(rows[rows.length - 1].nav, 0) : num(f.nav, 0);
      return {
        fund: String(f.name || f.code || ''),
        code: String(f.code || ''),
        company: String(f.company || ''),
        periodStart: first,
        periodEnd: last,
        totalInvestment: 0,
        recurringInvestment: 0,
        spotInvestment: 0,
        valuation: latestNav,
        annualVol: NaN,
        maxDrawdown: NaN,
        sharpe: NaN,
        beta: NaN,
        trackingError: NaN
      };
    });
  }

  function convertFundSearchToSeries(funds) {
    return (Array.isArray(funds) ? funds : []).map((f, i) => ({
      name: String(f.name || f.code || `Fund ${i + 1}`),
      key: `fund_search_${i}`,
      code: String(f.code || ''),
      company: String(f.company || ''),
      visibleRows: (Array.isArray(f.rows) ? f.rows : []).map(r => ({
        date: r.dateISO || r.date || '',
        nav: num(r.nav, NaN)
      })).filter(r => r.date && Number.isFinite(r.nav))
    }));
  }

  function convertFundSearchToBenchmarkRows(funds) {
    const bm = (Array.isArray(funds) ? funds : []).find(f => String(f.transferKind || 'fund') === 'benchmark');
    if (!bm) return [];
    return (Array.isArray(bm.rows) ? bm.rows : [])
      .map(r => ({
        date: r.dateISO || r.date || '',
        price: num(r.nav, NaN)
      }))
      .filter(r => r.date && Number.isFinite(r.price));
  }

  function normalizeFundSearchForTransfer(funds) {
    const list = Array.isArray(funds) ? funds.slice() : [];
    const fundOnly = list.filter(f => String(f.transferKind || 'fund') !== 'benchmark');
    const hasFundRows = fundOnly.some(f => Array.isArray(f && f.rows) && f.rows.length > 0);
    if (hasFundRows) return list;

    const fallback = list.find(f => Array.isArray(f && f.rows) && f.rows.length > 0);
    if (!fallback) return list;

    const code = String((fallback && fallback.code) || '').trim();
    if (!code) return list;

    const existsFund = list.some(f => String((f && f.code) || '').trim() === code && String(f.transferKind || 'fund') !== 'benchmark');
    if (existsFund) return list;

    const clone = {
      ...fallback,
      transferKind: 'fund'
    };
    return [clone, ...list];
  }

  function extractFundRowsForTransfer(funds) {
    const list = Array.isArray(funds) ? funds : [];
    // Prefer explicit fund rows; if none, fallback to any rows so fund pipeline never becomes empty.
    const explicit = list.filter(f => String(f.transferKind || 'fund') !== 'benchmark' && Array.isArray(f && f.rows) && f.rows.length > 0);
    if (explicit.length) return explicit;
    return list.filter(f => Array.isArray(f && f.rows) && f.rows.length > 0);
  }

  function buildMergedRowsForApp2(fundRows, benchmarkRows, reinvest) {
    const f = (Array.isArray(fundRows) ? fundRows : [])
      .map(r => ({
        date: normalizeDateKey(r && (r.date || r.dateISO)),
        nav: num(r && r.nav, NaN),
        distribution: 0
      }))
      .filter(r => r.date && Number.isFinite(r.nav))
      .sort((a, b) => a.date.localeCompare(b.date));
    if (!f.length) return [];

    const bList = (Array.isArray(benchmarkRows) ? benchmarkRows : [])
      .map(r => ({ date: normalizeDateKey(r && (r.date || r.dateISO)), price: num(r && (r.price ?? r.nav), NaN) }))
      .filter(r => r.date && Number.isFinite(r.price))
      .sort((a, b) => a.date.localeCompare(b.date));
    const bMap = new Map(bList.map(r => [r.date, r.price]));

    const dates = f.map(x => x.date);
    const baseIndex = 100;
    const out = [];
    let units = 1;
    let bi = 0;
    let lastBenchPrice = NaN;
    for (let i = 0; i < dates.length; i += 1) {
      const date = dates[i];
      const fr = f.find(x => x.date === date);
      if (!fr) continue;
      if (i > 0 && reinvest && Number.isFinite(fr.distribution) && fr.distribution !== 0) {
        const prevNav = out[i - 1] && out[i - 1].fundNav;
        if (prevNav > 0) units *= (1 + fr.distribution / prevNav);
      }
      const fundAdj = fr.nav * units;
      let benchPrice = bMap.has(date) ? bMap.get(date) : NaN;
      // If benchmark dates are not perfectly aligned, use the latest benchmark value up to the fund date.
      if (!Number.isFinite(benchPrice) && bList.length) {
        while (bi < bList.length && bList[bi].date <= date) {
          lastBenchPrice = bList[bi].price;
          bi += 1;
        }
        benchPrice = lastBenchPrice;
      }
      out.push({
        date,
        fundNav: fr.nav,
        distribution: fr.distribution,
        fundAdj,
        benchPrice,
        fundReturn: NaN,
        benchReturn: NaN,
        excessReturn: NaN,
        fundIndex: baseIndex,
        benchIndex: Number.isFinite(benchPrice) ? baseIndex : NaN,
        fundDrawdown: NaN,
        benchDrawdown: NaN,
        rollingExcess: NaN
      });
    }
    if (!out.length) return [];
    for (let i = 0; i < out.length; i += 1) {
      out[i].fundIndex = (out[i].fundAdj / out[0].fundAdj) * baseIndex;
      if (Number.isFinite(out[i].benchPrice) && Number.isFinite(out[0].benchPrice) && out[0].benchPrice !== 0) {
        out[i].benchIndex = (out[i].benchPrice / out[0].benchPrice) * baseIndex;
      }
    }
    for (let i = 1; i < out.length; i += 1) {
      const p = out[i - 1];
      const c = out[i];
      c.fundReturn = p.fundAdj !== 0 ? (c.fundAdj / p.fundAdj) - 1 : NaN;
      c.benchReturn = (Number.isFinite(c.benchPrice) && Number.isFinite(p.benchPrice) && p.benchPrice !== 0)
        ? (c.benchPrice / p.benchPrice) - 1
        : NaN;
      c.excessReturn = Number.isFinite(c.fundReturn) && Number.isFinite(c.benchReturn) ? (c.fundReturn - c.benchReturn) : NaN;
    }
    const rollingWindow = 60;
    for (let i = 0; i < out.length; i += 1) {
      if (i < rollingWindow || !Number.isFinite(out[i].fundReturn) || !Number.isFinite(out[i].benchReturn)) {
        out[i].rollingExcess = NaN;
        continue;
      }
      let fundCum = 1;
      let benchCum = 1;
      let ok = true;
      for (let j = i - rollingWindow + 1; j <= i; j += 1) {
        const fr = out[j] && out[j].fundReturn;
        const br = out[j] && out[j].benchReturn;
        if (!Number.isFinite(fr) || !Number.isFinite(br)) {
          ok = false;
          break;
        }
        fundCum *= (1 + fr);
        benchCum *= (1 + br);
      }
      out[i].rollingExcess = ok ? ((fundCum - 1) - (benchCum - 1)) : NaN;
    }
    let fundPeak = -Infinity;
    let benchPeak = -Infinity;
    for (let i = 0; i < out.length; i += 1) {
      fundPeak = Math.max(fundPeak, out[i].fundIndex);
      out[i].fundDrawdown = fundPeak > 0 ? (out[i].fundIndex / fundPeak) - 1 : NaN;
      if (Number.isFinite(out[i].benchIndex)) {
        benchPeak = Math.max(benchPeak, out[i].benchIndex);
        out[i].benchDrawdown = benchPeak > 0 ? (out[i].benchIndex / benchPeak) - 1 : NaN;
      }
    }
    return out;
  }

  function deriveWorkbookRowsFromSeries(seriesList) {
    const list = Array.isArray(seriesList) ? seriesList : [];
    return list.map((series, idx) => {
      const rows = (Array.isArray(series && series.visibleRows) ? series.visibleRows : [])
        .map(r => ({
          date: normalizeDateKey(r && (r.date || r.dateISO)),
          nav: num(r && (r.fundAdj ?? r.fundNav ?? r.nav ?? r.price ?? r.value), NaN)
        }))
        .filter(r => r.date && Number.isFinite(r.nav))
        .sort((a, b) => a.date.localeCompare(b.date));
      if (!rows.length) return null;
      const first = rows[0];
      const last = rows[rows.length - 1];
      const daily = [];
      for (let i = 1; i < rows.length; i += 1) {
        const p = rows[i - 1].nav;
        const c = rows[i].nav;
        if (p > 0 && Number.isFinite(c)) daily.push((c / p) - 1);
      }
      const mean = daily.length ? (daily.reduce((a, b) => a + b, 0) / daily.length) : 0;
      const variance = daily.length ? (daily.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / daily.length) : 0;
      const annualVol = Math.sqrt(Math.max(0, variance)) * Math.sqrt(252);
      return {
        fund: String((series && series.name) || `ファンド${idx + 1}`),
        periodStart: first.date,
        periodEnd: last.date,
        totalInvestment: 1000000,
        recurringInvestment: 0,
        spotInvestment: 1000000,
        valuation: Math.max(0, Math.round(last.nav * 100)),
        annualVol: Number.isFinite(annualVol) ? annualVol : NaN,
        maxDrawdown: NaN,
        sharpe: NaN,
        beta: NaN,
        trackingError: NaN,
        returnPct: first.nav > 0 ? (last.nav / first.nav) - 1 : 0
      };
    }).filter(Boolean);
  }

  function buildFundCsvTextFromRows(rows) {
    const list = Array.isArray(rows) ? rows : [];
    const lines = ['Date,NAV'];
    list.forEach(r => {
      const d = String((r && (r.date || r.dateISO)) || '').trim();
      const nav = num(r && r.nav, NaN);
      if (!d || !Number.isFinite(nav)) return;
      lines.push(`${d},${Math.round(nav)}`);
    });
    return lines.join('\n');
  }

  function buildBenchmarkCsvTextFromRows(rows) {
    const list = Array.isArray(rows) ? rows : [];
    const lines = ['Date,Price'];
    list.forEach(r => {
      const d = String((r && r.date) || '').trim();
      const p = num(r && (r.price != null ? r.price : r.nav), NaN);
      if (!d || !Number.isFinite(p)) return;
      lines.push(`${d},${Math.round(p)}`);
    });
    return lines.join('\n');
  }

  function setInputFilesByText(inputEl, filesSpec) {
    if (!inputEl) return false;
    try {
      const dt = new DataTransfer();
      (Array.isArray(filesSpec) ? filesSpec : []).forEach(spec => {
        if (!spec || !spec.name) return;
        const content = String(spec.content || '');
        const file = new File([content], spec.name, { type: 'text/csv;charset=utf-8' });
        dt.items.add(file);
      });
      inputEl.files = dt.files;
      inputEl.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    } catch (e) {
      return false;
    }
  }

  function applyFundSearchToApp1Direct(app1, funds, benchmarkRows, bmDisplayName) {
    const s = safeGetAppState(app1);
    if (!s || !Array.isArray(s.funds)) return false;
    const list = (Array.isArray(funds) ? funds : []).filter(f => String(f.transferKind || 'fund') !== 'benchmark');
    let appliedFundCount = 0;
    list.forEach(f => {
      const code = String((f && f.code) || '').trim();
      if (!code) return;
      const rows = (Array.isArray(f && f.rows) ? f.rows : []).map(r => {
        const dateISO = String((r && (r.dateISO || r.date)) || '').trim();
        const d = new Date(dateISO);
        const nav = num(r && r.nav, NaN);
        if (!dateISO || !Number.isFinite(nav) || Number.isNaN(d.getTime())) return null;
        return { dateISO, date: d, nav };
      }).filter(Boolean);
      const exists = s.funds.some(x => x && x.source === 'api' && x.sourceCode === code);
      if (exists) return;
      const slot = s.funds.find(x => {
        const rowsEmpty = !Array.isArray(x && x.rows) || (x.rows || []).length === 0;
        return rowsEmpty;
      });
      const first = rows[0];
      const last = rows.length ? rows[rows.length - 1] : null;
      const payload = {
        id: (slot && slot.id) ? slot.id : (window.crypto && window.crypto.randomUUID ? window.crypto.randomUUID() : `api-${code}-${Date.now()}`),
        name: String((f && f.name) || code),
        rows,
        fileName: `API:${code}`,
        fileStatus: rows.length ? `API:${code}：${rows.length}件（${first ? first.dateISO : ''} ～ ${last ? last.dateISO : ''}）` : `API:${code}：0件`,
        investAmount: Number((slot && slot.investAmount) || 0),
        lumpBuyDate: String((slot && slot.lumpBuyDate) || ''),
        sipAmount: Number((slot && slot.sipAmount) || 10000),
        sipDay: Number((slot && slot.sipDay) || 1),
        sipStartDate: String((slot && slot.sipStartDate) || ''),
        visible: true,
        lastEncoding: 'api',
        source: 'api',
        sourceCode: code,
        sourceKind: 'fund'
      };
      if (slot) Object.assign(slot, payload);
      else s.funds.push(payload);
      appliedFundCount += 1;
    });
    if (s.bm && typeof s.bm === 'object') {
      s.bm.name = benchmarkRows.length ? bmDisplayName : (s.bm.name || '');
      s.bm.rows = (Array.isArray(benchmarkRows) ? benchmarkRows : []).map(r => {
        const dateISO = String((r && r.date) || '').trim();
        const d = new Date(dateISO);
        const nav = num(r && (r.price != null ? r.price : r.nav), NaN);
        if (!dateISO || !Number.isFinite(nav) || Number.isNaN(d.getTime())) return null;
        return { dateISO, date: d, nav };
      }).filter(Boolean);
    }
    return appliedFundCount > 0;
  }

  async function applyFundSearchToApp1ViaInputs(app1, seriesAll, benchmarkRows, bmDisplayName) {
    try {
      const doc = app1 && app1.document ? app1.document : null;
      if (!doc) return false;
      const series = (Array.isArray(seriesAll) ? seriesAll : [])
        .filter(s => s && Array.isArray(s.visibleRows) && s.visibleRows.length);
      if (!series.length) return false;

      const addBtn = doc.getElementById('btnAddFund');
      const getCards = () => Array.from(doc.querySelectorAll('#fundList .fundCard'));
      let cards = getCards();
      while (cards.length < series.length && addBtn) {
        addBtn.click();
        cards = getCards();
      }
      if (!cards.length) return false;

      let applied = 0;
      for (let i = 0; i < series.length && i < cards.length; i += 1) {
        const s = series[i];
        const card = cards[i];
        const nameInput = card.querySelector('[data-k="name"]');
        const fileInput = card.querySelector('[data-k="file"]');
        if (nameInput) {
          nameInput.value = String(s.name || `ファンド${i + 1}`);
          nameInput.dispatchEvent(new Event('input', { bubbles: true }));
        }
        if (!fileInput) continue;
        const ok = setInputFilesByText(fileInput, [{
          name: `${String((s && s.name) || `fund-${i + 1}`).replace(/[\\/:*?"<>|]/g, '_')}.csv`,
          content: buildFundCsvTextFromRows(s.visibleRows),
        }]);
        if (ok) applied += 1;
      }

      const bmInput = doc.getElementById('bmFile');
      const bmNameInput = doc.getElementById('bmName');
      if (bmInput) {
        if (Array.isArray(benchmarkRows) && benchmarkRows.length) {
          setInputFilesByText(bmInput, [{
            name: `${String(bmDisplayName || 'benchmark').replace(/[\\/:*?"<>|]/g, '_')}.csv`,
            content: buildBenchmarkCsvTextFromRows(benchmarkRows),
          }]);
        } else {
          setInputFilesByText(bmInput, []);
        }
      }
      if (bmNameInput && benchmarkRows.length) {
        bmNameInput.value = bmDisplayName;
        bmNameInput.dispatchEvent(new Event('input', { bubbles: true }));
      }

      if (typeof app1.runCalcSafe === 'function') {
        try { app1.runCalcSafe(); } catch (e) { /* no-op */ }
      } else if (typeof app1.renderAllFromVisibleRange === 'function') {
        app1.renderAllFromVisibleRange();
      }
      return applied > 0;
    } catch (e) {
      return false;
    }
  }

  function countApp1TransferredFunds(app1, codes) {
    try {
      const s = safeGetAppState(app1);
      if (!s || !Array.isArray(s.funds)) return 0;
      const codeSet = new Set((Array.isArray(codes) ? codes : []).map(c => String(c || '').trim()).filter(Boolean));
      return s.funds.filter(f => {
        if (!f) return false;
        const isApi = String(f.source || '') === 'api';
        const c = String(f.sourceCode || '').trim();
        return isApi && c && (codeSet.size === 0 || codeSet.has(c));
      }).length;
    } catch (e) {
      return 0;
    }
  }

  function applyFundSearchToApp2Direct(app2, workbookRows, seriesAll, benchmarkRows, bmDisplayName) {
    const s = safeGetAppState(app2);
    if (!s) return false;
    s.workbookRows = Array.isArray(workbookRows) ? workbookRows : [];
    s.workbookAnalysis = null;
    s.workbookLastError = '';
    const mainSeries = Array.isArray(seriesAll) && seriesAll.length ? seriesAll[0] : null;
    const addSeries = Array.isArray(seriesAll) && seriesAll.length > 1 ? seriesAll.slice(1) : [];
    const merged = buildMergedRowsForApp2(mainSeries ? mainSeries.visibleRows : [], benchmarkRows, false);
    s.merged = merged;
    s.additionalFundSeries = addSeries.map((x, idx) => ({
      name: String((x && x.name) || `追加ファンド${idx + 1}`),
      key: String((x && x.key) || `series${idx + 1}`),
      visibleRows: buildMergedRowsForApp2(x && x.visibleRows, benchmarkRows, false)
    }));
    s.fundSeriesAll = [
      ...(merged.length ? [{ name: String((mainSeries && mainSeries.name) || '主ファンド'), key: 'series0', visibleRows: merged }] : []),
      ...s.additionalFundSeries
    ];
    s.benchmarkRows = Array.isArray(benchmarkRows) ? benchmarkRows : [];
    s.hasBenchmark = s.benchmarkRows.length > 0;
    if (merged.length) {
      s.chartRange = s.chartRange || {};
      s.chartRange.start = merged[0].date;
      s.chartRange.end = merged[merged.length - 1].date;
    }
    try {
      const doc = app2.document;
      if (doc) {
        const fundNameInput = doc.getElementById('fundName');
        const bmNameInput = doc.getElementById('benchmarkName');
        if (fundNameInput && seriesAll.length) fundNameInput.value = String(seriesAll[0].name || 'ファンド');
        if (bmNameInput && s.benchmarkRows.length) bmNameInput.value = bmDisplayName;
      }
    } catch (e) { /* no-op */ }
    return true;
  }

  async function applyFundSearchToApp2ViaInputs(app2, seriesAll, benchmarkRows, bmDisplayName) {
    try {
      const doc = app2 && app2.document ? app2.document : null;
      if (!doc) return false;
      const fundInput = doc.getElementById('fundFile');
      const bmInput = doc.getElementById('benchmarkFile');
      const fundNameInput = doc.getElementById('fundName');
      const bmNameInput = doc.getElementById('benchmarkName');
      const firstSeries = Array.isArray(seriesAll) && seriesAll.length ? seriesAll[0] : null;
      const fundRows = firstSeries && Array.isArray(firstSeries.visibleRows) ? firstSeries.visibleRows : [];
      if (!fundRows.length || !fundInput) return false;

      const okFund = setInputFilesByText(fundInput, [{
        name: `${String((firstSeries && firstSeries.name) || 'fund').replace(/[\\/:*?"<>|]/g, '_')}.csv`,
        content: buildFundCsvTextFromRows(fundRows)
      }]);
      if (!okFund) return false;

      if (bmInput) {
        if (Array.isArray(benchmarkRows) && benchmarkRows.length) {
          setInputFilesByText(bmInput, [{
            name: `${String(bmDisplayName || 'benchmark').replace(/[\\/:*?"<>|]/g, '_')}.csv`,
            content: buildBenchmarkCsvTextFromRows(benchmarkRows)
          }]);
        } else {
          setInputFilesByText(bmInput, []);
        }
      }

      if (fundNameInput && firstSeries) fundNameInput.value = String(firstSeries.name || 'ファンド');
      if (bmNameInput && benchmarkRows.length) bmNameInput.value = bmDisplayName;
      if (typeof app2.updateSelectedFileInfos === 'function') {
        try { app2.updateSelectedFileInfos(); } catch (e) { /* no-op */ }
      }
      if (typeof app2.renderAllFromVisibleRange === 'function') {
        app2.renderAllFromVisibleRange();
      } else if (typeof app2.renderWorkbookArea === 'function') {
        app2.renderWorkbookArea();
      }
      return true;
    } catch (e) {
      return false;
    }
  }

  async function syncFundSearchToAllApps() {
    const funds = normalizeFundSearchForTransfer(getFundSearchSharedSelection());
    if (!funds.length) {
      throw new Error('Fund Search 側で選択済みファンドがありません。先に project/web/index.html で追加してください。');
    }
    const fundsForCsv = extractFundRowsForTransfer(funds);
    const fundsForFundPipeline = fundsForCsv.length ? fundsForCsv : funds;
    const workbookRows = convertFundSearchToWorkbookRows(fundsForFundPipeline);
    const seriesAll = convertFundSearchToSeries(fundsForFundPipeline);
    const benchmarkRows = convertFundSearchToBenchmarkRows(funds);
    const bmSelectedFund = (Array.isArray(funds) ? funds : []).find(f => String(f.transferKind || 'fund') === 'benchmark');
    const bmDisplayName = String((bmSelectedFund && (bmSelectedFund.name || bmSelectedFund.code)) || 'Benchmark').trim() || 'Benchmark';

    const app1 = safeGetAppWindow('app1');
    const app2 = safeGetAppWindow('app2');
    let appliedApp1 = false;
    let appliedApp2 = false;

    if (app1 && typeof app1.__codexImportApiFundsToState === 'function') {
      const app1FundCodes = fundsForFundPipeline
        .map(f => String(f.code || '').trim())
        .filter(Boolean);
      if (app1FundCodes.length) {
        try { await app1.__codexImportApiFundsToState(app1FundCodes); } catch (e) { /* no-op */ }
      }
      let importedCount = countApp1TransferredFunds(app1, app1FundCodes);
      if (!importedCount && typeof app1.addFundFromApi === 'function') {
        for (const code of app1FundCodes) {
          try { await app1.addFundFromApi(code); } catch (e) { /* no-op */ }
        }
        importedCount = countApp1TransferredFunds(app1, app1FundCodes);
      }
      if (!importedCount) {
        const directOk = applyFundSearchToApp1Direct(app1, fundsForFundPipeline, benchmarkRows, bmDisplayName);
        importedCount = directOk ? countApp1TransferredFunds(app1, app1FundCodes) : 0;
      }
      if (!importedCount) {
        const viaInputOk = await applyFundSearchToApp1ViaInputs(app1, seriesAll, benchmarkRows, bmDisplayName);
        importedCount = viaInputOk ? countApp1TransferredFunds(app1, app1FundCodes) : 0;
      }
      if (typeof app1.__codexSetBenchmarkRows === 'function') app1.__codexSetBenchmarkRows(benchmarkRows);
      try {
        const app1Doc = app1.document;
        if (app1Doc) {
          const bmNameInput = app1Doc.getElementById('bmName');
          const bmStatus = app1Doc.getElementById('bmStatus');
          if (bmNameInput && benchmarkRows.length) bmNameInput.value = bmDisplayName;
          if (bmStatus) {
            bmStatus.textContent = benchmarkRows.length
              ? `${bmDisplayName}：${benchmarkRows.length}件（${benchmarkRows[0].date} ～ ${benchmarkRows[benchmarkRows.length - 1].date}）`
              : '';
          }
        }
      } catch (error) { /* no-op */ }
      if (typeof app1.__codexMarkFundSearchImport === 'function') {
        app1.__codexMarkFundSearchImport({
          fundCount: workbookRows.length,
          benchmarkCount: benchmarkRows.length ? 1 : 0,
          timestamp: new Date().toLocaleString()
        });
      }
      if (typeof app1.renderAllFromVisibleRange === 'function') app1.renderAllFromVisibleRange();
      if (typeof app1.renderFunds === 'function') app1.renderFunds();
      if (typeof app1.renderWorkbookArea === 'function') app1.renderWorkbookArea();
      resizeIframe(els.frameApp1);
      appliedApp1 = importedCount > 0;
    } else if (app1) {
      const app1FundCodes = fundsForFundPipeline
        .map(f => String(f.code || '').trim())
        .filter(Boolean);
      if (typeof app1.addFundFromApi === 'function') {
        for (const code of app1FundCodes) {
          try { await app1.addFundFromApi(code); } catch (e) { /* no-op */ }
        }
      }
      appliedApp1 = applyFundSearchToApp1Direct(app1, fundsForFundPipeline, benchmarkRows, bmDisplayName);
      try {
        const d = app1.document;
        if (d) {
          const bmNameInput = d.getElementById('bmName');
          const bmStatus = d.getElementById('bmStatus');
          if (bmNameInput && benchmarkRows.length) bmNameInput.value = bmDisplayName;
          if (bmStatus) bmStatus.textContent = benchmarkRows.length
            ? `${bmDisplayName}：${benchmarkRows.length}件（${benchmarkRows[0].date} ～ ${benchmarkRows[benchmarkRows.length - 1].date}）`
            : '';
        }
      } catch (e) { /* no-op */ }
      if (typeof app1.renderFunds === 'function') app1.renderFunds();
      if (typeof app1.renderAllFromVisibleRange === 'function') app1.renderAllFromVisibleRange();
      if (typeof app1.renderWorkbookArea === 'function') app1.renderWorkbookArea();
      resizeIframe(els.frameApp1);
      if (!appliedApp1) {
        appliedApp1 = await applyFundSearchToApp1ViaInputs(app1, seriesAll, benchmarkRows, bmDisplayName);
        if (appliedApp1) resizeIframe(els.frameApp1);
      }
    }

    if (app2 && typeof app2.__codexSetWorkbookRows === 'function') {
      let okRows = app2.__codexSetWorkbookRows(workbookRows);
      const okSeries = (typeof app2.__codexSetFundSeriesAll === 'function') ? app2.__codexSetFundSeriesAll(seriesAll) : false;
      const okBm = (typeof app2.__codexSetBenchmarkRows === 'function') ? app2.__codexSetBenchmarkRows(benchmarkRows) : false;
      if (!okRows && typeof app2.__codexImportApiFundsToState === 'function') {
        const app2FundCodes = fundsForFundPipeline
          .map(f => String(f.code || '').trim())
          .filter(Boolean);
        if (app2FundCodes.length) {
          try {
            const result = await app2.__codexImportApiFundsToState(app2FundCodes);
            okRows = !!(result && result.ok);
          } catch (error) { /* no-op */ }
        }
      }
      try {
        const s = typeof app2.__codexGetState === 'function' ? app2.__codexGetState() : null;
        const doc = app2.document;
        if (doc) {
          const fundNameInput = doc.getElementById('fundName');
          const bmNameInput = doc.getElementById('benchmarkName');
          if (fundNameInput && seriesAll.length) fundNameInput.value = String(seriesAll[0].name || 'ファンド');
          if (bmNameInput && benchmarkRows.length) bmNameInput.value = bmDisplayName;
        }
        if (s && typeof s === 'object') {
          const mainSeries = Array.isArray(seriesAll) && seriesAll.length ? seriesAll[0] : null;
          const addSeries = Array.isArray(seriesAll) && seriesAll.length > 1 ? seriesAll.slice(1) : [];
          const merged = buildMergedRowsForApp2(mainSeries ? mainSeries.visibleRows : [], benchmarkRows, false);
          s.merged = merged;
          s.additionalFundSeries = addSeries.map((x, idx) => ({
            name: String((x && x.name) || `追加ファンド${idx + 1}`),
            key: String((x && x.key) || `series${idx + 1}`),
            visibleRows: buildMergedRowsForApp2(x && x.visibleRows, benchmarkRows, false)
          }));
          s.fundSeriesAll = [
            ...(merged.length ? [{ name: String((mainSeries && mainSeries.name) || '主ファンド'), key: 'series0', visibleRows: merged }] : []),
            ...s.additionalFundSeries
          ];
          s.benchmarkRows = benchmarkRows;
          s.hasBenchmark = benchmarkRows.length > 0;
          if (merged.length) {
            s.chartRange = s.chartRange || {};
            s.chartRange.start = merged[0].date;
            s.chartRange.end = merged[merged.length - 1].date;
          }
        }
      } catch (error) { /* no-op */ }
      if (typeof app2.__codexMarkFundSearchImport === 'function') {
        app2.__codexMarkFundSearchImport({
          fundCount: workbookRows.length,
          benchmarkCount: benchmarkRows.length ? 1 : 0,
          timestamp: new Date().toLocaleString()
        });
      }
      if (typeof app2.renderAllFromVisibleRange === 'function') {
        try { app2.renderAllFromVisibleRange(); } catch (error) { /* no-op */ }
      } else if (typeof app2.renderWorkbookArea === 'function') {
        app2.renderWorkbookArea();
      }
      resizeIframe(els.frameApp2);
      appliedApp2 = !!(okRows || okSeries || okBm);
      if (!appliedApp2) {
        appliedApp2 = await applyFundSearchToApp2ViaInputs(app2, seriesAll, benchmarkRows, bmDisplayName);
      }
    } else if (app2) {
      appliedApp2 = applyFundSearchToApp2Direct(app2, workbookRows, seriesAll, benchmarkRows, bmDisplayName);
      if (!appliedApp2) {
        appliedApp2 = await applyFundSearchToApp2ViaInputs(app2, seriesAll, benchmarkRows, bmDisplayName);
      }
      if (typeof app2.renderAllFromVisibleRange === 'function') {
        try { app2.renderAllFromVisibleRange(); } catch (error) { /* no-op */ }
      } else if (typeof app2.renderWorkbookArea === 'function') {
        app2.renderWorkbookArea();
      }
      resizeIframe(els.frameApp2);
    }

    if (!appliedApp1 && !appliedApp2) {
      throw new Error('App1/App2 へデータ注入できませんでした。ページ再読込後に再実行してください。');
    }
    if (shellState.frameReady.app2 && !appliedApp2) {
      throw new Error('App2 へのデータ注入に失敗しました。App2タブを一度開いてから再実行してください。');
    }

    shellState.portfolioItems = workbookRows.map((row, index) => workbookRowToItem(row, index, 'fund-search'));
    const fundCount = workbookRows.length;
    const bmCount = benchmarkRows.length ? 1 : 0;
    if (!fundCount && !bmCount) {
      throw new Error('送信対象データがありません。Fund Search で詳細取得済みの銘柄を選択してください。');
    }
    shellState.portfolioSource = `Fund Search transfer: funds ${fundCount} / benchmark ${bmCount}`;
    renderPortfolio();
    setSyncMark('app1', appliedApp1 ? 'ok' : 'ng');
    setSyncMark('app2', appliedApp2 ? 'ok' : 'ng');
    setSyncMark('app3', fundCount > 0 ? 'ok' : 'ng');
    showBanner(`Fund Search を反映: ファンド ${fundCount}件 / ベンチマーク ${bmCount}件`, 'success');
    setShellStatus('Fund Search -> App1/App2/App3 synced');
  }

  async function syncApp2ToPortfolio() {
    if (!shellState.frameReady.app2) {
      throw new Error('App2 がまだ読み込み中です。少し待ってから再実行してください。');
    }
    const app2 = safeGetAppWindow('app2');
    const s0 = safeGetAppState(app2);
    const hasSeriesBefore = !!(s0 && Array.isArray(s0.fundSeriesAll) && s0.fundSeriesAll.length);
    // Avoid auto-calling runAnalysisFromInputs here; it can trigger
    // "ファンドCSVまたは積立・スポット投資実績XLSXを選択してください。" alerts.
    let workbookRows = [];
    try {
      workbookRows = getApp2WorkbookRows();
    } catch (error) {
      workbookRows = [];
    }
    const workbookFiles = getApp2SelectedWorkbookFiles();
    if (workbookFiles.length && window.XLSX) {
      const parsedByName = new Map();
      for (const file of workbookFiles) {
        const rawRows = await parseWorkbookRowsFromXlsx(file, window.XLSX);
        const normalized = normalizeWorkbookObjectRows(rawRows, 'app2-xlsx');
        normalized.forEach(row => parsedByName.set(slugify(row.fund), row));
      }
      const hasValuation = workbookRows.some(row => num(row.valuation ?? row.value ?? row.currentValue, 0) > 0);
      if (!workbookRows.length || !hasValuation) {
        workbookRows = Array.from(parsedByName.values()).map(row => ({
          fund: row.fund,
          totalInvestment: row.totalInvestment,
          recurringInvestment: row.recurringInvestment,
          spotInvestment: row.spotInvestment,
          units: row.units,
          valuation: row.valuation,
          annualVol: row.annualVol,
          maxDrawdown: row.maxDrawdown,
          returnPct: row.returnPct
        }));
      } else {
        workbookRows = workbookRows.map(row => {
          const key = slugify(row.fund || row.name || '');
          const src = parsedByName.get(key);
          if (!src) return row;
          return {
            ...row,
            units: num(row.units, NaN) > 0 ? row.units : src.units,
            valuation: num(row.valuation, NaN) > 0 ? row.valuation : src.valuation,
            totalInvestment: num(row.totalInvestment, NaN) > 0 ? row.totalInvestment : src.totalInvestment
          };
        });
      }
    }
    const avgDdFallback = getApp2AverageMaxDrawdown();
    const riskTol = getApp2RiskTolerance();
    if (Number.isFinite(riskTol)) shellState.drawdownTolerance = riskTol;
    const riskFree = getApp2RiskFreeRate();
    if (Number.isFinite(riskFree) && els.portfolioRiskFree) {
      els.portfolioRiskFree.value = String(riskFree);
    }
    const colors = buildResolvedApp2ColorMap(Object.assign({}, getApp2SeriesColors(), getApp2FundColorMap()));
    const workbookItems = workbookRows.map((row, index) => {
      const rowName = row.fund || row.name;
      const rowColor = row.color || row.fundColor || pickColorByName(colors, rowName);
      const item = workbookRowToItem(row, index, 'app2');
      const c = rowColor || pickColorByName(colors, item.name);
      if (c) item.color = c;
      if (!Number.isFinite(item.maxDrawdown) || Math.abs(item.maxDrawdown) < 1e-8) {
        item.maxDrawdown = Number.isFinite(avgDdFallback) ? avgDdFallback : item.maxDrawdown;
      }
      return item;
    });
    const app2SeriesCandidates = getApp2FundSeriesCandidates();
    const fileCandidates = [];
    const app2FundFiles = getApp2SelectedFundFiles();
    for (const file of app2FundFiles) {
      const stats = await parseAdditionalFundCsv(file);
      fileCandidates.push(normalizeItem({
        id: `app2-fundcsv-${slugify(stats.name)}`,
        name: stats.name,
        source: 'app2-csv',
        currentValue: 0,
        targetWeight: 0,
        units: 0,
        nav: stats.nav,
        expectedReturn: stats.annualReturn,
        expectedRisk: stats.annualVol,
        maxDrawdown: NaN,
        color: pickColorByName(colors, stats.name)
      }, fileCandidates.length));
    }
    const mergedItems = workbookItems.slice();
    const hasName = name => mergedItems.some(item => slugify(item.name) === slugify(name));
    app2SeriesCandidates.forEach(candidate => {
      if (hasName(candidate.name)) return;
      mergedItems.push(normalizeItem(candidate, mergedItems.length));
    });
    fileCandidates.forEach(candidate => {
      if (hasName(candidate.name)) return;
      mergedItems.push(candidate);
    });
    shellState.portfolioItems = mergedItems;
    if (!shellState.portfolioItems.length) {
      throw new Error('App2 に転送可能な CSV/XLSX データがありません。');
    }
    shellState.portfolioSource = `App2 transfer: XLSX ${workbookItems.length} / CSV ${Math.max(0, mergedItems.length - workbookItems.length)}`;
    renderPortfolio();
    showBanner('App2 の CSV/XLSX データを App3 に反映しました。', 'success');
    switchTab('panelApp3');
  }

  async function syncApp1ToApp2() {
    let rows = [];
    try {
      rows = getApp1WorkbookRows();
    } catch (_e) {
      rows = getApp1WorkbookRowsFromFundsState();
      if (!rows.length) throw new Error('App1 に再計算データまたはCSV読込済みファンドがありません。');
    }
    const app2 = safeGetAppWindow('app2');
    if (!app2) throw new Error('App2 フレームがまだ準備中です。');
    const app1 = safeGetAppWindow('app1');
    const app1State = app1 ? safeGetAppState(app1) : null;
    const benchmarkRows = (() => {
      if (!app1State || typeof app1State !== 'object') return [];
      const direct = Array.isArray(app1State.benchmarkRows) ? app1State.benchmarkRows : [];
      if (direct.length) {
        return direct.map(r => ({
          date: normalizeDateKey(r && (r.date || r.dateISO)),
          price: num(r && (r.price ?? r.nav ?? r.value), NaN)
        })).filter(r => r.date && Number.isFinite(r.price));
      }
      const bmRows = Array.isArray(app1State.bm && app1State.bm.rows) ? app1State.bm.rows : [];
      return bmRows.map(r => ({
        date: normalizeDateKey(r && (r.date || r.dateISO)),
        price: num(r && (r.price ?? r.nav ?? r.value), NaN)
      })).filter(r => r.date && Number.isFinite(r.price));
    })();
    const app1Series = (() => {
      if (!app1State || typeof app1State !== 'object') return [];
      const fromAll = Array.isArray(app1State.fundSeriesAll) ? app1State.fundSeriesAll : [];
      if (fromAll.length) return fromAll;
      const funds = Array.isArray(app1State.funds) ? app1State.funds : [];
      return funds.map((f, idx) => ({
        name: String((f && (f.name || f.fundName)) || `ファンド${idx + 1}`),
        key: String((f && (f.key || f.code)) || `series${idx + 1}`),
        visibleRows: Array.isArray(f && f.rows) ? f.rows : []
      }));
    })();
    let normalizedSeries = app1Series.map((series, idx) => ({
      name: String((series && series.name) || `ファンド${idx + 1}`),
      key: String((series && series.key) || `series${idx + 1}`),
      visibleRows: (Array.isArray(series && series.visibleRows) ? series.visibleRows : [])
        .map(r => ({
          date: normalizeDateKey(r && (r.date || r.dateISO)),
          nav: num(r && (r.nav ?? r.fundAdj ?? r.fundNav ?? r.price ?? r.value ?? r.baseValue), NaN)
        }))
        .filter(r => r.date && Number.isFinite(r.nav))
    })).filter(s => s.visibleRows.length);
    if (!normalizedSeries.length && app1State && Array.isArray(app1State.merged) && app1State.merged.length) {
      const mergedRows = app1State.merged
        .map(r => ({
          date: normalizeDateKey(r && (r.date || r.dateISO)),
          nav: num(r && (r.fundAdj ?? r.fundNav ?? r.nav ?? r.price), NaN)
        }))
        .filter(r => r.date && Number.isFinite(r.nav));
      if (mergedRows.length) normalizedSeries = [{ name: '主ファンド', key: 'series0', visibleRows: mergedRows }];
      if (!benchmarkRows.length) {
        const bmFromMerged = app1State.merged
          .map(r => ({
            date: normalizeDateKey(r && (r.date || r.dateISO)),
            price: num(r && (r.benchPrice ?? r.benchmarkPrice ?? r.benchNav ?? r.benchIndex ?? r.benchmarkIndex), NaN)
          }))
          .filter(r => r.date && Number.isFinite(r.price));
        if (bmFromMerged.length) benchmarkRows.push(...bmFromMerged);
      }
    }
    if (app1State && Array.isArray(app1State.additionalFundSeries) && app1State.additionalFundSeries.length) {
      const extras = app1State.additionalFundSeries.map((series, idx) => ({
        name: String((series && series.name) || `追加ファンド${idx + 1}`),
        key: String((series && series.key) || `extra${idx + 1}`),
        visibleRows: (Array.isArray(series && series.visibleRows) ? series.visibleRows : [])
          .map(r => ({
            date: normalizeDateKey(r && (r.date || r.dateISO)),
            nav: num(r && (r.fundAdj ?? r.fundNav ?? r.nav ?? r.price ?? r.value), NaN)
          }))
          .filter(r => r.date && Number.isFinite(r.nav))
      })).filter(x => x.visibleRows.length);
      if (extras.length) {
        const exists = new Set(normalizedSeries.map(s => `${s.key}::${s.name}`));
        extras.forEach(x => {
          const k = `${x.key}::${x.name}`;
          if (!exists.has(k)) normalizedSeries.push(x);
        });
      }
    }

    if (typeof app2.__codexSetWorkbookRows === 'function') {
      app2.__codexSetWorkbookRows(rows);
    } else {
      const state = safeGetAppState(app2);
      if (!state) throw new Error('App2 の状態にアクセスできません。');
      state.workbookRows = rows;
      state.workbookAnalysis = null;
      state.workbookLastError = '';
    }
    if (typeof app2.__codexSetBenchmarkRows === 'function') {
      try { app2.__codexSetBenchmarkRows(benchmarkRows); } catch (e) { /* no-op */ }
    }
    if (typeof app2.__codexSetFundSeriesAll === 'function') {
      try { app2.__codexSetFundSeriesAll(normalizedSeries); } catch (e) { /* no-op */ }
    }
    const app2State = safeGetAppState(app2);
    if (app2State && typeof app2State === 'object') {
      const repairRollingRows = (rows, bmRows) => {
        const list = Array.isArray(rows) ? rows : [];
        if (!list.length) return list;
        const bList = (Array.isArray(bmRows) ? bmRows : [])
          .map(r => ({ date: normalizeDateKey(r && (r.date || r.dateISO)), price: num(r && (r.price ?? r.nav), NaN) }))
          .filter(x => x.date && Number.isFinite(x.price))
          .sort((a, b) => a.date.localeCompare(b.date));
        let bi = 0;
        let lastB = NaN;
        const out = list.map(r => ({ ...r }));
        for (let i = 0; i < out.length; i += 1) {
          const d = normalizeDateKey(out[i] && out[i].date);
          const fund = num(out[i] && (out[i].fundAdj ?? out[i].fundNav ?? out[i].nav ?? out[i].price), NaN);
          out[i].date = d;
          out[i].fundAdj = Number.isFinite(fund) ? fund : num(out[i] && out[i].fundAdj, NaN);
          while (bi < bList.length && bList[bi].date <= d) {
            lastB = bList[bi].price;
            bi += 1;
          }
          if (!Number.isFinite(num(out[i] && out[i].benchPrice, NaN))) out[i].benchPrice = lastB;
        }
        for (let i = 1; i < out.length; i += 1) {
          const p = out[i - 1];
          const c = out[i];
          const pFund = num(p && p.fundAdj, NaN);
          const cFund = num(c && c.fundAdj, NaN);
          const pBench = num(p && p.benchPrice, NaN);
          const cBench = num(c && c.benchPrice, NaN);
          c.fundReturn = Number.isFinite(pFund) && Number.isFinite(cFund) && pFund !== 0 ? (cFund / pFund) - 1 : NaN;
          c.benchReturn = Number.isFinite(pBench) && Number.isFinite(cBench) && pBench !== 0 ? (cBench / pBench) - 1 : NaN;
          c.excessReturn = Number.isFinite(c.fundReturn) && Number.isFinite(c.benchReturn) ? (c.fundReturn - c.benchReturn) : NaN;
        }
        const win = 60;
        for (let i = 0; i < out.length; i += 1) {
          if (i < win) { out[i].rollingExcess = NaN; continue; }
          let f = 1; let b = 1; let ok = true;
          for (let j = i - win + 1; j <= i; j += 1) {
            const fr = num(out[j] && out[j].fundReturn, NaN);
            const br = num(out[j] && out[j].benchReturn, NaN);
            if (!Number.isFinite(fr) || !Number.isFinite(br)) { ok = false; break; }
            f *= (1 + fr); b *= (1 + br);
          }
          out[i].rollingExcess = ok ? ((f - 1) - (b - 1)) : NaN;
        }
        return out;
      };

      const mainSeries = normalizedSeries.length ? normalizedSeries[0] : null;
      const addSeries = normalizedSeries.length > 1 ? normalizedSeries.slice(1) : [];
      const safeBenchmarkRows = benchmarkRows.length
        ? benchmarkRows
        : ((app1State && Array.isArray(app1State.merged))
          ? app1State.merged.map(r => ({
            date: normalizeDateKey(r && (r.date || r.dateISO)),
            price: num(r && (r.benchPrice ?? r.benchmarkPrice ?? r.benchNav ?? r.benchIndex ?? r.benchmarkIndex), NaN)
          })).filter(x => x.date && Number.isFinite(x.price))
          : []);
      const merged = repairRollingRows(
        buildMergedRowsForApp2(mainSeries ? mainSeries.visibleRows : [], safeBenchmarkRows, false),
        safeBenchmarkRows
      );
      app2State.merged = merged;
      app2State.additionalFundSeries = addSeries.map((x, idx) => ({
        name: String((x && x.name) || `追加ファンド${idx + 1}`),
        key: String((x && x.key) || `series${idx + 1}`),
        visibleRows: repairRollingRows(
          buildMergedRowsForApp2(x && x.visibleRows, safeBenchmarkRows, false),
          safeBenchmarkRows
        )
      }));
      app2State.fundSeriesAll = [
        ...(merged.length ? [{ name: String((mainSeries && mainSeries.name) || '主ファンド'), key: 'series0', visibleRows: merged }] : []),
        ...app2State.additionalFundSeries
      ];
      if (!Array.isArray(app2State.workbookRows) || !app2State.workbookRows.length) {
        app2State.workbookRows = deriveWorkbookRowsFromSeries(app2State.fundSeriesAll);
      } else {
        const hasVal = app2State.workbookRows.some(r => num(r && (r.valuation ?? r.currentValue ?? r.value), 0) > 0);
        if (!hasVal) app2State.workbookRows = deriveWorkbookRowsFromSeries(app2State.fundSeriesAll);
      }
      app2State.benchmarkRows = safeBenchmarkRows;
      const hasBenchmarkLikeRows = rows => (Array.isArray(rows) ? rows : []).some(r => (
        Number.isFinite(num(r && r.benchPrice, NaN)) ||
        Number.isFinite(num(r && r.benchReturn, NaN)) ||
        Number.isFinite(num(r && r.benchIndex, NaN)) ||
        Number.isFinite(num(r && r.excessReturn, NaN)) ||
        Number.isFinite(num(r && r.rollingExcess, NaN))
      ));
      app2State.hasBenchmark = safeBenchmarkRows.length > 0 || hasBenchmarkLikeRows(merged) || app2State.additionalFundSeries.some(x => hasBenchmarkLikeRows(x && x.visibleRows));
      app2State.bm = app2State.bm && typeof app2State.bm === 'object' ? app2State.bm : {};
      app2State.bm.rows = safeBenchmarkRows.map(r => ({
        dateISO: String(r.date || ''),
        date: new Date(String(r.date || '')),
        nav: num(r.price, NaN)
      })).filter(x => x.dateISO && Number.isFinite(x.nav) && !Number.isNaN(x.date.getTime()));
      if (!app2State.bm.name && app2State.bm.rows.length) app2State.bm.name = 'BM';
      if (merged.length) {
        app2State.chartRange = app2State.chartRange || {};
        app2State.chartRange.start = merged[0].date;
        app2State.chartRange.end = merged[merged.length - 1].date;
      }
      try {
        const doc = app2.document;
        const startInput = doc && doc.getElementById('startDate');
        const endInput = doc && doc.getElementById('endDate');
        const start = merged.length ? String(merged[0].date || '').slice(0, 10) : '';
        const end = merged.length ? String(merged[merged.length - 1].date || '').slice(0, 10) : '';
        if (startInput && start) startInput.value = start;
        if (endInput && end) endInput.value = end;
        if (startInput && endInput && startInput.value && endInput.value && startInput.value > endInput.value) {
          const tmp = startInput.value;
          startInput.value = endInput.value;
          endInput.value = tmp;
        }
        if (app2State.chartRange && startInput && endInput) {
          app2State.chartRange.start = startInput.value || app2State.chartRange.start;
          app2State.chartRange.end = endInput.value || app2State.chartRange.end;
        }
        const rollingOpt = doc && doc.getElementById('optRolling');
        if (rollingOpt && !rollingOpt.checked) {
          rollingOpt.checked = true;
          try { rollingOpt.dispatchEvent(new Event('change', { bubbles: true })); } catch (e) { /* no-op */ }
        }
        ['optIndex','optDrawdown','optRolling','optHoldingPerf','optWorkbook'].forEach(id => {
          const cb = doc && doc.getElementById(id);
          if (cb && !cb.checked) {
            cb.checked = true;
            try { cb.dispatchEvent(new Event('change', { bubbles: true })); } catch (e) { /* no-op */ }
          }
        });
      } catch (e) { /* no-op */ }
    }
    if (typeof app2.renderWorkbookArea === 'function') app2.renderWorkbookArea();
    if (typeof app2.renderAdditionalFundSections === 'function') {
      try { app2.renderAdditionalFundSections(); } catch (e) { /* no-op */ }
    }
    applyApp2AlertGuard();
    applyApp2RollingVisibilityFix();
    if (typeof app2.renderAllFromVisibleRange === 'function') app2.renderAllFromVisibleRange();
    resizeIframe(els.frameApp2);
    showBanner('App1 の再計算データを App2 に反映しました。', 'success');
    setShellStatus('App1 -> App2 synced');
  }

  function installFrameContent(frame, html, kind) {
    frame.srcdoc = preprocessFrameHtml(html);
    frame.addEventListener('load', function onLoad() {
      if (kind === 'app1') shellState.frameReady.app1 = true;
      if (kind === 'app2') shellState.frameReady.app2 = true;
      if (kind === 'app1') {
        installApp1TransferButton();
        ensureApp1DrawModeCheckboxes();
        applyApp1ChartDrawModeFilter();
      }
      resizeIframe(frame);
      setShellStatus(`App1 ${shellState.frameReady.app1 ? 'ready' : 'loading'} / App2 ${shellState.frameReady.app2 ? 'ready' : 'loading'}`);
      setTimeout(() => resizeIframe(frame), 300);
      setTimeout(() => resizeIframe(frame), 1200);
    }, { once: true });
  }

  function restoreShellLabels() {
    const tabButtons = Array.from(document.querySelectorAll('.tab-btn'));
    if (tabButtons[0]) tabButtons[0].textContent = 'Fund Search';
    if (tabButtons[1]) tabButtons[1].textContent = 'App1 xlsx.Builder';
    if (tabButtons[2]) tabButtons[2].textContent = 'App 2 / Fund Analytics';
    if (tabButtons[3]) tabButtons[3].textContent = 'App 3 / Rebalance';
    renderTabSyncMarks();

    const actionMap = {
      syncApp1ToApp2Btn: 'Send App 1 workbook to App 2',
      syncApp1ToPortfolioBtn: 'Send App 1 data to App 3',
      syncApp2ToPortfolioBtn: 'Send App 2 data to App 3',
      syncFundSearchAllBtn: 'Import Fund Search to App1/2/3',
      panelSyncFundSearchAllBtn: 'Fund Search -> App1/2/3',
      panelSyncApp1ToApp2Btn: 'App 1 -> App 2',
      panelSyncApp1ToPortfolioBtn: 'App 1 -> App 3',
      panelSyncApp2ToPortfolioBtn: 'App 2 -> App 3',
      resetFromApp2Btn: 'Reset from App 2',
      resetTargetsBtn: 'Reset to current weights',
      exportPortfolioCsvBtn: 'Export portfolio CSV',
      openApp1Btn: 'Back to App 1',
      openApp2Btn: 'Back to App 2'
    };
    Object.entries(actionMap).forEach(([id, text]) => {
      const el = byId(id);
      if (el) {
        el.textContent = text;
        el.setAttribute('data-base-label', text);
      }
    });

    const h2s = Array.from(document.querySelectorAll('.panel-head h2'));
    if (h2s[0]) h2s[0].textContent = 'Fund Search / Web';
    if (h2s[1]) h2s[1].textContent = 'App 1 / XLSX Builder';
    if (h2s[2]) h2s[2].textContent = 'App 2 / Fund Analysis';
    if (h2s[3]) h2s[3].textContent = 'App 3 / Portfolio Rebalance';

    const shellStatus = byId('shellStatusText');
    if (shellStatus) shellStatus.textContent = 'Loading...';
  }

  function reorderHeroButtons() {
    const bar = document.querySelector('.hero-actions');
    if (!bar) return;
    const importBtn = byId('syncFundSearchAllBtn');
    const app1To2Btn = byId('syncApp1ToApp2Btn');
    if (!importBtn || !app1To2Btn) return;
    bar.insertBefore(importBtn, app1To2Btn);
  }

  function cacheElements() {
    [
      'frameFundSearch', 'frameApp1', 'frameApp2', 'shellStatusText',
      'syncApp1ToApp2Btn', 'syncApp1ToPortfolioBtn', 'syncApp2ToPortfolioBtn', 'syncFundSearchAllBtn',
      'panelSyncFundSearchAllBtn', 'panelSyncApp1ToApp2Btn', 'panelSyncApp1ToPortfolioBtn', 'panelSyncApp2ToPortfolioBtn',
      'portfolioBanner', 'portfolioInvestModeBtn', 'portfolioInvestModeNote', 'portfolioExtraCash', 'portfolioCorrelation', 'portfolioRiskFree',
      'extraCashBreakdown',
      'goalTargetAmount', 'goalYears', 'goalYearsRange', 'goalQuickBtn', 'goalQuickPopover', 'goalQuickValue', 'goalProjectionSummary', 'goalProjectionChart', 'goalPlannerModeBtn', 'goalPlannerModeNote', 'goalTargetReturnLowPct', 'goalTargetReturnMidPct', 'goalTargetReturnHighPct', 'goalTargetReturnMidWrap', 'goalTargetReturnHighWrap', 'goalReturnWrap',
      'portfolioCsvInput', 'portfolioSourceText', 'portfolioSummaryGrid', 'portfolioDonut', 'donutCenter',
      'currentPortfolioDonut', 'currentPortfolioCenter', 'editableMixDonut', 'editableMixCenter', 'unitBarRows',
      'portfolioLegend', 'riskReturnScatter', 'riskReturnLegend', 'riskReturnSelected', 'drawdownBars',
      'tradePlanBody', 'portfolioTableBody', 'resetFromApp2Btn', 'resetTargetsBtn', 'exportPortfolioCsvBtn',
      'openApp1Btn', 'openApp2Btn'
    ].forEach(id => { els[id] = byId(id); });
  }

  function bindEvents() {
    const syncManualTargetReturnChoiceUi = () => {
      const active = shellState.manualTargetReturnChoice === 'high' ? 'high' : 'mid';
      if (els.goalTargetReturnMidWrap) {
        els.goalTargetReturnMidWrap.style.borderColor = active === 'mid' ? '#3b82f6' : '#d8e5f2';
        els.goalTargetReturnMidWrap.style.background = active === 'mid' ? '#eef6ff' : 'transparent';
        const note = els.goalTargetReturnMidWrap.querySelector('.inline-note');
        if (note) {
          note.textContent = active === 'mid' ? '期待値（選択中）' : '期待値';
          note.style.color = active === 'mid' ? '#1d4ed8' : '#5d7288';
          note.style.fontWeight = active === 'mid' ? '700' : '400';
        }
      }
      if (els.goalTargetReturnHighWrap) {
        els.goalTargetReturnHighWrap.style.borderColor = active === 'high' ? '#3b82f6' : '#d8e5f2';
        els.goalTargetReturnHighWrap.style.background = active === 'high' ? '#eef6ff' : 'transparent';
        const note = els.goalTargetReturnHighWrap.querySelector('.inline-note');
        if (note) {
          note.textContent = active === 'high' ? '最大値（選択中）' : '最大値';
          note.style.color = active === 'high' ? '#1d4ed8' : '#5d7288';
          note.style.fontWeight = active === 'high' ? '700' : '400';
        }
      }
    };
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => switchTab(btn.getAttribute('data-tab-target')));
    });

    const app1to2Buttons = [els.syncApp1ToApp2Btn, els.panelSyncApp1ToApp2Btn];
    app1to2Buttons.forEach(btn => {
      if (!btn) return;
      btn.addEventListener('click', async () => {
        toggleSendPrefix(app1to2Buttons, true);
        try {
          await syncApp1ToApp2();
          setSyncMark('app1', 'ok');
          setSyncMark('app2', 'ok');
        } catch (error) {
          setSyncMark('app1', 'ng');
          setSyncMark('app2', 'ng');
          showBanner(error.message, 'error');
        } finally {
          toggleSendPrefix(app1to2Buttons, false);
        }
      });
    });

    const app1to3Buttons = [els.syncApp1ToPortfolioBtn, els.panelSyncApp1ToPortfolioBtn];
    app1to3Buttons.forEach(btn => {
      if (!btn) return;
      btn.addEventListener('click', () => {
        toggleSendPrefix(app1to3Buttons, true);
        try {
          syncApp1ToPortfolio();
          setShellStatus('App1 -> App3 synced');
          setSyncMark('app1', 'ok');
          setSyncMark('app3', 'ok');
        } catch (error) {
          setSyncMark('app1', 'ng');
          setSyncMark('app3', 'ng');
          showBanner(error.message, 'error');
        } finally {
          toggleSendPrefix(app1to3Buttons, false);
        }
      });
    });

    const app2to3Buttons = [els.syncApp2ToPortfolioBtn, els.panelSyncApp2ToPortfolioBtn];
    app2to3Buttons.forEach(btn => {
      if (!btn) return;
      btn.addEventListener('click', async () => {
        toggleSendPrefix(app2to3Buttons, true);
        try {
          await syncApp2ToPortfolio();
          setShellStatus('App2 -> App3 synced');
          setSyncMark('app2', 'ok');
          setSyncMark('app3', 'ok');
        } catch (error) {
          setSyncMark('app2', 'ng');
          setSyncMark('app3', 'ng');
          showBanner(error.message, 'error');
        } finally {
          toggleSendPrefix(app2to3Buttons, false);
        }
      });
    });

    if (els.syncFundSearchAllBtn) {
      els.syncFundSearchAllBtn.addEventListener('click', async () => {
        toggleSendPrefix([els.syncFundSearchAllBtn, els.panelSyncFundSearchAllBtn], true);
        try { await syncFundSearchToAllApps(); } catch (error) {
          setSyncMark('app1', 'ng');
          setSyncMark('app2', 'ng');
          setSyncMark('app3', 'ng');
          showBanner(error.message, 'error');
        } finally {
          toggleSendPrefix([els.syncFundSearchAllBtn, els.panelSyncFundSearchAllBtn], false);
        }
      });
    }
    if (els.panelSyncFundSearchAllBtn) {
      els.panelSyncFundSearchAllBtn.addEventListener('click', async () => {
        toggleSendPrefix([els.syncFundSearchAllBtn, els.panelSyncFundSearchAllBtn], true);
        try { await syncFundSearchToAllApps(); } catch (error) {
          setSyncMark('app1', 'ng');
          setSyncMark('app2', 'ng');
          setSyncMark('app3', 'ng');
          showBanner(error.message, 'error');
        } finally {
          toggleSendPrefix([els.syncFundSearchAllBtn, els.panelSyncFundSearchAllBtn], false);
        }
      });
    }

    if (els.openApp1Btn) els.openApp1Btn.addEventListener('click', () => switchTab('panelApp1'));
    if (els.openApp2Btn) els.openApp2Btn.addEventListener('click', () => switchTab('panelApp2'));
    if (els.resetTargetsBtn) {
      els.resetTargetsBtn.addEventListener('click', () => {
        normalizePortfolio();
        shellState.portfolioItems.forEach(item => { item.targetWeight = item.currentWeight; });
        renderPortfolio();
      });
    }
    if (els.resetFromApp2Btn) {
      els.resetFromApp2Btn.addEventListener('click', async () => {
        try { await syncApp2ToPortfolio(); } catch (error) { showBanner(error.message, 'error'); }
      });
    }
    if (els.exportPortfolioCsvBtn) els.exportPortfolioCsvBtn.addEventListener('click', exportPortfolioCsv);

    [els.portfolioExtraCash, els.portfolioCorrelation, els.portfolioRiskFree, els.goalTargetAmount, els.goalYears, els.goalTargetReturnLowPct, els.goalTargetReturnMidPct, els.goalTargetReturnHighPct].forEach(input => {
      if (!input) return;
      input.addEventListener('input', renderPortfolio);
    });
    if (els.portfolioExtraCash) {
      els.portfolioExtraCash.addEventListener('input', () => {
        if (shellState.investmentMode === 'savings') shellState.savingsExtraCashEdited = true;
      });
      els.portfolioExtraCash.addEventListener('blur', () => {
        els.portfolioExtraCash.value = formatMoneyInputValue(els.portfolioExtraCash.value);
      });
      els.portfolioExtraCash.value = formatMoneyInputValue(els.portfolioExtraCash.value);
    }
    if (els.goalYears) {
      els.goalYears.addEventListener('input', () => {
        const y = Math.max(1, Math.round(num(els.goalYears.value, shellState.lumpYearsBase || 10)));
        if (els.goalYearsRange) els.goalYearsRange.value = String(y);
        if (shellState.investmentMode === 'savings') {
          shellState.savingsYearsEdited = true;
        } else {
          shellState.lumpYearsBase = y;
        }
      });
    }
    if (els.goalYearsRange) {
      els.goalYearsRange.addEventListener('input', () => {
        const y = Math.max(1, Math.round(num(els.goalYearsRange.value, 10)));
        if (els.goalYears) els.goalYears.value = String(y);
        if (shellState.investmentMode === 'savings') shellState.savingsYearsEdited = true;
        else shellState.lumpYearsBase = y;
        renderPortfolio();
      });
    }
    if (els.goalTargetAmount) {
      els.goalTargetAmount.addEventListener('blur', () => {
        els.goalTargetAmount.value = formatMoneyInputValue(els.goalTargetAmount.value);
        if (els.goalQuickValue) els.goalQuickValue.textContent = els.goalTargetAmount.value;
        renderPortfolio();
      });
      els.goalTargetAmount.value = formatMoneyInputValue(els.goalTargetAmount.value);
      if (els.goalQuickValue) els.goalQuickValue.textContent = els.goalTargetAmount.value;
    }
    if (els.goalTargetReturnLowPct && !String(els.goalTargetReturnLowPct.value || '').trim()) els.goalTargetReturnLowPct.value = '-3';
    if (els.goalTargetReturnMidPct && !String(els.goalTargetReturnMidPct.value || '').trim()) els.goalTargetReturnMidPct.value = '3';
    if (els.goalTargetReturnHighPct && !String(els.goalTargetReturnHighPct.value || '').trim()) els.goalTargetReturnHighPct.value = '6';
    if (els.goalYearsRange && els.goalYears) els.goalYearsRange.value = String(Math.max(1, Math.round(num(els.goalYears.value, 10))));
    if (els.goalReturnWrap) {
      els.goalReturnWrap.style.display = shellState.goalPlannerMode === 'manual' ? '' : 'none';
    }
    syncManualTargetReturnChoiceUi();
    if (els.portfolioInvestModeBtn) {
      els.portfolioInvestModeBtn.dataset.mode = shellState.investmentMode;
      els.portfolioInvestModeBtn.textContent = shellState.investmentMode === 'savings' ? '積立投資' : '一括投資';
    }
    if (els.portfolioInvestModeNote) {
      els.portfolioInvestModeNote.textContent = shellState.investmentMode === 'savings'
        ? '積立: Extra cash は最小額 + 最小額×√2（手入力で上書き可）'
        : '最小実現額を Extra cash に反映';
    }
    document.addEventListener('click', event => {
      const amt = event.target.closest('[data-goal-amount]');
      if (amt && els.goalTargetAmount) {
        els.goalTargetAmount.value = formatMoneyInputValue(amt.getAttribute('data-goal-amount'));
        if (els.goalQuickPopover) els.goalQuickPopover.style.display = 'none';
        if (els.goalQuickValue) els.goalQuickValue.textContent = els.goalTargetAmount.value;
        renderPortfolio();
        return;
      }
      const deltaBtn = event.target.closest('[data-goal-delta]');
      if (deltaBtn && els.goalTargetAmount) {
        const cur = num(els.goalTargetAmount.value, 0);
        const delta = num(deltaBtn.getAttribute('data-goal-delta'), 0);
        const next = Math.max(0, Math.round(cur + delta));
        els.goalTargetAmount.value = formatMoneyInputValue(next);
        if (els.goalQuickValue) els.goalQuickValue.textContent = els.goalTargetAmount.value;
        renderPortfolio();
        return;
      }
      const quickBtn = event.target.closest('#goalQuickBtn');
      if (quickBtn && els.goalQuickPopover) {
        if (els.goalQuickPopover.style.display === 'block') {
          els.goalQuickPopover.style.display = 'none';
        } else {
          const r = quickBtn.getBoundingClientRect();
          els.goalQuickPopover.style.top = `${Math.min(window.innerHeight - 120, r.bottom + 8)}px`;
          els.goalQuickPopover.style.left = `${Math.min(window.innerWidth - 220, r.left)}px`;
          els.goalQuickPopover.style.display = 'block';
          if (els.goalQuickValue && els.goalTargetAmount) els.goalQuickValue.textContent = formatMoneyInputValue(els.goalTargetAmount.value);
        }
        return;
      }
      if (els.goalQuickPopover && !event.target.closest('#goalQuickPopover')) {
        els.goalQuickPopover.style.display = 'none';
      }
      const modeBtn = event.target.closest('#goalPlannerModeBtn');
      if (modeBtn) {
        shellState.goalPlannerMode = shellState.goalPlannerMode === 'chart' ? 'manual' : 'chart';
        modeBtn.dataset.mode = shellState.goalPlannerMode;
        modeBtn.textContent = shellState.goalPlannerMode === 'chart' ? 'Chart Click Input' : 'Manual Input';
        const note = byId('goalPlannerModeNote');
        if (note) note.textContent = `グラフクリックで設定: ${shellState.goalPlannerMode === 'chart' ? 'ON' : 'OFF'}`;
        if (els.goalReturnWrap) els.goalReturnWrap.style.display = shellState.goalPlannerMode === 'manual' ? '' : 'none';
        if (shellState.goalPlannerMode === 'manual') {
          if (els.goalTargetReturnLowPct && !Number.isFinite(num(els.goalTargetReturnLowPct.value, NaN))) els.goalTargetReturnLowPct.value = '-3';
          if (els.goalTargetReturnMidPct && !Number.isFinite(num(els.goalTargetReturnMidPct.value, NaN))) els.goalTargetReturnMidPct.value = '3';
          if (els.goalTargetReturnHighPct && !Number.isFinite(num(els.goalTargetReturnHighPct.value, NaN))) els.goalTargetReturnHighPct.value = '6';
        }
        renderPortfolio();
        return;
      }
      const midWrap = event.target.closest('#goalTargetReturnMidWrap');
      if (midWrap) {
        shellState.manualTargetReturnChoice = 'mid';
        syncManualTargetReturnChoiceUi();
        renderPortfolio();
        return;
      }
      const highWrap = event.target.closest('#goalTargetReturnHighWrap');
      if (highWrap) {
        shellState.manualTargetReturnChoice = 'high';
        syncManualTargetReturnChoiceUi();
        renderPortfolio();
        return;
      }
      const investModeBtn = event.target.closest('#portfolioInvestModeBtn');
      if (investModeBtn) {
        shellState.investmentMode = shellState.investmentMode === 'savings' ? 'lump' : 'savings';
        investModeBtn.dataset.mode = shellState.investmentMode;
        investModeBtn.textContent = shellState.investmentMode === 'savings' ? '積立投資' : '一括投資';
        if (els.portfolioInvestModeNote) {
          els.portfolioInvestModeNote.textContent = shellState.investmentMode === 'savings'
            ? '積立: Extra cash は最小額 + 最小額×√2（手入力で上書き可）'
            : '最小実現額を Extra cash に反映';
        }
        if (shellState.investmentMode === 'savings') {
          shellState.savingsExtraCashEdited = false;
          shellState.savingsYearsEdited = false;
          const baseYears = Math.max(1, Math.round(num(els.goalYears && els.goalYears.value, shellState.lumpYearsBase || 10)));
          shellState.lumpYearsBase = baseYears;
          const savingsYears = Math.max(1, Math.round(baseYears * Math.SQRT2));
          if (els.goalYears) els.goalYears.value = String(savingsYears);
          if (els.goalYearsRange) els.goalYearsRange.value = String(savingsYears);
        } else {
          shellState.savingsExtraCashEdited = false;
          shellState.savingsYearsEdited = false;
          if (els.goalYears) els.goalYears.value = String(Math.max(1, Math.round(shellState.lumpYearsBase || num(els.goalYears.value, 10) || 10)));
          if (els.goalYearsRange && els.goalYears) els.goalYearsRange.value = String(Math.max(1, Math.round(num(els.goalYears.value, 10))));
        }
        renderPortfolio();
        return;
      }
      const toggleSeries = event.target.closest('[data-toggle-series]');
      if (toggleSeries) {
        const name = toggleSeries.getAttribute('data-toggle-series');
        if (!name) return;
        if (shellState.scatterHidden.has(name)) shellState.scatterHidden.delete(name);
        else shellState.scatterHidden.add(name);
        renderPortfolio();
        return;
      }
    });

    if (els.portfolioCsvInput) {
      els.portfolioCsvInput.addEventListener('change', async event => {
        try {
          await importPortfolioCsvFiles(event.target.files);
        } catch (error) {
          showBanner(error.message, 'error');
        } finally {
          event.target.value = '';
        }
      });
    }

    if (els.portfolioTableBody) {
      els.portfolioTableBody.addEventListener('input', handlePortfolioTableInput);
      els.portfolioTableBody.addEventListener('click', event => {
        const button = event.target.closest('[data-remove-item]');
        if (button) removePortfolioItem(button.getAttribute('data-remove-item'));
      });
    }
    if (els.portfolioLegend) {
      els.portfolioLegend.addEventListener('click', event => {
        const button = event.target.closest('[data-remove-item]');
        if (button) removePortfolioItem(button.getAttribute('data-remove-item'));
      });
    }

    if (els.unitBarRows) {
      els.unitBarRows.addEventListener('mousedown', startUnitBarDrag);
      els.unitBarRows.addEventListener('change', event => {
        const toggle = event.target.closest('#buyOnlyToggleInline');
        if (!toggle) return;
        shellState.buyOnlyMode = !!toggle.checked;
        renderPortfolio();
      });
      els.unitBarRows.addEventListener('click', event => {
        const btn = event.target.closest('[data-hide-item]');
        if (!btn) return;
        const id = btn.getAttribute('data-hide-item');
        if (!id) return;
        shellState.hiddenFunds.add(id);
        renderPortfolio();
      });
    }
    if (els.editableMixDonut) {
      els.editableMixDonut.addEventListener('mousedown', handleEditableMixMouseDown);
    }
    window.addEventListener('mousemove', handleGlobalPointerMove);
    window.addEventListener('mouseup', handleGlobalPointerUp);
  }

  function applyApp2RollingDefaults() {
    try {
      if (!els.frameApp2 || !els.frameApp2.contentWindow) return;
      const app2 = els.frameApp2.contentWindow;
      const s = safeGetAppState(app2);
      if (s && s.rollingChartMode !== 'bar') s.rollingChartMode = 'bar';
      if (typeof app2.updateRollingToggleButton === 'function') {
        try { app2.updateRollingToggleButton(); } catch (e) { /* no-op */ }
      }
      const doc = els.frameApp2.contentDocument;
      const btn = doc && doc.getElementById('toggleRollingModeBtn');
      if (btn) {
        const text = String(btn.textContent || '').trim();
        // App2本体の状態と表示がズレるケースがあるため、必要時は実クリックで確実にbarへ寄せる
        if (text.includes('棒グラフへ切替')) {
          try { btn.click(); } catch (e) { /* no-op */ }
        }
        btn.textContent = '折れ線へ切替';
      }
      if (!shellState.__forcedRollingBarOnce && typeof app2.renderAllFromVisibleRange === 'function') {
        shellState.__forcedRollingBarOnce = true;
        try { app2.renderAllFromVisibleRange(); } catch (e) { /* no-op */ }
      }
    } catch (e) {
      // no-op
    }
  }

  function applyApp2PanelLayoutFix() {
    try {
      if (!els.frameApp2 || !els.frameApp2.contentDocument) return;
      const doc = els.frameApp2.contentDocument;
      const left = doc.getElementById('leftPanel') || doc.querySelector('.left-panel');
      const settings = doc.getElementById('settingsSection');
      const cards = Array.from(doc.querySelectorAll('.card, .section, .panel, .sticky-main-card'));
      const settingsCard = cards.find(card => {
        const t = (card.querySelector('.section-title, .card-title-text, h2, h3') || {}).textContent || '';
        return String(t).indexOf('入力・設定') >= 0;
      });
      const applyBox = el => {
        if (!el || !el.style) return;
        el.style.boxSizing = 'border-box';
        el.style.width = '100%';
        el.style.maxWidth = '100%';
        el.style.overflow = 'hidden';
      };
      applyBox(left);
      if (left && window.innerWidth <= 900) {
        left.style.flexBasis = '100%';
        left.style.resize = 'none';
      }
      applyBox(settings);
      applyBox(settingsCard);
      const body = settingsCard && settingsCard.querySelector('.body, .section-body, .card-body');
      applyBox(body);
      const controls = doc.querySelectorAll('#settingsSection input, #settingsSection select, #settingsSection textarea, #settingsSection button');
      controls.forEach(el => {
        if (!el || !el.style) return;
        el.style.boxSizing = 'border-box';
        el.style.maxWidth = '100%';
      });
    } catch (e) {
      // no-op
    }
  }

  function applyApp2RollingVisibilityFix() {
    try {
      if (!els.frameApp2 || !els.frameApp2.contentWindow) return;
      const app2 = els.frameApp2.contentWindow;
      if (app2.__codexRollingVisibilityPatched || typeof app2.renderAdditionalFundSections !== 'function') return;
      const fillRollingExcess = rows => {
        const list = Array.isArray(rows) ? rows : [];
        const win = Math.max(2, num((els.frameApp2 && els.frameApp2.contentDocument && els.frameApp2.contentDocument.getElementById('rollingWindow') && els.frameApp2.contentDocument.getElementById('rollingWindow').value), 60));
        // Normalize numeric fields and prepare bench/fund level series.
        const fundLevels = [];
        const benchLevels = [];
        for (let i = 0; i < list.length; i += 1) {
          const r = list[i];
          if (!r || typeof r !== 'object') continue;
          const fAdj = num(r.fundAdj, num(r.fundNav, NaN));
          const bPx = num(r.benchPrice, num(r.benchNav, NaN));
          if (Number.isFinite(fAdj)) r.fundAdj = fAdj;
          if (Number.isFinite(bPx)) r.benchPrice = bPx;
          fundLevels[i] = num(r.fundAdj, NaN);
          benchLevels[i] = num(r.benchPrice, NaN);
          const rr = num(r.rollingExcess, NaN);
          if (Number.isFinite(rr)) {
            r.rollingExcess = rr;
            continue;
          }
          const er = num(r.excessReturn, NaN);
          if (Number.isFinite(er)) {
            r.rollingExcess = er;
            continue;
          }
        }
        // Forward/back fill benchmark levels when sparse.
        let prevBench = NaN;
        for (let i = 0; i < benchLevels.length; i += 1) {
          if (Number.isFinite(benchLevels[i])) prevBench = benchLevels[i];
          else if (Number.isFinite(prevBench)) benchLevels[i] = prevBench;
        }
        let nextBench = NaN;
        for (let i = benchLevels.length - 1; i >= 0; i -= 1) {
          if (Number.isFinite(benchLevels[i])) nextBench = benchLevels[i];
          else if (Number.isFinite(nextBench)) benchLevels[i] = nextBench;
        }
        // Recompute rollingExcess from level series if still missing.
        for (let i = 0; i < list.length; i += 1) {
          const r = list[i];
          if (!r || typeof r !== 'object') continue;
          if (Number.isFinite(num(r.rollingExcess, NaN))) continue;
          if (i < win - 1) continue;
          const sf = fundLevels[i - (win - 1)];
          const ef = fundLevels[i];
          const sb = benchLevels[i - (win - 1)];
          const eb = benchLevels[i];
          if (!Number.isFinite(sf) || !Number.isFinite(ef) || !Number.isFinite(sb) || !Number.isFinite(eb) || sf === 0 || sb === 0) continue;
          const fund = (ef / sf) - 1;
          const bench = (eb / sb) - 1;
          r.rollingExcess = (fund - bench);
        }
        return list;
      };
      if (!app2.__codexRollingStatsPatched) {
        if (typeof app2.renderRollingStats === 'function') {
          const origStats = app2.renderRollingStats.bind(app2);
          app2.renderRollingStats = function patchedRenderRollingStats(rows) {
            return origStats(fillRollingExcess(Array.isArray(rows) ? rows : []));
          };
        }
        if (typeof app2.renderRollingStatsInto === 'function') {
          const origInto = app2.renderRollingStatsInto.bind(app2);
          app2.renderRollingStatsInto = function patchedRenderRollingStatsInto(rows, targetEl) {
            return origInto(fillRollingExcess(Array.isArray(rows) ? rows : []), targetEl);
          };
        }
        app2.__codexRollingStatsPatched = true;
      }
      const orig = app2.renderAdditionalFundSections;
      app2.renderAdditionalFundSections = function patchedRenderAdditionalFundSections() {
        try {
          const s0 = safeGetAppState(app2);
          if (s0 && typeof s0 === 'object') {
            fillRollingExcess(s0.merged);
            if (Array.isArray(s0.additionalFundSeries)) {
              s0.additionalFundSeries.forEach(x => fillRollingExcess(x && x.visibleRows));
            }
          }
        } catch (e) {
          // no-op
        }
        const out = orig.apply(this, arguments);
        try {
          const s = safeGetAppState(app2);
          if (s && typeof s === 'object') {
            const merged = Array.isArray(s.merged) ? s.merged : [];
            const benchFromMerged = merged
              .map(r => ({
                date: normalizeDateKey(r && (r.date || r.dateISO)),
                price: num(r && (r.benchPrice ?? r.benchmarkPrice ?? r.benchNav ?? r.benchIndex ?? r.benchmarkIndex), NaN)
              }))
              .filter(x => x.date && Number.isFinite(x.price));
            if ((!Array.isArray(s.benchmarkRows) || !s.benchmarkRows.length) && benchFromMerged.length) {
              s.benchmarkRows = benchFromMerged;
            }
            if ((!s.bm || !Array.isArray(s.bm.rows) || !s.bm.rows.length) && Array.isArray(s.benchmarkRows) && s.benchmarkRows.length) {
              s.bm = s.bm && typeof s.bm === 'object' ? s.bm : {};
              s.bm.rows = s.benchmarkRows.map(r => ({
                dateISO: String(r.date || ''),
                date: new Date(String(r.date || '')),
                nav: num(r.price, NaN)
              })).filter(x => x.dateISO && Number.isFinite(x.nav) && !Number.isNaN(x.date.getTime()));
            }
            const hasBenchmarkLikeRows = rows => (Array.isArray(rows) ? rows : []).some(r => (
              Number.isFinite(num(r && r.benchPrice, NaN)) ||
              Number.isFinite(num(r && r.benchReturn, NaN)) ||
              Number.isFinite(num(r && r.benchIndex, NaN)) ||
              Number.isFinite(num(r && r.excessReturn, NaN)) ||
              Number.isFinite(num(r && r.rollingExcess, NaN))
            ));
            const hasBenchmarkLike = hasBenchmarkLikeRows(s.merged) ||
              (Array.isArray(s.additionalFundSeries) && s.additionalFundSeries.some(x => hasBenchmarkLikeRows(x && x.visibleRows)));
            if ((Array.isArray(s.benchmarkRows) && s.benchmarkRows.length) || (s.bm && Array.isArray(s.bm.rows) && s.bm.rows.length) || hasBenchmarkLike) {
              s.hasBenchmark = true;
            }
            // Fallback: if rolling chart is not rendered despite table/state values,
            // redraw every rollingChart* panel directly from series rows.
            const doc = els.frameApp2 && els.frameApp2.contentDocument;
            const drawRollingFallback = (gd, rows, titleText) => {
              if (!gd || !Array.isArray(rows) || rows.length < 2 || !app2.Plotly || typeof app2.Plotly.newPlot !== 'function') return;
              const xs = [];
              const ys = [];
              for (let i = 0; i < rows.length; i += 1) {
                const row = rows[i];
                const d = normalizeDateKey(row && row.date);
                if (!d) continue;
                xs.push(d);
                if (row && typeof row === 'object' && !Number.isFinite(num(row.rollingExcess, NaN))) {
                  fillRollingExcess(rows);
                }
                const rr = num(row && row.rollingExcess, NaN);
                if (Number.isFinite(rr)) {
                  ys.push(rr * 100);
                  continue;
                }
                const er = num(row && row.excessReturn, NaN);
                ys.push(Number.isFinite(er) ? (er * 100) : NaN);
              }
              if (!ys.some(v => Number.isFinite(v))) return;
              const mode = (s.rollingChartMode === 'line') ? 'line' : 'bar';
              const trace = mode === 'bar'
                ? {
                    type: 'bar',
                    x: xs,
                    y: ys,
                    name: 'Rolling超過',
                    marker: { color: ys.map(v => (Number.isFinite(v) ? (v >= 0 ? '#16a34a' : '#ef4444') : '#94a3b8')) }
                  }
                : {
                    type: 'scatter',
                    mode: 'lines',
                    x: xs,
                    y: ys,
                    name: 'Rolling超過',
                    line: { color: '#ef4444', width: 1.8 }
                  };
              const layout = {
                title: { text: titleText || 'III. Rolling超過リターン', x: 0.5, y: 0.98 },
                margin: { l: 60, r: 24, t: 64, b: 120 },
                xaxis: { title: { text: '日付', standoff: 20 }, automargin: false },
                yaxis: { title: '超過（%）', automargin: false },
                legend: { orientation: 'h', x: 0, y: -0.24, xanchor: 'left', yanchor: 'top' },
                paper_bgcolor: '#ffffff',
                plot_bgcolor: '#ffffff'
              };
              const config = { responsive: false, displayModeBar: true };
              try { app2.Plotly.newPlot(gd, [trace], layout, config); } catch (e) { /* no-op */ }
            };
            const rollingTargets = doc ? Array.from(doc.querySelectorAll('[id^="rollingChart"]')) : [];
            const mergedRows = Array.isArray(s.merged) ? s.merged : [];
            const addSeries = Array.isArray(s.additionalFundSeries) ? s.additionalFundSeries : [];
            rollingTargets.forEach((gd, idx) => {
              const hasGraphData = !!(gd && Array.isArray(gd.data) && gd.data.some(tr => Array.isArray(tr && tr.y) && tr.y.some(v => Number.isFinite(Number(v)))));
              const isFlatGraph = !!(gd && Array.isArray(gd.data) && gd.data.some(tr => {
                if (!Array.isArray(tr && tr.y)) return false;
                const vals = tr.y.map(v => Number(v)).filter(Number.isFinite);
                if (!vals.length) return false;
                const mn = Math.min.apply(null, vals);
                const mx = Math.max.apply(null, vals);
                return Math.abs(mx - mn) < 1e-10;
              }));
              if (hasGraphData && !isFlatGraph) return;
              const rows = idx === 0
                ? (mergedRows.length ? mergedRows : (addSeries[0] && Array.isArray(addSeries[0].visibleRows) ? addSeries[0].visibleRows : []))
                : (addSeries[idx - 1] && Array.isArray(addSeries[idx - 1].visibleRows) ? addSeries[idx - 1].visibleRows : []);
              const card = gd.closest('.card, .section, .panel');
              const titleEl = card && card.querySelector('.section-title, .card-title-text, h2, h3');
              drawRollingFallback(gd, rows, titleEl ? String(titleEl.textContent || '').trim() : 'III. Rolling超過リターン');
            });

            // Ensure main rolling stats table reflects backfilled rollingExcess values.
            if (typeof app2.renderRollingStats === 'function') {
              const baseRows = Array.isArray(s.merged) ? s.merged : [];
              const visible = (typeof app2.recomputeVisibleRows === 'function')
                ? app2.recomputeVisibleRows(baseRows)
                : baseRows;
              try { app2.renderRollingStats(visible); } catch (e) { /* no-op */ }
            }

            // Fallback for IV chart: if table has data but chart is blank, draw grouped bars from table.
            const hp = doc && doc.getElementById('holdingPerfChart');
            const hpHasData = !!(hp && Array.isArray(hp.data) && hp.data.some(tr => Array.isArray(tr && tr.y) && tr.y.some(v => Number.isFinite(Number(v)))));
            const hpTableRows = doc ? Array.from(doc.querySelectorAll('#holdingPerfCard table tbody tr')) : [];
            if ((!hpHasData) && hp && hpTableRows.length && app2.Plotly && typeof app2.Plotly.newPlot === 'function') {
              const labels = [];
              const fundAvg = [];
              const benchAvg = [];
              hpTableRows.forEach(tr => {
                const tds = Array.from(tr.querySelectorAll('td')).map(td => String(td.textContent || '').trim());
                if (tds.length < 7) return;
                const term = tds[0];
                const fAvg = num(tds[3], NaN);
                const bAvg = num(tds[6], NaN);
                if (!term || (!Number.isFinite(fAvg) && !Number.isFinite(bAvg))) return;
                labels.push(term);
                fundAvg.push(Number.isFinite(fAvg) ? fAvg : NaN);
                benchAvg.push(Number.isFinite(bAvg) ? bAvg : NaN);
              });
              if (labels.length) {
                const data = [
                  { type: 'bar', name: 'ファンド平均', x: labels, y: fundAvg, marker: { color: '#22c55e' } },
                  { type: 'bar', name: 'ベンチ平均', x: labels, y: benchAvg, marker: { color: '#a78bfa' } }
                ];
                const layout = {
                  barmode: 'group',
                  title: { text: 'IV. 保有期間別パフォーマンス', x: 0.5, y: 0.98 },
                  margin: { l: 60, r: 24, t: 64, b: 140 },
                  xaxis: { title: { text: '保有期間', standoff: 20 }, automargin: false },
                  yaxis: { title: 'リターン (%)', automargin: false },
                  legend: { orientation: 'h', x: 0, y: -0.24, xanchor: 'left', yanchor: 'top' },
                  paper_bgcolor: '#ffffff',
                  plot_bgcolor: '#ffffff'
                };
                const config = { responsive: false, displayModeBar: true };
                try { app2.Plotly.newPlot(hp, data, layout, config); } catch (e) { /* no-op */ }
              }
            }
          }
        } catch (e) {
          // no-op
        }
        return out;
      };
      app2.__codexRollingVisibilityPatched = true;
      try { app2.renderAdditionalFundSections(); } catch (e) { /* no-op */ }
    } catch (e) {
      // no-op
    }
  }

  function applyApp2AlertGuard() {
    try {
      if (!els.frameApp2 || !els.frameApp2.contentWindow) return;
      const app2 = els.frameApp2.contentWindow;
      if (app2.__codexAlertGuardPatched) return;
      const originalAlert = typeof app2.alert === 'function' ? app2.alert.bind(app2) : null;
      app2.alert = function guardedAlert(message) {
        const text = String(message == null ? '' : message);
        if (text.includes('ファンドCSVまたは積立・スポット投資実績XLSXを選択してください。')) {
          return;
        }
        if (originalAlert) return originalAlert(message);
      };
      app2.__codexAlertGuardPatched = true;
    } catch (e) {
      // no-op
    }
  }

  function applyApp1ResultBase100Fix() {
    try {
      if (!els.frameApp1 || !els.frameApp1.contentWindow) return;
      const app1 = els.frameApp1.contentWindow;
      const plotly = app1.Plotly;
      if (!plotly || typeof plotly.react !== 'function' || app1.__codexApp1Base100Patched) return;
      const origReact = plotly.react.bind(plotly);
      const origNewPlot = (typeof plotly.newPlot === 'function') ? plotly.newPlot.bind(plotly) : null;
      const normalizeTraceName = name => String(name || '').replace(/\s+/g, '').toLowerCase();
      const getDrawMode = () => {
        const doc = els.frameApp1 && els.frameApp1.contentDocument;
        return {
          sumi: !!(doc && doc.getElementById('codexDrawTsumitate') && doc.getElementById('codexDrawTsumitate').checked),
          lump: !!(doc && doc.getElementById('codexDrawLump') && doc.getElementById('codexDrawLump').checked),
          same: !!(doc && doc.getElementById('codexDrawSameLump') && doc.getElementById('codexDrawSameLump').checked)
        };
      };
      const shouldKeepApp1Trace = name => {
        const { sumi, lump, same } = getDrawMode();
        if (!sumi && !lump && !same) return true;
        const n = normalizeTraceName(name);
        if (/bm|benchmark|ベンチ/.test(n)) return true;
        const label = String(name || '');
        const isSame = /同額一括|積立総額→同額一括|積立同額一括/.test(label);
        const isLump = /任意一括/.test(label);
        const isSumi = /積立/.test(label) && !isSame;
        if (isSame) return same;
        if (isLump) return lump;
        if (isSumi) return sumi;
        return true;
      };
      const enforceVisibilityOnGraph = gd => {
        if (!gd || !Array.isArray(gd.data) || typeof plotly.restyle !== 'function') return;
        const vis = gd.data.map(tr => (shouldKeepApp1Trace(tr && tr.name) ? true : 'legendonly'));
        try { plotly.restyle(gd, { visible: vis }); } catch (e) { /* no-op */ }
      };
      app1.__codexEnforceDrawModeVisibility = function enforceDrawModeVisibility() {
        try {
          const doc = els.frameApp1 && els.frameApp1.contentDocument;
          if (!doc) return;
          const targets = Array.from(doc.querySelectorAll('.js-plotly-plot'));
          targets.forEach(gd => enforceVisibilityOnGraph(gd));
        } catch (e) { /* no-op */ }
      };
      const normalizeAndFilterApp1Data = (gd, data, layout) => {
        const id = typeof gd === 'string' ? gd : (gd && gd.id ? gd.id : '');
        const titleText = String(
          (layout && layout.title && (layout.title.text || layout.title)) ||
          ''
        );
        const shouldNormalize =
          /評価額推移|基準価額推移/.test(titleText) ||
          /valuation|nav|result/i.test(String(id || ''));
        if (shouldNormalize && Array.isArray(data)) {
          data = data.filter(trace => shouldKeepApp1Trace(trace && trace.name));
          data.forEach(trace => {
            if (!trace || !Array.isArray(trace.y)) return;
            const firstFinite = trace.y.find(v => Number.isFinite(Number(v)));
            const base = Number(firstFinite);
            if (!Number.isFinite(base) || base === 0) return;
            trace.y = trace.y.map(v => {
              const n = Number(v);
              return Number.isFinite(n) ? (n / base) * 100 : v;
            });
          });
          layout = Object.assign({}, layout || {}, {
            yaxis: Object.assign({}, (layout && layout.yaxis) || {}, { title: '指数化値（起点 100）' })
          });
        }
        return { data, layout };
      };
      plotly.react = function patchedReact(gd, data, layout, config) {
        try {
          const out = normalizeAndFilterApp1Data(gd, data, layout);
          data = out.data;
          layout = out.layout;
        } catch (e) {
          // no-op
        }
        const out = origReact(gd, data, layout, config);
        try {
          if (out && typeof out.then === 'function') out.then(() => app1.__codexEnforceDrawModeVisibility());
          else app1.__codexEnforceDrawModeVisibility();
        } catch (e) { /* no-op */ }
        return out;
      };
      if (origNewPlot) {
        plotly.newPlot = function patchedNewPlot(gd, data, layout, config) {
          try {
            const out = normalizeAndFilterApp1Data(gd, data, layout);
            data = out.data;
            layout = out.layout;
          } catch (e) {
            // no-op
          }
          const out = origNewPlot(gd, data, layout, config);
          try {
            if (out && typeof out.then === 'function') out.then(() => app1.__codexEnforceDrawModeVisibility());
            else app1.__codexEnforceDrawModeVisibility();
          } catch (e) { /* no-op */ }
          return out;
        };
      }
      app1.__codexApp1Base100Patched = true;
      ensureApp1DrawModeCheckboxes();
      applyApp1ChartDrawModeFilter();
      try { app1.__codexEnforceDrawModeVisibility(); } catch (e) { /* no-op */ }
      if (typeof app1.renderAllFromVisibleRange === 'function') {
        try { app1.renderAllFromVisibleRange(); } catch (e) { /* no-op */ }
      }
    } catch (e) {
      // no-op
    }
  }

  function applyApp2ResultBase100Fix() {
    try {
      if (!els.frameApp2 || !els.frameApp2.contentWindow) return;
      const app2 = els.frameApp2.contentWindow;
      const plotly = app2.Plotly;
      if (!plotly || app2.__codexApp2Base100Patched) return;

      const normalizeTracesToBase100 = (gd, data, layout) => {
        try {
          const id = typeof gd === 'string' ? gd : (gd && gd.id ? gd.id : '');
          const titleText = String(
            (layout && layout.title && (layout.title.text || layout.title)) || ''
          );
          const isValuationChart = /評価額推移/.test(titleText);
          const isNavChart = /基準価額推移/.test(titleText);
          const shouldNormalize = isValuationChart || isNavChart;
          if (shouldNormalize && Array.isArray(data)) {
            let startDateKey = '';
            try {
              const doc = els.frameApp2 && els.frameApp2.contentDocument;
              const startInput = doc && doc.getElementById('startDate');
              startDateKey = normalizeDateKey(startInput && startInput.value);
            } catch (e) { /* no-op */ }
            data.forEach(trace => {
              if (!trace || !Array.isArray(trace.y) || !Array.isArray(trace.x)) return;
              let base = NaN;
              for (let i = 0; i < trace.y.length; i += 1) {
                const xv = normalizeDateKey(trace.x[i]);
                const yv = Number(trace.y[i]);
                if (!Number.isFinite(yv)) continue;
                if (startDateKey && xv && xv < startDateKey) continue;
                base = yv;
                break;
              }
              if (!Number.isFinite(base)) {
                const firstFinite = trace.y.find(v => Number.isFinite(Number(v)));
                base = Number(firstFinite);
              }
              if (!Number.isFinite(base) || base === 0) return;
              trace.y = trace.y.map(v => {
                const n = Number(v);
                return Number.isFinite(n) ? (n / base) * 100 : v;
              });
            });
            layout = Object.assign({}, layout || {}, {
              yaxis: Object.assign({}, (layout && layout.yaxis) || {}, {
                title: isValuationChart
                  ? '指数化値（投資日起点 100）'
                  : '指数化NAV（投資日起点 100）'
              })
            });
          }
        } catch (e) {
          // no-op
        }
        return { data, layout };
      };

      const origReact = typeof plotly.react === 'function' ? plotly.react.bind(plotly) : null;
      if (origReact) {
        plotly.react = function patchedReact(gd, data, layout, config) {
          const normalized = normalizeTracesToBase100(gd, data, layout);
          return origReact(gd, normalized.data, normalized.layout, config);
        };
      }

      const origNewPlot = typeof plotly.newPlot === 'function' ? plotly.newPlot.bind(plotly) : null;
      if (origNewPlot) {
        plotly.newPlot = function patchedNewPlot(gd, data, layout, config) {
          const normalized = normalizeTracesToBase100(gd, data, layout);
          return origNewPlot(gd, normalized.data, normalized.layout, config);
        };
      }

      const origRedraw = typeof plotly.redraw === 'function' ? plotly.redraw.bind(plotly) : null;
      if (origRedraw) {
        plotly.redraw = function patchedRedraw(gd) {
          try {
            if (gd && gd.data && gd.layout) normalizeTracesToBase100(gd, gd.data, gd.layout);
          } catch (e) { /* no-op */ }
          return origRedraw(gd);
        };
      }

      app2.__codexApp2Base100Patched = true;
      if (typeof app2.renderAllFromVisibleRange === 'function') {
        try { app2.renderAllFromVisibleRange(); } catch (e) { /* no-op */ }
      }
    } catch (e) {
      // no-op
    }
  }

  function applyApp2LegendLayoutFix() {
    try {
      if (!els.frameApp2 || !els.frameApp2.contentWindow) return;
      const app2 = els.frameApp2.contentWindow;
      const plotly = app2.Plotly;
      if (!plotly || typeof plotly.react !== 'function' || app2.__codexLegendLayoutPatched) return;

      const origReact = plotly.react.bind(plotly);
      plotly.react = function patchedReact(gd, data, layout, config) {
        try {
          const id = typeof gd === 'string' ? gd : (gd && gd.id ? gd.id : '');
          const targets = new Set([
            'indexChart',
            'drawdownChart',
            'rollingChart',
            'holdingPerfChart',
            'holdingPeriodChart',
            'workbookRiskReturnChart',
            'workbookScatterChart',
            'fundContributionPie',
            'lumpContributionPie',
            'sipContributionPie'
          ]);
          if (targets.has(id) && layout && typeof layout === 'object') {
            const denseChart = (id === 'indexChart' || id === 'drawdownChart' || id === 'holdingPerfChart');
            const pieLike = (id === 'fundContributionPie' || id === 'lumpContributionPie' || id === 'sipContributionPie');
            const rollingOnly = (id === 'rollingChart');
            layout.legend = Object.assign({}, layout.legend || {}, {
              orientation: 'h',
              x: 0,
              y: rollingOnly ? -0.26 : (denseChart ? -0.52 : (pieLike ? -0.2 : -0.30)),
              xanchor: 'left',
              yanchor: 'top',
              bgcolor: 'rgba(255,255,255,0.92)',
              bordercolor: '#d8e5f2',
              borderwidth: 1,
              font: Object.assign({}, (layout.legend && layout.legend.font) || {}, { size: denseChart ? 11 : 12 }),
              entrywidthmode: 'pixels',
              entrywidth: rollingOnly ? 300 : (denseChart ? 420 : 360),
              itemwidth: rollingOnly ? 300 : (denseChart ? 420 : 360),
              traceorder: 'normal',
              valign: 'top',
            });
            layout.margin = Object.assign({}, layout.margin || {}, {
              r: Math.max(24, Number((layout.margin && layout.margin.r) || 0)),
              l: Math.max(pieLike ? 96 : 56, Number((layout.margin && layout.margin.l) || 0)),
              t: Math.max(pieLike ? 104 : 72, Number((layout.margin && layout.margin.t) || 0)),
              b: Math.max(rollingOnly ? 150 : (denseChart ? 280 : (pieLike ? 170 : 200)), Number((layout.margin && layout.margin.b) || 0)),
            });
            layout.xaxis = Object.assign({}, layout.xaxis || {}, {
              automargin: false,
              title: Object.assign({}, (layout.xaxis && layout.xaxis.title) || {}, { standoff: 26 })
            });
            layout.yaxis = Object.assign({}, layout.yaxis || {}, { automargin: false });
            layout.xaxis2 = Object.assign({}, layout.xaxis2 || {}, { automargin: false });
            layout.yaxis2 = Object.assign({}, layout.yaxis2 || {}, { automargin: false });
            if (rollingOnly && Array.isArray(data)) {
              const hasRolling = data.some(tr => Array.isArray(tr && tr.y) && tr.y.some(v => Number.isFinite(Number(v))));
              if (!hasRolling) {
                const st = safeGetAppState(app2);
                const merged = st && Array.isArray(st.merged) ? st.merged : [];
                const xs = merged.length
                  ? merged.map(r => r.date)
                  : (Array.isArray(data[0] && data[0].x) ? data[0].x : []);
                const ys = merged.length
                  ? merged.map((_, i, arr) => {
                    if (i < 60) return NaN;
                    let f = 1; let b = 1; let ok = true;
                    for (let j = i - 59; j <= i; j += 1) {
                      const fr = Number(arr[j] && arr[j].fundReturn);
                      const br = Number(arr[j] && arr[j].benchReturn);
                      if (!Number.isFinite(fr) || !Number.isFinite(br)) { ok = false; break; }
                      f *= (1 + fr); b *= (1 + br);
                    }
                    return ok ? ((f - 1) - (b - 1)) * 100 : NaN;
                  })
                  : xs.map(() => 0);
                data.length = 0;
                data.push({
                  type: 'scatter',
                  mode: 'lines',
                  name: 'Rolling超過',
                  x: xs,
                  y: ys,
                  line: { color: '#ef4444', width: 1.5 }
                });
              }
            }
            if (config && typeof config === 'object') {
              config.responsive = false;
              config.displayModeBar = config.displayModeBar === false ? false : true;
            }
          }
        } catch (e) {
          // no-op
        }
        return origReact(gd, data, layout, config);
      };

      app2.__codexLegendLayoutPatched = true;
      if (typeof app2.renderAllFromVisibleRange === 'function') {
        try { app2.renderAllFromVisibleRange(); } catch (e) { /* no-op */ }
      }
    } catch (e) {
      // no-op
    }
  }

  function init() {
    installParentFileGuard();
    cacheElements();
    restoreShellLabels();
    reorderHeroButtons();
    bindEvents();
    const app1Base64 = typeof APP1_HTML_BASE64 !== 'undefined' ? APP1_HTML_BASE64 : window.APP1_HTML_BASE64;
    const app2Base64 = typeof APP2_HTML_BASE64 !== 'undefined' ? APP2_HTML_BASE64 : window.APP2_HTML_BASE64;
    installFrameContent(els.frameApp1, decodeBase64Utf8(app1Base64), 'app1');
    installFrameContent(els.frameApp2, decodeBase64Utf8(app2Base64), 'app2');
    setTimeout(applyApp2RollingDefaults, 300);
    setTimeout(applyApp2RollingDefaults, 1200);
    setTimeout(applyApp1ChartDrawModeFilter, 280);
    setTimeout(applyApp1ChartDrawModeFilter, 1100);
    setTimeout(applyApp1ResultBase100Fix, 360);
    setTimeout(applyApp1ResultBase100Fix, 1300);
    setTimeout(applyApp2ResultBase100Fix, 420);
    setTimeout(applyApp2ResultBase100Fix, 1500);
    setTimeout(applyApp2PanelLayoutFix, 320);
    setTimeout(applyApp2PanelLayoutFix, 1400);
    setTimeout(applyApp2AlertGuard, 340);
    setTimeout(applyApp2AlertGuard, 1500);
    setTimeout(applyApp2RollingVisibilityFix, 380);
    setTimeout(applyApp2RollingVisibilityFix, 1500);
    setTimeout(applyApp2LegendLayoutFix, 350);
    setTimeout(applyApp2LegendLayoutFix, 1400);
    if (els.frameFundSearch) {
      if (shellState.localMode) {
        els.frameFundSearch.src = 'http://127.0.0.1:8000/project/web/index.html?v=20260407-10';
      }
      els.frameFundSearch.addEventListener('load', () => {
        resizeIframe(els.frameFundSearch);
      });
    }
    window.__codexImportApiFundsToApp1 = importApiFundsToApp1;
    setTimeout(applyApp2RollingDefaults, 2200);
    setTimeout(applyApp1ChartDrawModeFilter, 2000);
    setTimeout(applyApp1ResultBase100Fix, 2100);
    setTimeout(applyApp2ResultBase100Fix, 2300);
    setTimeout(applyApp2PanelLayoutFix, 2300);
    setTimeout(applyApp2AlertGuard, 2300);
    setTimeout(applyApp2RollingVisibilityFix, 2300);
    setTimeout(applyApp2LegendLayoutFix, 2400);
    renderPortfolio();
    switchTab('panelFundSearch');
  }

  document.addEventListener('DOMContentLoaded', init);
})();








