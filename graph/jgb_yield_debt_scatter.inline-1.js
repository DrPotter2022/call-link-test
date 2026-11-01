(function () {
  // ---------- data ----------
  // debt = general government gross debt, % of GDP (approx, 2024)
  const data10 = [
    { code: "AU", name: "オーストラリア", debt: 51,    yield: 4.99 },
    { code: "NZ", name: "ニュージーランド", debt: 53,    yield: 4.71 },
    { code: "GB", name: "英国",           debt: 103,   yield: 4.99 },
    { code: "US", name: "米国",           debt: 125,   yield: 4.70 },
    { code: "FR", name: "フランス",        debt: 113,   yield: 3.98 },
    { code: "SI", name: "スロベニア",      debt: 67,    yield: 3.52 },
    { code: "BE", name: "ベルギー",        debt: 104.7, yield: 3.70 },
    { code: "FI", name: "フィンランド",     debt: 82.1,  yield: 3.47 },
    { code: "DE", name: "ドイツ",          debt: 62.5,  yield: 3.15 },
    { code: "AT", name: "オーストリア",     debt: 81.8,  yield: 3.41 },
    { code: "PT", name: "ポルトガル",       debt: 94.9,  yield: 3.50 },
    { code: "CA", name: "カナダ",          debt: 114,   yield: 3.68 },
    { code: "NL", name: "オランダ",         debt: 43.3,  yield: 3.25 },
    { code: "DK", name: "デンマーク",       debt: 31.1,  yield: 2.31 },
    { code: "CH", name: "スイス",          debt: 37,    yield: 0.38 },
  ];
  const jp10 = { code: "JP", name: "日本", debt: 230, yield: 2.87 };

  const data20 = [
    { code: "US", name: "米国",     debt: 125, yield: 5.25 },
    { code: "GB", name: "英国",     debt: 103, yield: 5.64 },
    { code: "FR", name: "フランス", debt: 113, yield: 4.55 },
    { code: "CA", name: "カナダ",   debt: 114, yield: 3.98 },
    { code: "AU", name: "豪州",     debt: 51,  yield: 5.46 },
  ];
  const jp20 = { code: "JP", name: "日本", debt: 230, yield: 3.74 };
  const missing20 = ["NZ", "CH", "NL", "DK", "BE", "AT", "PT", "FI", "SI", "DE"];

  // ---------- stats ----------
  function linreg(points) {
    const n = points.length;
    const mx = points.reduce((s, p) => s + p.debt, 0) / n;
    const my = points.reduce((s, p) => s + p.yield, 0) / n;
    let sxy = 0, sxx = 0, syy = 0;
    points.forEach(p => {
      const dx = p.debt - mx, dy = p.yield - my;
      sxy += dx * dy; sxx += dx * dx; syy += dy * dy;
    });
    const slope = sxy / sxx;
    const intercept = my - slope * mx;
    const r = sxy / Math.sqrt(sxx * syy);
    return { slope, intercept, r, r2: r * r, n };
  }

  const reg10 = linreg(data10);
  const reg20 = linreg(data20);
  const jp10Pred = reg10.intercept + reg10.slope * jp10.debt;
  const jp20Pred = reg20.intercept + reg20.slope * jp20.debt;

  document.getElementById("stat10").textContent =
    `他${reg10.n}か国の回帰: r = ${reg10.r.toFixed(2)} (R² = ${reg10.r2.toFixed(2)}) ／ 日本の予測値 ${jp10Pred.toFixed(2)}% に対し実際は ${jp10.yield.toFixed(2)}%(乖離 ${(jp10.yield - jp10Pred).toFixed(2)}pt)`;
  document.getElementById("stat20").textContent =
    `他${reg20.n}か国の回帰(サンプル僅少): r = ${reg20.r.toFixed(2)} (R² = ${reg20.r2.toFixed(2)}) ／ 日本の予測値 ${jp20Pred.toFixed(2)}% に対し実際は ${jp20.yield.toFixed(2)}%(乖離 ${(jp20.yield - jp20Pred).toFixed(2)}pt)`;

  document.getElementById("note10").textContent =
    "10年ゾーンは政策金利サイクルの影響が大きく、豪州・NZ(低債務・高金利)とスイス(低債務・超低金利)が回帰線を大きく振らすため、相関はあるものの弱め(R²≈0.25)。それでも日本は回帰予測値より " + Math.abs(jp10.yield - jp10Pred).toFixed(1) + "pt程度低い位置にあり、30年ゾーンと同様に「割高」な状態が確認できる。";
  document.getElementById("note20").textContent =
    "20年債は米・英・日以外では標準的なベンチマーク銘柄が存在しない国が多く、比較可能なのは" + reg20.n + "か国のみ(" + missing20.join("・") + "はデータなし)。サンプルが極端に少ないため回帰線・R²は参考値に留め、過度な解釈は禁物。";

  // ---------- render ----------
  function renderChart(svgId, points, jp, reg, jpPred, xMax, yMax) {
    const svg = document.getElementById(svgId);
    const W = 620, H = 440;
    const M = { top: 16, right: 20, bottom: 40, left: 42 };
    const plotW = W - M.left - M.right;
    const plotH = H - M.top - M.bottom;

    const xScale = v => M.left + (v / xMax) * plotW;
    const yScale = v => M.top + plotH - (v / yMax) * plotH;

    let svgContent = "";

    // gridlines + y ticks
    const yTicks = [];
    const yStep = yMax <= 6 ? 1 : (yMax <= 8 ? 1 : 2);
    for (let v = 0; v <= yMax + 0.001; v += yStep) yTicks.push(Math.round(v * 10) / 10);
    yTicks.forEach(v => {
      const y = yScale(v);
      svgContent += `<line class="grid-line" x1="${M.left}" y1="${y}" x2="${W - M.right}" y2="${y}"></line>`;
      svgContent += `<text class="tick-label" x="${M.left - 8}" y="${y + 3}" text-anchor="end">${v}%</text>`;
    });

    // x ticks
    const xTicks = [0, 50, 100, 150, 200, 250].filter(v => v <= xMax + 1);
    xTicks.forEach(v => {
      const x = xScale(v);
      svgContent += `<text class="tick-label" x="${x}" y="${H - M.bottom + 16}" text-anchor="middle">${v}</text>`;
    });

    // baseline (x-axis)
    svgContent += `<line class="baseline" x1="${M.left}" y1="${H - M.bottom}" x2="${W - M.right}" y2="${H - M.bottom}"></line>`;
    svgContent += `<line class="baseline" x1="${M.left}" y1="${M.top}" x2="${M.left}" y2="${H - M.bottom}"></line>`;

    // axis labels
    svgContent += `<text class="axis-label" x="${M.left + plotW / 2}" y="${H - 4}" text-anchor="middle">一般政府粗債務残高、対GDP比(%)</text>`;
    svgContent += `<text class="axis-label" transform="translate(12, ${M.top + plotH / 2}) rotate(-90)" text-anchor="middle">利回り(%)</text>`;

    // trend line (across full x range using regression)
    const trendCls = reg.n <= 6 ? "trend-line weak" : "trend-line";
    const xA = 0, xB = xMax;
    const yA = reg.intercept + reg.slope * xA;
    const yB = reg.intercept + reg.slope * xB;
    svgContent += `<line class="${trendCls}" x1="${xScale(xA)}" y1="${yScale(Math.max(0, Math.min(yMax, yA)))}" x2="${xScale(xB)}" y2="${yScale(Math.max(0, Math.min(yMax, yB)))}"></line>`;

    // arrow between actual JP and predicted JP
    const jpX = xScale(jp.debt);
    const jpYActual = yScale(jp.yield);
    const jpYPred = yScale(Math.max(0, Math.min(yMax, jpPred)));
    svgContent += `<line class="arrow" x1="${jpX}" y1="${jpYActual - 10}" x2="${jpX}" y2="${jpYPred + 10}" marker-end="url(#arrowhead-${svgId})"></line>`;
    svgContent += `<defs><marker id="arrowhead-${svgId}" markerWidth="8" markerHeight="8" refX="4" refY="7" orient="auto"><path d="M0,0 L4,7 L8,0" fill="none" stroke="var(--series-red)" stroke-width="1.4"></path></marker></defs>`;

    // hollow predicted JP marker
    svgContent += `<circle class="dot-hollow" cx="${jpX}" cy="${jpYPred}" r="6.5"></circle>`;

    // country points
    points.forEach(p => {
      const x = xScale(p.debt), y = yScale(p.yield);
      svgContent += `<circle class="dot-blue" cx="${x}" cy="${y}" r="5.5" data-tip="${p.name}(${p.code})｜債務 ${p.debt}%｜利回り ${p.yield.toFixed(2)}%"></circle>`;
      svgContent += `<text class="point-label" x="${x + 8}" y="${y - 7}">${p.code}</text>`;
      svgContent += `<circle class="hit" cx="${x}" cy="${y}" r="11" data-tip="${p.name}(${p.code})｜債務 ${p.debt}%｜利回り ${p.yield.toFixed(2)}%"></circle>`;
    });

    // JP actual point
    svgContent += `<circle class="dot-red" cx="${jpX}" cy="${jpYActual}" r="6.5" data-tip="日本(JP)｜債務 ${jp.debt}%｜利回り ${jp.yield.toFixed(2)}%"></circle>`;
    svgContent += `<text class="jp-label" x="${jpX + 10}" y="${jpYActual + 4}">JP</text>`;
    svgContent += `<circle class="hit" cx="${jpX}" cy="${jpYActual}" r="12" data-tip="日本(JP)｜債務 ${jp.debt}%｜利回り ${jp.yield.toFixed(2)}%"></circle>`;
    svgContent += `<circle class="hit" cx="${jpX}" cy="${jpYPred}" r="12" data-tip="日本｜回帰予測値 ${jpPred.toFixed(2)}%"></circle>`;

    svg.innerHTML = svgContent;
  }

  renderChart("chart10", data10, jp10, reg10, jp10Pred, 250, 7);
  renderChart("chart20", data20, jp20, reg20, jp20Pred, 250, 6.5);

  // ---------- tooltip ----------
  const tooltip = document.getElementById("tooltip");
  document.querySelectorAll("circle[data-tip]").forEach(el => {
    el.addEventListener("mousemove", e => {
      tooltip.textContent = el.getAttribute("data-tip");
      tooltip.style.left = e.clientX + "px";
      tooltip.style.top = e.clientY + "px";
      tooltip.classList.add("show");
    });
    el.addEventListener("mouseleave", () => tooltip.classList.remove("show"));
  });

  // ---------- data tables ----------
  function renderTable(tableId, points, jp, caption) {
    const table = document.getElementById(tableId);
    const rows = [...points, jp].sort((a, b) => a.debt - b.debt);
    let html = `<caption>${caption}</caption><thead><tr><th>国</th><th>債務/GDP(%)</th><th>利回り(%)</th></tr></thead><tbody>`;
    rows.forEach(p => {
      const cls = p.code === "JP" ? ' class="jp-row"' : "";
      html += `<tr${cls}><td>${p.name}</td><td>${p.debt}</td><td>${p.yield.toFixed(2)}</td></tr>`;
    });
    html += "</tbody>";
    table.innerHTML = html;
  }
  renderTable("table10", data10, jp10, "元データ(10年、債務比率の低い順)");
  renderTable("table20", data20, jp20, "元データ(20年、データ取得できた国のみ)");

  // ---------- theme toggle ----------
  const root = document.documentElement;
  document.getElementById("themeToggle").addEventListener("click", () => {
    const cur = root.getAttribute("data-theme");
    root.setAttribute("data-theme", cur === "dark" ? "light" : "dark");
  });
})();
