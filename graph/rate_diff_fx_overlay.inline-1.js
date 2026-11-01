(function () {
  const rows = [
    { y: 2013, diff: 1.629, fx: 97.60 },
    { y: 2014, diff: 1.975, fx: 105.94 },
    { y: 2015, diff: 1.760, fx: 121.04 },
    { y: 2016, diff: 1.871, fx: 108.79 },
    { y: 2017, diff: 2.269, fx: 112.17 },
    { y: 2018, diff: 2.828, fx: 110.42 },
    { y: 2019, diff: 2.230, fx: 109.01 },
    { y: 2020, diff: 0.885, fx: 106.77 },
    { y: 2021, diff: 1.385, fx: 109.75 },
    { y: 2022, diff: 2.739, fx: 131.50 },
    { y: 2023, diff: 3.389, fx: 140.49 },
    { y: 2024, diff: 3.297, fx: 151.37 },
  ];
  const current = { y: 2026.5, label: "2026/8(スポット)", diff: 1.83, fx: 159.5 };

  const tooltip = document.getElementById("tooltip");
  function bindTip(el, text) {
    el.setAttribute("data-tip", text);
    el.addEventListener("mousemove", e => {
      tooltip.textContent = el.getAttribute("data-tip");
      tooltip.style.left = e.clientX + "px";
      tooltip.style.top = e.clientY + "px";
      tooltip.classList.add("show");
    });
    el.addEventListener("mouseleave", () => tooltip.classList.remove("show"));
  }

  const svg = document.getElementById("chartIndex");
  const noteEl = document.getElementById("indexNote");
  const tableEl = document.getElementById("dataTable");
  const baseSelect = document.getElementById("baseYearSelect");

  function render(baseYear) {
    const baseRow = rows.find(r => r.y === baseYear);
    const baseDiff = baseRow.diff, baseFx = baseRow.fx;
    const allPts = rows.concat([current]).map(r => ({
      y: r.y, label: r.label, diff: r.diff, fx: r.fx,
      idxDiff: (r.diff / baseDiff) * 100, idxFx: (r.fx / baseFx) * 100,
    }));

    const W = 820, H = 320, M = { top: 16, right: 60, bottom: 26, left: 44 };
    const plotW = W - M.left - M.right, plotH = H - M.top - M.bottom;
    const xMin = 2013, xMax = 2026.8;
    const idxVals = allPts.flatMap(p => [p.idxDiff, p.idxFx]);
    const yMin = Math.min(0, Math.floor(Math.min(...idxVals) / 20) * 20);
    const yMax = Math.ceil(Math.max(...idxVals) / 20) * 20 + 10;
    const xScale = yr => M.left + ((yr - xMin) / (xMax - xMin)) * plotW;
    const yScale = v => M.top + plotH - ((v - yMin) / (yMax - yMin)) * plotH;

    let s = "";
    const step = (yMax - yMin) > 250 ? 50 : 25;
    for (let v = Math.ceil(yMin / step) * step; v <= yMax; v += step) {
      const y = yScale(v);
      s += `<line class="grid-line" x1="${M.left}" y1="${y}" x2="${W - M.right}" y2="${y}"></line>`;
      s += `<text class="tick-label" x="${M.left - 8}" y="${y + 3}" text-anchor="end">${v}</text>`;
    }
    rows.forEach(r => { s += `<text class="tick-label" x="${xScale(r.y)}" y="${H - M.bottom + 15}" text-anchor="middle">${r.y}</text>`; });
    s += `<text class="tick-label" x="${xScale(2026)}" y="${H - M.bottom + 15}" text-anchor="middle">2026</text>`;
    s += `<line class="baseline" x1="${M.left}" y1="${M.top + plotH}" x2="${W - M.right}" y2="${M.top + plotH}"></line>`;
    s += `<line class="zero-line" x1="${M.left}" y1="${yScale(100)}" x2="${W - M.right}" y2="${yScale(100)}"></line>`;
    s += `<text class="tick-label" x="${W - M.right + 4}" y="${yScale(100) + 3}">基準(100)</text>`;

    function drawSeries(key, color, label) {
      let path = "";
      rows.forEach((r, i) => {
        const p = allPts[i];
        path += (i === 0 ? "M" : "L") + xScale(r.y) + "," + yScale(p[key]) + " ";
      });
      s += `<path class="series-line" d="${path}" stroke="${color}"></path>`;
      rows.forEach((r, i) => {
        const p = allPts[i];
        s += `<circle class="series-dot" cx="${xScale(r.y)}" cy="${yScale(p[key])}" r="3.3" fill="${color}"></circle>`;
        const rawLabel = key === "idxDiff" ? `${r.diff.toFixed(2)}pt` : `${r.fx.toFixed(1)}円`;
        s += `<circle class="hit" cx="${xScale(r.y)}" cy="${yScale(p[key])}" r="9" data-tip="${label}｜${r.y}年｜指数 ${p[key].toFixed(1)}(実数 ${rawLabel})"></circle>`;
      });
      const lastRow = rows[rows.length - 1];
      const lastP = allPts[allPts.length - 2];
      const curP = allPts[allPts.length - 1];
      s += `<line class="series-line dashed" x1="${xScale(lastRow.y)}" y1="${yScale(lastP[key])}" x2="${xScale(current.y)}" y2="${yScale(curP[key])}" stroke="var(--s-red)"></line>`;
      s += `<circle class="series-dot" cx="${xScale(current.y)}" cy="${yScale(curP[key])}" r="4" fill="var(--s-red)"></circle>`;
      const rawLabelCur = key === "idxDiff" ? `${current.diff.toFixed(2)}pt` : `${current.fx.toFixed(1)}円`;
      s += `<circle class="hit" cx="${xScale(current.y)}" cy="${yScale(curP[key])}" r="10" data-tip="${label}｜${current.label}｜指数 ${curP[key].toFixed(1)}(実数 ${rawLabelCur})"></circle>`;
      s += `<text class="end-label" x="${xScale(current.y) + 8}" y="${yScale(curP[key]) + 3}" fill="${color}">${label}</text>`;
    }
    drawSeries("idxDiff", "var(--s-blue)", "金利差");
    drawSeries("idxFx", "var(--s-orange)", "ドル円");

    s += `<text class="axis-label" transform="translate(12, ${M.top + plotH/2}) rotate(-90)" text-anchor="middle">指数(${baseYear}年 = 100)</text>`;

    svg.innerHTML = s;
    svg.querySelectorAll(".hit").forEach(el => bindTip(el, el.getAttribute("data-tip")));

    const curIdxDiff = allPts[allPts.length - 1].idxDiff;
    const curIdxFx = allPts[allPts.length - 1].idxFx;
    const higher = curIdxDiff > curIdxFx ? "金利差指数" : "ドル円指数";
    const lower = curIdxDiff > curIdxFx ? "ドル円指数" : "金利差指数";
    const baseYearNote = baseYear === 2020
      ? "なお2020年は金利差が期間中最も縮小した「谷」の年のため、これを基準にすると金利差指数の変動幅が機械的に誇張される点に注意。基準年の選び方自体が印象を左右するため、複数の基準年で見比べることを推奨する。"
      : "";
    noteEl.textContent = `基準年=${baseYear}年。2026年8月時点で金利差指数は${curIdxDiff.toFixed(0)}(${baseYear}年比${curIdxDiff >= 100 ? "+" : ""}${(curIdxDiff-100).toFixed(0)}%)、ドル円指数は${curIdxFx.toFixed(0)}(同${curIdxFx >= 100 ? "+" : ""}${(curIdxFx-100).toFixed(0)}%)で、${higher}が${lower}を上回っている。${baseYearNote}`;

    let html = `<caption>年次データ一覧(基準年=${baseYear}年)</caption><thead><tr><th>年</th><th>金利差(pt)</th><th>金利差指数</th><th>ドル円(円)</th><th>ドル円指数</th></tr></thead><tbody>`;
    rows.forEach((r, i) => {
      const p = allPts[i];
      html += `<tr><td>${r.y}</td><td>${r.diff.toFixed(2)}</td><td>${p.idxDiff.toFixed(1)}</td><td>${r.fx.toFixed(2)}</td><td>${p.idxFx.toFixed(1)}</td></tr>`;
    });
    const curP = allPts[allPts.length - 1];
    html += `<tr class="cur-row"><td>${current.label}</td><td>${current.diff.toFixed(2)}</td><td>${curP.idxDiff.toFixed(1)}</td><td>${current.fx.toFixed(1)}</td><td>${curP.idxFx.toFixed(1)}</td></tr>`;
    html += "</tbody>";
    tableEl.innerHTML = html;
  }

  baseSelect.addEventListener("change", () => render(parseInt(baseSelect.value, 10)));
  render(2013);

  document.getElementById("themeToggle").addEventListener("click", () => {
    const root = document.documentElement;
    const cur = root.getAttribute("data-theme");
    root.setAttribute("data-theme", cur === "dark" ? "light" : "dark");
  });
})();
