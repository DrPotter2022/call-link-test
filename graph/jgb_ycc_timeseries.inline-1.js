(function () {
  const debt = [
    { y: 2013, v: 201.2 }, { y: 2014, v: 203.6 }, { y: 2015, v: 200.1 }, { y: 2016, v: 202.1 },
    { y: 2017, v: 203.1 }, { y: 2018, v: 203.7 }, { y: 2019, v: 206.3 }, { y: 2020, v: 228.8 },
    { y: 2021, v: 222.7 }, { y: 2022, v: 227.8 }, { y: 2023, v: 220.3 },
  ];
  const y10 = [0.721,0.565,0.380,-0.031,0.061,0.082,-0.090,0.005,0.065,0.211,0.571,0.883];
  const y20 = [1.638,1.402,1.138,0.437,0.591,0.566,0.298,0.351,0.447,0.841,1.247,1.712];
  const y30 = [1.762,1.636,1.411,0.564,0.834,0.794,0.448,0.530,0.672,1.130,1.498,2.027];
  const years = [2013,2014,2015,2016,2017,2018,2019,2020,2021,2022,2023,2024];

  const YCC_START = 2016.72; // 2016/9
  const YCC_END = 2024.22;   // 2024/3
  const KURODA_START = 2013.27; // 2013/4
  const KURODA_END = 2023.27;   // 2023/4

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

  // ---------- Debt panel ----------
  (function renderDebt() {
    const svg = document.getElementById("chartDebt");
    const W = 820, H = 220, M = { top: 14, right: 30, bottom: 26, left: 46 };
    const plotW = W - M.left - M.right, plotH = H - M.top - M.bottom;
    const xMin = 2013, xMax = 2023.6;
    const yMin = 190, yMax = 235;
    const xScale = yr => M.left + ((yr - xMin) / (xMax - xMin)) * plotW;
    const yScale = v => M.top + plotH - ((v - yMin) / (yMax - yMin)) * plotH;

    let s = "";
    // YCC band
    s += `<rect class="ycc-band" x="${xScale(YCC_START)}" y="${M.top}" width="${xScale(xMax) - xScale(YCC_START)}" height="${plotH}"></rect>`;
    s += `<line class="ycc-band-edge" x1="${xScale(YCC_START)}" y1="${M.top}" x2="${xScale(YCC_START)}" y2="${M.top + plotH}"></line>`;

    // gridlines
    [190,200,210,220,230].forEach(v => {
      const y = yScale(v);
      s += `<line class="grid-line" x1="${M.left}" y1="${y}" x2="${W - M.right}" y2="${y}"></line>`;
      s += `<text class="tick-label" x="${M.left - 8}" y="${y + 3}" text-anchor="end">${v}%</text>`;
    });
    debt.forEach(d => {
      const x = xScale(d.y);
      s += `<text class="tick-label" x="${x}" y="${H - M.bottom + 15}" text-anchor="middle">${d.y}</text>`;
    });
    s += `<line class="baseline" x1="${M.left}" y1="${M.top + plotH}" x2="${W - M.right}" y2="${M.top + plotH}"></line>`;

    // Kuroda markers
    [KURODA_START, KURODA_END].forEach(k => {
      const x = xScale(k);
      s += `<line class="kuroda-line" x1="${x}" y1="${M.top}" x2="${x}" y2="${M.top + plotH}"></line>`;
    });
    s += `<text class="kuroda-text" x="${xScale(KURODA_START) + 4}" y="${M.top + 10}">黒田総裁 就任(2013/4)</text>`;
    s += `<text class="kuroda-text" x="${xScale(KURODA_END) + 4}" y="${M.top + 22}">退任(2023/4)</text>`;

    // line
    let path = "";
    debt.forEach((d, i) => { path += (i === 0 ? "M" : "L") + xScale(d.y) + "," + yScale(d.v) + " "; });
    s += `<path class="series-line" d="${path}" stroke="var(--s-blue)"></path>`;
    debt.forEach(d => {
      s += `<circle class="series-dot" cx="${xScale(d.y)}" cy="${yScale(d.v)}" r="3.5" fill="var(--s-blue)"></circle>`;
      s += `<circle class="hit" cx="${xScale(d.y)}" cy="${yScale(d.v)}" r="10" data-tip="${d.y}年｜債務/GDP ${d.v}%"></circle>`;
    });
    s += `<text class="ycc-text" x="${xScale(YCC_START) + 4}" y="${M.top + plotH - 6}">YCC導入(2016/9)〜</text>`;

    svg.innerHTML = s;
    svg.querySelectorAll(".hit").forEach(el => bindTip(el, el.getAttribute("data-tip")));
  })();

  // ---------- Yield panel ----------
  (function renderYield() {
    const svg = document.getElementById("chartYield");
    const W = 820, H = 260, M = { top: 14, right: 30, bottom: 26, left: 46 };
    const plotW = W - M.left - M.right, plotH = H - M.top - M.bottom;
    const xMin = 2013, xMax = 2024.6;
    const yMin = -0.5, yMax = 2.3;
    const xScale = yr => M.left + ((yr - xMin) / (xMax - xMin)) * plotW;
    const yScale = v => M.top + plotH - ((v - yMin) / (yMax - yMin)) * plotH;

    let s = "";
    s += `<rect class="ycc-band" x="${xScale(YCC_START)}" y="${M.top}" width="${xScale(YCC_END) - xScale(YCC_START)}" height="${plotH}"></rect>`;
    s += `<line class="ycc-band-edge" x1="${xScale(YCC_START)}" y1="${M.top}" x2="${xScale(YCC_START)}" y2="${M.top + plotH}"></line>`;
    s += `<line class="ycc-band-edge" x1="${xScale(YCC_END)}" y1="${M.top}" x2="${xScale(YCC_END)}" y2="${M.top + plotH}"></line>`;

    [0,0.5,1,1.5,2].forEach(v => {
      const y = yScale(v);
      s += `<line class="grid-line" x1="${M.left}" y1="${y}" x2="${W - M.right}" y2="${y}"></line>`;
      s += `<text class="tick-label" x="${M.left - 8}" y="${y + 3}" text-anchor="end">${v}%</text>`;
    });
    years.forEach(yr => {
      const x = xScale(yr);
      s += `<text class="tick-label" x="${x}" y="${H - M.bottom + 15}" text-anchor="middle">${yr}</text>`;
    });
    s += `<line class="baseline" x1="${M.left}" y1="${yScale(0)}" x2="${W - M.right}" y2="${yScale(0)}"></line>`;

    [KURODA_START, KURODA_END].forEach(k => {
      const x = xScale(k);
      s += `<line class="kuroda-line" x1="${x}" y1="${M.top}" x2="${x}" y2="${M.top + plotH}"></line>`;
    });

    function drawSeries(vals, color, label) {
      let path = "";
      years.forEach((yr, i) => { path += (i === 0 ? "M" : "L") + xScale(yr) + "," + yScale(vals[i]) + " "; });
      s += `<path class="series-line" d="${path}" stroke="${color}"></path>`;
      years.forEach((yr, i) => {
        s += `<circle class="series-dot" cx="${xScale(yr)}" cy="${yScale(vals[i])}" r="3" fill="${color}"></circle>`;
        s += `<circle class="hit" cx="${xScale(yr)}" cy="${yScale(vals[i])}" r="9" data-tip="${label}｜${yr}年｜${vals[i].toFixed(3)}%"></circle>`;
      });
      const lastX = xScale(years[years.length - 1]);
      const lastY = yScale(vals[vals.length - 1]);
      s += `<text class="end-label" x="${lastX + 6}" y="${lastY + 3}" fill="${color}">${label}</text>`;
    }
    drawSeries(y10, "var(--s-blue)", "10年");
    drawSeries(y20, "var(--s-orange)", "20年");
    drawSeries(y30, "var(--s-aqua)", "30年");

    s += `<text class="ycc-text" x="${xScale(YCC_START) + 4}" y="${M.top + 12}">YCC期間</text>`;

    svg.innerHTML = s;
    svg.querySelectorAll(".hit").forEach(el => bindTip(el, el.getAttribute("data-tip")));
  })();

  // ---------- Table ----------
  (function renderTable() {
    const table = document.getElementById("dataTable");
    let html = `<caption>年次データ一覧(黄色網掛け = YCC実施期間中の年)</caption>
      <thead><tr><th>年</th><th>債務/GDP(%)</th><th>10年(%)</th><th>20年(%)</th><th>30年(%)</th></tr></thead><tbody>`;
    years.forEach((yr, i) => {
      const d = debt.find(x => x.y === yr);
      const inYcc = yr >= 2016 && yr <= 2024;
      html += `<tr class="${inYcc ? 'ycc-row' : ''}"><td>${yr}</td><td>${d ? d.v : '—'}</td><td>${y10[i].toFixed(3)}</td><td>${y20[i].toFixed(3)}</td><td>${y30[i].toFixed(3)}</td></tr>`;
    });
    html += "</tbody>";
    table.innerHTML = html;
  })();

  document.getElementById("themeToggle").addEventListener("click", () => {
    const root = document.documentElement;
    const cur = root.getAttribute("data-theme");
    root.setAttribute("data-theme", cur === "dark" ? "light" : "dark");
  });
})();
