let data = [];

    fetch('data/warasse_nebuta.json')
      .then(r => r.json())
      .then(json => {
        data = json;
        populateYearSelector();
        displayDataForYear('all');
      });

    function populateYearSelector() {
      const selector = document.getElementById("yearSelector");
      selector.innerHTML = "";
      // 「すべて表示」
      const allOpt = document.createElement("option");
      allOpt.value = "all";
      allOpt.textContent = "すべて表示";
      selector.appendChild(allOpt);
      // 年度一覧
      const years = [...new Set(data.map(item => String(item["年度"])))]
                      .sort((a,b) => a.localeCompare(b));
      years.forEach(yr => {
        const opt = document.createElement("option");
        opt.value = yr;
        opt.textContent = yr;
        selector.appendChild(opt);
      });
      selector.addEventListener("change", () => displayDataForYear(selector.value));
    }

    function displayDataForYear(year) {
      const tbody = document.querySelector("#nebutaTable tbody");
      tbody.innerHTML = "";
      // 年度の型を文字列に変換して比較
      const list = (year === 'all')
        ? data
        : data.filter(item => String(item["年度"]) === year);
      list.forEach(row => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${row["年度"]||""}</td>
          <td>${row["団体名"]||""}</td>
          <td><a href="${row["URL"]||"#"}" target="_blank">${row["題材"]||""}</a></td>
          <td>${row["制作者"]||""}</td>
          <td>${row["備考"]||""}</td>
        `;
        tbody.appendChild(tr);
      });
    }
  