const metrics = [
    { key: '対人賠償',     id: 'chart-tainin'   },
    { key: '対物賠償',     id: 'chart-taibutsu' },
    { key: '人身傷害',     id: 'chart-jinshin'  },
    { key: '搭乗者傷害',   id: 'chart-toujousha'},
    { key: '車両保険',     id: 'chart-sharyo'   }
  ];

  const select     = document.getElementById('pref-select');
  const compareSel = document.getElementById('compare-select');
  const boxes      = document.getElementById('checkboxes');
  const ctxSingle  = document.getElementById('singleChart').getContext('2d');

  let allData, singleChart, compareCharts = {};

  // JSONを読み込んで、セレクト類を初期化
  fetch('data/insurance_data.json')
    .then(res => {
      if(!res.ok) throw new Error(`HTTP ${res.status}`);
      return res.json();
    })
    .then(json => {
      allData = json;

      // 単一選択セレクトに追加
      Object.keys(allData).forEach(pref => {
        const opt = new Option(pref, pref);
        select.add(opt);
      });
      select.value = '全国平均';

      // 複数選択セレクトに追加
      Object.keys(allData).forEach(pref => {
        const opt = new Option(pref, pref);
        compareSel.add(opt);
      });
      // デフォルトで全国平均のみ
      compareSel.value = [];

      // Chart.js インスタンスを生成
      singleChart = new Chart(ctxSingle, {
        type: 'line',
        data: { labels: [], datasets: [] },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: { y: { beginAtZero: true } }
        }
      });

      metrics.forEach(m => {
        const ctx = document.getElementById(m.id).getContext('2d');
        compareCharts[m.key] = new Chart(ctx, {
          type: 'line',
          data: { labels: [], datasets: [] },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { title: { display: true, text: m.key } },
            scales: { y: { beginAtZero: true } }
          }
        });
      });

      // 初期描画
      updateSingleChart();
      updateCompareCharts();
    })
    .catch(err => {
      alert('JSON読み込み失敗：' + err);
    });

  function updateSingleChart() {
    const pref = select.value;
    if (!allData[pref]) {
      singleChart.data.labels = [];
      singleChart.data.datasets = [];
      singleChart.update();
      return;
    }
    const raw = allData[pref];
    // 配列形式 or オブジェクト内配列形式の両対応
    const labels = Array.isArray(raw)
      ? raw.map(r => r.年度)
      : raw.年度;
    const getArr = field => Array.isArray(raw)
      ? raw.map(r => r[field])
      : raw[field];

    const datasets = Array.from(boxes.querySelectorAll('input[type=checkbox]'))
      .filter(cb => cb.checked)
      .map(cb => ({
        label: cb.value,
        data: getArr(cb.value),
        borderWidth: 2,
        fill: false
      }));

    singleChart.data.labels    = labels;
    singleChart.data.datasets  = datasets;
    singleChart.update();
  }

  function updateCompareCharts() {
    // 選択された都道府県の配列（全国平均は必ず先頭に）
    const selected = Array.from(compareSel.selectedOptions).map(o => o.value);
    const prefs = ['全国平均', ...selected.filter(p => p!=='全国平均')];

    // ラベル（年度）は全国平均のデータを使って取得すればOK
    const raw0 = allData['全国平均'];
    const labels = Array.isArray(raw0)
      ? raw0.map(r => r.年度)
      : raw0.年度;

    metrics.forEach(m => {
      const chart = compareCharts[m.key];
      const datasets = prefs.map(pref => {
        const raw = allData[pref];
        const arr = Array.isArray(raw)
          ? raw.map(r => r[m.key])
          : raw[m.key];
        return {
          label: pref,
          data: arr,
          borderWidth: 2,
          fill: false
        };
      });
      chart.data.labels   = labels;
      chart.data.datasets = datasets;
      chart.update();
    });
  }

  // イベント登録
  select.addEventListener('change', updateSingleChart);
  boxes.querySelectorAll('input[type=checkbox]')
    .forEach(cb => cb.addEventListener('change', updateSingleChart));
  compareSel.addEventListener('change', updateCompareCharts);
