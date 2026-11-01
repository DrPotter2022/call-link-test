const JSON_PATH = 'data/insurance_data.json';

    const select = document.getElementById('pref-select');
    const checkboxContainer = document.getElementById('checkboxes');
    const ctx = document.getElementById('myChart').getContext('2d');
    const statusEl = document.getElementById('status');

    const autoYAxis = document.getElementById('auto-y-axis');
    const yMinRange = document.getElementById('y-min-range');
    const yMaxRange = document.getElementById('y-max-range');
    const yMinInput = document.getElementById('y-min-input');
    const yMaxInput = document.getElementById('y-max-input');
    const autoFitBtn = document.getElementById('auto-fit-btn');
    const resetAxisBtn = document.getElementById('reset-axis-btn');

    let allData;
    let chartInstance;
    let currentDatasets = [];

    fetch(JSON_PATH)
      .then(res => {
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return res.json();
      })
      .then(json => {
        allData = json;
        Object.keys(allData).forEach(pref => {
          const opt = document.createElement('option');
          opt.value = pref;
          opt.textContent = pref;
          select.appendChild(opt);
        });
        select.value = allData['全国平均'] ? '全国平均' : Object.keys(allData)[0];
        statusEl.textContent = `JSON読み込み完了：${JSON_PATH}`;
        renderChart();
      })
      .catch(err => {
        statusEl.textContent = `JSON読み込み失敗：${JSON_PATH} / ${err.message}`;
        alert('JSON読み込み失敗：' + err.message + '\n\nGitHub上で HTML と同じ階層に data/insurance_data.json があるか確認してください。');
      });

    function toNumber(value) {
      if (value === null || value === undefined || value === '') return null;
      const num = Number(String(value).replace('%', '').trim());
      return Number.isFinite(num) ? num : null;
    }

    function getVisibleValues(datasets) {
      return datasets
        .flatMap(d => d.data)
        .map(toNumber)
        .filter(v => v !== null);
    }

    function getAutoAxis(values) {
      if (!values.length) return { min: 0, max: 100 };
      const minValue = Math.min(...values);
      const maxValue = Math.max(...values);
      const span = Math.max(maxValue - minValue, 1);
      const padding = Math.max(span * 0.15, 0.5);
      const min = Math.max(0, Math.floor((minValue - padding) * 10) / 10);
      const max = Math.min(100, Math.ceil((maxValue + padding) * 10) / 10);
      return min >= max ? { min: Math.max(0, min - 1), max: Math.min(100, max + 1) } : { min, max };
    }

    function syncAxisInputs(min, max) {
      yMinRange.value = min;
      yMinInput.value = min;
      yMaxRange.value = max;
      yMaxInput.value = max;
    }

    function getManualAxis() {
      let min = toNumber(yMinInput.value);
      let max = toNumber(yMaxInput.value);
      if (min === null) min = 0;
      if (max === null) max = 100;
      min = Math.max(0, Math.min(100, min));
      max = Math.max(0, Math.min(100, max));
      if (min >= max) max = Math.min(100, min + 1);
      return { min, max };
    }

    function setAxisControlsDisabled(disabled) {
      yMinRange.disabled = disabled;
      yMinInput.disabled = disabled;
      yMaxRange.disabled = disabled;
      yMaxInput.disabled = disabled;
    }

    function getAxisForChart(datasets) {
      if (autoYAxis.checked) {
        const axis = getAutoAxis(getVisibleValues(datasets));
        syncAxisInputs(axis.min, axis.max);
        setAxisControlsDisabled(true);
        return axis;
      }
      setAxisControlsDisabled(false);
      const axis = getManualAxis();
      syncAxisInputs(axis.min, axis.max);
      return axis;
    }

    function renderChart() {
      if (!allData) return;

      const raw = allData[select.value];
      if (!raw) {
        if (chartInstance) chartInstance.destroy();
        return;
      }

      let labels, getData;
      if (Array.isArray(raw)) {
        labels = raw.map(r => r.年度);
        getData = field => raw.map(r => toNumber(r[field]));
      } else {
        labels = raw.年度;
        getData = field => raw[field].map(toNumber);
      }

      const datasets = Array.from(
        checkboxContainer.querySelectorAll('input[type=checkbox]')
      ).filter(cb => cb.checked).map(cb => ({
        label: cb.value,
        data: getData(cb.value),
        borderWidth: 2,
        fill: false,
        tension: 0.15,
        pointRadius: 3
      }));

      currentDatasets = datasets;
      const axis = getAxisForChart(datasets);

      if (chartInstance) chartInstance.destroy();

      chartInstance = new Chart(ctx, {
        type: 'line',
        data: { labels, datasets },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            tooltip: {
              callbacks: {
                label: context => `${context.dataset.label}: ${Number(context.parsed.y).toFixed(1)}%`
              }
            }
          },
          scales: {
            y: {
              min: axis.min,
              max: axis.max,
              ticks: {
                callback: value => Number(value).toFixed(1) + '%'
              }
            }
          }
        }
      });
    }

    function updateManualPair(source, target) {
      target.value = source.value;
      autoYAxis.checked = false;
      setAxisControlsDisabled(false);
      renderChart();
    }

    select.addEventListener('change', renderChart);
    checkboxContainer.querySelectorAll('input[type=checkbox]').forEach(cb => cb.addEventListener('change', renderChart));

    autoYAxis.addEventListener('change', renderChart);
    yMinRange.addEventListener('input', () => updateManualPair(yMinRange, yMinInput));
    yMinInput.addEventListener('input', () => updateManualPair(yMinInput, yMinRange));
    yMaxRange.addEventListener('input', () => updateManualPair(yMaxRange, yMaxInput));
    yMaxInput.addEventListener('input', () => updateManualPair(yMaxInput, yMaxRange));

    autoFitBtn.addEventListener('click', () => {
      const axis = getAutoAxis(getVisibleValues(currentDatasets));
      autoYAxis.checked = false;
      setAxisControlsDisabled(false);
      syncAxisInputs(axis.min, axis.max);
      renderChart();
    });

    resetAxisBtn.addEventListener('click', () => {
      autoYAxis.checked = false;
      setAxisControlsDisabled(false);
      syncAxisInputs(0, 100);
      renderChart();
    });
  