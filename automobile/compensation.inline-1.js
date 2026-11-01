let chart, rawData;
    async function loadData() {
      const res = await fetch('data/insurance_data_by_year_cleaned.json');
      rawData = await res.json();
      const years = Object.keys(rawData).map(y => parseInt(y, 10)).sort((a, b) => a - b);
      const slider = document.getElementById('yearSlider');
      const display = document.getElementById('yearDisplay');
      slider.min = years[0];
      slider.max = years[years.length - 1];
      slider.step = 1;
      slider.value = slider.max;
      display.textContent = slider.value;
      slider.addEventListener('input', () => {
        display.textContent = slider.value;
        updateVehicleTypes();
      });
      updateVehicleTypes();
    }
    function updateVehicleTypes() {
      // 保持しているチェック状態取得
      const prevChecked = Array.from(
        document.querySelectorAll('#vehicleTypes input:checked')
      ).map(cb => cb.value);
      const year = String(document.getElementById('yearSlider').value);
      const container = document.getElementById('vehicleTypes');
      container.innerHTML = '';
      if (!rawData[year]) return;
      Object.keys(rawData[year]).forEach(type => {
        const lbl = document.createElement('label');
        const cb = document.createElement('input');
        cb.type = 'checkbox'; cb.value = type;
        if (prevChecked.includes(type)) cb.checked = true;
        cb.addEventListener('change', drawChart);
        lbl.append(cb, document.createTextNode(' ' + type));
        container.append(lbl);
      });
      drawChart();
    }
    function drawChart() {
      const year = String(document.getElementById('yearSlider').value);
      const choices = Array.from(
        document.querySelectorAll('#vehicleTypes input:checked')
      ).map(cb => cb.value);
      if (!year || choices.length === 0) {
        if (chart) chart.destroy();
        return;
      }
      const labels = Object.keys(rawData[year][choices[0]]);
      const datasets = choices.map((type, idx) => ({
        label: type,
        data: labels.map(lbl => rawData[year][type][lbl] || 0),
        backgroundColor: labels.map((_, i) => `hsl(${i*60},70%,${60-idx*10}%)`),
        borderColor: '#fff', borderWidth: 1
      }));
      const ctx = document.getElementById('chartCanvas').getContext('2d');
      if (chart) chart.destroy();
      chart = new Chart(ctx, {
        type: 'doughnut',
        data: { labels, datasets },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          layout: { padding: 0 },
          plugins: {
            legend: { position: 'right', labels: { boxWidth: 12 } },
            title: {
              display: true,
              text: `${year}年度 選択車種の保険金額帯別加入割合（%）`,
              padding: { top: 0, bottom: 0 }
            }
          }
        }
      });
    }
    window.onload = loadData;
  