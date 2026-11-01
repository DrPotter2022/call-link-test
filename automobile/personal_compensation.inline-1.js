let chart, rawData;

  async function loadData() {
    const res = await fetch('data/insurance_data_by_year_cleaned.json');
    rawData = await res.json();
    const yearSel = document.getElementById('year');
    Object.keys(rawData).sort((a,b)=>b-a)
      .forEach(y=>{
        const opt = document.createElement('option');
        opt.value = y; opt.textContent = y;
        yearSel.append(opt);
      });
    yearSel.addEventListener('change', updateVehicleTypes);
    updateVehicleTypes();
  }

  function updateVehicleTypes() {
    const year = document.getElementById('year').value;
    const vDiv = document.getElementById('vehicleTypes');
    vDiv.innerHTML = '';
    if (!rawData[year]) return;
    Object.keys(rawData[year]).forEach(type=>{
      const lbl = document.createElement('label');
      const cb  = document.createElement('input');
      cb.type  = 'checkbox';
      cb.value = type;
      cb.addEventListener('change', drawChart);
      lbl.append(cb, document.createTextNode(' '+type));
      vDiv.append(lbl);
    });
    drawChart();
  }

  function drawChart() {
    const year = document.getElementById('year').value;
    const checks = Array.from(document.querySelectorAll('#vehicleTypes input:checked'))
                        .map(cb=>cb.value);
    if (!year || !checks.length) {
      if (chart) chart.destroy();
      return;
    }
    const yearData = rawData[year];
    const labels = Object.keys(yearData[checks[0]]);

    const datasets = checks.map((type, idx) => ({
      label: type,
      data: labels.map(lbl => yearData[type][lbl] || 0),
      backgroundColor: labels.map((_,i)=>`hsl(${i*60},70%,${60 - idx*10}%)`),
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
          legend: { position: 'right', labels:{ boxWidth:12 } },
          title: {
            display: true,
            text: `${year}年度 選択車種の保険金額帯別加入割合（%）`,
            padding: { top:0, bottom:0 }
          }
        }
      }
    });
  }

  window.onload = loadData;
