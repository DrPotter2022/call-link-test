const ageLabels = ["0歳", "1～4歳", "5～9歳", "10～14歳", "15～19歳", "20～24歳", "25～29歳", "30～34歳",
      "35～39歳", "40～44歳", "45～49歳", "50～54歳", "55～59歳", "60～64歳",
      "65～69歳", "70～74歳", "75～79歳", "80～84歳", "85～89歳", "90歳以上"];

    const population = {
      total: [757, 3331, 4839, 5248, 5494, 5737, 5670, 6115, 6582, 7290, 7821, 7765, 7629, 8142, 8131, 7783, 7095, 5794, 3873, 2756],
      male:  [388, 1705, 2478, 2687, 2818, 2952, 2923, 3175, 3388, 3749, 4023, 4020, 3983, 4297, 4281, 4090, 3692, 2923, 1875, 1176],
      female:[369, 1625, 2361, 2558, 2676, 2785, 2747, 2940, 3194, 3541, 3798, 3745, 3646, 3845, 3850, 3693, 3403, 2871, 1998, 1580]
    };

    const diseaseData = {
      "胃がん": {
        total: [0.1,0.2,0.5,0.7,1.0,1.2,1.3,1.4,1.6,1.9,2.2,2.5,2.7,2.9,3.0,2.8,2.5,2.0,1.5,1.0],
        male:  [0.1,0.2,0.5,0.7,1.0,1.2,1.3,1.4,1.6,1.9,2.2,2.5,2.7,2.9,3.0,2.8,2.5,2.0,1.5,1.0],
        female:[0.1,0.2,0.5,0.7,1.0,1.2,1.3,1.4,1.6,1.9,2.2,2.5,2.7,2.9,3.0,2.8,2.5,2.0,1.5,1.0]
      },
      "前立腺がん": {
        total: [0,0,0,0,0,0.1,0.2,0.4,0.8,1.2,1.5,1.9,2.4,3.0,3.2,3.5,3.8,3.5,3.0,2.4],
        male:  [0,0,0,0,0,0.1,0.2,0.4,0.8,1.2,1.5,1.9,2.4,3.0,3.2,3.5,3.8,3.5,3.0,2.4],
        female:[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
      },
      "尿管結石": {
        total: [0.1,0.2,0.3,0.4,0.6,0.8,1.0,1.2,1.4,1.3,1.2,1.1,1.0,0.9,0.8,0.7,0.6,0.5,0.4,0.3],
        male:  [0.1,0.2,0.3,0.4,0.6,0.8,1.0,1.2,1.4,1.3,1.2,1.1,1.0,0.9,0.8,0.7,0.6,0.5,0.4,0.3],
        female:[0.1,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1.0,1.0,0.9,0.8,0.7,0.6,0.5,0.4,0.3,0.2]
      }
    };

    const diseaseSelect = document.getElementById("diseaseSelect");
    Object.keys(diseaseData).forEach(name => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      diseaseSelect.appendChild(option);
    });

    const ctx = document.getElementById("chartCanvas").getContext("2d");
    let chart;

    function updateChart() {
      const disease = diseaseSelect.value;
      const isActual = document.getElementById("toggleActual").checked;

      const totalOn = document.getElementById("totalCheck").checked;
      const maleOn = document.getElementById("maleCheck").checked;
      const femaleOn = document.getElementById("femaleCheck").checked;

      const datasets = [];

      if (totalOn) {
        datasets.push({
          label: "総数",
          backgroundColor: "gray",
          borderColor: "gray",
          fill: false,
          data: diseaseData[disease].total.map((val, i) => isActual ? +(val * population.total[i]).toFixed(2) : val)
        });
      }
      if (maleOn) {
        datasets.push({
          label: "男性",
          backgroundColor: "blue",
          borderColor: "blue",
          fill: false,
          data: diseaseData[disease].male.map((val, i) => isActual ? +(val * population.male[i]).toFixed(2) : val)
        });
      }
      if (femaleOn) {
        datasets.push({
          label: "女性",
          backgroundColor: "pink",
          borderColor: "pink",
          fill: false,
          data: diseaseData[disease].female.map((val, i) => isActual ? +(val * population.female[i]).toFixed(2) : val)
        });
      }

      if (chart) chart.destroy();
      chart = new Chart(ctx, {
        type: 'line',
        data: { labels: ageLabels, datasets },
        options: {
          responsive: true,
          plugins: {
            title: { display: true, text: disease + "の年齢別推移" },
            tooltip: { mode: 'index', intersect: false }
          },
          interaction: { mode: 'nearest', axis: 'x', intersect: false },
          scales: {
            y: {
              title: { display: true, text: isActual ? '推定入院者数（人）' : '人口千人あたり入院率' }
            }
          }
        }
      });
    }

    diseaseSelect.addEventListener("change", updateChart);
    document.getElementById("toggleActual").addEventListener("change", updateChart);
    document.getElementById("totalCheck").addEventListener("change", updateChart);
    document.getElementById("maleCheck").addEventListener("change", updateChart);
    document.getElementById("femaleCheck").addEventListener("change", updateChart);

    updateChart();
  