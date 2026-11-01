let incomeData = {};
    let leibnizData = {};
    const disabilityRates = {"1":100,"2":100,"3":100,"4":92,"5":79,"6":67,"7":56,"8":45,"9":35,"10":27,"11":20,"12":14,"13":9,"14":5};

    async function loadData() {
      try {
        const res1 = await fetch('data/average_income.json');
        incomeData = await res1.json();
      } catch(e) {
        console.error('平均給与 JSON 読込失敗', e);
      }
      try {
        const res2 = await fetch('data/leibniz_coeff.json');
        leibnizData = await res2.json();
      } catch(e) {
        console.error('係数 JSON 読込失敗', e);
      }
    }

    function populateAgeOptions() {
      const sel = document.getElementById('age');
      sel.innerHTML = '';
      for (let i = 0; i <= 102; i++) {
        const o = document.createElement('option');
        o.value = i;
        o.textContent = i + '\u6B73';
        sel.appendChild(o);
      }
    }

    function getAgeClass(a) {
      if (a <= 19) return '19歳以下';
      if (a <= 24) return '20~24';
      if (a <= 29) return '25~29';
      if (a <= 34) return '30~34';
      if (a <= 39) return '35~39';
      if (a <= 44) return '40~44';
      if (a <= 49) return '45~49';
      if (a <= 54) return '50~54';
      if (a <= 59) return '55~59';
      if (a <= 64) return '60~64';
      if (a <= 69) return '65~69';
      return '70歳以上';
    }

    function updateIncome() {
      const age = Number(document.getElementById('age').value);
      const gender = document.getElementById('gender').value;
      const cls = getAgeClass(age);
      const key = '平均給与（円）_' + gender;
      const val = (incomeData[cls] || {})[key] || 0;
      document.getElementById('income').value = val.toLocaleString();
    }

    function updateCoefficient() {
      const age = document.getElementById('age').value;
      document.getElementById('coefficient').value = leibnizData[age] ?? '';
    }

    function updateDisabilityRate() {
      const r = disabilityRates[document.getElementById('grade').value];
      document.getElementById('disability-rate').value = r != null ? r : '';
    }

    function setHousewife() {
      document.getElementById('income').value = '4,194,400';
    }

    function setStudentIncome(g, e) {
      let v = 0;
      if (g === '男性' && e === '全学歴') v = 5908100;
      if (g === '男性' && e === '大卒')   v = 6835500;
      if (g === '女性' && e === '全学歴') v = 4194400;
      if (g === '女性' && e === '大卒')   v = 4700200;
      document.getElementById('income').value = v.toLocaleString();
    }

    function updateStudentButtons() {
      const age = Number(document.getElementById('age').value);
      const g   = document.getElementById('gender').value;
      const show = age <= 30;
      const container = document.getElementById('student-buttons');
      container.classList.toggle('hidden', !show);
      ['btn-male-all','btn-male-uni','btn-female-all','btn-female-uni'].forEach(id => {
        const b = document.getElementById(id);
        if (b) b.style.display = 'none';
      });
      if (show) {
        if (g === '男性') {
          document.getElementById('btn-male-all').style.display = 'inline-block';
          document.getElementById('btn-male-uni').style.display = 'inline-block';
        } else if (g === '女性') {
          document.getElementById('btn-female-all').style.display = 'inline-block';
          document.getElementById('btn-female-uni').style.display = 'inline-block';
        }
      }
    }

    function calculate() {
      const inc  = parseFloat(document.getElementById('income').value.replace(/,/g,'')) || 0;
      const coef = parseFloat(document.getElementById('coefficient').value) || 0;
      let res = 0;
      if (document.getElementById('calc-type').value === 'death') {
        const d = (parseFloat(document.getElementById('deduction').value) || 0) / 100;
        res = inc * (1 - d) * coef;
      } else {
        const d = (parseFloat(document.getElementById('disability-rate').value) || 0) / 100;
        res = inc * d * coef;
      }
      document.getElementById('result').textContent = '逸失利益：' + Math.round(res).toLocaleString() + ' 円';
    }

    function formatIncome() {
      const el = document.getElementById('income');
      const v  = el.value.replace(/,/g,'');
      if (!isNaN(v) && v !== '') {
        el.value = v.replace(/\B(?=(\d{3})+(?!\d))/g,',');
      }
    }

    function unformatIncome() {
      document.getElementById('income').value = document.getElementById('income').value.replace(/,/g,'');
    }

    function reformatIncome() {
      formatIncome();
    }

    function switchMode() {
      const m = document.getElementById('calc-type').value;
      document.getElementById('death-section').classList.toggle('hidden', m !== 'death');
      document.getElementById('disability-section').classList.toggle('hidden', m !== 'disability');
      updateIncome();
      updateCoefficient();
      updateDisabilityRate();
      updateStudentButtons();
    }

    document.addEventListener('DOMContentLoaded', async () => {
      await loadData();
      populateAgeOptions();
      document.getElementById('calc-type').addEventListener('change', switchMode);
      document.getElementById('deduction-select').addEventListener('change', () => {
        document.getElementById('deduction').value = parseFloat(document.getElementById('deduction-select').value);
      });
      document.getElementById('age').addEventListener('change', () => { updateIncome(); updateCoefficient(); updateStudentButtons(); });
      document.getElementById('gender').addEventListener('change', () => { updateIncome(); updateCoefficient(); updateStudentButtons(); });
      document.getElementById('grade').addEventListener('change', updateDisabilityRate);
      switchMode();
    });
  