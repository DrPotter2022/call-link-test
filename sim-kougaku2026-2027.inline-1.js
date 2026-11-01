// ---------------- Income options ----------------
  const INCOME_U70 = [
    {value:"A", text:"約1,650万円〜（標報:127万円〜）"},
    {value:"B", text:"約1,410〜約1,650万円（標報:103〜121万円）"},
    {value:"C", text:"約1,160〜約1,410万円（標報:83〜98万円）"},
    {value:"D", text:"約1,040〜約1,160万円（標報:71〜79万円）"},
    {value:"E", text:"約950〜約1,040万円（標報:62〜68万円）"},
    {value:"F", text:"約770〜約950万円（標報:53〜59万円）"},
    {value:"G", text:"約650〜約770万円（標報:44〜50万円）"},
    {value:"H", text:"約510〜約650万円（標報:36〜41万円）"},
    {value:"I", text:"約370〜約510万円（標報:28〜34万円）"},
    {value:"J", text:"約200〜約260万円（標報:16〜19万円）"},
    {value:"NTU70", text:"非課税（70歳未満）"},
  ];

  const INCOME_O70 = [
    {value:"A", text:"年収 約1,650万円〜（70歳以上・外来特例なし）"},
    {value:"D", text:"年収 約950〜約1,650万円相当（70歳以上・外来特例なし）"},
    {value:"G", text:"年収 約370〜約950万円相当（70歳以上・外来特例なし）"},

    {value:"J2", text:"年収 約260〜約370万円（R9.8〜のみ 外来特例あり）"},
    {value:"J",  text:"年収 約200〜約260万円（外来特例あり）"},
    {value:"K",  text:"年収 〜約200万円（外来特例あり）"},

    {value:"NTO70", text:"非課税（70歳以上）"},
    {value:"FIXO70", text:"一定所得以下（70歳以上）"},
  ];

  function loadIncomeOptions(){
    const age = document.getElementById("ageGroup").value;
    const sel = document.getElementById("income");
    sel.innerHTML = "";
    const items = age==="u70" ? INCOME_U70 : INCOME_O70;
    items.forEach(it=>{
      const opt=document.createElement("option");
      opt.value=it.value;
      opt.textContent=it.text;
      sel.appendChild(opt);
    });
    // sensible defaults
    sel.value = age==="u70" ? "G" : "NTO70";
    document.getElementById("incomeLabel").textContent = "所得区分（画像の表に準拠）";
    // show outpatient toggle only for 70+
    document.getElementById("outpatientRow").style.display = (age==="o70") ? "flex" : "none";
    if (age==="u70") document.getElementById("outpatientOnly").checked = false;

    // default copay rate
    const copaySel = document.getElementById("copayRate");
    if (age==="u70") copaySel.value = "0.3";
    else if (copaySel.value==="0.3") copaySel.value = "0.2"; // 70+は2割が多い前提で初期値だけ調整
  }

  // ---------------- Examples ----------------
  const EXAMPLES = [
    {key:"custom", label:"（手入力）", monthlyTotal:null},

    // 厚労省資料の治療例（病名で選択）
    {key:"cml", label:"慢性骨髄性白血病（分子標的薬治療の例）", monthlyTotal:2872000},
    {key:"ird", label:"遺伝性網膜ジストロフィー（ルクスターナ治療の例）", monthlyTotal:49600000},
    {key:"breast", label:"乳がん（抗がん剤等の例）", monthlyTotal:6582000},

    // 胃がん：ユーザー要望。総医療費は暫定（資料値に差替可）
    {key:"stomach", label:"胃がん（治療例）", monthlyTotal:1000000},

    {key:"ref1m", label:"参考：総医療費 100万円 / 月", monthlyTotal:1000000},
  ];

  // ---------------- Rules (from your image) ----------------
  // base + 1% * max(0, total - threshold)
  // multi = 多数回該当の上限（<...>）
  // annual = 年間上限（ある場合のみ）
  // outpatient = 外来特例（70歳以上）月額上限（ある場合のみ）
  const RULES = {
    2025: {
      // 高所得〜一般（外来特例なし）
      A:{type:"formula", base:252600, threshold:842000, rate:0.01, multi:140100, annual:null, outpatient:null},
      B:{type:"formula", base:252600, threshold:842000, rate:0.01, multi:140100, annual:null, outpatient:null},
      C:{type:"formula", base:252600, threshold:842000, rate:0.01, multi:140100, annual:null, outpatient:null},

      D:{type:"formula", base:167400, threshold:558000, rate:0.01, multi:93000, annual:null, outpatient:null},
      E:{type:"formula", base:167400, threshold:558000, rate:0.01, multi:93000, annual:null, outpatient:null},
      F:{type:"formula", base:167400, threshold:558000, rate:0.01, multi:93000, annual:null, outpatient:null},

      G:{type:"formula", base:80100, threshold:267000, rate:0.01, multi:44400, annual:null, outpatient:null},
      H:{type:"formula", base:80100, threshold:267000, rate:0.01, multi:44400, annual:null, outpatient:null},
      I:{type:"formula", base:80100, threshold:267000, rate:0.01, multi:44400, annual:null, outpatient:null},

      // 所得低め（外来特例あり）
      J:{type:"fixed", base:57600, threshold:0, rate:0, multi:44400, annual:null, outpatient:18000, outpatientAnnual:144000}, // 年14.4万は表示だけ（年額計算は months×月額でOK）
      K:{type:"fixed", base:null, threshold:0, rate:0, multi:null, annual:null, outpatient:18000, outpatientAnnual:144000}, // 〜約200は現行欄に月額上限の明記がないため空（後で使う場合は入れてください）

      // 非課税
      NTU70:{type:"fixed", base:35400, threshold:0, rate:0, multi:24600, annual:null, outpatient:null},
      NTO70:{type:"fixed", base:24600, threshold:0, rate:0, multi:null, annual:null, outpatient:8000, outpatientAnnual:96000},
      FIXO70:{type:"fixed", base:15000, threshold:0, rate:0, multi:null, annual:null, outpatient:8000, outpatientAnnual:96000},
      J2:{type:"fixed", base:null, threshold:0, rate:0, multi:null, annual:null, outpatient:null},
    },
    2026: {
      A:{type:"formula", base:270300, threshold:842000, rate:0.01, multi:140100, annual:1680000, outpatient:null},
      B:{type:"formula", base:270300, threshold:842000, rate:0.01, multi:140100, annual:1680000, outpatient:null},
      C:{type:"formula", base:270300, threshold:842000, rate:0.01, multi:140100, annual:1680000, outpatient:null},

      D:{type:"formula", base:179100, threshold:558000, rate:0.01, multi:93000, annual:1110000, outpatient:null},
      E:{type:"formula", base:179100, threshold:558000, rate:0.01, multi:93000, annual:1110000, outpatient:null},
      F:{type:"formula", base:179100, threshold:558000, rate:0.01, multi:93000, annual:1110000, outpatient:null},

      G:{type:"formula", base:85800, threshold:267000, rate:0.01, multi:44400, annual:530000, outpatient:null},
      H:{type:"formula", base:85800, threshold:267000, rate:0.01, multi:44400, annual:530000, outpatient:null},
      I:{type:"formula", base:85800, threshold:267000, rate:0.01, multi:44400, annual:530000, outpatient:null},

      J:{type:"fixed", base:61500, threshold:0, rate:0, multi:44400, annual:530000, outpatient:22000, outpatientAnnual:216000},
      K:{type:"fixed", base:null, threshold:0, rate:0, multi:null, annual:null, outpatient:22000, outpatientAnnual:216000}, // 〜約200は（※1）で年額上限の導入時期が限定されているため、ここは空

      NTU70:{type:"fixed", base:36900, threshold:0, rate:0, multi:24600, annual:290000, outpatient:null},
      NTO70:{type:"fixed", base:25700, threshold:0, rate:0, multi:24600, annual:290000, outpatient:11000, outpatientAnnual:96000},
      FIXO70:{type:"fixed", base:15700, threshold:0, rate:0, multi:null, annual:180000, outpatient:8000, outpatientAnnual:96000},
      J2:{type:"fixed", base:null, threshold:0, rate:0, multi:null, annual:null, outpatient:null},
    },
    2027: {
      // 上位区分が細分化（画像のR9.8列）
      A:{type:"formula", base:342000, threshold:842000, rate:0.01, multi:140100, annual:1680000, outpatient:null},
      B:{type:"formula", base:303000, threshold:842000, rate:0.01, multi:140100, annual:1680000, outpatient:null},
      C:{type:"formula", base:270300, threshold:842000, rate:0.01, multi:140100, annual:1680000, outpatient:null},

      D:{type:"formula", base:209400, threshold:558000, rate:0.01, multi:93000, annual:1110000, outpatient:null},
      E:{type:"formula", base:194400, threshold:558000, rate:0.01, multi:93000, annual:1110000, outpatient:null},
      F:{type:"formula", base:179100, threshold:558000, rate:0.01, multi:93000, annual:1110000, outpatient:null},

      G:{type:"formula", base:110400, threshold:267000, rate:0.01, multi:44400, annual:530000, outpatient:null},
      H:{type:"formula", base:98100, threshold:267000, rate:0.01, multi:44400, annual:530000, outpatient:null},
      I:{type:"formula", base:85800, threshold:267000, rate:0.01, multi:44400, annual:530000, outpatient:null},

      // 低所得帯（画像下段）
      J:{type:"fixed", base:65400, threshold:0, rate:0, multi:44400, annual:530000, outpatient:28000, outpatientAnnual:216000},
      K:{type:"fixed", base:61500, threshold:0, rate:0, multi:34500, annual:410000, outpatient:22000, outpatientAnnual:216000},

      NTU70:{type:"fixed", base:36900, threshold:0, rate:0, multi:24600, annual:290000, outpatient:null},
      NTO70:{type:"fixed", base:25700, threshold:0, rate:0, multi:24600, annual:290000, outpatient:13000, outpatientAnnual:96000},
      FIXO70:{type:"fixed", base:15700, threshold:0, rate:0, multi:null, annual:180000, outpatient:8000, outpatientAnnual:96000},

      J2:{type:"fixed", base:69600, threshold:0, rate:0, multi:44400, annual:530000, outpatient:28000, outpatientAnnual:216000},
    }
  };

  // ---------------- Helpers ----------------

  function parseYenInput(str){
    if (str===null || str===undefined) return 0;
    return Number(String(str).replace(/,/g,"").trim() || 0);
  }
  function formatWithCommas(n){
    if (n===null || n===undefined || Number.isNaN(n)) return "";
    return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }
  const yen = (n) => {
    if (n === null || n === undefined || Number.isNaN(n)) return "—";
    const x = Math.round(n);
    return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + "円";
  };

  const calcMonthlyLimit = (rule, total, isMultiple) => {
    if (!rule) return null;
    if (isMultiple && rule.multi !== null && rule.multi !== undefined) return rule.multi;
    if (rule.type === "fixed") return rule.base;
    const extra = Math.max(0, total - rule.threshold) * rule.rate;
    return rule.base + extra;
  };

  const ruleLogicText = (rule) => {
    if (!rule || rule.base === null || rule.base === undefined) return "—";
    if (rule.type === "fixed"){
      const multiTxt = (rule.multi!==null && rule.multi!==undefined) ? `（多数回 ${yen(rule.multi)}）` : "";
      return `固定 ${yen(rule.base)}${multiTxt}`;
    }
    const multiTxt = (rule.multi!==null && rule.multi!==undefined) ? `（多数回 ${yen(rule.multi)}）` : "";
    return `${yen(rule.base)} + 1% × max(0, 総医療費 - ${yen(rule.threshold).replace("円","")})${multiTxt}`;
  };

  function pickedIncomeLabel(){
    const sel = document.getElementById("income");
    return sel.options[sel.selectedIndex]?.text || "—";
  }

  // ---------------- UI wiring ----------------
  const totalCostInput = document.getElementById("totalCost");
  const exampleSel = document.getElementById("example");
  EXAMPLES.forEach(e=>{
    const opt=document.createElement("option");
    opt.value=e.key;
    opt.textContent=e.label;
    exampleSel.appendChild(opt);
  });

  function setExampleToCost(){
    const picked = EXAMPLES.find(x=>x.key===exampleSel.value);
    if (picked && picked.monthlyTotal){
      totalCostInput.value = formatWithCommas(picked.monthlyTotal);
    }
  }

  exampleSel.addEventListener("change", ()=>{ setExampleToCost(); recalc(); });

  document.getElementById("ageGroup").addEventListener("change", ()=>{
    loadIncomeOptions();
    recalc();
  });

  document.getElementById("calcBtn").addEventListener("click", ()=>{ setExampleToCost(); recalc(); });
  document.getElementById("resetBtn").addEventListener("click", ()=>{
    document.getElementById("ageGroup").value="u70";
    document.getElementById("copayRate").value="0.3";
    loadIncomeOptions();
    document.getElementById("example").value="custom";
    document.getElementById("totalCost").value="";
    document.getElementById("months").value="12";
    document.getElementById("isMultiple").checked=false;
    document.getElementById("outpatientOnly").checked=false;
    recalc();
  });


  // 総医療費：カンマ区切り表示（入力中は邪魔しない・フォーカスアウトで整形）
  totalCostInput.addEventListener("focus", ()=>{
    totalCostInput.value = String(totalCostInput.value || "").replace(/,/g,"");
  });
  totalCostInput.addEventListener("blur", ()=>{
    const v = parseYenInput(totalCostInput.value);
    totalCostInput.value = formatWithCommas(v);
    recalc();
  });
  ["income","totalCost","months","isMultiple","copayRate","outpatientOnly"].forEach(id=>{
    document.getElementById(id).addEventListener("input", recalc);
    document.getElementById(id).addEventListener("change", recalc);
  });

  function recalc(){
    const income = document.getElementById("income").value;
    let total = parseYenInput(totalCostInput.value);

    // AUTO-FILL: 総医療費が空/0なら、選択中の治療例（病名）の金額を採用
    const exKey = document.getElementById("example").value;
    const exPicked = EXAMPLES.find(x=>x.key===exKey);
    if ((!Number.isFinite(total) || total === 0) && exPicked && exPicked.monthlyTotal && exKey !== "custom"){
      total = exPicked.monthlyTotal;
      totalCostInput.value = formatWithCommas(total);
    }
const months = Math.min(12, Math.max(1, Number(document.getElementById("months").value || 12)));
    document.getElementById("months").value = months;

    const copayRate = Number(document.getElementById("copayRate").value || 0.3);
    const copay = total * copayRate;

    document.getElementById("copayRaw").textContent = yen(copay);
    document.getElementById("copayRawInput").value = yen(copay);
    document.getElementById("pickedIncome").textContent = pickedIncomeLabel();
    document.getElementById("pickedMonths").textContent = months + "か月";

    const age = document.getElementById("ageGroup").value;
    const isMultipleEl = document.getElementById("isMultiple");
    const outpatientOnly = document.getElementById("outpatientOnly").checked;
    // Badge / Note toggle (外来特例の説明)
    
    // === Mode tabs (外来のみ / 外来+入院+手術 / 入院+手術) ===
    const outpatientOnlyEl = document.getElementById("outpatientOnly");
    const modeNoteEl = document.getElementById("modeNote");
    const thOutGroup = document.getElementById("thOutpatientGroup");
    const thInGroup = document.getElementById("thInpatientGroup");
    const modeTabs = Array.from(document.querySelectorAll(".modeTab"));

    // Derive currentMode from checkbox (backward-compatible)
    let currentMode = outpatientOnlyEl && outpatientOnlyEl.checked ? "outpatient" : "combo";
    // If tabs exist, prefer the active one
    const activeBtn = modeTabs.find(b => b.classList.contains("active"));
    if (activeBtn && activeBtn.dataset && activeBtn.dataset.mode) currentMode = activeBtn.dataset.mode;

    function applyModeUI(mode){
      modeTabs.forEach(btn=>{
        const on = btn.dataset.mode === mode;
        btn.classList.toggle("active", on);
        btn.setAttribute("aria-selected", on ? "true" : "false");
      });

      if (thOutGroup) thOutGroup.textContent = "外来のみ（外来特例）";
      if (thInGroup){
        thInGroup.textContent =
          mode === "combo" ? "外来＋入院＋手術" :
          mode === "inpatient" ? "入院＋手術" :
          "入院＋手術（参考）";
      }

      if (modeNoteEl){
        if (mode === "outpatient"){
          modeNoteEl.innerHTML = '※外来特例（70歳以上・外来のみ・個人単位）の上限を優先して試算します。<br>※外来のみのため、右側（入院＋手術）は参考表示です。';
        }else if (mode === "combo"){
          modeNoteEl.innerHTML = '※外来特例（70歳以上・外来のみ・個人単位）は、入院を伴う高額療養費とは合算されません。外来＋入院＋手術の合算を想定する場合は、右側（外来＋入院＋手術）で試算します。';
        }else{
          modeNoteEl.innerHTML = '※入院＋手術（高額療養費）を想定する場合は、右側（入院＋手術）で試算します。外来特例（外来のみ）は適用しません。';
        }
      }
    }

    // Wire tab clicks (do once)
    if (!window.__modeTabsWired){
      modeTabs.forEach(btn=>{
        btn.addEventListener("click", ()=>{
          if (btn.classList.contains("disabled")) return;
          const mode = btn.dataset.mode;
          if (!mode) return;
          if (outpatientOnlyEl) outpatientOnlyEl.checked = (mode === "outpatient"); // keep calculation logic
          applyModeUI(mode);
          recalc();
        });
      });
      window.__modeTabsWired = true;
    }

    applyModeUI(currentMode);

// Table dimming + unavailable note
    const resultTable = document.getElementById("resultTable");
    const noteUnavailable = document.getElementById("noteOutpatientUnavailable");

    // eligibility: 70+ and selected income group has outpatient cap in 2025 rules (same key is used for all years in this simulator)
    let hasOutCap = false;
    try{
      const key = document.getElementById("income").value;
      const r0 = (RULES && RULES[2025] && RULES[2025][key]) ? RULES[2025][key] : null;
      hasOutCap = !!(r0 && r0.outpatient !== null && r0.outpatient !== undefined);
    }catch(e){ hasOutCap = false; }

    
    // Disable "外来のみ（外来特例）" tab when not applicable (e.g., 70歳未満)
    const outpatientTabBtn = modeTabs.find(b => b.dataset && b.dataset.mode === "outpatient");
    const canSelectOutTab = (age === "o70") && hasOutCap;
    if (outpatientTabBtn){
      outpatientTabBtn.classList.toggle("disabled", !canSelectOutTab);
      outpatientTabBtn.setAttribute("aria-disabled", !canSelectOutTab ? "true" : "false");
      outpatientTabBtn.title = !canSelectOutTab ? "外来特例は70歳以上かつ対象区分のみ選択できます" : "";
    }
    // If currently on outpatient mode but it became unavailable, fall back to combo
    if (!canSelectOutTab && currentMode === "outpatient"){
      if (outpatientOnlyEl) outpatientOnlyEl.checked = false;
      currentMode = "combo";
      applyModeUI("combo");
    }

    const eligibleOut = (age === "o70") && outpatientOnly && hasOutCap;

    if (resultTable){
      if (eligibleOut) resultTable.classList.add("dim-inpatient");
      else resultTable.classList.remove("dim-inpatient");
    }
    if (noteUnavailable){
      noteUnavailable.style.display = (outpatientOnly && !eligibleOut) ? "block" : "none";
    }

    // 外来特例（70+・外来のみ）には多数回該当を適用しない
    if (age === "o70" && outpatientOnly){
      isMultipleEl.checked = false;
      isMultipleEl.disabled = true;
    } else {
      isMultipleEl.disabled = false;
    }
    const isMultiple = isMultipleEl.checked;

    const years = [2025,2026,2027];
    const tbody = document.getElementById("resultBody");
    tbody.innerHTML = "";

    const yearRows = [];
    years.forEach(y=>{
      const rule = RULES[y][income];
      const monthlyLimit = calcMonthlyLimit(rule, total, isMultiple);

      // outpatient cap (70+ & outpatientOnly)
      let outpatientCap = null;
      if (age==="o70" && outpatientOnly && rule && rule.outpatient !== null && rule.outpatient !== undefined){
        outpatientCap = rule.outpatient;
      }

      const effectiveCap = (outpatientCap !== null && outpatientCap !== undefined) ? outpatientCap : monthlyLimit;
      const monthlyOop = (effectiveCap !== null && effectiveCap !== undefined) ? Math.min(copay, effectiveCap) : copay;

      const annualCap = (rule && rule.annual !== null && rule.annual !== undefined) ? rule.annual : null;
      const annualRaw = monthlyOop * months;

      // 外来特例（70+）は年額上限（例：年21.6万円）を優先適用
      const outpatientAnnualCap = (outpatientCap !== null && outpatientCap !== undefined)
        ? ((rule && rule.outpatientAnnual !== null && rule.outpatientAnnual !== undefined) ? rule.outpatientAnnual : (outpatientCap*12))
        : null;

      
      // Text for outpatient cap column (only when eligible outpatient cap exists)
      const outpatientText = (outpatientCap !== null && outpatientCap !== undefined)
        ? `月 ${yen(outpatientCap)}（年 ${yen(outpatientAnnualCap)}）`
        : "—";
const annualOop = (outpatientAnnualCap !== null && outpatientAnnualCap !== undefined)
        ? Math.min(annualRaw, outpatientAnnualCap)
        : (annualCap ? Math.min(annualRaw, annualCap) : annualRaw);

      const capText = `<div>${ruleLogicText(rule)}</div>`;


      yearRows.push({
        year: y,
        outpatientCapText: outpatientText,
        outpatientMonthlyCap: outpatientCap ?? null,
        outpatientAnnualCap: outpatientAnnualCap ?? null,
        inpatientRuleText: ruleLogicText(rule),
        inpatientAnnualCap: annualCap ?? null,
        inpatientMonthlyOop: monthlyOop ?? null,
        inpatientAnnualOop: annualOop ?? null
      });

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td data-label="年"><span class="val"><b>${y}</b></span></td>
        <td data-label="外来特例上限（月額／年額）" class="mono"><span class="val"><b>${outpatientText}</b></span></td>
        <td data-label="入院 月額上限（ロジック）"><span class="val">${capText}</span></td>
        <td data-label="入院 年間上限" class="mono"><span class="val"><b>${annualCap?yen(annualCap):"—"}</b></span></td>
        <td data-label="入院 月額自己負担" class="mono"><span class="val"><b>${yen(monthlyOop)}</b></span></td>
        <td data-label="入院 年額自己負担" class="mono"><span class="val"><b>${yen(annualOop)}</b></span></td>
      `;
      tbody.appendChild(tr);
    });

    // Mobile: 年度タブ + カード表示（横スクロール回避）
    window.__yearRows = yearRows;
    window.__currentMode = currentMode;
    if (!window.__activeYear) window.__activeYear = 2025;
    renderYearTabs(yearRows, currentMode);
    renderMobileYearCard(yearRows, window.__activeYear, currentMode);
  }

  // initial
  loadIncomeOptions();
  totalCostInput.value = formatWithCommas(parseYenInput(totalCostInput.value));
  recalc();

// --- Comma formatting for 総医療費 input (UI only; calc uses comma-stripped value) ---
function formatNumberWithComma(value){
  if(value===null||value===undefined) return '';
  const num = Number(String(value).replace(/,/g,''));
  if(Number.isNaN(num)) return '';
  return num.toLocaleString('ja-JP');
}
function removeComma(value){
  return String(value||'').replace(/,/g,'');
}

  function isMobileView(){
    return window.matchMedia && window.matchMedia('(max-width: 720px)').matches;
  }

  function renderYearTabs(yearRows, currentMode){
    const tabWrap = document.getElementById('yearTabs');
    if(!tabWrap) return;
    tabWrap.innerHTML = '';
    yearRows.forEach(r=>{
      const b = document.createElement('button');
      b.type = 'button';
      b.textContent = String(r.year);
      b.setAttribute('role','tab');
      b.setAttribute('aria-selected', String(window.__activeYear===r.year));
      b.addEventListener('click', ()=>{
        window.__activeYear = r.year;
        renderYearTabs(window.__yearRows||yearRows, window.__currentMode||currentMode);
        renderMobileYearCard(window.__yearRows||yearRows, window.__activeYear, window.__currentMode||currentMode);
      });
      tabWrap.appendChild(b);
    });
    [...tabWrap.querySelectorAll('button')].forEach(btn=>{
      btn.setAttribute('aria-selected', String(Number(btn.textContent)===window.__activeYear));
    });
  }

  function renderMobileYearCard(yearRows, year, currentMode){
    const box = document.getElementById('mobileYearCard');
    if(!box) return;
    const r = yearRows.find(x=>x.year===year) || yearRows[0];
    if(!r){
      box.innerHTML = '';
      return;
    }

    const parts = [];

    // 外来（外来特例）
    parts.push(`
      <div class="ycard">
        <h4>外来（70歳以上・外来特例）</h4>
        <div class="kv">
          <div class="row"><div class="k">外来特例上限（月額／年額）</div><div class="v">${r.outpatientCapText || '—'}</div></div>
        </div>
      </div>
    `);

    // 入院（高額療養費）
    if(currentMode !== 'outpatient'){
      parts.push(`
        <div class="ycard">
          <h4>${currentMode === 'combo' ? '外来＋入院＋手術（高額療養費）' : '入院＋手術（高額療養費）'}</h4>
          <div class="kv">
            <div class="row"><div class="k">月額上限（ロジック）</div><div class="v">${r.inpatientRuleText || '—'}</div></div>
            <div class="row"><div class="k">年間上限</div><div class="v">${r.inpatientAnnualCap!=null ? yen(r.inpatientAnnualCap) : '—'}</div></div>
            <div class="row"><div class="k">月額自己負担</div><div class="v">${r.inpatientMonthlyOop!=null ? yen(r.inpatientMonthlyOop) : '—'}</div></div>
            <div class="row"><div class="k">年間自己負担</div><div class="v">${r.inpatientAnnualOop!=null ? yen(r.inpatientAnnualOop) : '—'}</div></div>
          </div>
        </div>
      `);
    }

    box.innerHTML = parts.join('');

    if(!isMobileView()){
      box.style.display = 'none';
    }else{
      box.style.display = 'block';
    }
  }

  (function bindMobileRerender(){
    if(window.__mobileBindDone) return;
    window.__mobileBindDone = true;
    window.addEventListener('resize', ()=>{
      if(window.__yearRows){
        renderYearTabs(window.__yearRows, window.__currentMode||'combo');
        renderMobileYearCard(window.__yearRows, window.__activeYear||2025, window.__currentMode||'combo');
      }
    }, {passive:true});
  })();

function getTotalCostValue(){
  const el = document.getElementById('totalCost');
  if(!el) return 0;
  return Number(removeComma(el.value));
}
(function attachTotalCostFormat(){
  const el = document.getElementById('totalCost');
  if(!el) return;
  el.addEventListener('input', () => {
    el.value = el.value.replace(/[^\d,]/g,'');
  });
  el.addEventListener('blur', () => {
    el.value = formatNumberWithComma(el.value);
  });
})();



/* ===== Patch: stable rendering for result table (v17) ===== */
function fmtYen(n){
  if(!isFinite(n)) return "—";
  return Math.round(n).toLocaleString("ja-JP") + "円";
}
function getRulesByYear(y){
  // RULES is a map keyed by year (e.g., 2025/2026/2027)
  return RULES[y] ?? RULES[String(y)] ?? null;
}
function getRuleForYearIncome(y, incomeCode){
  const rules = getRulesByYear(y);
  return rules[incomeCode] || null;
}

  // 外来特例（70歳以上・外来のみ）の月額上限（年別）
  // 対象外の場合は null
  function outpatientCapFor(year, incomeKey){
    const rule = getRuleForYearIncome(year, incomeKey);
    if(!rule || !rule.outpatient) return null;
    return rule.outpatient;
  }


function calcMonthlyCap(rule, totalCost){
  if(!rule) return null;
  if(rule.type==="fixed") return rule.cap;
  if(rule.type==="formula"){
    return rule.base + Math.floor(rule.rate * Math.max(0, totalCost - rule.threshold));
  }
  return null;
}
function calcMonthlyCapMulti(rule, totalCost){
  if(!rule) return null;
  if(rule.type==="fixed") return rule.multi ?? rule.cap;
  if(rule.type==="formula"){
    // 多数回該当は固定値（rule.multi）を優先
    return (rule.multi ?? calcMonthlyCap(rule,totalCost));
  }
  return null;
}
