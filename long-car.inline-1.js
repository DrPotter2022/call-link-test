/* ========= Helpers ========= */
const $ = (id)=>document.getElementById(id);
const clamp = (x,a,b)=>Math.min(Math.max(x,a),b);

function fmtYYMMDD(d){
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const dd = String(d.getDate()).padStart(2,'0');
  return `${yy}/${mm}/${dd}`;
}

// yy/mm/dd, yyyy/mm/dd, yymmdd, yyyymmdd
function parseFlexDate(s){
  if(!s) return null;
  const t = String(s).trim();
  let y,m,d;
  if(/^\d{2}\/\d{2}\/\d{2}$/.test(t)){
    const [yy,mm,dd]=t.split('/').map(Number);
    y = 2000 + yy; m=mm; d=dd;
  }else if(/^\d{4}\/\d{2}\/\d{2}$/.test(t)){
    const [yyyy,mm,dd]=t.split('/').map(Number);
    y=yyyy; m=mm; d=dd;
  }else if(/^\d{6}$/.test(t)){
    y = 2000 + Number(t.slice(0,2));
    m = Number(t.slice(2,4));
    d = Number(t.slice(4,6));
  }else if(/^\d{8}$/.test(t)){
    y = Number(t.slice(0,4));
    m = Number(t.slice(4,6));
    d = Number(t.slice(6,8));
  }else{
    return null;
  }
  if(!y||m<1||m>12||d<1||d>31) return null;
  const dt = new Date(y, m-1, d);
  if(dt.getFullYear()!==y || dt.getMonth()!==m-1 || dt.getDate()!==d) return null;
  return dt;
}
function addYears(date, years){
  const d = new Date(date);
  const m = d.getMonth();
  d.setFullYear(d.getFullYear() + years);
  if(d.getMonth() !== m){
    d.setDate(0);
  }
  return d;
}

/* ========= 等級別割引増率テーブル（目安） ========= */
/* 注意：実務は保険会社・商品で異なる可能性があります。ここではWeb参照の目安値として利用。 */
const rateNoAcc = {
  20:-63,19:-50,18:-46,17:-44,16:-42,15:-41,14:-40,13:-38,12:-37,11:-35,
  10:-33, 9:-31, 8:-29, 7:-27, 6:-13, 5: -2, 4: +7, 3:+23, 2:+49, 1:+64
};
// 事故有（目安）：無事故から+15%相当で近似（例：20等級 -63 → -44）
const rateAcc = {
  20:-44,19:-34,18:-33,17:-32,16:-31,15:-30,14:-29,13:-28,12:-27,11:-26,
  10:-25, 9:-24, 8:-23, 7:-22, 6:-15, 5: -5, 4:+13, 3:+32, 2:+58, 1:+64
};

function getRate(grade, coefYears){
  const g = clamp(grade,1,20);
  const isAcc = coefYears>=1;
  return isAcc ? rateAcc[g] : rateNoAcc[g];
}

/* ========= 事故区分 ========= */
let accidents = []; // {type:"D3"/"D1"/"NC", dateText:"", memo:""}
const typeMeta = {
  D3: { label:"3等級ダウン", tip:"3等級ダウン：事故有+3年、等級-3（例：車同士の事故等）", coefAdd:3, gradeDown:3 },
  D1: { label:"1等級ダウン", tip:"1等級ダウン：事故有+1年、等級-1（例：飛び石等）", coefAdd:1, gradeDown:1 },
  NC: { label:"ノーカウント", tip:"ノーカウント：等級・事故有係数に影響なし", coefAdd:0, gradeDown:0 },
};

function makeAccRow(item, idx){
  const wrap = document.createElement("div");
  wrap.className="accItem";

  // type
  const typeBox = document.createElement("div");
  typeBox.className="accType";
  const sel = document.createElement("select");
  [["D3","3等級ダウン"],["D1","1等級ダウン"],["NC","ノーカウント"]].forEach(([v,txt])=>{
    const o=document.createElement("option");
    o.value=v; o.textContent=txt;
    sel.appendChild(o);
  });
  sel.value = item.type;
  sel.addEventListener("change", ()=>{
    item.type = sel.value;
    renderAccidents();
    compute();
  });
  typeBox.appendChild(sel);

  // info
  const info = document.createElement("button");
  info.type="button";
  info.className="infoBtn";
  info.textContent="i";
  info.addEventListener("click",(e)=>{
    e.stopPropagation();
    const meta = typeMeta[item.type];
    const r = info.getBoundingClientRect();
    showTip(r.left, r.top, meta.tip);
  });

  // date (text + date picker)
  const dateWrap = document.createElement("div");
  dateWrap.className="accDate";

  const dateText = document.createElement("input");
  dateText.type="text";
  dateText.placeholder="事故日";
  dateText.value = item.dateText || "";
  dateText.addEventListener("blur", ()=>{
    const d = parseFlexDate(dateText.value);
    if(d){
      dateText.value = fmtYYMMDD(d);
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth()+1).padStart(2,'0');
      const dd = String(d.getDate()).padStart(2,'0');
      datePick.value = `${yyyy}-${mm}-${dd}`;
    }
    item.dateText = dateText.value;
  });

  const dateBtn = document.createElement("div");
  dateBtn.className="dateBtn";
  dateBtn.title="カレンダー";
  const datePick = document.createElement("input");
  datePick.type="date";
  // init from text if possible
  const d0 = parseFlexDate(dateText.value);
  if(d0){
    const yyyy=d0.getFullYear();
    const mm=String(d0.getMonth()+1).padStart(2,'0');
    const dd=String(d0.getDate()).padStart(2,'0');
    datePick.value = `${yyyy}-${mm}-${dd}`;
  }
  datePick.addEventListener("change", ()=>{
    if(!datePick.value) return;
    const d = new Date(datePick.value);
    dateText.value = fmtYYMMDD(d);
    item.dateText = dateText.value;
  });
  dateBtn.appendChild(datePick);

  // クリックで確実にカレンダーを開く（対応ブラウザでは showPicker）
  dateBtn.addEventListener('click', (ev)=>{
    ev.preventDefault();
    try{
      if (typeof datePick.showPicker === 'function') datePick.showPicker();
    }catch(_){}
    datePick.focus();
    datePick.click();
  });

  dateWrap.appendChild(dateText);
  dateWrap.appendChild(dateBtn);

  // memo
  const memo = document.createElement("div");
  memo.className="accMemo";
  const memoIn = document.createElement("input");
  memoIn.type="text";
  memoIn.placeholder="メモ（任意：例）車両";
  memoIn.value = item.memo || "";
  memoIn.addEventListener("input", ()=>{
    item.memo = memoIn.value;
  });
  memo.appendChild(memoIn);

  // del
  const del = document.createElement("button");
  del.type="button";
  del.className="delBtn";
  del.textContent="×";
  del.addEventListener("click", ()=>{
    accidents.splice(idx,1);
    renderAccidents();
    compute();
  });

  wrap.appendChild(typeBox);
  wrap.appendChild(info);
  wrap.appendChild(dateWrap);
  wrap.appendChild(memo);
  wrap.appendChild(del);
  return wrap;
}

function renderAccidents(){
  const list = $("accList");
  list.innerHTML = "";
  accidents.forEach((it,i)=>list.appendChild(makeAccRow(it,i)));

  let c3=0,c1=0,cn=0;
  for(const a of accidents){
    if(a.type==="D3") c3++;
    else if(a.type==="D1") c1++;
    else cn++;
  }
  $("accSummary").textContent = `自動集計：3等級ダウン件数=${c3} / 1等級ダウン件数=${c1} / ノーカウント=${cn}`;
}

/* ========= Tip popup ========= */
const tipPop = $("tipPop");
function showTip(x,y,text){
  if(!text) return;
  tipPop.textContent = text;
  tipPop.style.display="block";
  const pad=12;
  const w = 320;
  const h = 120;
  let left = x + 10;
  let top = y + 10;
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  if(left + w > vw - pad) left = vw - w - pad;
  if(top + h > vh - pad) top = vh - h - pad;
  tipPop.style.left = left+"px";
  tipPop.style.top = top+"px";
}
document.addEventListener("click", ()=> tipPop.style.display="none");

/* ========= 保険始期日 ========= */
(function initStartDate(){
  const t = $("startDateText");
  const p = $("startDatePick");
  const btn = $("startDateBtn");
  if(btn){
    btn.addEventListener("click", ()=>{
      try{ p.showPicker && p.showPicker(); }catch(_){}
      p.focus(); p.click();
    });
  }
  p.addEventListener("change", ()=>{
    if(!p.value) return;
    const d = new Date(p.value);
    t.value = fmtYYMMDD(d);
  });
  t.addEventListener("blur", ()=>{
    const d = parseFlexDate(t.value);
    if(d){
      t.value = fmtYYMMDD(d);
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth()+1).padStart(2,'0');
      const dd = String(d.getDate()).padStart(2,'0');
      p.value = `${yyyy}-${mm}-${dd}`;
    }
  });
})();

/* ========= 契約年数 ========= */
function getTermYears(){
  const v = $("term").value;
  const n = parseInt(v,10);
  return Number.isFinite(n) ? n : 1;
}

/* ========= 計算 ========= */
function compute(){
  const gradeNow = parseInt($("gradeNow").value,10);
  const coefNow = parseInt($("coefNow").value,10);
  const years = getTermYears();

  let cnt3=0,cnt1=0;
  for(const a of accidents){
    if(a.type==="D3") cnt3++;
    if(a.type==="D1") cnt1++;
  }
  const downgr = 3*cnt3 + 1*cnt1;
  const addCoef = 3*cnt3 + 1*cnt1;

  const addYearsForGrade = Math.max(0, years - 1);
  let outGrade = clamp(Math.floor(gradeNow + addYearsForGrade - downgr), 1, 20);
  let outCoef = clamp(coefNow + addCoef - addYearsForGrade, 0, 6);

  const hadAcc = (cnt3 + cnt1) > 0;
  let refGrade = hadAcc ? clamp(Math.floor(gradeNow - downgr), 1, 20)
                        : clamp(Math.floor(gradeNow + 1), 1, 20);
  let refCoef = clamp(coefNow + addCoef - 1, 0, 6);
  if(!hadAcc){ refCoef = clamp(coefNow - 1, 0, 6); }

  $("outGrade").textContent = outGrade;
  $("outCoef").textContent = outCoef + "年";

  const outR = getRate(outGrade, outCoef);
  $("outRate").textContent = `割引増率：${outR}%`;

  $("refGrade").textContent = refGrade;
  $("refCoef").textContent = refCoef;
  const refR = getRate(refGrade, refCoef);
  $("refRate").textContent = `割引増率：${refR}%`;

  const bNow = coefNow>=1 ? ["事故有","bad"] : ["無事故","good"];
  $("badgeNow").textContent=bNow[0];
  $("badgeNow").className="badge "+bNow[1];

  const bOut = outCoef>=1 ? ["事故有","bad"] : ["無事故","good"];
  $("badgeOut").textContent=bOut[0];
  $("badgeOut").className="badge "+bOut[1];

  const sd = parseFlexDate($("startDateText").value);
  if(sd){
    const nd = addYears(sd, years);
    const yyyy = nd.getFullYear();
    const mm = String(nd.getMonth()+1).padStart(2,'0');
    const dd = String(nd.getDate()).padStart(2,'0');
    $("nextDate").textContent = `次回更改日：${yyyy}/${mm}/${dd}`;
  }else{
    $("nextDate").textContent = "次回更改日：-";
  }

  const refRes = sumUntilZero(refGrade, refCoef);
  const N = refRes.years;
  const mainRes = sumForYears(outGrade, outCoef, N);

  $("sumMain").textContent = `合計${fmtPct(mainRes.sum)}`;
  $("sumMainList").textContent = mainRes.lines.join("\n");

  $("sumRef").textContent = `合計${fmtPct(refRes.sum)}`;
  $("sumRefList").textContent = refRes.lines.join("\n");

  const diff = mainRes.sum - refRes.sum;
  $("sumDiff").textContent = fmtPct(diff);
  $("sumDiffMsg").textContent = diffMessage(diff);

  const outLabel = `主表示：等級${outGrade}/${outCoef>=1?"事故有":"無事故"}`;
  const refLabel = `参考：等級${refGrade}/${refCoef>=1?"事故有":"無事故"}`;
  $("chartHint").textContent = `${outLabel} / ${refLabel}`;

  renderChart(outGrade, outCoef, refGrade, refCoef);

  $("msg").textContent = "計算しました。";
}

function fmtPct(x){
  const v = Math.round(x);
  if(v===0) return "0%";
  return v>0 ? `+${v}%` : `${v}%`;
}
function fmtPctAbs(x){
  const v = Math.round(Math.abs(x));
  return `+${v}%`;
}
function diffMessage(diff){
  const v = Math.round(diff);
  if(v===0) return "差：同等（0%）";
  if(v<0) return `差：主表示が ${fmtPctAbs(v)} 有利`;
  return `差：参考（1年更新）が ${fmtPctAbs(v)} 有利`;
}

function sumUntilZero(startGrade, startCoef){
  let g = startGrade;
  let c = startCoef;
  let sum = 0;
  const lines = [];
  let years = 0;
  for(let year=1; year<=6; year++){
    if(c<=0) break;
    years = year;
    const r = getRate(g, c);
    sum += r;
    lines.push(`年${year}：等級${g} / ${c>=1?"事故有":"無事故"} / ${r}%`);
    g = clamp(g+1,1,20);
    c = clamp(c-1,0,6);
  }
  return { sum, lines, years };
}

function sumForYears(startGrade, startCoef, years){
  let g = startGrade;
  let c = startCoef;
  let sum = 0;
  const lines = [];
  for(let year=1; year<=years; year++){
    const r = getRate(g, c);
    sum += r;
    lines.push(`年${year}：等級${g} / ${c>=1?"事故有":"無事故"} / ${r}%`);
    g = clamp(g+1,1,20);
    c = clamp(c-1,0,6);
  }
  return { sum, lines, years };
}

/* ========= hover tooltip ========= */
const hoverTip = $("hoverTip");
function showHover(x,y, html){
  hoverTip.innerHTML = html;
  hoverTip.style.display="block";
  const pad=12;
  const vw=innerWidth, vh=innerHeight;
  let left = x + 12;
  let top  = y + 12;
  const rectW = 280, rectH = 120;
  if(left + rectW > vw - pad) left = vw - rectW - pad;
  if(top + rectH > vh - pad) top = vh - rectH - pad;
  hoverTip.style.left = left+"px";
  hoverTip.style.top = top+"px";
}
function hideHover(){ hoverTip.style.display="none"; }

/* ========= グラフ ========= */
function widthFromRate(r){
  const maxAbs = 70;
  const a = Math.abs(r);
  const pct = (a / maxAbs) * 49;
  return clamp(pct, 0, 49);
}

function renderChart(mainG, mainC, refG, refC){
  const host = $("nfChart");
  host.innerHTML = "";

  for(let g=20; g>=1; g--){
    const row = document.createElement("div");
    row.className="nfRow";

    const y = document.createElement("div");
    y.className="nfY";
    y.textContent = g;

    const rail = document.createElement("div");
    rail.className="nfRail";

    const center = document.createElement("div");
    center.className="nfCenter";
    rail.appendChild(center);

    const rNo = getRate(g, 0);
    const rYes = getRate(g, 1);

    const barNo = document.createElement("div");
    barNo.className = "bar good " + (rNo<0 ? "neg" : "pos");
    barNo.style.width = widthFromRate(rNo) + "%";

    const barYes = document.createElement("div");
    barYes.className = "bar bad " + (rYes<0 ? "neg" : "pos");
    barYes.style.width = widthFromRate(rYes) + "%";

    const mainIsAcc = mainC >= 1;
    const refIsAcc  = refC >= 1;
    if(g===mainG){ (mainIsAcc ? barYes : barNo).classList.add("blink"); }
    if(g===refG){ (refIsAcc ? barYes : barNo).classList.add("blink"); }

    rail.addEventListener("mousemove",(e)=>{
      const html = `
        <div style="font-weight:900;margin-bottom:6px;">等級${g}</div>
        <div>事故なし：${rNo}%</div>
        <div>事故あり：${rYes}%</div>
      `;
      showHover(e.clientX, e.clientY, html);
    });
    rail.addEventListener("mouseleave", hideHover);

    rail.appendChild(barNo);
    rail.appendChild(barYes);

    row.appendChild(y);
    row.appendChild(rail);
    host.appendChild(row);
  }
}

/* ========= イベント ========= */
$("btnAdd").addEventListener("click", ()=>{
  accidents.push({type:"D3", dateText:"", memo:""});
  renderAccidents();
  compute();
});
$("btnClear").addEventListener("click", ()=>{
  accidents = [];
  renderAccidents();
  compute();
});

$("btnCalc").addEventListener("click", compute);

$("btnReset").addEventListener("click", ()=>{
  $("gradeNow").value="20";
  $("coefNow").value="0";
  $("term").value="1";
  $("startDateText").value="";
  $("startDatePick").value="";
  accidents=[];
  renderAccidents();
  $("msg").textContent="";
  compute();
});

$("gradeNow").addEventListener("change", compute);
$("coefNow").addEventListener("change", compute);
$("term").addEventListener("change", compute);
$("startDateText").addEventListener("input", ()=>{$("msg").textContent="";});

/* ========= 初期描画 ========= */
renderAccidents();
compute();

/* ========= 同意モーダル ========= */
(function consent(){
  const modal = $("modal");
  const root = $("appRoot");

  function lock(on){
    if(on) root.classList.add("disabledAll");
    else root.classList.remove("disabledAll");
  }

  modal.classList.add("on");
  lock(true);

  $("btnAgree").addEventListener("click", ()=>{
    modal.classList.remove("on");
    lock(false);
  });
})();
