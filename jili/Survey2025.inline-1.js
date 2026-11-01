(() => {
'use strict';
const DATA = JSON.parse(document.getElementById('embedded-data').textContent);
const $ = id => document.getElementById(id);
const chapterSel=$('chapterSelect'), qSel=$('questionSelect'), sheetChecks=$('sheetChecks'), answerChecks=$('answerChecks'), checks=$('checks'), filterText=$('filterText'), paletteSelect=$('paletteSelect');
const chart=$('chart'), tooltip=$('tooltip'), table=$('dataTable');
let sortMode=false;
let sheetOrder=[], answerOrder=[];
const hiddenSeries=new Set();
const palettes={default:['#1f6feb','#0a7f5a','#d97706','#7c3aed','#dc2626','#0891b2','#4f46e5','#be123c','#65a30d','#9333ea','#0369a1','#b45309'],blue:['#dbeafe','#bfdbfe','#93c5fd','#60a5fa','#3b82f6','#2563eb','#1d4ed8','#1e40af','#1e3a8a'],green:['#dcfce7','#bbf7d0','#86efac','#4ade80','#22c55e','#16a34a','#15803d','#166534','#14532d'],warm:['#ffedd5','#fed7aa','#fdba74','#fb923c','#f97316','#ea580c','#dc2626','#b91c1c','#991b1b'],purple:['#f3e8ff','#e9d5ff','#d8b4fe','#c084fc','#a855f7','#9333ea','#7e22ce','#6b21a8','#581c87'],mono:['#f3f4f6','#e5e7eb','#d1d5db','#9ca3af','#6b7280','#4b5563','#374151','#1f2937','#111827'],vivid:['#2563eb','#f97316','#22c55e','#ef4444','#a855f7','#06b6d4','#eab308','#ec4899','#14b8a6','#64748b']};
const CHAPTERS={
  '1':'第Ⅰ章　生活設計と生活保障意識',
  '2':'第Ⅱ章　医療保障',
  '3':'第Ⅲ章　老後保障',
  '4':'第Ⅳ章　死亡保障',
  '5':'第Ⅴ章　介護保障',
  '6':'第Ⅵ章　生命保険の加入状況',
  '7':'第Ⅶ章　直近加入契約の状況と今後の加入意向',
  '8':'補章',
  '9':'属性'
};
function chapterOf(q){return String(q?.id||'').split('-')[0];}

function activeColors(){return palettes[(paletteSelect&&paletteSelect.value)||'default']||palettes.default;}
function updatePalettePreview(){/* カラーサンプルは非表示 */}
function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}
function fmt(v){return (typeof v==='number' && isFinite(v)) ? (Math.round(v*10)/10).toFixed(1) : String(v??'');}
function currentQ(){const idx=+qSel.value; return DATA.questions[idx] || filteredQuestions()[0]?.q || DATA.questions[0];}
function filteredQuestions(){const ch=chapterSel.value; return DATA.questions.map((q,i)=>({q,i})).filter(x=>!ch || chapterOf(x.q)===ch);}
function populateChapters(){const present=[...new Set(DATA.questions.map(chapterOf))].filter(Boolean).sort((a,b)=>+a-+b); chapterSel.innerHTML='<option value="">すべて</option>'+present.map(k=>`<option value="${esc(k)}">${esc(k)}-${esc(CHAPTERS[k]||('第'+k+'章'))}</option>`).join('');}
function allAnswersForSheets(sheets){const set=new Set(); sheets.forEach(s=>s.rows.forEach(r=>Object.keys(r.values||{}).forEach(k=>set.add(k)))); return [...set];}
function selectedSheetIndexes(){return [...sheetChecks.querySelectorAll('input:checked')].map(i=>+i.value);}
function selectedSheets(){const q=currentQ(); return selectedSheetIndexes().map(i=>q.sheets[i]).filter(Boolean);}
function selectedAnswers(){return [...answerChecks.querySelectorAll('input:checked')].map(i=>i.value);}
function updateOrder(container, order, isNumber=false){const checked=[...container.querySelectorAll('input:checked')].map(i=>isNumber?+i.value:i.value); for(const v of checked){if(!order.includes(v)) order.push(v);} return order.filter(v=>checked.includes(v));}
function populateQuestions(){const items=filteredQuestions(); qSel.innerHTML=items.map(({q,i})=>`<option value="${i}">${esc(q.id)}：${esc(q.title)}</option>`).join(''); if(!items.length){qSel.innerHTML='<option value="0">該当する設問がありません</option>';} sortMode=false; hiddenSeries.clear(); populateSheets();}
function populateSheets(){const q=currentQ(); sheetOrder=[]; sheetChecks.innerHTML=q.sheets.map((s,i)=>`<label class="check"><input type="checkbox" value="${i}" ${i===0?'checked':''}> <span>${esc(s.sheet)} <span class="small">${s.rows.length}属性</span></span></label>`).join(''); sheetOrder=[0]; populateAnswers();}
function populateAnswers(){const sheets=selectedSheets(); const ans=allAnswersForSheets(sheets.length?sheets:[currentQ().sheets[0]]); const prev=new Set(answerOrder.length?answerOrder:ans.slice(0,1)); answerChecks.innerHTML=ans.map((a,i)=>`<label class="check"><input type="checkbox" value="${esc(a)}" ${prev.has(a)||(!answerOrder.length&&i===0)?'checked':''}> <span>${esc(a)}</span></label>`).join('') || '<div class="empty">選択可能な回答項目がありません。</div>'; answerOrder=[...answerChecks.querySelectorAll('input:checked')].map(i=>i.value); populateChecks(true);}
function visibleRows(){const kw=filterText.value.trim().toLowerCase(); const map=new Map(); selectedSheets().forEach(s=>s.rows.forEach(r=>{const key=r.category+'||'+r.label; if(!kw || (r.category+' '+r.label).toLowerCase().includes(kw)){ if(!map.has(key)) map.set(key,{category:r.category,label:r.label,n:r.n,key}); }})); return [...map.values()];}
function populateChecks(initial){const rows=visibleRows(); const prev=new Set([...checks.querySelectorAll('input:checked')].map(i=>i.value)); let html='', last=''; rows.forEach((r,idx)=>{const key=r.key; if(r.category!==last){html+=`<div class="section">${esc(r.category)}</div>`; last=r.category;} const checked=initial ? idx===0 : prev.has(key); html+=`<label class="check"><input type="checkbox" value="${esc(key)}" ${checked?'checked':''}> <span>${esc(r.label)} <span class="small">N=${esc(r.n)}</span></span></label>`;}); checks.innerHTML=html || '<div class="empty">該当する属性がありません。</div>'; render();}
function selectedAttrKeys(){return [...checks.querySelectorAll('input:checked')].map(i=>i.value);}
function makeSeries(){const q=currentQ(); const sidx=selectedSheetIndexes(); const answers=selectedAnswers(); const sOrdered=sheetOrder.filter(i=>sidx.includes(i)); const aOrdered=answerOrder.filter(a=>answers.includes(a)); const list=[]; sOrdered.forEach(si=>{aOrdered.forEach(ans=>{const sheet=q.sheets[si]; if(sheet && sheet.rows.some(r=>typeof r.values[ans]==='number')) list.push({sheetIndex:si,sheet:sheet.sheet,answer:ans,label:sheet.sheet+' / '+ans});});}); return list;}
function collectRows(){const q=currentQ(); const keys=selectedAttrKeys(); const series=makeSeries(); let rows=keys.map(key=>{const [category,label]=key.split('||'); const vals=series.map((ser,idx)=>{const row=q.sheets[ser.sheetIndex].rows.find(r=>r.category+'||'+r.label===key); return {seriesIndex:idx, value:row&&typeof row.values[ser.answer]==='number'?row.values[ser.answer]:null, n:row?row.n:null, category, label, sheet:ser.sheet, answer:ser.answer};}); return {key,category,label,values:vals};}); if(sortMode && series[0]) rows=rows.slice().sort((a,b)=>(b.values[0].value??-Infinity)-(a.values[0].value??-Infinity)); return {rows,series};}


function buildStackWarning(excludedItems){
  const items = Array.isArray(excludedItems) ? excludedItems : [];
  const details = document.createElement('details');
  details.className = 'warn autoFix autoFixWarning';
  details.open = false;

  const summary = document.createElement('summary');
  summary.innerHTML = '⚠️';
  summary.title = '積み上げ表示の自動補正';
  details.appendChild(summary);

  const p = document.createElement('p');
  p.textContent = '積み上げ表示の自動補正：選択項目の合計が100%を超えたため、重複集計項目を除外し、100%構成になる回答項目のみで表示しています。';
  details.appendChild(p);

  if(items.length){
    const title = document.createElement('div');
    title.textContent = '除外項目';
    title.style.fontWeight = '800';
    details.appendChild(title);

    const ul = document.createElement('ul');
    items.forEach(x=>{
      const li = document.createElement('li');
      li.textContent = String(x);
      ul.appendChild(li);
    });
    details.appendChild(ul);
  }
  return details;
}


function normalizeStackWarnings(){
  const candidates = Array.from(document.querySelectorAll('.warn,.warningBox,.stackWarning,.autoFixWarning'));
  candidates.forEach(node=>{
    if(node.tagName && node.tagName.toLowerCase()==='details') return;
    const txt = (node.textContent||'').trim();
    if(!txt.includes('積み上げ表示の自動補正') && !txt.includes('積み上げ表示の注意')) return;

    const excluded = [];
    const m = txt.match(/除外[:：]\s*(.+)$/);
    if(m){
      m[1].split(/[、,]/).map(s=>s.trim()).filter(Boolean).forEach(s=>excluded.push(s));
    }

    const details = buildStackWarning(excluded);
    node.replaceWith(details);
  });
}

function render(){sheetOrder=updateOrder(sheetChecks,sheetOrder,true); answerOrder=updateOrder(answerChecks,answerOrder,false); const q=currentQ(); const {rows,series}=collectRows(); $('qBadge').textContent=`${q.id} / ${series.filter(s=>!hiddenSeries.has(s.label)).length}表示 / ${series.length}系列`;  $('chartTitle').textContent=q.title; $('chartSub').textContent=`選択属性 ${rows.length}件 / 集計シート ${selectedSheetIndexes().length}件 / 回答項目 ${selectedAnswers().length}件 / データ内蔵設問 ${DATA.questions.length}件`; drawChart(rows, series); drawTable(rows, series);}
function findHundredPercentSubset(rows, visibleSeries){
  const n=visibleSeries.length;
  if(n<=1 || n>14) return null;
  const totals=rows.map(r=>r.values.reduce((a,v)=>a+(typeof v.value==='number'?v.value:0),0)).filter(v=>Number.isFinite(v));
  const hasOver=totals.some(v=>v>101);
  if(!hasOver) return null;
  let best=null;
  const totalMasks=1<<n;
  for(let mask=1; mask<totalMasks; mask++){
    const idx=[];
    for(let i=0;i<n;i++) if(mask&(1<<i)) idx.push(i);
    if(idx.length<2 || idx.length===n) continue;
    let tested=0, close=0, absSum=0, over=0;
    rows.forEach(r=>{
      let sum=0, any=false;
      idx.forEach(i=>{const val=r.values[i]?.value; if(typeof val==='number'){sum+=val; any=true;}});
      if(any){
        tested++;
        const d=Math.abs(sum-100);
        absSum+=d;
        if(d<=1.2) close++;
        if(sum>101.2) over++;
      }
    });
    if(!tested) continue;
    const closeRate=close/tested;
    if(closeRate<0.70 || over>Math.max(1, tested*0.15)) continue;
    const avgAbs=absSum/tested;
    const score=closeRate*100 - avgAbs - idx.length*0.03;
    if(!best || score>best.score) best={idx,score,closeRate,avgAbs};
  }
  return best;
}
function drawChart(rows, series){
chart.innerHTML=''; tooltip.style.display='none';
const mode=(document.querySelector('input[name="chartMode"]:checked')||{}).value||'overlay';
const colors=activeColors();
const baseVisibleSeries=series.map((s,i)=>({...s,origIndex:i})).filter(s=>!hiddenSeries.has(s.label));
const baseDisplayRows=rows.map(r=>({...r,values:r.values.filter((v,i)=>!hiddenSeries.has(series[i]?.label))}));
let visibleSeries=baseVisibleSeries;
let displayRows=baseDisplayRows;
const warnEl=$('stackWarn');
if(warnEl){warnEl.style.display='none'; warnEl.innerHTML='';}
let autoStack=false, excluded=[];
if(mode==='stack'){
  const originalSums=baseDisplayRows.map(r=>r.values.reduce((a,v)=>a+(typeof v.value==='number'?v.value:0),0));
  const hasOver=originalSums.some(v=>v>101);
  if(hasOver){
    const subset=findHundredPercentSubset(baseDisplayRows, baseVisibleSeries);
    if(subset){
      const keep=new Set(subset.idx);
      excluded=baseVisibleSeries.filter((_,i)=>!keep.has(i)).map(s=>s.label);
      visibleSeries=baseVisibleSeries.filter((_,i)=>keep.has(i));
      displayRows=baseDisplayRows.map(r=>({...r,values:r.values.filter((_,i)=>keep.has(i))}));
      autoStack=true;
      if(warnEl){
        warnEl.innerHTML=`<details class="stackIconDetails"><summary title="積み上げ表示の自動補正">⚠️</summary><div class="stackIconBody"><strong>積み上げ表示の自動補正：</strong>選択項目の合計が100%を超えたため、重複集計項目を除外し、100%構成になる回答項目のみで表示しています。<div style="font-weight:800;margin-top:6px">除外項目</div><ul>${excluded.map(x=>`<li>${esc(x)}</li>`).join('')}</ul></div></details>`;
        warnEl.style.display='block';
      }
    }else if(warnEl){
      warnEl.innerHTML='<details class="stackIconDetails"><summary title="積み上げ表示の注意">⚠️</summary><div class="stackIconBody"><strong>積み上げ表示の注意：</strong>選択された回答項目には重複集計項目が含まれている可能性があります。100%構成になる組み合わせを自動判定できなかったため、選択項目を見直してください。</div></details>';
      warnEl.style.display='block';
    }
  }
}
const w=1150,h=560,ml=190,mr=35,mt=28,mb=88; chart.setAttribute('viewBox',`0 0 ${w} ${h}`);
const nums=[]; displayRows.forEach(r=>r.values.forEach(v=>{if(typeof v.value==='number') nums.push(v.value);}));
const rowSums=displayRows.map(r=>r.values.reduce((a,v)=>a+(typeof v.value==='number'?v.value:0),0));
const max= mode==='stack' ? 100 : Math.max(100,...nums);
const avg=nums.length?nums.reduce((a,b)=>a+b,0)/nums.length:0;
const plotW=w-ml-mr, plotH=h-mt-mb; const n=displayRows.length; const gap=8;
const slotH=Math.max(26, Math.min(54, (plotH-(n-1)*gap)/Math.max(n,1)));
const groupGap=2;
const barH= mode==='stack' ? Math.max(8, slotH*.72) : Math.max(4, Math.min(14, (slotH-groupGap*Math.max(0,series.length-1))/Math.max(1,visibleSeries.length)));
const scale=v=>ml+(Math.min(v,max)/max)*plotW;
function line(x1,y1,x2,y2,stroke='#d0d5dd'){const e=document.createElementNS('http://www.w3.org/2000/svg','line'); e.setAttribute('x1',x1);e.setAttribute('y1',y1);e.setAttribute('x2',x2);e.setAttribute('y2',y2);e.setAttribute('stroke',stroke); chart.appendChild(e);}
const step=20;
for(let t=0;t<=max;t+=step){const x=scale(t); line(x,mt,x,mt+plotH,'#edf1f7'); const tx=document.createElementNS('http://www.w3.org/2000/svg','text'); tx.setAttribute('x',x);tx.setAttribute('y',mt+plotH+22);tx.setAttribute('text-anchor','middle');tx.setAttribute('font-size','11');tx.setAttribute('fill','#667085');tx.textContent=t+'%'; chart.appendChild(tx);}
line(ml,mt+plotH,w-mr,mt+plotH,'#98a2b3');
const ax=scale(Math.min(avg,max)); /* average line hidden */
displayRows.forEach((r,i)=>{const y=mt+i*(slotH+gap); const lab=document.createElementNS('http://www.w3.org/2000/svg','text'); lab.setAttribute('x',ml-8);lab.setAttribute('y',y+slotH*.58);lab.setAttribute('text-anchor','end');lab.setAttribute('font-size', n>18?'10':'12');lab.setAttribute('fill','#344054');lab.textContent=(r.label.length>22?r.label.slice(0,21)+'…':r.label); chart.appendChild(lab);
let acc=0;
r.values.forEach((v,si)=>{if(typeof v.value!=='number') return; const rect=document.createElementNS('http://www.w3.org/2000/svg','rect');
let x=ml, width=Math.max(2,scale(v.value)-ml), yPos=y+(slotH-barH)/2, thisH=barH, opacity=1;
if(mode==='stack'){
  x=scale(acc); width=Math.max(2,scale(acc+v.value)-scale(acc)); yPos=y+(slotH-barH)/2; opacity=.9; acc+=v.value;
}else{
  const totalH=visibleSeries.length*barH+(visibleSeries.length-1)*groupGap;
  yPos=y+Math.max(0,(slotH-totalH)/2)+si*(barH+groupGap);
  opacity=Math.max(.38,1-si*.10);
}
const ser=visibleSeries[si]; rect.setAttribute('class','bar'); rect.setAttribute('x',x);rect.setAttribute('y',yPos);rect.setAttribute('width',width);rect.setAttribute('height',thisH);rect.setAttribute('rx',mode==='stack'?2:3);rect.setAttribute('fill',colors[(ser?.origIndex??si)%colors.length]);rect.setAttribute('fill-opacity',String(opacity));rect.setAttribute('stroke','#fff');rect.setAttribute('stroke-opacity',.75);
rect.addEventListener('mousemove',ev=>{const total=r.values.reduce((a,z)=>a+(typeof z.value==='number'?z.value:0),0); tooltip.innerHTML=`<b>${esc(r.label)}</b><br>${esc(v.category)}<br>${esc(v.sheet)} / ${esc(v.answer)}：${fmt(v.value)}%<br>N：${esc(v.n)}<br>${mode==='stack'?'表示項目合計：'+fmt(total)+'%':''}`; const box=$('chartWrap').getBoundingClientRect(); tooltip.style.left=(ev.clientX-box.left+12)+'px'; tooltip.style.top=(ev.clientY-box.top+12)+'px'; tooltip.style.display='block';});
rect.addEventListener('mouseleave',()=>tooltip.style.display='none'); rect.addEventListener('click',ev=>{ev.stopPropagation(); if(ser){hiddenSeries.add(ser.label); render();}}); chart.appendChild(rect);
if(mode==='overlay' && visibleSeries.length<=3){const tv=document.createElementNS('http://www.w3.org/2000/svg','text'); tv.setAttribute('x',scale(v.value)+5);tv.setAttribute('y',yPos+barH*.78);tv.setAttribute('font-size','10');tv.setAttribute('fill','#344054');tv.textContent=fmt(v.value)+'%'; chart.appendChild(tv);}
});
if(mode==='stack' && acc>0){/* 100% label intentionally omitted */}
});

/* average label hidden */

if(!displayRows.length || !visibleSeries.length){const tx=document.createElementNS('http://www.w3.org/2000/svg','text'); tx.setAttribute('x',w/2);tx.setAttribute('y',h/2);tx.setAttribute('text-anchor','middle');tx.setAttribute('fill','#667085');tx.textContent='集計シート・回答項目・属性を選択してください'; chart.appendChild(tx);}
const legendSeries = mode==='stack' ? visibleSeries : series;
$('legend').innerHTML=legendSeries.map((s,i)=>`<span data-series="${esc(s.label)}" class="${hiddenSeries.has(s.label)?'off':''}"><span class="sw" style="background:${colors[(s.origIndex??i)%colors.length]};opacity:${hiddenSeries.has(s.label)?.25:(mode==='stack'?.9:Math.max(.38,1-i*.10))}"></span>${esc(i+1)}. ${esc(s.label)}</span>`).join('') + (autoStack?'<span class="small">※100%構成に自動補正中</span>':'');
$('legend').querySelectorAll('[data-series]').forEach(el=>el.addEventListener('click',()=>{const k=el.getAttribute('data-series'); if(!k)return; hiddenSeries.has(k)?hiddenSeries.delete(k):hiddenSeries.add(k); render();}));
}

function drawTable(rows, series){let headers=['属性分類','属性']; series.forEach((s,i)=>{headers.push(`${i+1}. ${s.label}`,`N${i+1}`,`差${i+1}(pt)`);}); const nums=[]; rows.forEach(r=>r.values.forEach(v=>{if(typeof v.value==='number') nums.push(v.value);})); const avg=nums.length?nums.reduce((a,b)=>a+b,0)/nums.length:0; table.innerHTML='<thead><tr>'+headers.map(h=>`<th>${esc(h)}</th>`).join('')+'</tr></thead><tbody>'+rows.map(r=>'<tr><td>'+esc(r.category)+'</td><td>'+esc(r.label)+'</td>'+r.values.map(v=>`<td>${typeof v.value==='number'?fmt(v.value)+'%':'-'}</td><td>${v.n??'-'}</td><td>${typeof v.value==='number'?fmt(v.value-avg):'-'}</td>`).join('')+'</tr>').join('')+'</tbody>';}
chapterSel.addEventListener('change',populateQuestions);
qSel.addEventListener('change',()=>{sortMode=false; hiddenSeries.clear(); populateSheets();});
sheetChecks.addEventListener('change',()=>{hiddenSeries.clear(); sheetOrder=updateOrder(sheetChecks,sheetOrder,true); populateAnswers();});
answerChecks.addEventListener('change',()=>{hiddenSeries.clear(); answerOrder=updateOrder(answerChecks,answerOrder,false); render();});
filterText.addEventListener('input',()=>populateChecks(false));
checks.addEventListener('change',render);
$('sheetAll').addEventListener('click',()=>{sheetChecks.querySelectorAll('input').forEach(i=>i.checked=true); sheetOrder=updateOrder(sheetChecks,sheetOrder,true); populateAnswers();});
$('sheetClear').addEventListener('click',()=>{sheetChecks.querySelectorAll('input').forEach(i=>i.checked=false); sheetOrder=[]; populateAnswers();});
$('answerAll').addEventListener('click',()=>{hiddenSeries.clear(); answerChecks.querySelectorAll('input').forEach(i=>i.checked=true); answerOrder=updateOrder(answerChecks,answerOrder,false); render();});
$('answerClear').addEventListener('click',()=>{hiddenSeries.clear(); answerChecks.querySelectorAll('input').forEach(i=>i.checked=false); answerOrder=[]; render();});
$('selectShown').addEventListener('click',()=>{checks.querySelectorAll('input').forEach(i=>i.checked=true); render();});
$('clearAll').addEventListener('click',()=>{checks.querySelectorAll('input').forEach(i=>i.checked=false); render();});
$('sortBtn').addEventListener('click',()=>{sortMode=true; render();});
$('resetSort').addEventListener('click',()=>{sortMode=false; render();});
document.querySelectorAll('input[name="chartMode"]').forEach(i=>i.addEventListener('change',render));
paletteSelect.addEventListener('change',()=>{updatePalettePreview();render();});
$('sideToggle').addEventListener('click',()=>{document.body.classList.toggle('sidebar-collapsed'); const closed=document.body.classList.contains('sidebar-collapsed'); $('sideToggle').innerHTML=closed?'<span class="mark">＋</span>':'<span class="mark">−</span>'; $('sideToggle').setAttribute('aria-label',closed?'左メニューを展開する':'左メニューを折りたたむ');});
$('agreeBtn').addEventListener('click',()=>document.body.classList.remove('consent-pending'));
updatePalettePreview();
populateChapters();


function normalizeChartSvgLayout(){
  const svg=document.getElementById('chart');
  const wrap=document.getElementById('chartWrap');
  if(!svg||!wrap) return;
  const w=Math.max(480, Math.floor(wrap.clientWidth||svg.clientWidth||900));
  const h=parseFloat(svg.getAttribute('height')||svg.style.height||560)||560;
  svg.setAttribute('width', w);
  svg.setAttribute('viewBox', '0 0 '+w+' '+h);
  svg.style.width='100%';
  svg.style.maxWidth='100%';
}


function forceCompactWarningIcon(){
  setTimeout(()=>{try{normalizeStackWarnings()}catch(e){}},0);
  setTimeout(()=>{try{normalizeStackWarnings()}catch(e){}},50);
}

const __originalRenderForWarnings = render;
render = function(){
  const result = __originalRenderForWarnings.apply(this, arguments);
  try{ normalizeChartSvgLayout(); normalizeStackWarnings(); forceCompactWarningIcon(); }catch(e){}
  return result;
};

populateQuestions();
try{forceCompactWarningIcon()}catch(e){}
if(window.ResizeObserver){new ResizeObserver(()=>{try{render()}catch(e){}}).observe(document.getElementById('chartWrap'));}
})();
