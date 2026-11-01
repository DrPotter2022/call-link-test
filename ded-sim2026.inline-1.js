const fmt = n => (isNaN(n)?0:n).toLocaleString();
const unfmt = s => (s+"").replace(/[^\d\-]/g,"")*1;
// 計算後は入力変更のたびに自動で再計算（スマホで「計算」を押し直さなくてよい）
function autoCalc(){ if(document.getElementById('resultsCard').style.display==='block') calculate(); }
function bindCommaInput(el){ el.addEventListener('blur',()=> el.value = fmt(unfmt(el.value))); el.addEventListener('focus',()=> el.select()); }

const buckets = [
  { key:'A', label:'新・一般（A）', color:'bg-new'   },
  { key:'B', label:'旧・一般（B）', color:'bg-old'   },
  { key:'C', label:'介護医療（C｜新のみ）', color:'bg-care' },
  { key:'D', label:'新・個人年金（D）', color:'bg-dnew'},
  { key:'E', label:'旧・個人年金（E）', color:'bg-eold'}
];

const state = { A:[0], B:[0], C:[0], D:[0], E:[0], income:4520000, special2026:false };

function renderInputs(){
  const wrap = document.getElementById('inputs');
  wrap.innerHTML='';
  buckets.forEach(b=>{
    const box = document.createElement('div');
    box.className='card '+b.color; box.style.margin='10px 0';
    box.innerHTML = state[b.key].map((v,i)=>`
      <div class="inrow">
        <div class="inhead">
          <span class="chip">✏️ ${b.label}${state[b.key].length>1?`・${i+1}`:''}</span>
          ${state[b.key].length>1?`<button class="btn del" aria-label="この行を削除" data-del="${b.key}" data-idx="${i}">×</button>`:''}
        </div>
        <div class="amt"><input class="mono" data-k="${b.key}" data-idx="${i}" type="text" inputmode="numeric" enterkeyhint="done" autocomplete="off" value="${fmt(v)}" /><span>円</span><button class="btn calc-btn" type="button" data-calc="1" data-k="${b.key}" data-idx="${i}" aria-label="電卓で入力">🧮</button></div>
      </div>
    `).join('') + `<div><button class="btn addbtn" data-add="${b.key}">＋証明書を追加</button></div>`;
    wrap.appendChild(box);
  });
  wrap.querySelectorAll('input[type="text"]').forEach(el=>{
    bindCommaInput(el);
    el.addEventListener('input',()=>{ state[el.dataset.k][el.dataset.idx*1] = unfmt(el.value); autoCalc(); });
  });
  wrap.querySelectorAll('button[data-calc]').forEach(b=>{
    b.addEventListener('click',()=>{
      const k=b.dataset.k, idx=b.dataset.idx*1;
      openKeypad({title:'金額を入力', value:state[k][idx]||0, onOk:v=>{ state[k][idx]=v; renderInputs(); autoCalc(); }});
    });
  });
  wrap.querySelectorAll('button[data-add]').forEach(b=> b.addEventListener('click',()=>{ state[b.dataset.add].push(0); renderInputs(); }));
  wrap.querySelectorAll('button[data-del]').forEach(b=> b.addEventListener('click',()=>{ const k=b.dataset.del, i=b.dataset.idx*1; state[k].splice(i,1); if(state[k].length===0) state[k]=[0]; renderInputs(); autoCalc(); }));
}
const incEl=document.getElementById('income'), incSl=document.getElementById('incSlider');
function setIncome(v){ state.income=Math.max(0,v); incEl.value=fmt(state.income); incSl.value=Math.min(state.income,+incSl.max); }
incSl.addEventListener('input',()=>{ setIncome(+incSl.value); autoCalc(); });
incEl.addEventListener('input',()=>{ incSl.value=Math.min(Math.max(0,unfmt(incEl.value)),+incSl.max); });
document.getElementById('incCalc').addEventListener('click',()=>{
  openKeypad({title:'年収を入力', value:unfmt(incEl.value)||0, onOk:v=>{ setIncome(v); autoCalc(); }});
});

// ---- 電卓ポップアップ ----
let kpCb=null, kpVal='';
const kp=document.getElementById('kp'), kpDisp=document.getElementById('kpDisp');
function kpShow(){ kpDisp.textContent = fmt(+kpVal||0); }
function openKeypad({title,value,onOk}){
  kpCb=onOk; kpVal=String(value||0)==='0'?'':String(value);
  document.getElementById('kpTitle').textContent=title; kpShow(); kp.style.display='flex';
}
function closeKeypad(){ kp.style.display='none'; kpCb=null; }
kp.addEventListener('click',e=>{
  if(e.target===kp){ closeKeypad(); return; }
  const b=e.target.closest('button'); if(!b) return;
  const t=b.dataset.t;
  if(t==='ok'){ const cb=kpCb, v=+kpVal||0; closeKeypad(); if(cb) cb(v); return; }
  if(t==='cancel'){ closeKeypad(); return; }
  if(t==='back') kpVal=kpVal.slice(0,-1);
  else if(t==='clear') kpVal='';
  else if(t==='add'){ kpVal=String(Math.max(0,(+kpVal||0)+(+b.dataset.v))); }
  else kpVal=(kpVal+t).replace(/^0+/,'');
  if(kpVal.length>9) kpVal=kpVal.slice(0,9);
  kpShow();
});

function dedNew(a){a=Math.max(0,a); if(a<=20000)return a; if(a<=40000)return Math.floor(a/2)+10000; if(a<=80000)return Math.floor(a/4)+20000; return 40000;}
function dedOld(a){a=Math.max(0,a); if(a<=25000)return a; if(a<=50000)return Math.floor(a/2)+12500; if(a<=100000)return Math.floor(a/4)+25000; return 50000;}
function dedNew2026(a){a=Math.max(0,a); if(a<=30000)return a; if(a<=60000)return Math.floor(a/2)+15000; if(a<=120000)return Math.floor(a/4)+30000; return 60000;}
const sum = arr => arr.reduce((s,v)=>s+Math.max(0,v|0),0);

function calculate(){
  state.income = unfmt(document.getElementById('income').value);
  state.special2026 = document.getElementById('sp2026').checked;

  const A=sum(state.A), B=sum(state.B), C=sum(state.C), D=sum(state.D), E=sum(state.E);
  const dep = state.special2026;

  // 控除証明書（2026年分〜）の様式に対応：①〜⑧・㋑㋺㋩
  const d1 = dedNew(A);                 // ① 計算式Ⅰ
  const d2 = dedNew2026(A);             // ② 計算式Ⅱ（23歳未満の扶養親族あり）
  const d3 = dedOld(B);                 // ③ 計算式Ⅲ
  const d4 = Math.min(d1+d3, 40000);    // ④ ①＋③
  const d5 = Math.min(d2+d3, 60000);    // ⑤ ②＋③
  const vI = dep? d5 : Math.max(d3, d4); // ㋑（有：⑤／無：③と④の大きい方）
  const dRo = dedNew(C);                // ㋺
  const d6 = dedNew(D);                 // ⑥
  const d7 = dedOld(E);                 // ⑦
  const d8 = Math.min(d6+d7, 40000);    // ⑧
  const vP = Math.max(d7, d8);          // ㋩

  const grand = Math.min(vI + dRo + vP, 120000); // 合計上限12万円

  // ② 表示
  document.getElementById('resultsCard').style.display='block';
  document.getElementById('depState').textContent = dep? '有':'無';
  document.getElementById('rowA2').style.opacity = dep? 1 : 0.4;
  document.getElementById('aTotal').textContent = fmt(A);
  document.getElementById('bTotal').textContent = fmt(B);
  document.getElementById('cTotal').textContent = fmt(C);
  document.getElementById('dTotal').textContent = fmt(D);
  document.getElementById('eTotal').textContent = fmt(E);

  const setv=(id,v)=>document.getElementById(id).textContent=fmt(v);
  setv('v1',d1); setv('v2',d2); setv('v3',d3); setv('v4',d4); setv('v5',d5); setv('vI',vI);
  setv('vRo',dRo); setv('v6',d6); setv('v7',d7); setv('v8',d8); setv('vP',vP);
  setv('grand',grand);
  // 申告書イメージへ反映
  document.getElementById('formCard').style.display='block';
  const ov=(k,v)=>{ document.getElementById('ov-'+k).textContent = (v===null)?'':fmt(v); };
  ov('A',A); ov('B',B); ov('C',C); ov('D',D); ov('E',E);
  ov('o1',d1); ov('o2',dep?d2:null); ov('o3',d3); ov('o4',d4); ov('o5',dep?d5:null); ov('oI',vI);
  ov('oRo',dRo); ov('o6',d6); ov('o7',d7); ov('o8',d8); ov('oHa',vP); ov('oSum',grand);


  // ③ 年収レンジ（★修正：入力年収に連動してハイライト）
  document.getElementById('rangeCard').style.display='block';
  const ranges=[
    {label:'〜4,000,000',         min:0,        max:4000000,  rate:0.05},
    {label:'4,000,001〜6,000,000',min:4000001,  max:6000000,  rate:0.10},
    {label:'6,000,001〜9,000,000',min:6000001,  max:9000000,  rate:0.20},
    {label:'9,000,001〜10,000,000',min:9000001, max:10000000, rate:0.23},
    {label:'10,000,001〜',        min:10000001, max:Infinity, rate:0.33}
  ];

  const tb=document.querySelector('#rangeTable tbody');
  tb.innerHTML='';

  const income = Math.max(0, state.income|0);

  ranges.forEach(r=>{
    const it=Math.round(grand*r.rate), lt=7000;
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${r.label}</td>
                  <td class="mono" style="text-align:right">${Math.round(r.rate*100)}%</td>
                  <td class="mono" style="text-align:right">${fmt(it)}</td>
                  <td class="mono" style="text-align:right">${fmt(lt)}</td>
                  <td class="mono" style="text-align:right;font-weight:700">${fmt(it+lt)}</td>`;

    // 入力年収がこのレンジならピンクで強調（スクショと同じ連動）
    if(income >= r.min && income <= r.max){
      tr.style.background = 'var(--pink)';
    }
    tb.appendChild(tr);
  });

  // ④ 未使用枠
  document.getElementById('unusedCard').style.display='block';
  const capA2 = dep?60000:40000, capD=40000, capC=40000;
  const unusedA = Math.max(0, capA2 - (dep? d5 : d4));
  const unusedD = Math.max(0, capD  - d8);
  const unusedC = Math.max(0, capC  - dRo);
  const utb=document.querySelector('#unusedTable tbody'); utb.innerHTML='';
  [{nm:'一般生命保険料（新制度）',now:A,rest:unusedA},
   {nm:'介護医療保険料（新制度）',now:C,rest:unusedC},
   {nm:'個人年金保険料（新制度）',now:D,rest:unusedD}]
   .forEach(r=>{
     const tr=document.createElement('tr');
     tr.innerHTML=`<td>${r.nm}</td><td class="mono" style="text-align:right">${fmt(r.now)}</td>
                   <td class="mono" style="text-align:right">${fmt(r.rest)}</td>
                   <td class="mono" style="text-align:right">${fmt(Math.ceil(r.rest/12))}</td>`;
     utb.appendChild(tr);
   });
}

/* 初期セット */
document.getElementById('income').value=fmt(state.income);
bindCommaInput(document.getElementById('income'));
renderInputs();
document.getElementById('calc').addEventListener('click',calculate);
document.getElementById('agreeOk').addEventListener('click',()=>document.getElementById('agree').style.display='none');
document.addEventListener('keydown',e=>{ if(e.key==='Enter'){ calculate(); if(e.target&&e.target.blur) e.target.blur(); } });
document.getElementById('income').addEventListener('input',autoCalc);
document.getElementById('sp2026').addEventListener('change',autoCalc);
