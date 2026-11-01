/* ===== 同意ゲート ===== */
(function(){
  const gate=document.getElementById('consentGate');
  const app=document.getElementById('app');
  const KEY='lifeDeductionSim.consent';
  if(localStorage.getItem(KEY)==='yes'){ gate.style.display='none'; app.classList.remove('hidden'); }
  document.getElementById('agree').onclick=()=>{ localStorage.setItem(KEY,'yes'); gate.style.display='none'; app.classList.remove('hidden'); };
  document.getElementById('decline').onclick=()=>alert('同意いただけない場合はご利用いただけません。');
})();

/* ===== ユーティリティ ===== */
const fmt=n=>(Math.round(n||0)).toLocaleString('ja-JP');
const yen=s=>{ if(typeof s==='number')return s; return +String(s||'0').replace(/,/g,'').replace(/[^\d.-]/g,'')||0; };
document.addEventListener('input',e=>{
  if(e.target.matches('input[type=text]')) e.target.value=fmt(yen(e.target.value));
});

/* ===== 入力行 追加・削除・ステップ ===== */
function stepButtons(){
  return `<div class="stepbox">
    <button class="step" data-delta="-10000">−10,000</button>
    <button class="step" data-delta="-1000">−1,000</button>
    <button class="step reset" data-reset="true">0</button>
    <button class="step" data-delta="1000">＋1,000</button>
    <button class="step" data-delta="10000">＋10,000</button>
  </div>`;
}
function addRow(bucket){
  const host=document.getElementById('list-'+bucket);
  const row=document.createElement('div');
  row.className='grid2';
  row.innerHTML=`<div class="rowwrap"><input type="text" value="0">${stepButtons()}</div><button class="del">✕</button>`;
  host.appendChild(row);
}
['genNew','genOld','medNew','annNew','annOld'].forEach(addRow);

/* ▼ 「＋証明書を追加」 クリックで行追加（イベント委譲） */
document.addEventListener('click',e=>{
  const add=e.target.closest('[data-add]'); if(add){ addRow(add.dataset.add); }
});

/* ▼ 行削除（✕） */
document.addEventListener('click',e=>{
  const del=e.target.closest('.del'); if(del){ const r=del.closest('.grid2'); r?.remove(); }
});

/* ▼ 汎用ステップ（年収用 data-target ボタンは除外） */
document.addEventListener('click',e=>{
  const btn=e.target.closest('.step'); if(!btn) return;
  if(btn.dataset.target) return;                 // 年収は別ハンドラ
  const wrap=e.target.closest('.rowwrap'), inp=wrap?.querySelector('input'); if(!inp) return;
  if(btn.dataset.reset){ inp.value='0'; return; }
  inp.value=fmt(Math.max(0, yen(inp.value)+(+btn.dataset.delta||0)));
});

/* ▼ 年収ステップ専用 */
document.addEventListener('click',e=>{
  const btn=e.target.closest('.step'); if(!btn||!btn.dataset.target) return;
  const inp=document.getElementById(btn.dataset.target);
  if(btn.dataset.reset){ inp.value='0'; return; }
  inp.value=fmt(Math.max(0, yen(inp.value)+(+btn.dataset.delta||0)));
});

/* A〜E 合計セルクリックで直接上書き */
document.addEventListener('click',e=>{
  const t=e.target.closest('[data-edit]'); if(!t) return;
  const bucket=t.dataset.edit;
  const cur=sum(bucket);
  const v=window.prompt('この区分の合計額を入力', fmt(cur));
  if(v==null) return;
  const n=yen(v);
  const inputs=[...document.querySelectorAll('#list-'+bucket+' input')];
  if(inputs.length===0) addRow(bucket);
  const all=[...document.querySelectorAll('#list-'+bucket+' input')];
  all[0].value=fmt(n); for(let i=1;i<all.length;i++) all[i].value='0';
});

/* ===== 計算式 ===== */
function newSys(x){ if(x<=20000)return x; if(x<=40000)return x/2+10000; if(x<=80000)return x/4+20000; return 40000; }
function oldSys(x){ if(x<=25000)return x; if(x<=50000)return x/2+12500; if(x<=100000)return x/4+25000; return 50000; }
function invNewSys(d){ if(d<=0)return 0; if(d<=20000)return d; if(d<=30000)return 2*(d-10000); if(d<=40000)return 4*(d-20000); return 80000; }
function sum(bucket){ let t=0; document.querySelectorAll('#list-'+bucket+' input').forEach(i=>t+=yen(i.value)); return t; }

/* ===== ②計算（手動） ===== */
function compute(){
  const A=sum('genNew'), B=sum('genOld'), C=sum('medNew'), D=sum('annNew'), E=sum('annOld');
  const one=Math.min(40000,newSys(A));
  const two=Math.min(50000,oldSys(B));
  const three=Math.min(40000, one+two);
  const medI=Math.min(40000,newSys(C));
  const four=Math.min(40000,newSys(D));
  const five=Math.min(50000,oldSys(E));
  const six=Math.min(40000, four+five);
  const ina=Math.max(two,three);
  const inu=medI;
  const ino=Math.max(five,six);
  const totalResident=ina+inu+ino;
  const totalIncome  =three+medI+six;
  return {A,B,C,D,E,one,two,three,medI,four,five,six,ina,inu,ino,totalResident,totalIncome};
}
function fill(){
  const r=compute(), set=(id,v)=>{const el=document.getElementById(id); if(el) el.textContent=fmt(v);};
  set('A',r.A); set('B',r.B); set('C',r.C); set('D',r.D); set('E',r.E);
  set('one',r.one); set('two',r.two); set('three',r.three);
  set('inu',r.inu); set('four',r.four); set('five',r.five); set('six',r.six);
  document.getElementById('ina').textContent=fmt(r.ina);
  document.getElementById('ino').textContent=fmt(r.ino);
  document.getElementById('kanaTotal').textContent=fmt(r.totalResident);
  buildIncomeTable(r); buildUnusedTable(r);
  document.getElementById('fs-res').classList.remove('hidden');
  document.getElementById('fs-inc').classList.remove('hidden');
  document.getElementById('fs-unused').classList.remove('hidden');
}
document.getElementById('calc').addEventListener('click', fill);

/* ===== ③ 年収レンジ表 ===== */
function marginalRate(annual){
  const y=annual;
  if(y<=1950000) return 0.05;
  if(y<=3300000) return 0.10;
  if(y<=6950000) return 0.20;
  if(y<=9000000) return 0.23;
  if(y<=18000000) return 0.33;
  if(y<=40000000) return 0.40;
  return 0.45;
}
function buildIncomeTable(r){
  const host=document.getElementById('incTable');
  const base=yen(document.getElementById('income').value)||0;
  const STEP=1000000, MIN=2000000, MAX=12000000;
  let start=Math.floor(base/STEP)*STEP - 4000000;
  start=Math.max(MIN, Math.min(start, MAX-10*STEP));
  const incomes=[]; for(let v=start; v<=Math.min(MAX,start+10*STEP); v+=STEP) incomes.push(v);
  let html='<tr><th>年収</th><th>限界税率</th><th>控除額（所得税）</th><th>控除額（住民税）</th><th>合計軽減額</th></tr>';
  incomes.forEach(v=>{
    const rate=marginalRate(v);
    const incomeSave=r.totalIncome*rate;
    const residentSave=r.totalResident*0.10;
    const total=incomeSave+residentSave;
    const hit=(v===Math.floor(base/STEP)*STEP)?' class="hit"':'';
    html+=`<tr${hit}><td>${fmt(v)}</td><td>${Math.round(rate*100)}%</td><td>${fmt(incomeSave)}</td><td>${fmt(residentSave)}</td><td>${fmt(total)}</td></tr>`;
  });
  host.innerHTML=html;
}

/* ===== ④ 未使用枠（旧制度控除を考慮） ===== */
function invNewSys(d){
  if(d<=0) return 0;
  if(d<=20000) return d;
  if(d<=30000) return 2*(d-10000);
  if(d<=40000) return 4*(d-20000);
  return 80000;
}
function remainForNew(currentNewPremium, oldDeduction){
  const targetDed=Math.max(0, 40000 - Math.min(40000, oldDeduction||0));
  const curDed=newSys(currentNewPremium);
  if(curDed>=targetDed) return 0;
  const neededPremium=invNewSys(targetDed);
  return Math.max(0, neededPremium - currentNewPremium);
}
function buildUnusedTable(r){
  const remainA=remainForNew(r.A, r.two);
  const remainC=Math.max(0, 80000 - r.C);
  const remainD=remainForNew(r.D, r.five);
  const rows=[
    {label:'一般生命保険料（新制度）', amt:r.A, remain:remainA},
    {label:'介護医療保険料（新制度）', amt:r.C, remain:remainC},
    {label:'個人年金保険料（新制度）', amt:r.D, remain:remainD},
  ];
  let html='<tr><th>区分</th><th>現在の年額</th><th>未使用枠（年額）</th><th>未使用枠（月額目安）</th></tr>';
  rows.forEach(x=>{
    html+=`<tr>
      <td><span class="badge-new">新</span> ${x.label}</td>
      <td>${fmt(x.amt)}</td>
      <td>${fmt(x.remain)}</td>
      <td>${fmt(x.remain/12)}</td>
    </tr>`;
  });
  document.getElementById('unusedTable').innerHTML=html;
}
