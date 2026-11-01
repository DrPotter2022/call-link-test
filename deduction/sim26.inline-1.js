/* ===== 共通ユーティリティ ===== */
const $ = s=>document.querySelector(s);
const $$ = s=>document.querySelectorAll(s);
const numRaw = t => Number(String(t).replace(/[^0-9.-]/g,''))||0;
const withComma = n => n.toLocaleString('ja-JP');

/* mailto 起動の確実化（_self → iframe → コピー救済） */
function bindMailtoFix(){
  const copyToClipboard = async (txt) => {
    try {
      await navigator.clipboard.writeText(txt);
      alert('メールアドレスをコピーしました。お使いのメールアプリで貼り付けてご利用ください。');
    } catch {
      prompt('以下をコピーしてご利用ください：', txt);
    }
  };
  $$('.mailLink').forEach(a=>{
    a.removeAttribute('target');
    a.addEventListener('click', async (e)=>{
      e.preventDefault();
      const href=a.getAttribute('href');
      const addr=(href.match(/^mailto:([^?]+)/)||[])[1]||'seminar.wize@gmail.com';
      try{ window.open(href,'_self'); }catch(_){}
      setTimeout(()=>{
        if(document.visibilityState==='visible'){
          try{
            const ifr=document.createElement('iframe');
            ifr.style.display='none'; ifr.src=href; document.body.appendChild(ifr);
            setTimeout(()=>ifr.remove(),2500);
          }catch(_){}
        }
        setTimeout(()=>{ if(document.visibilityState==='visible'){ copyToClipboard(addr); } },1000);
      },400);
    });
  });
}

/* 入力整形 */
function syncFormatInput(el){ const v=numRaw(el.value); el.value=withComma(v); }
function syncAllMoney(){ $$('.money').forEach(el=>syncFormatInput(el)); syncFormatInput($('#income')); }

/* 証明書管理 */
const certCounters={a:0,b:0,c:0,d:0,e:0};
document.addEventListener('click',e=>{
  if(e.target.classList.contains('adjBtn')){
    const tId=e.target.dataset.target, diff=Number(e.target.dataset.diff);
    const el=document.getElementById(tId); if(!el) return;
    if(diff===0){ el.value='0'; } else { let nv=numRaw(el.value)+diff; if(nv<0) nv=0; el.value=withComma(nv); }
    sumLines();
  }
  if(e.target.matches('.addBtn[data-add-cat]')) addCertificateRow(e.target.dataset.addCat);
});
function addCertificateRow(cat){
  certCounters[cat]+=1;
  const id=cat+'Cert'+certCounters[cat];
  const row=document.createElement('div'); row.className='certRow';
  const inp=document.createElement('input'); inp.type='text'; inp.id=id; inp.className='mono money certInput'; inp.placeholder='＋証明書（円）'; inp.value='0'; inp.dataset.cat=cat;
  inp.addEventListener('input',()=>{ syncFormatInput(inp); sumLines(); });
  row.appendChild(inp);
  [-10000,-1000,-100,-10,-1,0,1,10,100,1000,10000].forEach(v=>{
    const b=document.createElement('button'); b.className='pill adjBtn'; b.dataset.target=id; b.dataset.diff=String(v); b.textContent=(v>0?`+${withComma(v)}`:withComma(v)); row.appendChild(b);
  });
  document.getElementById(cat+'CertWrap').appendChild(row); sumLines();
}
function sumLines(){
  ['a','b','c','d','e'].forEach(k=>{
    const base=document.getElementById(k+'Sum'); if(!base) return;
    let total=numRaw(base.value);
    const wrap=document.getElementById(k+'CertWrap'); if(wrap) wrap.querySelectorAll('input.money.certInput').forEach(el=>total+=numRaw(el.value));
    const edit=document.getElementById(k.toUpperCase()+'_edit'); if(edit) edit.value=withComma(total);
  });
}
;['A_edit','B_edit','C_edit','D_edit','E_edit'].forEach(id=>{
  const el=document.getElementById(id);
  el.addEventListener('input',()=>{
    syncFormatInput(el);
    const k=id[0].toLowerCase();
    const base=document.getElementById(k+'Sum'); if(!base) return;
    const wrap=document.getElementById(k+'CertWrap'); if(wrap) wrap.querySelectorAll('input.money.certInput').forEach(x=>x.value='0');
    base.value=el.value; sumLines();
  });
});

/* ===== 計算ロジック ===== */
/* 所得税側（新制度 計算式Ⅰ） */
function newSchemeI(p,cap=40000){ if(p<=0) return 0; if(p<=20000) return Math.min(p,cap); if(p<=40000) return Math.min(Math.floor(p/2)+10000,cap); if(p<=80000) return Math.min(Math.floor(p/4)+20000,cap); return cap; }
/* 旧制度（所得税） */
function oldSchemeII(p){ if(p<=0) return 0; if(p<=25000) return p; if(p<=50000) return Math.floor(p/2)+12500; if(p<=100000) return Math.floor(p/4)+25000; return 50000; }
/* 住民税（新制度） */
function newSchemeLT(p){ if(p<=0) return 0; if(p<=12000) return p; if(p<=32000) return Math.floor(p/2)+6000; if(p<=56000) return Math.floor(p/4)+14000; return 28000; }
/* 住民税（旧制度） */
function oldSchemeLT(p){ if(p<=0) return 0; if(p<=15000) return p; if(p<=40000) return Math.floor(p/2)+7500; if(p<=70000) return Math.floor(p/4)+17500; return 35000; }

function bestLTForGeneral(aNew,bOld){ const n=newSchemeLT(aNew), o=oldSchemeLT(bOld); return Math.max(o,n,Math.min(28000,o+n)); }
function bestLTForPension(dNew,eOld){ const n=newSchemeLT(dNew), o=oldSchemeLT(eOld); return Math.max(o,n,Math.min(28000,o+n)); }
function computeResidentGrand(aP,bP,cP,dP,eP){ const gen=bestLTForGeneral(aP,bP), med=newSchemeLT(cP), pen=bestLTForPension(dP,eP); return Math.min(70000,gen+med+pen); }

/* 逆算（控除 → 必要保険料（年））新制度のみ */
function inverseNewSchemeI(d,cap=40000){ d=Math.max(0,Math.min(d,cap)); if(d<=20000) return d; if(d<=30000) return 2*(d-10000); if(d<=40000) return 4*(d-20000); return 80000; }

/* ③ Beforeテーブル */
function buildRangeRows(tbody,income,grandIT,grandLT){
  const ranges=[
    {label:'〜4,000,000',rate:0.05,match:x=>x<=4000000},
    {label:'4,000,001〜6,000,000',rate:0.10,match:x=>x>=4000001&&x<=6000000},
    {label:'6,000,001〜9,000,000',rate:0.20,match:x=>x>=6000001&&x<=9000000},
    {label:'9,000,001〜10,000,000',rate:0.23,match:x=>x>=9000001&&x<=10000000},
    {label:'10,000,001〜',rate:0.33,match:x=>x>=10000001},
  ];
  tbody.innerHTML='';
  const ltTax=Math.round(Math.min(70000,grandLT)*0.10);
  ranges.forEach(r=>{
    const it=Math.round(grandIT*r.rate);
    const tr=document.createElement('tr');
    tr.innerHTML=`<th>${r.label}</th><td class="r">${(r.rate*100).toFixed(0)}%</td><td class="r">${withComma(it)}</td><td class="r">${withComma(ltTax)}</td><td class="r sum">${withComma(it+ltTax)}</td>`;
    if(r.match(income)) tr.classList.add('rangeHighlight');
    tbody.appendChild(tr);
  });
}

/* ④テーブル（新制度のみ）：a,b,c,d */
const reuseFlags = {a:true, c:true, d:true};
function fillUnusedTable(u){
  const tb=$('#unusedTable tbody'); tb.innerHTML='';
  const {aNow,cNow,dNow,aNeed,cNeed,deNeed,d1_now,d4C_now,d4D_now,capGeneral}=u;

  const rowsA = [
    {cat:'a', label:'一般生命保険料（新制度）', a:d1_now, b:aNeed,
      c: (()=>{ const before=inverseNewSchemeI(d1_now,capGeneral); const after=inverseNewSchemeI(d1_now+aNeed,capGeneral); return Math.max(0,Math.ceil((after-before)/12)); })()
    },
    {cat:'c', label:'介護医療保険料（新のみ）', a:d4C_now, b:cNeed,
      c: (()=>{ const before=inverseNewSchemeI(d4C_now,40000); const after=inverseNewSchemeI(d4C_now+cNeed,40000); return Math.max(0,Math.ceil((after-before)/12)); })()
    },
    {cat:'d', label:'新 個人年金保険料',       a:d4D_now, b:deNeed,
      c: (()=>{ const before=inverseNewSchemeI(d4D_now,40000); const after=inverseNewSchemeI(d4D_now+deNeed,40000); return Math.max(0,Math.ceil((after-before)/12)); })()
    }
  ];

  rowsA.forEach(r=>{
    const tr=document.createElement('tr');
    const d = r.c*12;
    tr.innerHTML=`
      <td>
        <label style="display:flex;align-items:center;gap:6px;">
          <input type="checkbox" class="reuseChk" data-cat="${r.cat}" ${reuseFlags[r.cat]?'checked':''}/>
          <span>${r.label}</span>
        </label>
      </td>
      <td class="r">${withComma(r.a)}</td>
      <td class="r">${withComma(r.b)}</td>
      <td class="r">${withComma(r.c)}</td>
      <td class="r">${withComma(d)}</td>`;
    tb.appendChild(tr);
  });
}
document.addEventListener('change',e=>{
  if(e.target.classList.contains('reuseChk')){
    const cat=e.target.dataset.cat;
    reuseFlags[cat]=e.target.checked;
  }
});

/* ⑤ After（所得税側） */
function computeGrandAfterUsingUnused_IT(ctx){
  const capGeneral=ctx.special2026?60000:40000;
  const extraA = reuseFlags.a ? ctx.unusedA : 0;
  const d3_after=Math.min(capGeneral, ctx.d1 + ctx.d2 + extraA);
  const vI_after=Math.max(ctx.d2, d3_after);
  const d4C_after=Math.min(40000, ctx.d4C + (reuseFlags.c?ctx.unusedC:0));
  const d6_after =Math.min(40000, ctx.d4D + ctx.d5 + (reuseFlags.d?ctx.unusedD:0));
  const vP_after=Math.max(ctx.d5, d6_after);
  return vI_after + d4C_after + vP_after;
}
/* ⑤ After（住民税側） */
function computeGrandAfterUsingUnused_LT(ctx,aP,bP,cP,dP,eP){
  const capGeneral=ctx.special2026?60000:40000;
  const d3_target = Math.min(capGeneral, ctx.d1 + ctx.d2 + (reuseFlags.a?ctx.unusedA:0));
  const aP_after  = reuseFlags.a ? Math.max(aP, inverseNewSchemeI(d3_target,capGeneral)) : aP;
  const d4C_target= Math.min(40000, ctx.d4C + (reuseFlags.c?ctx.unusedC:0));
  const cP_after  = reuseFlags.c ? Math.max(cP, inverseNewSchemeI(d4C_target,40000)) : cP;
  const d6_target = Math.min(40000, ctx.d4D + ctx.d5 + (reuseFlags.d?ctx.unusedD:0));
  const dP_after  = reuseFlags.d ? Math.max(dP, inverseNewSchemeI(d6_target,40000)) : dP;
  return computeResidentGrand(aP_after,bP,cP_after,dP_after,eP);
}
function buildCompareRangeRows(tbody,income,bIT,aIT,bLT,aLT){
  const ranges=[
    {label:'〜4,000,000',rate:0.05,match:x=>x<=4000000},
    {label:'4,000,001〜6,000,000',rate:0.10,match:x=>x>=4000001&&x<=6000000},
    {label:'6,000,001〜9,000,000',rate:0.20,match:x=>x>=6000001&&x<=9000000},
    {label:'9,000,001〜10,000,000',rate:0.23,match:x=>x>=9000001&&x<=10000000},
    {label:'10,000,001〜',rate:0.33,match:x=>x>=10000001},
  ];
  tbody.innerHTML='';
  const ltB=Math.round(Math.min(70000,bLT)*0.10);
  const ltA=Math.round(Math.min(70000,aLT)*0.10);
  ranges.forEach(r=>{
    const itB=Math.round(bIT*r.rate), itA=Math.round(aIT*r.rate);
    const tr=document.createElement('tr');
    tr.innerHTML=`<th>${r.label}</th>
      <td class="r">${withComma(itB)}</td><td class="r">${withComma(ltB)}</td><td class="r sum">${withComma(itB+ltB)}</td>
      <td class="r">${withComma(itA)}</td><td class="r">${withComma(ltA)}</td><td class="r sum">${withComma(itA+ltA)}</td>`;
    if(r.match(income)) tr.classList.add('rangeHighlight');
    tbody.appendChild(tr);
  });
}

/* ===== メイン計算 ===== */
function calculate(){
  const income=numRaw($('#income').value);
  const a=numRaw($('#A_edit').value||$('#aSum').value);
  const b=numRaw($('#B_edit').value||$('#bSum').value);
  const c=numRaw($('#C_edit').value||$('#cSum').value);
  const d=numRaw($('#D_edit').value||$('#dSum').value);
  const e=numRaw($('#E_edit').value||$('#eSum').value);

  const special2026=$('#sp2026').checked;
  const capGeneral=special2026?60000:40000;
  $('#capGenLbl').textContent=withComma(capGeneral);
  $('#capGenLbl2').textContent=withComma(capGeneral);
  $('#capGenSide').textContent=withComma(capGeneral);

  // ② 現状控除
  const d1=newSchemeI(a,capGeneral);
  const d2=oldSchemeII(b);
  const d3=Math.min(capGeneral,d1+d2);
  const vI=Math.max(d2,d3);
  const d4C=newSchemeI(c,40000);
  const d4D=newSchemeI(d,40000);
  const d5=oldSchemeII(e);
  const d6=Math.min(40000,d4D+d5);
  const vP=Math.max(d5,d6);
  const grandIT=vI+d4C+vP;
  const grandLT=computeResidentGrand(a,b,c,d,e);

  // ② 表示
  $('#d1').textContent=withComma(d1); $('#d2').textContent=withComma(d2); $('#d3').textContent=withComma(d3); $('#vI').textContent=withComma(vI);
  $('#d4C').textContent=withComma(d4C); $('#d4D').textContent=withComma(d4D); $('#d5').textContent=withComma(d5); $('#d6').textContent=withComma(d6); $('#vP').textContent=withComma(vP);
  $('#grand').textContent=withComma(grandIT); $('#grandSide').textContent=withComma(grandIT);

  // ③ Before 表
  buildRangeRows($('#rangeTable tbody'),income,grandIT,grandLT);

  // ④ 未使用枠（新制度のみ）
  const aNeed=Math.max(0,capGeneral - Math.min(capGeneral,d1+d2));
  const cNeed=Math.max(0,40000 - d4C);
  const deNeed=Math.max(0,40000 - Math.min(40000,d4D+d5));
  fillUnusedTable({aNow:a,cNow:c,dNow:d,aNeed,cNeed,deNeed,d1_now:d1,d4C_now:d4C,d4D_now:d4D,capGeneral});

  // ⑤ 再計算
  $('#recalcCard').style.display='block';
  $('#recalcBtn').onclick=()=>{
    const ctx={special2026,d1,d2,d3,d4C,d4D,d5,d6,vI,vP,grand:grandIT,unusedA:aNeed,unusedC:cNeed,unusedD:deNeed};
    const afterIT=computeGrandAfterUsingUnused_IT(ctx);
    const afterLT=computeGrandAfterUsingUnused_LT(ctx,a,b,c,d,e);
    buildCompareRangeRows($('#rangeTableCompare tbody'),income,grandIT,afterIT,grandLT,afterLT);
  };

  syncAllMoney();
}

/* 同意モーダル */
function openConsent(){ $('#agreeMask').style.display='block'; $('#agreePanel').style.display='block'; }
function closeConsent(){ $('#agreeMask').style.display='none'; $('#agreePanel').style.display='none'; }
$('#agreeBtn').onclick=closeConsent;

/* 初期ロード */
window.addEventListener('load',()=>{
  bindMailtoFix();          // mailto起動の確実化
  openConsent();            // モーダル
  syncAllMoney(); sumLines();
  calculate();
  $('#calcBtn').addEventListener('click', calculate);
});
