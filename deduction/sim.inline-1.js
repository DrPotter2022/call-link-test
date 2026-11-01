/* ユーティリティ */
const $  = s=>document.querySelector(s);
const $$ = s=>document.querySelectorAll(s);
const numRaw = t => Number(String(t).replace(/[^0-9.-]/g,''))||0;
const withComma = n => n.toLocaleString('ja-JP');

/* ④→⑤で使うフラグ（どの区分を未使用枠充当対象にするか） */
const reuseFlags = {a:true,b:true,c:true,d:true,e:true};

/* 入力系をカンマ付きで整形 */
function syncFormatInput(el){
  const v = numRaw(el.value);
  el.value = withComma(v);
}
function syncAllMoney(){
  $$('.money').forEach(el=>syncFormatInput(el));
  syncFormatInput($('#income'));
}

/* 各区分ごとの証明書カウンタ */
const certCounters = {a:0,b:0,c:0,d:0,e:0};

/* ±ボタン押下処理（年収含む） */
document.addEventListener('click',e=>{
  if(e.target.classList.contains('adjBtn')){
    const tId  = e.target.dataset.target;
    const diff = Number(e.target.dataset.diff);
    const el = document.getElementById(tId);
    if(!el) return;

    if(diff===0){
      el.value = '0';
    }else{
      const now = numRaw(el.value);
      let nv = now + diff;
      if(nv<0) nv=0;
      el.value = withComma(nv);
    }
    sumLines();
  }
});

/* ＋証明書を追加 */
document.addEventListener('click',e=>{
  if(e.target.matches('.addBtn[data-add-cat]')){
    const cat = e.target.dataset.addCat;
    addCertificateRow(cat);
  }
});

function addCertificateRow(cat){
  certCounters[cat] += 1;
  const id = cat + 'Cert' + certCounters[cat];

  const row = document.createElement('div');
  row.className='certRow';

  const inp = document.createElement('input');
  inp.type='text';
  inp.id=id;
  inp.className='mono money certInput';
  inp.placeholder='＋証明書（円）';
  inp.value='0';
  inp.dataset.cat=cat;
  inp.addEventListener('input',()=>{
    syncFormatInput(inp);
    sumLines();
  });
  row.appendChild(inp);

  const diffsForCert=[-10000,-1000,-100,-10,-1,0,1,10,100,1000,10000];
  diffsForCert.forEach(v=>{
    const b=document.createElement('button');
    b.className='pill adjBtn';
    b.dataset.target=id;
    b.dataset.diff=String(v);
    b.textContent=(v>0?`+${withComma(v)}`:withComma(v));
    row.appendChild(b);
  });

  const wrap = document.getElementById(cat+'CertWrap');
  wrap.appendChild(row);

  sumLines();
}

/* A〜Eの合計を再集計し、②の編集ボックスにも反映 */
function sumLines(){
  ['a','b','c','d','e'].forEach(k=>{
    const baseEl = document.getElementById(k+'Sum');
    let total = numRaw(baseEl.value);

    const wrap = document.getElementById(k+'CertWrap');
    if(wrap){
       wrap.querySelectorAll('input.money.certInput').forEach(el=>{
         total += numRaw(el.value);
       });
    }

    const editEl = document.getElementById(k.toUpperCase()+'_edit');
    if(editEl){
      editEl.value = withComma(total);
    }
  });
}

/* ②側の直接編集に対応。編集したらベースをその値で置き換え＋証明書はリセット */
;['A_edit','B_edit','C_edit','D_edit','E_edit'].forEach(id=>{
  const el=document.getElementById(id);
  el.addEventListener('input',()=>{
    syncFormatInput(el);

    const k = id[0].toLowerCase(); // a,b,c,d,e
    const baseEl=document.getElementById(k+'Sum');
    if(!baseEl)return;

    // 証明書側はリセット
    const wrap=document.getElementById(k+'CertWrap');
    if(wrap){
      wrap.querySelectorAll('input.money.certInput').forEach(x=>{
        x.value='0';
      });
    }

    baseEl.value=el.value;
    sumLines();
  });
});

/* 控除ロジック（所得税側） */
function newSchemeI(p, cap=40000){
  if(p<=0) return 0;
  if(p<=20000) return Math.min(p, cap);
  if(p<=40000) return Math.min(Math.floor(p/2)+10000, cap);
  if(p<=80000) return Math.min(Math.floor(p/4)+20000, cap);
  return cap;
}
function oldSchemeII(p){
  if(p<=0) return 0;
  if(p<=25000) return p;
  if(p<=50000) return Math.floor(p/2)+12500;
  if(p<=100000) return Math.floor(p/4)+25000;
  return 50000;
}

/* 住民税側（練馬区式。一般/年金は新旧どちらが有利か、かつ合算上限あり。
   合計の控除上限は7万円。そこに10%を掛けて③の住民税軽減目安を出す。 */
function newSchemeLT(p){
  if(p<=0) return 0;
  if(p<=12000) return p;
  if(p<=32000) return Math.floor(p/2)+6000;
  if(p<=56000) return Math.floor(p/4)+14000;
  return 28000;
}
function oldSchemeLT(p){
  if(p<=0) return 0;
  if(p<=15000) return p;
  if(p<=40000) return Math.floor(p/2)+7500;
  if(p<=70000) return Math.floor(p/4)+17500;
  return 35000;
}
function bestLTForGeneral(newP, oldP){
  const n=newSchemeLT(newP);
  const o=oldSchemeLT(oldP);
  return Math.max(o, n, Math.min(28000, o+n));
}
function bestLTForPension(newP, oldP){
  const n=newSchemeLT(newP);
  const o=oldSchemeLT(oldP);
  return Math.max(o, n, Math.min(28000, o+n));
}
function computeResidentGrand(aP,bP,cP,dP,eP){
  const gen = bestLTForGeneral(aP,bP); // 一般(A/B)
  const med = newSchemeLT(cP);         // 介護医療(C)
  const pen = bestLTForPension(dP,eP); // 年金(D/E)
  return Math.min(70000, gen+med+pen);
}

/* ③ 年収レンジ別テーブル */
function buildRangeRows(tbody, income, grandIT, grandLT){
  const ranges=[
    {label:'〜4,000,000',rate:0.05, match:x=>x<=4000000},
    {label:'4,000,001〜6,000,000',rate:0.10, match:x=>x>=4000001 && x<=6000000},
    {label:'6,000,001〜9,000,000',rate:0.20, match:x=>x>=6000001 && x<=9000000},
    {label:'9,000,001〜10,000,000',rate:0.23, match:x=>x>=9000001 && x<=10000000},
    {label:'10,000,001〜',rate:0.33, match:x=>x>=10000001},
  ];
  tbody.innerHTML='';

  // 住民税分の軽減額(概算)→ 住民税控除額(上限7万円)×10%
  const ltTax=Math.round(Math.min(70000,grandLT)*0.10);

  ranges.forEach(r=>{
    const it=Math.round(grandIT*r.rate);
    const tr=document.createElement('tr');
    tr.innerHTML=`
      <td>${r.label}</td>
      <td class="r">${(r.rate*100).toFixed(0)}%</td>
      <td class="r">${withComma(it)}</td>
      <td class="r">${withComma(ltTax)}</td>
      <td class="r sum">${withComma(it+ltTax)}</td>
    `;
    if(r.match(income)){
      tr.classList.add('rangeHighlight');
    }
    tbody.appendChild(tr);
  });
}

/* ④ 未使用枠テーブル表示
   reuseFlags.[a,b,c,d,e] にチェックボックスの状態を入れる */
function fillUnusedTable(unused){
  const tb=$('#unusedTable tbody');
  tb.innerHTML='';

  const rows=[
    {cat:'a', label:'一般生命保険料（新制度）', now:unused.aNow, year:unused.aNeed, mon:Math.ceil(unused.aNeed/12)},
    {cat:'b', label:'一般生命保険料（旧制度）', now:unused.bNow, year:0,          mon:0},
    {cat:'c', label:'介護医療保険料（新のみ）', now:unused.cNow, year:unused.cNeed, mon:Math.ceil(unused.cNeed/12)},
    {cat:'d', label:'新 個人年金保険料',       now:unused.dNow, year:0,          mon:0},
    {cat:'e', label:'旧 個人年金保険料',       now:unused.eNow, year:unused.deNeed, mon:Math.ceil(unused.deNeed/12)},
  ];

  rows.forEach(r=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`
      <td>
        <label style="display:flex;align-items:center;gap:6px;">
          <input type="checkbox" class="reuseChk" data-cat="${r.cat}" ${reuseFlags[r.cat]?'checked':''}/>
          <span>${r.label}</span>
        </label>
      </td>
      <td class="r">${withComma(r.now)}</td>
      <td class="r">${withComma(r.year)}</td>
      <td class="r">${r.mon?withComma(r.mon):'-'}</td>`;
    tb.appendChild(tr);
  });
}

/* チェック変更時、reuseFlagsを更新 */
document.addEventListener('change',e=>{
  if(e.target.classList.contains('reuseChk')){
    const cat=e.target.dataset.cat;
    reuseFlags[cat]=e.target.checked;
  }
});

/* 控除額→追加で必要な保険料を逆算する近似関数（新制度側） */
function inverseNewSchemeI(d, cap=40000){
  d=Math.max(0,Math.min(d,cap));
  if(d<=20000) return d;
  if(d<=30000) return 2*(d-10000);
  if(d<=40000) return 4*(d-20000);
  return 80000;
}

/* ⑤ After(未使用枠を追加) の所得税側合計控除額 */
function computeGrandAfterUsingUnused_IT(ctx){
  const capGeneral=ctx.special2026?60000:40000;

  const unusedAtoUse = reuseFlags.a ? ctx.unusedA : 0;
  const d3_after=Math.min(capGeneral,ctx.d1+ctx.d2+unusedAtoUse);
  const vI_after=Math.max(ctx.d2,d3_after);

  const unusedCtoUse = reuseFlags.c ? ctx.unusedC : 0;
  const d4C_after=Math.min(40000,ctx.d4C+unusedCtoUse);

  const unusedDtoUse = reuseFlags.d ? ctx.unusedD : 0;
  const d6_after =Math.min(40000,ctx.d4D+ctx.d5+unusedDtoUse);
  const vP_after=Math.max(ctx.d5,d6_after);

  return vI_after+d4C_after+vP_after;
}

/* ⑤ After(未使用枠を追加) の住民税側合計控除額 */
function computeGrandAfterUsingUnused_LT(ctx,aP,bP,cP,dP,eP){
  const capGeneral=ctx.special2026?60000:40000;

  // A(新) 分
  const d1_target=Math.min(
    capGeneral,
    ctx.d1 + (reuseFlags.a ? ctx.unusedA : 0)
  );
  const aP_after =Math.max(aP,inverseNewSchemeI(d1_target,capGeneral));

  // C(介護医療)
  const d4C_target=Math.min(
    40000,
    ctx.d4C + (reuseFlags.c ? ctx.unusedC : 0)
  );
  const cP_after =Math.max(cP,inverseNewSchemeI(d4C_target,40000));

  // D(新個人年金)
  const d4D_target=Math.min(
    40000,
    ctx.d4D + (reuseFlags.d ? ctx.unusedD : 0)
  );
  const dP_after =Math.max(dP,inverseNewSchemeI(d4D_target,40000));

  // 住民税トータル
  return computeResidentGrand(aP_after,bP,cP_after,dP_after,eP);
}

/* ⑤ 比較テーブル（Before / After） */
function buildCompareRangeRows(
  tbody,
  income,
  beforeGrandIT,
  afterGrandIT,
  beforeGrandLT,
  afterGrandLT
){
  const ranges=[
    {label:'〜4,000,000',rate:0.05, match:x=>x<=4000000},
    {label:'4,000,001〜6,000,000',rate:0.10, match:x=>x>=4000001 && x<=6000000},
    {label:'6,000,001〜9,000,000',rate:0.20, match:x=>x>=6000001 && x<=9000000},
    {label:'9,000,001〜10,000,000',rate:0.23, match:x=>x>=9000001 && x<=10000000},
    {label:'10,000,001〜',rate:0.33, match:x=>x>=10000001},
  ];
  tbody.innerHTML='';

  const ltB=Math.round(Math.min(70000,beforeGrandLT)*0.10);
  const ltA=Math.round(Math.min(70000,afterGrandLT )*0.10);

  ranges.forEach(r=>{
    const itB=Math.round(beforeGrandIT*r.rate);
    const itA=Math.round(afterGrandIT *r.rate);
    const tr=document.createElement('tr');
    tr.innerHTML=`
      <td>${r.label}</td>
      <td class="r">${withComma(itB)}</td>
      <td class="r">${withComma(ltB)}</td>
      <td class="r sum">${withComma(itB+ltB)}</td>
      <td class="r">${withComma(itA)}</td>
      <td class="r">${withComma(ltA)}</td>
      <td class="r sum">${withComma(itA+ltA)}</td>`;
    if(r.match(income)){
      tr.classList.add('rangeHighlight');
    }
    tbody.appendChild(tr);
  });
}

/* メイン計算 */
function calculate(){
  const income=numRaw($('#income').value);

  const a=numRaw($('#A_edit').value||$('#aSum').value);
  const b=numRaw($('#B_edit').value||$('#bSum').value);
  const c=numRaw($('#C_edit').value||$('#cSum').value);
  const d=numRaw($('#D_edit').value||$('#dSum').value);
  const e=numRaw($('#E_edit').value||$('#eSum').value);

  const special2026=$('#sp2026').checked;
  const capGeneral =special2026?60000:40000;

  // 上限ラベル更新
  $('#capGenLbl').textContent  =withComma(capGeneral);
  $('#capGenLbl2').textContent =withComma(capGeneral);
  $('#capGenSide').textContent =withComma(capGeneral);

  // 所得税側の控除
  const d1=newSchemeI(a,capGeneral);     //①
  const d2=oldSchemeII(b);               //②
  const d3=Math.min(capGeneral,d1+d2);   //③
  const vI=Math.max(d2,d3);              //㋑

  const d4C=newSchemeI(c,40000);         //㋺
  const d4D=newSchemeI(d,40000);         //④
  const d5 =oldSchemeII(e);              //⑤
  const d6 =Math.min(40000,d4D+d5);      //⑥
  const vP =Math.max(d5,d6);             //㋩

  const grandIT=vI+d4C+vP;               // 所得税側の合計控除額

  // 住民税側（合計7万円上限）
  const grandLT=computeResidentGrand(a,b,c,d,e);

  // ② 表示更新
  $('#d1').textContent   =withComma(d1);
  $('#d2').textContent   =withComma(d2);
  $('#d3').textContent   =withComma(d3);
  $('#vI').textContent   =withComma(vI);
  $('#d4C').textContent  =withComma(d4C);
  $('#d4D').textContent  =withComma(d4D);
  $('#d5').textContent   =withComma(d5);
  $('#d6').textContent   =withComma(d6);
  $('#vP').textContent   =withComma(vP);
  $('#grand').textContent=withComma(grandIT);
  $('#grandSide').textContent=withComma(grandIT);

  // ③ 年収レンジ別テーブル
  buildRangeRows($('#rangeTable tbody'),income,grandIT,grandLT);

  // ④ 未使用枠の逆算
  // 一般（新）：capGeneralに対し d1+d2 で取り切れてない分
  const aNeed =Math.max(0,capGeneral - Math.min(capGeneral,d1+d2));
  // 介護医療(C)：40,000に対して d4C
  const cNeed =Math.max(0,40000     - d4C);
  // 年金(D/E)：40,000に対して d4D+d5
  const deNeed=Math.max(0,40000     - Math.min(40000,d4D+d5));

  fillUnusedTable({
    aNow:a,bNow:b,cNow:c,dNow:d,eNow:e,
    aNeed,cNeed,deNeed
  });

  // ⑤ 再計算（After）
  $('#recalcCard').style.display='block';
  $('#recalcBtn').onclick=()=>{
    const ctx={
      special2026,
      d1,d2,d3,d4C,d4D,d5,d6,vI,vP,
      grand:grandIT,
      unusedA:aNeed,
      unusedC:cNeed,
      unusedD:deNeed
    };

    const afterGrandIT=computeGrandAfterUsingUnused_IT(ctx);
    const afterGrandLT=computeGrandAfterUsingUnused_LT(ctx,a,b,c,d,e);

    buildCompareRangeRows(
      $('#rangeTableCompare tbody'),
      income,
      grandIT,
      afterGrandIT,
      grandLT,
      afterGrandLT
    );
  };

  syncAllMoney();
}

/* 同意モーダル */
function openConsent(){
  $('#agreeMask').style.display='block';
  $('#agreePanel').style.display='block';
}
function closeConsent(){
  $('#agreeMask').style.display='none';
  $('#agreePanel').style.display='none';
}
$('#agreeBtn').onclick=closeConsent;

/* イベント */
$('#calcBtn').addEventListener('click',calculate);

/* 初期ロード処理 */
window.addEventListener('load',()=>{
  openConsent();   // 初回に同意ダイアログを表示
  syncAllMoney();  // カンマ整形
  sumLines();      // A〜E合計の同期
  calculate();     // 全体計算
});
