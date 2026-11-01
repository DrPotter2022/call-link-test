const lifeTable = {
  male: {0:74,1:74,2:73,3:72,4:71,5:70,6:69,7:68,8:67,9:66,10:65,11:64,12:63,13:62,14:61,15:60,16:59,17:58,18:57,19:56,20:55,21:54,22:53,23:52,24:51,25:50,26:50,27:49,28:48,29:47,30:46,31:45,32:44,33:43,34:42,35:41,36:40,37:39,38:38,39:37,40:36,41:35,42:34,43:33,44:32,45:32,46:31,47:30,48:29,49:28,50:27,51:26,52:25,53:25,54:24,55:23,56:22,57:21,58:20,59:20,60:19,61:18,62:17,63:17,64:16,65:15,66:14,67:14,68:13,69:12,70:12,71:11,72:10,73:10,74:9,75:8,76:8,77:7,78:7,79:6,80:6,81:6,82:5,83:5,84:4,85:4,86:4,87:4,88:3,89:3,90:3,91:3,92:2,93:2,94:2,95:2,96:2,97:1},
  female:{0:80,1:79,2:78,3:77,4:77,5:76,6:75,7:74,8:73,9:72,10:71,11:70,12:69,13:68,14:67,15:66,16:65,17:64,18:63,19:62,20:61,21:60,22:59,23:58,24:57,25:56,26:55,27:54,28:53,29:52,30:51,31:50,32:49,33:48,34:47,35:46,36:45,37:44,38:43,39:42,40:41,41:40,42:39,43:38,44:37,45:36,46:36,47:35,48:34,49:33,50:32,51:31,52:30,53:29,54:28,55:27,56:26,57:25,58:25,59:24,60:23,61:22,62:21,63:20,64:19,65:18,66:18,67:17,68:16,69:15,70:14,71:14,72:13,73:12,74:11,75:11,76:10,77:9,78:9,79:8,80:8,81:7,82:7,83:6,84:6,85:5,86:5,87:4,88:4,89:4,90:3,91:3,92:3,93:3,94:2,95:2,96:2,97:1}
};

const nf = new Intl.NumberFormat('ja-JP');
function parseMoney(v){ const s=String(v??'').replace(/[,\s　]/g,''); if(s==='')return null; const n=Number(s); return Number.isFinite(n)?n:null; }
function parseIntSafe(v){ const n=parseMoney(v); if(n===null) return null; return Math.trunc(n); }
function yen(n){ if(!Number.isFinite(n)) return '—'; return nf.format(Math.round(n)); }
function roundUp2(x){ return Math.ceil((x-1e-12)*100)/100; }
function showMsg(t){ const el=document.getElementById('msg'); if(!t){el.style.display='none';el.textContent='';return;} el.style.display='block'; el.textContent=t; }
function formatCommaInput(el){ const n=parseMoney(el.value); if(n===null)return; el.value=nf.format(n); }
['premium','payment','death'].forEach(id=> document.getElementById(id).addEventListener('blur', e=>formatCommaInput(e.target)));
function getLifeYears(sex, age){ if(age>=97) return 1; return lifeTable[sex]?.[age]; }

function closeAllPanels(exceptId=null){
  document.querySelectorAll('.panel').forEach(p=>{ if(p.id!==exceptId) p.classList.remove('open'); });
  document.querySelectorAll('.iconBtn').forEach(b=>{ const t=b.getAttribute('data-toggle'); if(t!==exceptId) b.classList.remove('open'); });
}
function togglePanel(panelId, btn){
  const p=document.getElementById(panelId);
  const isOpen=p.classList.contains('open');
  closeAllPanels(panelId);
  if(isOpen){ p.classList.remove('open'); btn.classList.remove('open'); }
  else{ p.classList.add('open'); btn.classList.add('open'); }
}
document.addEventListener('click', (e)=>{
  const btn=e.target.closest('.iconBtn');
  if(btn){ e.preventDefault(); togglePanel(btn.getAttribute('data-toggle'), btn); return; }
  if(e.target.closest('.panel')) return;
  closeAllPanels(null);
});

function renderAgeDecade(decadeBoxId, detailBoxId, targetInputId){
  const decadeBox=document.getElementById(decadeBoxId);
  const detailBox=document.getElementById(detailBoxId);
  decadeBox.innerHTML=''; detailBox.innerHTML='';
  for(let d=0; d<=90; d+=10){
    const b=document.createElement('button');
    b.type='button'; b.className='chip'; b.textContent=String(d);
    b.addEventListener('click', ()=>renderAgeDetail(detailBox,d,targetInputId));
    decadeBox.appendChild(b);
  }
  const clear=document.createElement('button');
  clear.type='button'; clear.className='chip secondary'; clear.textContent='クリア';
  clear.addEventListener('click', ()=>{ document.getElementById(targetInputId).value=''; detailBox.innerHTML=''; });
  decadeBox.appendChild(clear);
}
function renderAgeDetail(detailBox, decade, targetInputId){
  detailBox.innerHTML='';
  const start=decade, end=Math.min(decade+9,90);
  for(let a=start; a<=end; a++){
    const b=document.createElement('button');
    b.type='button'; b.className='chip good'; b.textContent=String(a);
    b.addEventListener('click', ()=>{ document.getElementById(targetInputId).value=String(a); });
    detailBox.appendChild(b);
  }
}

function attachAutoRepeat(button, actionFn, opt={}){
  const delay=opt.delay??350, interval=opt.interval??90;
  let t1=null,t2=null;
  const clear=()=>{ if(t1)clearTimeout(t1); if(t2)clearInterval(t2); t1=null; t2=null; };
  const start=(e)=>{ e.preventDefault(); actionFn(); clear(); t1=setTimeout(()=>{ t2=setInterval(actionFn, interval); }, delay); };
  const stop=()=>clear();
  button.addEventListener('pointerdown', start);
  button.addEventListener('pointerup', stop);
  button.addEventListener('pointercancel', stop);
  button.addEventListener('pointerleave', stop);
  button.addEventListener('touchend', stop);
}

function MoneyPad(inputId, addBoxId, adjustBoxId, options){
  this.inputId=inputId; this.addBoxId=addBoxId; this.adjustBoxId=adjustBoxId;
  this.minValue=options?.minValue??0;
  this.lastStep=options?.defaultStep??10_000;
  this.signMode=options?.signMode??false;
  this.sign=+1;
  this.steps=options?.steps??[
    {base:'+1万', v:10_000},{base:'+10万', v:100_000},{base:'+100万', v:1_000_000},{base:'+1000万', v:10_000_000},
  ];
}
MoneyPad.prototype.getValue=function(){ const v=parseMoney(document.getElementById(this.inputId).value); return v??0; };
MoneyPad.prototype.setValue=function(v){ const n=Math.max(this.minValue, v); document.getElementById(this.inputId).value=nf.format(n); };
MoneyPad.prototype.apply=function(delta){ this.setValue(this.getValue()+delta); };
MoneyPad.prototype.add=function(step){ this.lastStep=step; const delta=(this.signMode?this.sign:+1)*step; this.apply(delta); };
MoneyPad.prototype.adjust=function(sign){ if(this.signMode) this.setSign(sign); this.apply(sign*this.lastStep); };
MoneyPad.prototype.clear=function(){ document.getElementById(this.inputId).value=''; };
MoneyPad.prototype.setSign=function(sign){
  this.sign=(sign>=0)?+1:-1;
  if(!this.signMode) return;
  this.updateStepButtons();
  this.updateSignButtons();
};
MoneyPad.prototype.updateStepButtons=function(){
  if(!this._stepButtons) return;
  const prefix=(this.sign===1)?'+':'-';
  this._stepButtons.forEach(({btn,labelCore})=>{ btn.textContent=`${prefix}${labelCore}`; });
};
MoneyPad.prototype.updateSignButtons=function(){
  if(!this._signBtns) return;
  const {minusBtn, plusBtn}=this._signBtns;
  if(this.sign===-1){ minusBtn.classList.add('active'); plusBtn.classList.remove('active'); }
  else{ plusBtn.classList.add('active'); minusBtn.classList.remove('active'); }
};
MoneyPad.prototype.render=function(){
  const addBox=document.getElementById(this.addBoxId);
  const adjBox=document.getElementById(this.adjustBoxId);
  addBox.innerHTML=''; adjBox.innerHTML='';
  this._stepButtons=[];
  this.steps.forEach(s=>{
    const btn=document.createElement('button');
    btn.type='button'; btn.className='chip';
    const labelCore=(s.base.startsWith('+')||s.base.startsWith('-'))?s.base.slice(1):s.base;
    btn.textContent=this.signMode?((this.sign===1?'+':'-')+labelCore):s.base;
    attachAutoRepeat(btn, ()=>this.add(s.v));
    addBox.appendChild(btn);
    this._stepButtons.push({btn,labelCore});
  });
  const clear=document.createElement('button');
  clear.type='button'; clear.className='chip secondary'; clear.textContent='クリア';
  clear.addEventListener('click', ()=>this.clear());
  addBox.appendChild(clear);

  const minus=document.createElement('button');
  minus.type='button'; minus.className='chip danger'; minus.textContent='－';
  attachAutoRepeat(minus, ()=>this.adjust(-1));
  adjBox.appendChild(minus);

  const plus=document.createElement('button');
  plus.type='button'; plus.className='chip good'; plus.textContent='＋';
  attachAutoRepeat(plus, ()=>this.adjust(+1));
  adjBox.appendChild(plus);

  this._signBtns={minusBtn:minus, plusBtn:plus};
  this.updateSignButtons();
};

function renderStartAgeSpecial(){
  const box=document.getElementById('startAgeSpecial');
  box.innerHTML='';
  const plus1=document.createElement('button');
  plus1.type='button'; plus1.className='chip good'; plus1.textContent='契約年齢+1歳';
  plus1.addEventListener('click', ()=>{
    const a=parseIntSafe(document.getElementById('age').value);
    if(a!==null) document.getElementById('startAge').value=String(a+1);
  });
  box.appendChild(plus1);

  const same=document.createElement('button');
  same.type='button'; same.className='chip secondary'; same.textContent='契約年齢と同じ';
  same.addEventListener('click', ()=>{
    const a=parseIntSafe(document.getElementById('age').value);
    if(a!==null) document.getElementById('startAge').value=String(a);
  });
  box.appendChild(same);

  const clear=document.createElement('button');
  clear.type='button'; clear.className='chip secondary'; clear.textContent='クリア';
  clear.addEventListener('click', ()=>{ document.getElementById('startAge').value=''; document.getElementById('startAgeDetail').innerHTML=''; });
  box.appendChild(clear);
}
function renderDeathTopButtons(){
  const box=document.getElementById('deathTop');
  box.innerHTML='';
  const same=document.createElement('button');
  same.type='button'; same.className='chip good'; same.textContent='払込保険料と同額';
  same.addEventListener('click', ()=>{
    const p=parseMoney(document.getElementById('premium').value);
    if(p!==null) document.getElementById('death').value=nf.format(p);
  });
  box.appendChild(same);
}

function calc(){
  showMsg(null);
  const age=parseIntSafe(document.getElementById('age').value);
  const sex=document.getElementById('sex').value;
  const premium=parseMoney(document.getElementById('premium').value);
  const payment=parseMoney(document.getElementById('payment').value);
  const deathIn=parseMoney(document.getElementById('death').value);
  const startAgeIn=parseIntSafe(document.getElementById('startAge').value);

  if(age===null||age<0||age>130) return showMsg('「契約時年齢」を正しく入力してください。');
  if(premium===null||premium<0) return showMsg('「払込保険料相当額」を正しく入力してください。');
  if(payment===null||payment<0) return showMsg('「定期支払額」を正しく入力してください。');

  const startAge=(startAgeIn===null)?(age+1):startAgeIn;
  if(startAge<0||startAge>130) return showMsg('「第1回定期支払日時点の年齢」を正しく入力してください。');

  const death=(deathIn===null)?premium:deathIn;

  const life=getLifeYears(sex,startAge);
  if(!life && life!==0) return showMsg(`余命年数表に存在しない年齢です：${startAge}歳`);

  const total=payment*life;
  const denom=total+death;
  if(denom<=0) return showMsg('計算の分母（受取予定総額＋死亡保険金額）が0以下です。入力を見直してください。');

  const ratioRaw=premium/denom;
  const ratio=roundUp2(ratioRaw);

  const expense=Math.round(payment*ratio);
  const taxable=Math.round(payment-expense);

  let withholdText='源泉徴収なし';
  if(taxable>=250000){
    const w=Math.floor(taxable*0.1021);
    withholdText=`源泉徴収あり（概算）：${yen(w)} 円（10.21%）`;
  }

  document.getElementById('life').textContent=String(life);
  document.getElementById('total').textContent=yen(total);
  document.getElementById('ratio').innerHTML=`${ratio.toFixed(2)}<span class="subline">（端数前：${ratioRaw.toFixed(6)}）</span>`;
  document.getElementById('expense').textContent=yen(expense);
  document.getElementById('taxable').textContent=yen(taxable);
  document.getElementById('withhold').textContent=withholdText;
}

function resetAll(){
  document.getElementById('age').value='60';
  document.getElementById('sex').value='female';
  document.getElementById('premium').value='10,000,000';
  document.getElementById('payment').value='430,000';
  document.getElementById('death').value='';
  document.getElementById('startAge').value='';
  ['life','total','ratio','expense','taxable','withhold'].forEach(id=>document.getElementById(id).textContent='—');
  showMsg(null);
  closeAllPanels();
}

const CONSENT_KEY='teiki_keihi_consent_v9';
function openConsent(){ document.getElementById('overlay').style.display='flex'; }
function closeConsent(){ document.getElementById('overlay').style.display='none'; }
document.getElementById('agreeBtn').addEventListener('click', ()=>{ localStorage.setItem(CONSENT_KEY,'yes'); closeConsent(); });
document.getElementById('denyBtn').addEventListener('click', ()=>alert('同意がないため利用できません。'));

const premiumPad=new MoneyPad('premium','premiumAdd','premiumAdjust',{minValue:0, signMode:true});
const deathPad  =new MoneyPad('death','deathAdd','deathAdjust',{minValue:0, signMode:true});
const paymentPad=new MoneyPad('payment','paymentAdd','paymentAdjust',{
  minValue:0, signMode:true,
  steps:[
    {base:'+1万', v:10_000},{base:'+10万', v:100_000},{base:'+100万', v:1_000_000},{base:'+1000万', v:10_000_000},
    {base:'+1十', v:10},{base:'+1百', v:100},{base:'+1千', v:1000},
  ]
});

premiumPad.render(); paymentPad.render(); deathPad.render();
renderAgeDecade('ageDecade','ageDetail','age');
renderAgeDecade('startAgeDecade','startAgeDetail','startAge');
renderStartAgeSpecial();
renderDeathTopButtons();

document.getElementById('calcBtn').addEventListener('click', ()=>{
  calc();
  ['premium','payment','death'].forEach(id=>formatCommaInput(document.getElementById(id)));
});
document.getElementById('resetBtn').addEventListener('click', resetAll);

resetAll();
if(localStorage.getItem(CONSENT_KEY)!=='yes') openConsent();
else closeConsent();
