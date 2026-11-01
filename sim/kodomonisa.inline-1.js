/* ===================== 計算エンジン ===================== */
var Engine=(function(){
function dn(y,m,d){return Date.UTC(y,m-1,d)/864e5;}
function isLeap(y){return y%4===0&&(y%100!==0||y%400===0);}
function ymd(t){var d=new Date(t*864e5);return {y:d.getUTCFullYear(),m:d.getUTCMonth()+1,d:d.getUTCDate()};}
function monthlyRate(a,f,method){
  if(method==='simple')return (a-f)/12;
  var b=1+a-f;if(b<=0)b=1e-9;return Math.pow(b,1/12)-1;
}
var CH_YEAR=600000,CH_LIFE=6000000,AD_YEAR=3600000,ALL_LIFE=18000000;

function simulate(P,pat){
  var errors=[],warnings=[],notes=[],noteSeen={};
  function note(k,msg){if(!noteSeen[k]){noteSeen[k]=1;notes.push(msg);}}
  var by=P.by,bm=P.bm,bd=P.bd;
  function bday(a){var y=by+a;return dn(y,bm,(bm===2&&bd===29&&!isLeap(y))?28:bd);}
  var tBirth=bday(0),t18=bday(18);
  var rmC=monthlyRate(P.rChild,P.fee,P.method),rmA=monthlyRate(P.rAdult,P.fee,P.method);
  var buys=[],y,m,t;
  function stopped(ix){return (P.stops||[]).some(function(c){return ix>=c.from&&(c.to==null||ix<=c.to);});}

  /* こども期の買付予定 */
  var b15m=(bm<=3?(by+15):(by+16))*12+2, b3m=(by*12+bm-1)+36;
  y=P.sy;m=P.sm;
  while(true){
    t=dn(y,m,P.buyDay);
    if(t>=t18)break;
    if(t>=tBirth&&!stopped(y*12+m-1)){
      var mo=0,lu=0,mi=y*12+m-1;
      if(pat===0){mo=P.monthly;(P.changes||[]).forEach(function(c){if(mi>=c.from&&(c.to==null||mi<=c.to))mo=c.amount;});lu=(m===P.lumpM)?P.lumpAmt:0;}
      else if(pat===1){mo=(mi<=b3m)?15000:(mi<=b15m?10000:0);}
      else mo=50000;
      if(mo+lu>0)buys.push({t:t,y:y,m:m,ph:'c',mo:mo,lu:lu,fm:0});
    }
    m++;if(m>12){m=1;y++;}
  }

  /* 18歳以降の追加投資予定 */
  var B=P.B;
  if(B.on){
    var sa=Math.max(18,B.startAge),ea=B.endAge;
    if(ea<=sa)errors.push('B：終了年齢は開始年齢より大きくしてください。');
    else{
      var eligY=by+18+((bm===1&&bd===1)?0:1);
      var effStart=Math.max(bday(sa),dn(eligY,1,1));
      if(effStart>bday(sa)){var es=ymd(effStart);note('elig','成人NISAの追加投資は「その年の1月1日時点で18歳以上」の年から可能なため、開始を'+es.y+'年1月に繰り下げました。');}
      var endT=bday(ea),s=ymd(effStart);y=s.y;m=s.m;
      while(true){
        t=dn(y,m,P.buyDay);
        if(t>=endT)break;
        if(t>=effStart){
          var lu2=(m===B.lumpM)?B.lumpAmt:0;
          if(B.monthly+lu2>0)buys.push({t:t,y:y,m:m,ph:'a',mo:B.monthly,lu:lu2,fm:0});
        }
        m++;if(m>12){m=1;y++;}
      }
    }
  }
  buys.sort(function(a,b){return a.t-b.t;});
  var fut={},i;
  for(i=buys.length-1;i>=0;i--){var k=buys[i].ph+buys[i].y;buys[i].fm=fut[k]||0;fut[k]=(fut[k]||0)+buys[i].mo;}

  /* イベント */
  var ev=[],n=0;
  ev.push({t:t18,pr:-1,type:'sw',n:n++});
  buys.forEach(function(b){ev.push({t:b.t,pr:0,type:'buy',d:b,n:n++});});
  var wAges=[];
  if(P.C.on){
    P.C.items.forEach(function(w){
      if(!(w.amount>0))return;
      if(w.age<12){errors.push('C：'+w.age+'歳の取崩しは不可です（払出しは12歳の誕生日以降）。');return;}
      if(w.age>100){errors.push('C：取崩し年齢は100歳以下にしてください。');return;}
      ev.push({t:bday(w.age),pr:1,type:'wd',d:w,n:n++});wAges.push(w.age);
    });
  }
  for(i=0;i<=100;i++)ev.push({t:bday(i),pr:2,type:'snap',d:i,n:n++});
  ev.sort(function(a,b){return a.t-b.t||a.pr-b.pr||a.n-b.n;});

  var V=0,basis=0,cumBuy=0,cumW=0,acc=0,accW=0,rm=rmC,last=ev[0].t;
  var cy={},ay={},cTot=0,aTot=0,rows=[],purchases=[],reach600=null,reach1800=null;
  ev.forEach(function(e){
    if(e.t>last){V*=Math.pow(1+rm,(e.t-last)*12/365);last=e.t;}
    if(e.type==='sw'){rm=rmA;}
    else if(e.type==='buy'){
      var b=e.d,child=b.ph==='c',cap=child?CH_YEAR:AD_YEAR,ymap=child?cy:ay;
      var used=ymap[b.y]||0,life=child?CH_LIFE-cTot:ALL_LIFE-cTot-aTot;
      var mo2=Math.max(0,Math.min(b.mo,cap-used,life));
      if(mo2<b.mo){var lifeHit=life<=cap-used;note('m'+b.ph+(lifeHit?'L':'Y'),(child?'こども期':'18歳以降')+'の'+b.y+'年'+b.m+'月：'+(lifeHit?'生涯投資枠':'年間投資枠')+'の上限により毎月積立を減額（以降も同様に自動調整）。');}
      var lu2=Math.max(0,Math.min(b.lu,cap-used-mo2-b.fm,life-mo2));
      if(lu2<b.lu)note('l'+b.ph,(child?'こども期':'18歳以降')+'の'+b.y+'年：年間枠（'+(cap/10000)+'万円）または生涯枠の上限のため、年1回の一括追加を'+lu2.toLocaleString('ja-JP')+'円に減額（以降も同様に自動調整）。');
      var amt=mo2+lu2;
      if(amt>0){
        ymap[b.y]=used+amt;
        if(child)cTot+=amt;else aTot+=amt;
        V+=amt;basis+=amt;cumBuy+=amt;acc+=amt;
        purchases.push({t:e.t,amount:amt,ph:b.ph});
        if(!reach600&&cTot>=CH_LIFE)reach600={t:e.t};
        if(!reach1800&&cTot+aTot>=ALL_LIFE)reach1800={t:e.t};
      }
    }else if(e.type==='wd'){
      var w=e.d;
      if(V<=1e-9){warnings.push('C：'+w.age+'歳時点の残高がないため取崩せません。');}
      else{
        var take=Math.min(w.amount,V);
        if(w.amount>V+1e-6)warnings.push('C：'+w.age+'歳の取崩し額が残高（'+Math.round(V).toLocaleString('ja-JP')+'円）を超えるため、残高全額を取崩しました。');
        var ratio=take/V;basis*=(1-ratio);V-=take;cumW+=take;accW+=take;
      }
    }else{
      rows.push({age:e.d,year:by+e.d,t:e.t,buy:acc,principal:cumBuy,withdraw:accW,cumW:cumW,value:V,basis:basis,gain:V-basis});
      acc=0;accW=0;
    }
  });
  function ageAt(t){
    var a=ymd(t),age=a.y-by,mon=a.m-bm;
    if(a.d<bd)mon--;
    if(mon<0){age--;mon+=12;}
    return {years:age,months:mon,y:a.y,m:a.m};
  }
  return {rows:rows,purchases:purchases,errors:errors,warnings:warnings,notes:notes,wAges:wAges,
    reach600:reach600?ageAt(reach600.t):null,reach1800:reach1800?ageAt(reach1800.t):null,
    rmC:rmC,rmA:rmA,cTot:cTot,aTot:aTot,totalBuy:cumBuy};
}
return {simulate:simulate,monthlyRate:monthlyRate,dn:dn};
})();

/* ===================== UI ===================== */
if(typeof document!=='undefined'){(function(){
var $=function(id){return document.getElementById(id);};
var NAMES=['①任意金額','②児童手当','③MAX'];
var COLS=['#2563eb','#16a34a','#ea580c'];
var tab=1,res=null,cur={1:0,2:0},last=null;
var fmtYen=function(v){return Math.round(v).toLocaleString('ja-JP');};
var yen=function(v){return fmtYen(v)+'円';};
var man=function(v){return (v/10000).toLocaleString('ja-JP',{minimumFractionDigits:1,maximumFractionDigits:1})+'万円';};
var esc=function(s){return String(s).replace(/[&<>]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c];});};
function pv(el){var v=parseFloat(String(el.value).replace(/[,\s，]/g,''));return isFinite(v)?v:NaN;}
function pn(id,def){var v=pv($(id));return isNaN(v)?def:v;}

/* --- 初期化 --- */
(function init(){
  var i,o;
  for(i=1;i<=28;i++){o=document.createElement('option');o.value=i;o.textContent='毎月'+i+'日';$('buyDay').appendChild(o);}
  ['lumpM','bLumpM'].forEach(function(id){for(var k=1;k<=12;k++){var q=document.createElement('option');q.value=k;q.textContent=k+'月';$(id).appendChild(q);}});
  addRow(2,18,1000000);
})();

/* --- スライダー連動 --- */
function pair(rid,nid,money,step){
  var r=$(rid),n=$(nid);
  r.addEventListener('input',function(){n.value=money?fmtYen(r.value):r.value;});
  n.addEventListener('input',function(){var v=pv(n);if(!isNaN(v))r.value=v;});
}
pair('monthlyR','monthly',true);pair('rChildR','rChild',false);pair('rAdultR','rAdult',false);pair('dispAgeR','dispAge',false);
function ymIdx(v){var a=String(v).split('-').map(Number);return a.length>=2&&a[0]?a[0]*12+a[1]-1:NaN;}
function ymStr(n){return Math.floor(n/12)+'-'+('0'+(n%12+1)).slice(-2);}
function birthIdx(){var b=$('birth').value.split('-').map(Number);return b[0]?b[0]*12+b[1]-1:NaN;}
function agePair(box){
  var ym=box.querySelector('.ym'),y=box.querySelector('.ay'),m=box.querySelector('.am');
  function fromAge(){var b=birthIdx();if(isNaN(b)||y.value===''||m.value==='')return;ym.value=ymStr(b+Math.floor(+y.value)*12+Math.floor(+m.value));}
  box.toAge=function(){var b=birthIdx(),n=ymIdx(ym.value);
    if(isNaN(n)){y.value='';m.value='';return;}if(isNaN(b))return;n-=b;if(n<0)n=0;y.value=Math.floor(n/12);m.value=n%12;};
  y.addEventListener('input',fromAge);m.addEventListener('input',fromAge);ym.addEventListener('input',box.toAge);
}
function pbox(cls,label){return '<div class="f pb '+cls+'"><label>'+label+'<input type="month" class="ym"></label><span class="sl"><input type="number" class="ay" min="0" max="17" step="1"><span class="unit">歳</span><input type="number" class="am" min="0" max="11" step="1"><span class="unit">ヶ月</span></span></div>';}
function addPeriod(cid,kind,fromAgeM,toAgeM,amount){
  var w=document.createElement('div');
  w.innerHTML='<div class="wrow">'+pbox('pf','開始（年月 または ○歳○ヶ月）')+pbox('pt','終了（空欄＝18歳まで）')+
    (kind==='cg'?'<div class="f"><label>月額（円・1,000円単位）<input type="text" inputmode="numeric" class="money pa"></label></div>':'')+
    '<button type="button" class="sub">削除</button></div><div class="fm wmsg"></div>';
  var b=birthIdx();
  [].forEach.call(w.querySelectorAll('.pb'),agePair);
  if(!isNaN(b)){
    w.querySelector('.pf .ym').value=ymStr(b+fromAgeM);
    if(toAgeM!=null)w.querySelector('.pt .ym').value=ymStr(b+toAgeM);
    [].forEach.call(w.querySelectorAll('.pb'),function(x){x.toAge();});
  }
  if(kind==='cg')w.querySelector('.pa').value=fmtYen(amount);
  w.querySelector('button').onclick=function(){w.remove();sched();};
  $(cid).appendChild(w);
}
$('addCg').onclick=function(){addPeriod('cgRows','cg',60,119,20000);sched();};
$('addSt').onclick=function(){addPeriod('stRows','st',120,null);sched();};
$('birth').addEventListener('input',function(){[].forEach.call(document.querySelectorAll('.pb'),function(x){x.toAge();});});


document.addEventListener('change',function(e){
  var t=e.target;
  if(t.classList&&t.classList.contains('money')){var v=pv(t);t.value=isNaN(v)?'0':fmtYen(v);}
});

/* --- 取崩し行 --- */
function addRow(t,age,amt){
  var d=document.createElement('div');d.className='wrow';
  d.innerHTML='<div class="f"><label>取崩し年齢（歳）<input type="number" class="wa" step="1"></label></div>'+
   '<div class="f"><label>取崩し金額（円）<input type="text" inputmode="numeric" class="wm money"></label></div><button type="button" class="sub">削除</button>';
  var m=document.createElement('div');m.className='fm wmsg';
  var w=document.createElement('div');w.appendChild(d);w.appendChild(m);
  d.querySelector('.wa').value=age;d.querySelector('.wm').value=fmtYen(amt);
  d.querySelector('button').onclick=function(){w.remove();sched();};
  $('wRows'+t).appendChild(w);
}
$('addW1').onclick=function(){addRow(1,12,300000);sched();};
$('addW2').onclick=function(){addRow(2,60,1000000);sched();};

/* --- 入力の読み取り・検証 --- */
function rowsOf(t){
  return [].map.call($('wRows'+t).children,function(w){
    var r=w.querySelector('.wrow');
    return {age:Math.floor(parseFloat(r.querySelector('.wa').value)),amount:pv(r.querySelector('.wm'))||0,msg:w.querySelector('.wmsg')};
  });
}
function read(){
  var b=$('birth').value.split('-').map(Number),s=$('start').value.split('-').map(Number),er=[];
  if(b.length<3||!b[0])er.push('生年月日を入力してください。');
  if(s.length<2||!s[0])er.push('開始年月を入力してください。');
  var msgs={monthly:'',lump:'',bMonthly:''};
  var monthly=Math.max(0,Math.min(50000,Math.round(pn('monthly',0)/1000)*1000));
  var lumpAmt=pn('lumpAmt',0);if(lumpAmt>600000){msgs.lump='上限60万円に補正しました。';lumpAmt=600000;}
  lumpAmt=Math.max(0,lumpAmt);
  var bm=pn('bMonthly',0);if(bm>300000){msgs.bMonthly='上限30万円に補正しました。';bm=300000;}bm=Math.max(0,bm);
  var stops=[],changes=[];
  [['stRows','st'],['cgRows','cg']].forEach(function(k){
    [].forEach.call($(k[0]).children,function(w){
      var msg=w.querySelector('.wmsg'),f=ymIdx(w.querySelector('.pf .ym').value),tv=w.querySelector('.pt .ym').value,t=tv===''?null:ymIdx(tv),m='',bad=false;
      if(isNaN(f)){m='開始年月を入力してください。';bad=true;}
      else if(t!==null&&isNaN(t)){m='終了年月が不正です。';bad=true;}
      else if(t!==null&&t<f){m='終了は開始以降にしてください。';bad=true;}
      var a=0;
      if(k[1]==='cg'){
        a=pv(w.querySelector('.pa'));if(isNaN(a))a=0;
        var a2=Math.max(0,Math.min(50000,Math.round(a/1000)*1000));
        if(!m&&a2!==a)m='月額は0〜50,000円（1,000円単位）に補正しました。';
        a=a2;
      }
      msg.className='fm wmsg'+(bad?' e':'');msg.textContent=m;
      if(!bad)(k[1]==='st'?stops:changes).push(k[1]==='st'?{from:f,to:t}:{from:f,to:t,amount:a});
    });
  });
  var items=[],valid=[];
  [1,2].forEach(function(t){
    var on=t===1||$('cOn').checked;
    rowsOf(t).forEach(function(r){
      var m='';
      if(on&&r.amount>0){
        if(isNaN(r.age))m='年齢を入力してください。';
        else if(t===1&&r.age<12)m='12歳未満は払出しできません。';
        else if(t===1&&r.age>=18)m='18歳以降の取崩しはタブ②で入力してください。';
        else if(t===2&&r.age<18)m='18歳未満の取崩しはタブ①で入力してください。';
        else if(r.age>100)m='100歳以下で入力してください。';
        else{items.push({age:r.age,amount:r.amount});valid.push(r);}
      }
      r.err=m;
      r.msg.className='fm wmsg'+(m?' e':'');r.msg.textContent=m;
    });
  });
  return {er:er,items:items,valid:valid,msgs:msgs,P:{
    by:b[0],bm:b[1],bd:b[2],sy:s[0],sm:s[1],buyDay:parseInt($('buyDay').value,10),
    rChild:pn('rChild',5)/100,rAdult:pn('rAdult',5)/100,fee:pn('fee',0.1)/100,method:$('method').value,
    stops:stops,changes:changes,
    monthly:monthly,lumpM:parseInt($('lumpM').value,10),lumpAmt:lumpAmt,
    B:{on:$('bOn').checked,startAge:Math.floor(pn('bStart',18)),endAge:Math.floor(pn('bEnd',30)),monthly:bm,
       lumpM:parseInt($('bLumpM').value,10),lumpAmt:Math.max(0,pn('bLumpAmt',0))},
    C:{on:items.length>0,items:items}
  },disp:Math.max(18,Math.min(100,Math.floor(pn('dispAge',60))))};
}

/* --- 計算・描画 --- */
var tmr;
function sched(){clearTimeout(tmr);tmr=setTimeout(calc,300);}
function syncVis(){$('bBox').hidden=!$('bOn').checked;$('cBox').hidden=!$('cOn').checked;}

function setMsg(id,arr,err){var e=$(id);e.className='fm'+(err?' e':'');e.innerHTML=arr.map(esc).join('<br>');}
function calc(){
  syncVis();
  var r=read();
  if(r.er.length){setMsg('notes1',r.er,true);return;}
  NAMES[0]='①任意金額（月'+fmtYen(r.P.monthly)+'円'+(r.P.changes.length?'＋変更あり':'')+'）';
  [].forEach.call(document.querySelectorAll('.ptab[data-p="0"]'),function(b){b.textContent=NAMES[0];});
  res=[0,1,2].map(function(p){return Engine.simulate(r.P,p);});
  last=r;
  var e0=res[0],mname=r.P.method==='simple'?'単純按分':'実効換算';
  $('rateInfo1').textContent='月利の方式：'+mname+'／月利（18歳未満）'+(e0.rmC*100).toFixed(4)+'%（信託報酬等控除後）';
  $('rateInfo2').textContent='月利の方式：'+mname+'／月利（18歳以降）'+(e0.rmA*100).toFixed(4)+'%（信託報酬等控除後）';
  $('dispLbl').textContent=r.disp;$('dispLbl2').textContent=r.disp;

  /* 行ごとの警告（残高超過） */
  r.valid.forEach(function(row){
    var lines=[];
    res.forEach(function(x,p){x.warnings.forEach(function(w){
      var mm=w.match(/^C：(\d+)歳(?:の|時点の)?(.*)$/);
      if(mm&&+mm[1]===row.age)lines.push(NAMES[p]+'：'+mm[2]);
    });});
    row.msg.className='fm wmsg';row.msg.innerHTML=lines.map(esc).join('<br>');
  });

  /* 上限による自動調整・エラーの表示 */
  var childOther=[],m1=[],l1=[],bs=[],bmn=[],bl=[],be=[];
  if(r.msgs.lump)l1.push(r.msgs.lump);
  res.forEach(function(x,p){
    x.notes.forEach(function(n){
      if(n.indexOf('こども期')===0){
        if(p===0&&n.indexOf('一括')>=0)l1.push(n);
        else if(p===0&&n.indexOf('毎月積立')>=0)m1.push(n);
        else childOther.push(NAMES[p]+'：'+n);
      }else if(n.indexOf('成人NISA')===0){if(p===0)bs.push(n);}
      else if(n.indexOf('一括')>=0)bl.push(NAMES[p]+'：'+n);
      else bmn.push(NAMES[p]+'：'+n);
    });
  });
  e0.errors.forEach(function(x){if(x.indexOf('B：')===0)be.push(x.slice(2));});
  if(r.msgs.bMonthly)bmn.unshift(r.msgs.bMonthly);
  setMsg('monthlyMsg',m1);setMsg('lumpMsg',l1);setMsg('notes1',childOther);
  setMsg('bStartMsg',bs);setMsg('bEndMsg',be,true);setMsg('bMonthlyMsg',bmn);setMsg('bLumpMsg',bl);
  $('notes2').innerHTML='';

  renderCards1();renderCards2(r.disp);renderCarry();
  renderTable(1);renderTable(2);renderLegends();renderInSum(r);drawCharts();
}

function card(p,title,rows){
  return '<div class="sc p'+(p+1)+'"><h3 style="color:'+COLS[p]+'">'+NAMES[p]+'</h3><table>'+
    rows.map(function(x){return '<tr><td>'+x[0]+'</td><td>'+x[1]+'</td></tr>';}).join('')+'</table></div>';
}
function ageTxt(a){return a?(a.years+'歳'+a.months+'ヶ月（'+a.y+'年'+a.m+'月）'):'—（到達せず）';}
function renderCards1(){
  $('sum1').innerHTML=res.map(function(x,p){var a=x.rows[18];
    return card(p,'',[['積立投資額',man(a.principal)],['運用評価額',man(a.value)],['運用益',man(a.gain)],['600万円到達',ageTxt(x.reach600)]]);}).join('');
}
function renderCards2(d){
  $('sum2').innerHTML=res.map(function(x,p){var a=x.rows[d];
    return card(p,'',[['積立投資額',man(a.principal)],['運用評価額',man(a.value)],['運用益（評価額−元本残高）',man(a.gain)],['取崩し累計',man(a.cumW)],
      ['総合損益',man(a.value+a.cumW-a.principal)],['1,800万円到達',ageTxt(x.reach1800)]]);}).join('');
}
function renderCarry(){
  $('carry').innerHTML=res.map(function(x,p){var a=x.rows[18];
    return '<div style="border-color:'+COLS[p]+'"><b style="color:'+COLS[p]+'">'+NAMES[p]+'</b><br>評価額 '+man(a.value)+'<br>元本 '+man(a.principal)+'</div>';}).join('');
}
function range(t){return t===1?[0,18]:[18,100];}
function renderTable(t){
  var rg=range(t),x=res[cur[t]];
  var h='<thead><tr><th>年齢</th><th>西暦年</th><th>当年買付額</th><th>買付元本累計</th><th>取崩し額</th><th>元本残高(簿価)</th><th>評価額</th><th>運用益</th></tr></thead><tbody>';
  for(var i=rg[0];i<=rg[1];i++){var w=x.rows[i];
    h+='<tr><td>'+w.age+'</td><td>'+w.year+'</td><td>'+fmtYen(w.buy)+'</td><td>'+fmtYen(w.principal)+'</td><td>'+fmtYen(w.withdraw)+'</td><td>'+fmtYen(w.basis)+'</td><td>'+fmtYen(w.value)+'</td><td>'+fmtYen(w.gain)+'</td></tr>';}
  $('tbl'+t).innerHTML=h+'</tbody>';
}
[].forEach.call(document.querySelectorAll('button.ptab'),function(b){
  b.onclick=function(){var t=+b.dataset.t;cur[t]=+b.dataset.p;
    [].forEach.call(document.querySelectorAll('button.ptab[data-t="'+t+'"]'),function(x){x.classList.toggle('on',x===b);});
    if(res)renderTable(t);};
});

var vis={},hov={cv:null,age:-1,p:-1},pins={1:[],2:[]},PW=190,zc=10;
function chartT(cv){return cv.id==='chart1'?1:2;}
function clamp(v,a,b){return Math.max(a,Math.min(b,v));}
function pinHTML(pn){
  var w=res[pn.p].rows[pn.age];
  return '<b style="color:'+COLS[pn.p]+'">'+esc(NAMES[pn.p])+'</b><br>'+pn.age+'歳（'+w.year+'年）<br>元本残高　'+fmtYen(w.basis)+'円<br>運用益　　'+fmtYen(w.gain)+'円<br><b>評価額　　'+fmtYen(w.value)+'円</b><br>買付元本累計　'+fmtYen(w.principal)+'円<br>取崩し累計　'+fmtYen(w.cumW)+'円';
}
var rq={};
function redraw(t){if(rq[t])return;rq[t]=requestAnimationFrame(function(){rq[t]=0;var cv=$('chart'+t);if(cv._args)drawChart.apply(null,[cv].concat(cv._args));});}
function renderPins(t){
  var cw=$('cw'+t),cv=$('chart'+t);
  [].slice.call(cw.querySelectorAll('.pin')).forEach(function(e){e.remove();});
  if(!res||!cv._args)return;
  pins[t].forEach(function(pn){
    if(pn.age<cv._args[0]||pn.age>cv._args[1])return;
    var d=document.createElement('div');d.className='pin';d.style.borderColor=COLS[pn.p];
    d.style.left=(pn.fx*100)+'%';d.style.top=(pn.fy*100)+'%';
    d.innerHTML='<button type="button" class="x" title="削除">×</button>'+pinHTML(pn);
    d.querySelector('.x').onclick=function(){pins[t].splice(pins[t].indexOf(pn),1);renderPins(t);redraw(t);};
    d.addEventListener('pointerdown',function(e){
      if(e.target.classList.contains('x'))return;
      var r=cw.getBoundingClientRect(),sx=e.clientX,sy=e.clientY,ox=pn.fx,oy=pn.fy;
      d.style.zIndex=++zc;d.setPointerCapture(e.pointerId);e.preventDefault();
      function mv(ev){
        pn.fx=clamp(ox+(ev.clientX-sx)/r.width,0,Math.max(0,1-PW/r.width));
        pn.fy=clamp(oy+(ev.clientY-sy)/r.height,0,Math.max(0,1-(pn.h||120)/r.height));
        d.style.left=(pn.fx*100)+'%';d.style.top=(pn.fy*100)+'%';redraw(t);
      }
      function up(){d.removeEventListener('pointermove',mv);d.removeEventListener('pointerup',up);}
      d.addEventListener('pointermove',mv);d.addEventListener('pointerup',up);
    });
    cw.appendChild(d);pn.h=d.offsetHeight;
  });
}
function addPin(cv,e){
  var r=cv.getBoundingClientRect(),x=e.clientX-r.left,y=e.clientY-r.top,hit=null,t=chartT(cv);
  (cv._bars||[]).forEach(function(b){if(x>=b.x0&&x<=b.x0+b.bw)hit=b;});
  if(!hit||y<hit.top-6)return;
  if(pins[t].some(function(q){return q.age===hit.age&&q.p===hit.p;}))return;
  pins[t].push({p:hit.p,age:hit.age,h:120,
    fx:clamp((x+14)/r.width,0,Math.max(0,1-PW/r.width)),fy:clamp((y-40)/r.height,0,Math.max(0,1-120/r.height))});
  $('tip').style.display='none';
  renderPins(t);redraw(t);
}
function tipShow(cv,e){
  var r=cv.getBoundingClientRect(),x=e.clientX-r.left,y=e.clientY-r.top,hit=null,tip=$('tip');
  (cv._bars||[]).forEach(function(b){if(x>=b.x0&&x<=b.x0+b.bw)hit=b;});
  if(!hit||y<hit.top-6){tipHide();return;}
  if(hov.cv!==cv||hov.age!==hit.age||hov.p!==hit.p){hov={cv:cv,age:hit.age,p:hit.p};drawChart.apply(null,[cv].concat(cv._args));}
  var w=res[hit.p].rows[hit.age],bs=Math.min(Math.max(w.basis,0),Math.max(w.value,0)),gn=Math.max(0,w.value-bs);
  var inP=hit.seg.p&&y>=hit.seg.p[0]&&y<=hit.seg.p[1],inV=hit.seg.v&&y>=hit.seg.v[0]&&y<=hit.seg.v[1];
  var mk=function(on){return on?' <span class="hl">◀</span>':'';};
  tip.innerHTML='<b style="color:'+COLS[hit.p]+';filter:brightness(1.6)">'+esc(NAMES[hit.p])+'</b>　'+hit.age+'歳（'+w.year+'年）<br>'+
    '元本残高　'+fmtYen(w.basis)+'円'+mk(inP)+'<br>運用益　　'+fmtYen(w.gain)+'円'+mk(inV)+'<br><b>評価額　　'+fmtYen(w.value)+'円</b><br>'+
    '買付元本累計　'+fmtYen(w.principal)+'円<br>取崩し累計　'+fmtYen(w.cumW)+'円';
  tip.style.display='block';
  var tw=tip.offsetWidth,th=tip.offsetHeight,lx=e.clientX+14,ly=e.clientY+14;
  if(lx+tw>innerWidth-4)lx=e.clientX-tw-14;if(ly+th>innerHeight-4)ly=e.clientY-th-14;
  tip.style.left=Math.max(4,lx)+'px';tip.style.top=Math.max(4,ly)+'px';
}
function tipHide(){
  $('tip').style.display='none';
  if(hov.cv){var c=hov.cv;hov={cv:null,age:-1,p:-1};drawChart.apply(null,[c].concat(c._args));}
}
['chart1','chart2'].forEach(function(id){var c=$(id);
  c.addEventListener('mousemove',function(e){tipShow(c,e);});c.addEventListener('click',function(e){addPin(c,e);});
  c.addEventListener('mouseleave',tipHide);});
function renderLegends(){
  var h='';
  for(var p=0;p<3;p++){
    h+='<button type="button" data-k="'+p+'v" class="'+(vis[p+'v']===false?'off':'')+'"><b style="background:'+COLS[p]+';opacity:.45"></b>'+esc(NAMES[p])+' 運用益</button>';
    h+='<button type="button" data-k="'+p+'p" class="'+(vis[p+'p']===false?'off':'')+'"><b style="background:'+COLS[p]+';width:8px"></b>'+esc(NAMES[p])+' 元本残高</button>';
  }
  h+='<span class="note">赤縦線＝取崩し／凡例クリックで表示切替／棒をクリックで内訳を固定（ドラッグで移動・×で削除・PDFにも出力）</span>';
  $('legend1').innerHTML=h;$('legend2').innerHTML=h;
}
['legend1','legend2'].forEach(function(id){$(id).addEventListener('click',function(e){
  var b=e.target.closest('button');if(!b)return;var k=b.dataset.k;vis[k]=vis[k]===false;renderLegends();drawCharts();});});
function drawChart(cv,from,to,wl){
  var dpr=cv._dpr||window.devicePixelRatio||1,W=cv._forceW||cv.clientWidth,H=cv._forceH||cv.clientHeight;
  if(!W||!H||!res)return;
  if(!cv._forceW)cv._args=[from,to,wl];cv._bars=[];
  cv.width=W*dpr;cv.height=H*dpr;
  var g=cv.getContext('2d');g.setTransform(dpr,0,0,dpr,0,0);g.clearRect(0,0,W,H);
  var L=64,R=10,T=14,B=28,pw=W-L-R,ph=H-T-B,max=1,i,p,n=to-from+1;
  for(p=0;p<3;p++)for(i=from;i<=to;i++){var w=res[p].rows[i];
    var bs=Math.min(Math.max(w.basis,0),Math.max(w.value,0)),gn=Math.max(0,w.value-bs);
    max=Math.max(max,(vis[p+'p']!==false?bs:0)+(vis[p+'v']!==false?gn:0));}
  var step=Math.pow(10,Math.floor(Math.log10(max/4))),f=[1,2,5,10],s=step;
  for(i=0;i<4;i++){if(max/(f[i]*step)<=6){s=f[i]*step;break;}}
  var top=Math.ceil(max/s)*s,sw=pw/n;
  var CX=function(a){return L+(a-from+0.5)*sw;},Y=function(v){return T+ph-(Math.max(0,v)/top)*ph;};
  g.font='11px sans-serif';g.textAlign='right';g.textBaseline='middle';g.lineWidth=1;
  for(var v=0;v<=top+1e-6;v+=s){g.strokeStyle='#e5e7eb';g.beginPath();g.moveTo(L,Y(v));g.lineTo(W-R,Y(v));g.stroke();g.fillStyle='#6b7280';g.fillText((v/10000).toLocaleString('ja-JP')+'万',L-4,Y(v));}
  g.textAlign='center';g.textBaseline='top';
  var xs=n>60?10:(n>30?5:(n>12?2:1));
  for(i=from;i<=to;i+=xs){g.fillStyle='#6b7280';g.fillText(i+'歳',CX(i),T+ph+6);}
  var gw=sw*0.86,slot=gw/3;
  for(i=from;i<=to;i++)for(p=0;p<3;p++){
    var w2=res[p].rows[i],x0=CX(i)-gw/2+p*slot,bw=slot*0.94;
    var b2=Math.min(Math.max(w2.basis,0),Math.max(w2.value,0)),g2=Math.max(0,w2.value-b2),acc=0;
    var bar={x0:x0,bw:bw,age:i,p:p,seg:{}};
    if(vis[p+'p']!==false){g.globalAlpha=1;g.fillStyle=COLS[p];g.fillRect(x0,Y(b2),bw,T+ph-Y(b2));bar.seg.p=[Y(b2),T+ph];acc=b2;}
    if(vis[p+'v']!==false){g.globalAlpha=.45;g.fillStyle=COLS[p];g.fillRect(x0,Y(acc+g2),bw,Y(acc)-Y(acc+g2));bar.seg.v=[Y(acc+g2),Y(acc)];}
    if(bar.seg.p||bar.seg.v){bar.top=Math.min(bar.seg.p?bar.seg.p[0]:1e9,bar.seg.v?bar.seg.v[0]:1e9);cv._bars.push(bar);}
    if(hov.cv===cv&&hov.age===i&&hov.p===p&&bar.top!=null){g.globalAlpha=1;g.strokeStyle='#111';g.lineWidth=1.5;g.strokeRect(x0,bar.top,bw,T+ph-bar.top);}
  }
  g.globalAlpha=1;
  var pt=chartT(cv);
  (cv._pins||pins[pt]).forEach(function(pn){
    var bar=null;cv._bars.forEach(function(b){if(b.age===pn.age&&b.p===pn.p)bar=b;});
    if(!bar)return;
    g.strokeStyle='#111';g.lineWidth=1.5;g.strokeRect(bar.x0,bar.top,bar.bw,T+ph-bar.top);
    var pl=pn.fx*W,ptp=pn.fy*H,ph2=pn.h||120,bx=bar.x0+bar.bw/2,by=bar.top;
    var qx=Math.max(pl,Math.min(bx,pl+PW)),qy=Math.max(ptp,Math.min(by,ptp+ph2));
    if(qx!==bx||qy!==by){g.strokeStyle=COLS[pn.p];g.lineWidth=1;g.setLineDash([3,2]);g.beginPath();g.moveTo(bx,by);g.lineTo(qx,qy);g.stroke();g.setLineDash([]);}
  });
  var seen={};
  wl.forEach(function(a){
    if(a<from||a>to||seen[a])return;seen[a]=1;
    g.strokeStyle='#dc2626';g.setLineDash([4,3]);g.beginPath();g.moveTo(CX(a),T);g.lineTo(CX(a),T+ph);g.stroke();g.setLineDash([]);
    g.fillStyle='#dc2626';g.textAlign='left';g.textBaseline='top';g.fillText('取崩'+a,CX(a)+3,T+2);
  });
}
var inRows=[];
function renderInSum(r){
  var P=r.P,rows=[],ym=function(n){return Math.floor(n/12)+'年'+(n%12+1)+'月';},
      per=function(c){return ym(c.from)+'〜'+(c.to==null?'18歳まで':ym(c.to));};
  rows.push(['お子様の生年月日',P.by+'年'+P.bm+'月'+P.bd+'日']);
  rows.push(['シミュレーション開始',P.sy+'年'+P.sm+'月／買付日：毎月'+P.buyDay+'日']);
  rows.push(['①毎月の積立額',yen(P.monthly)]);
  if(P.changes.length)rows.push(['①積立額の変更',P.changes.map(function(c){return per(c)+'：'+yen(c.amount);}).join('<br>')]);
  if(P.lumpAmt>0)rows.push(['①年1回の一括追加',P.lumpM+'月に'+yen(P.lumpAmt)]);
  if(P.stops.length)rows.push(['積立の停止（3パターン共通）',P.stops.map(per).join('<br>')]);
  rows.push(['想定リターン（年率）','18歳未満 '+(P.rChild*100)+'%／18歳以降 '+(P.rAdult*100)+'%']);
  rows.push(['信託報酬等（年率）',(P.fee*100)+'%（'+(P.method==='simple'?'単純按分':'実効換算')+'）']);
  if(P.B.on)rows.push(['18歳以降の追加投資',P.B.startAge+'歳〜'+P.B.endAge+'歳（誕生日の前日まで）／毎月'+yen(P.B.monthly)+(P.B.lumpAmt>0?'／'+P.B.lumpM+'月に'+yen(P.B.lumpAmt):'')]);
  if(r.items.length)rows.push(['取崩し',r.items.map(function(x){return x.age+'歳：'+yen(x.amount);}).join('<br>')]);
  rows.push(['確認年齢（タブ②）',r.disp+'歳']);
  inRows=rows;
  $('inSum').innerHTML='<table>'+rows.map(function(x){return '<tr><td>'+x[0]+'</td><td>'+x[1]+'</td></tr>';}).join('')+'</table>';
}
function drawAll(){
  if(!res||!last)return;
  var w=last.items.map(function(x){return x.age;});
  drawChart($('chart1'),0,18,w);drawChart($('chart2'),18,last.disp,w);renderPins(1);renderPins(2);
}
function drawCharts(){
  if(!res||!last)return;
  var w=last.items.map(function(x){return x.age;});
  if(tab===1)drawChart($('chart1'),0,18,w);else drawChart($('chart2'),18,last.disp,w);
  renderPins(tab);
}

/* --- CSV --- */
function csv(t){
  var rg=range(t),q=function(x){return '"'+String(x).replace(/"/g,'""')+'"';};
  var lines=[['パターン','年齢','西暦年','当年買付額','買付元本累計','取崩し額','元本残高(簿価)','評価額','運用益'].join(',')];
  res.forEach(function(x,p){for(var i=rg[0];i<=rg[1];i++){var w=x.rows[i];
    lines.push([q(NAMES[p]),w.age,w.year,Math.round(w.buy),Math.round(w.principal),Math.round(w.withdraw),Math.round(w.basis),Math.round(w.value),Math.round(w.gain)].join(','));}});
  var blob=new Blob(['﻿'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'});
  var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='children_nisa_'+(t===1?'age0-18':'age18-100')+'.csv';
  document.body.appendChild(a);a.click();a.remove();setTimeout(function(){URL.revokeObjectURL(a.href);},1000);
}
[].forEach.call(document.querySelectorAll('button.csvb'),function(b){b.onclick=function(){if(res)csv(+b.dataset.t);};});

/* ===== PDF直接生成（外部ライブラリなし・A4縦・JPEG貼り付け） ===== */
function buildPdf(keys){
  var S=2,PWd=794,PHd=1123,M=42,CWd=PWd-2*M,BOT=PHd-M-16;
  var FONT='"Hiragino Sans","Yu Gothic UI","Meiryo",sans-serif';
  var pages=[],g,y;
  function newPage(){var c=document.createElement('canvas');c.width=PWd*S;c.height=PHd*S;g=c.getContext('2d');
    g.fillStyle='#fff';g.fillRect(0,0,c.width,c.height);g.scale(S,S);g.textBaseline='top';pages.push(c);y=M;}
  function ensure(h){if(y+h>BOT)newPage();}
  function font(sz,b){g.font=(b?'bold ':'')+sz+'px '+FONT;}
  function wrap(t,w){var out=[];String(t).split('\n').forEach(function(par){var line='';
    for(var i=0;i<par.length;i++){var c=par[i];if(g.measureText(line+c).width>w&&line){out.push(line);line=c;}else line+=c;}out.push(line);});return out;}
  function title(t){font(14,true);g.fillStyle='#2563eb';g.fillRect(M,y+1,4,17);g.fillStyle='#1f2937';g.fillText(t,M+10,y+1);y+=28;}

  newPage();
  font(20,true);g.fillStyle='#1f2937';g.fillText('こどもNISA 積立シミュレーション結果',M,y);
  var d=new Date();font(10);g.fillStyle='#6b7280';g.textAlign='right';
  g.fillText('作成日：'+d.getFullYear()+'年'+(d.getMonth()+1)+'月'+d.getDate()+'日',PWd-M,y+8);g.textAlign='left';
  y+=34;g.strokeStyle='#d8dee9';g.beginPath();g.moveTo(M,y-6);g.lineTo(PWd-M,y-6);g.stroke();

  function secIn(){
    font(11.5);var lw=190,vw=CWd-lw-10,items=inRows.map(function(r){
      var ls=wrap(String(r[1]).replace(/<br>/g,'\n'),vw);return {l:r[0],ls:ls,h:ls.length*16+6};});
    var h=28+items.reduce(function(a,b){return a+b.h;},0);ensure(h);title('入力条件');
    items.forEach(function(it){font(11.5);g.fillStyle='#6b7280';g.fillText(it.l,M+4,y+2);g.fillStyle='#1f2937';
      it.ls.forEach(function(t,i){g.fillText(t,M+lw,y+2+i*16);});
      g.strokeStyle='#e5e7eb';g.beginPath();g.moveTo(M,y+it.h);g.lineTo(PWd-M,y+it.h);g.stroke();y+=it.h;});
    y+=10;
  }
  function secCards(ttl,cards){
    var gap=10,cw=(CWd-2*gap)/3,fs=11.5;font(fs);
    cards.forEach(function(c){c.rr=c.rows.map(function(r){var two=g.measureText(r[0]).width+g.measureText(r[1]).width+12>cw-16;return {two:two,h:two?34:20};});
      c.h=32+c.rr.reduce(function(a,b){return a+b.h;},0)+8;});
    var H=Math.max.apply(null,cards.map(function(c){return c.h;}));ensure(28+H+10);title(ttl);
    cards.forEach(function(c,i){var x=M+i*(cw+gap);g.strokeStyle=c.color;g.lineWidth=2;g.strokeRect(x,y,cw,H);g.lineWidth=1;
      font(12,true);g.fillStyle=c.color;
      var nl=wrap(c.name,cw-16);g.fillText(nl[0]+(nl.length>1?'…':''),x+8,y+7);var yy=y+30;font(fs);
      c.rows.forEach(function(r,k){var rr=c.rr[k];g.fillStyle='#4b5563';g.textAlign='left';g.fillText(r[0],x+8,yy+2);
        g.fillStyle='#111';g.textAlign='right';g.fillText(r[1],x+cw-8,yy+2+(rr.two?16:0));g.textAlign='left';
        g.strokeStyle='#e5e7eb';g.beginPath();g.setLineDash([1,2]);g.moveTo(x+8,yy+rr.h-1);g.lineTo(x+cw-8,yy+rr.h-1);g.stroke();g.setLineDash([]);yy+=rr.h;});});
    y+=H+12;
  }
  function secChart(t,ttl){
    var from=t===1?0:18,to=t===1?18:last.disp,W=CWd,H=300;
    font(11);var items=[];
    for(var p=0;p<3;p++){if(vis[p+'p']!==false)items.push([COLS[p],NAMES[p]+' 元本残高',1]);if(vis[p+'v']!==false)items.push([COLS[p],NAMES[p]+' 運用益',.45]);}
    var rows=[[]],rw=0;items.forEach(function(it){var w=g.measureText(it[1]).width+26;if(rw+w>CWd&&rows[rows.length-1].length){rows.push([]);rw=0;}rows[rows.length-1].push(it);rw+=w;});
    var lh=rows.length*18+4;ensure(28+H+lh+10);title(ttl);
    var tmp=document.createElement('canvas');tmp.id='chart'+t;tmp._forceW=W;tmp._forceH=H;tmp._dpr=S;
    var w0=last.items.map(function(x){return x.age;});
    var adj=pins[t].filter(function(q){return q.age>=from&&q.age<=to;}).map(function(q){
      return {p:q.p,age:q.age,h:122,fx:clamp(q.fx,0,Math.max(0,1-PW/W)),fy:clamp(q.fy,0,Math.max(0,1-122/H))};});
    tmp._pins=adj;drawChart(tmp,from,to,w0);
    g.drawImage(tmp,M,y,W,H);
    adj.forEach(function(q){var x=M+q.fx*W,yy=y+q.fy*H,w=res[q.p].rows[q.age];
      g.fillStyle='#fff';g.fillRect(x,yy,PW,122);g.strokeStyle=COLS[q.p];g.lineWidth=2;g.strokeRect(x,yy,PW,122);g.lineWidth=1;
      var ls=[[NAMES[q.p],1],[q.age+'歳（'+w.year+'年）'],['元本残高　'+fmtYen(w.basis)+'円'],['運用益　　'+fmtYen(w.gain)+'円'],['評価額　　'+fmtYen(w.value)+'円',2],['買付元本累計　'+fmtYen(w.principal)+'円'],['取崩し累計　'+fmtYen(w.cumW)+'円']];
      ls.forEach(function(l,i){font(11,l[1]===1||l[1]===2);g.fillStyle=l[1]===1?COLS[q.p]:'#111';var tx=l[0];
        if(l[1]===1){while(g.measureText(tx).width>PW-14&&tx.length>3)tx=tx.slice(0,-1);}g.fillText(tx,x+7,yy+6+i*16);});});
    y+=H+4;font(11);
    rows.forEach(function(r){var x=M;r.forEach(function(it){g.globalAlpha=it[2];g.fillStyle=it[0];g.fillRect(x,y+4,it[2]===1?8:14,9);g.globalAlpha=1;
      g.fillStyle='#374151';g.fillText(it[1],x+18,y+2);x+=g.measureText(it[1]).width+26;});y+=18;});
    y+=10;
  }
  function secTable(t,ttl){
    var rg=range(t),x=res[cur[t]],ws=[44,52,88,104,88,104,120,110],hd=['年齢','西暦年','当年買付額','買付元本累計','取崩し額','元本残高(簿価)','評価額','運用益'],RH=15.5;
    function head(cont){ensure(28+18+RH*3);title(ttl+(cont?'（続き）':''));var xx=M;font(10,true);
      g.fillStyle='#eef2f9';g.fillRect(M,y,CWd,18);g.fillStyle='#1f2937';
      hd.forEach(function(h,i){g.textAlign=i<2?'center':'right';g.fillText(h,i<2?xx+ws[i]/2:xx+ws[i]-5,y+3);xx+=ws[i];});g.textAlign='left';y+=18;}
    head(false);
    for(var i=rg[0];i<=rg[1];i++){
      if(y+RH>BOT){newPage();head(true);}
      var w=x.rows[i],vals=[w.age,w.year,w.buy,w.principal,w.withdraw,w.basis,w.value,w.gain],xx=M;font(10);g.fillStyle='#111';
      if((i-rg[0])%2)  {g.fillStyle='#f8fafc';g.fillRect(M,y,CWd,RH);g.fillStyle='#111';}
      vals.forEach(function(v,k){var tx=k<2?String(v):fmtYen(v);g.textAlign=k<2?'center':'right';g.fillText(tx,k<2?xx+ws[k]/2:xx+ws[k]-5,y+2.5);xx+=ws[k];});g.textAlign='left';
      g.strokeStyle='#e5e7eb';g.beginPath();g.moveTo(M,y+RH);g.lineTo(PWd-M,y+RH);g.stroke();y+=RH;
    }
    y+=10;
  }
  function secCarry(){
    var cards=res.map(function(x,p){var a=x.rows[18];return {name:NAMES[p],color:COLS[p],rows:[['評価額',man(a.value)],['元本',man(a.principal)]]};});
    secCards('18歳時点の評価額（タブ①の結果）',cards);
  }
  function secAssump(){
    var lis=[].map.call(document.querySelectorAll('[data-pdf="as"] li'),function(l){return l.textContent;});
    var disc=document.querySelector('[data-pdf="as"] p').textContent;
    font(10.5);ensure(28+40);title('前提条件・注記');
    lis.forEach(function(t){font(10.5);var ls=wrap(t,CWd-14);if(y+ls.length*15+3>BOT)newPage();g.fillStyle='#374151';g.fillText('・',M,y);
      ls.forEach(function(l,i){g.fillText(l,M+12,y+i*15);});y+=ls.length*15+3;});
    y+=6;font(11,true);var dl=wrap(disc,CWd-20);var h=dl.length*16+14;ensure(h);
    g.strokeStyle='#b45309';g.lineWidth=2;g.strokeRect(M,y,CWd,h);g.lineWidth=1;g.fillStyle='#78350f';
    dl.forEach(function(l,i){g.fillText(l,M+10,y+7+i*16);});y+=h+8;
  }
  var disp=last.disp;
  var map={
    'in':secIn,
    'c1':function(){secCards('結果：18歳時点',res.map(function(x,p){var a=x.rows[18];return {name:NAMES[p],color:COLS[p],rows:[['積立投資額',man(a.principal)],['運用評価額',man(a.value)],['運用益',man(a.gain)],['600万円到達',ageTxt(x.reach600)]]};}));},
    'g1':function(){secChart(1,'評価額の推移（0〜18歳・積み上げ：元本＋運用益）');},
    't1':function(){secTable(1,'年齢別推移表（0〜18歳）'+' '+NAMES[cur[1]]);},
    'cy':secCarry,
    'c2':function(){secCards('結果：'+disp+'歳時点',res.map(function(x,p){var a=x.rows[disp];return {name:NAMES[p],color:COLS[p],rows:[['積立投資額',man(a.principal)],['運用評価額',man(a.value)],['運用益（評価額−元本残高）',man(a.gain)],['取崩し累計',man(a.cumW)],['総合損益',man(a.value+a.cumW-a.principal)],['1,800万円到達',ageTxt(x.reach1800)]]};}));},
    'g2':function(){secChart(2,'評価額の推移（18歳〜'+disp+'歳・積み上げ：元本＋運用益）');},
    't2':function(){secTable(2,'年齢別推移表（18〜100歳）'+' '+NAMES[cur[2]]);},
    'as':secAssump
  };
  keys.forEach(function(k){map[k]();});
  pages.forEach(function(c,i){var gg=c.getContext('2d');gg.setTransform(S,0,0,S,0,0);gg.textBaseline='top';
    gg.strokeStyle='#999';gg.lineWidth=1;gg.beginPath();gg.moveTo(M,PHd-M);gg.lineTo(PWd-M,PHd-M);gg.stroke();
    gg.font='9px '+FONT;gg.fillStyle='#555';gg.textAlign='left';
    gg.fillText('本試算は将来の運用成果を示唆・保証するものではありません。制度は2027年開始予定であり、詳細は今後変更される可能性があります。',M,PHd-M+4);
    gg.textAlign='right';gg.fillText((i+1)+' / '+pages.length,PWd-M,PHd-M+4);});
  return pagesToPdf(pages);
}
function pagesToPdf(pages){
  var chunks=[],pos=0,offs=[],enc=new TextEncoder();
  function put(x){var u=typeof x==='string'?enc.encode(x):x;chunks.push(u);pos+=u.length;}
  var n=pages.length,kids=[];for(var i=0;i<n;i++)kids.push((3+3*i)+' 0 R');
  put('%PDF-1.4\n');
  offs[1]=pos;put('1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n');
  offs[2]=pos;put('2 0 obj\n<< /Type /Pages /Kids ['+kids.join(' ')+'] /Count '+n+' >>\nendobj\n');
  pages.forEach(function(c,i){
    var b=atob(c.toDataURL('image/jpeg',0.92).split(',')[1]),u=new Uint8Array(b.length);for(var k=0;k<b.length;k++)u[k]=b.charCodeAt(k);
    var pg=3+3*i,ct=pg+1,im=pg+2,cs='q 595.28 0 0 841.89 0 0 cm /Im'+i+' Do Q';
    offs[pg]=pos;put(pg+' 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.28 841.89] /Resources << /XObject << /Im'+i+' '+im+' 0 R >> >> /Contents '+ct+' 0 R >>\nendobj\n');
    offs[ct]=pos;put(ct+' 0 obj\n<< /Length '+cs.length+' >>\nstream\n'+cs+'\nendstream\nendobj\n');
    offs[im]=pos;put(im+' 0 obj\n<< /Type /XObject /Subtype /Image /Width '+c.width+' /Height '+c.height+' /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length '+u.length+' >>\nstream\n');
    put(u);put('\nendstream\nendobj\n');
  });
  var total=3+3*n,xr=pos;put('xref\n0 '+total+'\n0000000000 65535 f \n');
  for(var j=1;j<total;j++)put(('0000000000'+offs[j]).slice(-10)+' 00000 n \n');
  put('trailer\n<< /Size '+total+' /Root 1 0 R >>\nstartxref\n'+xr+'\n%%EOF\n');
  return new Blob(chunks,{type:'application/pdf'});
}
$('pdfGo').onclick=function(){
  if(!res||!last)return;
  var keys=['in','c1','g1','t1','cy','c2','g2','t2','as'].filter(function(k){return $('pk_'+k).checked;});
  if(!keys.length){$('pdfMsg').className='fm e';$('pdfMsg').textContent='出力する項目を1つ以上選択してください。';return;}
  $('pdfMsg').className='fm note';$('pdfMsg').textContent='PDFを作成中…';
  setTimeout(function(){
    try{
      var blob=buildPdf(keys),a=document.createElement('a');
      a.href=URL.createObjectURL(blob);a.download='children_nisa_simulation.pdf';document.body.appendChild(a);a.click();a.remove();
      setTimeout(function(){URL.revokeObjectURL(a.href);},2000);
      $('pdfMsg').textContent='PDFを作成しました（'+Math.round(blob.size/1024)+'KB）。';
    }catch(err){$('pdfMsg').className='fm e';$('pdfMsg').textContent='PDFの作成に失敗しました：'+err.message;}
  },30);
};

/* --- タブ・印刷 --- */
function setTab(n){
  tab=n;
  $('tp1').classList.toggle('on',n===1);$('tp2').classList.toggle('on',n===2);
  $('tb1').classList.toggle('on',n===1);$('tb2').classList.toggle('on',n===2);
  window.scrollTo(0,0);drawCharts();
}
$('tb1').onclick=function(){setTab(1);};$('tb2').onclick=function(){setTab(2);};
$('goNext').onclick=function(){setTab(2);};
function syncStick(){
  var h=document.querySelector('.topbar').offsetHeight;
  document.querySelector('.cta').style.top=h+'px';
  var h2=h+document.querySelector('.cta').offsetHeight;
  document.querySelector('.tabs').style.top=h2+'px';
}
syncStick();window.addEventListener('resize',syncStick);
$('menuBtn').onclick=function(e){e.stopPropagation();$('menuList').hidden=!$('menuList').hidden;};
document.addEventListener('click',function(e){if(!e.target.closest('.menuwrap'))$('menuList').hidden=true;});
document.addEventListener('keydown',function(e){if(e.key==='Escape')$('menuList').hidden=true;});
$('pdfPrint').onclick=function(){
  if(!res)return;
  var any=false;
  [].forEach.call(document.querySelectorAll('[data-pdf]'),function(el){var on=$('pk_'+el.dataset.pdf).checked;el.classList.toggle('pdf-hide',!on);if(on)any=true;});
  if(!any){$('pdfMsg').textContent='出力する項目を1つ以上選択してください。';return;}
  $('pdfMsg').textContent='';
  var d=new Date();
  $('pdfHead').textContent='作成日：'+d.getFullYear()+'年'+(d.getMonth()+1)+'月'+d.getDate()+'日';
  document.body.classList.add('pdfmode','pdfshow');
  drawAll();
  setTimeout(function(){window.print();},80);
};
var opened=[];
window.addEventListener('beforeprint',function(){
  opened=[];[].forEach.call(document.querySelectorAll(document.body.classList.contains('pdfmode')?'.tp details':'.tp.on details'),function(d){if(!d.open){d.open=true;opened.push(d);}});
  if(document.body.classList.contains('pdfmode'))drawAll();else drawCharts();
});
window.addEventListener('afterprint',function(){
  opened.forEach(function(d){d.open=false;});opened=[];
  document.body.classList.remove('pdfmode','pdfshow');
  [].forEach.call(document.querySelectorAll('.pdf-hide'),function(el){el.classList.remove('pdf-hide');});
  drawCharts();
});
window.addEventListener('resize',drawCharts);
document.addEventListener('input',sched);
document.addEventListener('change',function(){syncVis();sched();});
calc();
})();}
