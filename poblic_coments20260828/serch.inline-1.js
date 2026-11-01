const DATA=JSON.parse(document.getElementById('embedded-data').textContent);
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const KEY='fsaNavV9_';
const key=r=>r.source+'-'+r.no;

const SOURCES={
 '0330f':{label:'3/30 内閣府令等',count:203,url:'https://www.fsa.go.jp/news/r7/hoken/20260330/20260330.html',color:'#f0a64a'},
 '0330':{label:'3/30 監督指針',count:152,url:'https://www.fsa.go.jp/news/r7/hoken/20260330-2/20260330-2.html',color:'#8a67d3'},
 '0717c':{label:'7/17 監督指針',count:76,url:'https://www.fsa.go.jp/news/r8/sonota/20260717-2/20260717.html',color:'#42a77c'},
 '0828':{label:'8/28 比較推奨',count:759,url:'https://www.fsa.go.jp/news/r8/hoken/20260828/20260828.html',color:'#3b91c9'}
};
const SOURCE_ORDER=['0330f','0330','0717c','0828'];

const PROFILES={
 INS_LARGE:{group:'insurer',label:'大手損保（既存シェア防衛型）',desc:'大手損害保険会社（営業推進・チャネル企画・経営陣）',
  must:['大手損保','大手損害保険','引受シェア','シェア調整','既存シェア','保険契約数又は保険引受シェア'],
  related:['損害保険会社','販売目標','手数料ポイント','営業上の配慮','既存契約','取扱量','委託解除']},
 INS_FAIR:{group:'insurer',label:'中堅・外資・生保（公正競争型）',desc:'中堅損保、外資系損保、生命保険会社各社',
  must:['中堅','外資','生命保険会社','生命保険募集人','生命保険','公正な競争','健全な競争','イコールフッティング'],
  related:['競争環境','公平性','公正な保険募集','乗合生命','少額短期保険','保険仲立人']},
 INS_DIRECT:{group:'insurer',label:'ダイレクト系損保（通販・価格優位型）',desc:'通販型・ネットダイレクト損害保険会社',
  must:['ダイレクト型','通販型','ネット専業','オンラインによるダイレクト販売','インターネットを通じた募集'],
  related:['オンライン募集','オンライン手続','WEBサイト','ウェブサイト','インターネット取引','メールやSMS','低廉な保険料']},
 INS_HQ_FIELD:{group:'insurer',label:'保険会社 本部管理・現場ソリシタ',desc:'営業推進（1線）、コンプラ・監査（2線）、現場ソリシタ',
  must:['管理責任者','ソリシタ','本部','内部監査','コンプライアンス部門','営業部門','保険会社の営業部門'],
  related:['代理店管理','教育・管理・指導','管理態勢','体制整備','監査','営業担当','指導等','所属保険会社']},
 AGENCY_MOTOR:{group:'agency',label:'モーター・自動車兼業系代理店',desc:'正規ディーラー、中古車販売店、整備・板金工場、商用車ディーラー',
  must:['自動車ディーラー','ディーラー代理店','自動車販売','自動車修理業','自動車の修理業務','修理工場','整備工場'],
  related:['車検','修理','事故修理','自動車保険','ブランド保険','納車','レッカー','レンタカー','テレマティクス']},
 AGENCY_RETAIL:{group:'agency',label:'来店型ショップ・大型リテール乗合代理店',desc:'モール内保険ショップ、大型訪問販売型乗合代理店（多社取扱型）',
  must:['来店型','保険ショップ','大型リテール','特定大規模乗合保険募集人','特定大規模乗合損害保険代理店','特定大規模乗合生命保険募集人','大規模乗合'],
  related:['乗合代理店','取扱保険会社','複数の保険会社','比較推奨販売','店舗','販売チャネル']},
 AGENCY_INDEPENDENT:{group:'agency',label:'独立系・地域密着プロ代理店',desc:'損保プロ代理店、事業統合・M&A代理店、専業代理店',
  must:['独立系','地域密着','プロ代理店','中小代理店','小規模代理店'],
  related:['代理店の規模','地域','募集人個人','保険募集人','専門的知識','担当者との信頼関係']},
 AGENCY_CORP:{group:'agency',label:'企業内・職域代理店',desc:'ハウスエージェンシー、職域・共済含特化代理店、企業グループ保険子会社',
  must:['企業内代理店','子会社企業代理店','企業代理店','職域代理店','親法人','子会社であるため','密接な関係を有する者'],
  related:['団体従業員','従業員','社員本人契約','福利厚生','団体扱い','団体保険','職域','グループ会社']},
 AGENCY_BANK:{group:'agency',label:'金融機関代理店',desc:'都市銀行、地方銀行、信用金庫、証券会社等の保険窓販',
  must:['銀行代理店','銀行窓販','金融機関代理店','銀行の別働隊','信用金庫','信用組合'],
  related:['銀行','金融機関','預金協力','融資','窓口販売','金融サービス仲介']},
 AGENCY_DIGITAL_LEASE:{group:'agency',label:'デジタル・WEB・リース系代理店',desc:'ネット完結型代理店、比較サイト、リース受託型代理店',
  must:['デジタル代理店','WEB代理店','ウェブ代理店','リース会社','リース系','オンライン代理店'],
  related:['WEB','ウェブ','オンライン','インターネット','デジタル','非対面','システム','比較ツール','プログラム','メール','SMS']},
 AGENCY_ALLIANCE:{group:'agency',label:'共同募集・アライアンス型代理店',desc:'専属×乗合共同募集、異業種提携・紹介スキーム、事務分担型',
  must:['共同募集','アライアンス','共同保険','再委託','再受託者'],
  related:['提携','協業','共同','委託者','再委託者','共同して','紹介']}
};

const OFFICIAL_THEME_ORDER=['顧客の最善の利益','意向がない場合','指名買い','意向推定','比較ツール','比較可能な同種の保険契約','推奨商品群','合理的かつ一定の具体性を有する基準や理由','情報提供','共同募集','共同保険','更改・継続契約','団体・団体扱い','別個登録','自賠責保険','体制整備','証跡保存','施行時期','ご意見','その他'];

const EXTERNAL_REFS=[
 {url:'https://www.fsa.go.jp/news/26/hoken/20150527-1.html',phrases:['「平成26年改正保険業法(2年以内施行)に係る政府令・監督指針案」に対するパブリックコメント結果等について','「平成26年改正保険業法（2年以内施行）に係る政府令・監督指針案」に対するパブリックコメント結果等について','平成27年5月27日パブリックコメント']},
 {url:'https://www.fsa.go.jp/news/r5/sonota/20240327/20240327.html',phrases:['「令和5年金融商品取引法等改正に係る政令・内閣府令案等に関するパブリックコメントの結果等について」','「令和５年金融商品取引法等改正に係る政令・内閣府令案等に関するパブリックコメントの結果等について」']},
 {url:'https://www.fsa.go.jp/singi/singi_kinyu/tosin/20241225.html',phrases:['金融審議会「損害保険業等に関する制度等ワーキング・グループ」報告書','「損害保険業等に関する制度等ワーキング・グループ」報告書']},
 {url:'https://www.fsa.go.jp/singi/sonpo/',phrases:['「損害保険業の構造的課題と競争のあり方に関する有識者会議」','損害保険業の構造的課題と競争のあり方に関する有識者会議']},
 {url:'https://www.fsa.go.jp/common/law/guide/kaisya/14.pdf',phrases:['事務ガイドライン（第三分冊：金融会社関係 １４ 資金移動業者関係）','事務ガイドライン（第三分冊：金融会社関係 14 資金移動業者関係）']},
 {url:'https://www.fsa.go.jp/common/noact/kaitou_2/hoken/index.html',phrases:['「広く共有することが有効な相談事例(保険業法関係)」','「広く共有することが有効な相談事例（保険業法関係）」']}
];

let profileCode=null;
let group=null;
let displayMode='all';
let activeSources=new Set(SOURCE_ORDER);
let page=1;
const pageSize=50;
let favs=new Set(JSON.parse(localStorage.getItem(KEY+'favs')||'[]'));
let notes=JSON.parse(localStorage.getItem(KEY+'notes')||'{}');
let selected=new Set();

function textOf(r){return [r.theme,r.section,r.location,r.comment,r.answer,(r.tags||[]).join(' ')].join(' ')}
function relation(r){
 if(!profileCode)return 'raw';
 const p=PROFILES[profileCode],t=textOf(r);
 if(p.must.some(k=>t.includes(k)))return 'must';
 if(p.related.some(k=>t.includes(k)))return 'related';
 return 'none';
}
function addRefs(raw){
 let x=esc(raw||'');
 const refs=[];
 EXTERNAL_REFS.forEach(o=>o.phrases.forEach(p=>refs.push({p,url:o.url})));
 refs.sort((a,b)=>b.p.length-a.p.length);
 refs.forEach(r=>{const needle=esc(r.p);if(x.includes(needle))x=x.split(needle).join(`<a class="extRef" href="${r.url}" target="_blank" rel="noopener">${needle} ↗</a>`)});
 return x;
}
function hi(raw){
 let x=addRefs(raw);
 const qs=$('q').value.trim().split(/\s+/).filter(Boolean).sort((a,b)=>b.length-a.length);
 if(!qs.length)return x;
 return x.split(/(<[^>]+>)/g).map(part=>{
  if(part.startsWith('<'))return part;
  qs.forEach(q=>{const re=new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'gi');part=part.replace(re,m=>`<mark>${m}</mark>`)});
  return part;
 }).join('');
}
function brief(s,n=160){s=String(s||'').replace(/\s+/g,' ').trim();return s.length>n?s.slice(0,n)+'…':s}
function sourceOrder(s){return SOURCE_ORDER.indexOf(s)}

function showStep1(){
 $('chooseStep1').classList.remove('hidden');$('chooseStep2').classList.add('hidden');$('chooseStep3').classList.add('hidden');
}
function showStep2(g){
 group=g;$('chooseStep1').classList.add('hidden');$('chooseStep2').classList.remove('hidden');$('chooseStep3').classList.add('hidden');
 const list=Object.entries(PROFILES).filter(([k,p])=>p.group===g);
 $('profileGrid').innerHTML=list.map(([k,p])=>`<button class="profileCard" data-profile="${k}"><b>${esc(p.label)}</b><span>${esc(p.desc)}</span><em>この属性で見る →</em></button>`).join('');
 document.querySelectorAll('[data-profile]').forEach(b=>b.onclick=()=>selectProfile(b.dataset.profile));
}
function selectProfile(code){
 profileCode=code||null;
 displayMode=code?'all':'raw';
 page=1;
 showSourceStep();
}

function showSourceStep(){
 $('chooseStep1').classList.add('hidden');$('chooseStep2').classList.add('hidden');$('chooseStep3').classList.remove('hidden');
 renderSourceStep();
}
function renderSourceStep(){
 $('sourceStepGrid').innerHTML=SOURCE_ORDER.map(src=>{
  const o=SOURCES[src],on=activeSources.has(src);
  return `<button class="sourceStepCard ${on?'active':''}" data-step-src="${src}"><div><b>${esc(o.label)}</b><span>${o.count}件</span></div><em>${on?'選択中 ✓':'選択する →'}</em></button>`;
 }).join('');
 document.querySelectorAll('[data-step-src]').forEach(b=>b.onclick=()=>{
  const src=b.dataset.stepSrc;
  activeSources.has(src)?activeSources.delete(src):activeSources.add(src);
  if(!activeSources.size)activeSources.add(src);
  renderSourceStep();
 });
}
function startApp(){
 $('chooser').classList.add('hidden');$('app').classList.remove('hidden');$('changeProfile').classList.remove('hidden');
 syncProfile();populateThemes();render();scrollTo({top:0,behavior:'instant'});
}
function syncProfile(){
 const p=profileCode?PROFILES[profileCode]:null;
 $('breadcrumb').textContent=p?`ホーム ＞ ${p.group==='insurer'?'保険会社':'保険代理店'} ＞ ${p.label}`:'ホーム ＞ 属性指定なし';
 $('profileName').textContent=p?p.label:'属性を指定せず全て表示';
 $('profileDesc').textContent=p?p.desc:'4つの公表資料・全1,190件を属性分類なしで検索します。';
 $('profileMatchCount').textContent=profileCode?DATA.filter(r=>['must','related'].includes(relation(r))).length:DATA.length;
 document.querySelectorAll('.modeTab').forEach(b=>{
  const m=b.dataset.mode;
  b.classList.toggle('hidden',!profileCode&&['must','related','all'].includes(m));
  b.classList.toggle('active',m===displayMode);
 });
}

function buildSourceUI(){
 $('sourceBtns').innerHTML=SOURCE_ORDER.map(s=>`<button class="srcBtn ${activeSources.has(s)?'active':''}" data-src="${s}">${SOURCES[s].label}</button>`).join('');
 document.querySelectorAll('[data-src]').forEach(b=>b.onclick=()=>{
  const s=b.dataset.src;
  if(activeSources.has(s)){if(activeSources.size===1)return;activeSources.delete(s)}else activeSources.add(s);
  b.classList.toggle('active',activeSources.has(s));page=1;populateThemes();render();
 });
 $('sourceLinks').innerHTML=SOURCE_ORDER.map(s=>`<a href="${SOURCES[s].url}" target="_blank" rel="noopener">${SOURCES[s].label} 公表ページ ↗</a>`).join('');
}
function populateThemes(){
 const cur=$('theme').value,counts={};
 DATA.filter(r=>activeSources.has(r.source)).forEach(r=>counts[r.theme]=(counts[r.theme]||0)+1);
 let arr=Object.keys(counts);
 arr.sort((a,b)=>{const ia=OFFICIAL_THEME_ORDER.indexOf(a),ib=OFFICIAL_THEME_ORDER.indexOf(b);if(ia>=0||ib>=0){if(ia<0)return 1;if(ib<0)return-1;return ia-ib}return a.localeCompare(b,'ja')});
 $('theme').innerHTML='<option value="">すべて</option>'+arr.map(t=>`<option value="${esc(t)}">${esc(t)} (${counts[t]})</option>`).join('');
 if(arr.includes(cur))$('theme').value=cur;
 buildSourceUI();
}

function filtered(){
 const q=$('q').value.trim().toLowerCase(),qs=q.split(/\s+/).filter(Boolean);
 const theme=$('theme').value,from=+$('fromNo').value||1,to=+$('toNo').value||9999;
 let a=DATA.filter(r=>activeSources.has(r.source)&&r.no>=from&&r.no<=to&&(!theme||r.theme===theme));
 if(q)a=a.filter(r=>{const h=[r.no,textOf(r)].join(' ').toLowerCase();return qs.every(x=>h.includes(x))});
 if(displayMode==='fav')a=a.filter(r=>favs.has(key(r)));
 else if(profileCode){
  if(displayMode==='all')a=a.filter(r=>['must','related'].includes(relation(r)));
  if(displayMode==='must')a=a.filter(r=>relation(r)==='must');
  if(displayMode==='related')a=a.filter(r=>relation(r)==='related');
  // raw = no attribute filtering
 }
 return a.sort((a,b)=>sourceOrder(a.source)-sourceOrder(b.source)||a.no-b.no);
}

function updateBreakdown(a){
 const total=a.length,counts={};SOURCE_ORDER.forEach(s=>counts[s]=0);a.forEach(r=>counts[r.source]++);
 $('heroPieTotal').textContent=total.toLocaleString('ja-JP');
 let start=0,segs=[];
 SOURCE_ORDER.forEach(src=>{const c=counts[src],pct=total?c/total*100:0;if(c){segs.push(`${SOURCES[src].color} ${start}% ${start+pct}%`);start+=pct}});
 $('heroPie').style.background=segs.length?`conic-gradient(${segs.join(',')})`:'#ffffff28';
 $('heroLegend').innerHTML=SOURCE_ORDER.filter(src=>counts[src]).map(src=>{
  const pct=total?counts[src]/total*100:0;
  return `<div class="legendRow"><span class="dot" style="background:${SOURCES[src].color}"></span><span>${SOURCES[src].label}</span><b>${counts[src]}件 (${pct.toFixed(1)}%)</b></div>`;
 }).join('');
}

function render(){
 syncProfile();
 const all=filtered(),totalPages=Math.max(1,Math.ceil(all.length/pageSize));if(page>totalPages)page=totalPages;
 const rows=all.slice((page-1)*pageSize,page*pageSize);
 $('count').textContent=all.length+'件';$('shownCount').textContent=all.length;
 $('countNote').textContent=`（表示 ${all.length?((page-1)*pageSize+1):0}–${Math.min(page*pageSize,all.length)}）`;
 updateBreakdown(all);
 if(!rows.length){$('results').innerHTML='<div class="empty">条件に一致するコメントはありません。</div>';renderPager(totalPages);updateSelection();return}
 $('results').innerHTML=rows.map(r=>{
  const k=key(r),rel=relation(r),fav=favs.has(k);
  const relText=rel==='must'?'● 必須・直接該当':rel==='related'?'○ 推奨・関連':'◇ 原文全件';
  const tags=(r.tags||[]).slice(0,3).map(t=>`<span class="miniTag">#${esc(t)}</span>`).join('');
  return `<article class="result" data-key="${k}">
   <div class="resultHead">
    <input class="pick" type="checkbox" data-pick="${k}" ${selected.has(k)?'checked':''}>
    <div class="noBadge">No.${r.no}</div>
    <div class="relBadge ${rel}">${relText}</div>
    <div class="resultTitle"><b>${esc(r.theme||'その他')}</b><div class="snip">${hi(brief(r.comment||r.answer,180))}</div><div class="miniTags"><span class="miniTag">#${esc(r.sourceLabel)}</span>${tags}</div></div>
    <div class="headActions"><button class="iconBtn ${fav?'fav':''}" data-fav="${k}" title="保存">${fav?'★':'☆'}</button><button class="iconBtn" data-open="${k}" title="開く">＋</button></div>
   </div>
   <div class="detail hidden" id="detail-${k}">
    <div class="previewGrid">
     <div class="preview comment"><h4>🗣 コメント概要・冒頭</h4><p>${hi(brief(r.comment,220))}</p></div>
     <div class="preview answer"><h4>🏛 金融庁の考え方・冒頭</h4><p>${hi(brief(r.answer,220))}</p></div>
    </div>
    <details class="original"><summary>📄 コメント概要の原文</summary><div class="copyRow"><button class="copyBtn" data-copy="comment-${k}">コピー</button></div><div class="originalText" id="comment-${k}">${hi(r.comment)}</div></details>
    <details class="original"><summary>🏛 金融庁の考え方の原文</summary><div class="copyRow"><button class="copyBtn" data-copy="answer-${k}">コピー</button></div><div class="originalText" id="answer-${k}">${hi(r.answer)}</div></details>
    <div class="metaRow"><span class="metaPill">${esc(r.sourceLabel)}</span><span class="metaPill">${esc(r.section||'')}</span><span class="metaPill">PDF ${r.pageStart}${r.pageEnd>r.pageStart?'–'+r.pageEnd:''}頁</span><a class="metaPill" href="${SOURCES[r.source].url}" target="_blank" rel="noopener">金融庁公表ページ ↗</a></div>
    <details class="noteBox"><summary style="font-size:10px;font-weight:900;cursor:pointer">自分用メモ</summary><textarea data-note="${k}" placeholder="確認事項、実務メモなど">${esc(notes[k]||'')}</textarea></details>
   </div>
  </article>`;
 }).join('');
 bindResults();renderPager(totalPages);updateSelection();
}
function bindResults(){
 document.querySelectorAll('[data-open]').forEach(b=>b.onclick=()=>{const d=$('detail-'+b.dataset.open);const open=d.classList.toggle('hidden')===false;b.textContent=open?'×':'＋'});
 document.querySelectorAll('[data-fav]').forEach(b=>b.onclick=()=>{const k=b.dataset.fav;favs.has(k)?favs.delete(k):favs.add(k);localStorage.setItem(KEY+'favs',JSON.stringify([...favs]));render()});
 document.querySelectorAll('[data-pick]').forEach(c=>c.onchange=()=>{c.checked?selected.add(c.dataset.pick):selected.delete(c.dataset.pick);updateSelection()});
 document.querySelectorAll('[data-note]').forEach(t=>t.oninput=()=>{notes[t.dataset.note]=t.value;localStorage.setItem(KEY+'notes',JSON.stringify(notes))});
 document.querySelectorAll('[data-copy]').forEach(b=>b.onclick=()=>{const el=$(b.dataset.copy);navigator.clipboard?.writeText(el.innerText);b.textContent='コピー済';setTimeout(()=>b.textContent='コピー',800)});
}
function renderPager(total){
 if(total<=1){$('pager').innerHTML='';return}
 const pages=[];for(let i=Math.max(1,page-2);i<=Math.min(total,page+2);i++)pages.push(i);
 $('pager').innerHTML=(page>1?`<button data-page="${page-1}">‹</button>`:'')+pages.map(p=>`<button class="${p===page?'active':''}" data-page="${p}">${p}</button>`).join('')+(page<total?`<button data-page="${page+1}">›</button>`:'');
 document.querySelectorAll('[data-page]').forEach(b=>b.onclick=()=>{page=+b.dataset.page;render();scrollTo({top:$('profileName').getBoundingClientRect().top+scrollY-80,behavior:'smooth'})});
}
function updateSelection(){
 $('selBar').classList.toggle('hidden',selected.size===0);$('selCount').textContent=selected.size+'件選択';
 $('compareSelected').disabled=selected.size<2||selected.size>4;
}
function selectedRecords(){return [...selected].map(k=>DATA.find(r=>key(r)===k)).filter(Boolean)}
function printSelected(){
 const arr=selectedRecords();if(!arr.length)return;
 const w=window.open('','_blank');if(!w){alert('ポップアップを許可してください');return}
 const rows=arr.map(r=>`<article><h2>${esc(r.sourceLabel)} No.${r.no}　${esc(r.theme)}</h2><h3>コメントの概要</h3><div>${esc(r.comment)}</div><h3>金融庁の考え方</h3><div>${esc(r.answer)}</div></article>`).join('');
 w.document.write(`<html lang="ja"><head><meta charset="utf-8"><title>金融庁パブコメ選択結果</title><style>@page{size:A4;margin:14mm}body{font-family:sans-serif;font-size:10pt;line-height:1.6}article{break-inside:avoid-page;border-bottom:1px solid #aaa;padding-bottom:12px;margin-bottom:15px}h2{font-size:13pt;color:#075f59}h3{font-size:10pt;border-left:4px solid #087b72;padding-left:6px}</style></head><body><h1>金融庁パブリックコメント 選択結果</h1>${rows}</body></html>`);w.document.close();setTimeout(()=>w.print(),200);
}
function compareSelected(){
 const arr=selectedRecords();$('compareGrid').innerHTML=arr.map(r=>`<div class="compareCard"><b>${esc(r.sourceLabel)} No.${r.no}</b><h3>${esc(r.theme)}</h3><h4>コメント</h4><div>${hi(r.comment)}</div><h4>金融庁の考え方</h4><div>${hi(r.answer)}</div></div>`).join('');$('compareModal').classList.remove('hidden');
}
function exportCsv(){
 const a=filtered(),q=x=>'"'+String(x??'').replaceAll('"','""')+'"';
 const rows=[['資料','No.','テーマ','区分','関連度','コメントの概要','金融庁の考え方'],...a.map(r=>[r.sourceLabel,r.no,r.theme,r.section,relation(r),r.comment,r.answer])];
 const blob=new Blob(['\uFEFF'+rows.map(r=>r.map(q).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}),u=URL.createObjectURL(blob),x=document.createElement('a');x.href=u;x.download='fsa_public_comments.csv';x.click();URL.revokeObjectURL(u);
}

/* controls */
$('consentCheck').onchange=()=>$('consentBtn').disabled=!$('consentCheck').checked;
$('consentBtn').onclick=()=>{$('consent').classList.add('hidden');sessionStorage.setItem(KEY+'consent','1')};
if(sessionStorage.getItem(KEY+'consent')==='1')$('consent').classList.add('hidden');
document.querySelectorAll('[data-group]').forEach(b=>b.onclick=()=>showStep2(b.dataset.group));
$('backGroup').onclick=showStep1;
$('backProfile').onclick=()=>{if(profileCode&&PROFILES[profileCode])showStep2(PROFILES[profileCode].group);else showStep1()};
$('sourceAllAndStart').onclick=()=>{activeSources=new Set(SOURCE_ORDER);startApp()};
$('sourceStart').onclick=()=>{if(!activeSources.size)activeSources=new Set(SOURCE_ORDER);startApp()};
$('showAllNoAttr').onclick=()=>selectProfile(null);
$('changeProfile').onclick=()=>{$('app').classList.add('hidden');$('chooser').classList.remove('hidden');$('changeProfile').classList.add('hidden');showStep1()};
document.querySelectorAll('[data-mode]').forEach(b=>b.onclick=()=>{displayMode=b.dataset.mode;page=1;render()});
$('q').oninput=()=>{page=1;render()};
document.querySelectorAll('.quickKeyword').forEach(b=>b.onclick=()=>{$('q').value=b.textContent.trim();page=1;render();$('q').focus()});
$('clearQ').onclick=()=>{$('q').value='';page=1;render();$('q').focus()};
$('theme').onchange=()=>{page=1;render()};
$('fromNo').oninput=()=>{page=1;render()};
$('toNo').oninput=()=>{page=1;render()};
$('toggleAdvanced').onclick=()=>{
 const closed=$('advanced').classList.toggle('hidden');
 $('toggleAdvanced').textContent=closed?'詳細条件を開く':'詳細条件を閉じる';
};
$('expandAll').onclick=()=>document.querySelectorAll('.detail').forEach(d=>d.classList.remove('hidden'));
$('collapseAll').onclick=()=>document.querySelectorAll('.detail').forEach(d=>d.classList.add('hidden'));
$('csv').onclick=exportCsv;
$('clearSelected').onclick=()=>{selected.clear();render()};
$('printSelected').onclick=printSelected;
$('compareSelected').onclick=compareSelected;
$('closeCompare').onclick=()=>$('compareModal').classList.add('hidden');

showStep1();populateThemes();
