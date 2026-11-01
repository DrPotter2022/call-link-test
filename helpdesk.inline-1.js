const LOG_KEY='helpdesk_phonebook_logs_v3';
const APPS_SCRIPT_WEB_APP_URL='https://script.google.com/macros/s/AKfycbxkcp4LpbJhd_ec_PoYMwPqilggK8WTVRXaGQs0Jvnoy8hh71c6dKK96zAjLk6uYZSn/exec';
let currentCode='';

function isRemoteEnabled(){
  return APPS_SCRIPT_WEB_APP_URL && !APPS_SCRIPT_WEB_APP_URL.includes('PASTE_APPS_SCRIPT');
}

function getLogs(){
  try{return JSON.parse(localStorage.getItem(LOG_KEY)||'[]')}catch(e){return []}
}

function saveLogs(logs){
  localStorage.setItem(LOG_KEY,JSON.stringify(logs))
}

function localLog(entry){
  const logs=getLogs();
  logs.push(entry);
  saveLogs(logs.slice(-1000));
}

function sendRemoteLog(entry){
  if(!isRemoteEnabled()) return;
  try{
    const payload = new URLSearchParams();
    Object.entries(entry).forEach(([k,v])=>payload.append(k, String(v ?? '')));
    fetch(APPS_SCRIPT_WEB_APP_URL, {
      method:'POST',
      mode:'no-cors',
      headers:{'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8'},
      body:payload
    });
  }catch(e){}
}

function addLog(type,target,extra={}){
  const entry={
    time:new Date().toISOString(),
    code:currentCode||'',
    type,
    target,
    company:extra.company||'',
    label:extra.label||'',
    phone:extra.phone||'',
    sheet:document.querySelector('.section.active')?.id==='life'?'生命保険':'損害保険',
    userAgent:navigator.userAgent,
    page:location.pathname.split('/').pop()||'helpdesk_phonebook.html'
  };
  localLog(entry);
  sendRemoteLog(entry);
}


async function authByAppsScript(code){
  if(!isRemoteEnabled()) {
    throw new Error('Apps Script URLが未設定です。');
  }

  const url = APPS_SCRIPT_WEB_APP_URL
    + '?action=auth'
    + '&code=' + encodeURIComponent(code)
    + '&t=' + Date.now();

  const res = await fetch(url, {
    method: 'GET',
    mode: 'cors',
    credentials: 'omit',
    cache: 'no-store',
    redirect: 'follow'
  });

  if(!res.ok){
    throw new Error('認証サーバーに接続できません。');
  }

  const text = await res.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch(e) {
    throw new Error('認証サーバーの応答を読み取れません。');
  }

  return !!data.ok;
}

async function authenticate(){
  const input=document.getElementById('staffCode');
  const btn=document.getElementById('authBtn');
  const err=document.getElementById('authError');
  const code=(input.value||'').trim();

  if(!/^.+\d{6}$/.test(code)){
    err.textContent='担当者コードの形式が正しくありません。';
    return;
  }

  btn.disabled=true;
  btn.textContent='確認中...';
  err.textContent='';

  try{
    const ok=await authByAppsScript(code);
    if(!ok){
      err.textContent='担当者コードが正しくありません。';
      return;
    }
    currentCode=code;
    document.body.classList.remove('locked');
    document.getElementById('authOverlay').style.display='none';
    addLog('login','page_open');
  }catch(e){
    err.textContent=e.message || '認証に失敗しました。';
  }finally{
    btn.disabled=false;
    btn.textContent='同意する';
  }
}

document.getElementById('authBtn').addEventListener('click',authenticate);
document.getElementById('staffCode').addEventListener('keydown',e=>{if(e.key==='Enter')authenticate();});

document.querySelectorAll('.tab').forEach(btn=>btn.addEventListener('click',()=>{
  document.querySelectorAll('.tab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.getElementById(btn.dataset.target).classList.add('active');
  applySearch();
}));

function applySearch(){
  const q=(document.getElementById('q').value||'').toLowerCase().trim();
  document.querySelectorAll('.company').forEach(c=>{
    const hit=!q||c.textContent.toLowerCase().includes(q);
    c.style.display=hit?'':'none';
  });
}
document.getElementById('q').addEventListener('input',applySearch);

document.querySelectorAll('a[href^="tel:"]').forEach(a=>{
  a.addEventListener('click',()=>{
    const entry=a.closest('.entry');
    const company=a.closest('.company')?.dataset.company||a.closest('.company')?.querySelector('summary span')?.textContent?.trim()||'';
    const label=entry?.querySelector('.item')?.textContent?.trim()||'';
    const phone=a.textContent.trim();
    addLog('tel','call_tap',{company,label,phone});
  });
});
