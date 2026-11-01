import { initializeApp } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-app.js";
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";

const firebaseConfig={apiKey:"AIzaSyB6Dq_RR65s0w-fYcfiNN7XELpS5TGf_oQ",authDomain:"index-6c554.firebaseapp.com",projectId:"index-6c554",appId:"1:571533029170:web:8d84392edc1f44ec45312b"};
const auth=getAuth(initializeApp(firebaseConfig));
const API={directory:"https://helpdeskdirectory-uk7kp3auqq-uc.a.run.app",log:"https://helpdesklog-uk7kp3auqq-uc.a.run.app"};
const login=document.getElementById("login"),directoryElement=document.getElementById("directory"),errorElement=document.getElementById("error"),signOutButton=document.getElementById("signOut");
const setError=(message)=>{errorElement.textContent=message};

async function api(name,options={}){const token=await auth.currentUser.getIdToken();const response=await fetch(API[name],{...options,headers:{Authorization:`Bearer ${token}`,...(options.headers||{})},cache:"no-store"});const body=await response.json().catch(()=>({}));if(!response.ok||!body.ok)throw new Error(body.error||"request_failed");return body}
async function logEvent(event){try{await api("log",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(event)})}catch(_){/* Logging must not affect the directory. */}}

function renderDirectory(html){if(typeof html!=="string")throw new Error("invalid_directory");const template=document.createElement("template");template.innerHTML=html;directoryElement.replaceChildren(template.content);directoryElement.querySelectorAll('a[href^="tel:"]').forEach((link)=>link.addEventListener("click",()=>{const entry=link.closest(".entry"),company=link.closest(".company");void logEvent({type:"tel",target:"call_tap",company:company?.querySelector("summary")?.textContent?.trim()||"",label:entry?.querySelector(".item")?.textContent?.trim()||"",phone:link.textContent.trim(),section:link.closest("section")?.id||"",page:"helpdesk.html"})}))}
async function loadDirectory(){const result=await api("directory");renderDirectory(result.html);login.classList.add("hidden");directoryElement.classList.remove("hidden");signOutButton.classList.remove("hidden");void logEvent({type:"login",target:"page_open",page:"helpdesk.html"})}

document.getElementById("signIn").addEventListener("click",async()=>{setError("");const email=document.getElementById("email").value.trim(),password=document.getElementById("password").value;if(!email||!password)return setError("メールアドレスとパスワードを入力してください。");try{await signInWithEmailAndPassword(auth,email,password)}catch(_){setError("ログインできません。入力内容または利用権限を確認してください。")}});
signOutButton.addEventListener("click",()=>signOut(auth));
onAuthStateChanged(auth,async(user)=>{if(!user){directoryElement.replaceChildren();directoryElement.classList.add("hidden");login.classList.remove("hidden");signOutButton.classList.add("hidden");return}try{await loadDirectory()}catch(error){await signOut(auth);setError(error.message==="helpdesk_access_required"?"このアカウントには利用権限がありません。":"電話帳を読み込めません。")}});
