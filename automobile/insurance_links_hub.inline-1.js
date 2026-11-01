// ==== 設定用データ（ここを書き換えるだけで更新できます）====
const DATA = [
  {
    id: "spread_auto_insurance_adoption",
    title: "自動車保険の普及率",
    emoji: "📈",
    items: [
        { title: "都道府県別 普及率（2013-2024）", url: "https://www.wize.uno/automobile/enroll-insurance.html", desc: "損害保険料算出機構「自動車保険の概況」より集計", tags: ["普及率","都道府県","自動車保険"] },
    ]
  },
  {
    id: "lost_profit_calc",
    title: "逸失利益　簡易計算ツール",
    emoji: "🧮",
    items: [
        { title: "逸失利益のざっくり試算", url: "https://www.wize.uno/automobile/embedded.html", desc: "基礎収入×労働能力喪失率×ライプニッツ係数などを手早く", tags: ["逸失利益","人損","交通事故","ライプニッツ係数","ライプニッツ"] },
    ]
  },
  {
    id: "claim_enrollment",
    title: "自動車保険の「保険金」加入割合",
    emoji: "🧾",
    items: [
      // セクション内の小分類はタグで管理します（対人賠償・対物賠償・人身傷害）
       { title: "対人賠償の加入割合", url: "https://www.wize.uno/automobile/compensation.html", desc: "対人賠償の全国加入率", tags: ["加入割合","対人賠償","対人"] },
       { title: "対物賠償の加入割合", url: "https://www.wize.uno/automobile/property.html", desc: "対物賠償の全国加入率", tags: ["加入割合","対物賠償","対物"] },
       { title: "人身傷害の加入割合", url: "https://www.wize.uno/automobile/injury.html", desc: "人身傷害の全国加入率", tags: ["加入割合","人身傷害","人身"] },
    ]
  }
];

// 任意でよく使うタグ（クリックでフィルタ）
const FAVORITE_TAGS = ["普及率","加入割合","対人賠償","対物賠償","人身傷害","都道府県","人損","交通事故","資料","ツール"];

// ==== 以下はUIロジック（編集不要）====
const qs = (sel, el=document) => el.querySelector(sel);
const qsa = (sel, el=document) => [...el.querySelectorAll(sel)];

const chipRow = qs("#chipRow");
const sectionsEl = qs("#sections");
const searchInput = qs("#search");

let activeTag = null;
let keyword = "";

function buildChips() {
  chipRow.innerHTML = "";
  FAVORITE_TAGS.forEach(tag => {
    const b = document.createElement("button");
    b.className = "chip";
    b.textContent = `#${tag}`;
    b.onclick = () => {
      if (activeTag === tag) { activeTag = null; }
      else { activeTag = tag; }
      updateChips();
      render();
    };
    chipRow.appendChild(b);
  });
  updateChips();
}
function updateChips() {
  qsa(".chip", chipRow).forEach(btn => {
    const tag = btn.textContent.replace(/^#/, "");
    btn.classList.toggle("active", tag === activeTag);
  });
}

function matchItem(item) {
  const kw = keyword.trim().toLowerCase();
  const tagOk = !activeTag || (item.tags||[]).includes(activeTag);
  const kwOk = !kw ||
    (item.title && item.title.toLowerCase().includes(kw)) ||
    (item.desc && item.desc.toLowerCase().includes(kw)) ||
    (item.url && item.url.toLowerCase().includes(kw)) ||
    (item.tags && item.tags.some(t => t.toLowerCase().includes(kw)));
  return tagOk && kwOk;
}

function card(item) {
  const el = document.createElement("div");
  el.className = "card";
  el.innerHTML = `
    <div class="top">
      <div class="emoji">🔗</div>
      <div class="title">${escapeHtml(item.title || "(無題)")}</div>
    </div>
    <div class="desc">${escapeHtml(item.desc || "")}</div>
    <div class="url" title="${escapeHtml(item.url||"")}">${escapeHtml(item.url || "")}</div>
    <div class="tags">${(item.tags||[]).map(t => `<span class="tag">${escapeHtml(t)}</span>`).join("")}</div>
    <div class="actions">
      ${item.url ? `<a class="btn external" target="_blank" rel="noopener" href="${escapeAttr(item.url)}">開く</a>` : `<button class="btn" disabled>URL未設定</button>`}
      ${item.url ? `<button class="btn" data-copy="${escapeAttr(item.url)}">URLコピー</button>` : ""}
    </div>
  `;
  const copyBtn = el.querySelector('[data-copy]');
  if (copyBtn) {
    copyBtn.addEventListener("click", async (e) => {
      const url = e.currentTarget.getAttribute("data-copy");
      try {
        await navigator.clipboard.writeText(url);
        e.currentTarget.textContent = "コピーしました";
        setTimeout(() => (e.currentTarget.textContent = "URLコピー"), 1200);
      } catch {
        prompt("コピーできない場合は手動でコピーしてください:", url);
      }
    });
  }
  return el;
}

function section(sec) {
  const wrap = document.createElement("section");
  const items = (sec.items||[]).filter(matchItem);
  wrap.innerHTML = `
    <div class="sec-head">
      <h2>${sec.emoji || "📚"} ${escapeHtml(sec.title)}</h2>
      <div class="count">${items.length} 件</div>
    </div>
  `;
  if (items.length === 0) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.innerHTML = "該当するリンクがありません。<br>右上の検索・タグを変更するか、下記のJSONにURLを追加してください。";
    wrap.appendChild(empty);
    return wrap;
  }
  const grid = document.createElement("div");
  grid.className = "grid";
  items.forEach(it => grid.appendChild(card(it)));
  wrap.appendChild(grid);
  return wrap;
}

function render() {
  sectionsEl.innerHTML = "";
  DATA.forEach(s => sectionsEl.appendChild(section(s)));
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, m => ({
    "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"
  }[m]));
}
function escapeAttr(s) { return escapeHtml(s); }

searchInput.addEventListener("input", (e) => {
  keyword = e.target.value || "";
  render();
});

buildChips();
render();

