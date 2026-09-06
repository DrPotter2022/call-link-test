// ui.js — サイドパネル・凡例・統計カード・ツールチップの描画
const UIModule = (function () {

  const INFO_TEXT = {
    count: '「データセンター総数」とは、各国・地域に立地する物理的なデータセンター施設の数です。<br><br><strong>注意：</strong>データセンター数が多い＝AI計算能力が大きい、とは限りません。従来型の小規模施設が多数を占める国もあれば、少数の超大型施設に集中する国もあります。',
    hyperscale: '「Hyperscale」とは、AWS・Microsoft・Google・Metaなど大手クラウド/AI事業者が運営する大規模データセンターを指します。AIトレーニング・推論向けの専用施設を含みます。',
    cluster: '「クラスター」とは、地理的に近接した複数のデータセンターが集積している地域を指します。単独施設ではなく地域全体の傾向を示します。',
    power: '電力規模（MW/GW）を見る理由：GPU台数や計算能力は非公表であることが多い一方、電力容量は事業者・自治体の発表や送電契約から把握しやすく、AI計算インフラの実質的な規模を測る代表的な指標として使われます。'
  };

  function showInfoTooltip(target, key) {
    document.querySelectorAll('.info-tooltip').forEach(el => el.remove());
    const tip = document.createElement('div');
    tip.className = 'info-tooltip';
    tip.innerHTML = INFO_TEXT[key] || '';
    document.body.appendChild(tip);
    const rect = target.getBoundingClientRect();
    let top = rect.bottom + 8;
    let left = rect.left;
    if (left + 280 > window.innerWidth) left = window.innerWidth - 290;
    tip.style.top = top + 'px';
    tip.style.left = Math.max(8, left) + 'px';
    setTimeout(() => {
      document.addEventListener('click', function handler(e) {
        if (!tip.contains(e.target)) { tip.remove(); document.removeEventListener('click', handler); }
      });
    }, 10);
  }

  // ---------- 凡例 ----------
  const LEGEND_THRESHOLDS = {
    count: [
      { label: '0～20', color: '#3b82f6' },
      { label: '20～100', color: '#60a5fa' },
      { label: '100～300', color: '#f59e0b' },
      { label: '300～1,000', color: '#f97316' },
      { label: '1,000+', color: '#dc2626' }
    ],
    power: [
      { label: '0～20MW', color: '#3b82f6' },
      { label: '20～100MW', color: '#60a5fa' },
      { label: '100～500MW', color: '#f59e0b' },
      { label: '500MW～1GW', color: '#f97316' },
      { label: '1GW+', color: '#dc2626' }
    ]
  };

  function renderLegend(mode) {
    const box = document.getElementById('legend-box');
    let html = '';
    if (mode === 'count') {
      html += '<h4>データセンター総数</h4>';
      LEGEND_THRESHOLDS.count.forEach(r => {
        html += `<div class="legend-row"><span class="legend-swatch" style="background:${r.color}"></span>${r.label}</div>`;
      });
    } else if (mode === 'power') {
      html += '<h4>電力容量 (MW)</h4>';
      LEGEND_THRESHOLDS.power.forEach(r => {
        html += `<div class="legend-row"><span class="legend-swatch" style="background:${r.color}"></span>${r.label}</div>`;
      });
      html += '<div class="legend-row" style="margin-top:4px;"><span class="legend-swatch" style="background:#c3ccd6"></span>電力規模 非公表</div>';
    } else if (mode === 'hyperscale') {
      html += '<h4>AI / Hyperscale拠点</h4>';
      html += '<div class="legend-row"><span class="legend-swatch" style="background:#1f6fd6;border-radius:50%;"></span>拠点</div>';
    } else if (mode === 'cluster') {
      html += '<h4>主要集積地域</h4>';
      html += '<div class="legend-row"><span class="legend-swatch" style="background:#1f6fd6;opacity:0.5;"></span>AI/Hyperscale集積地</div>';
      html += '<div class="legend-row"><span class="legend-swatch" style="background:#8ba0b8;opacity:0.5;"></span>その他DC集積地</div>';
    }
    html += `<div class="legend-status-row">
      <span>● 稼働中</span><span>▲ 建設中</span><span>○ 計画</span>
    </div>`;
    box.innerHTML = html;
  }

  // ---------- 統計カード ----------
  function renderStats(mode, ctx) {
    const bar = document.getElementById('stats-bar');
    const cards = [];
    const { countries, clusters } = ctx;

    if (mode === 'count') {
      const total = countries.reduce((s, c) => s + (c.dataCenterCount || 0), 0);
      const us = countries.find(c => c.code === 'US');
      const usShare = us && total ? ((us.dataCenterCount / total) * 100).toFixed(0) + '%' : 'N/A';
      const known = countries.filter(c => c.dataCenterCount).length;
      cards.push(card('世界DC総数(登録国合計)', total.toLocaleString('ja-JP') + '件', '公開データのみ'));
      cards.push(card('米国シェア', usShare, '登録国合計に対する比率'));
      cards.push(card('最大国', us ? (us.nameJa) : 'N/A', ''));
      cards.push(card('データ登録国数', known + '/' + countries.length, ''));
    } else if (mode === 'hyperscale') {
      const hs = clusters.filter(c => c.hyperscale);
      const construct = hs.filter(c => c.status === 'under_construction').length;
      const opCount = {};
      hs.forEach(it => (it.operators || []).forEach(op => { opCount[op] = (opCount[op] || 0) + 1; }));
      const topOp = Object.entries(opCount).sort((a, b) => b[1] - a[1])[0];
      cards.push(card('Hyperscale拠点数', hs.length + '件', '公開データのみ'));
      cards.push(card('建設中施設数', construct + '件', ''));
      cards.push(card('主要事業者', topOp ? topOp[0] : 'N/A', '登場回数最多'));
      cards.push(card('稼働状況区分', '稼働/建設中/計画', 'ステータス別に表示'));
    } else if (mode === 'cluster') {
      const cl = clusters.filter(c => c.granularity === 'cluster');
      const hsCl = cl.filter(c => c.hyperscale).length;
      const regions = new Set(cl.map(c => c.region));
      cards.push(card('登録クラスター数', cl.length + '件', '公開データのみ'));
      cards.push(card('AI/Hyperscale集積地', hsCl + '件', ''));
      cards.push(card('対象地域数', regions.size + '地域', ''));
      cards.push(card('最大集積地域', 'Northern Virginia', '世界最大級とされる'));
    } else if (mode === 'power') {
      const withPower = clusters.filter(c => typeof c.powerMW === 'number');
      const totalMW = withPower.reduce((s, c) => s + c.powerMW, 0);
      const maxItem = withPower.sort((a, b) => b.powerMW - a.powerMW)[0];
      const construct = clusters.filter(c => c.status === 'under_construction' && typeof c.powerMW === 'number');
      const constructMW = construct.reduce((s, c) => s + c.powerMW, 0);
      const gigaPlans = clusters.filter(c => typeof c.powerMW === 'number' && c.powerMW >= 1000).length;
      cards.push(card('確認可能な総電力容量', totalMW ? totalMW.toLocaleString('ja-JP') + 'MW' : 'N/A', '公開データのみ・粒度混在に注意'));
      cards.push(card('最大キャンパス', maxItem ? (maxItem.nameJa || maxItem.name) : 'N/A', maxItem ? maxItem.powerMW.toLocaleString('ja-JP') + 'MW' : ''));
      cards.push(card('1GW超の計画/施設数', gigaPlans + '件', ''));
      cards.push(card('建設中の確認可能容量', constructMW ? constructMW.toLocaleString('ja-JP') + 'MW' : 'N/A', ''));
    }

    bar.innerHTML = cards.join('');
  }

  function card(label, value, sub) {
    return `<div class="stat-card">
      <div class="stat-label">${label}</div>
      <div class="stat-value">${value}</div>
      ${sub ? `<div class="stat-sub">${sub}</div>` : ''}
    </div>`;
  }

  // ---------- サイドパネル ----------
  function qTag(q) {
    if (q === 'confirmed') return '<span class="quality-tag quality-confirmed">確認値</span>';
    if (q === 'estimated') return '<span class="quality-tag quality-estimated">推定値</span>';
    return '<span class="quality-tag quality-unknown">非公表</span>';
  }

  function field(label, value) {
    return `<div class="panel-field"><div class="pf-label">${label}</div><div class="pf-value">${value}</div></div>`;
  }

  function showEmptyPanel() {
    const panel = document.getElementById('side-panel');
    panel.innerHTML = `<div class="panel-drag-handle" id="panel-drag-handle"></div><div class="panel-empty">地図上の地点を選択してください</div>`;
    bindDragHandle();
    panel.classList.remove('expanded');
  }

  function showCountryPanel(c, total) {
    const panel = document.getElementById('side-panel');
    const share = c.dataCenterCount && total ? ((c.dataCenterCount / total) * 100).toFixed(0) + '%' : 'N/A';
    let html = `<div class="panel-drag-handle" id="panel-drag-handle"></div>`;
    html += `<h2>${c.nameJa} <span style="font-weight:400;font-size:12px;color:#8a94a1;">${c.name}</span></h2>`;
    html += `<div class="panel-sub">国別データセンター概況</div>`;
    html += field('データセンター数', fmtNum(c.dataCenterCount) + ' ' + qTag(c.dataCenterCountQuality));
    html += field('世界順位', c.rank ? `第${c.rank}位` : '非公表');
    html += field('世界全体に占める比率(登録国内)', share);
    html += field('主要都市', (c.majorLocations || []).join(' / ') || '非公表');
    html += field('主要事業者', (c.majorOperators || []).join(' / ') || '非公表');
    if (c.sourceUrl) html += `<a class="panel-source-link" href="${c.sourceUrl}" target="_blank" rel="noopener">出典を見る</a>`;
    else html += field('出典', c.source || '非公表');
    panel.innerHTML = html;
    bindDragHandle();
    panel.classList.add('expanded');
  }

  function showItemPanel(it) {
    const panel = document.getElementById('side-panel');
    const statusLabel = MapModule.STATUS_LABEL_JA[it.status] || it.status || '不明';
    let html = `<div class="panel-drag-handle" id="panel-drag-handle"></div>`;
    html += `<h2>${it.nameJa || it.name}</h2>`;
    html += `<div class="panel-sub">${it.name} ｜ ${it.countryCode || ''}</div>`;
    html += `<span class="panel-status-badge status-${it.status}">${MapModule.STATUS_SYMBOL[it.status] || ''} ${statusLabel}</span>`;
    html += field('国', it.country || '不明');
    html += field('地域', it.region || '不明');
    html += field('運営事業者', (it.operators || []).join(' / ') || '非公表');
    html += field('区分', it.hyperscale ? 'AI / Hyperscale' : '従来型DC集積地');
    html += field('粒度', it.granularity === 'cluster' ? 'クラスター(地域)' : it.granularity === 'campus' ? 'キャンパス' : '施設');
    html += field('電力容量', fmtNum(it.powerMW, 'MW') + ' ' + qTag(it.powerMWQuality));
    html += field('GPU情報', it.gpuType ? `${it.gpuType}${it.gpuCount ? ' ／ 約' + it.gpuCount.toLocaleString('ja-JP') + '基 (報道ベース)' : ''}` : '非公表');
    html += field('稼働開始年', it.startYear || '非公表');
    html += field('完成予定年', it.completionYear || '非公表');
    if (it.aiUse) html += field('AI用途', it.aiUse);
    if (it.powerSituation) html += field('電力事情', it.powerSituation);
    if (it.locationFeatures) html += field('立地上の特徴', it.locationFeatures);
    if (it.notes) html += field('備考', it.notes);
    if (it.sourceUrl) html += `<a class="panel-source-link" href="${it.sourceUrl}" target="_blank" rel="noopener">出典を見る</a>`;
    else html += field('データ出典', it.source || '非公表');
    html += field('最終更新', it.sourceDate || DataStore.lastUpdated || '不明');
    panel.innerHTML = html;
    bindDragHandle();
    panel.classList.add('expanded');
  }

  function showFlowPanel(f) {
    const panel = document.getElementById('side-panel');
    let html = `<div class="panel-drag-handle" id="panel-drag-handle"></div>`;
    html += `<h2>インフラ移動・拡大</h2>`;
    html += field('移動元', f.fromName);
    html += field('移動先', f.toName);
    html += field('背景', f.background);
    html += field('理由', (f.reasonsJa || f.reasons || []).join(' / '));
    panel.innerHTML = html;
    bindDragHandle();
    panel.classList.add('expanded');
  }

  function showComparePanel(selected) {
    const panel = document.getElementById('side-panel');
    let html = `<div class="panel-drag-handle" id="panel-drag-handle"></div>`;
    html += `<h2>国・地域比較</h2><div class="panel-sub">最大3件まで選択できます</div>`;
    if (!selected.length) {
      html += `<div class="panel-empty" style="margin-top:20px;">地図上で比較したい国を最大3つクリックしてください</div>`;
    } else {
      html += `<table class="compare-table"><tr><th></th>${selected.map(c => `<th>${c.nameJa}</th>`).join('')}</tr>`;
      html += `<tr><td>DC総数</td>${selected.map(c => `<td>${fmtNum(c.dataCenterCount)}</td>`).join('')}</tr>`;
      html += `<tr><td>Hyperscale拠点</td>${selected.map(c => {
        const n = DataStore.clusters.filter(cl => cl.hyperscale && cl.countryCode === c.code).length;
        return `<td>${n || 'N/A'}</td>`;
      }).join('')}</tr>`;
      html += `<tr><td>主要クラスター</td>${selected.map(c => {
        const cl = DataStore.clusters.filter(x => x.granularity === 'cluster' && x.countryCode === c.code).map(x => x.nameJa).slice(0, 2).join('/');
        return `<td>${cl || 'N/A'}</td>`;
      }).join('')}</tr>`;
      html += `<tr><td>確認可能MW</td>${selected.map(c => {
        const items = DataStore.clusters.filter(x => x.countryCode === c.code && typeof x.powerMW === 'number');
        const sum = items.reduce((s, x) => s + x.powerMW, 0);
        return `<td>${items.length ? sum.toLocaleString('ja-JP') + 'MW' : 'N/A'}</td>`;
      }).join('')}</tr>`;
      html += `</table>`;
    }
    panel.innerHTML = html;
    bindDragHandle();
    panel.classList.add('expanded');
  }

  function bindDragHandle() {
    const handle = document.getElementById('panel-drag-handle');
    if (!handle) return;
    handle.addEventListener('click', () => {
      document.getElementById('side-panel').classList.toggle('expanded');
    });
  }

  function fmtNum(n, unit) {
    if (n === null || n === undefined) return '非公表';
    return n.toLocaleString('ja-JP') + (unit || '');
  }

  return {
    showInfoTooltip, renderLegend, renderStats,
    showEmptyPanel, showCountryPanel, showItemPanel, showFlowPanel, showComparePanel
  };
})();
