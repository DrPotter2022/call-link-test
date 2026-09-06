// app.js — 全体の状態管理とイベント配線
const AppController = (function () {
  const state = {
    mode: 'count',
    region: 'all',
    operator: 'all',
    expansionOn: false,
    compareMode: false,
    compareSelection: []
  };

  function filteredClusters() {
    let items = DataStore.clusters;
    if (state.region !== 'all') items = items.filter(it => it.region === state.region);
    if (state.operator !== 'all') {
      if (state.operator === '__other') {
        const known = ['Amazon Web Services', 'Microsoft', 'Google', 'Meta', 'Oracle', 'OpenAI', 'xAI', 'CoreWeave', 'Alibaba Cloud', 'Tencent Cloud', 'ByteDance', 'SoftBank', 'NTT'];
        items = items.filter(it => (it.operators || []).some(op => !known.includes(op)));
      } else {
        items = items.filter(it => (it.operators || []).includes(state.operator));
      }
    }
    return items;
  }

  function filteredCountries() {
    let items = DataStore.countries;
    if (state.region !== 'all') items = items.filter(c => c.region === state.region);
    return items;
  }

  function render() {
    try {
      renderInner();
    } catch (err) {
      console.warn('[AI DCマップ] 地図コンテナのサイズ確定前に描画されたため再試行します:', err.message);
      // 地図コンテナのサイズが確定する前に描画された場合など、一時的な描画失敗はリトライする
      const map = MapModule.getMap();
      if (map) map.invalidateSize();
      setTimeout(() => {
        try { renderInner(); } catch (err2) { console.error('[AI DCマップ] 描画エラー:', err2); }
      }, 150);
    }
  }

  function renderInner() {
    MapModule.clearDataLayers();
    const clusters = filteredClusters();
    const countries = filteredCountries();

    if (state.mode === 'count') {
      MapModule.renderCountMode(countries);
    } else if (state.mode === 'hyperscale') {
      MapModule.renderHyperscaleMode(clusters);
    } else if (state.mode === 'cluster') {
      MapModule.renderClusterMode(clusters);
    } else if (state.mode === 'power') {
      MapModule.renderPowerMode(clusters);
    }

    MapModule.renderFlows(DataStore.flows, state.expansionOn);
    UIModule.renderLegend(state.mode);
    UIModule.renderStats(state.mode, { countries, clusters });

    // 企業フィルターの表示切替 (mode②④のみ)
    document.getElementById('operator-filter-group').style.display =
      (state.mode === 'hyperscale' || state.mode === 'power') ? 'flex' : 'none';
  }

  function setMode(mode) {
    state.mode = mode;
    document.querySelectorAll('.mode-btn').forEach(b => b.classList.toggle('active', b.dataset.mode === mode));
    render();
  }

  function setRegion(region) {
    state.region = region;
    document.querySelectorAll('.chip[data-region]').forEach(b => b.classList.toggle('active', b.dataset.region === region));
    MapModule.zoomToRegion(region);
    render();
  }

  function setOperator(op) {
    state.operator = op;
    document.querySelectorAll('.chip[data-op]').forEach(b => b.classList.toggle('active', b.dataset.op === op));
    render();
  }

  function toggleExpansion() {
    state.expansionOn = !state.expansionOn;
    document.getElementById('expansion-switch').classList.toggle('on', state.expansionOn);
    render();
  }

  function toggleCompare() {
    state.compareMode = !state.compareMode;
    state.compareSelection = [];
    document.getElementById('compare-btn').classList.toggle('active', state.compareMode);
    if (state.compareMode) {
      UIModule.showComparePanel(state.compareSelection);
    } else {
      UIModule.showEmptyPanel();
    }
  }

  function onZoomChange(zoom) {
    render();
  }

  function selectCountry(c) {
    if (state.compareMode) {
      const exists = state.compareSelection.find(x => x.code === c.code);
      if (exists) {
        state.compareSelection = state.compareSelection.filter(x => x.code !== c.code);
      } else if (state.compareSelection.length < 3) {
        state.compareSelection.push(c);
      }
      UIModule.showComparePanel(state.compareSelection);
      return;
    }
    const total = DataStore.countries.reduce((s, x) => s + (x.dataCenterCount || 0), 0);
    UIModule.showCountryPanel(c, total);
  }

  function selectItem(it) {
    UIModule.showItemPanel(it);
  }

  function selectFlow(f) {
    UIModule.showFlowPanel(f);
  }

  function bindEvents() {
    document.querySelectorAll('.mode-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        if (e.target.classList.contains('info-dot')) return;
        setMode(btn.dataset.mode);
      });
    });

    document.querySelectorAll('.info-dot').forEach(dot => {
      dot.addEventListener('click', (e) => {
        e.stopPropagation();
        UIModule.showInfoTooltip(dot, dot.dataset.info);
      });
    });

    document.querySelectorAll('.chip[data-region]').forEach(btn => {
      btn.addEventListener('click', () => setRegion(btn.dataset.region));
    });

    document.querySelectorAll('.chip[data-op]').forEach(btn => {
      btn.addEventListener('click', () => setOperator(btn.dataset.op));
    });

    document.getElementById('expansion-switch').addEventListener('click', toggleExpansion);
    document.getElementById('compare-btn').addEventListener('click', toggleCompare);
  }

  async function init() {
    const statusEl = document.getElementById('load-status');
    try {
      MapModule.init();
      await DataStore.loadAll();
      document.getElementById('meta-updated').textContent = `データ最終更新：${DataStore.lastUpdated}`;
      bindEvents();
      MapModule.getMap().invalidateSize();
      await new Promise(resolve => requestAnimationFrame(resolve));
      setMode('count');
      UIModule.showEmptyPanel();
      statusEl.classList.add('hidden');
    } catch (err) {
      console.error('[AI DCマップ] データ読み込みエラー:', err);
      statusEl.classList.remove('error');
      statusEl.classList.add('error');
      if (err && err.message === 'FILE_PROTOCOL') {
        statusEl.innerHTML = 'ローカルサーバーから起動してください。<br><br>例：<code>python -m http.server 8000</code><br>を実行し、<code>http://localhost:8000</code> を開いてください。<br><br>(file:// で直接開くとブラウザのCORS制限によりデータを読み込めません)';
      } else {
        statusEl.innerHTML = `データを読み込めませんでした。<br><br>ブラウザのConsoleで詳細をご確認ください。<br><span style="font-size:11px;color:#999;">${(err && err.message) || err}</span>`;
      }
    }
  }

  return { init, onZoomChange, selectCountry, selectItem, selectFlow };
})();

window.AppController = AppController;

window.addEventListener('DOMContentLoaded', () => {
  AppController.init();
});
