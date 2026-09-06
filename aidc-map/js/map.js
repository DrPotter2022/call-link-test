// map.js — Leaflet地図の初期化と各モードの描画ロジック
const MapModule = (function () {
  let map = null;
  let heatLayer = null;
  let markerLayer = null;
  let circleLayer = null;
  let flowLayer = null;
  let currentZoom = 2;

  const STATUS_SYMBOL = {
    operational: '●',
    under_construction: '▲',
    planned: '○'
  };
  const STATUS_LABEL_JA = {
    operational: '稼働中',
    under_construction: '建設中',
    planned: '計画'
  };
  const STATUS_COLOR = {
    operational: '#16a34a',
    under_construction: '#d97706',
    planned: '#6b7280'
  };

  function init() {
    map = L.map('map', {
      center: [25, 20],
      zoom: 2,
      minZoom: 2,
      maxZoom: 12,
      worldCopyJump: true,
      zoomControl: false
    });

    L.control.zoom({ position: 'topright' }).addTo(map);

    // Esri World Light Gray Canvas: ビジネス向けの明るい配色、APIキー不要
    L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}', {
      attribution: 'Tiles &copy; Esri &mdash; Esri, DeLorme, NAVTEQ',
      maxZoom: 16
    }).addTo(map);
    L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Reference/MapServer/tile/{z}/{y}/{x}', {
      maxZoom: 16, opacity: 0.8
    }).addTo(map);

    markerLayer = L.layerGroup().addTo(map);
    circleLayer = L.layerGroup().addTo(map);
    flowLayer = L.layerGroup();

    map.on('zoomend', () => {
      currentZoom = map.getZoom();
      if (window.AppController) window.AppController.onZoomChange(currentZoom);
    });

    return map;
  }

  function clearDataLayers() {
    if (heatLayer) { map.removeLayer(heatLayer); heatLayer = null; }
    markerLayer.clearLayers();
    circleLayer.clearLayers();
  }

  function zoomToRegion(regionKey) {
    const bounds = {
      all: [[-55, -170], [70, 170]],
      'North America': [[15, -130], [60, -60]],
      'China': [[18, 95], [50, 130]],
      'Europe': [[35, -12], [65, 32]],
      'Japan': [[26, 128], [46, 146]],
      'Asia Pacific': [[-40, 70], [40, 140]]
    };
    const b = bounds[regionKey] || bounds.all;
    map.fitBounds(b, { animate: true, duration: 0.6 });
  }

  function qualityLabel(q) {
    if (q === 'confirmed') return '確認値';
    if (q === 'estimated') return '推定値';
    return '非公表';
  }

  function fmtNum(n, unit) {
    if (n === null || n === undefined) return '非公表';
    return n.toLocaleString('ja-JP') + (unit || '');
  }

  // ---------- Mode 1: データセンター総数 (国別ヒートマップ) ----------
  function renderCountMode(countries) {
    const points = [];
    let maxCount = 1;
    countries.forEach(c => { if (c.dataCenterCount) maxCount = Math.max(maxCount, c.dataCenterCount); });

    countries.forEach(c => {
      if (!c.dataCenterCount) return;
      const intensity = c.dataCenterCount / maxCount;
      // 複数点を撒いてヒートマップらしい面を作る（国の中心から少しばらす）
      const n = Math.max(3, Math.round(intensity * 18));
      for (let i = 0; i < n; i++) {
        const jitterLat = (Math.random() - 0.5) * 6;
        const jitterLng = (Math.random() - 0.5) * 8;
        points.push([c.lat + jitterLat, c.lng + jitterLng, intensity]);
      }
    });

    heatLayer = L.heatLayer(points, {
      radius: 34, blur: 28, maxZoom: 6, max: 1.0,
      gradient: { 0.2: '#3b82f6', 0.45: '#60a5fa', 0.65: '#f59e0b', 0.85: '#f97316', 1.0: '#dc2626' }
    }).addTo(map);

    // 国クリック用の透明マーカー
    countries.forEach(c => {
      const marker = L.circleMarker([c.lat, c.lng], {
        radius: 16, fillOpacity: 0, opacity: 0, weight: 0
      });
      marker.on('click', () => window.AppController.selectCountry(c));
      marker.bindTooltip(`${c.nameJa || c.name}: ${fmtNum(c.dataCenterCount)}件`, { direction: 'top' });
      markerLayer.addLayer(marker);

      // ズームが大きい場合は国名ラベルも表示
      if (currentZoom >= 4) {
        const label = L.marker([c.lat, c.lng], {
          icon: L.divIcon({
            className: '', html: `<div style="font-size:11px;font-weight:700;color:#0f2540;text-shadow:0 0 3px #fff, 0 0 3px #fff;">${c.nameJa}</div>`,
            iconSize: [0, 0]
          }),
          interactive: false
        });
        markerLayer.addLayer(label);
      }
    });
  }

  // ---------- Mode 2: AI / Hyperscale ----------
  function renderHyperscaleMode(items) {
    const hs = items.filter(it => it.hyperscale);
    hs.forEach(it => {
      const color = STATUS_COLOR[it.status] || '#6b7280';
      const symbol = STATUS_SYMBOL[it.status] || '●';
      const size = it.granularity === 'campus' || it.granularity === 'facility' ? 9 : 13;

      const icon = L.divIcon({
        className: 'marker-op',
        html: `<div class="dc-marker-dot" style="width:${size}px;height:${size}px;background:${color};"></div>`,
        iconSize: [size, size],
        iconAnchor: [size / 2, size / 2]
      });
      const marker = L.marker([it.lat, it.lng], { icon });
      marker.on('click', () => window.AppController.selectItem(it));

      let tooltip = `${symbol} ${it.nameJa || it.name}`;
      if (currentZoom >= 4) tooltip += ` (${STATUS_LABEL_JA[it.status] || ''})`;
      marker.bindTooltip(tooltip, { direction: 'top' });
      markerLayer.addLayer(marker);

      // クリック/タップしやすいよう見た目より大きい透明な当たり判定を重ねる
      const hitArea = L.circleMarker([it.lat, it.lng], { radius: 16, fillOpacity: 0, opacity: 0, weight: 0 });
      hitArea.on('click', () => window.AppController.selectItem(it));
      hitArea.bindTooltip(tooltip, { direction: 'top' });
      markerLayer.addLayer(hitArea);
    });
  }

  // ---------- Mode 3: 主要集積地域 (クラスター) ----------
  function renderClusterMode(items) {
    const clusters = items.filter(it => it.granularity === 'cluster');
    clusters.forEach(it => {
      const color = it.hyperscale ? '#1f6fd6' : '#8ba0b8';
      const circle = L.circle([it.lat, it.lng], {
        radius: 90000,
        color: color,
        weight: 2,
        fillColor: color,
        fillOpacity: 0.18
      });
      circle.on('click', () => window.AppController.selectItem(it));
      circle.bindTooltip(it.nameJa || it.name, { direction: 'top', sticky: true });
      circleLayer.addLayer(circle);

      if (currentZoom >= 3) {
        const label = L.marker([it.lat, it.lng], {
          icon: L.divIcon({
            className: '', html: `<div style="font-size:11px;font-weight:700;color:#0f2540;text-shadow:0 0 3px #fff, 0 0 3px #fff;white-space:nowrap;">${it.nameJa || it.name}</div>`,
            iconSize: [0, 0]
          }),
          interactive: false
        });
        circleLayer.addLayer(label);
      }
    });
  }

  // ---------- Mode 4: 計算能力・電力規模 ----------
  function renderPowerMode(items) {
    const withPower = items.filter(it => typeof it.powerMW === 'number' && it.powerMW > 0);
    const withoutPower = items.filter(it => it.hyperscale && !(typeof it.powerMW === 'number' && it.powerMW > 0));

    let maxPower = 1;
    withPower.forEach(it => { maxPower = Math.max(maxPower, it.powerMW); });

    const points = withPower.map(it => [it.lat, it.lng, Math.max(0.25, it.powerMW / maxPower)]);
    if (points.length) {
      heatLayer = L.heatLayer(points, {
        radius: 42, blur: 34, maxZoom: 8, max: 1.0,
        gradient: { 0.2: '#3b82f6', 0.45: '#60a5fa', 0.65: '#f59e0b', 0.85: '#f97316', 1.0: '#dc2626' }
      }).addTo(map);
    }

    withPower.forEach(it => {
      const marker = L.circleMarker([it.lat, it.lng], {
        radius: 14, fillOpacity: 0, opacity: 0, weight: 0
      });
      marker.on('click', () => window.AppController.selectItem(it));
      const qtag = it.powerMWQuality === 'confirmed' ? '確認値' : '推定値';
      marker.bindTooltip(`${it.nameJa || it.name}: ${fmtNum(it.powerMW, 'MW')} (${qtag})`, { direction: 'top' });
      markerLayer.addLayer(marker);
    });

    // 電力非公表のhyperscale拠点はグレーの小さいマーカーで存在のみ示す
    withoutPower.forEach(it => {
      const marker = L.circleMarker([it.lat, it.lng], {
        radius: 5, color: '#9aa5b1', fillColor: '#c3ccd6', fillOpacity: 0.8, weight: 1
      });
      marker.on('click', () => window.AppController.selectItem(it));
      marker.bindTooltip(`${it.nameJa || it.name}: 電力規模 非公表`, { direction: 'top' });
      markerLayer.addLayer(marker);
    });
  }

  // ---------- 拡大方向フロー ----------
  function renderFlows(flows, show) {
    flowLayer.clearLayers();
    if (!show) {
      if (map.hasLayer(flowLayer)) map.removeLayer(flowLayer);
      return;
    }
    flows.forEach(f => {
      const from = f.fromLatLng, to = f.toLatLng;
      const line = L.polyline([from, to], {
        color: '#dc2626', weight: 2.5, opacity: 0.75, dashArray: '1 8', lineCap: 'round'
      });
      line.on('click', () => window.AppController.selectFlow(f));
      line.bindTooltip(`${f.fromName} → ${f.toName}`, { sticky: true });
      flowLayer.addLayer(line);

      // クリックしやすいよう太い透明な当たり判定ラインを重ねる
      const hitLine = L.polyline([from, to], { color: '#dc2626', weight: 18, opacity: 0 });
      hitLine.on('click', () => window.AppController.selectFlow(f));
      hitLine.bindTooltip(`${f.fromName} → ${f.toName}`, { sticky: true });
      flowLayer.addLayer(hitLine);

      // 矢印head
      const angle = Math.atan2(to[0] - from[0], to[1] - from[1]);
      const arrowIcon = L.divIcon({
        className: '',
        html: `<div style="transform:rotate(${90 - angle * 180 / Math.PI}deg);font-size:16px;color:#dc2626;">➤</div>`,
        iconSize: [16, 16], iconAnchor: [8, 8]
      });
      const arrowMarker = L.marker(to, { icon: arrowIcon, interactive: false });
      flowLayer.addLayer(arrowMarker);
    });
    if (!map.hasLayer(flowLayer)) flowLayer.addTo(map);
  }

  function getMap() { return map; }
  function getZoom() { return currentZoom; }

  return {
    init, clearDataLayers, zoomToRegion,
    renderCountMode, renderHyperscaleMode, renderClusterMode, renderPowerMode,
    renderFlows, getMap, getZoom,
    STATUS_LABEL_JA, STATUS_SYMBOL
  };
})();
