// data.js — JSONデータの読み込みとメモリ上のデータストア
const DataStore = {
  countries: [],
  clusters: [],
  flows: [],
  lastUpdated: null,

  async loadAll() {
    if (window.location.protocol === 'file:') {
      throw new Error('FILE_PROTOCOL');
    }
    const [countriesRes, clustersRes, flowsRes] = await Promise.all([
      fetch('data/countries.json'),
      fetch('data/clusters.json'),
      fetch('data/expansion-flows.json')
    ]);

    if (!countriesRes.ok || !clustersRes.ok || !flowsRes.ok) {
      throw new Error('HTTP_ERROR: ' +
        [countriesRes.status, clustersRes.status, flowsRes.status].join(','));
    }

    const countriesJson = await countriesRes.json();
    const clustersJson = await clustersRes.json();
    const flowsJson = await flowsRes.json();

    this.countries = countriesJson.countries || [];
    this.clusters = clustersJson.items || [];
    this.flows = flowsJson.flows || [];

    // 最も新しい lastUpdated を採用
    const dates = [countriesJson.lastUpdated, clustersJson.lastUpdated, flowsJson.lastUpdated].filter(Boolean);
    this.lastUpdated = dates.sort().pop() || '不明';

    return true;
  }
};
