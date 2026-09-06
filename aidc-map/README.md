# 世界AIデータセンター・インフラマップ

世界のデータセンター／AI計算インフラの分布を、Google Mapsのようにドラッグ・ズームできる世界地図上でヒートマップ表示するインタラクティブWebツールです。API キー不要で動作します。

## 1. ツール概要

4つのモードをボタンで切り替えながら、世界地図上でデータセンター／AIインフラの分布を確認できます。

| モード | 表示内容 |
|---|---|
| ① データセンター総数 | 国別のDC総数をヒートマップ表示 |
| ② AI / Hyperscale | AWS・Microsoft・Google・Meta・OpenAI等の大型AI/Hyperscale拠点 |
| ③ 主要集積地域 | 地域単位のデータセンター・クラスター |
| ④ 計算能力・電力規模 | 電力容量(MW)を中心としたAI計算インフラの規模 |

「インフラ移動・拡大方向」スイッチをONにすると、北バージニア→オハイオ／北京・上海→内モンゴル・貴州など、世界的なAIインフラの拡大トレンドを矢印で表示します。

## 2. ファイル構成

```
/index.html              メイン画面
/css/style.css           スタイル
/js/data.js              JSONデータの読み込み
/js/map.js               Leaflet地図描画ロジック(各モード・ヒートマップ・凡例用データ)
/js/ui.js                サイドパネル・凡例・統計カード・ツールチップ描画
/js/app.js               状態管理・イベント配線(コントローラー)
/data/countries.json     国別データセンター概況
/data/clusters.json      AI/Hyperscale拠点・クラスターの詳細データ(施設/キャンパス/クラスター粒度混在)
/data/expansion-flows.json  インフラ移動・拡大トレンドのデータ
/data/data-centers.json  個別施設データの拡張用に予約(現状は clusters.json に統合)
/README.md               本ファイル
```

## 3. ローカルでの起動方法

`file://` で `index.html` を直接開くと、ブラウザのCORS制限によりJSONデータを読み込めません。必ずローカルサーバー経由で起動してください。

```bash
# このフォルダで実行
python -m http.server 8000
```

その後、ブラウザで `http://localhost:8000` を開いてください。

Node.js がある場合は以下でも起動できます。

```bash
npx serve .
```

## 4. GitHub Pages 等への公開方法

1. このフォルダの内容をGitHubリポジトリにpushします。
2. リポジトリの Settings → Pages で公開ブランチ・フォルダ(ルート)を指定します。
3. 数分後に `https://<username>.github.io/<repo>/` でアクセスできます。

単一の静的サイトなので、Netlify・Vercel・Cloudflare Pages 等、他の静的ホスティングサービスでもそのまま公開できます。

## 5. JSONデータの追加方法

`data/clusters.json` の `items` 配列に、以下の形式でオブジェクトを追加してください。

```json
{
  "id": "us-va-northern-virginia",
  "name": "英語名",
  "nameJa": "日本語名",
  "country": "国名(英語)",
  "countryCode": "ISO 2文字コード",
  "region": "North America | China | Europe | Japan | Asia Pacific",
  "lat": 39.0438,
  "lng": -77.4874,
  "granularity": "facility | campus | cluster",
  "hyperscale": true,
  "operators": ["事業者名の配列"],
  "dataCenterCount": null,
  "dataCenterCountQuality": "confirmed | estimated | unknown",
  "powerMW": null,
  "powerMWQuality": "confirmed | estimated | unknown",
  "gpuCount": null,
  "gpuType": null,
  "status": "operational | under_construction | planned",
  "startYear": null,
  "completionYear": null,
  "aiUse": "",
  "powerSituation": "",
  "locationFeatures": "",
  "notes": "",
  "source": "出典名",
  "sourceUrl": "出典URL",
  "sourceDate": "YYYY-MM"
}
```

**重要：数値を推測して埋めないでください。** 確認できない数値は `null` にし、`...Quality` フィールドを `unknown` としてください。確認済みの公式発表値がある場合のみ `confirmed`、報道等に基づく推計は `estimated` としてください。

国データを追加する場合は `data/countries.json` の `countries` 配列に、拡大トレンドを追加する場合は `data/expansion-flows.json` の `flows` 配列に、同様の考え方でオブジェクトを追加してください。

## 6. ヒートマップ閾値の変更方法

凡例の区分値は `js/ui.js` 内の `LEGEND_THRESHOLDS` オブジェクトで管理しています。

```js
const LEGEND_THRESHOLDS = {
  count: [
    { label: '0～20', color: '#3b82f6' },
    ...
  ],
  power: [
    { label: '0～20MW', color: '#3b82f6' },
    ...
  ]
};
```

ヒートマップ自体の色のグラデーションは `js/map.js` 内の `L.heatLayer(..., { gradient: {...} })` で調整できます。半径・ぼかしの強さは同箇所の `radius` / `blur` パラメータで変更できます。

## 7. 新しい事業者の追加方法

1. `index.html` の企業フィルター(`#operator-filter-group`)に `<button class="chip" data-op="事業者名">表示名</button>` を追加します。`data-op` の値は `data/clusters.json` の `operators` 配列内の文字列と完全一致させてください。
2. 各データレコードの `operators` 配列に事業者名を追加します。
3. 「その他」フィルターは `js/app.js` の `filteredClusters()` 内の `known` 配列に含まれない事業者を自動的に拾う仕組みになっているため、`known` 配列に新事業者を追加すると「その他」から独立したフィルターとして扱われます。

## 8. 出典の更新方法

各データレコードの `source` / `sourceUrl` / `sourceDate` フィールドを更新してください。`sourceUrl` を設定すると、サイドパネルに「出典を見る」リンクが自動的に表示されます。ファイル全体の `lastUpdated` フィールドも合わせて更新してください。

**優先すべき出典：** 各社公式発表・IR資料、各国政府・自治体資料、EU/米国/中国政府資料、日本経済産業省資料、Synergy Research Group、Stanford AI Index、Cloudscene、Data Center Map、信頼できる報道機関。Wikipediaのみを根拠とすることは避けてください。

## 9. データ品質について

本ツールは正確性を最優先しています。

- `confirmed`：企業公式発表・政府資料等で確認済みの数値
- `estimated`：報道・業界レポート等に基づく推計値
- `unknown` / `null`：非公表・未確認（推測値で埋めていません）

「データセンター数が多い＝AI計算能力が大きい、とは限らない」点にご注意ください。各モードの「？」アイコンから詳細説明を確認できます。

## 10. 技術構成

- HTML5 / CSS3 / Vanilla JavaScript(フレームワーク不使用)
- [Leaflet.js](https://leafletjs.com/)（地図描画、OSS・APIキー不要）
- [leaflet.heat](https://github.com/Leaflet/Leaflet.heat)（ヒートマップ表示）
- タイル：CARTO Positron（OpenStreetMapベースの明るいビジネス向けデザイン、APIキー不要）
- 外部ライブラリはCDN経由で読み込むため、初回アクセス時のみインターネット接続が必要です

## 11. 既知の制約

- 各種数値（DC数・電力容量・GPU台数）は公開情報の集計時点によって変動します。最新の一次情報は各出典リンクをご確認ください。
- 中国国内の一部データセンターは西側の集計トラッカーで捕捉しきれていない可能性があります。
- GPU台数・電力容量は多くの施設で非公表であり、「非公表」と表示される項目が多数存在します。これは意図的な仕様です。
