# GSC Local Keywords Analyzer

Google Search Console の検索パフォーマンスデータを **47都道府県別** に自動分析する Google Apps Script です。

地域キーワードの割合、都道府県ごとの平均順位、改善余地の大きいキーワードなどを、ボタン1つで可視化します。

## Features

- **地域キーワード自動分類** — 47都道府県名 + 主要都市200以上の辞書でクエリを自動判定
- **9種類の分析シート** を自動生成（詳細は後述）
- **日本語/英語 両対応** — GSCの言語設定に関わらず動作
- **1,000件以上対応** — 外部ツールで取得した大量データもOK
- **加重平均順位** — 表示回数を考慮した実態に近い平均順位を算出

## Output Sheets

| Sheet | Description |
|-------|-------------|
| **サマリー** | 地域KW比率・カバー率・全体 vs 地域のパフォーマンス比較 |
| **都道府県別分析** | 47都道府県ごとのKW数・クリック・表示・CTR・順位（単純/加重）・Top3/Top10率 |
| **地方別分析** | 7地方ブロック（北海道東北〜九州沖縄）単位の集計 |
| **地域KW一覧** | 地域と判定された全KWの一覧（マッチ根拠付き） |
| **チャンスKW** | 高表示・低順位のKW（チャンススコア付き） |
| **低CTR改善KW** | 順位帯別CTRベンチマークと比較し、CTRが低いKWを抽出 |
| **未カバー地域** | KWが0件/手薄/カバー済を色分け表示 |
| **順位帯別分布** | 都道府県 × 順位帯（1-3, 4-10, 11-20, 21-50, 51+）のクロス集計 |
| **地域×サービス** | 都道府県 × サービスKWのマトリクス（要・設定シート） |

## Setup

### 1. データの準備

1. [Google Search Console](https://search.google.com/search-console/) を開く
2. 「検索パフォーマンス」→ 期間を設定（推奨：過去3ヶ月）
3. 右上「エクスポート」→「Google スプレッドシート」

### 2. スクリプトの導入

1. スプレッドシートで「拡張機能」→「Apps Script」を開く
2. 既存コードを削除し、[`Code.gs`](./Code.gs) の内容を貼り付けて保存
3. スプレッドシートに戻り、**ページをリロード**

### 3. 実行

1. メニュー「**地域KW分析**」→「**分析を実行**」をクリック
2. 初回は認証画面が出るので「許可」を選択
3. 各シートが自動生成される

### 4. 地域×サービス分析（任意）

1. メニュー「地域KW分析」→「設定シートを作成」
2. サービスキーワードを1行1件で入力
3. 再度「分析を実行」

## Supported Headers

日本語・英語どちらのヘッダーにも対応しています。

| Japanese | English |
|----------|---------|
| 上位のクエリ | Top queries / Queries / Query |
| クリック数 | Clicks |
| 表示回数 | Impressions |
| CTR | CTR |
| 掲載順位 | Position / Average position |

Sheet name: `クエリ` or `Queries`

## Customization

### 都市辞書の編集

`Code.gs` 内の `PREFECTURE_DICT` を編集することで、マッチング対象の都市名を追加・削除できます。

```javascript
'東京都': {
  aliases: ['東京'],
  cities: ['新宿','渋谷','池袋','品川','銀座','八王子','町田','立川',...]
},
```

### チャンスKWの抽出条件

`writeOpportunitySheet_()` 内の条件を変更できます。

```javascript
// デフォルト: 順位11位以下 & 表示回数10以上
if (q.position > 10 && q.impressions >= 10) {
```

### CTRベンチマーク

`writeLowCTRSheet_()` 内の `ctrBenchmark` 関数で、順位帯ごとの期待CTRを調整できます。

## Documentation

各分析シートの詳細な説明は [GUIDE.md](./GUIDE.md) を参照してください。

## File Structure

```
gsc-local-keywords/
├── Code.gs          # Apps Script 本体
├── GUIDE.md         # 各分析項目の詳細説明書
├── README.md        # このファイル
├── CHANGELOG.md     # 変更履歴
├── LICENSE          # MIT License
└── .gitignore
```

## License

[MIT](./LICENSE)
