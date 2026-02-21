/**
 * GSC 地域キーワード分析スクリプト
 *
 * Google Search Console の検索パフォーマンスデータを
 * 都道府県別に分析し、地域SEOの改善ポイントを可視化します。
 *
 * 使い方:
 * 1. GSCからクエリデータをスプレッドシートにエクスポート
 * 2. このスクリプトをスプレッドシートに追加
 * 3. メニュー「地域KW分析」→「分析を実行」をクリック
 */

// ============================================================
// ヘッダーマッピング（日本語/英語両対応）
// ============================================================
var HEADER_CANDIDATES = {
  query:       ['上位のクエリ', 'Top queries', 'Queries', 'Query', 'クエリ'],
  clicks:      ['クリック数', 'Clicks'],
  impressions: ['表示回数', 'Impressions'],
  ctr:         ['CTR'],
  position:    ['掲載順位', 'Position', 'Average position']
};

var QUERY_SHEET_NAMES = ['クエリ', 'Queries'];

// ============================================================
// 47都道府県 + 主要都市辞書
// ============================================================
var PREFECTURE_DICT = {
  '北海道':   { aliases: ['北海道'], cities: ['札幌','旭川','函館','帯広','釧路','苫小牧','小樽','北見','江別','千歳'] },
  '青森県':   { aliases: ['青森'], cities: ['八戸','弘前','十和田','むつ'] },
  '岩手県':   { aliases: ['岩手'], cities: ['盛岡','一関','花巻','奥州','北上'] },
  '宮城県':   { aliases: ['宮城'], cities: ['仙台','石巻','大崎','名取'] },
  '秋田県':   { aliases: ['秋田'], cities: ['横手','大仙','由利本荘'] },
  '山形県':   { aliases: ['山形'], cities: ['酒田','鶴岡','米沢','天童'] },
  '福島県':   { aliases: ['福島'], cities: ['郡山','いわき','会津若松','須賀川'] },
  '茨城県':   { aliases: ['茨城'], cities: ['水戸','つくば','日立','土浦','古河','取手'] },
  '栃木県':   { aliases: ['栃木'], cities: ['宇都宮','小山','足利','那須塩原','佐野'] },
  '群馬県':   { aliases: ['群馬'], cities: ['前橋','高崎','太田','伊勢崎','桐生'] },
  '埼玉県':   { aliases: ['埼玉'], cities: ['さいたま','川口','川越','所沢','越谷','大宮','春日部','熊谷','草加','上尾'] },
  '千葉県':   { aliases: ['千葉'], cities: ['船橋','松戸','市川','浦安','木更津','成田','佐倉'] },
  '東京都':   { aliases: ['東京'], cities: ['新宿','渋谷','池袋','品川','銀座','八王子','町田','立川','上野','秋葉原','六本木','丸の内','多摩'] },
  '神奈川県': { aliases: ['神奈川'], cities: ['横浜','川崎','相模原','藤沢','横須賀','厚木','小田原','鎌倉','茅ヶ崎'] },
  '新潟県':   { aliases: ['新潟'], cities: ['長岡','上越','三条','燕','柏崎'] },
  '富山県':   { aliases: ['富山'], cities: ['高岡','射水','氷見'] },
  '石川県':   { aliases: ['石川'], cities: ['金沢','小松','白山','加賀'] },
  '福井県':   { aliases: ['福井'], cities: ['敦賀','鯖江','越前'] },
  '山梨県':   { aliases: ['山梨'], cities: ['甲府','富士吉田','甲斐','南アルプス'] },
  '長野県':   { aliases: ['長野'], cities: ['松本','上田','飯田','諏訪','佐久','安曇野'] },
  '岐阜県':   { aliases: ['岐阜'], cities: ['大垣','各務原','多治見','高山','関'] },
  '静岡県':   { aliases: ['静岡'], cities: ['浜松','沼津','磐田','藤枝','焼津','富士宮','掛川','三島'] },
  '愛知県':   { aliases: ['愛知'], cities: ['名古屋','豊田','豊橋','岡崎','一宮','春日井','刈谷','安城','豊川'] },
  '三重県':   { aliases: ['三重'], cities: ['四日市','鈴鹿','松阪','伊勢','桑名'] },
  '滋賀県':   { aliases: ['滋賀'], cities: ['大津','草津','彦根','近江八幡','長浜'] },
  '京都府':   { aliases: ['京都'], cities: ['宇治','舞鶴','亀岡','福知山'] },
  '大阪府':   { aliases: ['大阪'], cities: ['堺','東大阪','枚方','豊中','吹田','高槻','八尾','茨木','寝屋川','岸和田'] },
  '兵庫県':   { aliases: ['兵庫'], cities: ['神戸','姫路','西宮','尼崎','明石','加古川','宝塚','伊丹','川西'] },
  '奈良県':   { aliases: ['奈良'], cities: ['橿原','生駒','天理','大和郡山'] },
  '和歌山県': { aliases: ['和歌山'], cities: ['田辺','海南','橋本','紀の川'] },
  '鳥取県':   { aliases: ['鳥取'], cities: ['米子','倉吉','境港'] },
  '島根県':   { aliases: ['島根'], cities: ['松江','出雲','浜田','益田'] },
  '岡山県':   { aliases: ['岡山'], cities: ['倉敷','津山','総社'] },
  '広島県':   { aliases: ['広島'], cities: ['福山','呉','東広島','尾道','廿日市'] },
  '山口県':   { aliases: ['山口'], cities: ['下関','周南','宇部','岩国','防府'] },
  '徳島県':   { aliases: ['徳島'], cities: ['鳴門','阿南','小松島'] },
  '香川県':   { aliases: ['香川'], cities: ['高松','丸亀','三豊','坂出'] },
  '愛媛県':   { aliases: ['愛媛'], cities: ['松山','今治','新居浜','西条','四国中央'] },
  '高知県':   { aliases: ['高知'], cities: ['南国','四万十','香南'] },
  '福岡県':   { aliases: ['福岡'], cities: ['北九州','久留米','博多','小倉','飯塚','大牟田','春日'] },
  '佐賀県':   { aliases: ['佐賀'], cities: ['唐津','鳥栖','伊万里'] },
  '長崎県':   { aliases: ['長崎'], cities: ['佐世保','諫早','大村'] },
  '熊本県':   { aliases: ['熊本'], cities: ['八代','天草','玉名','菊池'] },
  '大分県':   { aliases: ['大分'], cities: ['別府','中津','日田','佐伯'] },
  '宮崎県':   { aliases: ['宮崎'], cities: ['都城','延岡','日南','日向'] },
  '鹿児島県': { aliases: ['鹿児島'], cities: ['霧島','鹿屋','薩摩川内','指宿'] },
  '沖縄県':   { aliases: ['沖縄'], cities: ['那覇','宜野湾','浦添','名護','糸満','豊見城'] }
};

// 地方区分
var REGION_MAP = {
  '北海道・東北': ['北海道','青森県','岩手県','宮城県','秋田県','山形県','福島県'],
  '関東':         ['茨城県','栃木県','群馬県','埼玉県','千葉県','東京都','神奈川県'],
  '中部':         ['新潟県','富山県','石川県','福井県','山梨県','長野県','岐阜県','静岡県','愛知県'],
  '近畿':         ['三重県','滋賀県','京都府','大阪府','兵庫県','奈良県','和歌山県'],
  '中国':         ['鳥取県','島根県','岡山県','広島県','山口県'],
  '四国':         ['徳島県','香川県','愛媛県','高知県'],
  '九州・沖縄':   ['福岡県','佐賀県','長崎県','熊本県','大分県','宮崎県','鹿児島県','沖縄県']
};


// ============================================================
// メニュー登録
// ============================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('地域KW分析')
    .addItem('分析を実行', 'runFullAnalysis')
    .addSeparator()
    .addItem('設定シートを作成', 'createSettingsSheet')
    .addToUi();
}


// ============================================================
// メインエントリポイント
// ============================================================
function runFullAnalysis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 1. クエリデータの読み込み
  var queryData = readQueryData_(ss);
  if (!queryData) {
    ui.alert('エラー', 'クエリシートが見つかりません。\n「クエリ」または「Queries」という名前のシートが必要です。', ui.ButtonSet.OK);
    return;
  }
  if (queryData.rows.length === 0) {
    ui.alert('エラー', 'クエリデータが空です。', ui.ButtonSet.OK);
    return;
  }

  // 2. 都道府県マッチング用インデックスの構築
  var prefIndex = buildPrefectureIndex_();

  // 3. 全クエリを都道府県に分類
  var classified = classifyAllQueries_(queryData.rows, prefIndex);

  // 4. サービスキーワードの読み込み（設定シートがあれば）
  var serviceKeywords = readServiceKeywords_(ss);

  // 5. 各種分析の実行と出力
  writeSummarySheet_(ss, queryData.rows, classified);
  writePrefectureSheet_(ss, classified);
  writeRegionSheet_(ss, classified);
  writeLocalKeywordListSheet_(ss, classified);
  writeOpportunitySheet_(ss, classified);
  writeLowCTRSheet_(ss, queryData.rows, classified);
  writeUncoveredSheet_(ss, classified);
  writePositionDistributionSheet_(ss, classified);
  if (serviceKeywords.length > 0) {
    writeServiceMatrixSheet_(ss, classified, serviceKeywords);
  }

  // 6. 完了通知
  ui.alert('完了', '地域キーワード分析が完了しました。\n出力シートを確認してください。', ui.ButtonSet.OK);
}


// ============================================================
// データ読み込み
// ============================================================

/**
 * クエリシートからデータを読み込む
 * @return {Object|null} { rows: [{query, clicks, impressions, ctr, position}] }
 */
function readQueryData_(ss) {
  var sheet = null;
  for (var i = 0; i < QUERY_SHEET_NAMES.length; i++) {
    sheet = ss.getSheetByName(QUERY_SHEET_NAMES[i]);
    if (sheet) break;
  }
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { rows: [] };

  // ヘッダー行の解析
  var headers = data[0];
  var colMap = parseHeaders_(headers);
  if (colMap.query < 0) return null;

  // データ行の解析
  var rows = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var query = String(row[colMap.query] || '').trim();
    if (!query) continue;

    var clicks = parseNumber_(row[colMap.clicks]);
    var impressions = parseNumber_(row[colMap.impressions]);
    var ctr = parseCTR_(row[colMap.ctr]);
    var position = parseNumber_(row[colMap.position]);

    rows.push({
      query: query,
      clicks: clicks,
      impressions: impressions,
      ctr: ctr,
      position: position
    });
  }

  return { rows: rows };
}

/**
 * ヘッダー行を解析し、各フィールドの列インデックスを返す
 */
function parseHeaders_(headerRow) {
  var colMap = { query: -1, clicks: -1, impressions: -1, ctr: -1, position: -1 };
  var fields = Object.keys(colMap);

  for (var c = 0; c < headerRow.length; c++) {
    var h = String(headerRow[c]).trim();
    for (var f = 0; f < fields.length; f++) {
      var field = fields[f];
      var candidates = HEADER_CANDIDATES[field];
      for (var k = 0; k < candidates.length; k++) {
        if (h === candidates[k]) {
          colMap[field] = c;
          break;
        }
      }
    }
  }
  return colMap;
}

/**
 * 数値をパースする（文字列・数値両対応）
 */
function parseNumber_(val) {
  if (typeof val === 'number') return val;
  if (typeof val === 'string') {
    val = val.replace(/,/g, '').replace(/%/g, '').trim();
    var n = parseFloat(val);
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

/**
 * CTRをパースする（"10.56%" → 0.1056, 0.1056 → 0.1056）
 */
function parseCTR_(val) {
  if (typeof val === 'number') {
    // 既に小数（0.1056）の場合はそのまま、100以下の%値の場合は変換
    return val > 1 ? val / 100 : val;
  }
  if (typeof val === 'string') {
    val = val.replace(/,/g, '').replace(/%/g, '').trim();
    var n = parseFloat(val);
    if (isNaN(n)) return 0;
    return n > 1 ? n / 100 : n;
  }
  return 0;
}

/**
 * 設定シートからサービスキーワードを読み込む
 */
function readServiceKeywords_(ss) {
  var sheet = ss.getSheetByName('設定');
  if (!sheet) sheet = ss.getSheetByName('Settings');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var keywords = [];
  var inSection = false;

  for (var r = 0; r < data.length; r++) {
    var cell = String(data[r][0]).trim();
    if (cell === 'サービスキーワード' || cell === 'Service Keywords') {
      inSection = true;
      continue;
    }
    if (inSection) {
      if (cell.indexOf('---') === 0) break;
      if (cell === '' || cell.charAt(0) === '(' || cell.charAt(0) === '（') continue;
      keywords.push(cell);
    }
  }
  return keywords;
}


// ============================================================
// 都道府県マッチング
// ============================================================

/**
 * マッチング用インデックスを構築
 * @return {Array} [{ keyword, prefecture }] 長い文字列を先にマッチさせるためソート済み
 */
function buildPrefectureIndex_() {
  var entries = [];
  var prefNames = Object.keys(PREFECTURE_DICT);

  for (var i = 0; i < prefNames.length; i++) {
    var pref = prefNames[i];
    var info = PREFECTURE_DICT[pref];

    // 都道府県フルネーム（例: 神奈川県）
    entries.push({ keyword: pref, prefecture: pref });

    // エイリアス（例: 神奈川、東京）
    for (var a = 0; a < info.aliases.length; a++) {
      if (info.aliases[a] !== pref) {
        entries.push({ keyword: info.aliases[a], prefecture: pref });
      }
    }

    // 主要都市名
    for (var c = 0; c < info.cities.length; c++) {
      entries.push({ keyword: info.cities[c], prefecture: pref });
    }
  }

  // 長い文字列を先にマッチ（「会津若松」が「若松」より先にマッチするように）
  entries.sort(function(a, b) { return b.keyword.length - a.keyword.length; });
  return entries;
}

/**
 * 1つのクエリがどの都道府県に該当するか判定
 * @return {Array} マッチした都道府県名の配列（重複なし）
 */
function classifyQuery_(query, prefIndex) {
  var matched = {};
  var matchDetails = [];
  var lowerQuery = query.toLowerCase();

  for (var i = 0; i < prefIndex.length; i++) {
    var entry = prefIndex[i];
    if (lowerQuery.indexOf(entry.keyword.toLowerCase()) >= 0) {
      if (!matched[entry.prefecture]) {
        matched[entry.prefecture] = true;
        matchDetails.push({
          prefecture: entry.prefecture,
          matchedKeyword: entry.keyword
        });
      }
    }
  }
  return matchDetails;
}

/**
 * 全クエリを分類
 * @return {Object} {
 *   localQueries: [{ query, clicks, impressions, ctr, position, prefectures: [{prefecture, matchedKeyword}] }],
 *   nonLocalQueries: [{ query, clicks, impressions, ctr, position }],
 *   byPrefecture: { '東京都': [{ query, ... }], ... }
 * }
 */
function classifyAllQueries_(rows, prefIndex) {
  var localQueries = [];
  var nonLocalQueries = [];
  var byPrefecture = {};

  // 初期化
  var prefNames = Object.keys(PREFECTURE_DICT);
  for (var p = 0; p < prefNames.length; p++) {
    byPrefecture[prefNames[p]] = [];
  }

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var matches = classifyQuery_(row.query, prefIndex);

    if (matches.length > 0) {
      var localRow = {
        query: row.query,
        clicks: row.clicks,
        impressions: row.impressions,
        ctr: row.ctr,
        position: row.position,
        prefectures: matches
      };
      localQueries.push(localRow);

      for (var m = 0; m < matches.length; m++) {
        byPrefecture[matches[m].prefecture].push(localRow);
      }
    } else {
      nonLocalQueries.push(row);
    }
  }

  return {
    localQueries: localQueries,
    nonLocalQueries: nonLocalQueries,
    byPrefecture: byPrefecture
  };
}


// ============================================================
// 出力: サマリー
// ============================================================
function writeSummarySheet_(ss, allRows, classified) {
  var sheet = getOrCreateSheet_(ss, 'サマリー');
  sheet.clear();

  var totalKW = allRows.length;
  var localKW = classified.localQueries.length;
  var nonLocalKW = classified.nonLocalQueries.length;
  var localRatio = totalKW > 0 ? localKW / totalKW : 0;

  // 全体指標
  var totalClicks = 0, totalImpressions = 0, totalPositionSum = 0;
  var localClicks = 0, localImpressions = 0, localPositionSum = 0, localImpressionWeightedPos = 0;
  var nonLocalClicks = 0, nonLocalImpressions = 0;

  for (var i = 0; i < allRows.length; i++) {
    totalClicks += allRows[i].clicks;
    totalImpressions += allRows[i].impressions;
    totalPositionSum += allRows[i].position;
  }
  for (var i = 0; i < classified.localQueries.length; i++) {
    var q = classified.localQueries[i];
    localClicks += q.clicks;
    localImpressions += q.impressions;
    localPositionSum += q.position;
    localImpressionWeightedPos += q.position * q.impressions;
  }
  for (var i = 0; i < classified.nonLocalQueries.length; i++) {
    nonLocalClicks += classified.nonLocalQueries[i].clicks;
    nonLocalImpressions += classified.nonLocalQueries[i].impressions;
  }

  var avgPosAll = totalKW > 0 ? totalPositionSum / totalKW : 0;
  var avgPosLocal = localKW > 0 ? localPositionSum / localKW : 0;
  var weightedAvgPosLocal = localImpressions > 0 ? localImpressionWeightedPos / localImpressions : 0;

  // カバー都道府県数
  var coveredCount = 0;
  var prefNames = Object.keys(PREFECTURE_DICT);
  for (var p = 0; p < prefNames.length; p++) {
    if (classified.byPrefecture[prefNames[p]].length > 0) coveredCount++;
  }

  var output = [
    ['GSC 地域キーワード分析レポート', '', ''],
    ['', '', ''],
    ['■ 全体概要', '', ''],
    ['総キーワード数', totalKW, ''],
    ['地域キーワード数', localKW, ''],
    ['非地域キーワード数', nonLocalKW, ''],
    ['地域キーワード比率', localRatio, ''],
    ['', '', ''],
    ['■ パフォーマンス比較', '全体', '地域KWのみ'],
    ['クリック数合計', totalClicks, localClicks],
    ['表示回数合計', totalImpressions, localImpressions],
    ['平均掲載順位（単純平均）', roundTo_(avgPosAll, 2), roundTo_(avgPosLocal, 2)],
    ['平均掲載順位（表示回数加重）', '', roundTo_(weightedAvgPosLocal, 2)],
    ['', '', ''],
    ['■ 都道府県カバレッジ', '', ''],
    ['カバー済み都道府県数', coveredCount, '/ 47'],
    ['未カバー都道府県数', 47 - coveredCount, ''],
    ['カバー率', coveredCount / 47, '']
  ];

  sheet.getRange(1, 1, output.length, 3).setValues(output);

  // 書式設定
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A3').setFontWeight('bold');
  sheet.getRange('A9').setFontWeight('bold');
  sheet.getRange('A15').setFontWeight('bold');
  sheet.getRange('B7').setNumberFormat('0.0%');
  sheet.getRange('B18').setNumberFormat('0.0%');
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
}


// ============================================================
// 出力: 都道府県別分析
// ============================================================
function writePrefectureSheet_(ss, classified) {
  var sheet = getOrCreateSheet_(ss, '都道府県別分析');
  sheet.clear();

  var headers = [
    '都道府県', 'キーワード数', 'クリック数合計', '表示回数合計',
    '平均CTR', '平均順位（単純）', '平均順位（加重）',
    'Top3内KW数', 'Top10内KW数', '順位10以下KW数'
  ];

  var rows = [];
  var prefNames = Object.keys(PREFECTURE_DICT);

  for (var p = 0; p < prefNames.length; p++) {
    var pref = prefNames[p];
    var queries = classified.byPrefecture[pref];
    if (queries.length === 0) {
      rows.push([pref, 0, 0, 0, 0, '-', '-', 0, 0, 0]);
      continue;
    }

    var clicks = 0, impressions = 0, posSum = 0, weightedPosSum = 0;
    var top3 = 0, top10 = 0, below10 = 0;

    for (var i = 0; i < queries.length; i++) {
      var q = queries[i];
      clicks += q.clicks;
      impressions += q.impressions;
      posSum += q.position;
      weightedPosSum += q.position * q.impressions;
      if (q.position <= 3) top3++;
      if (q.position <= 10) top10++;
      if (q.position > 10) below10++;
    }

    var avgCTR = impressions > 0 ? clicks / impressions : 0;
    var avgPos = roundTo_(posSum / queries.length, 2);
    var weightedAvgPos = impressions > 0 ? roundTo_(weightedPosSum / impressions, 2) : 0;

    rows.push([pref, queries.length, clicks, impressions, avgCTR, avgPos, weightedAvgPos, top3, top10, below10]);
  }

  // キーワード数の多い順にソート
  rows.sort(function(a, b) { return b[1] - a[1]; });

  var output = [headers].concat(rows);
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  // 書式設定
  formatHeaderRow_(sheet, headers.length);
  sheet.getRange(2, 5, rows.length, 1).setNumberFormat('0.00%');
  autoResizeColumns_(sheet, headers.length);
}


// ============================================================
// 出力: 地方別分析
// ============================================================
function writeRegionSheet_(ss, classified) {
  var sheet = getOrCreateSheet_(ss, '地方別分析');
  sheet.clear();

  var headers = ['地方', 'キーワード数', 'クリック数合計', '表示回数合計', '平均CTR', '平均順位（加重）', '含まれる都道府県'];

  var rows = [];
  var regionNames = Object.keys(REGION_MAP);

  for (var r = 0; r < regionNames.length; r++) {
    var region = regionNames[r];
    var prefs = REGION_MAP[region];
    var kwSet = {};
    var clicks = 0, impressions = 0, weightedPosSum = 0;
    var activePrefList = [];

    for (var p = 0; p < prefs.length; p++) {
      var queries = classified.byPrefecture[prefs[p]];
      if (queries.length > 0) activePrefList.push(prefs[p]);
      for (var i = 0; i < queries.length; i++) {
        var q = queries[i];
        if (!kwSet[q.query]) {
          kwSet[q.query] = true;
          clicks += q.clicks;
          impressions += q.impressions;
          weightedPosSum += q.position * q.impressions;
        }
      }
    }

    var kwCount = Object.keys(kwSet).length;
    var avgCTR = impressions > 0 ? clicks / impressions : 0;
    var weightedAvgPos = impressions > 0 ? roundTo_(weightedPosSum / impressions, 2) : '-';

    rows.push([region, kwCount, clicks, impressions, avgCTR, weightedAvgPos, activePrefList.join(', ')]);
  }

  var output = [headers].concat(rows);
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  formatHeaderRow_(sheet, headers.length);
  sheet.getRange(2, 5, rows.length, 1).setNumberFormat('0.00%');
  autoResizeColumns_(sheet, headers.length);
}


// ============================================================
// 出力: 地域KW一覧
// ============================================================
function writeLocalKeywordListSheet_(ss, classified) {
  var sheet = getOrCreateSheet_(ss, '地域KW一覧');
  sheet.clear();

  var headers = ['クエリ', 'クリック数', '表示回数', 'CTR', '掲載順位', '該当都道府県', 'マッチキーワード'];

  var rows = [];
  for (var i = 0; i < classified.localQueries.length; i++) {
    var q = classified.localQueries[i];
    var prefs = [];
    var matchedKws = [];
    for (var m = 0; m < q.prefectures.length; m++) {
      prefs.push(q.prefectures[m].prefecture);
      matchedKws.push(q.prefectures[m].matchedKeyword);
    }
    rows.push([q.query, q.clicks, q.impressions, q.ctr, roundTo_(q.position, 2), prefs.join(', '), matchedKws.join(', ')]);
  }

  // クリック数の多い順にソート
  rows.sort(function(a, b) { return b[1] - a[1]; });

  var output = [headers].concat(rows);
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  formatHeaderRow_(sheet, headers.length);
  if (rows.length > 0) {
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('0.00%');
  }
  autoResizeColumns_(sheet, headers.length);
}


// ============================================================
// 出力: チャンスKW（高表示・低順位）
// ============================================================
function writeOpportunitySheet_(ss, classified) {
  var sheet = getOrCreateSheet_(ss, 'チャンスKW');
  sheet.clear();

  var headers = ['クエリ', 'クリック数', '表示回数', 'CTR', '掲載順位', '都道府県', 'チャンススコア'];
  var rows = [];

  // 地域KWの中から高表示回数＆低順位のものを抽出
  for (var i = 0; i < classified.localQueries.length; i++) {
    var q = classified.localQueries[i];
    // 順位が10位以下 かつ 表示回数が中央値以上のものを抽出
    if (q.position > 10 && q.impressions >= 10) {
      // チャンススコア: 表示回数 × (順位 - 目標順位) で改善インパクトを推定
      var score = q.impressions * (q.position - 3);
      var prefs = [];
      for (var m = 0; m < q.prefectures.length; m++) {
        prefs.push(q.prefectures[m].prefecture);
      }
      rows.push([q.query, q.clicks, q.impressions, q.ctr, roundTo_(q.position, 2), prefs.join(', '), roundTo_(score, 0)]);
    }
  }

  // スコアの高い順にソート
  rows.sort(function(a, b) { return b[6] - a[6]; });

  if (rows.length === 0) {
    sheet.getRange(1, 1).setValue('該当するチャンスキーワードはありませんでした（順位10位以下 & 表示回数10以上の地域KW）');
    return;
  }

  var output = [headers].concat(rows);
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  formatHeaderRow_(sheet, headers.length);
  sheet.getRange(2, 4, rows.length, 1).setNumberFormat('0.00%');
  autoResizeColumns_(sheet, headers.length);
}


// ============================================================
// 出力: 高順位・低CTRキーワード
// ============================================================
function writeLowCTRSheet_(ss, allRows, classified) {
  var sheet = getOrCreateSheet_(ss, '低CTR改善KW');
  sheet.clear();

  var headers = ['クエリ', 'クリック数', '表示回数', 'CTR', '掲載順位', '地域KW', '都道府県', '想定CTRとの差'];

  // 順位帯別の平均的なCTRベンチマーク
  var ctrBenchmark = function(pos) {
    if (pos <= 1) return 0.30;
    if (pos <= 2) return 0.15;
    if (pos <= 3) return 0.10;
    if (pos <= 5) return 0.05;
    if (pos <= 10) return 0.02;
    return 0.01;
  };

  // 地域KWのクエリセットを作成
  var localQuerySet = {};
  for (var i = 0; i < classified.localQueries.length; i++) {
    var q = classified.localQueries[i];
    var prefs = [];
    for (var m = 0; m < q.prefectures.length; m++) {
      prefs.push(q.prefectures[m].prefecture);
    }
    localQuerySet[q.query] = prefs.join(', ');
  }

  var rows = [];
  for (var i = 0; i < allRows.length; i++) {
    var row = allRows[i];
    if (row.position > 10 || row.impressions < 10) continue;

    var expectedCTR = ctrBenchmark(row.position);
    var ctrGap = row.ctr - expectedCTR;

    if (ctrGap < -0.02) { // CTRが期待値より2%以上低い
      var isLocal = localQuerySet[row.query] !== undefined;
      rows.push([
        row.query, row.clicks, row.impressions, row.ctr,
        roundTo_(row.position, 2),
        isLocal ? 'はい' : 'いいえ',
        isLocal ? localQuerySet[row.query] : '',
        roundTo_(ctrGap, 4)
      ]);
    }
  }

  // CTR差の大きい順（悪い順）にソート
  rows.sort(function(a, b) { return a[7] - b[7]; });

  if (rows.length === 0) {
    sheet.getRange(1, 1).setValue('CTRが期待値を大きく下回るキーワードはありませんでした');
    return;
  }

  var output = [headers].concat(rows);
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  formatHeaderRow_(sheet, headers.length);
  sheet.getRange(2, 4, rows.length, 1).setNumberFormat('0.00%');
  sheet.getRange(2, 8, rows.length, 1).setNumberFormat('+0.00%;-0.00%');
  autoResizeColumns_(sheet, headers.length);
}


// ============================================================
// 出力: 未カバー地域
// ============================================================
function writeUncoveredSheet_(ss, classified) {
  var sheet = getOrCreateSheet_(ss, '未カバー地域');
  sheet.clear();

  var headers = ['都道府県', '地方', 'ステータス', 'キーワード数'];
  var rows = [];
  var prefNames = Object.keys(PREFECTURE_DICT);

  // 地方の逆引きマップ
  var prefToRegion = {};
  var regionNames = Object.keys(REGION_MAP);
  for (var r = 0; r < regionNames.length; r++) {
    var prefs = REGION_MAP[regionNames[r]];
    for (var p = 0; p < prefs.length; p++) {
      prefToRegion[prefs[p]] = regionNames[r];
    }
  }

  for (var p = 0; p < prefNames.length; p++) {
    var pref = prefNames[p];
    var count = classified.byPrefecture[pref].length;
    var status = count === 0 ? '未カバー' : (count <= 2 ? '手薄' : 'カバー済');
    rows.push([pref, prefToRegion[pref] || '', status, count]);
  }

  // 未カバー → 手薄 → カバー済の順
  var statusOrder = { '未カバー': 0, '手薄': 1, 'カバー済': 2 };
  rows.sort(function(a, b) {
    var diff = statusOrder[a[2]] - statusOrder[b[2]];
    if (diff !== 0) return diff;
    return a[3] - b[3];
  });

  var output = [headers].concat(rows);
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  formatHeaderRow_(sheet, headers.length);

  // 条件付き書式: 未カバーを赤、手薄をオレンジ
  for (var i = 0; i < rows.length; i++) {
    var rowNum = i + 2;
    if (rows[i][2] === '未カバー') {
      sheet.getRange(rowNum, 1, 1, headers.length).setBackground('#fce4ec');
    } else if (rows[i][2] === '手薄') {
      sheet.getRange(rowNum, 1, 1, headers.length).setBackground('#fff3e0');
    }
  }
  autoResizeColumns_(sheet, headers.length);
}


// ============================================================
// 出力: 順位帯別分布
// ============================================================
function writePositionDistributionSheet_(ss, classified) {
  var sheet = getOrCreateSheet_(ss, '順位帯別分布');
  sheet.clear();

  var positionBands = [
    { label: '1〜3位', min: 0, max: 3 },
    { label: '4〜10位', min: 3, max: 10 },
    { label: '11〜20位', min: 10, max: 20 },
    { label: '21〜50位', min: 20, max: 50 },
    { label: '51位以下', min: 50, max: 9999 }
  ];

  var headers = ['都道府県'];
  for (var b = 0; b < positionBands.length; b++) {
    headers.push(positionBands[b].label);
  }
  headers.push('合計');

  var rows = [];
  var prefNames = Object.keys(PREFECTURE_DICT);

  for (var p = 0; p < prefNames.length; p++) {
    var pref = prefNames[p];
    var queries = classified.byPrefecture[pref];
    if (queries.length === 0) continue;

    var row = [pref];
    var total = 0;
    for (var b = 0; b < positionBands.length; b++) {
      var count = 0;
      for (var i = 0; i < queries.length; i++) {
        if (queries[i].position > positionBands[b].min && queries[i].position <= positionBands[b].max) {
          count++;
        }
      }
      row.push(count);
      total += count;
    }
    row.push(total);
    rows.push(row);
  }

  // 合計の多い順
  rows.sort(function(a, b) { return b[b.length - 1] - a[a.length - 1]; });

  if (rows.length === 0) {
    sheet.getRange(1, 1).setValue('地域キーワードが見つかりませんでした');
    return;
  }

  var output = [headers].concat(rows);
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  formatHeaderRow_(sheet, headers.length);
  autoResizeColumns_(sheet, headers.length);
}


// ============================================================
// 出力: 地域 × サービスマトリクス
// ============================================================
function writeServiceMatrixSheet_(ss, classified, serviceKeywords) {
  var sheet = getOrCreateSheet_(ss, '地域×サービス');
  sheet.clear();

  var headers = ['都道府県'].concat(serviceKeywords).concat(['その他地域KW']);
  var rows = [];
  var prefNames = Object.keys(PREFECTURE_DICT);

  for (var p = 0; p < prefNames.length; p++) {
    var pref = prefNames[p];
    var queries = classified.byPrefecture[pref];
    if (queries.length === 0) continue;

    var row = [pref];
    var classifiedCount = 0;

    for (var s = 0; s < serviceKeywords.length; s++) {
      var svcKw = serviceKeywords[s].toLowerCase();
      var count = 0;
      for (var i = 0; i < queries.length; i++) {
        if (queries[i].query.toLowerCase().indexOf(svcKw) >= 0) {
          count++;
        }
      }
      row.push(count);
      classifiedCount += count;
    }

    // その他（どのサービスKWにもマッチしなかったもの）
    // 注意: 1つのクエリが複数サービスにマッチする可能性があるため厳密ではないが参考値
    row.push(Math.max(0, queries.length - classifiedCount));
    rows.push(row);
  }

  rows.sort(function(a, b) {
    var sumA = 0, sumB = 0;
    for (var i = 1; i < a.length; i++) sumA += a[i];
    for (var i = 1; i < b.length; i++) sumB += b[i];
    return sumB - sumA;
  });

  if (rows.length === 0) {
    sheet.getRange(1, 1).setValue('地域キーワードが見つかりませんでした');
    return;
  }

  var output = [headers].concat(rows);
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  formatHeaderRow_(sheet, headers.length);

  // 0のセルをグレー、値のあるセルを色付け
  for (var r = 0; r < rows.length; r++) {
    for (var c = 1; c < headers.length; c++) {
      var val = rows[r][c];
      var cell = sheet.getRange(r + 2, c + 1);
      if (val === 0) {
        cell.setBackground('#f5f5f5').setFontColor('#bdbdbd');
      } else if (val >= 3) {
        cell.setBackground('#c8e6c9');
      } else if (val >= 1) {
        cell.setBackground('#e8f5e9');
      }
    }
  }
  autoResizeColumns_(sheet, headers.length);
}


// ============================================================
// 設定シート作成
// ============================================================
function createSettingsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet_(ss, '設定');
  sheet.clear();

  var content = [
    ['地域KW分析 設定'],
    [''],
    ['サービスキーワード'],
    ['（以下に1行ずつサービスに関するキーワードを入力してください）'],
    ['ウォータージェット'],
    ['加工'],
    ['切断'],
    ['---'],
    [''],
    ['※ 上記のキーワードは「地域×サービス」マトリクス分析で使用されます。'],
    ['※ 「---」の行より上にキーワードを記入してください。']
  ];

  sheet.getRange(1, 1, content.length, 1).setValues(content);
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A3').setFontWeight('bold').setBackground('#e3f2fd');
  sheet.getRange('A4').setFontColor('#757575').setFontStyle('italic');
  sheet.setColumnWidth(1, 500);

  SpreadsheetApp.getUi().alert('設定シートを作成しました。\nサービスキーワードを編集してから分析を実行してください。');
}


// ============================================================
// ユーティリティ
// ============================================================

/**
 * シートを取得、なければ作成
 */
function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * 小数を指定桁数で丸める
 */
function roundTo_(num, digits) {
  if (typeof num !== 'number' || isNaN(num)) return 0;
  var factor = Math.pow(10, digits);
  return Math.round(num * factor) / factor;
}

/**
 * ヘッダー行の書式設定
 */
function formatHeaderRow_(sheet, colCount) {
  var headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1565c0');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

/**
 * 列幅を自動調整
 */
function autoResizeColumns_(sheet, colCount) {
  for (var c = 1; c <= colCount; c++) {
    sheet.autoResizeColumn(c);
  }
}
