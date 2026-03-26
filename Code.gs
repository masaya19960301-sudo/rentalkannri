// ============================================================
// Code.gs — サーバーサイド（ルーティング・スプレッドシート操作）
// ============================================================

/* ---------- 定数 ---------- */
// ★ スクリプトプロパティが設定できない場合、下の '' 内にスプレッドシートIDを直接貼り付けてください
var DIRECT_SPREADSHEET_ID = '';

function getSpreadsheet() {
  var id = DIRECT_SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('スプレッドシートIDが設定されていません。Code.gs の DIRECT_SPREADSHEET_ID にIDを貼り付けるか、スクリプトプロパティ SPREADSHEET_ID を設定してください。');
  }
  return SpreadsheetApp.openById(id);
}
function sheetRental()  { return getSpreadsheet().getSheetByName('レンタル品'); }
function sheetHistory() { return getSpreadsheet().getSheetByName('履歴'); }
function sheetConfig()  { return getSpreadsheet().getSheetByName('設定'); }

/* ---------- ルーティング ---------- */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'login';
  var isAdmin = (e && e.parameter && e.parameter.admin === '1');

  var file;
  switch (page) {
    case 'adminHome':  file = 'AdminHome'; break;
    case 'itemList':   file = 'ItemList';   break;
    case 'itemEdit':   file = 'ItemEdit';   break;
    case 'return':     file = 'Return';     break;
    case 'dataEdit':   file = 'DataEdit';   break;
    case 'history':    file = 'History';     break;
    default:           file = 'Login';      break;
  }

  var tmpl = HtmlService.createTemplateFromFile(file);
  tmpl.isAdmin = isAdmin;
  tmpl.params  = (e && e.parameter) ? e.parameter : {};
  return tmpl.evaluate()
    .setTitle('家具レンタル管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 共通：include ---------- */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================
   認証
   ============================================================ */
function verifyPassword(pw) {
  var sheet = sheetConfig();
  var correct = sheet.getRange('B1').getValue().toString();
  return pw === correct;
}

/* ============================================================
   レンタル品データ取得
   ============================================================ */
function getAllItems() {
  var sheet = sheetRental();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var data = sheet.getRange(2, 1, last - 1, 12).getValues();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return data.map(function(r, i) {
    var status = r[4];  // 貸出状況（シートの値）
    // 貸出日が入っている場合、日付に応じて自動判定
    if (status !== '貸出可能' && r[7]) {
      var lendStr = formatDate(r[7]);
      if (lendStr > today) {
        status = '貸出予定';
      } else {
        status = '貸出中';
      }
    }
    return {
      row: i + 2,
      item: r[0],    // 項目
      product: r[1], // 品名
      name: r[2],    // 名称
      id: r[3],      // ID
      status: status, // 貸出状況（自動判定）
      slipNo: r[5],  // 伝No
      customer: r[6],// お客様名
      lendDate: formatDate(r[7]),  // 貸出日
      returnDate: formatDate(r[8]),// 返却予定日
      condition: r[9],  // 状態
      shop: r[10],      // 店
      staff: r[11]      // 担当
    };
  });
}

function formatDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    var y = v.getFullYear();
    var m = ('0' + (v.getMonth() + 1)).slice(-2);
    var d = ('0' + v.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return v.toString();
}

/* ============================================================
   商品編集（IDクリック → 編集ページ）
   ============================================================ */
function getItemByRow(row) {
  var sheet = sheetRental();
  var r = sheet.getRange(row, 1, 1, 12).getValues()[0];
  return {
    row: row,
    item: r[0],
    product: r[1],
    name: r[2],
    id: r[3],
    status: r[4],
    slipNo: r[5],
    customer: r[6],
    lendDate: formatDate(r[7]),
    returnDate: formatDate(r[8]),
    condition: r[9],
    shop: r[10],
    staff: r[11]
  };
}

function updateItemByRow(row, data) {
  var sheet = sheetRental();
  sheet.getRange(row, 6).setValue(data.slipNo);     // F: 伝No
  sheet.getRange(row, 7).setValue(data.customer);    // G: お客様名
  sheet.getRange(row, 8).setValue(data.lendDate);    // H: 貸出日
  sheet.getRange(row, 9).setValue(data.returnDate);  // I: 返却予定日
  sheet.getRange(row, 11).setValue(data.shop);       // K: 店
  sheet.getRange(row, 12).setValue(data.staff);      // L: 担当
  return true;
}

/* ============================================================
   返却処理
   ============================================================ */
function getRentedCustomers() {
  var items = getAllItems();
  var map = {};
  items.forEach(function(it) {
    if (it.status === '貸出中' || it.status === '貸出予定') {
      var key = (it.slipNo || '') + '｜' + (it.customer || '');
      if (!map[key]) {
        map[key] = { slipNo: it.slipNo, customer: it.customer, label: it.slipNo + ' ' + it.customer };
      }
    }
  });
  var result = [];
  for (var k in map) result.push(map[k]);
  return result;
}

function getRentedItemsByCustomer(slipNo, customer) {
  var items = getAllItems();
  return items.filter(function(it) {
    return (it.status === '貸出中' || it.status === '貸出予定') && it.slipNo == slipNo && it.customer == customer;
  });
}

function saveInspection(row, condition) {
  var sheet = sheetRental();
  sheet.getRange(row, 10).setValue(condition); // J: 状態
  return true;
}

function finishReturn(processedItems) {
  // processedItems: [{ row, unreturned, condition }]
  var sheet = sheetRental();
  var histSheet = sheetHistory();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  processedItems.forEach(function(pi) {
    if (pi.unreturned) return; // 未返却はそのまま

    var row = pi.row;
    var vals = sheet.getRange(row, 1, 1, 12).getValues()[0];

    // 履歴へ追記
    var histRow = [
      vals[0], vals[1], vals[2], vals[3], vals[4],
      vals[5], vals[6], formatDate(vals[7]), formatDate(vals[8]),
      pi.condition || vals[9], vals[10], vals[11], today
    ];
    histSheet.appendRow(histRow);

    // レンタル品シート更新
    sheet.getRange(row, 5).setValue('貸出可能');     // E: 貸出状況
    sheet.getRange(row, 6).setValue('');              // F: 伝No
    sheet.getRange(row, 7).setValue('');              // G: お客様名
    sheet.getRange(row, 8).setValue('');              // H: 貸出日
    sheet.getRange(row, 9).setValue('');              // I: 返却予定日
    sheet.getRange(row, 10).setValue(pi.condition || vals[9]); // J: 状態
    sheet.getRange(row, 11).setValue('');             // K: 店
    sheet.getRange(row, 12).setValue('');             // L: 担当
  });

  return true;
}

/* ============================================================
   レンタル品データ編集（追加・削除・編集）
   ============================================================ */

/* --- ヘボン式ローマ字変換テーブル --- */
var ROMAJI_MAP = {
  'きゃ':'KYA','きゅ':'KYU','きょ':'KYO',
  'しゃ':'SHA','しゅ':'SHU','しょ':'SHO',
  'ちゃ':'CHA','ちゅ':'CHU','ちょ':'CHO',
  'にゃ':'NYA','にゅ':'NYU','にょ':'NYO',
  'ひゃ':'HYA','ひゅ':'HYU','ひょ':'HYO',
  'みゃ':'MYA','みゅ':'MYU','みょ':'MYO',
  'りゃ':'RYA','りゅ':'RYU','りょ':'RYO',
  'ぎゃ':'GYA','ぎゅ':'GYU','ぎょ':'GYO',
  'じゃ':'JA','じゅ':'JU','じょ':'JO',
  'びゃ':'BYA','びゅ':'BYU','びょ':'BYO',
  'ぴゃ':'PYA','ぴゅ':'PYU','ぴょ':'PYO',
  'が':'GA','ぎ':'GI','ぐ':'GU','げ':'GE','ご':'GO',
  'ざ':'ZA','じ':'JI','ず':'ZU','ぜ':'ZE','ぞ':'ZO',
  'だ':'DA','ぢ':'DI','づ':'DU','で':'DE','ど':'DO',
  'ば':'BA','び':'BI','ぶ':'BU','べ':'BE','ぼ':'BO',
  'ぱ':'PA','ぴ':'PI','ぷ':'PU','ぺ':'PE','ぽ':'PO',
  'か':'KA','き':'KI','く':'KU','け':'KE','こ':'KO',
  'さ':'SA','し':'SHI','す':'SU','せ':'SE','そ':'SO',
  'た':'TA','ち':'CHI','つ':'TSU','て':'TE','と':'TO',
  'な':'NA','に':'NI','ぬ':'NU','ね':'NE','の':'NO',
  'は':'HA','ひ':'HI','ふ':'FU','へ':'HE','ほ':'HO',
  'ま':'MA','み':'MI','む':'MU','め':'ME','も':'MO',
  'や':'YA','ゆ':'YU','よ':'YO',
  'ら':'RA','り':'RI','る':'RU','れ':'RE','ろ':'RO',
  'わ':'WA','ゐ':'WI','ゑ':'WE','を':'WO',
  'あ':'A','い':'I','う':'U','え':'E','お':'O',
  'ん':'N',
  'っ':'xTSU', // 促音は後続子音を重ねる（後処理）
  'ー':''
};

var KATAKANA_TO_HIRAGANA_OFFSET = 0x30A1 - 0x3041; // 96

function katakanaToHiragana(str) {
  var result = '';
  for (var i = 0; i < str.length; i++) {
    var code = str.charCodeAt(i);
    if (code >= 0x30A1 && code <= 0x30F6) {
      result += String.fromCharCode(code - KATAKANA_TO_HIRAGANA_OFFSET);
    } else {
      result += str.charAt(i);
    }
  }
  return result;
}

function toRomaji(text) {
  // カタカナ→ひらがな変換
  var hira = katakanaToHiragana(text);
  var result = '';
  var i = 0;
  while (i < hira.length) {
    // 2文字拗音チェック
    if (i + 1 < hira.length) {
      var two = hira.substring(i, i + 2);
      if (ROMAJI_MAP[two]) {
        result += ROMAJI_MAP[two];
        i += 2;
        continue;
      }
    }
    // 1文字
    var one = hira.charAt(i);
    if (ROMAJI_MAP[one]) {
      result += ROMAJI_MAP[one];
    } else if (/[A-Za-z0-9]/.test(one)) {
      result += one.toUpperCase();
    }
    // その他（記号等）は無視
    i++;
  }
  // 促音処理: xTSU + 子音 → 子音を重ねる
  result = result.replace(/xTSU([A-Z])/g, '$1$1');
  result = result.replace(/xTSU/g, 'TSU');
  return result;
}

function generatePrefix(item, product) {
  var itemR = toRomaji(item);
  var prodR = toRomaji(product);
  // すでにアルファベットのみならそのまま大文字
  if (!itemR) itemR = item.toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!prodR) prodR = product.toUpperCase().replace(/[^A-Z0-9]/g, '');
  return itemR + '_' + prodR;
}

function getNextSeqNumbers(prefix, count) {
  var items = getAllItems();
  var maxSeq = 0;
  var pat = prefix + '_';
  items.forEach(function(it) {
    var id = it.id || '';
    if (id.indexOf(pat) === 0) {
      var numPart = id.substring(pat.length);
      var n = parseInt(numPart, 10);
      if (!isNaN(n) && n > maxSeq) maxSeq = n;
    }
  });
  var ids = [];
  for (var i = 1; i <= count; i++) {
    var seq = maxSeq + i;
    ids.push(pat + ('000' + seq).slice(-3));
  }
  return ids;
}

function addItems(item, product, name, count, condition) {
  var prefix = generatePrefix(item, product);
  var ids = getNextSeqNumbers(prefix, count);
  var sheet = sheetRental();
  ids.forEach(function(id) {
    sheet.appendRow([item, product, name, id, '貸出可能', '', '', '', '', condition, '', '']);
  });
  return ids;
}

function deleteItemsByRows(rows) {
  var sheet = sheetRental();
  // 行番号の降順でないとズレる
  rows.sort(function(a, b) { return b - a; });
  rows.forEach(function(row) {
    sheet.deleteRow(row);
  });
  return true;
}

function updateItemMaster(row, item, product, name) {
  var sheet = sheetRental();
  var oldValues = sheet.getRange(row, 1, 1, 4).getValues()[0];
  var oldItem = oldValues[0];
  var oldProduct = oldValues[1];

  sheet.getRange(row, 1).setValue(item);   // A: 項目
  sheet.getRange(row, 2).setValue(product);// B: 品名
  sheet.getRange(row, 3).setValue(name);   // C: 名称

  // 項目 or 品名が変わった場合はID再生成
  if (oldItem !== item || oldProduct !== product) {
    var prefix = generatePrefix(item, product);
    var ids = getNextSeqNumbers(prefix, 1);
    sheet.getRange(row, 4).setValue(ids[0]); // D: ID
    return ids[0];
  }
  return oldValues[3]; // 変更なし
}

/* ============================================================
   履歴取得
   ============================================================ */
function getHistory() {
  var sheet = sheetHistory();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var data = sheet.getRange(2, 1, last - 1, 13).getValues();
  var result = data.map(function(r) {
    return {
      item: r[0],
      product: r[1],
      name: r[2],
      id: r[3],
      status: r[4],
      slipNo: r[5],
      customer: r[6],
      lendDate: formatDate(r[7]),
      returnDate: formatDate(r[8]),
      condition: r[9],
      shop: r[10],
      staff: r[11],
      returnedDate: formatDate(r[12])
    };
  });
  // 返却日の新しい順
  result.sort(function(a, b) {
    if (a.returnedDate > b.returnedDate) return -1;
    if (a.returnedDate < b.returnedDate) return 1;
    return 0;
  });
  return result;
}

/* ============================================================
   URL ヘルパー
   ============================================================ */
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}
