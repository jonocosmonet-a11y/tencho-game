// ============================================================
// 店長成長ゲーム - Google Apps Script v2
// 読み書き両対応・全デバイス同期版
// ============================================================

const SHEET_NAMES = {
  pt:       "pt履歴",
  todo:     "Todo完了ログ",
  mood:     "気分ログ",
  ichiba:   "市場実績",
  revenue:  "収益商材",
  daily:    "日次スナップショット",
  gamedata: "ゲームデータ",  // ← 追加：全データ同期用
};

const HEADERS = {
  pt:      ["日時","キャラID","キャラ名","店舗名","変動pt","変動理由","累計pt","フォーム名"],
  todo:    ["日時","キャラID","キャラ名","店舗名","TodoID","Todoラベル","種別","pt","完了/取消"],
  mood:    ["日時","キャラID","キャラ名","店舗名","アイコン","コメント","更新時刻"],
  ichiba:  ["日時","キャラID","キャラ名","店舗名","MNP目標","MNP累計","YS目標","YS累計","SB機変目標","SB機変累計","YM機変目標","YM機変累計","固定目標","固定累計","ペイトク加入率"],
  revenue: ["日時","キャラID","キャラ名","店舗名","商材名","目標","実績","進捗率"],
  daily:   ["日時","キャラID","キャラ名","店舗名","累計pt","フォーム","忠誠度","やる気","機嫌","MNP累計","YS累計","SB機変累計","YM機変累計","固定累計","ペイトク加入率"],
  gamedata:["更新日時","デバイス","state_json","kpi_json","shop_json","ichiba_json"],
};

// ============================================================
// POST: データ書き込み
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const type = data.type;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ゲーム全データ同期
    if (type === "sync") {
      saveGameData(ss, data);
      return jsonResponse({ok: true, type: "sync"});
    }

    // ログ系（既存）
    const sheet = getOrCreateSheet(ss, SHEET_NAMES[type], HEADERS[type]);
    const row = buildRow(type, data);
    if (row) sheet.appendRow(row);
    return jsonResponse({ok: true});

  } catch(err) {
    return jsonResponse({ok: false, error: err.message});
  }
}

// ============================================================
// GET: データ読み込み
// ============================================================
function doGet(e) {
  const action = e.parameter.action;

  // 疎通確認
  if (!action || action === "ping") {
    return jsonResponse({ok: true, message: "店長成長ゲーム GAS稼働中 v2"});
  }

  // ゲームデータ読み込み
  if (action === "load") {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const data = loadGameData(ss);
      return jsonResponse({ok: true, data: data});
    } catch(err) {
      return jsonResponse({ok: false, error: err.message});
    }
  }

  return jsonResponse({ok: false, error: "unknown action"});
}

// ============================================================
// ゲームデータ保存（ゲームデータシートに上書き）
// ============================================================
function saveGameData(ss, data) {
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.gamedata, HEADERS.gamedata);

  // 常に2行目に最新データを上書き（履歴は持たない→高速）
  const t = now();
  const device = data.device || "unknown";
  const row = [
    t,
    device,
    JSON.stringify(data.state   || {}),
    JSON.stringify(data.kpi     || {}),
    JSON.stringify(data.shop    || {}),
    JSON.stringify(data.ichiba  || {}),
  ];

  // 2行目があれば上書き、なければ追加
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

// ============================================================
// ゲームデータ読み込み
// ============================================================
function loadGameData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.gamedata);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const row = sheet.getRange(2, 1, 1, 6).getValues()[0];
  return {
    updatedAt: row[0],
    device:    row[1],
    state:     safeJson(row[2]),
    kpi:       safeJson(row[3]),
    shop:      safeJson(row[4]),
    ichiba:    safeJson(row[5]),
  };
}

// ============================================================
// ユーティリティ
// ============================================================
function safeJson(str) {
  try { return JSON.parse(str); } catch(e) { return {}; }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    const hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground("#534AB7").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function now() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
}

function buildRow(type, d) {
  const t = now();
  switch(type) {
    case "pt":
      return [t, d.charId, d.charName, d.shopName, d.delta, d.reason, d.totalPts, d.formName];
    case "todo":
      return [t, d.charId, d.charName, d.shopName, d.todoId, d.label, d.kind, d.pts, d.done?"完了":"取消"];
    case "mood":
      return [t, d.charId, d.charName, d.shopName, d.icon, d.comment, d.time];
    case "ichiba":
      return [t, d.charId, d.charName, d.shopName,
        d.mnpTarget, d.mnpActual, d.ysTarget, d.ysActual,
        d.sbTarget, d.sbActual, d.ymTarget, d.ymActual,
        d.koteiTarget, d.koteiActual, d.paytokuRate];
    case "revenue":
      return [t, d.charId, d.charName, d.shopName, d.name, d.target, d.actual,
        d.target > 0 ? Math.round(d.actual/d.target*1000)/10 : 0];
    case "daily":
      return [t, d.charId, d.charName, d.shopName,
        d.pts, d.formName, d.loyalty, d.energy, d.mood,
        d.mnpActual, d.ysActual, d.sbActual, d.ymActual, d.koteiActual, d.paytokuRate];
    default:
      return null;
  }
}
