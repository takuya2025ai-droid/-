/**
 * 全スクリプト共通の設定。シート名や列位置をここで一元管理する。
 * シート構成やレイアウトを変える場合はこのファイルだけ直せばよい。
 */

const SHEET = {
  PRICE_BILLING: '①単価マスタ_請求',
  PRICE_PAYOUT: '②単価マスタ_支払',
  JOB_LOG: '③案件実績_請求元',
  SHIFT_LOG: '④シフト実績_支払元',
  BILLING_SUMMARY: '⑤請求書サマリ_自動集計',
  PAYOUT_SUMMARY: '⑥支払明細サマリ_自動集計',
  WORKER_MASTER: '⑦作業者マスタ',
  STATUS_LOG: '⑧稼働ログ',
};

// ③案件実績_請求元 の列（1始まり）
const JOB_COL = {
  DATE: 1,
  CLIENT: 2,
  SITE: 3,
  LANE: 4,
  TOTAL_FEET: 5,
  TOTAL_COUNT: 6,
  WORKERS_COUNT: 7,
  COUNT_20F: 8,
  COUNT_40F: 9,
  BILLING_AMOUNT: 10,
};

// ④シフト実績_支払元 の列
const SHIFT_COL = {
  DATE: 1,
  SITE: 2,
  LANE: 3,
  WORKER_NAME: 4,
  WORKER_TYPE: 5,
  UNITS_HANDLED: 6,
  PAYOUT_AMOUNT: 7,
};

// ⑦作業者マスタ の列
const WORKER_COL = {
  NAME: 1,
  TYPE: 2,
  LINE_USER_ID: 3,
  PHONE: 4,
  NOTE: 5,
};

// ⑧稼働ログ の列
const STATUS_COL = {
  TIMESTAMP: 1,
  WORKER_NAME: 2,
  STATUS: 3,
  DATE: 4,
  SITE: 5,
};

const STATUS_LABEL = {
  DEPARTED: '出発',
  ARRIVED: '到着',
  FINISHED: '業務終了',
};

/**
 * スクリプトプロパティ（機密情報）は必ずここ経由で読む。
 * 設定方法: 拡張機能 > Apps Script > プロジェクトの設定 > スクリプト プロパティ
 *   LINE_CHANNEL_ACCESS_TOKEN : LINE Developersで発行したチャネルアクセストークン
 *   LINE_CHANNEL_SECRET       : Webhook署名検証用のチャネルシークレット
 *   NOTIFY_EMAIL              : 請求書/支払明細を自動送付するメールアドレス
 *   DRIVE_EXPORT_FOLDER_ID    : Excel出力の保存先Google DriveフォルダID
 */
function getScriptProp_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('スクリプト プロパティ "' + key + '" が未設定です。プロジェクトの設定から登録してください。');
  }
  return value;
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('シート "' + name + '" が見つかりません。シート名を確認してください。');
  }
  return sheet;
}

/** ヘッダー行を除いた実データ行だけを2次元配列で返す（末尾の空行は除外） */
function getDataRows_(sheetName, headerRowCount) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= headerRowCount) return [];
  return sheet.getRange(headerRowCount + 1, 1, lastRow - headerRowCount, lastCol).getValues()
    .filter(function (row) { return row.some(function (cell) { return cell !== '' && cell !== null; }); });
}

/** Dateまたは文字列をYYYY-MM-DD形式の文字列キーに正規化する */
function toDateKey_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(value).trim();
}

/** 指定した年月の月初・月末のDateを返す */
function monthRange_(year, month) {
  return {
    start: new Date(year, month - 1, 1),
    end: new Date(year, month, 0),
  };
}

/** 日付値(Dateまたは文字列)がstart〜end(両端含む)の範囲内か判定する */
function isDateInRange_(value, start, end) {
  const d = value instanceof Date ? value : new Date(value);
  const startKey = toDateKey_(start);
  const endKey = toDateKey_(end);
  const key = toDateKey_(d);
  return key >= startKey && key <= endKey;
}
