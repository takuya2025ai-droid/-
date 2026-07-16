/**
 * 新規のGoogleスプレッドシートに、この自動化一式が前提とするシート・見出しを一括生成する。
 * 使い方: Apps Scriptエディタでこのファイルを開き、関数選択で setupSpreadsheet を選んで実行する(1回だけでよい)。
 * 既にあるシートは壊さない(見出し行が無い場合だけ追記する)。
 */
function setupSpreadsheet() {
  setupPriceBillingSheet_();
  setupPricePayoutSheet_();
  setupJobLogSheet_();
  setupShiftLogSheet_();
  setupSummarySheet_(SHEET.BILLING_SUMMARY, ['取引先', '期間開始', '期間終了', '合計請求額']);
  setupSummarySheet_(SHEET.PAYOUT_SUMMARY, ['作業者名', '期間開始', '期間終了', '合計支払額']);
  setupWorkerMasterSheet_();
  setupStatusLogSheet_();
  SpreadsheetApp.getUi().alert('シートの初期セットアップが完了しました。①②単価マスタに実際の単価を入力してください。');
}

function ensureSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function writeHeaderIfEmpty_(sheet, row, headers) {
  const existing = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const isEmpty = existing.every(function (v) { return v === ''; });
  if (!isEmpty) return;
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(row, 1, 1, headers.length).setFontWeight('bold').setBackground('#2F5496').setFontColor('#FFFFFF');
}

function setupPriceBillingSheet_() {
  const sheet = ensureSheet_(SHEET.PRICE_BILLING);
  sheet.getRange(1, 1).setValue('単価マスタ（請求＝取引先向け課金ルール）');
  writeHeaderIfEmpty_(sheet, 4,
    ['取引先', '現場', '課金方式', 'サイズ/区分', '1本目(または2名)', '2本目(または3名)', '3本目(または4名)', '4本目以降']);
  sheet.setColumnWidths(1, 8, 130);
}

function setupPricePayoutSheet_() {
  const sheet = ensureSheet_(SHEET.PRICE_PAYOUT);
  sheet.getRange(1, 1).setValue('単価マスタ（支払＝外注さんへの支払ルール）');
  writeHeaderIfEmpty_(sheet, 4, ['区分', '1本目', '2本目', '3本目', '4本目以降']);
  sheet.setColumnWidths(1, 5, 130);
}

function setupJobLogSheet_() {
  const sheet = ensureSheet_(SHEET.JOB_LOG);
  sheet.getRange(1, 1).setValue('案件実績（請求の元データ。取引先フォームから自動追記される）');
  writeHeaderIfEmpty_(sheet, 3,
    ['日付', '取引先', '現場', 'レーン', '合計フィート', '本数(合計)', '人工(人数)',
      '20F本数(自動計算)', '40F本数(自動計算)', '請求額(自動計算)']);
  sheet.setColumnWidths(1, 10, 120);
}

function setupShiftLogSheet_() {
  const sheet = ensureSheet_(SHEET.SHIFT_LOG);
  sheet.getRange(1, 1).setValue('シフト実績（支払の元データ。ここに誰がどの現場に入ったかを入力する）');
  writeHeaderIfEmpty_(sheet, 4,
    ['日付', '現場', 'レーン', '作業者名', '区分(個人/協力会社)', '担当本数(通常1)', '支払額(自動計算)']);
  sheet.setColumnWidths(1, 7, 120);
}

function setupSummarySheet_(name, headers) {
  const sheet = ensureSheet_(name);
  sheet.getRange(1, 1).setValue(name.replace(/^[①-⑩]/, ''));
  writeHeaderIfEmpty_(sheet, 3, headers);
  sheet.setColumnWidths(1, headers.length, 140);
}

function setupWorkerMasterSheet_() {
  // 見出し行がConfig.gsの想定どおり1行目に来るよう、説明タイトルは付けずヘッダーのみ書く
  const sheet = ensureSheet_(SHEET.WORKER_MASTER);
  writeHeaderIfEmpty_(sheet, 1, ['作業者名', '区分', 'LINEユーザーID', '電話番号', '備考']);
  sheet.setColumnWidths(1, 5, 160);
}

function setupStatusLogSheet_() {
  const sheet = ensureSheet_(SHEET.STATUS_LOG);
  writeHeaderIfEmpty_(sheet, 1, ['タイムスタンプ', '作業者名', '種別', '対象日', '現場']);
  sheet.setColumnWidths(1, 5, 150);
}
