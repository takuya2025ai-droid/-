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
  insertSampleDataIfEmpty_();
  SpreadsheetApp.getUi().alert(
    '初期セットアップが完了しました。\n' +
    '③④シートには動作確認用のサンプル行を1件ずつ入れてあります。H〜J列・G列に金額が自動で入っているか確認してください。\n' +
    '確認できたらサンプル行は削除し、①②に実際の単価を入力してから運用を始めてください。');
}

/** 動作確認用のサンプルデータを③④に1行ずつ入れる(既にデータが入っている場合は何もしない) */
function insertSampleDataIfEmpty_() {
  const priceBilling = getSheet_(SHEET.PRICE_BILLING);
  if (priceBilling.getRange(5, 1, 1, 8).getValues()[0].every(function (v) { return v === ''; })) {
    priceBilling.getRange(5, 1, 3, 8).setValues([
      ['サンプル取引先A', 'サンプル現場', '本数ベース', '20F', 36000, 34000, 32000, 30000],
      ['サンプル取引先A', 'サンプル現場', '本数ベース', '40F', 25000, 25000, 25000, 25000],
      ['サンプル取引先B', 'サンプル現場2', '人工ベース', '人工数', 12000, 17000, 23000, 23000],
    ]);
  }

  const pricePayout = getSheet_(SHEET.PRICE_PAYOUT);
  if (pricePayout.getRange(5, 1, 1, 5).getValues()[0].every(function (v) { return v === ''; })) {
    pricePayout.getRange(5, 1, 2, 5).setValues([
      ['個人事業主(スズメ)', 4000, 7000, 9000, 10000],
      ['協力会社', 5000, 8000, 10000, 11000],
    ]);
  }

  const jobLog = getSheet_(SHEET.JOB_LOG);
  if (jobLog.getRange(4, 1, 1, 7).getValues()[0].every(function (v) { return v === ''; })) {
    const today = new Date();
    jobLog.getRange(4, 1, 1, 7).setValues([
      [today, 'サンプル取引先A', 'サンプル現場', '1レーン目', 80, 3, 3],
    ]);
    setJobRowFormulas_(jobLog, 4);
  }

  const shiftLog = getSheet_(SHEET.SHIFT_LOG);
  if (shiftLog.getRange(5, 1, 1, 5).getValues()[0].every(function (v) { return v === ''; })) {
    const today = new Date();
    shiftLog.getRange(5, 1, 3, 5).setValues([
      [today, 'サンプル現場', '1レーン目', '山田太郎(サンプル)', '個人事業主(スズメ)'],
      [today, 'サンプル現場', '1レーン目', '佐藤次郎(サンプル)', '個人事業主(スズメ)'],
      [today, 'サンプル現場', '1レーン目', '鈴木三郎(サンプル)', '個人事業主(スズメ)'],
    ]);
    for (let r = 5; r <= 7; r++) {
      shiftLog.getRange(r, SHIFT_COL.UNITS_HANDLED).setValue(1);
      shiftLog.getRange(r, SHIFT_COL.PAYOUT_AMOUNT).setFormula('=CALC_PAYOUT(E' + r + ',F' + r + ')');
    }
  }
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
