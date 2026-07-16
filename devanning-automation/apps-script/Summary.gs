/**
 * ⑤請求書サマリ・⑥支払明細サマリを、指定した月の③④実績から「その場で集計した値」として書き出す。
 * SUMIFSの数式に頼らず金額を確定値として書き込むことで、後日③④に追記・修正があっても
 * 発行済みの請求書・支払明細の金額が勝手に変わらないようにしている(請求は締めたら固定するのが安全)。
 */

function refreshMonthlySummaries(year, month) {
  const range = monthRange_(year, month);
  writeBillingSummary_(range.start, range.end);
  writePayoutSummary_(range.start, range.end);
  SpreadsheetApp.flush();
}

function writeBillingSummary_(start, end) {
  const rows = getDataRows_(SHEET.JOB_LOG, 3).filter(function (r) {
    return isDateInRange_(r[JOB_COL.DATE - 1], start, end);
  });

  const totals = {};
  rows.forEach(function (r) {
    const client = r[JOB_COL.CLIENT - 1];
    const amount = Number(r[JOB_COL.BILLING_AMOUNT - 1]) || 0;
    totals[client] = (totals[client] || 0) + amount;
  });

  const outputRows = Object.keys(totals).sort().map(function (client) {
    return [client, start, end, totals[client]];
  });

  writeSummarySheet_(SHEET.BILLING_SUMMARY,
    ['取引先', '期間開始', '期間終了', '合計請求額'], outputRows);
}

function writePayoutSummary_(start, end) {
  const rows = getDataRows_(SHEET.SHIFT_LOG, 4).filter(function (r) {
    return isDateInRange_(r[SHIFT_COL.DATE - 1], start, end);
  });

  const totals = {};
  rows.forEach(function (r) {
    const worker = r[SHIFT_COL.WORKER_NAME - 1];
    const amount = Number(r[SHIFT_COL.PAYOUT_AMOUNT - 1]) || 0;
    totals[worker] = (totals[worker] || 0) + amount;
  });

  const outputRows = Object.keys(totals).sort().map(function (worker) {
    return [worker, start, end, totals[worker]];
  });

  writeSummarySheet_(SHEET.PAYOUT_SUMMARY,
    ['作業者名', '期間開始', '期間終了', '合計支払額'], outputRows);
}

/** ヘッダー(3行目)を維持したまま、4行目以降を集計結果で置き換える */
function writeSummarySheet_(sheetName, headerLabels, outputRows) {
  const sheet = getSheet_(sheetName);
  const headerRow = 3;
  const dataStartRow = headerRow + 1;

  headerLabels.forEach(function (label, i) {
    sheet.getRange(headerRow, i + 1).setValue(label);
  });

  const lastRow = sheet.getLastRow();
  if (lastRow >= dataStartRow) {
    sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, headerLabels.length).clearContent();
  }

  if (outputRows.length === 0) return;

  const range = sheet.getRange(dataStartRow, 1, outputRows.length, headerLabels.length);
  range.setValues(outputRows);
  sheet.getRange(dataStartRow, 2, outputRows.length, 2).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(dataStartRow, 4, outputRows.length, 1).setNumberFormat('¥#,##0');
}

function refreshMonthlySummariesForPreviousMonth() {
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  refreshMonthlySummaries(prev.getFullYear(), prev.getMonth() + 1);
}
