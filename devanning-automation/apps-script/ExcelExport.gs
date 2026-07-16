/**
 * ⑤請求書サマリ・⑥支払明細サマリをExcel(.xlsx)としてGoogle Driveに保存し、メールで送付する。
 * 事前準備(スクリプト プロパティ):
 *   DRIVE_EXPORT_FOLDER_ID : 出力先Google DriveフォルダID
 *   NOTIFY_EMAIL           : 送付先メールアドレス(複数なら "a@x.com,b@x.com")
 */

function exportMonthlySummariesToExcel(year, month) {
  refreshMonthlySummaries(year, month);
  const periodLabel = year + '年' + month + '月分';
  return exportSheetsToExcel_([SHEET.BILLING_SUMMARY, SHEET.PAYOUT_SUMMARY], periodLabel);
}

function exportMonthlySummariesToExcelForPreviousMonth() {
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return exportMonthlySummariesToExcel(prev.getFullYear(), prev.getMonth() + 1);
}

/**
 * 指定シートだけを一時スプレッドシートにコピーしてxlsx変換し、Driveに保存＋メール添付する。
 * (Googleスプレッドシートのexport APIはブック単位でしか書き出せないため、
 *  必要なシートだけの一時ブックを作ってから変換する)
 */
function exportSheetsToExcel_(sheetNames, periodLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSs = SpreadsheetApp.create('devanning_export_tmp_' + periodLabel);

  sheetNames.forEach(function (name) {
    const src = getSheet_(name);
    const copied = src.copyTo(tempSs);
    copied.setName(name);
  });
  // 新規ブック作成時に自動で入る空のデフォルトシートを削除
  tempSs.getSheets().forEach(function (sheet) {
    if (sheetNames.indexOf(sheet.getName()) === -1) {
      tempSs.deleteSheet(sheet);
    }
  });
  SpreadsheetApp.flush();

  const url = 'https://docs.google.com/spreadsheets/d/' + tempSs.getId() + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
  });
  const fileName = '請求書_支払明細_' + periodLabel + '.xlsx';
  const blob = response.getBlob().setName(fileName);

  const folder = DriveApp.getFolderById(getScriptProp_('DRIVE_EXPORT_FOLDER_ID'));
  const file = folder.createFile(blob);

  // 変換専用に作った一時ブックは残さずゴミ箱へ
  DriveApp.getFileById(tempSs.getId()).setTrashed(true);

  const email = getScriptProp_('NOTIFY_EMAIL');
  MailApp.sendEmail({
    to: email,
    subject: '【自動出力】請求書・支払明細サマリ ' + periodLabel,
    body: periodLabel + 'の請求書サマリ・支払明細サマリをExcelで出力しました。\n' +
      'Driveの保存先: ' + file.getUrl() + '\n\n添付ファイルも合わせてご確認ください。\n' +
      'ここから金額を各社所定の請求書・支払明細書フォーマットに転記してください。',
    attachments: [blob],
  });

  return file.getUrl();
}
