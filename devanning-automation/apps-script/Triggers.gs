/**
 * カスタムメニューと、時間主導トリガーの一括インストール。
 * installAllTriggers() を一度だけ手動実行すれば、以降は自動で動く。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('デバンニング自動化')
    .addItem('カレンダーを今すぐ更新(当月)', 'generateCalendarThisMonth')
    .addItem('カレンダーを今すぐ更新(翌月)', 'generateCalendarNextMonth')
    .addSeparator()
    .addItem('請求書・支払明細を再集計(先月分)', 'refreshMonthlySummariesForPreviousMonth')
    .addItem('請求書・支払明細をExcel出力(先月分)', 'exportMonthlySummariesToExcelForPreviousMonth')
    .addSeparator()
    .addItem('【テスト】明日の確認LINEを送信', 'sendTomorrowConfirmations')
    .addItem('【テスト】来週のシフトLINEを送信', 'sendWeeklyShiftToWorkers')
    .addItem('【テスト】シフト提出リマインドを送信', 'sendShiftSubmissionReminders')
    .addSeparator()
    .addItem('自動化トリガーを一括登録', 'installAllTriggers')
    .addToUi();
}

/**
 * 既存の同名トリガーを削除してから登録し直す(何度実行しても重複しない)。
 * 初回セットアップ時にメニューの「自動化トリガーを一括登録」から1回実行する。
 */
function installAllTriggers() {
  const managedFunctions = [
    'sendTomorrowConfirmations',
    'sendWeeklyShiftToWorkers',
    'sendShiftSubmissionReminders',
    'exportMonthlySummariesToExcelForPreviousMonth',
    'generateCalendarThisMonth',
  ];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (managedFunctions.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 毎日15:00: 翌日勤務者への確認LINE
  ScriptApp.newTrigger('sendTomorrowConfirmations')
    .timeBased().everyDays(1).atHour(15).nearMinute(0).create();

  // 毎日6:00: カレンダー表示を最新化
  ScriptApp.newTrigger('generateCalendarThisMonth')
    .timeBased().everyDays(1).atHour(6).nearMinute(0).create();

  // 毎週金曜9:00: 翌週シフトを個別配信
  ScriptApp.newTrigger('sendWeeklyShiftToWorkers')
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).nearMinute(0).create();

  // 毎週水曜9:00: シフト提出(空き状況)リマインド
  ScriptApp.newTrigger('sendShiftSubmissionReminders')
    .timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(9).nearMinute(0).create();

  // 毎月1日9:00: 先月分の請求書・支払明細を自動でExcel出力
  ScriptApp.newTrigger('exportMonthlySummariesToExcelForPreviousMonth')
    .timeBased().onMonthDay(1).atHour(9).nearMinute(0).create();

  SpreadsheetApp.getUi().alert('自動化トリガーを登録しました。\n' +
    '取引先フォームのonFormSubmitトリガーだけは、フォーム作成後に別途\n' +
    '「トリガー」画面から onJobFormSubmit / フォーム送信時 を手動で追加してください。');
}

/**
 * 取引先フォームの回答先をこのスプレッドシートに設定した後、コードから直接トリガー登録したい場合に使う。
 * formUrl はフォーム編集画面のURL(共有せず自分だけが使うURLでよい)。
 */
function installFormTrigger_(formUrl) {
  const form = FormApp.openByUrl(formUrl);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onJobFormSubmit') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('onJobFormSubmit').forForm(form).onFormSubmit().create();
}
