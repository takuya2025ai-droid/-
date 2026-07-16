/**
 * 取引先向け依頼フォーム(Googleフォーム)の回答を受け取り、③案件実績_請求元へ自動追記する。
 *
 * 事前準備:
 *  1. Googleフォームを作成し、質問を以下の順番・型で用意する
 *       取引先名   (プルダウンまたは記述式)
 *       現場名     (記述式)
 *       レーン     (記述式・任意。使わない現場は空欄でよい)
 *       対象日     (日付)
 *       合計フィート (記述式・数値)
 *       本数(合計)  (記述式・数値)
 *       人工希望人数 (記述式・数値)
 *       備考       (記述式・任意)
 *  2. フォームの「回答」タブ → スプレッドシートのアイコン → このスプレッドシートを回答先に指定する
 *  3. Apps Scriptエディタ左の時計アイコン(トリガー) →
 *       実行する関数: onJobFormSubmit / イベントの種類: フォーム送信時 / 対象: このスプレッドシートの該当フォーム
 *     を追加する（installTriggers_ を1回実行すれば自動で登録される。Triggers.gs参照）
 */

/** フォームの質問順とこの配列のインデックスを対応させる（並び順を変えたらここも直す） */
const FORM_FIELD_ORDER = [
  '取引先名', '現場名', 'レーン', '対象日', '合計フィート', '本数(合計)', '人工希望人数', '備考',
];

function onJobFormSubmit(e) {
  try {
    const answers = extractFormAnswers_(e);
    validateJobFormAnswers_(answers);
    appendJobRow_(answers);
  } catch (err) {
    notifyIntakeError_(e, err);
  }
}

function extractFormAnswers_(e) {
  const result = {};
  if (e && e.namedValues) {
    FORM_FIELD_ORDER.forEach(function (label) {
      const arr = e.namedValues[label];
      result[label] = arr ? arr[0] : '';
    });
  } else if (e && e.values) {
    // e.values[0]はタイムスタンプなので1つずらす
    FORM_FIELD_ORDER.forEach(function (label, i) {
      result[label] = e.values[i + 1] !== undefined ? e.values[i + 1] : '';
    });
  } else {
    throw new Error('フォームの回答イベントを取得できませんでした。');
  }
  return result;
}

function validateJobFormAnswers_(a) {
  const required = ['取引先名', '現場名', '対象日', '合計フィート', '本数(合計)'];
  const missing = required.filter(function (key) { return a[key] === '' || a[key] === undefined; });
  if (missing.length > 0) {
    throw new Error('必須項目が空です: ' + missing.join(', '));
  }
  const totalFeet = Number(a['合計フィート']);
  const totalCount = Number(a['本数(合計)']);
  if (isNaN(totalFeet) || isNaN(totalCount) || totalCount <= 0) {
    throw new Error('合計フィート・本数(合計)が数値になっていません: ' +
      a['合計フィート'] + ' / ' + a['本数(合計)']);
  }
  // 20F/40Fの組み合わせとして整合するか(2n - T/20 と T/20 - n がどちらも0以上の整数)を事前検算
  const t20 = totalFeet / 20;
  const count20 = 2 * totalCount - t20;
  const count40 = t20 - totalCount;
  if (count20 < 0 || count40 < 0 || !Number.isInteger(count20) || !Number.isInteger(count40)) {
    throw new Error('合計フィート(' + totalFeet + ')と本数(' + totalCount +
      ')の組み合わせが20F/40Fとして成立しません。取引先へ確認してください。');
  }
}

function appendJobRow_(a) {
  const sheet = getSheet_(SHEET.JOB_LOG);
  const lastRow = Math.max(sheet.getLastRow(), 3); // 3行目までヘッダー
  const newRow = lastRow + 1;

  const dateValue = a['対象日'] instanceof Date ? a['対象日'] : new Date(a['対象日']);
  sheet.getRange(newRow, JOB_COL.DATE).setValue(dateValue);
  sheet.getRange(newRow, JOB_COL.CLIENT).setValue(a['取引先名']);
  sheet.getRange(newRow, JOB_COL.SITE).setValue(a['現場名']);
  sheet.getRange(newRow, JOB_COL.LANE).setValue(a['レーン'] || '-');
  sheet.getRange(newRow, JOB_COL.TOTAL_FEET).setValue(Number(a['合計フィート']));
  sheet.getRange(newRow, JOB_COL.TOTAL_COUNT).setValue(Number(a['本数(合計)']));
  sheet.getRange(newRow, JOB_COL.WORKERS_COUNT).setValue(a['人工希望人数'] ? Number(a['人工希望人数']) : '');

  copyFormulaDown_(sheet, newRow, JOB_COL.COUNT_20F, JOB_COL.BILLING_AMOUNT);
}

/**
 * H〜J列(20F本数・40F本数・請求額)の数式を、直前の行から相対参照でコピーする。
 * 直前行に数式がない(初回投入時など)場合は、プロトタイプと同じ数式をこの行番号向けに組み立てる。
 */
function copyFormulaDown_(sheet, newRow, fromCol, toCol) {
  const prevRow = newRow - 1;
  const prevFormulas = sheet.getRange(prevRow, fromCol, 1, toCol - fromCol + 1).getFormulas();
  const hasFormula = prevFormulas[0].some(function (f) { return f !== ''; });
  if (hasFormula) {
    sheet.getRange(prevRow, fromCol, 1, toCol - fromCol + 1)
      .copyTo(sheet.getRange(newRow, fromCol, 1, toCol - fromCol + 1));
    return;
  }
  const r = newRow;
  sheet.getRange(r, JOB_COL.COUNT_20F).setFormula(
    '=IF(OR(E' + r + '="",F' + r + '=""),"",2*F' + r + '-E' + r + '/20)');
  sheet.getRange(r, JOB_COL.COUNT_40F).setFormula(
    '=IF(OR(E' + r + '="",F' + r + '=""),"",E' + r + '/20-F' + r + ')');
  const m = "'" + SHEET.PRICE_BILLING + "'";
  const netforceCalc = 'INDEX(' + m + '!$E$7:$H$7,1,MIN(G' + r + ',4))';
  const sum20 = '(IF(H' + r + '>=1,' + m + '!$E$5,0)+IF(H' + r + '>=2,' + m + '!$F$5,0)' +
    '+IF(H' + r + '>=3,' + m + '!$G$5,0)+IF(H' + r + '>3,(H' + r + '-3)*' + m + '!$H$5,0))';
  const sum40 = '(IF(I' + r + '>=1,' + m + '!$E$6,0)+IF(I' + r + '>=2,' + m + '!$F$6,0)' +
    '+IF(I' + r + '>=3,' + m + '!$G$6,0)+IF(I' + r + '>3,(I' + r + '-3)*' + m + '!$H$6,0))';
  sheet.getRange(r, JOB_COL.BILLING_AMOUNT).setFormula(
    '=IF(B' + r + '="Netforce",' + netforceCalc + ',' + sum20 + '+' + sum40 + ')');
}

function notifyIntakeError_(e, err) {
  console.error('フォーム取り込みエラー: ' + err.message);
  try {
    const email = getScriptProp_('NOTIFY_EMAIL');
    const raw = e && e.namedValues ? JSON.stringify(e.namedValues) : (e && e.values ? e.values.join(', ') : '(データなし)');
    MailApp.sendEmail(email,
      '【要確認】取引先フォームの自動取り込みに失敗しました',
      'エラー内容: ' + err.message + '\n\n受信内容:\n' + raw +
      '\n\n③案件実績_請求元に手動で追記してください。');
  } catch (mailErr) {
    console.error('通知メールの送信にも失敗: ' + mailErr.message);
  }
}
