/**
 * ③④シートに直接手入力したときも、フォーム経由と同じように自動計算の数式を自動でセットする。
 * (=フォームを使わず手入力する取引先・作業者がいても、毎回ドラッグでコピーする必要がない)
 *
 * これは単純トリガー(onEdit)なので、Apps Scriptエディタでの権限承認や
 * トリガー登録は不要。スプレッドシートを人が編集した瞬間に自動で動く。
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheet.getName() === SHEET.JOB_LOG && row > 3 && col <= JOB_COL.WORKERS_COUNT) {
    autoFillJobRow_(sheet, row);
  } else if (sheet.getName() === SHEET.SHIFT_LOG && row > 4 && col <= SHIFT_COL.UNITS_HANDLED) {
    autoFillShiftRow_(sheet, row);
  }
}

function autoFillJobRow_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, JOB_COL.WORKERS_COUNT).getValues()[0];
  const hasCoreInput = values[JOB_COL.CLIENT - 1] && values[JOB_COL.SITE - 1] &&
    values[JOB_COL.TOTAL_FEET - 1] !== '' && values[JOB_COL.TOTAL_COUNT - 1] !== '';
  if (!hasCoreInput) return;
  setJobRowFormulas_(sheet, row);
}

function autoFillShiftRow_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, SHIFT_COL.WORKER_TYPE).getValues()[0];
  const hasCoreInput = values[SHIFT_COL.SITE - 1] && values[SHIFT_COL.WORKER_NAME - 1] && values[SHIFT_COL.WORKER_TYPE - 1];
  if (!hasCoreInput) return;

  const unitsCell = sheet.getRange(row, SHIFT_COL.UNITS_HANDLED);
  if (unitsCell.getValue() === '') {
    unitsCell.setValue(1); // 通常は「1人1本」なのでデフォルト1を自動で入れる
  }
  sheet.getRange(row, SHIFT_COL.PAYOUT_AMOUNT).setFormula(
    '=CALC_PAYOUT(E' + row + ',F' + row + ')');
}
