/**
 * ③案件実績_請求元 と ④シフト実績_支払元 から、元のカレンダー形式のシフト表を自動生成する。
 * 入力は③④の2枚だけにして、見た目のカレンダーは常にここから再生成する「使い捨てビュー」として扱う。
 * 手で③④を直接編集すればカレンダーは常に最新化できる（逆にカレンダー側は直接編集しない）。
 */

const CAL_START_COL = 2; // A列=ブロック見出し、B列から日付列
const CAL_HEADER_ROWS = 2; // 1行目:月タイトル, 2行目:日付+曜日

/** メニューやトリガーから呼ぶ。年月省略時は当月分を生成する。 */
function generateCalendar(year, month) {
  const now = new Date();
  const y = year || now.getFullYear();
  const m = month || (now.getMonth() + 1);

  const sheetName = 'カレンダー_' + y + '年' + m + '月';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  const daysInMonth = new Date(y, m, 0).getDate();
  const dateKeys = [];
  for (let d = 1; d <= daysInMonth; d++) {
    dateKeys.push(toDateKey_(new Date(y, m - 1, d)));
  }

  const jobRows = getDataRows_(SHEET.JOB_LOG, 3).filter(function (r) {
    return dateKeys.indexOf(toDateKey_(r[JOB_COL.DATE - 1])) !== -1;
  });
  const shiftRows = getDataRows_(SHEET.SHIFT_LOG, 4).filter(function (r) {
    return dateKeys.indexOf(toDateKey_(r[SHIFT_COL.DATE - 1])) !== -1;
  });

  const blocks = buildBlockOrder_(jobRows, shiftRows);

  writeTitleRow_(sheet, y, m, dateKeys.length);
  writeDateHeaderRow_(sheet, y, m, dateKeys);

  let row = CAL_HEADER_ROWS + 1;
  blocks.forEach(function (block) {
    row = writeBlock_(sheet, row, block, dateKeys, jobRows, shiftRows);
    row += 1; // ブロック間の空行
  });

  sheet.setFrozenRows(CAL_HEADER_ROWS);
  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 160);
  for (let c = 0; c < dateKeys.length; c++) {
    sheet.setColumnWidth(CAL_START_COL + c, 90);
  }
  SpreadsheetApp.flush();
}

/** (現場, レーン) の組み合わせを、④に最初に現れた順番でブロック化する */
function buildBlockOrder_(jobRows, shiftRows) {
  const seen = {};
  const blocks = [];
  shiftRows.forEach(function (r) {
    const site = r[SHIFT_COL.SITE - 1];
    const lane = r[SHIFT_COL.LANE - 1];
    const key = site + '|||' + lane;
    if (!seen[key]) {
      seen[key] = true;
      blocks.push({ site: site, lane: lane });
    }
  });
  // シフトに載っていないが案件実績だけにあるレーンも拾っておく
  jobRows.forEach(function (r) {
    const site = r[JOB_COL.SITE - 1];
    const lane = r[JOB_COL.LANE - 1];
    const key = site + '|||' + lane;
    if (!seen[key]) {
      seen[key] = true;
      blocks.push({ site: site, lane: lane });
    }
  });
  return blocks;
}

function writeTitleRow_(sheet, y, m, dayCount) {
  sheet.getRange(1, 1).setValue(y + '/' + m + '月予定');
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
}

function writeDateHeaderRow_(sheet, y, m, dateKeys) {
  const weekdayLabels = ['日', '月', '火', '水', '木', '金', '土'];
  dateKeys.forEach(function (key, i) {
    const d = new Date(key + 'T00:00:00');
    const col = CAL_START_COL + i;
    const cell = sheet.getRange(CAL_HEADER_ROWS, col);
    cell.setValue(d.getDate() + '(' + weekdayLabels[d.getDay()] + ')');
    cell.setFontWeight('bold');
    cell.setHorizontalAlignment('center');
    if (d.getDay() === 0) {
      cell.setFontColor('#CC0000');
      cell.setBackground('#FCE4E4');
    } else if (d.getDay() === 6) {
      cell.setFontColor('#1155CC');
      cell.setBackground('#E4EEFC');
    }
  });
}

/** 1つの(現場,レーン)ブロックを書き出し、次に使える行番号を返す */
function writeBlock_(sheet, startRow, block, dateKeys, jobRows, shiftRows) {
  const labelRow = startRow;
  const workerRow = startRow + 1;
  const fRow = startRow + 2;
  const countRow = startRow + 3;

  const label = block.lane && block.lane !== '-' ? block.site + '\n' + block.lane : block.site;
  const labelCell = sheet.getRange(labelRow, 1, 3, 1);
  labelCell.merge();
  labelCell.setValue(label);
  labelCell.setFontWeight('bold');
  labelCell.setVerticalAlignment('middle');
  labelCell.setWrap(true);
  labelCell.setBackground('#EFEFEF');

  sheet.getRange(countRow, 1).setValue('人数/本数');
  sheet.getRange(countRow, 1).setFontSize(9).setFontColor('#808080');

  dateKeys.forEach(function (dateKey, i) {
    const col = CAL_START_COL + i;

    const names = shiftRows.filter(function (r) {
      return r[SHIFT_COL.SITE - 1] === block.site &&
        r[SHIFT_COL.LANE - 1] === block.lane &&
        toDateKey_(r[SHIFT_COL.DATE - 1]) === dateKey;
    }).map(function (r) { return r[SHIFT_COL.WORKER_NAME - 1]; });

    const workerCell = sheet.getRange(workerRow, col);
    workerCell.setValue(names.join('\n'));
    workerCell.setWrap(true).setVerticalAlignment('top');

    const job = jobRows.filter(function (r) {
      return r[JOB_COL.SITE - 1] === block.site &&
        r[JOB_COL.LANE - 1] === block.lane &&
        toDateKey_(r[JOB_COL.DATE - 1]) === dateKey;
    })[0];

    const fCell = sheet.getRange(fRow, col);
    if (job) {
      const totalFeet = job[JOB_COL.TOTAL_FEET - 1];
      const totalCount = job[JOB_COL.TOTAL_COUNT - 1];
      fCell.setValue(totalFeet + '/' + totalCount);
      fCell.setFontColor('#CC0000');
    }
    fCell.setHorizontalAlignment('center');

    const countCell = sheet.getRange(countRow, col);
    countCell.setValue(names.length > 0 ? names.length : '');
    countCell.setFontSize(9).setFontColor('#808080').setHorizontalAlignment('center');
  });

  return countRow + 1;
}

/** スプレッドシートのカスタムメニューから当月・翌月を選んで再生成できるようにする */
function generateCalendarThisMonth() {
  const now = new Date();
  generateCalendar(now.getFullYear(), now.getMonth() + 1);
}

function generateCalendarNextMonth() {
  const now = new Date();
  const next = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  generateCalendar(next.getFullYear(), next.getMonth() + 1);
}
