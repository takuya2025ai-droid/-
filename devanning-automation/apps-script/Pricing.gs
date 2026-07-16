/**
 * ①②単価マスタを「行数固定」ではなく取引先名・区分で検索して計算する。
 * これにより①に何社追加しても(25社でも)、行の追加だけで自動計算に反映される。
 *
 * スプレッドシート上のセルから直接呼び出せる通常の関数(カスタム関数)として実装している。
 * 例: ③案件実績_請求元のJ列に  =CALC_BILLING(B4,C4,H4,I4,G4)
 */

/**
 * 案件実績1行分の請求額を計算する。
 * @param {string} client 取引先名
 * @param {string} site 現場名
 * @param {number} count20 20F本数
 * @param {number} count40 40F本数
 * @param {number} workerCount 人工(人数)
 * @return {number} 請求額
 * @customfunction
 */
function CALC_BILLING(client, site, count20, count40, workerCount) {
  if (!client || !site) return '';

  const rows = getDataRows_(SHEET.PRICE_BILLING, 4);
  const matches = rows.filter(function (r) { return r[0] === client && r[1] === site; });
  if (matches.length === 0) {
    throw new Error('①単価マスタ_請求に「' + client + ' / ' + site + '」の単価が見つかりません。行を追加してください。');
  }

  const laborRow = matches.filter(function (r) { return r[2] === '人工ベース'; })[0];
  if (laborRow) {
    return sumTieredPrice_(laborRow.slice(4, 7), laborRow[7], workerCount);
  }

  const row20 = matches.filter(function (r) { return r[3] === '20F'; })[0];
  const row40 = matches.filter(function (r) { return r[3] === '40F'; })[0];
  let total = 0;
  if (row20) total += sumTieredPrice_(row20.slice(4, 7), row20[7], count20);
  if (row40) total += sumTieredPrice_(row40.slice(4, 7), row40[7], count40);
  return total;
}

/**
 * シフト実績1行分の支払額を計算する。
 * @param {string} workerType 区分(個人事業主/協力会社などの②の1列目と一致する文字列)
 * @param {number} unitsHandled 担当本数
 * @return {number} 支払額
 * @customfunction
 */
function CALC_PAYOUT(workerType, unitsHandled) {
  if (!workerType) return '';

  const rows = getDataRows_(SHEET.PRICE_PAYOUT, 4);
  const match = rows.filter(function (r) { return r[0] === workerType; })[0];
  if (!match) {
    throw new Error('②単価マスタ_支払に区分「' + workerType + '」が見つかりません。行を追加してください。');
  }
  return sumTieredPrice_(match.slice(1, 4), match[4], unitsHandled);
}

/** 1本目・2本目・3本目 + 4本目以降(定額) の階段単価を、本数ぶん合算する */
function sumTieredPrice_(firstThreeTiers, fourthPlusTier, count) {
  let remaining = Number(count) || 0;
  let total = 0;
  for (let i = 0; i < firstThreeTiers.length && remaining > 0; i++) {
    total += Number(firstThreeTiers[i]) || 0;
    remaining--;
  }
  if (remaining > 0) {
    total += remaining * (Number(fourthPlusTier) || 0);
  }
  return total;
}
