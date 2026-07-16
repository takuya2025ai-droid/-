/**
 * LINE公式アカウント連携。
 *  - 出発/到着/業務終了 の報告をボタンのpostbackで受け取り⑧稼働ログに記録
 *  - 毎日15時: 翌日勤務者への確認メッセージ配信
 *  - 毎週金曜: 翌週シフトの個別配信
 *  - 毎週水曜: シフト提出(空き状況)のリマインド
 *
 * 事前準備:
 *  1. LINE Developersコンソールで Messaging API チャネルを作成する
 *  2. 「チャネルアクセストークン(長期)」を発行し、スクリプト プロパティ LINE_CHANNEL_ACCESS_TOKEN に保存
 *  3. 「チャネルシークレット」をスクリプト プロパティ LINE_CHANNEL_SECRET に保存
 *  4. このApps Scriptを「ウェブアプリとしてデプロイ」(実行者:自分 / アクセスできるユーザー:全員)し、
 *     発行されたURLをLINE DevelopersのWebhook URLに登録して「Webhookの利用」をONにする
 *  5. 応答メッセージ(自動応答)機能はOFFにしておく(このスクリプトのdoPostに完全に任せる)
 *  6. 外注さん各自に、LINE公式アカウントを友だち追加した直後に何かひとこと送ってもらう
 *     → ⑦作業者マスタに未紐付けのuserIdが自動追加されるので、氏名欄を手動で埋めて紐付ける
 */

const LINE_API = {
  PUSH: 'https://api.line.me/v2/bot/message/push',
  MULTICAST: 'https://api.line.me/v2/bot/message/multicast',
  REPLY: 'https://api.line.me/v2/bot/message/reply',
  PROFILE: 'https://api.line.me/v2/bot/profile/',
};

function doPost(e) {
  // 注意: Apps ScriptのdoPost(e)はリクエストヘッダーを取得する手段を提供していないため、
  // LINEが送るX-Line-Signatureを使った正規のHMAC検証はここでは実施できない。
  // 対策として、Webアプリのデプロイ後URL(推測困難なランダム文字列を含む)を外部に公開しない運用でリスクを抑える。
  const rawBody = e.postData.contents;
  const payload = JSON.parse(rawBody);
  const events = payload.events || [];
  events.forEach(function (event) {
    try {
      handleLineEvent_(event);
    } catch (err) {
      console.error('LINEイベント処理エラー: ' + err.message + ' / event=' + JSON.stringify(event));
    }
  });

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleLineEvent_(event) {
  if (event.type === 'follow') {
    handleFollowEvent_(event);
    return;
  }
  if (event.type === 'postback') {
    handlePostbackEvent_(event);
    return;
  }
  if (event.type === 'message' && event.message.type === 'text') {
    handleTextMessageEvent_(event);
    return;
  }
}

/** 友だち追加された直後: userIdを⑦作業者マスタの末尾に仮登録し、管理者が氏名を後から埋められるようにする */
function handleFollowEvent_(event) {
  const userId = event.source.userId;
  const sheet = getSheet_(SHEET.WORKER_MASTER);
  const existing = findWorkerRowByUserId_(userId);
  if (existing) return;

  const profile = fetchLineProfile_(userId);
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const newRow = lastRow + 1;
  sheet.getRange(newRow, WORKER_COL.NAME).setValue('【要紐付け】' + (profile ? profile.displayName : ''));
  sheet.getRange(newRow, WORKER_COL.LINE_USER_ID).setValue(userId);
  sheet.getRange(newRow, WORKER_COL.NOTE).setValue('管理者が氏名を確認して書き換えてください');

  replyLineMessage_(event.replyToken,
    'お友だち登録ありがとうございます。管理者が確認後、シフト連絡や確認メッセージをこちらにお送りします。');
}

/** ボタン(postback)での 出発/到着/業務終了 報告を⑧稼働ログに記録する */
function handlePostbackEvent_(event) {
  const data = parsePostbackData_(event.postback.data);
  if (data.action !== 'status') return;

  const userId = event.source.userId;
  const workerRow = findWorkerRowByUserId_(userId);
  const workerName = workerRow ? workerRow[WORKER_COL.NAME - 1] : '(未登録:' + userId + ')';

  const statusSheet = getSheet_(SHEET.STATUS_LOG);
  const newRow = statusSheet.getLastRow() + 1;
  statusSheet.getRange(newRow, STATUS_COL.TIMESTAMP).setValue(new Date());
  statusSheet.getRange(newRow, STATUS_COL.WORKER_NAME).setValue(workerName);
  statusSheet.getRange(newRow, STATUS_COL.STATUS).setValue(data.status);
  statusSheet.getRange(newRow, STATUS_COL.DATE).setValue(data.date || toDateKey_(new Date()));
  statusSheet.getRange(newRow, STATUS_COL.SITE).setValue(data.site || '');

  replyLineMessage_(event.replyToken, workerName + 'さん、「' + data.status + '」を記録しました。');
}

/** 想定外の自由文が来た場合は簡易な案内だけ返す(自動解析はしない) */
function handleTextMessageEvent_(event) {
  replyLineMessage_(event.replyToken,
    'メッセージありがとうございます。出発・到着・業務終了はお送りしたボタンからご報告ください。');
}

function parsePostbackData_(raw) {
  const result = {};
  raw.split('&').forEach(function (pair) {
    const kv = pair.split('=');
    result[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1] || '');
  });
  return result;
}

function findWorkerRowByUserId_(userId) {
  const rows = getDataRows_(SHEET.WORKER_MASTER, 1);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][WORKER_COL.LINE_USER_ID - 1] === userId) return rows[i];
  }
  return null;
}

function findUserIdByWorkerName_(name) {
  const rows = getDataRows_(SHEET.WORKER_MASTER, 1);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][WORKER_COL.NAME - 1] === name) return rows[i][WORKER_COL.LINE_USER_ID - 1];
  }
  return null;
}

function fetchLineProfile_(userId) {
  try {
    const resp = UrlFetchApp.fetch(LINE_API.PROFILE + userId, {
      headers: { Authorization: 'Bearer ' + getScriptProp_('LINE_CHANNEL_ACCESS_TOKEN') },
    });
    return JSON.parse(resp.getContentText());
  } catch (err) {
    console.error('プロフィール取得失敗: ' + err.message);
    return null;
  }
}

function replyLineMessage_(replyToken, text) {
  UrlFetchApp.fetch(LINE_API.REPLY, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + getScriptProp_('LINE_CHANNEL_ACCESS_TOKEN') },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: text }] }),
  });
}

function pushLineMessage_(userId, messages) {
  UrlFetchApp.fetch(LINE_API.PUSH, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + getScriptProp_('LINE_CHANNEL_ACCESS_TOKEN') },
    payload: JSON.stringify({ to: userId, messages: messages }),
  });
}

function statusButtonsTemplate_(dateKey, site) {
  const qs = function (status) {
    return 'action=status&status=' + encodeURIComponent(status) +
      '&date=' + encodeURIComponent(dateKey) + '&site=' + encodeURIComponent(site || '');
  };
  return {
    type: 'template',
    altText: dateKey + ' の稼働報告',
    template: {
      type: 'buttons',
      text: dateKey + ' 稼働報告をお願いします',
      actions: [
        { type: 'postback', label: '出発', data: qs(STATUS_LABEL.DEPARTED) },
        { type: 'postback', label: '到着', data: qs(STATUS_LABEL.ARRIVED) },
        { type: 'postback', label: '業務終了', data: qs(STATUS_LABEL.FINISHED) },
      ],
    },
  };
}

/** 毎日15時: 翌日勤務者ひとりひとりに、現場情報+稼働報告ボタンを送る */
function sendTomorrowConfirmations() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const dateKey = toDateKey_(tomorrow);

  const rows = getDataRows_(SHEET.SHIFT_LOG, 4).filter(function (r) {
    return toDateKey_(r[SHIFT_COL.DATE - 1]) === dateKey;
  });

  const bySite = {};
  rows.forEach(function (r) {
    const name = r[SHIFT_COL.WORKER_NAME - 1];
    if (!bySite[name]) bySite[name] = [];
    bySite[name].push(r[SHIFT_COL.SITE - 1] + (r[SHIFT_COL.LANE - 1] && r[SHIFT_COL.LANE - 1] !== '-' ? '(' + r[SHIFT_COL.LANE - 1] + ')' : ''));
  });

  Object.keys(bySite).forEach(function (name) {
    const userId = findUserIdByWorkerName_(name);
    if (!userId) {
      console.warn('LINE未紐付けのため送信スキップ: ' + name);
      return;
    }
    const sites = bySite[name].join('、');
    const textMsg = { type: 'text', text: '【明日の確認】' + dateKey + '\n現場: ' + sites + '\nよろしくお願いします。' };
    pushLineMessage_(userId, [textMsg, statusButtonsTemplate_(dateKey, bySite[name][0])]);
  });
}

/** 毎週金曜: 翌週(月〜日)のシフトを本人ごとにまとめて配信 */
function sendWeeklyShiftToWorkers() {
  const today = new Date();
  const nextMonday = new Date(today);
  nextMonday.setDate(today.getDate() + ((8 - today.getDay()) % 7 || 7));
  const weekDates = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(nextMonday);
    d.setDate(nextMonday.getDate() + i);
    weekDates.push(toDateKey_(d));
  }

  const rows = getDataRows_(SHEET.SHIFT_LOG, 4).filter(function (r) {
    return weekDates.indexOf(toDateKey_(r[SHIFT_COL.DATE - 1])) !== -1;
  });

  const byWorker = {};
  rows.forEach(function (r) {
    const name = r[SHIFT_COL.WORKER_NAME - 1];
    const dateKey = toDateKey_(r[SHIFT_COL.DATE - 1]);
    const site = r[SHIFT_COL.SITE - 1] + (r[SHIFT_COL.LANE - 1] && r[SHIFT_COL.LANE - 1] !== '-' ? '(' + r[SHIFT_COL.LANE - 1] + ')' : '');
    if (!byWorker[name]) byWorker[name] = [];
    byWorker[name].push(dateKey + ' : ' + site);
  });

  Object.keys(byWorker).forEach(function (name) {
    const userId = findUserIdByWorkerName_(name);
    if (!userId) {
      console.warn('LINE未紐付けのため送信スキップ: ' + name);
      return;
    }
    const body = '【来週のシフト】\n' + byWorker[name].join('\n');
    pushLineMessage_(userId, [{ type: 'text', text: body }]);
  });
}

/** 毎週水曜: 全登録作業者へ、来週の空き状況(シフト提出)を促すリマインド */
function sendShiftSubmissionReminders() {
  const rows = getDataRows_(SHEET.WORKER_MASTER, 1).filter(function (r) {
    return r[WORKER_COL.LINE_USER_ID - 1] && String(r[WORKER_COL.NAME - 1]).indexOf('【要紐付け】') !== 0;
  });

  const userIds = rows.map(function (r) { return r[WORKER_COL.LINE_USER_ID - 1]; });
  const chunks = chunkArray_(userIds, 500); // multicastは1回500件まで
  chunks.forEach(function (chunk) {
    UrlFetchApp.fetch(LINE_API.MULTICAST, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + getScriptProp_('LINE_CHANNEL_ACCESS_TOKEN') },
      payload: JSON.stringify({
        to: chunk,
        messages: [{ type: 'text', text: '来週の空き状況をいつもの方法でご提出ください。よろしくお願いします。' }],
      }),
    });
  });
}

function chunkArray_(arr, size) {
  const chunks = [];
  for (let i = 0; i < arr.length; i += size) {
    chunks.push(arr.slice(i, i + size));
  }
  return chunks;
}
