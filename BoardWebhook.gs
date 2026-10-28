// This file is server-side Apps Script code. Never copy it into index.html.
// Set GCAL_BOARD_WEBHOOK_URL in Script Properties, then run
// installBoardWebhookDelivery once in the Apps Script editor.
const BOARD_DELIVERY_SHEET = '掲示板通知キュー';
const BOARD_DELIVERY_HEADERS = ['通知ID','発生日時','本文','状態','試行回数','次回送信日時','処理期限','最終HTTP状態'];

function boardDeliverySheet_() {
  const book = getReservationSpreadsheet_();
  let sheet = book.getSheetByName(BOARD_DELIVERY_SHEET);
  if (!sheet) {
    sheet = book.insertSheet(BOARD_DELIVERY_SHEET);
    sheet.appendRow(BOARD_DELIVERY_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function boardReservationText_(title, reservation, before, extraFields) {
  // Escape values separately so authored Markdown remains active and member text stays literal.
  const escape = value => String(value).replace(/([\\`*_{}\[\]()<>#+.!|~\-])/g, '\\$1');
  const lines = ['## ' + escape(title), ''];
  reservationFields_(reservation, before).concat(extraFields || []).forEach(field => {
    const values = String(field.value).split('\n');
    lines.push('- **' + escape(field.name) + '：** ' + escape(values[0]));
    values.slice(1).forEach(value => lines.push('  - ' + escape(value)));
  });
  // Keep the authored heading, bold labels and list formatting for the Markdown-capable receiver.
  return lines.join('\n');
}

function queueBoardReservation_(title, reservation, before, extraFields) {
  if (!PropertiesService.getScriptProperties().getProperty('GCAL_BOARD_WEBHOOK_URL')) return;
  try {
    boardDeliverySheet_().appendRow([Utilities.getUuid(), new Date(), boardReservationText_(title, reservation, before, extraFields), '待機', 0, new Date(), '', '']);
  } catch (error) {
    // A notification failure must never undo a successful reservation.
    console.error('掲示板通知を保存できませんでした。予約管理シートから内容を確認してください。');
  }
}

function installBoardWebhookDelivery() {
  const url = PropertiesService.getScriptProperties().getProperty('GCAL_BOARD_WEBHOOK_URL');
  if (!url || !/^https:\/\/[^/?#]+\/hooks\/board\/[a-f0-9-]{36}\/[A-Za-z0-9_-]{43}$/.test(url))
    throw new Error('スクリプトプロパティ GCAL_BOARD_WEBHOOK_URL に掲示板の受信URLを保存してください。');
  boardDeliverySheet_();
  // Repeat setup does not create duplicate triggers. Other triggers are preserved.
  const triggers = ScriptApp.getProjectTriggers().filter(item => item.getHandlerFunction() === 'retryBoardWebhookNotifications');
  if (!triggers.length) ScriptApp.newTrigger('retryBoardWebhookNotifications').timeBased().everyMinutes(1).create();
  for (let index = 1; index < triggers.length; index++) ScriptApp.deleteTrigger(triggers[index]);
}

function retryBoardWebhookNotifications() {
  const url = PropertiesService.getScriptProperties().getProperty('GCAL_BOARD_WEBHOOK_URL');
  if (!url) return;
  if (!/^https:\/\/[^/?#]+\/hooks\/board\/[a-f0-9-]{36}\/[A-Za-z0-9_-]{43}$/.test(url))
    throw new Error('掲示板の受信URLの形式を確認してください。');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  let sheet, claimed = [];
  try {
    sheet = boardDeliverySheet_();
    const rows = sheet.getDataRange().getValues(), now = Date.now();
    for (let index = 1; index < rows.length && claimed.length < 5; index++) {
      const row = rows[index], state = row[3], due = row[5] ? new Date(row[5]).getTime() : 0;
      const lease = row[6] ? new Date(row[6]).getTime() : 0;
      if (!row[0] || ['送信済み','要確認'].includes(state) || due > now || lease > now) continue;
      const count = Number(row[4] || 0) + 1;
      sheet.getRange(index + 1, 4, 1, 4).setValues([['送信中', count, row[5], new Date(now + 10 * 60 * 1000)]]);
      claimed.push({id: String(row[0]), content: String(row[2]), count: count});
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  // External I/O happens only in this delivery function, outside reservation locks.
  for (const item of claimed) {
    let status = 0, retryAfter = 60;
    try {
      const response = UrlFetchApp.fetch(url, {method: 'post', contentType: 'application/json',
        headers: {'Idempotency-Key': item.id}, payload: JSON.stringify({content: item.content}),
        followRedirects: false, muteHttpExceptions: true});
      status = response.getResponseCode();
      const headers = response.getAllHeaders();
      const retryKey = Object.keys(headers).find(key => key.toLowerCase() === 'retry-after');
      if (retryKey && /^\d+$/.test(String(headers[retryKey]))) retryAfter = Math.min(86400, Math.max(1, Number(headers[retryKey])));
    } catch (error) { /* Keep the same notification ID. Never log its secret URL. */ }
    if (!lock.tryLock(1000)) continue; // Lease expiry recovers an interrupted acknowledgement.
    try {
      const rows = sheet.getDataRange().getValues();
      const rowIndex = rows.findIndex((row, index) => index > 0 && String(row[0]) === item.id);
      if (rowIndex < 1 || Number(rows[rowIndex][4]) !== item.count) continue;
      const done = status >= 200 && status < 300;
      const retry = !status || [408,425,429].includes(status) || status >= 500;
      const delay = Math.max(retryAfter, Math.min(3600, 60 * Math.pow(2, Math.min(item.count - 1, 6))));
      sheet.getRange(rowIndex + 1, 4, 1, 5).setValues([[done ? '送信済み' : retry ? '待機' : '要確認', item.count, done ? '' : new Date(Date.now() + delay * 1000), '', status || '通信失敗']]);
    } finally { lock.releaseLock(); }
  }
}

function testBoardWebhook() {
  queueBoardReservation_('掲示板への接続テスト', {name: '接続テスト', date: '', time: '', supportMethod: 'game'});
  retryBoardWebhookNotifications();
}
