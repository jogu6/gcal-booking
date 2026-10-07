const CONFIG = {
  SHEET_NAME: '予約管理',
  TIME_ZONE: 'Asia/Tokyo',
  START_DATE: '2026-10-11',
  END_DATE: '2026-10-18',
  START_MINUTE: 8 * 60,
  END_MINUTE: 24 * 60,
  SLOT_MINUTES: 30,
  HEADERS: [
    'キャラクター名',
    '所属DC',
    '予約状態',
    '予約日時',
    '集合DC',
    '予約時入力名',
    '名前照合',
    '対応状況',
    '確認済み',
    '元ステータス',
    '所属ワールド',
    '備考'
  ]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('予約管理')
    .addItem('初期設定・表示更新', 'setupSheet')
    .addToUi();
}

function setupSheet() {
  const sheet = getSheet_();
  sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  sheet.setFrozenRows(1);

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, CONFIG.HEADERS.length).getValues();
    values.forEach((row, i) => {
      if (row[0] && row[0] !== '⚠ 未照合' && !row[2]) {
        sheet.getRange(i + 2, 3).setValue('未予約');
      }
    });
  }

  sheet.getRange('D:D').setNumberFormat('yyyy/mm/dd hh:mm');
  sheet.autoResizeColumns(1, CONFIG.HEADERS.length);
  applyConditionalFormatting_(sheet);
}

function doGet(e) {
  const callback = sanitizeCallback_(e.parameter.callback);
  let result;

  try {
    const action = e.parameter.action || 'slots';

    if (action === 'slots') {
      result = { ok: true, booked: getBookedSlots_() };
    } else if (action === 'member') {
      result = member_(e.parameter.name);
    } else if (action === 'book') {
      result = book_({
        date: e.parameter.date,
        time: e.parameter.time,
        name: e.parameter.name,
        dc: e.parameter.dc || ''
      });
    } else {
      result = { ok: false, message: '不明な処理です。' };
    }
  } catch (err) {
    result = { ok: false, message: err && err.message ? err.message : String(err) };
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function book_(input) {
  validateBooking_(input);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const slotKey = input.date + ' ' + input.time;
    if (getBookedSlots_().includes(slotKey)) {
      return { ok: false, message: 'この時間は先に予約されました。別の時間を選んでください。' };
    }

    const sheet = getSheet_();
    const normalizedInput = normalizeName_(input.name);
    const member = findMember_(sheet, normalizedInput);
    const dt = slotDate_(input.date, input.time);

    if (!member) {
      return {
        ok: false,
        message: '名前が間違っています。メンバーリストに登録されているキャラクターフルネームを確認してください。'
      };
    }

    const sourceStatus = String(sheet.getRange(member.row, 10).getDisplayValue() || '').trim();
    if (sourceStatus !== '未招待') {
      return {
        ok: false,
        message: sourceStatus === '登録完了'
          ? 'このキャラクターはすでに登録完了しています。'
          : 'このキャラクターは招待不要になっています。'
      };
    }

    const state = String(sheet.getRange(member.row, 3).getValue() || '');
    if (state === '予約済' || state === '完了') {
      return { ok: false, message: 'このキャラクター名はすでに予約されています。' };
    }

    sheet.getRange(member.row, 3, 1, 8).setValues([[
      '予約済',
      dt,
      input.dc,
      input.name,
      '自動一致',
      '未対応',
      '',
      ''
    ]]);

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function applyManualLinks() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const deletes = [];

  values.forEach((row, index) => {
    const rowNumber = index + 2;
    if (row[0] !== '⚠ 未照合') return;

    const targetName = String(row[8] || '').trim();
    if (!targetName) return;

    const target = findMember_(sheet, normalizeName_(targetName));
    if (!target) {
      sheet.getRange(rowNumber, 10).setValue('手動紐付け先がメンバー一覧に見つかりません');
      return;
    }

    const targetState = String(sheet.getRange(target.row, 3).getValue() || '');
    if (targetState === '予約済' || targetState === '完了') {
      sheet.getRange(rowNumber, 10).setValue('紐付け先はすでに予約済みです');
      return;
    }

    sheet.getRange(target.row, 3, 1, 8).setValues([[
      '予約済',
      row[3],
      row[4],
      row[5],
      '手動一致',
      row[7] || '未対応',
      '',
      '未照合行から手動紐付け'
    ]]);

    deletes.push(rowNumber);
  });

  deletes.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
}

function getBookedSlots_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const set = new Set();

  rows.forEach(row => {
    const state = String(row[2] || '');
    const value = row[3];

    if (!value || !['予約済', '要確認', '完了'].includes(state)) return;

    const d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return;

    set.add(Utilities.formatDate(d, CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm'));
  });

  return Array.from(set).sort();
}

function member_(name) {
  const normalized = normalizeName_(name);
  if (!normalized) return { ok: true, found: false, dc: '' };

  const sheet = getSheet_();
  const member = findMember_(sheet, normalized);
  if (!member) return { ok: true, found: false, dc: '' };

  const dc = String(sheet.getRange(member.row, 2).getDisplayValue() || '').trim();
  const allowed = ['Elemental', 'Gaia', 'Mana', 'Meteor'];
  const sourceStatus = String(sheet.getRange(member.row, 10).getDisplayValue() || '').trim();
  const eligible = sourceStatus === '未招待';

  return {
    ok: true,
    found: true,
    dc: allowed.includes(dc) ? dc : '',
    eligible: eligible,
    sourceStatus: sourceStatus
  };
}

function findMember_(sheet, normalizedName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const names = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < names.length; i++) {
    const name = String(names[i][0] || '');
    if (!name || name === '⚠ 未照合') continue;
    if (normalizeName_(name) === normalizedName) {
      return { row: i + 2, name: name };
    }
  }
  return null;
}

function normalizeName_(value) {
  return String(value || '')
    .normalize('NFKC')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function validateBooking_(input) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(input.date || '')) throw new Error('日付が不正です。');
  if (!/^\d{2}:\d{2}$/.test(input.time || '')) throw new Error('時刻が不正です。');
  if (input.date < CONFIG.START_DATE || input.date > CONFIG.END_DATE) throw new Error('予約期間外です。');

  const [h, m] = input.time.split(':').map(Number);
  const minute = h * 60 + m;
  if (minute < CONFIG.START_MINUTE || minute >= CONFIG.END_MINUTE || m % CONFIG.SLOT_MINUTES !== 0) {
    throw new Error('予約可能時間外です。');
  }

  const normalized = normalizeName_(input.name);
  if (!/^[a-z][a-z'-]* [a-z][a-z'-]*$/i.test(normalized)) {
    throw new Error('キャラクターフルネームを「Aaaa Bbbb」の形式で入力してください。');
  }

  const needsDc = input.date === '2026-10-11' || input.date === '2026-10-12';
  const allowedDc = ['Elemental', 'Gaia', 'Mana', 'Meteor'];

  if (needsDc && !allowedDc.includes(input.dc)) {
    throw new Error('集合希望DCを選択してください。');
  }

  if (!needsDc) input.dc = '';
}

function slotDate_(date, time) {
  const [y, mo, d] = date.split('-').map(Number);
  const [h, mi] = time.split(':').map(Number);
  return new Date(y, mo - 1, d, h, mi, 0, 0);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('「' + CONFIG.SHEET_NAME + '」シートがありません。');
  return sheet;
}

function sanitizeCallback_(value) {
  const s = String(value || '');
  return /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(s) ? s : '';
}

function applyConditionalFormatting_(sheet) {
  const range = sheet.getRange('C2:C');
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('未予約').setBackground('#eeeeee').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('予約済').setBackground('#d9ead3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('要確認').setBackground('#fce5cd').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('完了').setBackground('#cfe2f3').setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}
