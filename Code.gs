const CONFIG = {
  SPREADSHEET_ID: '1lF_hqTu9Oe2VbH7-CEyKg14942SWgGKpwJ5Z-cGCz04',
  SHEET_NAME: '招待状況',
  TIME_ZONE: 'Asia/Tokyo',
  START_DATE: '2026-10-11',
  END_DATE: '2026-10-18',
  FIRST_DAY_START_TIME: '18:00',
  START_MINUTE: 8 * 60,
  END_MINUTE: 24 * 60,
  SLOT_MINUTES: 30,
  FIRST_DATA_ROW: 4,

  // A:確認済み B:キャラクター名 C:DC名/ワールド名
  // D:招待日 E:招待時刻 F:ステータス G:備考 H:集合DC
  COL_NAME: 2,
  COL_HOME: 3,
  COL_DATE: 4,
  COL_TIME: 5,
  COL_STATUS: 6,
  COL_NOTE: 7,
  COL_MEETING_DC: 8
};

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
        dc: e.parameter.dc || '',
        useX: e.parameter.useX === '1',
        xAccount: e.parameter.xAccount || ''
      });
    } else if (action === 'cancel') {
      result = cancel_(e.parameter.name);
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

function member_(name) {
  const normalized = normalizeName_(name);
  if (!normalized) return { ok: true, found: false, dc: '' };

  const sheet = getSheet_();
  const member = findMember_(sheet, normalized);
  if (!member) return { ok: true, found: false, dc: '' };

  const home = String(sheet.getRange(member.row, CONFIG.COL_HOME).getDisplayValue() || '').trim();
  const homeDc = home.includes('/') ? home.split('/')[0] : home;
  const status = String(sheet.getRange(member.row, CONFIG.COL_STATUS).getDisplayValue() || '').trim();
  const dateText = String(sheet.getRange(member.row, CONFIG.COL_DATE).getDisplayValue() || '').trim();
  const timeText = normalizeTime_(String(sheet.getRange(member.row, CONFIG.COL_TIME).getDisplayValue() || '').trim());
  const bookingDate = normalizeSheetDate_(dateText);
  const meetingDc = String(sheet.getRange(member.row, CONFIG.COL_MEETING_DC).getDisplayValue() || '').trim();
  const note = String(sheet.getRange(member.row, CONFIG.COL_NOTE).getDisplayValue() || '');
  const xMatch = note.match(/(?:^|\n)X:\s*(@[A-Za-z0-9_]{1,15})(?:$|\n)/);
  const hasBooking = status === '予約済' && !!bookingDate && !!timeText;

  return {
    ok: true,
    found: true,
    dc: ['Elemental','Gaia','Mana','Meteor'].includes(homeDc) ? homeDc : '',
    meetingDc: ['Elemental','Gaia','Mana','Meteor'].includes(meetingDc) ? meetingDc : '',
    eligible: status === '未招待' || status === '予約済',
    sourceStatus: status,
    canCancel: hasBooking,
    bookingDate: hasBooking ? bookingDate : '',
    bookingTime: hasBooking ? timeText : '',
    bookingKey: hasBooking ? bookingDate + ' ' + timeText : '',
    useX: !!xMatch,
    xAccount: xMatch ? xMatch[1] : ''
  };
}

function book_(input) {
  validateBooking_(input);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_();
    const member = findMember_(sheet, normalizeName_(input.name));

    if (!member) {
      return {
        ok: false,
        message: '名前が間違っています。メンバーリストに登録されているキャラクターフルネームを確認してください。'
      };
    }

    const status = String(sheet.getRange(member.row, CONFIG.COL_STATUS).getDisplayValue() || '').trim();

    if (status === '登録完了') {
      return { ok: false, message: 'このキャラクターはすでに登録完了しています。' };
    }
    if (status === '招待不要') {
      return {
        ok: false,
        message: 'このキャラクターは今回の予約対象ではありません。サブキャラの場合は、メインキャラでご予約ください。'
      };
    }
    if (status === '招待済み') {
      return { ok: false, message: 'このキャラクターはすでに招待済みです。' };
    }
    if (status !== '未招待' && status !== '予約済') {
      return { ok: false, message: 'このキャラクターは現在予約できません。' };
    }

    const currentDateText = String(sheet.getRange(member.row, CONFIG.COL_DATE).getDisplayValue() || '').trim();
    const currentTimeText = normalizeTime_(String(sheet.getRange(member.row, CONFIG.COL_TIME).getDisplayValue() || '').trim());
    const currentDate = normalizeSheetDate_(currentDateText);
    const currentKey = status === '予約済' && currentDate && currentTimeText
      ? currentDate + ' ' + currentTimeText
      : '';

    const requestedKey = input.date + ' ' + input.time;
    const booked = getBookedSlots_();

    if (booked.includes(requestedKey) && requestedKey !== currentKey) {
      return {
        ok: false,
        code: 'SLOT_TAKEN',
        message: 'この時間は先に予約されました。別の時間を選んでください。'
      };
    }

    const dateText = input.date.slice(5).replace('-', '/');
    sheet.getRange(member.row, CONFIG.COL_DATE).setValue(dateText);
    sheet.getRange(member.row, CONFIG.COL_TIME).setValue(input.time);
    sheet.getRange(member.row, CONFIG.COL_STATUS).setValue('予約済');
    sheet.getRange(member.row, CONFIG.COL_MEETING_DC).setValue(input.dc || '');
    updateXNote_(sheet, member.row, input.useX, input.xAccount);

    return {
      ok: true,
      oldKey: currentKey,
      bookedKey: requestedKey
    };
  } finally {
    lock.releaseLock();
  }
}

function cancel_(name) {
  const normalized = normalizeName_(name);
  if (!normalized) {
    return { ok: false, message: 'キャラクターフルネームを入力してください。' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_();
    const member = findMember_(sheet, normalized);

    if (!member) {
      return {
        ok: false,
        message: '名前が間違っています。メンバーリストに登録されているキャラクターフルネームを確認してください。'
      };
    }

    const status = String(sheet.getRange(member.row, CONFIG.COL_STATUS).getDisplayValue() || '').trim();
    if (status !== '予約済') {
      return { ok: false, message: '取り消せる予約がありません。' };
    }

    const dateText = String(sheet.getRange(member.row, CONFIG.COL_DATE).getDisplayValue() || '').trim();
    const timeText = normalizeTime_(String(sheet.getRange(member.row, CONFIG.COL_TIME).getDisplayValue() || '').trim());
    const bookingDate = normalizeSheetDate_(dateText);
    const oldKey = bookingDate && timeText ? bookingDate + ' ' + timeText : '';

    sheet.getRange(member.row, CONFIG.COL_DATE).clearContent();
    sheet.getRange(member.row, CONFIG.COL_TIME).clearContent();
    sheet.getRange(member.row, CONFIG.COL_MEETING_DC).clearContent();
    updateXNote_(sheet, member.row, false, '');
    sheet.getRange(member.row, CONFIG.COL_STATUS).setValue('未招待');

    return { ok: true, oldKey: oldKey };
  } finally {
    lock.releaseLock();
  }
}

function getBookedSlots_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.FIRST_DATA_ROW) return [];

  const rowCount = lastRow - CONFIG.FIRST_DATA_ROW + 1;
  const values = sheet.getRange(CONFIG.FIRST_DATA_ROW, CONFIG.COL_DATE, rowCount, 3).getDisplayValues();
  const set = new Set();

  values.forEach(row => {
    const dateText = String(row[0] || '').trim();
    const timeText = normalizeTime_(String(row[1] || '').trim());
    const status = String(row[2] || '').trim();
    if (!dateText || !timeText) return;
    if (!['予約済','招待済み','登録完了'].includes(status)) return;

    const date = normalizeSheetDate_(dateText);
    if (!date || date < CONFIG.START_DATE || date > CONFIG.END_DATE) return;

    set.add(date + ' ' + timeText);
  });

  return Array.from(set).sort();
}

function findMember_(sheet, normalizedName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.FIRST_DATA_ROW) return null;

  const names = sheet
    .getRange(CONFIG.FIRST_DATA_ROW, CONFIG.COL_NAME, lastRow - CONFIG.FIRST_DATA_ROW + 1, 1)
    .getDisplayValues();

  for (let i = 0; i < names.length; i++) {
    const name = String(names[i][0] || '').trim();
    if (!name) continue;
    if (normalizeName_(name) === normalizedName) {
      return { row: CONFIG.FIRST_DATA_ROW + i, name: name };
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

function normalizeTime_(value) {
  const m = value.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return '';
  return String(Number(m[1])).padStart(2,'0') + ':' + m[2];
}

function normalizeSheetDate_(value) {
  const s = String(value || '').trim();

  // 10/11 / 10/11/2026 / 2026/10/11 / 2026-10-11 を許容
  let m = s.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (m) return '2026-' + String(Number(m[1])).padStart(2,'0') + '-' + String(Number(m[2])).padStart(2,'0');

  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + String(Number(m[1])).padStart(2,'0') + '-' + String(Number(m[2])).padStart(2,'0');

  m = s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (m) return m[1] + '-' + String(Number(m[2])).padStart(2,'0') + '-' + String(Number(m[3])).padStart(2,'0');

  return '';
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

  if (input.date === CONFIG.START_DATE && input.time < CONFIG.FIRST_DAY_START_TIME) {
    throw new Error('10月11日の予約受付は18:00からです。');
  }

  const now = new Date();
  const nowDate = Utilities.formatDate(now, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  const nowTime = Utilities.formatDate(now, CONFIG.TIME_ZONE, 'HH:mm');

  const nowParts = nowDate.split('-').map(Number);
  const nowTimeParts = nowTime.split(':').map(Number);
  const jstNowMs = Date.UTC(
    nowParts[0],
    nowParts[1] - 1,
    nowParts[2],
    nowTimeParts[0],
    nowTimeParts[1]
  );
  const earliestMs = jstNowMs + 60 * 60 * 1000;

  const dateParts = input.date.split('-').map(Number);
  const timeParts = input.time.split(':').map(Number);
  const slotMs = Date.UTC(
    dateParts[0],
    dateParts[1] - 1,
    dateParts[2],
    timeParts[0],
    timeParts[1]
  );

  if (slotMs < earliestMs) {
    throw new Error('予約は現在時刻の1時間後以降を選択してください。');
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

  if (input.useX) {
    const raw = String(input.xAccount || '').trim();
    const handle = raw.startsWith('@') ? raw.slice(1) : raw;
    if (!/^[A-Za-z0-9_]{1,15}$/.test(handle)) {
      throw new Error('X アカウント名が正しくありません。@og_ff14 のように入力してください。');
    }
    input.xAccount = '@' + handle;
  } else {
    input.xAccount = '';
  }
}

function updateXNote_(sheet, row, useX, xAccount) {
  const range = sheet.getRange(row, CONFIG.COL_NOTE);
  const original = String(range.getDisplayValue() || '');
  const lines = original
    .split(/\r?\n/)
    .filter(line => !/^X:\s*@?[A-Za-z0-9_]{1,15}\s*$/.test(line.trim()));

  if (useX && xAccount) {
    lines.push('X: ' + xAccount);
  }

  const next = lines.join('\n').trim();
  if (next) range.setValue(next);
  else range.clearContent();
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('「' + CONFIG.SHEET_NAME + '」シートがありません。');
  return sheet;
}

function sanitizeCallback_(value) {
  const s = String(value || '');
  return /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(s) ? s : '';
}
