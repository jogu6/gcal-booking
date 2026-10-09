const CONFIG = {
  APP_VERSION: '2026.10.09.3',
  MEMBER_SPREADSHEET_ID: '1lF_hqTu9Oe2VbH7-CEyKg14942SWgGKpwJ5Z-cGCz04',
  MEMBER_SHEET_NAME: '招待状況',

  RESERVATION_SPREADSHEET_ID: '12mdeeY6y6xWv71CQdhxQCRxSI19oRPpTMkhlIIFX1vo',
  RESERVATION_SHEET_NAME: '予約管理',
  MAINTENANCE_SHEET_NAME: 'メンテナンス',

  TIME_ZONE: 'Asia/Tokyo',
  START_DATE: '2026-10-11',
  END_DATE: '2026-10-18',
  ROLLING_START_DATE: '2026-10-12',
  NOTIFY_ALL_FROM_DATE: '2026-10-19',
  ROLLING_DAYS: 7,
  FIRST_DAY_START_TIME: '18:00',
  START_MINUTE: 8 * 60,
  END_MINUTE: 24 * 60,
  SLOT_MINUTES: 30,

  MEMBER_FIRST_DATA_ROW: 4,
  MEMBER_COL_NAME: 2,
  MEMBER_COL_HOME: 3,
  MEMBER_COL_STATUS: 4,

  // 予約管理: A=キャラクター名 B=予約日 C=予約時刻 D=集合DC
  // E=X利用 F=Xアカウント名 G=更新日時
  RES_FIRST_DATA_ROW: 2,
  RES_COL_NAME: 1,
  RES_COL_DATE: 2,
  RES_COL_TIME: 3,
  RES_COL_MEETING_DC: 4,
  RES_COL_USE_X: 5,
  RES_COL_X_ACCOUNT: 6,
  RES_COL_UPDATED_AT: 7
};

function doGet(e) {
  const callback = sanitizeCallback_(e.parameter.callback);
  let result;

  try {
    const action = e.parameter.action || 'slots';

    if (action === 'version') {
      result = { ok: true, version: CONFIG.APP_VERSION };
    } else if (action === 'slots') {
      const privateSs = getReservationSpreadsheet_();
      const reservationSheet = getSheetFrom_(privateSs, CONFIG.RESERVATION_SHEET_NAME);
      const maintenanceSheet = getSheetFrom_(privateSs, CONFIG.MAINTENANCE_SHEET_NAME);
      const maintenance = getMaintenancePeriodsFromSheet_(maintenanceSheet);
      let reservations = readReservations_(reservationSheet);
      reservations = reconcileReservationsForMaintenance_(reservationSheet, reservations, maintenance);

      const window = getBookingWindow_();

      result = {
        ok: true,
        booked: getBookedSlotsFromReservations_(reservations, window),
        maintenance: maintenance.map(publicMaintenance_),
        windowStart: window.start,
        windowEnd: window.end,
        version: CONFIG.APP_VERSION
      };
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

  if (result && typeof result === 'object' && !('version' in result)) {
    result.version = CONFIG.APP_VERSION;
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

  const member = findMember_(getMemberSheet_(), normalized);
  if (!member) return { ok: true, found: false, dc: '' };

  const homeDc = member.home.includes('/') ? member.home.split('/')[0] : member.home;

  const reservation = findReservationInRows_(readReservations_(getReservationSheet_()), normalized);
  const hasBooking = !!reservation && !!reservation.date && !!reservation.time;

  return {
    ok: true,
    found: true,
    dc: ['Elemental','Gaia','Mana','Meteor'].includes(homeDc) ? homeDc : '',
    meetingDc: hasBooking && ['Elemental','Gaia','Mana','Meteor'].includes(reservation.meetingDc)
      ? reservation.meetingDc
      : '',
    eligible: member.status === '未招待',
    sourceStatus: member.status,
    canCancel: hasBooking,
    bookingDate: hasBooking ? reservation.date : '',
    bookingTime: hasBooking ? reservation.time : '',
    bookingKey: hasBooking ? reservation.date + ' ' + reservation.time : '',
    useX: hasBooking ? reservation.useX : false,
    xAccount: hasBooking ? reservation.xAccount : ''
  };
}

function book_(input) {
  validateBooking_(input);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const memberSheet = getMemberSheet_();
    const member = findMember_(memberSheet, normalizeName_(input.name));

    if (!member) {
      return {
        ok: false,
        message: '名前が間違っています。メンバーリストに登録されているキャラクターフルネームを確認してください。'
      };
    }

    const status = member.status;

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
    if (status !== '未招待') {
      return { ok: false, message: 'このキャラクターは現在予約できません。' };
    }

    const privateSs = getReservationSpreadsheet_();
    const reservationSheet = getSheetFrom_(privateSs, CONFIG.RESERVATION_SHEET_NAME);
    const maintenanceSheet = getSheetFrom_(privateSs, CONFIG.MAINTENANCE_SHEET_NAME);
    const maintenance = getMaintenancePeriodsFromSheet_(maintenanceSheet);

    if (isSlotInMaintenance_(input.date, input.time, maintenance) && !input.useX) {
      return {
        ok: false,
        code: 'MAINTENANCE_X_REQUIRED',
        message: 'メンテナンス時間帯は、Xでのやりとりを希望する予約のみ受け付けています。'
      };
    }

    const reservations = readReservations_(reservationSheet);
    const current = findReservationInRows_(reservations, normalizeName_(input.name));
    const currentKey = current && current.date && current.time
      ? current.date + ' ' + current.time
      : '';

    const requestedKey = input.date + ' ' + input.time;
    const booked = getBookedSlotsFromReservations_(reservations, getBookingWindow_());

    if (booked.includes(requestedKey) && requestedKey !== currentKey) {
      return {
        ok: false,
        code: 'SLOT_TAKEN',
        message: 'この時間は先に予約されました。別の時間を選んでください。'
      };
    }

    const savedReservation = {
      name: member.name,
      date: input.date,
      time: input.time,
      meetingDc: input.dc || '',
      useX: !!input.useX,
      xAccount: input.xAccount || ''
    };

    upsertReservation_(reservationSheet, current, savedReservation);

    if (shouldNotifyAllReservationEvents_()) {
      try {
        if (current) {
          notifyDiscordReservationChanged_(current, savedReservation);
        } else {
          notifyDiscordReservationCreated_(savedReservation);
        }
      } catch (err) {
        console.error('Discord予約通知に失敗しました: ' + err);
      }
    }

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
    const reservationSheet = getReservationSheet_();
    const reservation = findReservationInRows_(readReservations_(reservationSheet), normalized);

    if (!reservation || !reservation.date || !reservation.time) {
      return { ok: false, message: '取り消せる予約がありません。' };
    }

    const oldKey = reservation.date + ' ' + reservation.time;
    reservationSheet.deleteRow(reservation.row);

    if (shouldNotifyAllReservationEvents_()) {
      try {
        notifyDiscordReservationCancelled_(reservation);
      } catch (err) {
        console.error('Discord予約取消通知に失敗しました: ' + err);
      }
    }

    return { ok: true, oldKey: oldKey };
  } finally {
    lock.releaseLock();
  }
}

function getBookedSlots_() {
  return getBookedSlotsFromReservations_(
    readReservations_(getReservationSheet_()),
    getBookingWindow_()
  );
}

function readReservations_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.RES_FIRST_DATA_ROW) return [];

  const rowCount = lastRow - CONFIG.RES_FIRST_DATA_ROW + 1;
  const values = sheet
    .getRange(CONFIG.RES_FIRST_DATA_ROW, 1, rowCount, 7)
    .getDisplayValues();

  const rows = [];

  for (let i = 0; i < values.length; i++) {
    const name = String(values[i][0] || '').trim();
    if (!name) continue;

    rows.push({
      row: CONFIG.RES_FIRST_DATA_ROW + i,
      name: name,
      normalizedName: normalizeName_(name),
      date: normalizeSheetDate_(String(values[i][1] || '').trim()),
      time: normalizeTime_(String(values[i][2] || '').trim()),
      meetingDc: String(values[i][3] || '').trim(),
      useX: (() => {
        const xValue = String(values[i][4] || '').trim();
        return xValue.toLowerCase() === 'true' || xValue === 'Xでやりとりを希望';
      })(),
      xAccount: String(values[i][5] || '').trim(),
      updatedAt: String(values[i][6] || '').trim()
    });
  }

  return rows;
}

function findReservationInRows_(reservations, normalizedName) {
  for (let i = 0; i < reservations.length; i++) {
    if (reservations[i].normalizedName === normalizedName) return reservations[i];
  }
  return null;
}

function getBookedSlotsFromReservations_(reservations, window) {
  const set = new Set();
  const activeWindow = window || getBookingWindow_();

  reservations.forEach(reservation => {
    if (!reservation.date || !reservation.time) return;
    if (reservation.date < activeWindow.start || reservation.date > activeWindow.end) return;
    set.add(reservation.date + ' ' + reservation.time);
  });

  return Array.from(set).sort();
}

function getBookingWindow_() {
  const today = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd');

  if (today < CONFIG.ROLLING_START_DATE) {
    return {
      start: CONFIG.START_DATE,
      end: CONFIG.END_DATE
    };
  }

  return {
    start: today,
    end: addDaysToDateKey_(today, CONFIG.ROLLING_DAYS - 1)
  };
}

function addDaysToDateKey_(dateKey, days) {
  const parts = dateKey.split('-').map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + days));
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
}

function shouldNotifyAllReservationEvents_() {
  const today = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  return today >= CONFIG.NOTIFY_ALL_FROM_DATE;
}

function findMember_(sheet, normalizedName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.MEMBER_FIRST_DATA_ROW) return null;

  const rowCount = lastRow - CONFIG.MEMBER_FIRST_DATA_ROW + 1;
  const values = sheet
    .getRange(
      CONFIG.MEMBER_FIRST_DATA_ROW,
      CONFIG.MEMBER_COL_NAME,
      rowCount,
      3
    )
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const name = String(values[i][0] || '').trim();
    if (!name) continue;
    if (normalizeName_(name) !== normalizedName) continue;

    return {
      row: CONFIG.MEMBER_FIRST_DATA_ROW + i,
      name: name,
      home: String(values[i][1] || '').trim(),
      status: String(values[i][2] || '').trim()
    };
  }
  return null;
}

function getMaintenancePeriodsFromSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const periods = [];

  values.forEach(row => {
    if (row[0] !== true) return;
    if (!(row[1] instanceof Date) || !(row[2] instanceof Date)) return;

    const startMs = dateToJstPseudoMs_(row[1]);
    const endMs = dateToJstPseudoMs_(row[2]);
    if (!Number.isFinite(startMs) || !Number.isFinite(endMs) || endMs <= startMs) return;

    periods.push({
      startMs: startMs,
      endMs: endMs,
      start: Utilities.formatDate(row[1], CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm"),
      end: Utilities.formatDate(row[2], CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm"),
      message: String(row[3] || '').trim()
    });
  });

  periods.sort((a,b) => a.startMs - b.startMs);
  return periods;
}

function publicMaintenance_(period) {
  return {
    start: period.start,
    end: period.end,
    message: period.message || ''
  };
}

function dateToJstPseudoMs_(date) {
  const text = Utilities.formatDate(date, CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm');
  return localStampToPseudoMs_(text);
}

function localStampToPseudoMs_(value) {
  const m = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})$/);
  if (!m) return NaN;
  return Date.UTC(
    Number(m[1]),
    Number(m[2]) - 1,
    Number(m[3]),
    Number(m[4]),
    Number(m[5])
  );
}

function slotStartPseudoMs_(date, time) {
  return localStampToPseudoMs_(date + ' ' + time);
}

function isSlotInMaintenance_(date, time, maintenance) {
  const slotStart = slotStartPseudoMs_(date, time);
  const slotEnd = slotStart + CONFIG.SLOT_MINUTES * 60 * 1000;

  return maintenance.some(period =>
    slotStart < period.endMs && slotEnd > period.startMs
  );
}

function reconcileReservationsForMaintenance_(sheet, reservations, maintenance) {
  if (!maintenance.length || !reservations.length) return reservations;

  const removeRows = [];
  const keep = [];
  const autoCancelled = [];

  reservations.forEach(reservation => {
    const overlaps =
      reservation.date &&
      reservation.time &&
      isSlotInMaintenance_(reservation.date, reservation.time, maintenance);

    if (overlaps && !reservation.useX) {
      removeRows.push(reservation.row);
      autoCancelled.push(reservation);
    } else {
      keep.push(reservation);
    }
  });

  removeRows.sort((a,b) => b-a).forEach(row => sheet.deleteRow(row));

  if (autoCancelled.length) {
    notifyDiscordAutoCancelled_(autoCancelled, maintenance);
  }

  return keep;
}

function sendDiscordReservationEmbed_(embed) {
  const webhookUrl = PropertiesService
    .getScriptProperties()
    .getProperty('DISCORD_WEBHOOK_URL');

  if (!webhookUrl) return;

  const response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ embeds: [embed] }),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error(
      'Discord Webhook通知に失敗しました。HTTP ' +
      status +
      ': ' +
      response.getContentText()
    );
  }
}

function reservationFields_(reservation) {
  return [
    {
      name: 'キャラクター名',
      value: reservation.name || '不明',
      inline: false
    },
    {
      name: '予約日時',
      value: (reservation.date || '') + ' ' + (reservation.time || ''),
      inline: false
    },
    {
      name: '集合DC',
      value: reservation.meetingDc || '指定なし',
      inline: false
    },
    {
      name: 'X利用',
      value: reservation.useX ? 'あり' : 'なし',
      inline: false
    },
    {
      name: 'Xアカウント名',
      value: reservation.useX && reservation.xAccount ? reservation.xAccount : 'なし',
      inline: false
    }
  ];
}

function notifyDiscordReservationCreated_(reservation) {
  sendDiscordReservationEmbed_({
    title: '新規予約',
    fields: reservationFields_(reservation)
  });
}

function notifyDiscordReservationChanged_(before, after) {
  sendDiscordReservationEmbed_({
    title: '予約変更',
    fields: [
      {
        name: 'キャラクター名',
        value: after.name || before.name || '不明',
        inline: false
      },
      {
        name: '変更前',
        value: (before.date || '') + ' ' + (before.time || ''),
        inline: false
      },
      {
        name: '変更後',
        value: (after.date || '') + ' ' + (after.time || ''),
        inline: false
      },
      {
        name: '集合DC',
        value: after.meetingDc || '指定なし',
        inline: false
      },
      {
        name: 'X利用',
        value: after.useX ? 'あり' : 'なし',
        inline: false
      },
      {
        name: 'Xアカウント名',
        value: after.useX && after.xAccount ? after.xAccount : 'なし',
        inline: false
      }
    ]
  });
}

function notifyDiscordReservationCancelled_(reservation) {
  sendDiscordReservationEmbed_({
    title: '予約取り消し',
    fields: reservationFields_(reservation)
  });
}

function notifyDiscordAutoCancelled_(reservations, maintenance) {
  const webhookUrl = PropertiesService
    .getScriptProperties()
    .getProperty('DISCORD_WEBHOOK_URL');

  if (!webhookUrl) return;

  const embeds = reservations.map(reservation => {
    const matchingPeriods = maintenance.filter(period =>
      isSlotInMaintenance_(reservation.date, reservation.time, [period])
    );
    const period = matchingPeriods.length ? matchingPeriods[0] : null;

    const fields = [
      {
        name: 'キャラクター名',
        value: reservation.name || '不明',
        inline: false
      },
      {
        name: '予約日時',
        value: (reservation.date || '') + ' ' + (reservation.time || ''),
        inline: false
      },
      {
        name: '集合DC',
        value: reservation.meetingDc || '指定なし',
        inline: false
      },
      {
        name: 'X利用',
        value: 'なし',
        inline: false
      }
    ];

    if (period) {
      fields.push({
        name: 'メンテナンス時間',
        value: period.start.replace('T',' ') + ' ～ ' + period.end.replace('T',' '),
        inline: false
      });

      if (period.message) {
        fields.push({
          name: '案内',
          value: period.message,
          inline: false
        });
      }
    }

    return {
      title: 'メンテナンスによる予約の自動キャンセル',
      fields: fields,
      footer: {
        text: '予約管理シートから自動削除済み'
      }
    };
  });

  const payload = { embeds: embeds };

  const response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error(
      'Discord Webhook通知に失敗しました。HTTP ' +
      status +
      ': ' +
      response.getContentText()
    );
  }
}

function testDiscordWebhook() {
  const webhookUrl = PropertiesService
    .getScriptProperties()
    .getProperty('DISCORD_WEBHOOK_URL');

  if (!webhookUrl) {
    throw new Error('スクリプトプロパティ DISCORD_WEBHOOK_URL が設定されていません。');
  }

  const payload = {
    embeds: [{
      title: 'Discord Webhook 接続テスト',
      description: '予約システムからDiscordへの通知テストです。',
      footer: {
        text: Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm:ss')
      }
    }]
  };

  const response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error(
      'Discord Webhookテストに失敗しました。HTTP ' +
      status +
      ': ' +
      response.getContentText()
    );
  }

  console.log('Discord Webhookテスト成功: HTTP ' + status);
}

function upsertReservation_(sheet, existing, data) {
  const updatedAt = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
  const rowValues = [[
    data.name,
    data.date,
    data.time,
    data.meetingDc || '',
    data.useX ? 'Xでやりとりを希望' : '',
    data.useX ? data.xAccount : '',
    updatedAt
  ]];

  if (existing) {
    sheet.getRange(existing.row, 1, 1, 7).setValues(rowValues);
  } else {
    const row = Math.max(sheet.getLastRow() + 1, CONFIG.RES_FIRST_DATA_ROW);
    sheet.getRange(row, 1, 1, 7).setValues(rowValues);
  }

  sortReservations_(sheet);
}

function sortReservations_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.RES_FIRST_DATA_ROW) return;

  const rowCount = lastRow - CONFIG.RES_FIRST_DATA_ROW + 1;
  sheet
    .getRange(CONFIG.RES_FIRST_DATA_ROW, 1, rowCount, sheet.getLastColumn())
    .sort([
      { column: CONFIG.RES_COL_DATE, ascending: true },
      { column: CONFIG.RES_COL_TIME, ascending: true }
    ]);
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

  let m = s.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (m) {
    return '2026-' +
      String(Number(m[1])).padStart(2,'0') + '-' +
      String(Number(m[2])).padStart(2,'0');
  }

  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    return m[3] + '-' +
      String(Number(m[1])).padStart(2,'0') + '-' +
      String(Number(m[2])).padStart(2,'0');
  }

  m = s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (m) {
    return m[1] + '-' +
      String(Number(m[2])).padStart(2,'0') + '-' +
      String(Number(m[3])).padStart(2,'0');
  }

  return '';
}

function validateBooking_(input) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(input.date || '')) {
    throw new Error('日付が不正です。');
  }
  if (!/^\d{2}:\d{2}$/.test(input.time || '')) {
    throw new Error('時刻が不正です。');
  }
  const bookingWindow = getBookingWindow_();
  if (input.date < bookingWindow.start || input.date > bookingWindow.end) {
    throw new Error('予約期間外です。');
  }

  const [h, m] = input.time.split(':').map(Number);
  const minute = h * 60 + m;
  if (
    minute < CONFIG.START_MINUTE ||
    minute >= CONFIG.END_MINUTE ||
    m % CONFIG.SLOT_MINUTES !== 0
  ) {
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

  const allowedDc = ['Elemental', 'Gaia', 'Mana', 'Meteor'];
  const privateSs = getReservationSpreadsheet_();
  const maintenanceSheet = getSheetFrom_(privateSs, CONFIG.MAINTENANCE_SHEET_NAME);
  const maintenance = getMaintenancePeriodsFromSheet_(maintenanceSheet);
  const maintenanceSlot = isSlotInMaintenance_(input.date, input.time, maintenance);
  const needsDc =
    !maintenanceSlot &&
    (input.date === '2026-10-11' || input.date === '2026-10-12');

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

function getMemberSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.MEMBER_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.MEMBER_SHEET_NAME);
  if (!sheet) {
    throw new Error('「' + CONFIG.MEMBER_SHEET_NAME + '」シートがありません。');
  }
  return sheet;
}

function getReservationSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.RESERVATION_SPREADSHEET_ID);
}

function getSheetFrom_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('「' + sheetName + '」シートがありません。');
  return sheet;
}

function getReservationSheet_() {
  return getSheetFrom_(getReservationSpreadsheet_(), CONFIG.RESERVATION_SHEET_NAME);
}

function sanitizeCallback_(value) {
  const s = String(value || '');
  return /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(s) ? s : '';
}
