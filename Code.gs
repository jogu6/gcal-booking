const CONFIG = {
  APP_VERSION: '2026.10.16.1',
  MEMBER_SPREADSHEET_ID: '1lF_hqTu9Oe2VbH7-CEyKg14942SWgGKpwJ5Z-cGCz04',
  MEMBER_SHEET_NAME: '招待状況',

  RESERVATION_SPREADSHEET_ID: '12mdeeY6y6xWv71CQdhxQCRxSI19oRPpTMkhlIIFX1vo',
  HOLIDAY_SPREADSHEET_ID: '1pkRfYkVzq84vn11vkEZF6YHhMIIyVYgXQtjqO5lRpx0',
  HOLIDAY_SHEET_NAME: '祝日一覧',
  RESERVATION_SHEET_NAME: '予約管理',
  MAINTENANCE_SHEET_NAME: 'メンテナンス',
  SPECIAL_RECEPTION_SHEET_NAME: '特別受付',

  TIME_ZONE: 'Asia/Tokyo',
  START_DATE: '2026-10-11',
  END_DATE: '2026-10-18',
  ROLLING_START_DATE: '2026-10-12',
  ROLLING_DAYS: 7,
  FIRST_DAY_START_TIME: '18:00',
  START_MINUTE: 9 * 60,
  END_MINUTE: 23 * 60,
  SLOT_MINUTES: 30,

  MEMBER_FIRST_DATA_ROW: 4,
  MEMBER_COL_NAME: 2,
  MEMBER_COL_HOME: 3,
  MEMBER_COL_STATUS: 4,

  // 予約管理: A=キャラクター名 B=予約日 C=予約時刻 D=集合DC E=集合ワールド
  // F=対応方法 G=連絡先 H=更新日時 I=対応状況 J=対象者に連絡 K=受付方式 L=備考
  RES_FIRST_DATA_ROW: 2,
  RES_COL_NAME: 1,
  RES_COL_DATE: 2,
  RES_COL_TIME: 3,
  RES_COL_MEETING_DC: 4,
  RES_COL_MEETING_WORLD: 5,
  RES_COL_SUPPORT_METHOD: 6,
  RES_COL_CONTACT: 7,
  RES_COL_UPDATED_AT: 8,
  RES_COL_MODE: 11,
  RES_COL_REMARKS: 12
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
      const specialReceptionSheet = getSheetFrom_(privateSs, CONFIG.SPECIAL_RECEPTION_SHEET_NAME);
      const maintenance = getMaintenancePeriodsFromSheet_(maintenanceSheet);
      const specialReception = getSpecialReceptionPeriodsFromSheet_(specialReceptionSheet);
      let reservations = readReservations_(reservationSheet);
      reservations = reconcileReservationsForMaintenance_(reservationSheet, reservations, maintenance);

      const window = getBookingWindow_();

      result = {
        ok: true,
        booked: getBookedSlotsFromReservations_(reservations, window, specialReception),
        flexibleCounts: getFlexibleSlotCounts_(reservations, window, specialReception),
        maintenance: maintenance.map(publicMaintenance_),
        specialReception: specialReception.map(publicSpecialReception_),
        holidays: getHolidayDates_(window.start, window.end),
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
        world: e.parameter.world || '',
        supportMethod: e.parameter.supportMethod || (e.parameter.useX === '1' ? 'x' : 'game'),
        contact: e.parameter.contact || e.parameter.xAccount || '',
        language: e.parameter.lang === 'en' ? 'en' : 'ja'
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

  const homeParts = member.home.split('/');
  const homeDc = homeParts[0] || '';
  const homeWorld = homeParts.length > 1 ? homeParts.slice(1).join('/').trim() : '';

  const reservation = findReservationInRows_(readReservations_(getReservationSheet_()), normalized);
  const hasBooking = !!reservation && !!reservation.date && !!reservation.time;

  return {
    ok: true,
    found: true,
    dc: ['Elemental','Gaia','Mana','Meteor'].includes(homeDc) ? homeDc : '',
    world: homeWorld,
    meetingDc: hasBooking && ['Elemental','Gaia','Mana','Meteor'].includes(reservation.meetingDc)
      ? reservation.meetingDc
      : '',
    meetingWorld: hasBooking ? reservation.meetingWorld || '' : '',
    eligible: member.status === '招待日ご予約待ち',
    sourceStatus: member.status,
    canCancel: hasBooking,
    bookingDate: hasBooking ? reservation.date : '',
    bookingTime: hasBooking ? reservation.time : '',
    bookingKey: hasBooking ? reservation.date + ' ' + reservation.time : '',
    supportMethod: hasBooking ? reservation.supportMethod : '',
    contact: hasBooking ? reservation.contact : '',
    useX: hasBooking ? reservation.supportMethod === 'x' : false,
    xAccount: hasBooking && reservation.supportMethod === 'x' ? reservation.contact : ''
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
    if (status !== '招待日ご予約待ち') {
      return { ok: false, message: 'このキャラクターは現在予約できません。' };
    }

    const privateSs = getReservationSpreadsheet_();
    const reservationSheet = getSheetFrom_(privateSs, CONFIG.RESERVATION_SHEET_NAME);
    const maintenanceSheet = getSheetFrom_(privateSs, CONFIG.MAINTENANCE_SHEET_NAME);
    const specialReceptionSheet = getSheetFrom_(privateSs, CONFIG.SPECIAL_RECEPTION_SHEET_NAME);
    const maintenance = getMaintenancePeriodsFromSheet_(maintenanceSheet);
    const specialReception = getSpecialReceptionPeriodsFromSheet_(specialReceptionSheet);

    const maintenanceState = getMaintenanceStateForSlot_(input.date, input.time, maintenance);

    if (maintenanceState.closed) {
      return {
        ok: false,
        code: 'MAINTENANCE_CLOSED',
        message: 'この時間帯は受付を停止しています。'
      };
    }

    if (!isSupportMethodAvailableForMaintenance_(input.supportMethod, maintenanceState)) {
      if (input.supportMethod === 'game' && maintenanceState.blockGame) {
        return {
          ok: false,
          code: 'MAINTENANCE_FF14',
          message: 'この時間帯はゲーム内対応を受け付けていません。'
        };
      }
      if (input.supportMethod === 'x' && maintenanceState.blockX) {
        return {
          ok: false,
          code: 'MAINTENANCE_X',
          message: 'この時間帯はX対応を受け付けていません。'
        };
      }
      if (input.supportMethod === 'companion' && maintenanceState.blockCompanion) {
        return {
          ok: false,
          code: 'MAINTENANCE_COMPANION',
          message: 'この時間帯はコンパニオンアプリ対応を受け付けていません。'
        };
      }
    }

    const specialState = getSpecialReceptionStateForSlot_(
      input.date,
      input.time,
      specialReception
    );

    if (specialState && specialState.state === 'buffer') {
      return {
        ok: false,
        code: 'SPECIAL_BUFFER',
        message: 'この時間帯は随時受付から通常予約へ切り替えるためのアディショナルタイムです。'
      };
    }

    const reservations = readReservations_(reservationSheet);
    const current = findReservationInRows_(reservations, normalizeName_(input.name));
    const currentKey = current && current.date && current.time
      ? current.date + ' ' + current.time
      : '';

    const requestedKey = input.date + ' ' + input.time;
    const isFlexibleSlot = !!specialState && specialState.state === 'flexible';

    const requestedMode =
      current &&
      currentKey === requestedKey &&
      current.mode === 'normal'
        ? 'normal'
        : (isFlexibleSlot ? 'flexible' : 'normal');

    const collision = reservations.some(reservation => {
      if (current && reservation.row === current.row) return false;
      if (reservation.date !== input.date || reservation.time !== input.time) return false;

      if (requestedMode === 'flexible' && reservation.mode === 'flexible') {
        return false;
      }

      return true;
    });

    if (collision) {
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
      meetingWorld: input.world || '',
      supportMethod: input.supportMethod,
      contact: input.contact || '',
      useX: input.supportMethod === 'x',
      xAccount: input.supportMethod === 'x' ? (input.contact || '') : '',
      mode: requestedMode,
      language: input.language === 'en' ? 'en' : 'ja',
      remarks: current ? (current.remarks || '') : ''
    };

    upsertReservation_(reservationSheet, current, savedReservation);

    if (shouldNotifyAllReservationEvents_()) {
      try {
        if (current) {
          if (hasReservationChanged_(current, savedReservation)) {
            notifyDiscordReservationChanged_(current, savedReservation);
          }
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
      bookedKey: requestedKey,
      mode: requestedMode
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
  const privateSs = getReservationSpreadsheet_();
  const reservationSheet = getSheetFrom_(privateSs, CONFIG.RESERVATION_SHEET_NAME);
  const specialReceptionSheet = getSheetFrom_(privateSs, CONFIG.SPECIAL_RECEPTION_SHEET_NAME);

  return getBookedSlotsFromReservations_(
    readReservations_(reservationSheet),
    getBookingWindow_(),
    getSpecialReceptionPeriodsFromSheet_(specialReceptionSheet)
  );
}

function readReservations_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.RES_FIRST_DATA_ROW) return [];

  const rowCount = lastRow - CONFIG.RES_FIRST_DATA_ROW + 1;
  const values = sheet
    .getRange(
      CONFIG.RES_FIRST_DATA_ROW,
      1,
      rowCount,
      Math.max(CONFIG.RES_COL_MODE, CONFIG.RES_COL_REMARKS)
    )
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
      meetingDc: (() => {
        const dcValue = String(values[i][3] || '').trim();
        return dcValue === '(DC指定省略)' ? '' : dcValue;
      })(),
      meetingWorld: String(values[i][4] || '').trim(),
      supportMethod: (() => {
        const methodValue = String(values[i][5] || '').trim();
        if (methodValue === 'X' || methodValue.toLowerCase() === 'true' || methodValue === 'Xでやりとりを希望' || methodValue === '希望します') return 'x';
        if (methodValue === 'コンパニオンアプリ') return 'companion';
        return 'game';
      })(),
      contact: String(values[i][6] || '').trim(),
      useX: (() => {
        const methodValue = String(values[i][5] || '').trim();
        return methodValue === 'X' || methodValue.toLowerCase() === 'true' || methodValue === 'Xでやりとりを希望' || methodValue === '希望します';
      })(),
      xAccount: String(values[i][6] || '').trim(),
      updatedAt: String(values[i][7] || '').trim(),
      mode: String(values[i][10] || '').trim() === '随時受付' ? 'flexible' : 'normal',
      remarks: String(values[i][11] || '').trim()
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

function getBookedSlotsFromReservations_(reservations, window, specialReception) {
  const set = new Set();
  const activeWindow = window || getBookingWindow_();
  const special = specialReception || [];

  reservations.forEach(reservation => {
    if (!reservation.date || !reservation.time) return;
    if (reservation.date < activeWindow.start || reservation.date > activeWindow.end) return;

    const state = getSpecialReceptionStateForSlot_(
      reservation.date,
      reservation.time,
      special
    );

    if (state && state.state === 'flexible' && reservation.mode === 'flexible') {
      return;
    }

    set.add(reservation.date + ' ' + reservation.time);
  });

  return Array.from(set).sort();
}

function getFlexibleSlotCounts_(reservations, window, specialReception) {
  const counts = {};
  const activeWindow = window || getBookingWindow_();
  const special = specialReception || [];

  reservations.forEach(reservation => {
    if (!reservation.date || !reservation.time) return;
    if (reservation.date < activeWindow.start || reservation.date > activeWindow.end) return;
    if (reservation.mode !== 'flexible') return;

    const state = getSpecialReceptionStateForSlot_(
      reservation.date,
      reservation.time,
      special
    );

    if (!state || state.state !== 'flexible') return;

    const key = reservation.date + ' ' + reservation.time;
    counts[key] = (counts[key] || 0) + 1;
  });

  return counts;
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
  return true;
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

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const periods = [];

  values.forEach(row => {
    if (row[0] !== true) return;
    if (!(row[1] instanceof Date) || !(row[2] instanceof Date)) return;

    const startMs = dateToJstPseudoMs_(row[1]);
    const endMs = dateToJstPseudoMs_(row[2]);
    if (!Number.isFinite(startMs) || !Number.isFinite(endMs) || endMs <= startMs) return;

    const blockGame = row[4] === true;
    const blockX = row[5] === true;
    const blockCompanion = row[6] === true;

    if (!blockGame && !blockX && !blockCompanion) return;

    periods.push({
      startMs: startMs,
      endMs: endMs,
      start: Utilities.formatDate(row[1], CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm"),
      end: Utilities.formatDate(row[2], CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm"),
      message: String(row[3] || '').trim(),
      blockGame: blockGame,
      blockX: blockX,
      blockCompanion: blockCompanion
    });
  });

  periods.sort((a,b) => a.startMs - b.startMs);
  return periods;
}

function publicMaintenance_(period) {
  return {
    start: period.start,
    end: period.end,
    message: period.message || '',
    blockGame: !!period.blockGame,
    blockX: !!period.blockX,
    blockCompanion: !!period.blockCompanion
  };
}

function getSpecialReceptionPeriodsFromSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const periods = [];

  values.forEach(row => {
    if (row[0] !== true) return;
    if (!(row[1] instanceof Date) || !(row[2] instanceof Date) || !(row[3] instanceof Date)) return;

    const startMs = dateToJstPseudoMs_(row[1]);
    const cutoffMs = dateToJstPseudoMs_(row[2]);
    const normalStartMs = dateToJstPseudoMs_(row[3]);

    if (
      !Number.isFinite(startMs) ||
      !Number.isFinite(cutoffMs) ||
      !Number.isFinite(normalStartMs) ||
      cutoffMs <= startMs ||
      normalStartMs < cutoffMs
    ) {
      return;
    }

    periods.push({
      startMs: startMs,
      cutoffMs: cutoffMs,
      normalStartMs: normalStartMs,
      start: Utilities.formatDate(row[1], CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm"),
      cutoff: Utilities.formatDate(row[2], CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm"),
      normalStart: Utilities.formatDate(row[3], CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm"),
      message: String(row[4] || '').trim()
    });
  });

  periods.sort((a,b) => a.startMs - b.startMs);
  return periods;
}

function publicSpecialReception_(period) {
  return {
    start: period.start,
    cutoff: period.cutoff,
    normalStart: period.normalStart,
    message: period.message || ''
  };
}

function getSpecialReceptionStateForSlot_(date, time, periods) {
  const slotStart = slotStartPseudoMs_(date, time);
  if (!Number.isFinite(slotStart)) return null;

  const matches = [];

  periods.forEach(period => {
    if (slotStart >= period.startMs && slotStart < period.cutoffMs) {
      matches.push({ period: period, state: 'flexible' });
    } else if (slotStart >= period.cutoffMs && slotStart < period.normalStartMs) {
      matches.push({ period: period, state: 'buffer' });
    }
  });

  if (!matches.length) return null;

  const buffer = matches.find(match => match.state === 'buffer');
  return buffer || matches[0];
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

function getMaintenanceMatchesForSlot_(date, time, maintenance) {
  const slotStart = slotStartPseudoMs_(date, time);
  const slotEnd = slotStart + CONFIG.SLOT_MINUTES * 60 * 1000;

  return maintenance.filter(period =>
    slotStart < period.endMs && slotEnd > period.startMs
  );
}

function getMaintenanceStateForSlot_(date, time, maintenance) {
  const matches = getMaintenanceMatchesForSlot_(date, time, maintenance);
  const blockGame = matches.some(period => period.blockGame);
  const blockX = matches.some(period => period.blockX);
  const blockCompanion = matches.some(period => period.blockCompanion);

  return {
    matches: matches,
    blockGame: blockGame,
    blockX: blockX,
    blockCompanion: blockCompanion,
    closed: blockGame && blockX && blockCompanion
  };
}

function getMaintenanceForSlot_(date, time, maintenance) {
  const state = getMaintenanceStateForSlot_(date, time, maintenance);
  return state.matches.length ? state.matches[0] : null;
}

function isSlotInMaintenance_(date, time, maintenance) {
  return getMaintenanceStateForSlot_(date, time, maintenance).matches.length > 0;
}

function isSupportMethodAvailableForMaintenance_(method, state) {
  if (method === 'game') return !state.blockGame;
  if (method === 'x') return !state.blockX;
  if (method === 'companion') return !state.blockCompanion;
  return false;
}

function reconcileReservationsForMaintenance_(sheet, reservations, maintenance) {
  if (!maintenance.length || !reservations.length) return reservations;

  const removeRows = [];
  const keep = [];
  const autoCancelled = [];

  reservations.forEach(reservation => {
    const state =
      reservation.date && reservation.time
        ? getMaintenanceStateForSlot_(reservation.date, reservation.time, maintenance)
        : { matches: [], closed: false, ff14: false, companion: false };

    const shouldCancel =
      state.matches.length > 0 &&
      !isSupportMethodAvailableForMaintenance_(reservation.supportMethod || 'game', state);

    if (shouldCancel) {
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
      value: reservation.meetingDc || '(DC指定省略)',
      inline: false
    },
    {
      name: '集合ワールド',
      value: reservation.meetingWorld || '(ワールド指定省略)',
      inline: false
    },
    {
      name: '対応方法',
      value: reservation.supportMethod === 'x'
        ? 'X'
        : (reservation.supportMethod === 'companion' ? 'コンパニオンアプリ' : 'ゲーム内'),
      inline: false
    },
    {
      name: '連絡先',
      value: reservation.supportMethod === 'x' && reservation.contact ? reservation.contact : 'なし',
      inline: false
    }
  ];
}


function hasReservationChanged_(before, after) {
  return (
    String(before.date || '') !== String(after.date || '') ||
    String(before.time || '') !== String(after.time || '') ||
    String(before.meetingDc || '') !== String(after.meetingDc || '') ||
    String(before.meetingWorld || '') !== String(after.meetingWorld || '') ||
    String(before.supportMethod || 'game') !== String(after.supportMethod || 'game') ||
    String(before.contact || '') !== String(after.contact || '') ||
    String(before.mode || 'normal') !== String(after.mode || 'normal')
  );
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
        value: after.meetingDc || '(DC指定省略)',
        inline: false
      },
      {
        name: '集合ワールド',
        value: after.meetingWorld || '(ワールド指定省略)',
        inline: false
      },
      {
        name: '対応方法',
        value: after.supportMethod === 'x'
          ? 'X'
          : (after.supportMethod === 'companion' ? 'コンパニオンアプリ' : 'ゲーム内'),
        inline: false
      },
      {
        name: '連絡先',
        value: after.supportMethod === 'x' && after.contact ? after.contact : 'なし',
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
    const state = getMaintenanceStateForSlot_(
      reservation.date,
      reservation.time,
      maintenance
    );

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
        value: reservation.meetingDc || '(DC指定省略)',
        inline: false
      },
      {
        name: '集合ワールド',
        value: reservation.meetingWorld || '(ワールド指定省略)',
        inline: false
      },
      {
        name: '対応方法',
        value: reservation.supportMethod === 'x'
          ? 'X'
          : (reservation.supportMethod === 'companion' ? 'コンパニオンアプリ' : 'ゲーム内'),
        inline: false
      }
    ];

    if (state.matches.length) {
      const serviceLabels = [];
      if (state.blockGame) serviceLabels.push('ゲーム内停止');
      if (state.blockX) serviceLabels.push('X停止');
      if (state.blockCompanion) serviceLabels.push('コンパニオン停止');

      fields.push({
        name: '停止中の対応方法',
        value: serviceLabels.join(' / '),
        inline: false
      });

      const periodText = state.matches
        .map(period =>
          period.start.replace('T',' ') + ' ～ ' + period.end.replace('T',' ') +
          (period.message ? ' / ' + period.message : '')
        )
        .join('\n');

      fields.push({
        name: '該当時間',
        value: periodText,
        inline: false
      });
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
    data.meetingDc || '(DC指定省略)',
    data.meetingWorld || '',
    data.supportMethod === 'x'
      ? 'X'
      : (data.supportMethod === 'companion' ? 'コンパニオンアプリ' : 'ゲーム内'),
    data.supportMethod === 'x' ? (data.contact || '') : '',
    updatedAt
  ]];

  const targetRow = existing
    ? existing.row
    : Math.max(sheet.getLastRow() + 1, CONFIG.RES_FIRST_DATA_ROW);

  sheet.getRange(targetRow, 1, 1, 8).setValues(rowValues);
  sheet
    .getRange(targetRow, CONFIG.RES_COL_MODE)
    .setValue(data.mode === 'flexible' ? '随時受付' : '通常予約');

  let remarks = String(data.remarks || '').trim();
  if (data.language === 'en') {
    const note = '英語表示から予約';
    if (!remarks.includes(note)) {
      remarks = remarks ? remarks + ' / ' + note : note;
    }
  }
  sheet
    .getRange(targetRow, CONFIG.RES_COL_REMARKS)
    .setValue(remarks);

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
  const earliestMs = jstNowMs + 90 * 60 * 1000;

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
    throw new Error('予約は現在時刻の1時間30分後以降を選択してください。');
  }

  const normalized = normalizeName_(input.name);
  if (!/^[a-z][a-z'-]* [a-z][a-z'-]*$/i.test(normalized)) {
    throw new Error('キャラクターフルネームを「Aaaa Bbbb」の形式で入力してください。');
  }

  const allowedMethods = ['game', 'x', 'companion'];
  if (!allowedMethods.includes(input.supportMethod)) {
    throw new Error('招待・登録サポート方法を選択してください。');
  }

  const allowedDc = ['Elemental', 'Gaia', 'Mana', 'Meteor'];
  const privateSs = getReservationSpreadsheet_();
  const maintenanceSheet = getSheetFrom_(privateSs, CONFIG.MAINTENANCE_SHEET_NAME);
  const maintenance = getMaintenancePeriodsFromSheet_(maintenanceSheet);
  const maintenanceState = getMaintenanceStateForSlot_(input.date, input.time, maintenance);
  const needsDc =
    input.supportMethod === 'game' &&
    !maintenanceState.closed &&
    !maintenanceState.blockGame &&
    (input.date === '2026-10-11' || input.date === '2026-10-12');

  if (needsDc && !allowedDc.includes(input.dc)) {
    throw new Error('ゲーム内対応では、集合希望DCを選択してください。');
  }

  const needsWorld =
    needsDc &&
    input.date === '2026-10-12' &&
    input.time <= '18:30';

  if (needsWorld && !String(input.world || '').trim()) {
    throw new Error('ゲーム内対応では、集合希望ワールドを選択してください。');
  }

  if (!needsDc) input.dc = '';
  if (!needsWorld) input.world = '';

  if (input.supportMethod === 'x') {
    const raw = String(input.contact || '').trim();
    const handle = raw.startsWith('@') ? raw.slice(1) : raw;
    if (!/^[A-Za-z0-9_]{1,15}$/.test(handle)) {
      throw new Error('X アカウント名が正しくありません。@og_ff14 のように入力してください。');
    }
    input.contact = '@' + handle;
  } else {
    input.contact = '';
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

// 祝日の唯一の参照元は専用マスタースプレッドシート。
// 祝日マスターは日付昇順。受付期間に交差する行だけ読み取る。
function getHolidayDates_(startDate, endDate) {
  const sheet = SpreadsheetApp.openById(CONFIG.HOLIDAY_SPREADSHEET_ID)
    .getSheetByName(CONFIG.HOLIDAY_SHEET_NAME);
  if (!sheet) throw new Error('祝日マスターの「祝日一覧」シートがありません。');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // 日付列は昇順であることが前提。範囲外の過去分を一括取得しない。
  // TextFinderで受付開始日以降の最初の行を探す。
  const range = sheet.getRange(2, 1, lastRow - 1, 1);
  const first = range.createTextFinder('^' + startDate.replace(/-/g, '-') + '$')
    .useRegularExpression(true).matchEntireCell(true).findNext();
  let firstRow;
  if (first) {
    firstRow = first.getRow();
  } else {
    // 祝日が開始日と一致しない場合は、日付順の列を二分探索する。
    let low = 2, high = lastRow + 1;
    while (low < high) {
      const mid = Math.floor((low + high) / 2);
      const value = sheet.getRange(mid, 1).getValue();
      const key = value instanceof Date
        ? Utilities.formatDate(value, CONFIG.TIME_ZONE, 'yyyy-MM-dd')
        : String(value || '').trim();
      if (key < startDate) low = mid + 1;
      else high = mid;
    }
    firstRow = low;
  }
  if (firstRow > lastRow) return [];
  const result = [];
  for (let row = firstRow; row <= lastRow; row++) {
    const value = sheet.getRange(row, 1).getValue();
    const key = value instanceof Date
      ? Utilities.formatDate(value, CONFIG.TIME_ZONE, 'yyyy-MM-dd')
      : String(value || '').trim();
    if (key > endDate) break;
    if (/^\\d{4}-\\d{2}-\\d{2}$/.test(key) && key >= startDate) result.push(key);
  }
  return [...new Set(result)];
}
