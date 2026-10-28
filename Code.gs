const CONFIG = {
  APP_VERSION: '2026.10.28.5',
  MEMBER_SPREADSHEET_ID: '1lF_hqTu9Oe2VbH7-CEyKg14942SWgGKpwJ5Z-cGCz04',
  MEMBER_SHEET_NAME: '招待状況',

  RESERVATION_SPREADSHEET_ID: '12mdeeY6y6xWv71CQdhxQCRxSI19oRPpTMkhlIIFX1vo',
  HOLIDAY_SPREADSHEET_ID: '1pkRfYkVzq84vn11vkEZF6YHhMIIyVYgXQtjqO5lRpx0',
  HOLIDAY_SHEET_NAME: '祝日一覧',
  RESERVATION_SHEET_NAME: '予約管理',
  MAINTENANCE_SHEET_NAME: 'メンテナンス',
  SPECIAL_RECEPTION_SHEET_NAME: '特別受付',
  BOOKING_SETTINGS_SHEET_NAME: '予約設定',
  BOOKING_HOURS_SHEET_NAME: '受付時間',

  TIME_ZONE: 'Asia/Tokyo',
  START_DATE: '2026-10-11',
  END_DATE: '2026-10-18',
  ROLLING_START_DATE: '2026-10-12',
  ROLLING_DAYS: 7,
  FIRST_DAY_START_TIME: '18:00',
  START_MINUTE: 9 * 60,
  END_MINUTE: 23 * 60,
  SLOT_MINUTES: 30,

  MEMBER_FIRST_DATA_ROW: 3,
  MEMBER_COL_NAME: 2,
  MEMBER_COL_HOME: 3,
  MEMBER_COL_JOIN_DATE: 4,
  MEMBER_COL_STATUS: 5,
  MEMBER_COL_ONE_HOUR: 6,

  // 予約管理: A=キャラクター名 B=予約日 C=予約時刻 D=集合DC E=対応方法
  // F=連絡先 G=更新日時 H=対応状況 I=対象者に連絡 J=受付方式 K=備考
  RES_FIRST_DATA_ROW: 2,
  RES_COL_NAME: 1,
  RES_COL_DATE: 2,
  RES_COL_TIME: 3,
  RES_COL_MEETING_DC: 4,
  RES_COL_SUPPORT_METHOD: 5,
  RES_COL_CONTACT: 6,
  RES_COL_UPDATED_AT: 7,
  RES_COL_MODE: 10,
  RES_COL_REMARKS: 11
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
      cleanupExpiredMaintenance_(maintenanceSheet);
      const maintenance = getMaintenancePeriodsFromSheet_(maintenanceSheet);
      const specialReception = getSpecialReceptionPeriodsFromSheet_(specialReceptionSheet);
      let reservations = readReservations_(reservationSheet);
      reservations = reconcileReservationsForMaintenance_(reservationSheet, reservations, maintenance);

      const window = getBookingWindow_(privateSs);
      const holidays = getHolidayDates_(window.start, window.end);
      const dailyHours = getDailyBookingHours_(privateSs, window.start, window.end, holidays);

      result = {
        ok: true,
        booked: getBookedSlotsFromReservations_(reservations, window, specialReception, dailyHours),
        flexibleCounts: getFlexibleSlotCounts_(reservations, window, specialReception, dailyHours),
        maintenance: maintenance.map(publicMaintenance_),
        specialReception: specialReception.map(publicSpecialReception_),
        holidays: holidays,
        dailyHours: dailyHours,
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

  const reservation = findReservationInRows_(readReservations_(getReservationSheet_()), normalized);
  const hasBooking = !!reservation && !!reservation.date && !!reservation.time;
  const joinEligibility = getJoinEligibility_(member);

  return {
    ok: true,
    found: true,
    dc: ['Elemental','Gaia','Mana','Meteor'].includes(homeDc) ? homeDc : '',
    meetingDc: hasBooking && ['Elemental','Gaia','Mana','Meteor'].includes(reservation.meetingDc)
      ? reservation.meetingDc
      : '',
    eligible: member.status === 'ご予約待ち/ご予約中' && joinEligibility.allowed,
    code: joinEligibility.code || '',
    message: joinEligibility.message || '',
    bookingFrom: joinEligibility.bookingFrom || '',
    sourceStatus: member.status,
    oneHourBooking: !!member.oneHourBooking,
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

    const joinEligibility = getJoinEligibility_(member);
    if (!joinEligibility.allowed) {
      return {ok: false, code: joinEligibility.code, message: joinEligibility.message, bookingFrom: joinEligibility.bookingFrom || ''};
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
    if (status !== 'ご予約待ち/ご予約中') {
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

    const oneHourBooking = !!member.oneHourBooking;
    const secondTime = oneHourBooking
      ? addMinutesToTime_(input.time, CONFIG.SLOT_MINUTES)
      : '';

    if (oneHourBooking) {
      if (!secondTime || !isWithinBookingHours_(input.bookingHours, input.time, 60)) {
        return {
          ok: false,
          code: 'ONE_HOUR_OUTSIDE_HOURS',
          message: '1時間予約が1つの受付時間帯に収まる開始時刻を選択してください。'
        };
      }

      const secondMaintenanceState = getMaintenanceStateForSlot_(input.date, secondTime, maintenance);
      if (secondMaintenanceState.closed) {
        return {
          ok: false,
          code: 'ONE_HOUR_SECOND_CLOSED',
          message: '1時間予約の後半30分が受付停止のため、この時刻は予約できません。'
        };
      }
      if (!isSupportMethodAvailableForMaintenance_(input.supportMethod, secondMaintenanceState)) {
        const methodLabel = input.supportMethod === 'x'
          ? 'X'
          : (input.supportMethod === 'companion' ? 'コンパニオンアプリ' : 'ゲーム内');
        return {
          ok: false,
          code: 'ONE_HOUR_SECOND_METHOD_UNAVAILABLE',
          message: '1時間予約の後半30分は' + methodLabel + '対応を利用できないため、この時刻は予約できません。'
        };
      }

      const secondSpecialState = getSpecialReceptionStateForSlot_(
        input.date,
        secondTime,
        specialReception
      );
      if (secondSpecialState && secondSpecialState.state === 'buffer') {
        return {
          ok: false,
          code: 'ONE_HOUR_SECOND_BUFFER',
          message: '1時間予約の後半30分がアディショナルタイムのため、この時刻は予約できません。'
        };
      }

      const firstIsFlexible = !!specialState && specialState.state === 'flexible';
      const secondIsFlexible = !!secondSpecialState && secondSpecialState.state === 'flexible';
      if (firstIsFlexible !== secondIsFlexible) {
        return {
          ok: false,
          code: 'ONE_HOUR_MODE_CHANGE',
          message: '1時間予約の途中で受付方式が切り替わるため、この時刻は予約できません。'
        };
      }
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

    let collisionKey = '';
    const collision = reservations.some(reservation => {
      if (current && reservation.row === current.row) return false;
      if (requestedMode === 'flexible' && reservation.mode === 'flexible') return false;
      const overlaps = reservationOverlaps_(reservation, input.date, input.time, oneHourBooking ? 60 : 30);
      if (overlaps) collisionKey = reservation.date + ' ' + reservation.time;
      return overlaps;
    });

    if (collision) {
      const conflictTime = collisionKey ? collisionKey.slice(-5) : input.time;
      return {
        ok: false,
        code: 'SLOT_TAKEN',
        message: oneHourBooking
          ? '1時間予約のため' + conflictTime + '枠も必要ですが、すでに予約されています。別の時刻を選択してください。'
          : 'この時間は先に予約されました。別の時間を選んでください。'
      };
    }

    const savedReservation = {
      name: member.name,
      date: input.date,
      time: input.time,
      meetingDc: input.dc || '',
      supportMethod: input.supportMethod,
      contact: input.contact || '',
      useX: input.supportMethod === 'x',
      xAccount: input.supportMethod === 'x' ? (input.contact || '') : '',
      mode: requestedMode,
      language: input.language === 'en' ? 'en' : 'ja',
      oneHourBooking: oneHourBooking,
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
      mode: requestedMode,
      durationMinutes: oneHourBooking ? 60 : 30
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
      supportMethod: (() => {
        const methodValue = String(values[i][CONFIG.RES_COL_SUPPORT_METHOD - 1] || '').trim();
        if (methodValue === 'X' || methodValue.toLowerCase() === 'true' || methodValue === 'Xでやりとりを希望' || methodValue === '希望します') return 'x';
        if (methodValue === 'コンパニオンアプリ') return 'companion';
        return 'game';
      })(),
      contact: String(values[i][CONFIG.RES_COL_CONTACT - 1] || '').trim(),
      useX: (() => {
        const methodValue = String(values[i][CONFIG.RES_COL_SUPPORT_METHOD - 1] || '').trim();
        return methodValue === 'X' || methodValue.toLowerCase() === 'true' || methodValue === 'Xでやりとりを希望' || methodValue === '希望します';
      })(),
      xAccount: String(values[i][CONFIG.RES_COL_CONTACT - 1] || '').trim(),
      updatedAt: String(values[i][CONFIG.RES_COL_UPDATED_AT - 1] || '').trim(),
      mode: String(values[i][CONFIG.RES_COL_MODE - 1] || '').trim() === '随時受付' ? 'flexible' : 'normal',
      remarks: String(values[i][CONFIG.RES_COL_REMARKS - 1] || '').trim(),
      oneHourBooking: String(values[i][CONFIG.RES_COL_REMARKS - 1] || '').includes('1時間予約')
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

function addMinutesToTime_(time, minutes) {
  const m = String(time || '').match(/^(\d{2}):(\d{2})$/);
  if (!m) return '';
  const total = Number(m[1]) * 60 + Number(m[2]) + Number(minutes || 0);
  if (total < 0 || total >= 24 * 60) return '';
  return String(Math.floor(total / 60)).padStart(2,'0') + ':' +
    String(total % 60).padStart(2,'0');
}

function reservationOccupiedKeys_(reservation) {
  if (!reservation || !reservation.date || !reservation.time) return [];
  const keys = [reservation.date + ' ' + reservation.time];
  if (reservation.oneHourBooking) {
    const secondTime = addMinutesToTime_(reservation.time, CONFIG.SLOT_MINUTES);
    if (secondTime) keys.push(reservation.date + ' ' + secondTime);
  }
  return keys;
}

function getBookedSlotsFromReservations_(reservations, window, specialReception, dailyHours) {
  const set = new Set();
  const activeWindow = window || getBookingWindow_();
  const special = specialReception || [];
  const hours = dailyHours || getDailyBookingHours_(getReservationSpreadsheet_(), activeWindow.start, activeWindow.end);
  Object.keys(hours).forEach(date => {
    candidateBookingTimes_(hours[date]).forEach(time => {
      if (reservations.some(reservation => {
        if (!reservationOverlaps_(reservation, date, time, 30)) return false;
        const state = getSpecialReceptionStateForSlot_(date, time, special);
        return !(state && state.state === 'flexible' && reservation.mode === 'flexible');
      })) set.add(date + ' ' + time);
    });
  });
  return Array.from(set).sort();
}

function getFlexibleSlotCounts_(reservations, window, specialReception, dailyHours) {
  const counts = {};
  const activeWindow = window || getBookingWindow_();
  const special = specialReception || [];
  const hours = dailyHours || getDailyBookingHours_(getReservationSpreadsheet_(), activeWindow.start, activeWindow.end);
  Object.keys(hours).forEach(date => {
    candidateBookingTimes_(hours[date]).forEach(time => {
      const candidateState = getSpecialReceptionStateForSlot_(date, time, special);
      if (!candidateState || candidateState.state !== 'flexible') return;
      const count = reservations.filter(reservation => {
        if (reservation.mode !== 'flexible' || !reservationOverlaps_(reservation, date, time, 30)) return false;
        const state = getSpecialReceptionStateForSlot_(reservation.date, reservation.time, special);
        return state && state.state === 'flexible';
      }).length;
      if (count) counts[date + ' ' + time] = count;
    });
  });
  return counts;
}

function getBookingEndDateOverride_(reservationSs) {
  const ss = reservationSs || getReservationSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.BOOKING_SETTINGS_SHEET_NAME);
  if (!sheet) return '';

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key !== '受付終了日') continue;

    const value = values[i][1];
    if (value instanceof Date && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
    }

    const normalized = normalizeSheetDate_(String(value || '').trim());
    return normalized || '';
  }

  return '';
}

function bookingTimeMinute_(value, allowEndOfDay) {
  const match = String(value || '').trim().match(/^(\d{1,2}):([0-5]\d)$/);
  if (!match) return NaN;
  const minute = Number(match[1]) * 60 + Number(match[2]);
  return minute >= 0 && (minute < 1440 || (allowEndOfDay && minute === 1440))
    ? minute : NaN;
}

function getWeeklyBookingHours_(ss) {
  const keys = ['日', '月', '火', '水', '木', '金', '土', '祝日'];
  const schedule = {};
  const sheet = ss.getSheetByName(CONFIG.BOOKING_HOURS_SHEET_NAME);
  if (!sheet) {
    keys.forEach(key => { schedule[key] = {intervals: [{startMinute: CONFIG.START_MINUTE, endMinute: CONFIG.END_MINUTE}], stepMinutes: 30}; });
    return schedule;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('「受付時間」に曜日・祝日の設定がありません。');
  sheet.getRange(2, 1, lastRow - 1, Math.min(sheet.getMaxColumns(), 8)).getDisplayValues().forEach(row => {
    const key = String(row[0] || '').trim();
    if (row.every(value => !String(value || '').trim())) return;
    if (!keys.includes(key) || schedule[key]) throw new Error('「受付時間」の区分が不正または重複しています：' + key);
    const intervals = [];
    for (let index = 0; index < 3; index++) {
      const start = String(row[1 + index * 2] || '').trim();
      const end = String(row[2 + index * 2] || '').trim();
      if (!start && !end) continue;
      const startMinute = bookingTimeMinute_(start, false);
      const endMinute = bookingTimeMinute_(end, true);
      if (!Number.isFinite(startMinute) || !Number.isFinite(endMinute) || startMinute >= endMinute) {
        throw new Error('「受付時間」の' + key + '・時間帯' + (index + 1) + 'は、開始より後の終了時刻を指定してください。日をまたぐ設定はできません。');
      }
      intervals.push({startMinute: startMinute, endMinute: endMinute});
    }
    intervals.sort((left, right) => left.startMinute - right.startMinute);
    if (intervals.some((value, index) => index > 0 && value.startMinute < intervals[index - 1].endMinute)) {
      throw new Error('「受付時間」の' + key + 'の時間帯が重複しています。');
    }
    const stepMinutes = String(row[7] || '').trim() === '' ? 30 : Number(row[7]);
    if (![15, 30, 60].includes(stepMinutes)) throw new Error('「受付時間」の開始間隔は15・30・60分から指定してください：' + key);
    schedule[key] = {intervals: intervals, stepMinutes: stepMinutes};
  });
  keys.forEach(key => { if (!schedule[key]) throw new Error('「受付時間」に' + key + 'の行がありません。'); });
  return schedule;
}

function getDailyBookingHours_(ss, start, end, holidayDates) {
  const schedule = getWeeklyBookingHours_(ss);
  const holidays = new Set(holidayDates || getHolidayDates_(start, end));
  const hours = {};
  for (let date = start; date <= end; date = addDaysToDateKey_(date, 1)) {
    const parts = date.split('-').map(Number);
    const weekday = ['日', '月', '火', '水', '木', '金', '土'][new Date(Date.UTC(parts[0], parts[1] - 1, parts[2])).getUTCDay()];
    const key = holidays.has(date) ? '祝日' : weekday;
    hours[date] = {intervals: schedule[key].intervals, stepMinutes: schedule[key].stepMinutes, category: key};
  }
  return hours;
}

function isWithinBookingHours_(hours, time, duration) {
  const minute = bookingTimeMinute_(time, false);
  return !!hours && Number.isFinite(minute) && hours.intervals.some(interval =>
    minute >= interval.startMinute && minute + duration <= interval.endMinute &&
    (minute - interval.startMinute) % hours.stepMinutes === 0);
}

function candidateBookingTimes_(hours) {
  const times = [];
  hours.intervals.forEach(interval => {
    for (let minute = interval.startMinute; minute + 30 <= interval.endMinute; minute += hours.stepMinutes) {
      times.push(String(Math.floor(minute / 60)).padStart(2, '0') + ':' + String(minute % 60).padStart(2, '0'));
    }
  });
  return times;
}

function reservationOverlaps_(reservation, date, time, duration) {
  if (reservation.date !== date) return false;
  const start = bookingTimeMinute_(time, false);
  const existing = bookingTimeMinute_(reservation.time, false);
  return Number.isFinite(start) && Number.isFinite(existing) &&
    start < existing + (reservation.oneHourBooking ? 60 : 30) && start + duration > existing;
}

function getBookingWindow_(reservationSs) {
  const today = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd');

  let start;
  let end;

  if (today < CONFIG.ROLLING_START_DATE) {
    start = CONFIG.START_DATE;
    end = CONFIG.END_DATE;
  } else {
    start = today;
    end = addDaysToDateKey_(today, CONFIG.ROLLING_DAYS - 1);
  }

  const configuredEnd = getBookingEndDateOverride_(reservationSs);
  if (configuredEnd && configuredEnd < end) {
    end = configuredEnd;
  }

  return { start: start, end: end };
}

function addDaysToDateKey_(dateKey, days) {
  const parts = dateKey.split('-').map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + days));
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
}

function shouldNotifyAllReservationEvents_() {
  return true;
}

function getJoinEligibility_(member, today) {
  const raw = member.joinDate;
  // Empty cells explicitly mean that at least 30 days have already elapsed.
  if (raw === undefined || raw === null || String(raw).trim() === '') return {allowed: true};
  let date = '';
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    date = Utilities.formatDate(raw, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  } else {
    const match = String(raw).trim().match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
    if (match) {
      const candidate = match[1] + '-' + match[2].padStart(2, '0') + '-' + match[3].padStart(2, '0');
      const parsed = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
      if (Utilities.formatDate(parsed, 'UTC', 'yyyy-MM-dd') === candidate) date = candidate;
    }
  }
  if (!date) return {allowed: false, code: 'MEMBER_JOIN_DATE_INVALID', message: '加入日の記録を確認できません。管理者へご連絡ください。'};
  const bookingFrom = addDaysToDateKey_(date, 30);
  const currentDate = today || Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  if (currentDate >= bookingFrom) return {allowed: true, bookingFrom: bookingFrom};
  const parts = bookingFrom.split('-').map(Number);
  return {
    allowed: false,
    code: 'MEMBER_TOO_NEW',
    bookingFrom: bookingFrom,
    message: '加入から30日が経過していないため、まだ予約できません。' + parts[0] + '年' + parts[1] + '月' + parts[2] + '日から予約できます。'
  };
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
      CONFIG.MEMBER_COL_ONE_HOUR - CONFIG.MEMBER_COL_NAME + 1
    )
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const name = String(values[i][0] || '').trim();
    if (!name) continue;
    if (normalizeName_(name) !== normalizedName) continue;

    return {
      row: CONFIG.MEMBER_FIRST_DATA_ROW + i,
      name: name,
      home: String(values[i][1] || '').trim(),
      joinDate: values[i][CONFIG.MEMBER_COL_JOIN_DATE - CONFIG.MEMBER_COL_NAME],
      status: String(values[i][CONFIG.MEMBER_COL_STATUS - CONFIG.MEMBER_COL_NAME] || '').trim(),
      oneHourBooking: values[i][CONFIG.MEMBER_COL_ONE_HOUR - CONFIG.MEMBER_COL_NAME] === true
    };
  }
  return null;
}

function cleanupExpiredMaintenance_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return 0;

  try {
    const currentLastRow = sheet.getLastRow();
    if (currentLastRow < 2) return 0;

    const rows = sheet.getRange(2, 1, currentLastRow - 1, 3).getValues();
    const nowMs = Date.now();
    const deleteRows = [];

    rows.forEach((row, index) => {
      const end = row[2];
      if (!(end instanceof Date)) return;
      if (end.getTime() <= nowMs) {
        deleteRows.push(index + 2);
      }
    });

    deleteRows.sort((a,b) => b - a).forEach(rowNumber => {
      sheet.deleteRow(rowNumber);
    });

    return deleteRows.length;
  } finally {
    lock.releaseLock();
  }
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
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
  // A simultaneous slots request may already have cancelled these rows.
  // Reread while holding the same lock as book/cancel before deleting/notifying.
  reservations = readReservations_(sheet);

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
    try { notifyDiscordAutoCancelled_(autoCancelled, maintenance); }
    catch (error) { console.error('自動取消の通知に失敗しました。予約の取消結果は保存済みです。'); }
  }

  return keep;
  } finally { lock.releaseLock(); }
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

function supportMethodLabel_(method) {
  return method === 'x' ? 'X' : (method === 'companion' ? 'コンパニオンアプリ' : 'ゲーム内');
}

function reservationDateTime_(reservation) {
  const time = reservation.time || '';
  const minute = bookingTimeMinute_(time, false);
  const end = minute + (reservation.oneHourBooking ? 60 : 30);
  const endTime = Number.isFinite(minute)
    ? String(Math.floor(end / 60)).padStart(2, '0') + ':' + String(end % 60).padStart(2, '0') : '';
  return (reservation.date || '日付未設定') + ' ' + (time || '時刻未設定') + (endTime ? '～' + endTime : '');
}

function reservationFields_(reservation, before) {
  const field = (name, value) => ({name: name, value: value, inline: false});
  const fields = [field('キャラクター名', reservation.name || '不明')];
  if (before && reservationDateTime_(before) !== reservationDateTime_(reservation)) {
    fields.push(field('変更前日時', reservationDateTime_(before)));
  }
  fields.push(field('予約日時', reservationDateTime_(reservation)));
  fields.push(field('予約時間', reservation.oneHourBooking ? '60分' : '30分'));
  fields.push(field('対応方法', supportMethodLabel_(reservation.supportMethod)));
  fields.push(field('受付方式', reservation.mode === 'flexible' ? '随時受付' : '通常予約'));
  if (reservation.meetingDc && reservation.meetingDc !== '(DC指定省略)') fields.push(field('集合DC', reservation.meetingDc));
  if (reservation.supportMethod === 'x' && reservation.contact) fields.push(field('連絡先', reservation.contact));
  if (before) {
    const changes = [];
    const properties = [
      ['対応方法', item => supportMethodLabel_(item.supportMethod)],
      ['予約時間', item => item.oneHourBooking ? '60分' : '30分'],
      ['受付方式', item => item.mode === 'flexible' ? '随時受付' : '通常予約'],
      ['集合DC', item => item.meetingDc && item.meetingDc !== '(DC指定省略)' ? item.meetingDc : '指定なし'],
      ['連絡先', item => item.supportMethod === 'x' && item.contact ? item.contact : 'なし']
    ];
    properties.forEach(([label, value]) => { if (value(before) !== value(reservation)) changes.push(label + '：' + value(before) + ' → ' + value(reservation)); });
    if (changes.length) fields.push(field('変更内容', changes.join('\n')));
  }
  return fields;
}

function hasReservationChanged_(before, after) {
  return (
    String(before.date || '') !== String(after.date || '') ||
    String(before.time || '') !== String(after.time || '') ||
    String(before.meetingDc || '') !== String(after.meetingDc || '') ||
    String(before.supportMethod || 'game') !== String(after.supportMethod || 'game') ||
    String(before.contact || '') !== String(after.contact || '') ||
    String(before.mode || 'normal') !== String(after.mode || 'normal') ||
    !!before.oneHourBooking !== !!after.oneHourBooking
  );
}

function notifyDiscordReservationCreated_(reservation) {
  queueBoardReservation_('🆕 新規予約', reservation);
  sendDiscordReservationEmbed_({
    title: '🆕 新規予約',
    fields: reservationFields_(reservation)
  });
}

function notifyDiscordReservationChanged_(before, after) {
  queueBoardReservation_('🔄 予約変更', after, before);
  sendDiscordReservationEmbed_({title: '🔄 予約変更', fields: reservationFields_(after, before)});
}

function notifyDiscordReservationCancelled_(reservation) {
  queueBoardReservation_('❌ 予約取り消し', reservation);
  sendDiscordReservationEmbed_({
    title: '❌ 予約取り消し',
    fields: reservationFields_(reservation)
  });
}

function notifyDiscordAutoCancelled_(reservations, maintenance) {
  const title = 'メンテナンスによる予約の自動キャンセル';
  const embeds = reservations.map(reservation => {
    const state = getMaintenanceStateForSlot_(reservation.date, reservation.time, maintenance);
    const extra = [];
    if (state.matches.length) {
      const stopped = [];
      if (state.blockGame) stopped.push('ゲーム内停止');
      if (state.blockX) stopped.push('X停止');
      if (state.blockCompanion) stopped.push('コンパニオン停止');
      extra.push({name: '取消理由', value: stopped.join(' / '), inline: false});
      extra.push({name: '該当時間', value: state.matches.map(period => period.start.replace('T', ' ') + ' ～ ' + period.end.replace('T', ' ') + (period.message ? ' / ' + period.message : '')).join('\n'), inline: false});
    }
    queueBoardReservation_(title, reservation, null, extra);
    return {title: title, fields: reservationFields_(reservation).concat(extra), footer: {text: '予約管理シートから自動削除済み'}};
  });
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK_URL');
  if (!webhookUrl) return;
  // Discord permits up to 10 embeds in one request.
  for (let index = 0; index < embeds.length; index += 10) {
    const response = UrlFetchApp.fetch(webhookUrl, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify({embeds: embeds.slice(index, index + 10)}), muteHttpExceptions: true
    });
    const status = response.getResponseCode();
    if (status < 200 || status >= 300) throw new Error('Discord Webhook通知に失敗しました。HTTP ' + status);
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
    data.supportMethod === 'x'
      ? 'X'
      : (data.supportMethod === 'companion' ? 'コンパニオンアプリ' : 'ゲーム内'),
    data.supportMethod === 'x' ? (data.contact || '') : '',
    updatedAt
  ]];

  const targetRow = existing
    ? existing.row
    : Math.max(sheet.getLastRow() + 1, CONFIG.RES_FIRST_DATA_ROW);

  sheet.getRange(targetRow, 1, 1, CONFIG.RES_COL_UPDATED_AT).setValues(rowValues);
  sheet
    .getRange(targetRow, CONFIG.RES_COL_MODE)
    .setValue(data.mode === 'flexible' ? '随時受付' : '通常予約');

  let remarks = String(data.remarks || '').trim();
  let remarkParts = remarks
    ? remarks.split(' / ').map(value => value.trim()).filter(Boolean)
    : [];
  remarkParts = remarkParts.filter(value => value !== '1時間予約');
  if (data.oneHourBooking) remarkParts.push('1時間予約');
  if (data.language === 'en' && !remarkParts.includes('英語表示から予約')) {
    remarkParts.push('英語表示から予約');
  }
  remarks = remarkParts.join(' / ');
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
  if (input.supportMethod === 'x') {
    throw new Error('Xによるサポートの受付は終了しました。ゲーム内またはコンパニオンアプリを選択してください。');
  }
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

  const privateSs = getReservationSpreadsheet_();
  const hours = getDailyBookingHours_(privateSs, input.date, input.date)[input.date];
  input.bookingHours = hours;
  if (!isWithinBookingHours_(hours, input.time, CONFIG.SLOT_MINUTES)) {
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

  const allowedMethods = ['game', 'companion'];
  if (!allowedMethods.includes(input.supportMethod)) {
    throw new Error('招待・登録サポート方法を選択してください。');
  }

  const allowedDc = ['Elemental', 'Gaia', 'Mana', 'Meteor'];
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

  if (!needsDc) input.dc = '';

  input.contact = '';
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
  if (endDate < startDate) return [];
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
    if (/^\d{4}-\d{2}-\d{2}$/.test(key) && key >= startDate) result.push(key);
  }
  return [...new Set(result)];
}
