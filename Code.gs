/**
 * Farmasi Marketing Masterclass — Premium Backend
 * Google Apps Script Code v2.1
 *
 * Sheet structure:
 * - "Registrations": timestamp | first_name | last_name | city | phone | position | ticket_id | tier | seat_number
 * - "Config": key | value
 *
 * Premium features:
 * - Ticket ID generation (FM-2026-XXXXX format)
 * - Tier badges (Early Bird / VIP / Standard) based on registration order
 * - Registration progress data
 * - Last 5 registrations with masked phones
 * - Telegram bot notifications (optional)
 * - 🆕 v2.1: Auto-close registration when event_date + event_time has passed
 */

const SHEET_NAME = 'Registrations';
const CONFIG_SHEET_NAME = 'Config';
const CANCELLATIONS_SHEET_NAME = 'Cancellations';  // archive sheet — auto-created on first cancellation

// 🤖 Telegram Bot (არჩევითი)
// ⚠️ Token-ი Script Properties-ში დააყენე გასაღებით: TELEGRAM_BOT_TOKEN

// 📲 SMS settings (bulksms.ge / POSTA GUVERCINI)
// PRIVATE_KEY and PUBLIC_KEY must be set in Apps Script → Project Settings → Script Properties
const SMS_API_URL = 'https://api.bulksms.ge/gateway/api/sms/v1/message/send';
const SMS_DEFAULT_SENDER = 'FARMASI';
const SITE_BASE_URL = 'https://farmasi-masterclass.vercel.app';

// Tier thresholds
const EARLY_BIRD_LIMIT = 10;  // First 10 = Early Bird
const VIP_LIMIT = 30;          // 11–30 = VIP
                              // 31+ = Standard

// =================== GET ===================
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const regSheet = ss.getSheetByName(SHEET_NAME);
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    if (!regSheet || !configSheet) {
      return jsonResponse({ ok: false, error: 'Sheets not found' });
    }

    // 🎫 Lookup ticket by ID
    const action = e && e.parameter && e.parameter.action;
    if (action === 'getTicket') {
      return handleGetTicket(regSheet, configSheet, e.parameter.id);
    }

    // 🔐 Verify ticket by phone number
    if (action === 'verifyTicketByPhone') {
      return handleVerifyByPhone(regSheet, configSheet, e.parameter.phone);
    }

    // 📲 Send ticket link via SMS
    if (action === 'sendTicketLink') {
      return handleSendTicketLink(regSheet, e.parameter.phone, e.parameter.id);
    }

    const config = readConfig(configSheet);
    const data = regSheet.getDataRange().getValues();
    const rows = data.slice(1).filter(r => r[0]);

    const totalRegistrations = rows.length;

    // Today's registrations
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayRegistrations = rows.filter(r => {
      const ts = new Date(r[0]);
      return ts >= today;
    }).length;

    // Spots left
    const maxReg = parseInt(config.max_registrations);
    const spotsLeft = (!isNaN(maxReg) && maxReg > 0)
      ? Math.max(0, maxReg - totalRegistrations)
      : null;

    // Occupied seats — column I (index 8)
    const occupiedSeats = rows
      .map(r => parseInt(r[8]))
      .filter(n => !isNaN(n) && n > 0);

    // Last registrations
    const lastRegistrations = rows
      .slice(-5)
      .reverse()
      .map(r => {
        const ts = new Date(r[0]);
        return {
          name: String(r[1] || '').trim() + ' ' + String(r[2] || '').trim(),
          phone: maskPhone(String(r[4] || '')),
          time: Utilities.formatDate(ts, 'Asia/Tbilisi', 'HH:mm'),
          tier: r[7] || '',
          seat: r[8] || ''
        };
      });

    // 🆕 Auto-close registration if event has already started/passed.
    // This overrides the manual `registration_status` value in Config —
    // even if admin forgot to set it to "closed", expired events are auto-locked.
    const eventEnded = isEventEnded(config);
    let status = String(config.registration_status || 'open').toLowerCase();
    if (eventEnded) status = 'closed';

    // Compute progress percentage
    let progressPercent = 0;
    if (!isNaN(maxReg) && maxReg > 0) {
      progressPercent = Math.round((totalRegistrations / maxReg) * 100);
    }

    // Determine current tier
    const nextTier = getNextTier(totalRegistrations);

    return jsonResponse({
      ok: true,
      status: status,
      eventEnded: eventEnded,             // 🆕 frontend uses this to show "ღონისძიება დასრულდა" banner
      totalRegistrations: totalRegistrations,
      todayRegistrations: todayRegistrations,
      spotsLeft: spotsLeft,
      maxRegistrations: !isNaN(maxReg) ? maxReg : null,
      progressPercent: progressPercent,
      nextTier: nextTier,
      occupiedSeats: occupiedSeats,
      config: config,
      lastRegistrations: lastRegistrations
    });

  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// =================== POST ===================
function doPost(e) {
  try {
    const action = String((e.parameter && e.parameter.action) || '').trim();

    if (action === 'register') return handleRegister(e.parameter);
    if (action === 'cancel') return handleCancel(e.parameter);

    return jsonResponse({ ok: false, error: 'Unknown action' });

  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// =================== REGISTER ===================
function handleRegister(params) {
  // 🔒 Acquire script-level lock to prevent race conditions.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return jsonResponse({ ok: false, error: 'სერვერი დაკავებულია, სცადეთ ხელახლა' });
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const regSheet = ss.getSheetByName(SHEET_NAME);
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    const config = readConfig(configSheet);

    // 🆕 Server-side enforcement: if event already started, no more registrations.
    // This guards against direct API calls (Postman, curl, DevTools) that bypass the frontend lock.
    if (isEventEnded(config)) {
      return jsonResponse({ ok: false, error: 'Event ended' });
    }

    if (String(config.registration_status || 'open').toLowerCase() !== 'open') {
      return jsonResponse({ ok: false, error: 'Registration is closed' });
    }

    const phone = normalizePhone(params.phone);
    const firstName = String(params.first_name || '').trim();
    const lastName = String(params.last_name || '').trim();
    const seatRaw = String(params.seat_number || '').trim();
    const seatNumber = seatRaw ? parseInt(seatRaw) : null;

    if (!phone) return jsonResponse({ ok: false, error: 'Phone required' });
    if (!firstName) return jsonResponse({ ok: false, error: 'First name required' });
    if (!lastName) return jsonResponse({ ok: false, error: 'Last name required' });
    if (!seatNumber || isNaN(seatNumber) || seatNumber < 1 || seatNumber > 54) {
      return jsonResponse({ ok: false, error: 'Invalid seat' });
    }

    const data = regSheet.getDataRange().getValues();
    const rowCount = data.slice(1).filter(r => r[0]).length;

    const maxReg = parseInt(config.max_registrations);
    if (!isNaN(maxReg) && maxReg > 0 && rowCount >= maxReg) {
      return jsonResponse({ ok: false, error: 'No spots left' });
    }

    // Find max existing ticket ID number to prevent collisions after cancellations.
    let maxTicketNum = 0;
    for (let i = 1; i < data.length; i++) {
      const tid = String(data[i][6] || '');
      const m = tid.match(/FM-\d{4}-(\d+)/);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > maxTicketNum) maxTicketNum = n;
      }
    }

    // Duplicate phone check + seat availability check
    for (let i = 1; i < data.length; i++) {
      if (normalizePhone(data[i][4]) === phone) {
        return jsonResponse({ ok: false, error: 'Already registered' });
      }
      const existingSeat = parseInt(data[i][8]);
      if (!isNaN(existingSeat) && existingSeat === seatNumber) {
        return jsonResponse({ ok: false, error: 'Seat taken' });
      }
    }

    const ticketId = generateTicketId(maxTicketNum + 1);
    const tier = getNextTier(rowCount);

    regSheet.appendRow([
      new Date(),
      firstName,
      lastName,
      String(params.city_area_village || '').trim(),
      phone,
      String(params.position || '').trim(),
      ticketId,
      tier,
      seatNumber
    ]);

    // Telegram notification
    notifyTelegram(config, {
      first_name: firstName,
      last_name: lastName,
      city: params.city_area_village,
      phone: phone,
      position: params.position,
      ticket_id: ticketId,
      tier: tier,
      seat: seatNumber
    });

    // 📲 SMS to participant
    sendTicketSMS(phone, {
      first_name: firstName,
      last_name: lastName,
      ticket_id: ticketId,
      tier: tier,
      seat: seatNumber,
      event_time: config.event_time || '14:00',
      event_date: config.event_date || '',
      transport_time: config.transport_time || '',
      transport_address: config.transport_address || ''
    });

    // 📨 SMS to admins
    notifyAdmins(config, {
      first_name: firstName,
      last_name: lastName,
      phone: phone,
      city: params.city_area_village,
      position: params.position,
      ticket_id: ticketId,
      tier: tier,
      seat: seatNumber,
      registration_number: rowCount + 1
    });

    return jsonResponse({
      ok: true,
      ticket_id: ticketId,
      tier: tier,
      seat_number: seatNumber,
      registration_number: rowCount + 1
    });

  } finally {
    lock.releaseLock();
  }
}

// =================== CANCEL ===================
function handleCancel(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return jsonResponse({ ok: false, error: 'სერვერი დაკავებულია, სცადეთ ხელახლა' });
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const regSheet = ss.getSheetByName(SHEET_NAME);
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    const config = readConfig(configSheet);

    const phone = normalizePhone(params.phone);
    const ticketId = String(params.ticket_id || '').trim();
    const reason = String(params.reason || '').trim().slice(0, 500);

    if (!phone && !ticketId) {
      return jsonResponse({ ok: false, error: 'Phone or ticket_id required' });
    }

    const data = regSheet.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      const rowPhone = normalizePhone(data[i][4]);
      const rowTicketId = String(data[i][6] || '').trim();

      const matchesPhone = phone && rowPhone === phone;
      const matchesTicketId = ticketId && rowTicketId === ticketId;

      if (matchesPhone || matchesTicketId) {
        const cancelledRow = data[i];
        const cancelledData = {
          original_timestamp: cancelledRow[0],
          first_name:  String(cancelledRow[1] || '').trim(),
          last_name:   String(cancelledRow[2] || '').trim(),
          city:        String(cancelledRow[3] || '').trim(),
          phone:       String(cancelledRow[4] || '').trim(),
          position:    String(cancelledRow[5] || '').trim(),
          ticket_id:   String(cancelledRow[6] || '').trim(),
          tier:        String(cancelledRow[7] || '').trim(),
          seat_number: cancelledRow[8] || null,
          reason:      reason
        };

        try {
          archiveCancellation(ss, cancelledData);
        } catch (e) {
          Logger.log('Archive failed (non-fatal): ' + e);
        }

        regSheet.deleteRow(i + 1);

        try {
          sendCancellationSMS(cancelledData.phone, cancelledData);
        } catch (e) {
          Logger.log('Cancellation SMS failed (non-fatal): ' + e);
        }

        try {
          notifyAdminsCancellation(config, cancelledData);
        } catch (e) {
          Logger.log('Admin cancellation notification failed (non-fatal): ' + e);
        }

        return jsonResponse({
          ok: true,
          cancelled: {
            ticket_id: cancelledData.ticket_id,
            seat_number: cancelledData.seat_number
          }
        });
      }
    }

    return jsonResponse({ ok: false, error: 'Not found' });
  } finally {
    lock.releaseLock();
  }
}

function archiveCancellation(ss, data) {
  let cancelSheet = ss.getSheetByName(CANCELLATIONS_SHEET_NAME);

  if (!cancelSheet) {
    cancelSheet = ss.insertSheet(CANCELLATIONS_SHEET_NAME);
    cancelSheet.appendRow([
      'cancelled_at', 'original_timestamp', 'first_name', 'last_name', 'city',
      'phone', 'position', 'ticket_id', 'tier', 'seat_number', 'reason'
    ]);
    cancelSheet.getRange(1, 1, 1, 11).setFontWeight('bold');
    cancelSheet.setFrozenRows(1);
    cancelSheet.setColumnWidth(11, 280);
  } else {
    const headers = cancelSheet.getRange(1, 1, 1, cancelSheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('reason') === -1) {
      const newCol = cancelSheet.getLastColumn() + 1;
      cancelSheet.getRange(1, newCol).setValue('reason').setFontWeight('bold');
      cancelSheet.setColumnWidth(newCol, 280);
    }
  }

  cancelSheet.appendRow([
    new Date(),
    data.original_timestamp || '',
    data.first_name || '',
    data.last_name || '',
    data.city || '',
    data.phone || '',
    data.position || '',
    data.ticket_id || '',
    data.tier || '',
    data.seat_number || '',
    data.reason || ''
  ]);
}

function sendCancellationSMS(phone, data) {
  if (!phone) return;

  const firstName = String(data.first_name || '').trim();
  const greeting = firstName ? firstName + ', ' : '';

  const text =
    '💔 ' + greeting + 'რეგისტრაცია გაუქმდა.\n\n' +
    'ვწუხვართ, რომ ვერ შეუერთდები მასტერკლასს.\n' +
    'შენი ადგილი თავისუფლდება სხვისთვის.\n\n' +
    'ჩვენ მაინც დაგელოდებით — ხელახლა რეგისტრაცია ნებისმიერ დროს შეგიძლია.\n' +
    '🌟 Farmasi';

  try {
    const props = PropertiesService.getScriptProperties();
    const privateKey = props.getProperty('PRIVATE_KEY');
    const publicKey = props.getProperty('PUBLIC_KEY');
    if (!privateKey || !publicKey) {
      Logger.log('SMS keys missing — cancellation SMS skipped');
      return;
    }

    const intlPhone = toInternationalPhone(phone);
    if (!intlPhone) return;

    const payload = {
      Text: text,
      Purpose: 'INF',
      Options: {
        Originator: SMS_DEFAULT_SENDER,
        Encoding: 'UNICODE',
        SmsType: 'SMS',
        ReportLabel: 'Farmasi Cancellation'
      },
      Receivers: [{ Receiver: intlPhone }]
    };

    const url = SMS_API_URL + '?publicKey=' + encodeURIComponent(publicKey);
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + privateKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log('Cancellation SMS sent to ' + maskPhone(phone));
  } catch (e) {
    Logger.log('Cancellation SMS error: ' + e);
  }
}

function notifyAdminsCancellation(config, data) {
  const fullName = (String(data.first_name || '') + ' ' + String(data.last_name || '')).trim();
  const reason = String(data.reason || '').trim();

  const adminPhones = String(config.admin_phones || '')
    .split(/[,;\n]/)
    .map(s => s.trim())
    .filter(Boolean);

  if (adminPhones.length) {
    let text =
      '❌ რეგისტრაცია გაუქმდა\n' +
      '👤 ' + fullName + '\n' +
      '📞 ' + (data.phone || '') + '\n' +
      '🪑 ადგილი #' + (data.seat_number || '-') + ' · ' + (data.tier || '-') + '\n' +
      '🎫 ' + (data.ticket_id || '');
    if (reason) {
      text += '\n💬 ' + reason;
    }

    adminPhones.forEach(p => {
      try {
        sendBulkSMS(p, text, 'Cancellation Alert');
      } catch (e) {
        Logger.log('Admin cancellation SMS failed for ' + maskPhone(p) + ': ' + e);
      }
    });
  }

  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) return;

  const chatIds = String(config.admin_chat_ids || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
  if (!chatIds.length) return;

  let tgText =
    '❌ რეგისტრაცია გაუქმდა\n\n' +
    '👤 ' + fullName + '\n' +
    '📞 ' + (data.phone || '') + '\n' +
    '📍 ' + (data.city || '-') + '\n' +
    '💼 ' + (data.position || '-') + '\n\n' +
    '🪑 ადგილი #' + (data.seat_number || '-') + '\n' +
    '🥇 ' + (data.tier || '-') + '\n' +
    '🎫 ' + (data.ticket_id || '');
  if (reason) {
    tgText += '\n\n💬 მიზეზი:\n' + reason;
  }

  chatIds.forEach(chatId => {
    try {
      UrlFetchApp.fetch(
        'https://api.telegram.org/bot' + token + '/sendMessage',
        {
          method: 'post',
          payload: { chat_id: chatId, text: tgText },
          muteHttpExceptions: true
        }
      );
    } catch (e) {
      Logger.log('Telegram cancellation alert failed: ' + e);
    }
  });
}

// =================== HELPERS ===================
function readConfig(sheet) {
  const config = {};
  const data = sheet.getDataRange().getValues();

  const TIME_KEYS = [
    'event_time', 'meeting_end_time', 'transport_time', 'reminder_2h'
  ];

  const DATE_KEYS = [
    'event_date', 'transport_date', 'reminder_day'
  ];

  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    if (!key) continue;

    let val = data[i][1];

    if (val instanceof Date) {
      if (DATE_KEYS.indexOf(key) !== -1) {
        val = Utilities.formatDate(val, 'Asia/Tbilisi', 'yyyy-MM-dd');
      } else if (TIME_KEYS.indexOf(key) !== -1) {
        val = Utilities.formatDate(val, 'Asia/Tbilisi', 'HH:mm');
      } else {
        if (val.getFullYear() === 1899) {
          val = Utilities.formatDate(val, 'Asia/Tbilisi', 'HH:mm');
        } else {
          val = Utilities.formatDate(val, 'Asia/Tbilisi', 'yyyy-MM-dd HH:mm');
        }
      }
    }

    config[key] = String(val || '');
  }

  return config;
}

function normalizePhone(p) {
  return String(p || '').replace(/\D/g, '').trim();
}

function maskPhone(p) {
  const clean = normalizePhone(p);
  if (clean.length < 6) return clean;
  return clean.substring(0, 3) + '***' + clean.substring(clean.length - 3);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateTicketId(registrationNumber) {
  const year = new Date().getFullYear();
  const padded = String(registrationNumber).padStart(5, '0');
  return `FM-${year}-${padded}`;
}

function getNextTier(currentCount) {
  if (currentCount < EARLY_BIRD_LIMIT) return 'Early Bird';
  if (currentCount < VIP_LIMIT) return 'VIP';
  return 'Standard';
}

/**
 * 🆕 Check if the event has already started (or passed).
 * Returns true → registration must be locked.
 *
 * Uses event_date (YYYY-MM-DD) + event_time (HH:mm) from Config.
 * Returns false if event_date is empty or malformed (fail-safe — better
 * to allow registration than to lock it for no reason).
 *
 * Relies on Apps Script timezone being Asia/Tbilisi (Project Settings).
 * If event_time is missing, defaults to 23:59 (end of day).
 */
function isEventEnded(config) {
  const dateStr = String(config.event_date || '').trim();
  if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return false;

  const timeStr = String(config.event_time || '23:59').trim();
  const tm = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  const hh = tm ? parseInt(tm[1], 10) : 23;
  const mm = tm ? parseInt(tm[2], 10) : 59;

  const parts = dateStr.split('-').map(Number);
  const eventStart = new Date(parts[0], parts[1] - 1, parts[2], hh, mm, 0, 0);

  return Date.now() > eventStart.getTime();
}

function notifyTelegram(config, data) {
  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) {
    Logger.log('TELEGRAM_BOT_TOKEN not set in Script Properties — Telegram skipped');
    return;
  }

  const chatIds = String(config.admin_chat_ids || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);

  if (!chatIds.length) {
    Logger.log('No admin_chat_ids in Config — Telegram skipped');
    return;
  }

  const tierEmoji = data.tier === 'Early Bird' ? '🥇' : data.tier === 'VIP' ? '✨' : '🎫';

  const text =
    '🆕 ახალი რეგისტრაცია (მასტერკლასი)\n\n' +
    tierEmoji + ' ' + data.tier + '\n' +
    '🎫 ' + data.ticket_id + '\n' +
    '🪑 ადგილი #' + (data.seat || '-') + '\n\n' +
    '👤 ' + (data.first_name || '') + ' ' + (data.last_name || '') + '\n' +
    '📍 ' + (data.city || '-') + '\n' +
    '💼 ' + (data.position || '-') + '\n' +
    '📱 ' + (data.phone || '');

  let sentCount = 0;
  chatIds.forEach(chatId => {
    try {
      const response = UrlFetchApp.fetch(
        'https://api.telegram.org/bot' + token + '/sendMessage',
        {
          method: 'post',
          payload: { chat_id: chatId, text: text },
          muteHttpExceptions: true
        }
      );
      const code = response.getResponseCode();
      if (code === 200) {
        sentCount++;
      } else {
        Logger.log('Telegram error for chat ' + chatId + ': HTTP ' + code + ' — ' + response.getContentText());
      }
    } catch (e) {
      Logger.log('Telegram exception for chat ' + chatId + ': ' + e);
    }
  });

  Logger.log('Telegram sent to ' + sentCount + '/' + chatIds.length + ' admins');
}

// =================== GET TICKET BY ID ===================
function handleGetTicket(regSheet, configSheet, ticketId) {
  ticketId = String(ticketId || '').trim();
  if (!ticketId) {
    return jsonResponse({ ok: false, error: 'Ticket ID required' });
  }

  const data = regSheet.getDataRange().getValues();
  const config = readConfig(configSheet);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][6]).trim() === ticketId) {
      return jsonResponse({
        ok: true,
        ticket: {
          ticket_id: ticketId,
          first_name: String(data[i][1] || '').trim(),
          last_name: String(data[i][2] || '').trim(),
          city: String(data[i][3] || '').trim(),
          phone: maskPhone(String(data[i][4] || '')),
          position: String(data[i][5] || '').trim(),
          tier: String(data[i][7] || 'Standard').trim(),
          seat_number: parseInt(data[i][8]) || null
        },
        config: config
      });
    }
  }

  return jsonResponse({ ok: false, error: 'Ticket not found' });
}

// =================== SMS via bulksms.ge ===================
function sendTicketSMS(phone, data) {
  try {
    const props = PropertiesService.getScriptProperties();
    const privateKey = props.getProperty('PRIVATE_KEY');
    const publicKey = props.getProperty('PUBLIC_KEY');

    if (!privateKey || !publicKey) {
      Logger.log('SMS keys missing in Script Properties — SMS skipped');
      return;
    }

    const intlPhone = toInternationalPhone(phone);
    if (!intlPhone) {
      Logger.log('Invalid phone for SMS: ' + phone);
      return;
    }

    const text = composeTicketSMS(data);

    const payload = {
      Text: text,
      Purpose: 'INF',
      Options: {
        Originator: SMS_DEFAULT_SENDER,
        Encoding: 'UNICODE',
        SmsType: 'SMS',
        ReportLabel: 'Farmasi Masterclass Ticket'
      },
      Receivers: [{ Receiver: intlPhone }]
    };

    const url = SMS_API_URL + '?publicKey=' + encodeURIComponent(publicKey);

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + privateKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const body = response.getContentText();
    Logger.log('SMS API response (' + code + '): ' + body);

    if (code !== 200) {
      Logger.log('SMS sending failed for ' + intlPhone);
    }
  } catch (err) {
    Logger.log('SMS error: ' + err);
  }
}

function composeTicketSMS(data) {
  const tierEmoji = data.tier === 'Early Bird' ? '🌟'
                  : data.tier === 'VIP' ? '💎'
                  : '✦';

  const fullName = String((data.first_name || '') + ' ' + (data.last_name || '')).trim();

  let text = 'Farmasi Masterclass 2026\n';
  text += tierEmoji + ' ' + fullName + '\n';
  text += '🎫 ' + (data.ticket_id || '') + '\n';
  text += '🪑 ადგილი #' + (data.seat || '-') + '\n';

  if (data.transport_time) {
    text += '🚌 გასვლა: ' + data.transport_time + '\n';
  }

  if (data.transport_address) {
    const addr = data.transport_address.length > 50
      ? data.transport_address.substring(0, 47) + '...'
      : data.transport_address;
    text += '📍 ' + addr + '\n';
  }

  text += '\n🎫 ბილეთი: ' + SITE_BASE_URL + '/ticket.html?t=' + encodeURIComponent(data.ticket_id || '');

  return text;
}

function toInternationalPhone(phone) {
  let clean = String(phone || '').replace(/\D/g, '');

  if (!clean) return null;

  if (clean.startsWith('995') && clean.length === 12) {
    return clean;
  }

  if (clean.length === 9 && clean.startsWith('5')) {
    return '995' + clean;
  }

  if (clean.length >= 11 && clean.length <= 14) {
    return clean;
  }

  return null;
}

function testSMS() {
  const testPhone = '599772266';

  sendTicketSMS(testPhone, {
    first_name: 'ტესტი',
    last_name: 'ბერიძე',
    ticket_id: 'FM-2026-00001',
    tier: 'Early Bird',
    seat: 14,
    event_time: '14:00',
    event_date: '2026-06-15',
    transport_time: '12:00',
    transport_address: 'მეტრო ვაგზლის მოედანი'
  });

  Logger.log('Test SMS sent (check logs above for response)');
}


// =================== VERIFY TICKET BY PHONE ===================
function handleVerifyByPhone(regSheet, configSheet, rawPhone) {
  const inputDigits = String(rawPhone || '').replace(/\D/g, '');
  if (inputDigits.length < 9) {
    return jsonResponse({ ok: false, error: 'არასწორი ტელეფონის ფორმატი' });
  }

  const inputLast9 = inputDigits.slice(-9);
  const data = regSheet.getDataRange().getValues();
  const config = readConfig(configSheet);

  for (let i = 1; i < data.length; i++) {
    const sheetPhone = String(data[i][4] || '').replace(/\D/g, '');
    if (!sheetPhone) continue;

    if (sheetPhone.slice(-9) === inputLast9) {
      return jsonResponse({
        ok: true,
        ticket: {
          ticket_id:   String(data[i][6] || '').trim(),
          first_name:  String(data[i][1] || '').trim(),
          last_name:   String(data[i][2] || '').trim(),
          city:        String(data[i][3] || '').trim(),
          phone:       String(data[i][4] || '').trim(),
          position:    String(data[i][5] || '').trim(),
          tier:        String(data[i][7] || 'Standard').trim(),
          seat_number: parseInt(data[i][8]) || null
        },
        config: config
      });
    }
  }

  return jsonResponse({ ok: false, error: 'ამ ტელეფონით ბილეთი ვერ მოიძებნა' });
}


// =================== SEND TICKET LINK VIA SMS ===================
function handleSendTicketLink(regSheet, rawPhone, ticketId) {
  ticketId = String(ticketId || '').trim();
  const inputDigits = String(rawPhone || '').replace(/\D/g, '');

  if (!ticketId || inputDigits.length < 9) {
    return jsonResponse({ ok: false, error: 'ბილეთის ID ან ტელეფონი არ მოგვცეს' });
  }

  const inputLast9 = inputDigits.slice(-9);
  const data = regSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowTicket = String(data[i][6] || '').trim();
    const rowPhoneDigits = String(data[i][4] || '').replace(/\D/g, '');

    if (rowTicket === ticketId && rowPhoneDigits.slice(-9) === inputLast9) {
      const phone = String(data[i][4] || '').trim();
      const firstName = String(data[i][1] || '').trim();

      const link = SITE_BASE_URL + '/ticket.html?t=' + encodeURIComponent(ticketId);
      const text = '🎫 ' + (firstName ? firstName + ', ' : '') +
        'შენი ბილეთი:\n' + link;

      try {
        const props = PropertiesService.getScriptProperties();
        const privateKey = props.getProperty('PRIVATE_KEY');
        const publicKey = props.getProperty('PUBLIC_KEY');
        if (!privateKey || !publicKey) {
          return jsonResponse({ ok: false, error: 'SMS კონფიგურაცია არ არის' });
        }

        const intlPhone = toInternationalPhone(phone);
        if (!intlPhone) {
          return jsonResponse({ ok: false, error: 'არასწორი ტელეფონი' });
        }

        const payload = {
          Text: text,
          Purpose: 'INF',
          Options: {
            Originator: SMS_DEFAULT_SENDER,
            Encoding: 'UNICODE',
            SmsType: 'SMS',
            ReportLabel: 'Farmasi Ticket Link'
          },
          Receivers: [{ Receiver: intlPhone }]
        };

        const url = SMS_API_URL + '?publicKey=' + encodeURIComponent(publicKey);
        const response = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          headers: { 'Authorization': 'Bearer ' + privateKey },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
          return jsonResponse({ ok: true });
        }
        return jsonResponse({ ok: false, error: 'SMS API შეცდომა' });
      } catch (err) {
        return jsonResponse({ ok: false, error: 'SMS შეცდომა: ' + String(err) });
      }
    }
  }

  return jsonResponse({ ok: false, error: 'ბილეთი და ტელეფონი ვერ ემთხვევა' });
}


// =================================================================
// 📨 ADMIN NOTIFICATIONS
// =================================================================
function notifyAdmins(config, data) {
  try {
    const adminPhonesStr = String(config.admin_phones || '').trim();
    if (!adminPhonesStr) {
      Logger.log('No admin_phones in Config — admin SMS skipped');
      return;
    }

    const adminPhones = adminPhonesStr
      .split(/[,;\n]/)
      .map(p => p.trim())
      .filter(p => p.length >= 9);

    if (adminPhones.length === 0) {
      Logger.log('No valid admin phones found');
      return;
    }

    const tierEmoji = data.tier === 'Early Bird' ? '🌟'
                    : data.tier === 'VIP' ? '💎'
                    : '✦';
    const text =
      '🆕 ახალი რეგისტრაცია #' + data.registration_number + '\n' +
      tierEmoji + ' ' + data.first_name + ' ' + data.last_name + '\n' +
      '📞 ' + data.phone + '\n' +
      '🪑 ადგილი #' + data.seat + ' · ' + data.tier + '\n' +
      '🎫 ' + data.ticket_id;

    let sentCount = 0;
    adminPhones.forEach(phone => {
      const ok = sendBulkSMS(phone, text, 'Farmasi Admin Alert');
      if (ok) sentCount++;
    });

    Logger.log('Admin SMS sent to ' + sentCount + '/' + adminPhones.length + ' admins');
  } catch (err) {
    Logger.log('notifyAdmins error: ' + err);
  }
}


// =================================================================
// 📲 GENERIC BULK SMS HELPER
// =================================================================
function sendBulkSMS(phone, text, label) {
  try {
    const props = PropertiesService.getScriptProperties();
    const privateKey = props.getProperty('PRIVATE_KEY');
    const publicKey = props.getProperty('PUBLIC_KEY');

    if (!privateKey || !publicKey) {
      Logger.log('SMS keys missing — skipping');
      return false;
    }

    const intlPhone = toInternationalPhone(phone);
    if (!intlPhone) {
      Logger.log('Invalid phone: ' + phone);
      return false;
    }

    const payload = {
      Text: text,
      Purpose: 'INF',
      Options: {
        Originator: SMS_DEFAULT_SENDER,
        Encoding: 'UNICODE',
        SmsType: 'SMS',
        ReportLabel: label || 'Farmasi'
      },
      Receivers: [{ Receiver: intlPhone }]
    };

    const url = SMS_API_URL + '?publicKey=' + encodeURIComponent(publicKey);

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + privateKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    Logger.log('SMS to ' + intlPhone + ' (' + label + '): ' + code);
    return code === 200;
  } catch (err) {
    Logger.log('sendBulkSMS error: ' + err);
    return false;
  }
}


// =================================================================
// 🔔 REMINDER 1: One day before the event
// =================================================================
function sendReminderDayBefore() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const regSheet = ss.getSheetByName(SHEET_NAME);
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    if (!regSheet || !configSheet) {
      Logger.log('Sheets not found');
      return;
    }

    const config = readConfig(configSheet);
    const eventDateStr = String(config.event_date || '').trim();
    if (!eventDateStr) {
      Logger.log('event_date not set — reminder skipped');
      return;
    }

    const daysBefore = parseInt(config.reminder_days_before || '1', 10) || 1;

    const eventDate = new Date(eventDateStr + 'T00:00:00');
    const targetDate = new Date();
    targetDate.setDate(targetDate.getDate() + daysBefore);
    targetDate.setHours(0, 0, 0, 0);

    if (eventDate.toDateString() !== targetDate.toDateString()) {
      Logger.log('Today is not ' + daysBefore + ' day(s) before event (' + eventDateStr + ') — skipping');
      return;
    }

    const eventCity = config.event_city || '';
    const eventTime = config.event_time || '14:00';
    const transportTime = config.transport_time || '';
    const transportAddress = config.transport_address || '';

    let whenPhrase;
    if (daysBefore === 1) whenPhrase = 'ხვალ';
    else if (daysBefore === 2) whenPhrase = 'ზეგ';
    else if (daysBefore === 3) whenPhrase = 'მაზეგ';
    else whenPhrase = daysBefore + ' დღეში';

    let text = '🔔 შეხსენება!\n' +
      whenPhrase + ' Farmasi-ის მასტერკლასია 🎓\n' +
      '📅 ' + eventDateStr + ' · ' + eventTime + '\n' +
      '📍 ' + eventCity;

    if (transportTime && transportAddress) {
      text += '\n🚌 ავტობუსი: ' + transportTime + '\n📍 ' + transportAddress;
    }

    const data = regSheet.getDataRange().getValues();
    const phones = [];
    for (let i = 1; i < data.length; i++) {
      const phone = String(data[i][4] || '').trim();
      if (phone) phones.push(phone);
    }

    if (phones.length === 0) {
      Logger.log('No registered phones to remind');
      return;
    }

    let sent = 0;
    phones.forEach((phone, idx) => {
      const ok = sendBulkSMS(phone, text, 'Farmasi Day Before Reminder');
      if (ok) sent++;
      if ((idx + 1) % 10 === 0) {
        Utilities.sleep(1000);
      }
    });

    Logger.log('Day-before reminder sent to ' + sent + '/' + phones.length + ' participants');
  } catch (err) {
    Logger.log('sendReminderDayBefore error: ' + err);
  }
}


// =================================================================
// 🚌 REMINDER 2: Hours before bus departure
// =================================================================
function sendReminderBusDeparture() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const regSheet = ss.getSheetByName(SHEET_NAME);
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    if (!regSheet || !configSheet) {
      Logger.log('Sheets not found');
      return;
    }

    const config = readConfig(configSheet);
    const eventDateStr = String(config.event_date || '').trim();
    const transportTime = String(config.transport_time || '').trim();

    if (!eventDateStr || !transportTime) {
      Logger.log('event_date or transport_time missing — bus reminder skipped');
      return;
    }

    const today = new Date();
    const eventDate = new Date(eventDateStr + 'T00:00:00');
    if (today.toDateString() !== eventDate.toDateString()) {
      Logger.log('Today is not event day — bus reminder skipped');
      return;
    }

    const timeParts = transportTime.match(/^(\d{1,2}):(\d{2})$/);
    if (!timeParts) {
      Logger.log('Invalid transport_time format: ' + transportTime);
      return;
    }
    const busHour = parseInt(timeParts[1], 10);
    const busMin = parseInt(timeParts[2], 10);
    const busDateTime = new Date(today);
    busDateTime.setHours(busHour, busMin, 0, 0);

    const hoursBefore = parseFloat(config.reminder_hours_before || '3') || 3;

    const reminderTime = new Date(busDateTime.getTime() - hoursBefore * 60 * 60 * 1000);

    const now = new Date();
    const diffMinutes = Math.abs((now.getTime() - reminderTime.getTime()) / (60 * 1000));

    if (diffMinutes > 30) {
      Logger.log('Not within 30 min of bus reminder time (' + hoursBefore + 'h before). Now: ' + now + ', Reminder time: ' + reminderTime);
      return;
    }

    const transportAddress = config.transport_address || '';

    let hoursLabel;
    if (hoursBefore === 0.5) hoursLabel = '30 წუთში';
    else if (hoursBefore === 1) hoursLabel = '1 საათში';
    else if (Number.isInteger(hoursBefore)) hoursLabel = hoursBefore + ' საათში';
    else hoursLabel = hoursBefore + ' საათში';

    let text = '🚌 ავტობუსი ' + hoursLabel + '!\n' +
      'Farmasi მასტერკლასი 🎓\n' +
      '⏰ გასვლა: ' + transportTime;

    if (transportAddress) {
      text += '\n📍 ' + transportAddress;
    }
    text += '\n\n🎫 თან წაიღე ბილეთი (QR კოდი)';

    const data = regSheet.getDataRange().getValues();
    const phones = [];
    for (let i = 1; i < data.length; i++) {
      const phone = String(data[i][4] || '').trim();
      if (phone) phones.push(phone);
    }

    if (phones.length === 0) {
      Logger.log('No registered phones to remind');
      return;
    }

    let sent = 0;
    phones.forEach((phone, idx) => {
      const ok = sendBulkSMS(phone, text, 'Farmasi Bus Departure Reminder');
      if (ok) sent++;
      if ((idx + 1) % 10 === 0) Utilities.sleep(1000);
    });

    Logger.log('Bus reminder sent to ' + sent + '/' + phones.length + ' participants');
  } catch (err) {
    Logger.log('sendReminderBusDeparture error: ' + err);
  }
}


// =================================================================
// 🧪 TEST FUNCTIONS
// =================================================================
function testAdminNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig(ss.getSheetByName(CONFIG_SHEET_NAME));

  notifyAdmins(config, {
    first_name: 'ტესტ',
    last_name: 'ტესტიძე',
    phone: '599123456',
    city: 'თბილისი',
    position: 'ინფლუენსერი',
    ticket_id: 'FM-2026-99999',
    tier: 'Early Bird',
    seat: 1,
    registration_number: 1
  });
  Logger.log('Test admin notification sent');
}

function testTelegram() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig(ss.getSheetByName(CONFIG_SHEET_NAME));

  Logger.log('admin_chat_ids from Config: ' + (config.admin_chat_ids || '(empty)'));
  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  Logger.log('TELEGRAM_BOT_TOKEN in Script Properties: ' + (token ? '✅ set (' + token.length + ' chars)' : '❌ MISSING'));

  notifyTelegram(config, {
    first_name: 'ტესტ',
    last_name: 'ტესტიძე',
    phone: '599123456',
    city: 'თბილისი',
    position: 'ინფლუენსერი',
    ticket_id: 'FM-2026-99999',
    tier: 'Early Bird',
    seat: 1
  });
  Logger.log('Test Telegram notification finished — check your Telegram chat');
}

/**
 * 🆕 Test the event-ended check.
 * Run this after editing event_date in Config to verify the lock works.
 */
function testEventEndedCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig(ss.getSheetByName(CONFIG_SHEET_NAME));

  Logger.log('event_date: ' + config.event_date);
  Logger.log('event_time: ' + config.event_time);
  Logger.log('isEventEnded: ' + isEventEnded(config));
  Logger.log('Current time: ' + new Date());
}

function testDayReminder() {
  Logger.log('Run sendReminderDayBefore() manually to test (only works on actual day-before date)');
}
