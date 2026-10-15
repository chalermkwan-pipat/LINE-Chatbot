/**
 * LINE group-aware reminder bot backed by Google Sheets.
 * Start a chat with "hi bot" and stop it with "bye bot".
 * Configure Script Properties: LINE_CHANNEL_ACCESS_TOKEN, WEBHOOK_KEY.
 */
const CONFIG = {
  TASKS_SHEET: 'Tasks',
  SAVINGS_SHEET: 'Savings',
  USERS_SHEET: 'Users',
  TIMEZONE: 'Asia/Bangkok',
  DAILY_HOUR: 8
};

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, CONFIG.TASKS_SHEET,
    ['taskId', 'scopeKey', 'รายละเอียด', 'วัน', 'เวลา', 'สถานะ', 'สร้างเมื่อ', 'สถานที่', 'รูปแบบเตือน']);
  ensureSheet_(ss, CONFIG.SAVINGS_SHEET,
    ['recordId', 'scopeKey', 'วัน', 'ประเภท', 'จำนวนเงิน', 'หมายเหตุ', 'สร้างเมื่อ']);
  ensureSheet_(ss, CONFIG.USERS_SHEET,
    ['scopeKey', 'รับแจ้งเตือน', 'สมัครเมื่อ', 'ขั้นตอนบันทึก', 'targetId', 'targetType', 'sessionActive', 'lastActivityAt', 'pendingAppointment']);
  ScriptApp.getProjectTriggers()
    .filter(t => ['sendDailyReminders', 'expireIdleSessions'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyReminders').timeBased().everyDays(1)
    .atHour(CONFIG.DAILY_HOUR).inTimezone(CONFIG.TIMEZONE).create();
  ScriptApp.newTrigger('expireIdleSessions').timeBased().everyMinutes(1).create();
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }
  // Add any new columns at the end, keeping existing data intact.
  for (let col = sheet.getLastColumn() + 1; col <= headers.length; col++) {
    sheet.getRange(1, col).setValue(headers[col - 1]);
  }
}

// Deploy as a Web app (execute as you; access: anyone).
// Add ?key=WEBHOOK_KEY to the LINE webhook URL.
function doPost(e) {
  const expected = PropertiesService.getScriptProperties().getProperty('WEBHOOK_KEY');
  if (!expected || !e || !e.parameter || e.parameter.key !== expected) {
    return ContentService.createTextOutput('unauthorized');
  }
  const body = JSON.parse((e.postData && e.postData.contents) || '{}');
  (body.events || []).forEach(handleEvent_);
  return ContentService.createTextOutput('OK');
}

function handleEvent_(event) {
  if (event.type !== 'message' || !event.message || event.message.type !== 'text') return;
  const target = getTarget_(event.source);
  if (!target) return;
  const text = event.message.text.trim();
  const lower = text.toLowerCase();

  if (lower === 'hi bot') {
    setChatSession_(target, true);
    setRecordingState_(target, 'SELECT_TOPIC');
    replyLine_(event.replyToken, 'เริ่มทำงานแล้วครับ\n' + topicPrompt_());
    return;
  }
  if (lower === 'bye bot') {
    if (getChatSession_(target.ownerKey)) {
      endChatSession_(target);
      replyLine_(event.replyToken, 'จบการทำงานแล้วครับ พิมพ์ “hi bot” เมื่อต้องการเริ่มใหม่');
    }
    return;
  }
  // Ignore all messages in a chat until it starts the bot.
  if (!getChatSession_(target.ownerKey)) return;
  touchChatSession_(target);
  if (['ออก', 'ออกจากบอท', 'จบ'].includes(lower)) {
    endChatSession_(target);
    replyLine_(event.replyToken, 'ออกจากการใช้งานบอทแล้วครับ พิมพ์ “hi bot” เมื่อต้องการเริ่มใหม่');
    return;
  }

  const flowReply = handleRecordingFlow_(target, text);
  if (flowReply !== null) {
    replyLine_(event.replyToken, flowReply);
    return;
  }

  if (isBalanceCommand_(text)) {
    replyLine_(event.replyToken, getSavingsBalance_(target.ownerKey));
    return;
  }

  let reply;
  if (lower === '/help' || text === 'ช่วยเหลือ') {
    reply = helpText_();
  } else if (text === 'บันทึก') {
    setRecordingState_(target, 'SELECT_TOPIC');
    reply = topicPrompt_();
  } else if (text === 'เริ่ม') {
    upsertChat_(target, true);
    reply = 'เปิดรับแจ้งเตือนรายวันในแชตนี้แล้วครับ';
  } else if (text === 'หยุดแจ้งเตือน') {
    upsertChat_(target, false);
    reply = 'ปิดแจ้งเตือนรายวันในแชตนี้แล้วครับ';
  } else if (text === 'เริ่มแจ้งเตือน') {
    upsertChat_(target, true);
    reply = 'เปิดรับแจ้งเตือนรายวันในแชตนี้แล้วครับ';
  } else if (text.startsWith('เพิ่มงาน')) {
    reply = addTask_(target, text);
    if (reply.startsWith('บันทึกแล้ว')) {
      setRecordingState_(target, 'POST_SAVE_APPOINTMENT');
      reply += '\n\n' + afterSavePrompt_('นัดหมาย');
    }
  } else if (text === 'วันนี้') {
    reply = listToday_(target.ownerKey);
  } else if (text.startsWith('สรุปงาน')) {
    reply = summarizeDate_(target.ownerKey, text);
  } else if (text.startsWith('ลบงาน')) {
    reply = deleteTask_(target.ownerKey, text);
  } else {
    reply = 'ไม่รู้จักคำสั่งนี้ครับ\n' + helpText_();
  }
  replyLine_(event.replyToken, reply);
}

function getTarget_(source) {
  if (!source) return null;
  if (source.type === 'group' && source.groupId) {
    return {ownerKey: 'group:' + source.groupId, targetId: source.groupId, targetType: 'group'};
  }
  if (source.type === 'room' && source.roomId) {
    return {ownerKey: 'room:' + source.roomId, targetId: source.roomId, targetType: 'room'};
  }
  if (source.type === 'user' && source.userId) {
    return {ownerKey: source.userId, targetId: source.userId, targetType: 'user'};
  }
  return null;
}

function getChatSession_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === ownerKey) {
      const active = rows[i][6] === true || String(rows[i][6]).toLowerCase() === 'true';
      const lastActivity = rows[i][7];
      if (!active) return false;
      if (!(lastActivity instanceof Date) || Date.now() - lastActivity.getTime() >= 60000) {
        sheet.getRange(i + 1, 7).setValue(false);
        sheet.getRange(i + 1, 4).setValue('');
        return false;
      }
      return true;
    }
  }
  return false;
}

function setChatSession_(target, active) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === target.ownerKey) {
      sheet.getRange(i + 1, 5).setValue(target.targetId);
      sheet.getRange(i + 1, 6).setValue(target.targetType);
      sheet.getRange(i + 1, 7).setValue(active);
      sheet.getRange(i + 1, 8).setValue(active ? new Date() : '');
      return;
    }
  }
  sheet.appendRow([target.ownerKey, false, new Date(), '', target.targetId, target.targetType, active, active ? new Date() : '']);
}

function endChatSession_(target) {
  setRecordingState_(target, '');
  setChatSession_(target, false);
}

function touchChatSession_(target) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === target.ownerKey) {
      sheet.getRange(i + 1, 5).setValue(target.targetId);
      sheet.getRange(i + 1, 6).setValue(target.targetType);
      sheet.getRange(i + 1, 8).setValue(new Date());
      return;
    }
  }
}

function expireIdleSessions() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  const now = Date.now();
  for (let i = 1; i < rows.length; i++) {
    const active = rows[i][6] === true || String(rows[i][6]).toLowerCase() === 'true';
    const lastActivity = rows[i][7];
    if (active && (!(lastActivity instanceof Date) || now - lastActivity.getTime() >= 60000)) {
      sheet.getRange(i + 1, 7).setValue(false);
      sheet.getRange(i + 1, 4).setValue('');
    }
  }
}

function getRecordingState_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === ownerKey) return String(rows[i][3] || '');
  }
  return '';
}

function setRecordingState_(target, state) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === target.ownerKey) {
      sheet.getRange(i + 1, 4).setValue(state);
      sheet.getRange(i + 1, 5).setValue(target.targetId);
      sheet.getRange(i + 1, 6).setValue(target.targetType);
      return;
    }
  }
  sheet.appendRow([target.ownerKey, false, new Date(), state, target.targetId, target.targetType, false, '']);
}

function setPendingAppointment_(ownerKey, appointment) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === ownerKey) {
      sheet.getRange(i + 1, 9).setValue(appointment ? JSON.stringify(appointment) : '');
      return;
    }
  }
}

function getPendingAppointment_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === ownerKey && rows[i][8]) {
      try { return JSON.parse(rows[i][8]); } catch (err) { return null; }
    }
  }
  return null;
}

function upsertChat_(target, notificationsEnabled) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === target.ownerKey) {
      sheet.getRange(i + 1, 2).setValue(notificationsEnabled);
      sheet.getRange(i + 1, 5).setValue(target.targetId);
      sheet.getRange(i + 1, 6).setValue(target.targetType);
      return;
    }
  }
  const now = new Date();
  sheet.appendRow([target.ownerKey, notificationsEnabled, now, '', target.targetId, target.targetType, true, now]);
}

function helpText_() {
  return 'คำสั่งหลัก:\n' +
    'hi bot — เริ่มทำงานและแสดงหัวข้อหลัก\n' +
    'bye bot — จบการทำงานในแชตนี้\n' +
    '1 นัดหมาย / 2 เงินเก็บ — เลือกหัวข้อหลัก\n' +
    '3 เช็คตารางวันนี้ / 4 เช็คตารางวันพรุ่งนี้ / 5 เช็คยอดเงินฝาก balance — เมนูลัด\n' +
    'หลังเลือกหัวข้อ ให้เลือกเมนูย่อยตามที่บอทแสดง\n' +
    'หลังบันทึก: 1 บันทึกต่อ, 2 กลับหน้าหลัก, 3 ออกจากบอท\n' +
    'เมนูหลักเลือก 3 หรือเมนูย่อยเลือก 9 เพื่อออก โดยไม่ต้องพิมพ์ bye bot\n' +
    'ไม่มีข้อความเกิน 1 นาที บอทจะปิดการทำงานอัตโนมัติ\n' +
    'ยอดเงินเก็บ หรือ /balance — ดูยอดฝากสะสม หักยอดถอนสะสม\n' +
    'เริ่มแจ้งเตือน / หยุดแจ้งเตือน — เปิดหรือปิดสรุปรายวัน\n' +
    '/cancel — ยกเลิกขั้นตอนปัจจุบัน\n' +
    '/help — แสดงความช่วยเหลือนี้';
}

function topicPrompt_() {
  return '📋 เมนูหลัก — พิมพ์หมายเลข:\n' +
    '📅 เรื่องนัดหมาย\n' +
    '1 — บันทึก/จัดการนัดหมาย\n' +
    '2 — เช็คตารางวันนี้\n' +
    '3 — เช็คตารางวันพรุ่งนี้\n\n' +
    '💰 เรื่องเงิน\n' +
    '4 — เมนูเงินเก็บ\n' +
    '5 — ฝากเงิน\n' +
    '6 — ถอนเงิน\n' +
    '7 — เช็คยอดเงิน balance\n\n' +
    '9 — 🚪 ออกจากการใช้งานบอท';
}

function appointmentMenu_() {
  return '📅 หัวข้อนัดหมาย — พิมพ์เลขหรือชื่อคำสั่ง:\n' +
    '1 — ➕ บันทึกนัดหมาย\n2 — 🗓️ ดูงานวันนี้\n3 — 🧾 สรุปงานตามวันที่\n4 — 🗑️ ลบนัดหมาย\n0 — ↩️ กลับไปเลือกหัวข้อหลัก\n9 — 🚪 ออกจากการใช้งานบอท';
}

function savingsMenu_() {
  return '💰 หัวข้อเงินเก็บ — พิมพ์เลขหรือชื่อคำสั่ง:\n' +
    '1 — 📥 บันทึกเงินฝาก\n2 — 📤 บันทึกเงินถอน\n3 — 💵 ดูยอด balance\n4 — 🧾 ดูรายละเอียด 15 วันย้อนหลัง\n0 — ↩️ กลับไปเลือกหัวข้อหลัก\n9 — 🚪 ออกจากการใช้งานบอท';
}

function afterSavePrompt_(topic) {
  if (topic === 'เงินเก็บ') {
    return '💰 เมนูหลังบันทึก — เงินเก็บ\n' +
      '1 — 📥 บันทึกเงินฝาก\n' +
      '2 — 📤 บันทึกเงินถอน\n' +
      '3 — 💵 ดูยอด balance\n' +
      '4 — 🧾 ดูรายละเอียด 15 วันย้อนหลัง\n' +
      '0 — ↩️ กลับไปเลือกหัวข้อหลัก\n' +
      '9 — 🚪 ออกจากการใช้งานบอท';
  }
  return '📅 เมนูหลังบันทึก — ' + topic + '\n' +
    '1 — ➕ บันทึก ' + topic + ' ต่อ\n' +
    '2 — ↩️ กลับหน้าหลัก เพื่อเลือกหัวข้ออื่น\n' +
    '3 — 🚪 ออกจากการใช้งานบอท';
}

function handleRecordingFlow_(target, text) {
  const lower = text.toLowerCase();
  if (lower === '/help' || text === 'ช่วยเหลือ') return null;
  const state = getRecordingState_(target.ownerKey);
  if (!state) return null;

  if (lower === '/cancel') {
    setRecordingState_(target, 'SELECT_TOPIC');
    return 'ยกเลิกขั้นตอนนี้แล้วครับ\n' + topicPrompt_();
  }
  if (text === '0' && ['APPOINTMENT', 'APPOINTMENT_REMINDER', 'APPOINTMENT_DELETE_PICK', 'SAVINGS_DEPOSIT', 'SAVINGS_WITHDRAW'].includes(state)) {
    setPendingAppointment_(target.ownerKey, null);
    const menuState = state === 'SAVINGS_DEPOSIT' || state === 'SAVINGS_WITHDRAW' ? 'SAVINGS_MENU' : 'APPOINTMENT_MENU';
    setRecordingState_(target, menuState);
    return 'ออกจากการกรอกแล้วครับ\n' + (menuState === 'SAVINGS_MENU' ? savingsMenu_() : appointmentMenu_());
  }

  if (state === 'POST_SAVE_APPOINTMENT' || state === 'POST_SAVE_SAVINGS') {
    const returnState = state === 'POST_SAVE_APPOINTMENT' ? 'APPOINTMENT_MENU' : 'SAVINGS_MENU';
    if (state === 'POST_SAVE_SAVINGS') {
      if (text === '1' || ['ฝาก', 'บันทึกเงินฝาก'].includes(text)) {
        setRecordingState_(target, 'SAVINGS_DEPOSIT');
        return 'ส่งรายละเอียดเงินฝาก โดยระบบบันทึกวันที่ปัจจุบันให้อัตโนมัติ:\nจำนวนเงิน | หมายเหตุ\nตัวอย่าง: 500 | เก็บเงินเดือน';
      }
      if (text === '2' || ['ถอน', 'บันทึกเงินถอน'].includes(text)) {
        setRecordingState_(target, 'SAVINGS_WITHDRAW');
        return 'ส่งรายละเอียดเงินถอน โดยระบบบันทึกวันที่ปัจจุบันให้อัตโนมัติ:\nจำนวนเงิน | หมายเหตุ\nตัวอย่าง: 200 | ใช้จ่าย';
      }
      if (text === '3') return getSavingsBalance_(target.ownerKey) + '\n\n' + afterSavePrompt_('เงินเก็บ');
      if (text === '4') return getRecentSavings_(target.ownerKey) + '\n\n' + afterSavePrompt_('เงินเก็บ');
      if (text === '0') {
        setRecordingState_(target, 'SELECT_TOPIC');
        return topicPrompt_();
      }
      if (text === '9') {
        endChatSession_(target);
        return 'ออกจากการใช้งานบอทแล้วครับ พิมพ์ “hi bot” เมื่อต้องการเริ่มใหม่';
      }
      return afterSavePrompt_('เงินเก็บ');
    }
    if (text === '1' || text === 'บันทึกต่อ') {
      setRecordingState_(target, returnState);
      return returnState === 'APPOINTMENT_MENU' ? appointmentMenu_() : savingsMenu_();
    }
    if (text === '2' || text === 'หน้าหลัก' || text === 'กลับหน้าหลัก') {
      setRecordingState_(target, 'SELECT_TOPIC');
      return topicPrompt_();
    }
    if (text === '3' || text === 'ออก' || text === 'ออกจากบอท') {
      endChatSession_(target);
      return 'ออกจากการใช้งานบอทแล้วครับ พิมพ์ “hi bot” เมื่อต้องการเริ่มใหม่';
    }
    return afterSavePrompt_(state === 'POST_SAVE_APPOINTMENT' ? 'นัดหมาย' : 'เงินเก็บ');
  }

  if (state === 'SELECT_TOPIC') {
    if (text === '1' || text === 'นัดหมาย') {
      setRecordingState_(target, 'APPOINTMENT_MENU');
      return appointmentMenu_();
    }
    if (['2', 'เช็คตารางวันนี้', 'ตารางวันนี้'].includes(text)) {
      return listToday_(target.ownerKey) + '\n\n' + topicPrompt_();
    }
    if (['3', 'เช็คตารางวันพรุ่งนี้', 'ตารางวันพรุ่งนี้'].includes(text)) {
      return listTomorrow_(target.ownerKey) + '\n\n' + topicPrompt_();
    }
    if (text === '4' || text === 'เงินเก็บ') {
      setRecordingState_(target, 'SAVINGS_MENU');
      return savingsMenu_();
    }
    if (text === '5' || ['ฝากเงิน', 'ฝาก'].includes(text)) {
      setRecordingState_(target, 'SAVINGS_DEPOSIT');
      return '💰 ฝากเงิน — วันที่ปัจจุบันจะบันทึกให้อัตโนมัติ\nกรอก: จำนวนเงิน | หมายเหตุ\nตัวอย่าง: 500 | เก็บเงินเดือน\n0 — ออกจากการกรอก';
    }
    if (text === '6' || ['ถอนเงิน', 'ถอน'].includes(text)) {
      setRecordingState_(target, 'SAVINGS_WITHDRAW');
      return '💰 ถอนเงิน — วันที่ปัจจุบันจะบันทึกให้อัตโนมัติ\nกรอก: จำนวนเงิน | หมายเหตุ\nตัวอย่าง: 200 | ใช้จ่าย\n0 — ออกจากการกรอก';
    }
    if (['7', 'เช็คยอดเงินฝาก balance', 'เช็คยอดเงิน balance', 'ยอดเงินเก็บ', 'ยอดคงเหลือ', 'ยอด balance', 'balance', '/balance'].includes(text.toLowerCase())) {
      return getSavingsBalance_(target.ownerKey) + '\n\n' + topicPrompt_();
    }
    if (text === '9') {
      endChatSession_(target);
      return 'ออกจากการใช้งานบอทแล้วครับ พิมพ์ “hi bot” เมื่อต้องการเริ่มใหม่';
    }
    return 'กรุณาเลือกหมายเลขจากเมนูครับ\n' + topicPrompt_();
  }

  if (state === 'APPOINTMENT_MENU') {
    if (['1', 'บันทึกนัดหมาย'].includes(text)) {
      setRecordingState_(target, 'APPOINTMENT');
      return 'กรอกให้ครบ 4 ช่องตามรูปแบบนี้ครับ:\nYYYY-MM-DD | HH:mm | รายละเอียด | สถานที่\nถ้าไม่มีเวลา/สถานที่ ให้เว้นช่องว่างแต่คงเครื่องหมาย | ไว้\nตัวอย่างไม่มีเวลาและสถานที่: 2026-10-03 |  | ประชุมทีม |';
    }
    if (['2', 'วันนี้', 'ดูงานวันนี้'].includes(text)) return listToday_(target.ownerKey) + '\n\n' + appointmentMenu_();
    if (['3', 'สรุปงาน'].includes(text)) {
      setRecordingState_(target, 'APPOINTMENT_SUMMARY_DATE');
      return 'ส่งวันที่ที่ต้องการสรุปในรูปแบบ YYYY-MM-DD เช่น 2026-10-03';
    }
    if (['4', 'ลบงาน'].includes(text)) {
      return startAppointmentDelete_(target.ownerKey);
    }
    if (text === '0') {
      setRecordingState_(target, 'SELECT_TOPIC');
      return topicPrompt_();
    }
    if (text === '9') {
      endChatSession_(target);
      return 'ออกจากการใช้งานบอทแล้วครับ พิมพ์ “hi bot” เมื่อต้องการเริ่มใหม่';
    }
    return 'เลือก 1, 2, 3, 4, 0 หรือ 9 ครับ\n' + appointmentMenu_();
  }

  if (state === 'SAVINGS_MENU') {
    if (['1', 'ฝาก', 'บันทึกเงินฝาก'].includes(text)) {
      setRecordingState_(target, 'SAVINGS_DEPOSIT');
      return 'ส่งรายละเอียดเงินฝาก โดยระบบบันทึกวันที่ปัจจุบันให้อัตโนมัติ:\nจำนวนเงิน | หมายเหตุ\nตัวอย่าง: 500 | เก็บเงินเดือน';
    }
    if (['2', 'ถอน', 'บันทึกเงินถอน'].includes(text)) {
      setRecordingState_(target, 'SAVINGS_WITHDRAW');
      return 'ส่งรายละเอียดเงินถอน โดยระบบบันทึกวันที่ปัจจุบันให้อัตโนมัติ:\nจำนวนเงิน | หมายเหตุ\nตัวอย่าง: 200 | ใช้จ่าย';
    }
    if (['3', 'ยอดเงินเก็บ', 'ยอดคงเหลือ', 'ยอด balance', 'balance', '/balance'].includes(text.toLowerCase())) {
      return getSavingsBalance_(target.ownerKey) + '\n\n' + savingsMenu_();
    }
    if (['4', 'ย้อนหลัง 15 วัน', 'รายละเอียด 15 วันย้อนหลัง', 'รายการเงินเก็บย้อนหลัง'].includes(text.toLowerCase())) {
      return getRecentSavings_(target.ownerKey) + '\n\n' + savingsMenu_();
    }
    if (text === '0') {
      setRecordingState_(target, 'SELECT_TOPIC');
      return topicPrompt_();
    }
    if (text === '9') {
      endChatSession_(target);
      return 'ออกจากการใช้งานบอทแล้วครับ พิมพ์ “hi bot” เมื่อต้องการเริ่มใหม่';
    }
    return 'เลือก 1, 2, 3, 4, 0 หรือ 9 ครับ\n' + savingsMenu_();
  }

  if (state === 'APPOINTMENT') {
    const appointment = parseAppointment_(text);
    if (!appointment) return appointmentFormatError_();
    setPendingAppointment_(target.ownerKey, appointment);
    setRecordingState_(target, 'APPOINTMENT_REMINDER');
    return 'ต้องการให้แจ้งเตือนเมื่อใด? (ค่าเริ่มต้น: วันที่ระบุในตาราง)\n1 — ในวันที่ระบุในตาราง\n2 — ล่วงหน้า 1 วัน';
  }

  if (state === 'APPOINTMENT_REMINDER') {
      if (text !== '1' && text !== '2') return 'กรุณาเลือก 1 แจ้งเตือนในวันที่ระบุ (ค่าเริ่มต้น) หรือ 2 แจ้งเตือนล่วงหน้า 1 วัน\n0 — ออกจากการกรอกกลับเมนูนัดหมาย';
    const appointment = getPendingAppointment_(target.ownerKey);
    if (!appointment) {
      setRecordingState_(target, 'APPOINTMENT_MENU');
      return 'ไม่พบข้อมูลนัดหมายที่รอบันทึก กรุณาเริ่มรายการใหม่';
    }
    const reminderMode = text === '2' ? 'ADVANCE_1_DAY' : 'ON_DATE';
    const result = saveAppointment_(target, appointment, reminderMode);
    setPendingAppointment_(target.ownerKey, null);
    setRecordingState_(target, 'POST_SAVE_APPOINTMENT');
    return result + '\nแจ้งเตือน: ' + (text === '2' ? 'ล่วงหน้า 1 วัน' : 'ในวันที่ระบุ') +
      '\n\n' + afterSavePrompt_('นัดหมาย');
  }

  if (state === 'SAVINGS_DEPOSIT' || state === 'SAVINGS_WITHDRAW') {
    const type = state === 'SAVINGS_DEPOSIT' ? 'ฝาก' : 'ถอน';
    const parts = text.split('|').map(s => s.trim());
    let savingsEntryLine = text;
    if ((state === 'SAVINGS_DEPOSIT' || state === 'SAVINGS_WITHDRAW') && parts.length === 2) {
      const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
      savingsEntryLine = [today, type, parts[0], parts[1]].join(' | ');
    }
    const result = addSavings_(target, savingsEntryLine);
    if (result.startsWith('บันทึกเงินเก็บแล้ว')) {
      setRecordingState_(target, 'POST_SAVE_SAVINGS');
      return result + '\n\n' + getSavingsBalance_(target.ownerKey) + '\n\n' + afterSavePrompt_('เงินเก็บ');
    }
    return result + '\n0 — ออกจากการกรอกกลับเมนูเงินเก็บ';
  }

  if (state === 'APPOINTMENT_SUMMARY_DATE') {
    const result = summarizeDate_(target.ownerKey, 'สรุปงาน ' + text);
    setRecordingState_(target, 'APPOINTMENT_MENU');
    return result + '\n\n' + appointmentMenu_();
  }

  if (state === 'APPOINTMENT_DELETE_PICK') {
    const pending = getPendingAppointment_(target.ownerKey);
    const ids = pending && pending.deleteIds;
    const choice = Number(text);
    if (!ids || !ids.length) {
      setPendingAppointment_(target.ownerKey, null);
      setRecordingState_(target, 'APPOINTMENT_MENU');
      return 'ไม่พบรายการที่เลือกลบได้ครับ\n' + appointmentMenu_();
    }
    if (!Number.isInteger(choice) || choice < 1 || choice > ids.length) {
      return appointmentDeleteList_(target.ownerKey);
    }
    const result = deleteTask_(target.ownerKey, 'ลบงาน ' + ids[choice - 1]);
    setPendingAppointment_(target.ownerKey, null);
    setRecordingState_(target, 'APPOINTMENT_MENU');
    return result + '\n\n' + appointmentMenu_();
  }

  // Clear stale states from older script versions.
  setRecordingState_(target, 'SELECT_TOPIC');
  return topicPrompt_();
}

function isValidDate_(dateText) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) return false;
  const date = new Date(dateText + 'T00:00:00+07:00');
  return !isNaN(date.getTime()) && Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM-dd') === dateText;
}

function addTask_(target, text) {
  const prefix = 'เพิ่มงาน | ';
  if (!text.startsWith(prefix)) return appointmentFormatError_();
  const appointment = parseAppointment_(text.slice(prefix.length));
  if (!appointment) return appointmentFormatError_();
  return saveAppointment_(target, appointment, 'ON_DATE');
}

function parseAppointment_(text) {
  const parts = text.split('|').map(s => s.trim());
  const timePattern = /^([01]\d|2[0-3]):[0-5]\d$/;
  if (parts.length !== 4 || !isValidDate_(parts[0]) ||
      (parts[1] !== '' && !timePattern.test(parts[1])) || !parts[2]) return null;
  return {date: parts[0], time: parts[1], description: parts[2], location: parts[3]};
}

function saveAppointment_(target, appointment, reminderMode) {
  const date = appointment.date;
  const time = appointment.time;
  const description = appointment.description;
  const location = appointment.location;
  const id = Utilities.getUuid().slice(0, 8);
  SpreadsheetApp.getActive().getSheetByName(CONFIG.TASKS_SHEET)
    .appendRow([id, target.ownerKey, description, date, time, 'OPEN', new Date(), location, reminderMode]);
  upsertChat_(target, true);
  return '✅ บันทึกนัดหมายแล้วครับ (รหัส ' + id + ')\n📅 วันที่: ' + date +
    (time ? '\n🕒 เวลา: ' + time : '') + '\n📝 รายละเอียด: ' + description +
    (location ? '\n📍 สถานที่: ' + location : '');
}

function appointmentFormatError_() {
  return 'รูปแบบไม่ถูกต้องครับ ต้องกรอกให้ครบ 4 ช่อง:\nYYYY-MM-DD | HH:mm | รายละเอียด | สถานที่\nถ้าไม่มีเวลา/สถานที่ ให้เว้นช่องว่างแต่คงเครื่องหมาย | ไว้\nตัวอย่าง: 2026-10-03 |  | ประชุมทีม |\n0 — ออกจากการกรอกกลับเมนูนัดหมาย';
}

function listToday_(ownerKey) {
  return listTasksOnDate_(ownerKey, new Date(), 'งานวันนี้');
}

function listTomorrow_(ownerKey) {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  return listTasksOnDate_(ownerKey, tomorrow, 'งานวันพรุ่งนี้');
}

function listTasksOnDate_(ownerKey, dateValue, label) {
  const date = Utilities.formatDate(dateValue, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const rows = SpreadsheetApp.getActive().getSheetByName(CONFIG.TASKS_SHEET).getDataRange().getValues().slice(1);
  const items = rows.filter(r => String(r[1]) === ownerKey && cellDateKey_(r[3]) === date && String(r[5]).toUpperCase() === 'OPEN')
    .sort((a, b) => String(a[4]).localeCompare(String(b[4])));
  return items.length ? '📅 ' + label + ' (' + date + ')\n' + items.map(formatTask_).join('\n') : '📅 ไม่มีนัดหมายที่บันทึกไว้ใน' + label + ' (' + date + ') ครับ';
}

function formatTask_(row) {
  return '📌 นัดหมาย: ' + (row[4] ? row[4] + ' — ' : '') + row[2] +
    (row[7] ? '\n   📍 สถานที่: ' + row[7] : '') + '\n   🆔 รหัส: ' + row[0];
}

function cellDateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  const iso = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (iso) return iso[1];
  // Support legacy Sheets rows where a date was stored as a localized string.
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? '' : Utilities.formatDate(parsed, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function summarizeDate_(ownerKey, text) {
  const match = text.trim().match(/^สรุปงาน\s+(\d{4}-\d{2}-\d{2})$/);
  if (!match || !isValidDate_(match[1])) return 'วันที่ไม่ถูกต้องครับ ใช้รูปแบบ YYYY-MM-DD เช่น 2026-10-03';
  const date = match[1];
  const rows = SpreadsheetApp.getActive().getSheetByName(CONFIG.TASKS_SHEET).getDataRange().getValues().slice(1);
  const items = rows.filter(r => String(r[1]) === ownerKey && cellDateKey_(r[3]) === date && String(r[5]).toUpperCase() === 'OPEN')
    .sort((a, b) => String(a[4]).localeCompare(String(b[4])));
  return items.length ? '📅 สรุปงานวันที่ ' + date + ' (' + items.length + ' งาน)\n' + items.map(formatTask_).join('\n')
    : '📅 ไม่มีงานที่ยังใช้งานอยู่ในวันที่ ' + date + ' ครับ';
}

function deleteTask_(ownerKey, text) {
  const parts = text.trim().split(/\s+/);
  if (parts.length !== 2 || parts[0] !== 'ลบงาน') return 'รูปแบบไม่ถูกต้องครับ ใช้: ลบงาน <รหัสงาน>';
  const taskId = parts[1].toLowerCase();
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.TASKS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === taskId && String(rows[i][1]) === ownerKey && rows[i][5] === 'OPEN') {
      sheet.getRange(i + 1, 6).setValue('DELETED');
      return 'ลบนัดหมายรหัส ' + rows[i][0] + ' เรียบร้อยแล้วครับ';
    }
  }
  return 'ไม่พบงานที่ยังใช้งานอยู่รหัสนี้ในแชตปัจจุบันครับ';
}

function startAppointmentDelete_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.TASKS_SHEET);
  const rows = sheet.getDataRange().getValues().slice(1)
    .filter(function(row) { return String(row[1]) === ownerKey && String(row[5]).toUpperCase() === 'OPEN'; })
    .sort(function(a, b) {
      return cellDateKey_(a[3]).localeCompare(cellDateKey_(b[3])) || String(a[4]).localeCompare(String(b[4]));
    });
  if (!rows.length) return 'ไม่มีนัดหมายที่ยังใช้งานอยู่ให้ลบครับ\n\n' + appointmentMenu_();
  setPendingAppointment_(ownerKey, {deleteIds: rows.map(function(row) { return String(row[0]); })});
  setRecordingState_(getTargetFromOwner_(ownerKey), 'APPOINTMENT_DELETE_PICK');
  return '🗑️ เลือกนัดหมายที่ต้องการลบ โดยพิมพ์หมายเลข:\n' +
    rows.map(function(row, index) { return (index + 1) + '. ' + cellDateKey_(row[3]) + ' | ' + formatTask_(row); }).join('\n') +
    '\n0 — ยกเลิกและกลับเมนูนัดหมาย';
}

function getTargetFromOwner_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === ownerKey) {
      return {ownerKey: ownerKey, targetId: String(rows[i][4] || ownerKey), targetType: String(rows[i][5] || 'user')};
    }
  }
  return {ownerKey: ownerKey, targetId: ownerKey, targetType: 'user'};
}

function appointmentDeleteList_(ownerKey) {
  const pending = getPendingAppointment_(ownerKey);
  if (!pending || !pending.deleteIds) return 'ไม่พบรายการที่เลือกลบได้ครับ';
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.TASKS_SHEET);
  const rows = sheet.getDataRange().getValues().slice(1)
    .filter(function(row) { return pending.deleteIds.indexOf(String(row[0])) >= 0 && String(row[1]) === ownerKey && String(row[5]).toUpperCase() === 'OPEN'; });
  return rows.length
    ? '🗑️ เลือกหมายเลขนัดหมายที่ต้องการลบ:\n' + rows.map(function(row, index) { return (index + 1) + '. ' + cellDateKey_(row[3]) + ' | ' + formatTask_(row); }).join('\n') + '\n0 — ยกเลิกและกลับเมนูนัดหมาย'
    : 'ไม่มีนัดหมายที่ยังใช้งานอยู่ให้ลบครับ';
}

function addSavings_(target, text) {
  const parts = text.split('|').map(s => s.trim());
  if (parts.length !== 4 || !isValidDate_(parts[0]) || !['ฝาก', 'ถอน'].includes(parts[1]) ||
      !/^\d+(\.\d{1,2})?$/.test(parts[2]) || Number(parts[2]) <= 0 || !parts[3]) {
    return 'รูปแบบไม่ถูกต้องครับ ใช้: YYYY-MM-DD | ฝาก/ถอน | จำนวนเงิน | หมายเหตุ\nตัวอย่าง: 2026-10-03 | ฝาก | 500 | เก็บเงินเดือน';
  }
  const id = Utilities.getUuid().slice(0, 8);
  SpreadsheetApp.getActive().getSheetByName(CONFIG.SAVINGS_SHEET)
    .appendRow([id, target.ownerKey, parts[0], parts[1], Number(parts[2]), parts[3], new Date()]);
  return '✅ บันทึกเงินเก็บแล้วครับ (รหัส ' + id + ')\n📅 วันที่: ' + parts[0] +
    '\n' + (parts[1] === 'ฝาก' ? '📥 ฝาก' : '📤 ถอน') + ' ' + formatMoney_(Number(parts[2])) + ' บาท' +
    '\n📝 หมายเหตุ: ' + parts[3];
}

function isBalanceCommand_(text) {
  return ['ยอดเงินเก็บ', 'ยอดคงเหลือ', 'ยอด balance', 'balance', '/balance'].includes(text.toLowerCase());
}

function getSavingsBalance_(ownerKey) {
  const rows = SpreadsheetApp.getActive().getSheetByName(CONFIG.SAVINGS_SHEET).getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) === ownerKey);
  let deposits = 0;
  let withdrawals = 0;
  rows.forEach(r => {
    const amount = Number(r[4]) || 0;
    if (r[3] === 'ฝาก') deposits += amount;
    else if (r[3] === 'ถอน') withdrawals += amount;
  });
  return '💰 ยอดเงินเก็บของแชตนี้\n💵 Balance คงเหลือ: ' + formatMoney_(deposits - withdrawals) + ' บาท\n' +
    '📥 ยอดฝากสะสม: ' + formatMoney_(deposits) + ' บาท\n📤 ยอดถอนสะสม: ' + formatMoney_(withdrawals) + ' บาท';
}

function getRecentSavings_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SAVINGS_SHEET);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const dateKey = cellDateKey_(row[2]);
    if (String(row[1]) === ownerKey && isValidDate_(dateKey)) {
      rows.push({row: row, dateKey: dateKey, date: new Date(dateKey + 'T00:00:00+07:00')});
    }
  }
  const today = new Date(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd') + 'T00:00:00+07:00');
  const start = new Date(today);
  start.setDate(start.getDate() - 14); // today plus the preceding 14 days = 15 calendar days
  const recent = rows.filter(function(item) { return item.date >= start && item.date <= today; });
  recent.sort(function(a, b) {
    return b.date - a.date || new Date(b.row[6]).getTime() - new Date(a.row[6]).getTime();
  });
  if (!recent.length) return 'ไม่มีรายการเงินเก็บในช่วง 15 วันล่าสุดครับ';
  const totalDeposit = recent.reduce((sum, item) => sum + (item.row[3] === 'ฝาก' ? Number(item.row[4]) || 0 : 0), 0);
  const totalWithdrawal = recent.reduce((sum, item) => sum + (item.row[3] === 'ถอน' ? Number(item.row[4]) || 0 : 0), 0);
  return 'รายการเงินเก็บย้อนหลัง 15 วัน (' + Utilities.formatDate(start, CONFIG.TIMEZONE, 'yyyy-MM-dd') +
    ' ถึง ' + Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM-dd') + ')\n' +
    recent.map(item => (item.row[3] === 'ฝาก' ? '📥 ฝาก' : '📤 ถอน') + ' | ' + item.dateKey + ' | ' +
      formatMoney_(Number(item.row[4]) || 0) + ' บาท | 📝 ' + item.row[5]).join('\n') +
    '\n\n📥 รวมฝาก: ' + formatMoney_(totalDeposit) + ' บาท\n📤 รวมถอน: ' + formatMoney_(totalWithdrawal) + ' บาท';
}

function formatMoney_(amount) {
  return amount.toLocaleString('th-TH', {minimumFractionDigits: 0, maximumFractionDigits: 2});
}

function sendDailyReminders() {
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const ss = SpreadsheetApp.getActive();
  const tasks = ss.getSheetByName(CONFIG.TASKS_SHEET).getDataRange().getValues().slice(1);
  const chats = ss.getSheetByName(CONFIG.USERS_SHEET).getDataRange().getValues().slice(1);
  chats.filter(row => row[1] === true || String(row[1]).toLowerCase() === 'true').forEach(row => {
    const ownerKey = String(row[0]);
    const items = tasks.filter(t => String(t[1]) === ownerKey && t[5] === 'OPEN' && reminderDate_(t) === today)
      .sort((a, b) => String(a[4]).localeCompare(String(b[4])));
    if (items.length) {
      const recipient = String(row[4] || row[0]);
      pushLine_(recipient, '📅 แจ้งเตือนนัดหมาย (' + today + ')\n' + items.map(formatTask_).join('\n'));
    }
  });
}

function reminderDate_(taskRow) {
  const mode = String(taskRow[8] || 'ON_DATE');
  const scheduledDate = String(taskRow[3]);
  if (mode !== 'ADVANCE_1_DAY') return scheduledDate;
  const date = new Date(scheduledDate + 'T00:00:00+07:00');
  date.setDate(date.getDate() - 1);
  return Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function replyLine_(replyToken, message) {
  if (!replyToken) return;
  lineRequest_('https://api.line.me/v2/bot/message/reply', {
    replyToken: replyToken,
    messages: [{type: 'text', text: message}]
  });
}

function pushLine_(recipientId, message) {
  lineRequest_('https://api.line.me/v2/bot/message/push', {
    to: recipientId,
    messages: [{type: 'text', text: message}]
  });
}

function lineRequest_(url, payload) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error('Missing LINE_CHANNEL_ACCESS_TOKEN in Script Properties');
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + token},
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (response.getResponseCode() >= 300) {
    console.error('LINE API error ' + response.getResponseCode() + ': ' + response.getContentText());
  }
}

