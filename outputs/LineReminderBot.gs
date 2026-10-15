/**
 * LINE group-aware reminder bot backed by Google Sheets.
 * Commands can start a chat directly; "hi bot" shows the main menu.
 * Configure Script Properties: LINE_CHANNEL_ACCESS_TOKEN, WEBHOOK_KEY, BOT_ADMIN_USER_IDS.
 */
const CONFIG = {
  TASKS_SHEET: 'Tasks',
  SAVINGS_SHEET: 'Savings',
  INSTALLMENTS_SHEET: 'Installments',
  USERS_SHEET: 'Users',
  PERMISSIONS_SHEET: 'Permissions',
  TIMEZONE: 'Asia/Bangkok',
  DAILY_HOUR: 6
};

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, CONFIG.TASKS_SHEET,
    ['taskId', 'scopeKey', 'รายละเอียด', 'วัน', 'เวลา', 'สถานะ', 'สร้างเมื่อ', 'สถานที่', 'รูปแบบเตือน']);
  ensureSheet_(ss, CONFIG.SAVINGS_SHEET,
    ['recordId', 'scopeKey', 'วัน', 'ประเภท', 'จำนวนเงิน', 'หมายเหตุ', 'สร้างเมื่อ']);
  ensureSheet_(ss, CONFIG.INSTALLMENTS_SHEET,
    ['recordId', 'scopeKey', 'วัน', 'ประเภท', 'เงินต้น', 'อัตราดอกเบี้ย(%)', 'ดอกเบี้ย(บันทึก)', 'เงินต้นคงเหลือ', 'หมายเหตุ', 'สร้างเมื่อ', 'หัวเรื่อง', 'เจ้าหนี้', 'ลูกหนี้', 'ฐานดอกเบี้ย', 'loanId']);
  ensureSheet_(ss, CONFIG.USERS_SHEET,
    ['scopeKey', 'รับแจ้งเตือน', 'สมัครเมื่อ', 'ขั้นตอนบันทึก', 'targetId', 'targetType', 'sessionActive', 'lastActivityAt', 'pendingAppointment', 'pendingPermission', 'groupName', 'selectedInstallmentId']);
  ensureSheet_(ss, CONFIG.PERMISSIONS_SHEET,
    ['scopeKey', 'userId', 'appointmentAccess', 'financeAccess', 'updatedAt', 'updatedBy', 'installmentAccess']);
  ScriptApp.getProjectTriggers()
    .filter(t => ['sendDailyReminders', 'expireIdleSessions', 'accrueAnnualInterest'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyReminders').timeBased().everyDays(1)
    .atHour(CONFIG.DAILY_HOUR).inTimezone(CONFIG.TIMEZONE).create();
  ScriptApp.newTrigger('accrueAnnualInterest').timeBased().everyDays(1)
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

function isAdmin_(actorId) {
  if (!actorId) return false;
  const raw = PropertiesService.getScriptProperties().getProperty('BOT_ADMIN_USER_IDS') || '';
  return raw.split(',').map(function(id) { return id.trim(); }).includes(actorId);
}

function getPermissionRecord_(scopeKey, userId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.PERMISSIONS_SHEET);
  if (!sheet) return {sheet: null, rowIndex: 0, row: null};
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === scopeKey && String(rows[i][1] || '') === String(userId || '')) {
      return {sheet: sheet, rowIndex: i + 1, row: rows[i]};
    }
  }
  return {sheet: sheet, rowIndex: 0, row: null};
}

function hasTopicAccess_(target, topic) {
  const isAdminPersonalChat = target.targetType === 'user' && isAdmin_(target.actorId);
  const isAdminInGroup = target.targetType !== 'user' && isAdmin_(target.actorId);
  if (isAdminPersonalChat) return true;
  if (!target.actorId) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.PERMISSIONS_SHEET);
    if (sheet) {
      const rows = sheet.getDataRange().getValues();
      const column = permissionColumnIndex_(topic);
      const hasMemberOverride = rows.slice(1).some(function(row) {
        return String(row[0]) === target.ownerKey && String(row[1] || '') && String(row[column] || '').toUpperCase();
      });
      if (hasMemberOverride) return false;
    }
  }
  const userRecord = target.actorId && !isAdminInGroup ? getPermissionRecord_(target.ownerKey, target.actorId) : null;
  const column = permissionColumnIndex_(topic);
  if (userRecord && userRecord.row && String(userRecord.row[column]).toUpperCase()) {
    return String(userRecord.row[column]).toUpperCase() === 'ALLOW';
  }
  const groupRecord = getPermissionRecord_(target.ownerKey, '');
  if (groupRecord.row && String(groupRecord.row[column]).toUpperCase()) {
    return String(groupRecord.row[column]).toUpperCase() === 'ALLOW';
  }
  return true;
}

function permissionColumnIndex_(topic) {
  return topic === 'APPOINTMENT' ? 2 : topic === 'FINANCE' ? 3 : 6;
}

function setTopicAccess_(scopeKey, userId, topic, value, updatedBy) {
  const record = getPermissionRecord_(scopeKey, userId);
  if (!record.sheet) throw new Error('Run setup() to create the Permissions sheet before changing access.');
  const columns = topic === 'ALL' ? [3, 4, 7] : [permissionColumnIndex_(topic) + 1];
  if (!record.row) {
    const row = [scopeKey, userId || '', '', '', new Date(), updatedBy || '', ''];
    columns.forEach(function(column) { row[column - 1] = value; });
    record.sheet.appendRow(row);
    return;
  }
  columns.forEach(function(column) { record.sheet.getRange(record.rowIndex, column).setValue(value); });
  record.sheet.getRange(record.rowIndex, 5).setValue(new Date());
  record.sheet.getRange(record.rowIndex, 6).setValue(updatedBy || '');
}

function clearTopicAccess_(scopeKey, userId, topic) {
  const record = getPermissionRecord_(scopeKey, userId);
  if (!record.sheet) throw new Error('Run setup() to create the Permissions sheet before changing access.');
  if (!record.row) return;
  const columns = topic === 'ALL' ? [3, 4, 7] : [permissionColumnIndex_(topic) + 1];
  columns.forEach(function(column) { record.sheet.getRange(record.rowIndex, column).clearContent(); });
  record.sheet.getRange(record.rowIndex, 5).setValue(new Date());
}

function topicForStateOrCommand_(state, text, target) {
  if (state === 'TRANSACTION_DATE_PICK') {
    const pending = target ? getPendingAppointment_(target.ownerKey) : null;
    return pending && pending.resumeState ? topicForStateOrCommand_(pending.resumeState, '', target) : '';
  }
  if (['APPOINTMENT_MENU', 'APPOINTMENT', 'APPOINTMENT_DATE', 'APPOINTMENT_FIELDS',
      'APPOINTMENT_REMINDER', 'APPOINTMENT_SUMMARY_DATE', 'APPOINTMENT_DELETE_PICK',
      'POST_SAVE_APPOINTMENT'].includes(state)) return 'APPOINTMENT';
  if (['SAVINGS_MENU', 'SAVINGS_DEPOSIT', 'SAVINGS_WITHDRAW', 'POST_SAVE_SAVINGS'].includes(state)) return 'FINANCE';
  if (['INSTALLMENT_MENU', 'INSTALLMENT_PRINCIPAL', 'INSTALLMENT_PAYMENT', 'INSTALLMENT_INTEREST_PAYMENT', 'INSTALLMENT_INTEREST_CALC', 'INSTALLMENT_LOAN_PICK', 'POST_SAVE_INSTALLMENT'].includes(state)) return 'INSTALLMENT';
  if (state !== 'SELECT_TOPIC') return '';
  const command = String(text || '').trim().toUpperCase();
  if (/^(C1|C2|C3)(?:\s|$)/.test(command)) return 'APPOINTMENT';
  if (/^M[1-4](?:\s|$)/.test(command)) return 'FINANCE';
  if (/^L1(?:\s|$)/.test(command)) return 'INSTALLMENT';
  if (command.startsWith('เงินตั้งต้นผ่อน')) return 'INSTALLMENT';
  if (['C1', 'C2', 'C3', 'นัดหมาย', 'สรุปงานตามวันที่', 'วันนี้'].includes(command) || command.startsWith('เพิ่มงาน') || command.startsWith('สรุปงาน') || command.startsWith('ลบงาน')) return 'APPOINTMENT';
  if (['M1', 'M2', 'M3', 'M4', 'เงินเก็บ', 'ฝากเงิน', 'ฝาก', 'ถอนเงิน', 'ถอน',
      'เช็คยอดเงินฝาก BALANCE', 'เช็คยอดเงิน BALANCE', 'ยอดเงินเก็บ', 'ยอดคงเหลือ',
      'ยอด BALANCE', 'BALANCE', '/BALANCE'].includes(command)) return 'FINANCE';
  if (/^L[12](?:\s|$)/.test(command) || ['L1', 'L2', 'L3', 'ผ่อนเงิน', 'ตั้งยอดผ่อน', 'บันทึกยอดตั้งต้น', 'บันทึกค่างวด', 'ชำระค่างวด', 'ยอดผ่อน', 'สรุปผ่อน'].includes(command)) return 'INSTALLMENT';
  if (['ดอกเบี้ยค้าง', 'ดอกเบี้ยคงค้าง', 'สอบถามดอกเบี้ย', 'สอบถามดอกเบี้ยค้างชำระ', 'ดูดอกเบี้ยค้างชำระ'].includes(command)) return 'INSTALLMENT';
  return '';
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
  if (event.type === 'postback') {
    handlePostback_(event);
    return;
  }
  if (event.type !== 'message' || !event.message || event.message.type !== 'text') return;
  const target = getTarget_(event.source);
  if (!target) return;
  const text = event.message.text.trim();
  const lower = text.toLowerCase();

  if (lower === 'hi bot') {
    setChatSession_(target, true);
    setRecordingState_(target, 'SELECT_TOPIC');
    replyLine_(event.replyToken, 'เริ่มทำงานแล้วค่ะ\n\n' + topicPrompt_(target));
    return;
  }
  if (lower === 'bye bot') {
    if (getChatSession_(target.ownerKey)) {
      endChatSession_(target);
      replyLine_(event.replyToken, 'จบการทำงานแล้วค่ะ พิมพ์รหัสคำสั่งที่ต้องการเพื่อเริ่มใหม่ได้เลย');
    }
    return;
  }
  if (!getChatSession_(target.ownerKey)) {
    if (!isDirectStartCommand_(text, target)) return;
    setChatSession_(target, true);
    setRecordingState_(target, 'SELECT_TOPIC');
  }
  touchChatSession_(target);
  if (['x', 'ออก', 'ออกจากบอท', 'จบ'].includes(lower)) {
    endChatSession_(target);
    replyLine_(event.replyToken, 'ออกจากการใช้งานบอทแล้วค่ะ พิมพ์รหัสคำสั่งที่ต้องการเพื่อเริ่มใหม่ได้เลย');
    return;
  }
  if (['ย้อนกลับ', 'กลับ', 'back', 'b'].includes(lower)) {
    replyLine_(event.replyToken, goBack_(target));
    return;
  }

  const currentState = getRecordingState_(target.ownerKey);
  if (currentState.indexOf('PERMISSION_') === 0 && !isAdmin_(target.actorId)) {
    replyLine_(event.replyToken, 'เมนูกำหนดสิทธิ์สงวนไว้สำหรับผู้ดูแลระบบค่ะ');
    return;
  }
  const requiredTopic = topicForStateOrCommand_(currentState, text, target);
  if (requiredTopic && !hasTopicAccess_(target, requiredTopic)) {
    replyLine_(event.replyToken, 'คุณไม่มีสิทธิ์ใช้หัวข้อ ' + topicLabel_(requiredTopic) + ' ในแชตนี้ค่ะ');
    return;
  }

  const inlineRecordReply = handleInlineRecordCommand_(target, text);
  if (inlineRecordReply !== null) {
    replyLine_(event.replyToken, inlineRecordReply);
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
    reply = helpText_(target);
  } else if (text === 'บันทึก') {
    setRecordingState_(target, 'SELECT_TOPIC');
    reply = topicPrompt_(target);
  } else if (text === 'เริ่ม') {
    upsertChat_(target, true);
    reply = 'เปิดรับแจ้งเตือนรายวันในแชตนี้แล้วค่ะ';
  } else if (text === 'หยุดแจ้งเตือน') {
    upsertChat_(target, false);
    reply = 'ปิดแจ้งเตือนรายวันในแชตนี้แล้วค่ะ';
  } else if (text === 'เริ่มแจ้งเตือน') {
    upsertChat_(target, true);
    reply = 'เปิดรับแจ้งเตือนรายวันในแชตนี้แล้วค่ะ';
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
    reply = 'ไม่รู้จักคำสั่งนี้ค่ะ\n' + helpText_(target);
  }
  replyLine_(event.replyToken, reply);
}

function isDirectStartCommand_(text, target) {
  const value = String(text || '').trim();
  const command = value.toUpperCase();
  if (topicForStateOrCommand_('SELECT_TOPIC', value, target)) return true;
  if (['A1', '/HELP', 'HELP', 'ช่วยเหลือ', 'บันทึก', 'เริ่ม', 'เริ่มแจ้งเตือน', 'หยุดแจ้งเตือน',
      'วันนี้', 'ยอดเงินผ่อนคงค้าง', 'ดอกเบี้ยค้าง', 'ดอกเบี้ยคงค้าง', 'สอบถามดอกเบี้ย',
      'สอบถามดอกเบี้ยค้างชำระ', 'ดูดอกเบี้ยค้างชำระ'].includes(command)) return true;
  return /^(เพิ่มงาน|สรุปงาน|ลบงาน)\b/i.test(value);
}

function getTarget_(source) {
  if (!source) return null;
  if (source.type === 'group' && source.groupId) {
    return {ownerKey: 'group:' + source.groupId, targetId: source.groupId, targetType: 'group', actorId: source.userId || ''};
  }
  if (source.type === 'room' && source.roomId) {
    return {ownerKey: 'room:' + source.roomId, targetId: source.roomId, targetType: 'room', actorId: source.userId || ''};
  }
  if (source.type === 'user' && source.userId) {
    return {ownerKey: source.userId, targetId: source.userId, targetType: 'user', actorId: source.userId};
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

function setPendingPermission_(ownerKey, draft) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === ownerKey) {
      sheet.getRange(i + 1, 10).setValue(draft ? JSON.stringify(draft) : '');
      return;
    }
  }
  sheet.appendRow([ownerKey, false, new Date(), '', '', '', false, '', '', draft ? JSON.stringify(draft) : '']);
}

function getPendingPermission_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === ownerKey && rows[i][9]) {
      try { return JSON.parse(rows[i][9]); } catch (err) { return null; }
    }
  }
  return null;
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

function getSelectedInstallmentId_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  if (!sheet) return '';
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === ownerKey) return String(rows[i][11] || '');
  }
  return '';
}

function setSelectedInstallmentId_(ownerKey, loanId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === ownerKey) {
      sheet.getRange(i + 1, 12).setValue(loanId || '');
      return;
    }
  }
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

function helpText_(target) {
  const lines = [
    '🤖 /help — คำสั่งบอท',
    'พิมพ์รหัสคำสั่งได้เลย | hi bot — เรียกเมนู | X — ออก'
  ];
  if (!target || hasTopicAccess_(target, 'APPOINTMENT')) {
    lines.push('', '📅 นัดหมาย: C1 เมนู | C2 บันทึก | C3 สรุปงาน');
  }
  if (!target || hasTopicAccess_(target, 'FINANCE')) {
    lines.push('💰 เงิน: M1 เมนู | M2 ฝาก | M3 ถอน | M4 balance');
  }
  if (!target || hasTopicAccess_(target, 'INSTALLMENT')) {
    lines.push('💳 เงินผ่อน: L1 เมนู | L2 ยอดคงค้าง');
  }
  if (target && target.targetType === 'user' && isAdmin_(target.actorId)) lines.push('🔐 A1 — กำหนดสิทธิ์');
  lines.push('', 'C2: กรอกรายละเอียด | สถานที่ แล้วเลือกวัน/เวลา', '/cancel — ยกเลิกขั้นตอน');
  return lines.join('\n');
}

function topicPrompt_(target) {
  const lines = ['📋 เมนูหลัก — พิมพ์รหัสคำสั่ง'];
  const canUseAppointment = !target || hasTopicAccess_(target, 'APPOINTMENT');
  const canUseFinance = !target || hasTopicAccess_(target, 'FINANCE');
  if (canUseAppointment) {
    lines.push('', '📅 นัดหมาย', 'C1 จัดการ | C2 บันทึก | C3 สรุปตามวัน');
  }
  if (canUseFinance) {
    lines.push('', '💰 เงิน', 'M1 จัดการ | M2 ฝาก | M3 ถอน | M4 ยอดคงเหลือ');
  }
  if (!target || hasTopicAccess_(target, 'INSTALLMENT')) {
    lines.push('', '💳 ผ่อนเงิน', 'L1 จัดการ | L2 ยอดคงค้าง');
  }
  if (target && target.targetType === 'user' && isAdmin_(target.actorId)) lines.push('', '🔐 A1 กำหนดสิทธิ์');
  lines.push('', 'X ออกจากบอท');
  return lines.join('\n');
}

function permissionActionMenu_(isMember, memberName) {
  const memberLine = isMember ? '\nสมาชิก: ' + (memberName || 'ไม่พบชื่อใน LINE') : '';
  return '🔐 กำหนดสิทธิ์' + memberLine + '\n\nตอบ Y หรือ N ตามลำดับหัวข้อ\n' +
    'ตัวอย่าง: Y,N,Y\n\n' +
    '📅 เรื่องนัดหมาย (Y/N)\n' +
    '💰 เรื่องเงิน (Y/N)\n' +
    '💳 เงินผ่อน (Y/N)\n\n' +
    (isMember ? '9 — ใช้สิทธิ์ตามกลุ่ม\n' : '9 — คืนค่าเริ่มต้น\n') +
    '0 — ยกเลิก';
}

function topicLabel_(topic) {
  if (topic === 'APPOINTMENT') return 'เรื่องนัดหมาย';
  if (topic === 'INSTALLMENT') return 'เงินผ่อน';
  return 'เรื่องเงินเก็บ';
}

function getRegisteredGroupTargets_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  if (!sheet) return [];
  ensureSheet_(SpreadsheetApp.getActive(), CONFIG.USERS_SHEET,
    ['scopeKey', 'รับแจ้งเตือน', 'สมัครเมื่อ', 'ขั้นตอนบันทึก', 'targetId', 'targetType', 'sessionActive', 'lastActivityAt', 'pendingAppointment', 'pendingPermission', 'groupName', 'selectedInstallmentId']);
  const rows = sheet.getDataRange().getValues();
  const seen = {};
  const groups = [];
  for (let i = 1; i < rows.length; i++) {
    const type = String(rows[i][5] || '').toLowerCase();
    const scopeKey = String(rows[i][0] || '');
    const targetId = String(rows[i][4] || '');
    if (type !== 'group' || !scopeKey || !targetId || seen[scopeKey]) continue;
    seen[scopeKey] = true;
    let groupName = String(rows[i][10] || '').trim();
    const fetchedName = getLineGroupName_(targetId);
    if (fetchedName) {
      groupName = fetchedName;
      sheet.getRange(i + 1, 11).setValue(groupName);
    }
    groups.push({scopeKey: scopeKey, targetId: targetId, targetType: type, groupName: groupName});
  }
  return groups;
}

function getLineGroupName_(groupId) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) return '';
  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/group/' + encodeURIComponent(groupId) + '/summary', {
    method: 'get',
    headers: {Authorization: 'Bearer ' + token},
    muteHttpExceptions: true
  });
  if (response.getResponseCode() >= 300) {
    console.warn('Could not fetch LINE group name for ' + groupId + ': ' + response.getResponseCode());
    return '';
  }
  try {
    return String(JSON.parse(response.getContentText()).groupName || '').trim();
  } catch (err) {
    console.warn('Could not parse LINE group summary for ' + groupId + ': ' + err.message);
    return '';
  }
}

function getLineGroupMemberName_(groupId, userId) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token || !groupId || !userId) return '';
  const url = 'https://api.line.me/v2/bot/group/' + encodeURIComponent(groupId) +
    '/member/' + encodeURIComponent(userId);
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {Authorization: 'Bearer ' + token},
    muteHttpExceptions: true
  });
  if (response.getResponseCode() >= 300) {
    console.warn('Could not fetch LINE group member name for ' + userId + ': ' + response.getResponseCode());
    return '';
  }
  try {
    return String(JSON.parse(response.getContentText()).displayName || '').trim();
  } catch (err) {
    console.warn('Could not parse LINE group member profile for ' + userId + ': ' + err.message);
    return '';
  }
}

function permissionScopeMenu_() {
  const groups = getRegisteredGroupTargets_();
  if (!groups.length) return 'ยังไม่พบกลุ่มที่ลงทะเบียนกับบอทค่ะ\nกรุณาเพิ่มบอทเข้ากลุ่ม แล้วพิมพ์คำสั่งใดก็ได้ในกลุ่มก่อน จากนั้นกลับมาตั้งสิทธิ์ในแชตส่วนตัว';
  return '🔐 เลือกกลุ่มที่ต้องการกำหนดสิทธิ์\n\n' + groups.map(function(group, index) {
    return (index + 1) + ' — ' + (group.groupName || 'กลุ่ม (ID ' + group.targetId + ')');
  }).join('\n') + '\n\n0 — กลับเมนูหลัก';
}

function permissionMenu_(context) {
  const groupLine = context && context.targetId ? '\nกลุ่มเป้าหมาย: ' + (context.groupName || context.targetId) + '\n' : '';
  return '🔐 กำหนดสิทธิ์การใช้งาน' + groupLine + '\n\n' +
    '1 — กำหนดสิทธิ์ให้ทั้งกลุ่มนี้\n' +
    '2 — กำหนดสิทธิ์ให้สมาชิกในกลุ่มนี้\n' +
    '0 — เลือกกลุ่มอื่น';
}

function applyPermissionChoice_(target, userId, choice, scopeKey, memberName) {
  const permissionScope = scopeKey || target.ownerKey;
  const memberScope = Boolean(userId);
  const input = String(choice || '').trim().toUpperCase();
  if (input === '9') {
    clearTopicAccess_(permissionScope, userId, 'ALL');
    return memberScope ? 'ตั้งค่าสิทธิ์ของ ' + (memberName || 'สมาชิก') + ' ให้ใช้ตามกลุ่มแล้วค่ะ' : 'คืนสิทธิ์กลุ่มเป็นค่าเริ่มต้นแล้วค่ะ';
  }
  const selections = input.split(',').map(function(item) { return item.trim(); });
  if (selections.length !== 3 || selections.some(function(item) { return item !== 'Y' && item !== 'N'; })) return '';
  const topics = [
    {topic: 'APPOINTMENT', label: 'เรื่องนัดหมาย'},
    {topic: 'FINANCE', label: 'เรื่องเงิน'},
    {topic: 'INSTALLMENT', label: 'เงินผ่อน'}
  ];
  const settings = topics.map(function(item, index) {
    return {topic: item.topic, value: selections[index] === 'Y' ? 'ALLOW' : 'DENY', label: (selections[index] === 'Y' ? 'อนุญาต ' : 'ปิดสิทธิ์ ') + item.label};
  });
  settings.forEach(function(setting) {
    setTopicAccess_(permissionScope, userId, setting.topic, setting.value, target.actorId);
  });
  return (memberScope ? 'บันทึกสิทธิ์ของ ' + (memberName || 'สมาชิก') + ' แล้วค่ะ: ' : 'บันทึกสิทธิ์กลุ่มแล้วค่ะ: ') +
    settings.map(function(setting) { return setting.label; }).join(', ');
}

function goBack_(target) {
  const state = getRecordingState_(target.ownerKey);
  if (state.indexOf('PERMISSION_') === 0) {
    setPendingPermission_(target.ownerKey, null);
    setRecordingState_(target, 'SELECT_TOPIC');
    return '↩️ กลับไปเมนูหลักแล้วค่ะ\n\n' + topicPrompt_(target);
  }
  if (state === 'SELECT_TOPIC' || !state) {
    setRecordingState_(target, 'SELECT_TOPIC');
    return '📋 อยู่ที่เมนูหลักแล้วค่ะ\n\n' + topicPrompt_(target);
  }

  if (state === 'APPOINTMENT_MENU' || state === 'SAVINGS_MENU' || state === 'INSTALLMENT_MENU') {
    setRecordingState_(target, 'SELECT_TOPIC');
    return '↩️ กลับไปเมนูหลักแล้วค่ะ\n\n' + topicPrompt_(target);
  }

  if (state === 'SAVINGS_DEPOSIT' || state === 'SAVINGS_WITHDRAW' || state === 'POST_SAVE_SAVINGS') {
    setPendingAppointment_(target.ownerKey, null);
    setRecordingState_(target, 'SAVINGS_MENU');
    return '↩️ กลับไปเมนูเงินเก็บแล้วค่ะ\n\n' + savingsMenu_();
  }

  if (state === 'INSTALLMENT_PRINCIPAL' || state === 'INSTALLMENT_PAYMENT' || state === 'INSTALLMENT_INTEREST_PAYMENT' || state === 'POST_SAVE_INSTALLMENT') {
    setRecordingState_(target, 'INSTALLMENT_MENU');
    return '↩️ กลับไปเมนูผ่อนเงินแล้วค่ะ\n\n' + installmentMenu_();
  }

  if (state === 'APPOINTMENT' || state === 'APPOINTMENT_DATE' || state === 'APPOINTMENT_FIELDS' ||
      state === 'APPOINTMENT_REMINDER' ||
      state === 'APPOINTMENT_SUMMARY_DATE' || state === 'APPOINTMENT_DELETE_PICK' ||
      state === 'POST_SAVE_APPOINTMENT') {
    setPendingAppointment_(target.ownerKey, null);
    setRecordingState_(target, 'APPOINTMENT_MENU');
    return '↩️ กลับไปเมนูนัดหมายแล้วค่ะ\n\n' + appointmentMenu_();
  }

  setRecordingState_(target, 'SELECT_TOPIC');
  return '↩️ กลับไปเมนูหลักแล้วค่ะ\n\n' + topicPrompt_(target);
}

function appointmentMenu_() {
  return '📅 นัดหมาย\n1 บันทึก | 2 สรุปตามวัน | 3 ลบ\n0 เมนูหลัก | 9 ออก';
}

function appointmentDatePrompt_() {
  return '📅 เลือกวันและเวลา หรือเลือกวันอย่างเดียวค่ะ\n0 ยกเลิก';
}

function appointmentSummaryDatePrompt_() {
  return '📅 เลือกวันที่ต้องการสรุปค่ะ\n0 ยกเลิก';
}

function appointmentFieldsPrompt_() {
  return 'กรอก: รายละเอียด | สถานที่\nตัวอย่าง: ประชุมทีม | ห้องประชุม\nไม่มีสถานที่ให้เว้นช่องหลัง | ค่ะ';
}

function appointmentDateTimeQuickReply_() {
  const now = new Date();
  return {items: [
    {type: 'action', action: {type: 'datetimepicker', label: 'เลือกวันและเวลา', data: 'selectAppointmentDateTime', mode: 'datetime', initial: Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm")}},
    {type: 'action', action: {type: 'datetimepicker', label: 'เลือกวัน ไม่ระบุเวลา', data: 'selectAppointmentDateOnly', mode: 'date', initial: Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd')}}
  ]};
}

function appointmentReminderPrompt_() {
  return '🔔 แจ้งเตือนแบบใดคะ?\n1 ล่วงหน้า 1 วัน | 2 วันนัด | 3 ทั้งสองวัน';
}

function appointmentReminderLabel_(mode) {
  if (mode === 'ADVANCE_1_DAY') return 'ล่วงหน้า 1 วัน';
  if (mode === 'BOTH') return 'ล่วงหน้า 1 วันและวันนัดหมาย';
  return 'ในวันที่นัดหมาย';
}

function datePickerQuickReply_(actionData) {
  return {
    items: [{
      type: 'action',
      action: {
        type: 'datetimepicker',
        label: 'เลือกวันที่',
        data: actionData,
        mode: 'date',
        initial: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd')
      }
    }]
  };
}

function transactionDateQuickReply_() {
  return {
    items: [
      {type: 'action', action: {type: 'datetimepicker', label: 'เลือกวันที่', data: 'selectTransactionDate', mode: 'date', initial: todayDateKey_()}},
      {type: 'action', action: {type: 'message', label: 'วันนี้', text: 'วันนี้'}}
    ]
  };
}

function transactionDatePrompt_(label) {
  return '📅 วันที่' + (label ? ' — ' + label : '') + '\nเลือกจากปฏิทิน หรือแตะ “วันนี้” ค่ะ\n0 ยกเลิก';
}

function beginTransactionDateSelection_(target, resumeState, entryText, label) {
  setPendingAppointment_(target.ownerKey, {transactionDatePick: true, resumeState: resumeState, entryText: entryText});
  setRecordingState_(target, 'TRANSACTION_DATE_PICK');
  return transactionDatePrompt_(label);
}

function completeTransactionDateSelection_(target, selectedDate) {
  const pending = getPendingAppointment_(target.ownerKey);
  if (!pending || !pending.transactionDatePick || !isValidDate_(selectedDate)) {
    setPendingAppointment_(target.ownerKey, null);
    setRecordingState_(target, 'SELECT_TOPIC');
    return 'ไม่พบข้อมูลรายการที่รอเลือกวันที่ค่ะ กรุณาเริ่มบันทึกใหม่\n\n' + topicPrompt_(target);
  }
  setPendingAppointment_(target.ownerKey, null);
  setRecordingState_(target, pending.resumeState);
  return handleRecordingFlow_(target, pending.entryText + ' | ' + selectedDate);
}

function handlePostback_(event) {
  const target = getTarget_(event.source);
  if (!target || !getChatSession_(target.ownerKey)) return;
  touchChatSession_(target);

  const actionData = String(event.postback && event.postback.data || '');
  if (!['selectAppointmentDate', 'selectAppointmentDateTime', 'selectAppointmentDateOnly', 'selectSummaryDate', 'selectTransactionDate'].includes(actionData)) return;
  const state = getRecordingState_(target.ownerKey);
  const pendingTransaction = actionData === 'selectTransactionDate' ? getPendingAppointment_(target.ownerKey) : null;
  const requiredTopic = topicForStateOrCommand_(actionData === 'selectTransactionDate' && pendingTransaction ? pendingTransaction.resumeState : state, '');
  if (requiredTopic && !hasTopicAccess_(target, requiredTopic)) {
    replyLine_(event.replyToken, 'คุณไม่มีสิทธิ์ใช้หัวข้อนี้ในแชตค่ะ');
    return;
  }
  const params = event.postback && event.postback.params || {};
  const selectedDateTime = String(params.datetime || '');
  const selectedDate = selectedDateTime ? selectedDateTime.slice(0, 10) : String(params.date || '');
  const selectedTime = selectedDateTime ? selectedDateTime.slice(11, 16) : '';
  if (!isValidDate_(selectedDate)) {
    replyLine_(event.replyToken, 'ไม่ได้รับวันที่จากปฏิทินค่ะ กรุณาลองเลือกวันที่อีกครั้ง');
    return;
  }

  if (actionData === 'selectTransactionDate' && state === 'TRANSACTION_DATE_PICK') {
    replyLine_(event.replyToken, completeTransactionDateSelection_(target, selectedDate));
    return;
  }

  if (['selectAppointmentDate', 'selectAppointmentDateTime', 'selectAppointmentDateOnly'].includes(actionData) && state === 'APPOINTMENT_DATE') {
    const appointment = getPendingAppointment_(target.ownerKey) || {};
    appointment.date = selectedDate;
    if (actionData === 'selectAppointmentDateTime') appointment.time = selectedTime;
    if (actionData === 'selectAppointmentDateOnly') appointment.time = '';
    setPendingAppointment_(target.ownerKey, appointment);
    setRecordingState_(target, 'APPOINTMENT_REMINDER');
    replyLine_(event.replyToken, '📅 วันที่นัดหมาย: ' + formatDateWeekday_(selectedDate) +
      (appointment.time ? '\n🕒 เวลา: ' + appointment.time : '\n🕒 เวลา: ไม่ระบุ') +
      '\n\n' + appointmentReminderPrompt_());
    return;
  }

  if (actionData === 'selectSummaryDate' && state === 'APPOINTMENT_SUMMARY_DATE') {
    setRecordingState_(target, 'APPOINTMENT_MENU');
    const result = summarizeDate_(target.ownerKey, 'สรุปงาน ' + selectedDate);
    replyLine_(event.replyToken, result + '\n\n' + appointmentMenu_());
  }
}

function savingsMenu_() {
  return '💰 เงินเก็บ\n1 ฝาก | 2 ถอน | 3 ยอดคงเหลือ | 4 ย้อนหลัง 15 วัน\n0 เมนูหลัก | 9 ออก';
}

function savingsEntryPrompt_(type) {
  const deposit = type === 'ฝาก';
  return (deposit ? '📥 ฝากเงิน' : '📤 ถอนเงิน') + '\nจำนวนเงิน | หมายเหตุ\nตัวอย่าง: ' + (deposit ? '500 | เก็บเงินเดือน' : '200 | ใช้จ่าย') + '\nเลือกวันที่จากปฏิทิน หรือแตะ “วันนี้” | 0 ยกเลิก';
}

function installmentMenu_() {
  return '💳 เงินผ่อน\n1 ตั้งยอด | 2 จ่ายต้น | 3 จ่ายดอกเบี้ย\n4 ยอดคงเหลือ | 5 ประวัติ 16 เดือน | 6 ดอกเบี้ยค้าง\n7 คำนวณดอกเบี้ยจากยอดล่าสุด\n0 เมนูหลัก | 9 ออก';
}

function installmentPrincipalFormatError_() {
  return 'รูปแบบ: หัวเรื่อง | เจ้าหนี้ | ลูกหนี้ | เงินต้น | ดอกเบี้ย% | หมายเหตุ | ฐาน\nฐานเลือก “ยอดตั้งต้น” หรือ “ยอดคงเหลือ” (ไม่ใส่ได้)\nตัวอย่าง: รถ | Mama | Kwan | 300000 | 1% | | ยอดตั้งต้น\n0 ยกเลิก';
}

function installmentPaymentPrompt_(loan) {
  const current = loan || {balance: 0, originalPrincipal: 0, rate: 0, title: '', creditor: '', debtor: ''};
  return '💵 จ่ายเงินต้น' + (current.title ? ' — ' + current.title : '') + '\nคงเหลือ ' + formatMoney_(current.balance) + ' บาท\nกรอก: จำนวน | หมายเหตุ (เช่น 3000 | งวดนี้)\nจากนั้นเลือกวันที่ | 0 ยกเลิก';
}

function installmentInterestPaymentPrompt_(loan) {
  return '📈 จ่ายดอกเบี้ย' + (loan.title ? ' — ' + loan.title : '') + '\nค้าง ' + formatMoney_(loan.interestDue) + ' บาท\nกรอก: จำนวน | หมายเหตุ (เช่น 750 | จ่ายดอกเบี้ย)\nจากนั้นเลือกวันที่ | 0 ยกเลิก';
}

function installmentAfterSaveMenu_() {
  return '💳 ทำต่อ: 1 จ่ายต้น | 2 จ่ายดอกเบี้ย | 3 ยอดคงเหลือ | 4 ดอกเบี้ยค้าง\n0 เมนูหลัก | 9 ออก';
}

// Allows a record command and its complete data to be sent in one LINE message.
function handleInlineRecordCommand_(target, text) {
  const match = String(text || '').trim().match(/^(C2|M2|M3|L1)\s+(.+)$/i);
  if (!match) return null;
  const code = match[1].toUpperCase();
  const detail = match[2].trim();
  const state = getRecordingState_(target.ownerKey);
  const allowedInState = state === 'SELECT_TOPIC' ||
    (state === 'APPOINTMENT_MENU' && code === 'C2') ||
    (state === 'SAVINGS_MENU' && ['M2', 'M3'].includes(code)) ||
    (state === 'INSTALLMENT_MENU' && code === 'L1');
  if (!allowedInState) return null;
  const parts = detail.split('|').map(function(part) { return part.trim(); });

  if (code === 'C2') {
    if (parts.length === 2 && parts[0]) {
      setPendingAppointment_(target.ownerKey, {description: parts[0], location: parts[1], time: ''});
      setRecordingState_(target, 'APPOINTMENT_DATE');
      return appointmentDatePrompt_();
    }
    if (parts.length === 3 && (parts[0] === '' || /^([01]\d|2[0-3]):[0-5]\d$/.test(parts[0])) && parts[1]) {
      setPendingAppointment_(target.ownerKey, {time: parts[0], description: parts[1], location: parts[2]});
      setRecordingState_(target, 'APPOINTMENT_DATE');
      return appointmentDatePrompt_();
    }
    if (parts.length < 4 || parts.length > 5 || !isValidDate_(parts[0]) ||
        (parts[1] && !/^([01]\d|2[0-3]):[0-5]\d$/.test(parts[1])) || !parts[2] ||
        (parts.length === 5 && !['1', '2', '3'].includes(parts[4]))) {
      return 'รูปแบบบันทึกนัดหมายไม่ถูกต้องค่ะ\nกรอกรายละเอียดก่อนเลือกวันที่: C2 เวลา | รายละเอียด | สถานที่\nตัวอย่าง: C2 10:00 | ประชุม | ออฟฟิศ\nถ้าไม่มีเวลา/สถานที่ ให้เว้นช่องไว้';
    }
    const appointment = {date: parts[0], time: parts[1], description: parts[2], location: parts[3]};
    setPendingAppointment_(target.ownerKey, appointment);
    setRecordingState_(target, 'APPOINTMENT_REMINDER');
    return appointmentReminderPrompt_();
  }

  if (code === 'M2' || code === 'M3') {
    if ((parts.length !== 2 && parts.length !== 3) || !isPositiveMoney_(parts[0]) || !parts[1] ||
        (parts.length === 3 && parts[2] && !isValidDate_(parts[2]))) {
      return 'รูปแบบบันทึกเงินไม่ถูกต้องค่ะ\nใช้: ' + code + ' จำนวนเงิน | หมายเหตุ\nตัวอย่าง: ' + (code === 'M2' ? 'M2 500 | เก็บเงินเดือน' : 'M3 200 | ใช้จ่าย');
    }
    if (!parts[2]) return beginTransactionDateSelection_(target, code === 'M2' ? 'SAVINGS_DEPOSIT' : 'SAVINGS_WITHDRAW', parts[0] + ' | ' + parts[1], code === 'M2' ? 'ฝากเงิน' : 'ถอนเงิน');
    const transactionDate = optionalTransactionDate_(parts[2]);
    const type = code === 'M2' ? 'ฝาก' : 'ถอน';
    const result = addSavings_(target, [transactionDate, type, parts[0], parts[1]].join(' | '));
    if (result.startsWith('บันทึกเงินเก็บแล้ว')) {
      setRecordingState_(target, 'POST_SAVE_SAVINGS');
      return result + '\n\n' + getSavingsBalance_(target.ownerKey) + '\n\n' + afterSavePrompt_('เงินเก็บ');
    }
    return result;
  }

  const action = parts.shift();
  if (!['1', '2', '3'].includes(action)) return 'รูปแบบเงินผ่อนไม่ถูกต้องค่ะ ใช้ L1 1 (ตั้งยอด), L1 2 (ชำระเงินต้น) หรือ L1 3 (ชำระดอกเบี้ย) ตามด้วยรายละเอียดค่ะ';
  if (action === '1') {
    setRecordingState_(target, 'INSTALLMENT_PRINCIPAL');
    return handleRecordingFlow_(target, parts.join(' | '));
  }
  if ((parts.length !== 2 && parts.length !== 3) || !isPositiveMoney_(parts[0]) || !parts[1] || (parts.length === 3 && parts[2] && !isValidDate_(parts[2]))) return 'รูปแบบชำระเงินผ่อนไม่ถูกต้องค่ะ ใช้: L1 ' + action + ' | จำนวนเงิน | หมายเหตุ';
  return selectInstallmentForAction_(target, action === '2' ? 'INSTALLMENT_PAYMENT' : 'INSTALLMENT_INTEREST_PAYMENT', parts.join(' | '));
}

function isPositiveMoney_(value) {
  return /^\d+(\.\d{1,2})?$/.test(String(value)) && Number(value) > 0;
}

function normalizeMoneyField_(value) {
  const text = String(value || '').trim();
  if (!/^(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d{1,2})?$/.test(text)) return '';
  return text.replace(/,/g, '');
}

function normalizeRateField_(value) {
  return String(value || '').trim().replace(/%$/, '').trim();
}

function parseInterestBasis_(value) {
  const basis = String(value || '').trim().toUpperCase().replace(/\s+/g, '');
  if (['ยอดคงเหลือ', 'คงเหลือ', 'BALANCE', 'OUTSTANDING', '2'].includes(basis)) return 'BALANCE';
  if (['ยอดตั้งต้น', 'เงินต้นตั้งต้น', 'ตั้งต้น', 'ORIGINAL', '1'].includes(basis)) return 'ORIGINAL';
  return '';
}

function isNonNegativeRate_(value) {
  return /^\d+(\.\d{1,4})?$/.test(String(value)) && Number(value) >= 0;
}

function todayDateKey_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function optionalTransactionDate_(value) {
  const date = String(value || '').trim();
  if (!date) return todayDateKey_();
  return isValidDate_(date) ? date : '';
}

function getInstallmentLoanGroups_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.INSTALLMENTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const groups = [];
  const byId = {};
  let legacyCurrentId = '';
  sheet.getDataRange().getValues().slice(1).forEach(function(row) {
    if (String(row[1]) !== ownerKey) return;
    const type = String(row[3] || '');
    let loanId = String(row[14] || '');
    if (type === 'ตั้งต้น') {
      loanId = loanId || String(row[0]);
      legacyCurrentId = loanId;
      if (!byId[loanId]) {
        byId[loanId] = {loanId: loanId, entries: []};
        groups.push(byId[loanId]);
      }
      byId[loanId].entries.push(row);
      return;
    }
    loanId = loanId || legacyCurrentId;
    if (!loanId) return;
    if (!byId[loanId]) {
      byId[loanId] = {loanId: loanId, entries: []};
      groups.push(byId[loanId]);
    }
    byId[loanId].entries.push(row);
  });
  return groups;
}

function emptyInstallmentSummary_() {
  return {loanId: '', balance: 0, originalPrincipal: 0, rate: 0, interestAccrued: 0, interestPaid: 0, interestDue: 0, title: '', creditor: '', debtor: '', interestBasis: 'ORIGINAL', totalInterest: 0, latestPaymentDate: '', entries: []};
}

function summarizeInstallmentEntries_(loanId, entries) {
  let balance = 0;
  let originalPrincipal = 0;
  let rate = 0;
  let interestAccrued = 0;
  let interestPaid = 0;
  let latestPaymentDate = '';
  let title = '';
  let creditor = '';
  let debtor = '';
  let interestBasis = 'ORIGINAL';
  entries.forEach(function(row) {
    const hasRecordedBalance = row[7] !== '' && row[7] !== null && row[7] !== undefined && !isNaN(Number(row[7]));
    const recordedBalance = hasRecordedBalance ? Number(row[7]) : null;
    if (row[3] === 'ตั้งต้น') {
      balance = Number(row[4]) || 0;
      originalPrincipal = Number(row[4]) || 0;
      rate = Number(row[5]) || 0;
      interestAccrued = 0;
      interestPaid = 0;
      latestPaymentDate = '';
      title = String(row[10] || title);
      creditor = String(row[11] || creditor);
      debtor = String(row[12] || debtor);
      interestBasis = parseInterestBasis_(row[13]) || 'ORIGINAL';
    } else if (row[3] === 'ผ่อน') { // Legacy combined installment record.
      balance = hasRecordedBalance ? recordedBalance : Math.max(0, balance - (Number(row[4]) || 0));
      title = String(row[10] || title);
      creditor = String(row[11] || creditor);
      debtor = String(row[12] || debtor);
      interestAccrued += Number(row[6]) || 0;
    } else if (row[3] === 'ดอกเบี้ยรายปี' || row[3] === 'คำนวณดอกเบี้ยจากยอดล่าสุด') {
      if (hasRecordedBalance) balance = recordedBalance;
      interestAccrued += Number(row[6]) || 0;
    } else if (row[3] === 'จ่ายดอกเบี้ย') {
      if (hasRecordedBalance) balance = recordedBalance;
      interestPaid += Number(row[6]) || 0;
    } else if (row[3] === 'จ่ายเงินต้น') {
      balance = hasRecordedBalance ? recordedBalance : Math.max(0, balance - (Number(row[4]) || 0));
    }
    if (['ผ่อน', 'จ่ายเงินต้น', 'จ่ายดอกเบี้ย'].includes(String(row[3])) ) {
      const paymentDate = cellDateKey_(row[2]);
      if (paymentDate && paymentDate > latestPaymentDate) latestPaymentDate = paymentDate;
    }
  });
  return {loanId: loanId, balance: balance, originalPrincipal: originalPrincipal, rate: rate, interestAccrued: interestAccrued, interestPaid: interestPaid, interestDue: Math.max(0, interestAccrued - interestPaid), title: title, creditor: creditor, debtor: debtor, interestBasis: interestBasis, totalInterest: interestAccrued, latestPaymentDate: latestPaymentDate, entries: entries};
}

function getInstallmentSummaries_(ownerKey, accrue) {
  if (accrue !== false) accrueAnnualInterestForOwner_(ownerKey);
  return getInstallmentLoanGroups_(ownerKey).map(function(group) {
    return summarizeInstallmentEntries_(group.loanId, group.entries);
  });
}

function getInstallmentSummary_(ownerKey, loanId) {
  const summaries = getInstallmentSummaries_(ownerKey);
  const selectedId = loanId || getSelectedInstallmentId_(ownerKey);
  if (selectedId) {
    const selected = summaries.find(function(loan) { return loan.loanId === selectedId; });
    if (selected) return selected;
  }
  if (summaries.length) return summaries[summaries.length - 1];
  return emptyInstallmentSummary_();
}

function activeInstallmentLoans_(ownerKey) {
  return getInstallmentSummaries_(ownerKey).filter(function(loan) {
    return loan.balance > 0 || loan.interestDue > 0;
  });
}

function installmentHistoryWindow_() {
  const currentMonth = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM');
  const cutoffDate = new Date(currentMonth + '-01T00:00:00+07:00');
  cutoffDate.setMonth(cutoffDate.getMonth() - 15);
  return {currentMonth: currentMonth, cutoffMonth: Utilities.formatDate(cutoffDate, CONFIG.TIMEZONE, 'yyyy-MM')};
}

function hasInstallmentHistoryInWindow_(summary, window) {
  return summary.entries.some(function(row) {
    const month = cellDateKey_(row[2]).slice(0, 7);
    return month >= window.cutoffMonth && month <= window.currentMonth;
  });
}

function selectInstallmentForAction_(target, actionState, entryText) {
  const isHistory = actionState === 'INSTALLMENT_HISTORY';
  const historyWindow = isHistory ? installmentHistoryWindow_() : null;
  const loans = isHistory ? getInstallmentSummaries_(target.ownerKey).filter(function(loan) {
    return hasInstallmentHistoryInWindow_(loan, historyWindow);
  }) : activeInstallmentLoans_(target.ownerKey);
  const needsPrincipal = actionState === 'INSTALLMENT_PAYMENT';
  const isBalanceCalculation = actionState === 'INSTALLMENT_INTEREST_CALC';
  const available = isHistory ? loans : loans.filter(function(loan) { return isBalanceCalculation || needsPrincipal ? loan.balance > 0 : loan.interestDue > 0; });
  if (!available.length) return (isHistory ? 'ไม่มีสัญญาที่มีรายการในช่วงย้อนหลัง 16 เดือน' : (isBalanceCalculation || needsPrincipal ? 'ยังไม่มียอดเงินต้นคงเหลือ' : 'ยังไม่มีดอกเบี้ยค้างชำระ')) + 'ค่ะ\n\n' + installmentMenu_();
  if (available.length === 1) {
    setSelectedInstallmentId_(target.ownerKey, available[0].loanId);
    if (isHistory) return getInstallmentHistory_(target.ownerKey, available[0].loanId) + '\n\n' + installmentMenu_();
    setRecordingState_(target, actionState);
    if (isBalanceCalculation) return calculateInterestFromLatestBalance_(target, available[0]);
    return entryText ? handleRecordingFlow_(target, entryText) :
      (needsPrincipal ? installmentPaymentPrompt_(available[0]) : installmentInterestPaymentPrompt_(available[0]));
  }
  setPendingAppointment_(target.ownerKey, {installmentLoanPick: true, actionState: actionState, entryText: entryText || ''});
  setRecordingState_(target, 'INSTALLMENT_LOAN_PICK');
  return (isHistory ? '📜 เลือกสัญญาที่ต้องการดูประวัติ โดยพิมพ์หมายเลขค่ะ\n\n' : '💳 เลือกสัญญาผ่อนที่ต้องการทำรายการ โดยพิมพ์หมายเลขค่ะ\n\n') + available.map(function(loan, index) {
    return (index + 1) + ' — ' + loan.title + ' | เจ้าหนี้: ' + loan.creditor + ' | ลูกหนี้: ' + loan.debtor + '\n💵 เงินต้นคงเหลือ ' + formatMoney_(loan.balance) + ' บาท | 📈 ดอกเบี้ยค้าง ' + formatMoney_(loan.interestDue) + ' บาท';
  }).join('\n\n') + '\n\n0 — ยกเลิก';
}

function accrueAnnualInterest() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.INSTALLMENTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  const ownerKeys = Array.from(new Set(sheet.getDataRange().getValues().slice(1).map(function(row) { return String(row[1] || ''); }).filter(Boolean)));
  ownerKeys.forEach(accrueAnnualInterestForOwner_);
}

function accrueAnnualInterestForOwner_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.INSTALLMENTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  getInstallmentLoanGroups_(ownerKey).forEach(function(group) {
    accrueAnnualInterestForLoan_(ownerKey, group);
  });
}

function accrueAnnualInterestForLoan_(ownerKey, group) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.INSTALLMENTS_SHEET);
  const rows = group.entries;
  const startIndex = rows.findIndex(function(row) { return row[3] === 'ตั้งต้น'; });
  if (startIndex < 0) return;
  const start = rows[startIndex];
  const originalPrincipal = Number(start[4]) || 0;
  const annualRate = Number(start[5]) || 0;
  const startDate = cellDateKey_(start[2]);
  if (originalPrincipal <= 0 || !annualRate || !isValidDate_(startDate)) return;

  let lastAnnualAccrualDate = '';
  let lastLegacyInterestDate = '';
  for (let i = startIndex + 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[3] === 'ดอกเบี้ยรายปี') lastAnnualAccrualDate = cellDateKey_(row[2]);
    if (row[3] === 'ผ่อน' && Number(row[6]) > 0) lastLegacyInterestDate = cellDateKey_(row[2]);
  }

  let accrualAnchor = startDate;
  if (lastAnnualAccrualDate > accrualAnchor) accrualAnchor = lastAnnualAccrualDate;
  if (lastLegacyInterestDate > accrualAnchor) accrualAnchor = lastLegacyInterestDate;
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const interestBasis = parseInterestBasis_(start[13]) || 'ORIGINAL';
  let nextAnniversary = addYearsToDateKey_(accrualAnchor, 1);
  while (nextAnniversary && nextAnniversary <= today) {
    const balanceAtAnniversary = installmentPrincipalBalanceAt_(rows, startIndex, nextAnniversary, originalPrincipal);
    if (balanceAtAnniversary <= 0) break;
    const interestBase = interestBasis === 'BALANCE' ? balanceAtAnniversary : originalPrincipal;
    const interest = Math.round((interestBase * annualRate / 100 + Number.EPSILON) * 100) / 100;
    const id = Utilities.getUuid().slice(0, 8);
    sheet.appendRow([id, ownerKey, nextAnniversary, 'ดอกเบี้ยรายปี', 0, annualRate, interest, balanceAtAnniversary,
      'ดอกเบี้ยครบปี ' + accrualAnchor + ' ถึง ' + nextAnniversary, new Date(), start[10] || '', start[11] || '', start[12] || '', interestBasis, group.loanId]);
    accrualAnchor = nextAnniversary;
    nextAnniversary = addYearsToDateKey_(accrualAnchor, 1);
  }
}

function installmentPrincipalBalanceAt_(rows, startIndex, dateKey, originalPrincipal) {
  const payments = rows.slice(startIndex + 1).filter(function(row) {
    const type = String(row[3] || '');
    const date = cellDateKey_(row[2]);
    return ['ผ่อน', 'จ่ายเงินต้น'].includes(type) && date && date <= dateKey;
  }).sort(function(a, b) {
    return cellDateKey_(a[2]).localeCompare(cellDateKey_(b[2])) ||
      new Date(a[9]).getTime() - new Date(b[9]).getTime();
  });
  let balance = originalPrincipal;
  payments.forEach(function(row) {
    if (row[3] === 'จ่ายเงินต้น') {
      balance = Math.max(0, balance - (Number(row[4]) || 0));
    } else if (row[7] !== '' && row[7] !== null && row[7] !== undefined && !isNaN(Number(row[7]))) {
      balance = Number(row[7]);
    } else {
      balance = Math.max(0, balance - (Number(row[4]) || 0));
    }
  });
  return Math.max(0, balance);
}

function addYearsToDateKey_(dateKey, years) {
  if (!isValidDate_(dateKey)) return '';
  const date = new Date(dateKey + 'T00:00:00+07:00');
  const originalMonth = date.getMonth();
  date.setFullYear(date.getFullYear() + years);
  if (date.getMonth() !== originalMonth) date.setDate(0);
  return Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function addInstallmentPrincipal_(target, title, creditor, debtor, principal, rate, note, transactionDate, interestBasis) {
  const date = optionalTransactionDate_(transactionDate);
  if (!date) return 'รูปแบบวันที่ไม่ถูกต้องค่ะ ใช้ YYYY-MM-DD หรือเว้นว่างเพื่อใช้วันปัจจุบัน';
  const basis = interestBasis === 'BALANCE' ? 'BALANCE' : 'ORIGINAL';
  const id = Utilities.getUuid().slice(0, 8);
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.INSTALLMENTS_SHEET);
  sheet.appendRow([id, target.ownerKey, date, 'ตั้งต้น', principal, rate, 0, principal, note, new Date(), title, creditor, debtor, basis, id]);
  setSelectedInstallmentId_(target.ownerKey, id);
  const basisLabel = basis === 'BALANCE' ? 'เงินต้นคงเหลือ ณ วันครบรอบปี' : 'เงินต้นตั้งต้น';
  return '✅ บันทึกยอดตั้งต้นแล้วค่ะ (รหัส ' + id + ')\n📅 วันที่รายการ: ' + date + '\n🧾 หัวเรื่อง: ' + title + '\n🙋 เจ้าหนี้: ' + creditor + '\n👤 ลูกหนี้: ' + debtor + '\n💵 เงินต้นตั้งต้น: ' + formatMoney_(principal) + ' บาท\n📈 อัตราดอกเบี้ย: ' + formatMoney_(rate) + '% ต่อปี คิดจาก ' + basisLabel + ' เมื่อครบปี\nดอกเบี้ยจะบันทึกแยกและไม่หักจากเงินต้นค่ะ';
}

function addInstallmentPayment_(target, loan, principal, note, balanceAfter, transactionDate) {
  const date = optionalTransactionDate_(transactionDate);
  if (!date) return 'รูปแบบวันที่ไม่ถูกต้องค่ะ ใช้ YYYY-MM-DD หรือเว้นว่างเพื่อใช้วันปัจจุบัน';
  const id = Utilities.getUuid().slice(0, 8);
  SpreadsheetApp.getActive().getSheetByName(CONFIG.INSTALLMENTS_SHEET)
    .appendRow([id, target.ownerKey, date, 'จ่ายเงินต้น', principal, loan.rate, 0, balanceAfter, note, new Date(), loan.title, loan.creditor, loan.debtor, loan.interestBasis, loan.loanId]);
  return '✅ บันทึกชำระเงินต้นแล้วค่ะ (รหัส ' + id + ')\n📅 วันที่รายการ: ' + date + '\n🧾 หัวเรื่อง: ' + loan.title + '\n🙋 เจ้าหนี้: ' + loan.creditor + '\n👤 ลูกหนี้: ' + loan.debtor + '\n📥 ชำระเงินต้น: ' + formatMoney_(principal) + ' บาท\n💵 เงินต้นคงเหลือ: ' + formatMoney_(balanceAfter) + ' บาท\n📈 ดอกเบี้ยค้างชำระแยก: ' + formatMoney_(loan.interestDue) + ' บาท';
}

function addInstallmentInterestPayment_(target, loan, amount, note, transactionDate) {
  const date = optionalTransactionDate_(transactionDate);
  if (!date) return 'รูปแบบวันที่ไม่ถูกต้องค่ะ ใช้ YYYY-MM-DD หรือเว้นว่างเพื่อใช้วันปัจจุบัน';
  const id = Utilities.getUuid().slice(0, 8);
  SpreadsheetApp.getActive().getSheetByName(CONFIG.INSTALLMENTS_SHEET)
    .appendRow([id, target.ownerKey, date, 'จ่ายดอกเบี้ย', 0, loan.rate, amount, loan.balance, note, new Date(), loan.title, loan.creditor, loan.debtor, loan.interestBasis, loan.loanId]);
  return '✅ บันทึกชำระดอกเบี้ยแล้วค่ะ (รหัส ' + id + ')\n📅 วันที่รายการ: ' + date + '\n🧾 หัวเรื่อง: ' + loan.title + '\n📈 ชำระดอกเบี้ย: ' + formatMoney_(amount) + ' บาท\n📈 ดอกเบี้ยค้างชำระ: ' + formatMoney_(Math.max(0, loan.interestDue - amount)) + ' บาท\n💵 เงินต้นคงเหลือ: ' + formatMoney_(loan.balance) + ' บาท';
}

function calculateInterestFromLatestBalance_(target, loan) {
  setRecordingState_(target, 'INSTALLMENT_MENU');
  const balance = Number(loan.balance) || 0;
  const rate = Number(loan.rate) || 0;
  if (balance <= 0) return 'ไม่มียอดเงินต้นคงเหลือให้คำนวณค่ะ\n\n' + installmentMenu_();
  if (rate <= 0) return 'สัญญานี้ไม่ได้กำหนดอัตราดอกเบี้ยค่ะ\n\n' + installmentMenu_();
  const date = todayDateKey_();
  const marker = 'คำนวณจากยอดล่าสุด ' + balance + ' บาท ณ ' + date;
  const alreadyRecorded = loan.entries.some(function(row) {
    return String(row[3]) === 'คำนวณดอกเบี้ยจากยอดล่าสุด' &&
      String(row[8] || '') === marker;
  });
  if (alreadyRecorded) return 'วันนี้มีรายการคำนวณดอกเบี้ยจากยอด ' + formatMoney_(balance) + ' บาทแล้วค่ะ ไม่ได้บันทึกรายการซ้ำ\n\n💳 ' + loan.title + '\n' + installmentMenu_();
  const interest = Math.round((balance * rate / 100 + Number.EPSILON) * 100) / 100;
  const id = Utilities.getUuid().slice(0, 8);
  SpreadsheetApp.getActive().getSheetByName(CONFIG.INSTALLMENTS_SHEET).appendRow([
    id, target.ownerKey, date, 'คำนวณดอกเบี้ยจากยอดล่าสุด', 0, rate, interest, balance,
    marker, new Date(), loan.title, loan.creditor, loan.debtor, loan.interestBasis, loan.loanId
  ]);
  return '✅ คำนวณดอกเบี้ยจากยอดล่าสุดแล้วค่ะ (รหัส ' + id + ')\n📅 วันที่คำนวณ: ' + date + '\n🧾 หัวเรื่อง: ' + loan.title + '\n💵 ฐานคำนวณ: เงินต้นคงเหลือ ' + formatMoney_(balance) + ' บาท\n📈 อัตราดอกเบี้ย: ' + formatMoney_(rate) + '% ต่อปี\n💸 ดอกเบี้ยที่บันทึก: ' + formatMoney_(interest) + ' บาท\nดอกเบี้ยถูกบันทึกแยกและไม่หักจากเงินต้นค่ะ\n\n' + installmentMenu_();
}

function formatInstallmentBalance_(summary) {
  const basisText = summary.interestBasis === 'BALANCE' ? 'ยอดเงินต้นคงเหลือ ณ วันครบรอบปี' : 'เงินต้นตั้งต้น';
  const hasOutstanding = summary.balance > 0 || summary.interestDue > 0;
  const latestPaymentLine = hasOutstanding ? '\n📅 จ่ายล่าสุด: ' + (summary.latestPaymentDate || 'ยังไม่มีรายการจ่าย') : '';
  return '🧾 ' + (summary.title || 'สัญญาผ่อน') + '\n🙋 เจ้าหนี้: ' + (summary.creditor || '-') + ' | 👤 ลูกหนี้: ' + (summary.debtor || '-') + '\n' +
    '💵 เงินต้นคงเหลือ: ' + formatMoney_(summary.balance) + ' บาท\n📈 ดอกเบี้ยค้างชำระ: ' + formatMoney_(summary.interestDue) + ' บาท' +
    latestPaymentLine + '\nฐานดอกเบี้ย: ' + basisText + ' (' + formatMoney_(summary.rate) + '% ต่อปี)';
}

function getInstallmentBalance_(ownerKey, loanId) {
  const summaries = getInstallmentSummaries_(ownerKey);
  const summary = loanId ? summaries.find(function(item) { return item.loanId === loanId; }) : null;
  if (loanId) {
    if (!summary) return 'ไม่พบสัญญาผ่อนที่เลือกค่ะ';
    if (summary.balance <= 0) return 'สัญญา ' + (summary.title || 'ที่เลือก') + ' ไม่มีเงินต้นคงเหลือแล้วค่ะ';
    return '💳 ยอดเงินผ่อนคงค้าง\n\n' + formatInstallmentBalance_(summary);
  }
  const outstanding = summaries.filter(function(item) { return item.balance > 0 || item.interestDue > 0; });
  if (!outstanding.length) return 'ไม่มีสัญญาที่มียอดเงินต้นหรือดอกเบี้ยค้างชำระค่ะ';
  return '💳 ยอดเงินผ่อนคงค้าง\n\n' + outstanding.map(formatInstallmentBalance_).join('\n\n──────────\n\n');
}

function getInstallmentInterestDue_(ownerKey, loanId) {
  const summaries = getInstallmentSummaries_(ownerKey);
  const loans = loanId ? summaries.filter(function(item) { return item.loanId === loanId; }) : summaries;
  if (!loans.length) return loanId ? 'ไม่พบสัญญาผ่อนที่เลือกค่ะ' : 'ยังไม่มีรายการผ่อนเงินในแชตนี้ค่ะ';
  return '📈 ดอกเบี้ยค้างชำระ\n\n' + loans.map(function(loan) {
    const basisText = loan.interestBasis === 'BALANCE' ? 'ยอดเงินต้นคงเหลือ ณ วันครบรอบปี' : 'เงินต้นตั้งต้น';
    return '🧾 ' + loan.title + '\nเจ้าหนี้: ' + loan.creditor + ' | ลูกหนี้: ' + loan.debtor + '\nดอกเบี้ยค้างชำระ: ' + formatMoney_(loan.interestDue) + ' บาท\nอัตราดอกเบี้ย: ' + formatMoney_(loan.rate) + '% ต่อปี คิดจาก ' + basisText;
  }).join('\n\n──────────\n\n') + '\n\nดอกเบี้ยจะเพิ่มเมื่อครบปี และยอดนี้แยกจากเงินต้นค่ะ';
}

function getInstallmentHistory_(ownerKey, loanId) {
  let summaries = getInstallmentSummaries_(ownerKey, false);
  if (loanId) summaries = summaries.filter(function(summary) { return summary.loanId === loanId; });
  const historyWindow = installmentHistoryWindow_();
  const currentMonth = historyWindow.currentMonth;
  const cutoffMonth = historyWindow.cutoffMonth;
  const recent = [].concat.apply([], summaries.map(function(summary) { return summary.entries; })).filter(function(row) {
    const month = cellDateKey_(row[2]).slice(0, 7);
    return month >= cutoffMonth && month <= currentMonth;
  }).sort(function(a, b) {
    return cellDateKey_(b[2]).localeCompare(cellDateKey_(a[2])) || new Date(b[9]).getTime() - new Date(a[9]).getTime();
  });
  if (!recent.length) return 'ไม่มีรายการผ่อนเงินย้อนหลัง 16 เดือนค่ะ';
  const labels = {'ตั้งต้น': '🧾 ตั้งยอดเงินต้น', 'ดอกเบี้ยรายปี': '📈 คิดดอกเบี้ยครบปี', 'คำนวณดอกเบี้ยจากยอดล่าสุด': '📈 คำนวณดอกเบี้ย', 'จ่ายเงินต้น': '💵 ชำระเงินต้น', 'จ่ายดอกเบี้ย': '💸 ชำระดอกเบี้ย', 'ผ่อน': '💳 รายการเดิม'};
  const grouped = {};
  recent.forEach(function(row) {
    let summary = summaries.find(function(item) { return item.loanId === String(row[14] || ''); });
    if (!summary) summary = summaries.find(function(item) { return item.entries.indexOf(row) !== -1; });
    const key = summary ? summary.loanId : String(row[14] || row[0]);
    if (!grouped[key]) grouped[key] = {summary: summary, rows: []};
    grouped[key].rows.push(row);
  });
  const sections = Object.keys(grouped).map(function(key) {
    const group = grouped[key];
    const first = group.rows[0];
    const title = group.summary ? group.summary.title : String(first[10] || 'สัญญาผ่อน');
    const creditor = group.summary ? group.summary.creditor : String(first[11] || '-');
    const debtor = group.summary ? group.summary.debtor : String(first[12] || '-');
    const heading = '🧾 ' + title + '\n🙋 เจ้าหนี้: ' + creditor + ' | 👤 ลูกหนี้: ' + debtor;
    const entries = group.rows.map(function(row) {
      const date = cellDateKey_(row[2]);
      const label = labels[row[3]] || '💳 รายการผ่อน';
      const note = String(row[8] || '').trim();
      return '┌ 📅 ' + date + ' | ' + label + '\n│ 💵 เงินต้น: ' + formatMoney_(Number(row[4]) || 0) + ' บาท\n│ 📈 ดอกเบี้ย: ' + formatMoney_(Number(row[6]) || 0) + ' บาท\n│ 💰 เงินต้นคงเหลือ: ' + formatMoney_(Number(row[7]) || 0) + ' บาท' + (note ? '\n└ 📝 ' + note : '\n└──────────');
    });
    return heading + '\n\n' + entries.join('\n\n');
  });
  let currentBalance;
  if (loanId) {
    const summary = summaries[0];
    currentBalance = summary.balance <= 0 && summary.interestDue <= 0 ? 'สัญญานี้ปิดยอดแล้วค่ะ ไม่มีเงินต้นหรือดอกเบี้ยค้าง' : formatInstallmentBalance_(summary);
  } else {
    currentBalance = getInstallmentBalance_(ownerKey);
  }
  return '📜 ประวัติผ่อนเงินย้อนหลัง 16 เดือน (' + cutoffMonth + ' ถึง ' + currentMonth + ') — ' + recent.length + ' รายการ\n\n' + sections.join('\n\n──────────\n\n') + '\n\n📌 ยอดคงค้างปัจจุบัน\n' + currentBalance;
}

function afterSavePrompt_(topic) {
  if (topic === 'เงินเก็บ') {
    return '💰 ต่อไป: 1 ฝาก | 2 ถอน | 3 ยอด | 4 ย้อนหลัง 15 วัน\n0 เมนูหลัก | 9 ออก';
  }
  return '📅 ต่อไป: 1 บันทึกนัดอีก | 2 เมนูหลัก | 3 ออก';
}

function handleRecordingFlow_(target, text) {
  const lower = text.toLowerCase();
  if (lower === '/help' || text === 'ช่วยเหลือ') return null;
  const state = getRecordingState_(target.ownerKey);
  if (!state) return null;

  if (lower === '/cancel') {
    if (state.indexOf('PERMISSION_') === 0) setPendingPermission_(target.ownerKey, null);
    if (state === 'TRANSACTION_DATE_PICK' || state === 'INSTALLMENT_LOAN_PICK' || state.indexOf('APPOINTMENT') === 0) setPendingAppointment_(target.ownerKey, null);
    setRecordingState_(target, 'SELECT_TOPIC');
    return 'ยกเลิกขั้นตอนนี้แล้วค่ะ\n' + topicPrompt_(target);
  }

  if (state === 'INSTALLMENT_LOAN_PICK') {
    const pending = getPendingAppointment_(target.ownerKey) || {};
    const loans = pending.actionState === 'INSTALLMENT_HISTORY' ? getInstallmentSummaries_(target.ownerKey).filter(function(loan) {
      return hasInstallmentHistoryInWindow_(loan, installmentHistoryWindow_());
    }) :
      activeInstallmentLoans_(target.ownerKey).filter(function(loan) {
        return pending.actionState === 'INSTALLMENT_PAYMENT' || pending.actionState === 'INSTALLMENT_INTEREST_CALC' ? loan.balance > 0 : loan.interestDue > 0;
      });
    if (text === '0') {
      setPendingAppointment_(target.ownerKey, null);
      setRecordingState_(target, 'INSTALLMENT_MENU');
      return installmentMenu_();
    }
    const index = Number(text) - 1;
    if (!Number.isInteger(index) || index < 0 || index >= loans.length) return selectInstallmentForAction_(target, pending.actionState, pending.entryText);
    const loan = loans[index];
    setSelectedInstallmentId_(target.ownerKey, loan.loanId);
    setPendingAppointment_(target.ownerKey, null);
    if (pending.actionState === 'INSTALLMENT_HISTORY') {
      setRecordingState_(target, 'INSTALLMENT_MENU');
      return getInstallmentHistory_(target.ownerKey, loan.loanId) + '\n\n' + installmentMenu_();
    }
    setRecordingState_(target, pending.actionState);
    if (pending.actionState === 'INSTALLMENT_INTEREST_CALC') return calculateInterestFromLatestBalance_(target, loan);
    return pending.entryText ? handleRecordingFlow_(target, pending.entryText) :
      (pending.actionState === 'INSTALLMENT_PAYMENT' ? installmentPaymentPrompt_(loan) : installmentInterestPaymentPrompt_(loan));
  }

  if (state === 'TRANSACTION_DATE_PICK') {
    if (text === 'วันนี้' || lower === 'today') return completeTransactionDateSelection_(target, todayDateKey_());
    if (text === '0') {
      setPendingAppointment_(target.ownerKey, null);
      setRecordingState_(target, 'SELECT_TOPIC');
      return 'ยกเลิกการเลือกรายการวันที่แล้วค่ะ\n\n' + topicPrompt_(target);
    }
    return transactionDatePrompt_('');
  }

  if (state.indexOf('PERMISSION_') === 0 && !isAdmin_(target.actorId)) {
    return 'เมนูกำหนดสิทธิ์สงวนไว้สำหรับผู้ดูแลระบบค่ะ';
  }

  if (state === 'PERMISSION_SCOPE_SELECT') {
    if (text === '0') {
      setPendingPermission_(target.ownerKey, null);
      setRecordingState_(target, 'SELECT_TOPIC');
      return topicPrompt_(target);
    }
    const groups = getRegisteredGroupTargets_();
    const index = Number(text) - 1;
    if (!Number.isInteger(index) || index < 0 || index >= groups.length) return permissionScopeMenu_();
    const selected = groups[index];
    const context = {scopeKey: selected.scopeKey, targetId: selected.targetId, groupName: selected.groupName};
    setPendingPermission_(target.ownerKey, context);
    setRecordingState_(target, 'PERMISSION_MENU');
    return permissionMenu_(context);
  }

  if (state === 'PERMISSION_MENU') {
    const context = getPendingPermission_(target.ownerKey);
    if (!context || !context.scopeKey) {
      setRecordingState_(target, 'PERMISSION_SCOPE_SELECT');
      return permissionScopeMenu_();
    }
    if (text === '1') {
      setRecordingState_(target, 'PERMISSION_SET_GROUP');
      return permissionActionMenu_(false);
    }
    if (text === '2') {
      setRecordingState_(target, 'PERMISSION_MEMBER_ID');
      return 'ส่ง LINE User ID ของสมาชิกในกลุ่มนี้ค่ะ\nตัวอย่างรูปแบบ: Uxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx\n\nพิมพ์ 0 เพื่อกลับเมนูกลุ่ม';
    }
    if (text === '0') {
      setRecordingState_(target, 'PERMISSION_SCOPE_SELECT');
      return permissionScopeMenu_();
    }
    return permissionMenu_(context);
  }

  if (state === 'PERMISSION_MEMBER_ID') {
    const context = getPendingPermission_(target.ownerKey) || {};
    if (text === '0') {
      setRecordingState_(target, 'PERMISSION_MENU');
      return permissionMenu_(context);
    }
    if (!/^U[0-9a-f]{32}$/i.test(text)) {
      return 'LINE User ID ไม่ถูกต้องค่ะ กรุณาส่ง User ID ที่ขึ้นต้นด้วย U และมีอักขระ 32 ตัว';
    }
    context.userId = text;
    context.memberName = getLineGroupMemberName_(context.targetId, text);
    setPendingPermission_(target.ownerKey, context);
    setRecordingState_(target, 'PERMISSION_SET_MEMBER');
    return permissionActionMenu_(true, context.memberName);
  }

  if (state === 'PERMISSION_SET_GROUP' || state === 'PERMISSION_SET_MEMBER') {
    const draft = getPendingPermission_(target.ownerKey);
    if (!draft || !draft.scopeKey) {
      setRecordingState_(target, 'PERMISSION_SCOPE_SELECT');
      return 'ไม่พบกลุ่มเป้าหมายค่ะ\n' + permissionScopeMenu_();
    }
    if (text === '0') {
      delete draft.userId;
      delete draft.memberName;
      setPendingPermission_(target.ownerKey, draft);
      setRecordingState_(target, 'PERMISSION_MENU');
      return permissionMenu_(draft);
    }
    const userId = state === 'PERMISSION_SET_MEMBER' ? draft.userId : '';
    if (state === 'PERMISSION_SET_MEMBER' && !userId) {
      setRecordingState_(target, 'PERMISSION_MENU');
      return 'ไม่พบสมาชิกที่เลือกค่ะ\n\n' + permissionMenu_(draft);
    }
    const result = applyPermissionChoice_(target, userId, text, draft.scopeKey, draft.memberName);
    if (!result) return 'รูปแบบสิทธิ์ไม่ถูกต้องค่ะ กรุณาตอบ Y,N,Y หรือ 9 เพื่อคืนค่าเริ่มต้น แล้วลองใหม่\n\n' + permissionActionMenu_(state === 'PERMISSION_SET_MEMBER', draft.memberName);
    delete draft.userId;
    delete draft.memberName;
    setPendingPermission_(target.ownerKey, draft);
    setRecordingState_(target, 'PERMISSION_MENU');
    return result + '\n\n' + permissionMenu_(draft);
  }
  if (text === '0' && ['APPOINTMENT', 'APPOINTMENT_DATE', 'APPOINTMENT_FIELDS', 'APPOINTMENT_REMINDER', 'APPOINTMENT_SUMMARY_DATE', 'APPOINTMENT_DELETE_PICK', 'SAVINGS_DEPOSIT', 'SAVINGS_WITHDRAW'].includes(state)) {
    setPendingAppointment_(target.ownerKey, null);
    const menuState = state === 'SAVINGS_DEPOSIT' || state === 'SAVINGS_WITHDRAW' ? 'SAVINGS_MENU' : 'APPOINTMENT_MENU';
    setRecordingState_(target, menuState);
    return 'ออกจากการกรอกแล้วค่ะ\n' + (menuState === 'SAVINGS_MENU' ? savingsMenu_() : appointmentMenu_());
  }

  if (state === 'POST_SAVE_APPOINTMENT' || state === 'POST_SAVE_SAVINGS') {
    const returnState = state === 'POST_SAVE_APPOINTMENT' ? 'APPOINTMENT_MENU' : 'SAVINGS_MENU';
    if (state === 'POST_SAVE_SAVINGS') {
      if (text === '1' || ['ฝาก', 'บันทึกเงินฝาก'].includes(text)) {
        setRecordingState_(target, 'SAVINGS_DEPOSIT');
        return savingsEntryPrompt_('ฝาก');
      }
      if (text === '2' || ['ถอน', 'บันทึกเงินถอน'].includes(text)) {
        setRecordingState_(target, 'SAVINGS_WITHDRAW');
        return savingsEntryPrompt_('ถอน');
      }
      if (text === '3') return getSavingsBalance_(target.ownerKey) + '\n\n' + afterSavePrompt_('เงินเก็บ');
      if (text === '4') return getRecentSavings_(target.ownerKey) + '\n\n' + afterSavePrompt_('เงินเก็บ');
      if (text === '0') {
        setRecordingState_(target, 'SELECT_TOPIC');
        return topicPrompt_(target);
      }
      if (text === '9') {
        endChatSession_(target);
        return 'ออกจากการใช้งานบอทแล้วค่ะ พิมพ์รหัสคำสั่งที่ต้องการเพื่อเริ่มใหม่ได้เลย';
      }
      return afterSavePrompt_('เงินเก็บ');
    }
    if (text === '1' || text === 'บันทึกต่อ') {
      setRecordingState_(target, returnState);
      return returnState === 'APPOINTMENT_MENU' ? appointmentMenu_() : savingsMenu_();
    }
    if (text === '2' || text === 'หน้าหลัก' || text === 'กลับหน้าหลัก') {
      setRecordingState_(target, 'SELECT_TOPIC');
      return topicPrompt_(target);
    }
    if (text === '3' || text === 'ออก' || text === 'ออกจากบอท') {
      endChatSession_(target);
      return 'ออกจากการใช้งานบอทแล้วค่ะ พิมพ์รหัสคำสั่งที่ต้องการเพื่อเริ่มใหม่ได้เลย';
    }
    return afterSavePrompt_(state === 'POST_SAVE_APPOINTMENT' ? 'นัดหมาย' : 'เงินเก็บ');
  }

  if (state === 'POST_SAVE_INSTALLMENT') {
    if (text === '1') {
      return selectInstallmentForAction_(target, 'INSTALLMENT_PAYMENT', '');
    }
    if (text === '2') {
      return selectInstallmentForAction_(target, 'INSTALLMENT_INTEREST_PAYMENT', '');
    }
    if (text === '3') return getInstallmentBalance_(target.ownerKey) + '\n\n' + installmentAfterSaveMenu_();
    if (text === '4') return getInstallmentInterestDue_(target.ownerKey) + '\n\n' + installmentAfterSaveMenu_();
    if (text === '0') {
      setRecordingState_(target, 'SELECT_TOPIC');
      return topicPrompt_(target);
    }
    if (text === '9') {
      endChatSession_(target);
      return 'ออกจากการใช้งานบอทแล้วค่ะ พิมพ์รหัสคำสั่งที่ต้องการเพื่อเริ่มใหม่ได้เลย';
    }
    return installmentAfterSaveMenu_();
  }

  if (state === 'SELECT_TOPIC') {
    const command = text.trim().toUpperCase();
    if (command === 'A1') {
      if (!isAdmin_(target.actorId)) return 'เมนูกำหนดสิทธิ์สงวนไว้สำหรับผู้ดูแลระบบค่ะ';
      if (target.targetType !== 'user') return 'กรุณากำหนดสิทธิ์ผ่านแชตส่วนตัวกับบอทค่ะ';
      if (!SpreadsheetApp.getActive().getSheetByName(CONFIG.PERMISSIONS_SHEET)) {
        return 'ยังไม่พบชีต Permissions ค่ะ กรุณารัน setup() หนึ่งครั้งก่อนใช้งานเมนูกำหนดสิทธิ์';
      }
      setPendingPermission_(target.ownerKey, null);
      setRecordingState_(target, 'PERMISSION_SCOPE_SELECT');
      return permissionScopeMenu_();
    }
    if (['C1', 'นัดหมาย'].includes(command)) {
      setRecordingState_(target, 'APPOINTMENT_MENU');
      return appointmentMenu_();
    }
    if (command === 'C2') {
      setPendingAppointment_(target.ownerKey, null);
      setRecordingState_(target, 'APPOINTMENT_FIELDS');
      return appointmentFieldsPrompt_();
    }
    if (['C3', 'สรุปงานตามวันที่'].includes(command)) {
      setRecordingState_(target, 'APPOINTMENT_SUMMARY_DATE');
      return appointmentSummaryDatePrompt_();
    }
    if (['M1', 'เงินเก็บ'].includes(command)) {
      setRecordingState_(target, 'SAVINGS_MENU');
      return savingsMenu_();
    }
    if (['M2', 'ฝากเงิน', 'ฝาก'].includes(command)) {
      setRecordingState_(target, 'SAVINGS_DEPOSIT');
      return savingsEntryPrompt_('ฝาก');
    }
    if (['M3', 'ถอนเงิน', 'ถอน'].includes(command)) {
      setRecordingState_(target, 'SAVINGS_WITHDRAW');
      return savingsEntryPrompt_('ถอน');
    }
    if (['M4', 'เช็คยอดเงินฝาก BALANCE', 'เช็คยอดเงิน BALANCE', 'ยอดเงินเก็บ', 'ยอดคงเหลือ', 'ยอด BALANCE', 'BALANCE', '/BALANCE'].includes(command)) {
      return getSavingsBalance_(target.ownerKey) + '\n\n' + topicPrompt_(target);
    }
    if (['L1', 'ผ่อนเงิน'].includes(command)) {
      setRecordingState_(target, 'INSTALLMENT_MENU');
      return installmentMenu_();
    }
    if (command === 'L2' || command === 'ยอดเงินผ่อนคงค้าง') {
      return getInstallmentBalance_(target.ownerKey) + '\n\n' + topicPrompt_(target);
    }
    if (['ดอกเบี้ยค้าง', 'ดอกเบี้ยคงค้าง', 'สอบถามดอกเบี้ย', 'สอบถามดอกเบี้ยค้างชำระ', 'ดูดอกเบี้ยค้างชำระ'].includes(command)) {
      return getInstallmentInterestDue_(target.ownerKey) + '\n\n' + topicPrompt_(target);
    }
    if (command === 'X') {
      endChatSession_(target);
      return 'ออกจากการใช้งานบอทแล้วค่ะ พิมพ์รหัสคำสั่งที่ต้องการเพื่อเริ่มใหม่ได้เลย';
    }
    return 'กรุณาเลือกรหัสคำสั่งจากเมนูค่ะ\n\n' + topicPrompt_(target);
  }

  if (state === 'APPOINTMENT_MENU') {
    if (['1', 'บันทึกนัดหมาย'].includes(text)) {
      setPendingAppointment_(target.ownerKey, null);
      setRecordingState_(target, 'APPOINTMENT_FIELDS');
      return appointmentFieldsPrompt_();
    }
    if (['2', 'สรุปงาน'].includes(text)) {
      setRecordingState_(target, 'APPOINTMENT_SUMMARY_DATE');
      return appointmentSummaryDatePrompt_();
    }
    if (['3', 'ลบงาน'].includes(text)) {
      return startAppointmentDelete_(target.ownerKey);
    }
    if (text === '0') {
      setRecordingState_(target, 'SELECT_TOPIC');
      return topicPrompt_(target);
    }
    if (text === '9') {
      endChatSession_(target);
      return 'ออกจากการใช้งานบอทแล้วค่ะ พิมพ์รหัสคำสั่งที่ต้องการเพื่อเริ่มใหม่ได้เลย';
    }
    return 'เลือก 1, 2, 3, 0 หรือ 9 ค่ะ\n' + appointmentMenu_();
  }

  if (state === 'SAVINGS_MENU') {
    if (['1', 'ฝาก', 'บันทึกเงินฝาก'].includes(text)) {
      setRecordingState_(target, 'SAVINGS_DEPOSIT');
      return savingsEntryPrompt_('ฝาก');
    }
    if (['2', 'ถอน', 'บันทึกเงินถอน'].includes(text)) {
      setRecordingState_(target, 'SAVINGS_WITHDRAW');
      return savingsEntryPrompt_('ถอน');
    }
    if (['3', 'ยอดเงินเก็บ', 'ยอดคงเหลือ', 'ยอด balance', 'balance', '/balance'].includes(text.toLowerCase())) {
      return getSavingsBalance_(target.ownerKey) + '\n\n' + savingsMenu_();
    }
    if (['4', 'ย้อนหลัง 15 วัน', 'รายละเอียด 15 วันย้อนหลัง', 'รายการเงินเก็บย้อนหลัง'].includes(text.toLowerCase())) {
      return getRecentSavings_(target.ownerKey) + '\n\n' + savingsMenu_();
    }
    if (text === '0') {
      setRecordingState_(target, 'SELECT_TOPIC');
      return topicPrompt_(target);
    }
    if (text === '9') {
      endChatSession_(target);
      return 'ออกจากการใช้งานบอทแล้วค่ะ พิมพ์รหัสคำสั่งที่ต้องการเพื่อเริ่มใหม่ได้เลย';
    }
    return 'เลือก 1, 2, 3, 4, 0 หรือ 9 ค่ะ\n' + savingsMenu_();
  }

  if (state === 'INSTALLMENT_MENU') {
    if (['1', 'ตั้งยอด', 'กำหนดยอดเงินต้น'].includes(text.toLowerCase())) {
      setRecordingState_(target, 'INSTALLMENT_PRINCIPAL');
      return '💳 ตั้งสัญญาผ่อน\nหัวเรื่อง | เจ้าหนี้ | ลูกหนี้ | เงินต้น | ดอกเบี้ย% | หมายเหตุ | ฐาน\nฐาน: ยอดตั้งต้น/ยอดคงเหลือ (เว้นได้)\nตัวอย่าง: รถ | Mama | Kwan | 300000 | 1% | | ยอดตั้งต้น\nเลือกวันที่ถัดไป | 0 ยกเลิก';
    }
    if (['2', 'ผ่อน', 'บันทึกค่างวด', 'ชำระเงินต้น'].includes(text.toLowerCase())) {
      return selectInstallmentForAction_(target, 'INSTALLMENT_PAYMENT', '');
    }
    if (['3', 'ดอกเบี้ย', 'ชำระดอกเบี้ย'].includes(text.toLowerCase())) {
      return selectInstallmentForAction_(target, 'INSTALLMENT_INTEREST_PAYMENT', '');
    }
    if (['4', 'ยอด', 'ยอดคงเหลือ', 'balance'].includes(text.toLowerCase())) return getInstallmentBalance_(target.ownerKey) + '\n\n' + installmentMenu_();
    if (['5', 'ประวัติ', 'ย้อนหลัง'].includes(text.toLowerCase())) return selectInstallmentForAction_(target, 'INSTALLMENT_HISTORY', '');
    if (['6', 'ดอกเบี้ยค้าง', 'ดอกเบี้ยคงค้าง', 'สอบถามดอกเบี้ย', 'สอบถามดอกเบี้ยค้างชำระ', 'ดูดอกเบี้ยค้างชำระ'].includes(text.toLowerCase())) return getInstallmentInterestDue_(target.ownerKey) + '\n\n' + installmentMenu_();
    if (text === '7' || ['คำนวณดอกเบี้ย', 'คำนวณดอกเบี้ยจากยอดล่าสุด'].includes(text.toLowerCase())) return selectInstallmentForAction_(target, 'INSTALLMENT_INTEREST_CALC', '');
    if (text === '0') {
      setRecordingState_(target, 'SELECT_TOPIC');
      return topicPrompt_(target);
    }
    if (text === '9') {
      endChatSession_(target);
      return 'ออกจากการใช้งานบอทแล้วค่ะ พิมพ์ “hi bot” เมื่อต้องการเริ่มใหม่';
    }
    return 'เลือก 1, 2, 3, 4, 5, 6, 7, 0 หรือ 9 ค่ะ\n' + installmentMenu_();
  }

  if (state === 'INSTALLMENT_PRINCIPAL') {
    if (text === '0') {
      setRecordingState_(target, 'INSTALLMENT_MENU');
      return installmentMenu_();
    }
    const entryText = text.replace(/^เงินตั้งต้นผ่อน\s*/i, '');
    const inputParts = entryText.split('|').map(function(part) { return part.trim(); });
    if (inputParts.length < 5 || inputParts.length > 8) return installmentPrincipalFormatError_();
    const title = inputParts[0];
    const creditor = inputParts[1];
    const debtor = inputParts[2];
    const principalText = normalizeMoneyField_(inputParts[3]);
    const rateText = normalizeRateField_(inputParts[4]);
    let note = 'ตั้งยอดตั้งต้น';
    let interestBasis = 'ORIGINAL';
    let transactionDate = '';
    if (inputParts.length === 6) {
      const basis = parseInterestBasis_(inputParts[5]);
      if (basis) interestBasis = basis;
      else note = inputParts[5] || note;
    } else if (inputParts.length === 7) {
      const sixthBasis = parseInterestBasis_(inputParts[5]);
      const seventhBasis = parseInterestBasis_(inputParts[6]);
      if (sixthBasis) {
        interestBasis = sixthBasis;
        transactionDate = inputParts[6];
      } else {
        note = inputParts[5] || note;
        if (seventhBasis) interestBasis = seventhBasis;
        else transactionDate = inputParts[6];
      }
    } else if (inputParts.length === 8) {
      note = inputParts[5] || note;
      interestBasis = parseInterestBasis_(inputParts[6]);
      transactionDate = inputParts[7];
    }
    if (!title || !creditor || !debtor || !isPositiveMoney_(principalText) || !isNonNegativeRate_(rateText) ||
        !interestBasis || (transactionDate && !isValidDate_(transactionDate))) return installmentPrincipalFormatError_();
    const principalEntry = [title, creditor, debtor, principalText, rateText, note, interestBasis].join(' | ');
    if (!transactionDate) return beginTransactionDateSelection_(target, 'INSTALLMENT_PRINCIPAL', principalEntry, 'ตั้งยอดเงินผ่อน');
    const result = addInstallmentPrincipal_(target, title, creditor, debtor, Number(principalText), Number(rateText), note, transactionDate, interestBasis);
    setRecordingState_(target, 'POST_SAVE_INSTALLMENT');
    return result + '\n\n' + installmentAfterSaveMenu_();
  }

  if (state === 'INSTALLMENT_PAYMENT') {
    if (text === '0') {
      setRecordingState_(target, 'INSTALLMENT_MENU');
      return installmentMenu_();
    }
    const parts = text.split('|').map(function(part) { return part.trim(); });
    if ((parts.length !== 2 && parts.length !== 3) || !isPositiveMoney_(parts[0]) || !parts[1] || (parts.length === 3 && parts[2] && !isValidDate_(parts[2]))) {
      return 'รูปแบบไม่ถูกต้องค่ะ กรอก: จำนวนเงินต้นที่ชำระ | หมายเหตุ\nตัวอย่าง: 3000 | ชำระเงินต้นงวดนี้\nรายการนี้จะลดเฉพาะเงินต้นค่ะ\n0 — ยกเลิก';
    }
    const loan = getInstallmentSummary_(target.ownerKey, getSelectedInstallmentId_(target.ownerKey));
    const principal = Number(parts[0]);
    if (principal > loan.balance) return 'ยอดผ่อนเงินต้นมากกว่ายอดคงเหลือ ' + formatMoney_(loan.balance) + ' บาทค่ะ กรุณากรอกใหม่';
    if (!parts[2]) return beginTransactionDateSelection_(target, 'INSTALLMENT_PAYMENT', parts.slice(0, 2).join(' | '), 'จ่ายเงินต้น');
    const note = parts[1];
    const result = addInstallmentPayment_(target, loan, principal, note, loan.balance - principal, parts[2]);
    setRecordingState_(target, 'POST_SAVE_INSTALLMENT');
    return result + '\n\n' + installmentAfterSaveMenu_();
  }

  if (state === 'INSTALLMENT_INTEREST_PAYMENT') {
    if (text === '0') {
      setRecordingState_(target, 'INSTALLMENT_MENU');
      return installmentMenu_();
    }
    const parts = text.split('|').map(function(part) { return part.trim(); });
    if ((parts.length !== 2 && parts.length !== 3) || !isPositiveMoney_(parts[0]) || !parts[1] || (parts.length === 3 && parts[2] && !isValidDate_(parts[2]))) {
      return 'รูปแบบไม่ถูกต้องค่ะ กรอก: จำนวนดอกเบี้ยที่ชำระ | หมายเหตุ\nตัวอย่าง: 750 | ชำระดอกเบี้ยปีแรก\nรายการนี้จะลดเฉพาะดอกเบี้ยค้างชำระค่ะ\n0 — ยกเลิก';
    }
    const loan = getInstallmentSummary_(target.ownerKey, getSelectedInstallmentId_(target.ownerKey));
    const amount = Number(parts[0]);
    if (amount > loan.interestDue) return 'จำนวนเงินมากกว่าดอกเบี้ยค้างชำระ ' + formatMoney_(loan.interestDue) + ' บาทค่ะ กรุณากรอกใหม่';
    if (!parts[2]) return beginTransactionDateSelection_(target, 'INSTALLMENT_INTEREST_PAYMENT', parts.slice(0, 2).join(' | '), 'จ่ายดอกเบี้ย');
    const result = addInstallmentInterestPayment_(target, loan, amount, parts[1], parts[2]);
    setRecordingState_(target, 'POST_SAVE_INSTALLMENT');
    return result + '\n\n' + installmentAfterSaveMenu_();
  }

  if (state === 'APPOINTMENT') {
    const appointment = parseAppointment_(text);
    if (!appointment) return appointmentFormatError_();
    setPendingAppointment_(target.ownerKey, appointment);
    setRecordingState_(target, 'APPOINTMENT_REMINDER');
    return appointmentReminderPrompt_();
  }

  if (state === 'APPOINTMENT_DATE') {
    if (!isValidDate_(text)) return appointmentDatePrompt_();
    const appointment = getPendingAppointment_(target.ownerKey) || {};
    appointment.date = text;
    setPendingAppointment_(target.ownerKey, appointment);
    setRecordingState_(target, 'APPOINTMENT_REMINDER');
    return '📅 วันที่นัดหมาย: ' + formatDateWeekday_(text) + '\n\n' + appointmentReminderPrompt_();
  }

  if (state === 'APPOINTMENT_FIELDS') {
    const parts = text.split('|').map(function(part) { return part.trim(); });
    if ((parts.length !== 2 && parts.length !== 3) ||
        (parts.length === 2 && !parts[0]) ||
        (parts.length === 3 && ((parts[0] !== '' && !/^([01]\d|2[0-3]):[0-5]\d$/.test(parts[0])) || !parts[1]))) {
      return 'รูปแบบไม่ถูกต้องค่ะ กรุณากรอกรายละเอียด | สถานที่\nถ้าไม่มีสถานที่ ให้เว้นช่องว่างแต่คงเครื่องหมาย | ไว้\nตัวอย่าง: ประชุมทีม | ห้องประชุม';
    }
    const appointment = {
      time: parts.length === 3 ? parts[0] : '',
      description: parts.length === 3 ? parts[1] : parts[0],
      location: parts.length === 3 ? parts[2] : parts[1]
    };
    setPendingAppointment_(target.ownerKey, appointment);
    setRecordingState_(target, 'APPOINTMENT_DATE');
    return appointmentDatePrompt_();
  }

  if (state === 'APPOINTMENT_REMINDER') {
      if (!['1', '2', '3'].includes(text)) return 'กรุณาเลือก 1, 2 หรือ 3 ค่ะ\n\n' + appointmentReminderPrompt_();
    const appointment = getPendingAppointment_(target.ownerKey);
    if (!appointment) {
      setRecordingState_(target, 'APPOINTMENT_MENU');
      return 'ไม่พบข้อมูลนัดหมายที่รอบันทึก กรุณาเริ่มรายการใหม่';
    }
    const reminderMode = text === '1' ? 'ADVANCE_1_DAY' : (text === '3' ? 'BOTH' : 'ON_DATE');
    const result = saveAppointment_(target, appointment, reminderMode);
    setPendingAppointment_(target.ownerKey, null);
    setRecordingState_(target, 'POST_SAVE_APPOINTMENT');
    return result + '\nแจ้งเตือน: ' + appointmentReminderLabel_(reminderMode) +
      '\n\n' + afterSavePrompt_('นัดหมาย');
  }

  if (state === 'SAVINGS_DEPOSIT' || state === 'SAVINGS_WITHDRAW') {
    const type = state === 'SAVINGS_DEPOSIT' ? 'ฝาก' : 'ถอน';
    const parts = text.split('|').map(s => s.trim());
    let savingsEntryLine = text;
    if (parts.length === 2 || parts.length === 3) {
      if (!isPositiveMoney_(parts[0]) || !parts[1] || (parts.length === 3 && parts[2] && !isValidDate_(parts[2]))) {
        return 'รูปแบบไม่ถูกต้องค่ะ กรอก: จำนวนเงิน | หมายเหตุ\nตัวอย่าง: 500 | เก็บเงินเดือน\n0 — ออกจากการกรอกกลับเมนูเงินเก็บ';
      }
      if (!parts[2]) return beginTransactionDateSelection_(target, state, parts.slice(0, 2).join(' | '), type === 'ฝาก' ? 'ฝากเงิน' : 'ถอนเงิน');
      const transactionDate = optionalTransactionDate_(parts[2]);
      savingsEntryLine = [transactionDate, type, parts[0], parts[1]].join(' | ');
    }
    const result = addSavings_(target, savingsEntryLine);
    if (result.startsWith('บันทึกเงินเก็บแล้ว')) {
      setRecordingState_(target, 'POST_SAVE_SAVINGS');
      return result + '\n\n' + getSavingsBalance_(target.ownerKey) + '\n\n' + afterSavePrompt_('เงินเก็บ');
    }
    return result + '\n0 — ออกจากการกรอกกลับเมนูเงินเก็บ';
  }

  if (state === 'APPOINTMENT_SUMMARY_DATE') {
    if (!isValidDate_(text)) return appointmentSummaryDatePrompt_();
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
      return 'ไม่พบรายการที่เลือกลบได้ค่ะ\n' + appointmentMenu_();
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
  return topicPrompt_(target);
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
  return formatAppointmentLine_(date, time, description, location, id);
}

function appointmentFormatError_() {
  return 'รูปแบบไม่ถูกต้องค่ะ ต้องกรอกให้ครบ 4 ช่อง:\nYYYY-MM-DD | HH:mm | รายละเอียด | สถานที่\nถ้าไม่มีเวลา/สถานที่ ให้เว้นช่องว่างแต่คงเครื่องหมาย | ไว้\nตัวอย่าง: 2026-10-03 |  | ประชุมทีม |\n0 — ออกจากการกรอกกลับเมนูนัดหมาย';
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
    .sort((a, b) => formatTaskTime_(a[4]).localeCompare(formatTaskTime_(b[4])));
  return items.length ? '📅 ' + label + ' (' + formatDateWeekday_(date) + ')\n\n' + items.map(formatTask_).join('\n\n') : '📅 ไม่มีนัดหมายที่บันทึกไว้ใน' + label + ' (' + formatDateWeekday_(date) + ') ค่ะ';
}

function formatTask_(row) {
  const parts = ['📅'];
  const time = formatTaskTime_(row[4]);
  if (time) parts.push(time);
  if (row[2]) parts.push(String(row[2]));
  if (row[7]) parts.push(String(row[7]));
  if (row[0]) parts.push('[' + row[0] + ']');
  return parts.join(' ');
}

function formatReminderTask_(row) {
  const scheduledDate = cellDateKey_(row[3]);
  const details = formatTask_(row).replace(/^📅\s*/, '');
  return '📅 วันนัด: ' + formatDateWeekday_(scheduledDate) + (details ? '\n' + details : '');
}

function formatTaskTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, 'HH:mm');
  }
  const text = String(value || '').trim();
  const match = text.match(/(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?(?:\s|$)/);
  return match ? ('0' + match[1]).slice(-2) + ':' + match[2] : text;
}

function formatAppointmentLine_(date, time, description, location, id) {
  const parts = ['📅 ' + formatDateWeekday_(date)];
  if (time) parts.push(String(time));
  if (description) parts.push(String(description));
  if (location) parts.push(String(location));
  if (id) parts.push('[' + id + ']');
  return parts.join(' ');
}

function formatDateWeekday_(dateKey) {
  if (!isValidDate_(dateKey)) return dateKey || '-';
  const weekdays = ['วันอาทิตย์', 'วันจันทร์', 'วันอังคาร', 'วันพุธ', 'วันพฤหัสบดี', 'วันศุกร์', 'วันเสาร์'];
  const dayIndex = new Date(dateKey + 'T00:00:00Z').getUTCDay();
  const dateParts = dateKey.split('-');
  return weekdays[dayIndex] + 'ที่ ' + dateParts[2] + '/' + dateParts[1] + '/' + dateParts[0].slice(-2);
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
  if (!match || !isValidDate_(match[1])) return 'วันที่ไม่ถูกต้องค่ะ ใช้รูปแบบ YYYY-MM-DD เช่น 2026-10-03';
  const date = match[1];
  const rows = SpreadsheetApp.getActive().getSheetByName(CONFIG.TASKS_SHEET).getDataRange().getValues().slice(1);
  const items = rows.filter(r => String(r[1]) === ownerKey && cellDateKey_(r[3]) === date && String(r[5]).toUpperCase() === 'OPEN')
    .sort((a, b) => formatTaskTime_(a[4]).localeCompare(formatTaskTime_(b[4])));
  return items.length ? '📅 สรุปงานวันที่ ' + formatDateWeekday_(date) + ' (' + items.length + ' งาน)\n\n' + items.map(formatTask_).join('\n\n')
    : '📅 ไม่มีงานที่ยังใช้งานอยู่ในวันที่ ' + formatDateWeekday_(date) + ' ค่ะ';
}

function deleteTask_(ownerKey, text) {
  const parts = text.trim().split(/\s+/);
  if (parts.length !== 2 || parts[0] !== 'ลบงาน') return 'รูปแบบไม่ถูกต้องค่ะ ใช้: ลบงาน <รหัสงาน>';
  const taskId = parts[1].toLowerCase();
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.TASKS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === taskId && String(rows[i][1]) === ownerKey && rows[i][5] === 'OPEN') {
      sheet.getRange(i + 1, 6).setValue('DELETED');
      return 'ลบนัดหมายรหัส ' + rows[i][0] + ' เรียบร้อยแล้วค่ะ';
    }
  }
  return 'ไม่พบงานที่ยังใช้งานอยู่รหัสนี้ในแชตปัจจุบันค่ะ';
}

function startAppointmentDelete_(ownerKey) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.TASKS_SHEET);
  const today = todayDateKey_();
  const rows = sheet.getDataRange().getValues().slice(1)
    .filter(function(row) {
      const date = cellDateKey_(row[3]);
      return String(row[1]) === ownerKey && String(row[5]).toUpperCase() === 'OPEN' && date && date >= today;
    })
    .sort(function(a, b) {
      return cellDateKey_(a[3]).localeCompare(cellDateKey_(b[3])) || formatTaskTime_(a[4]).localeCompare(formatTaskTime_(b[4]));
    });
  if (!rows.length) return 'ไม่มีนัดหมายตั้งแต่วันนี้เป็นต้นไปให้ลบค่ะ\n\n' + appointmentMenu_();
  setPendingAppointment_(ownerKey, {deleteIds: rows.map(function(row) { return String(row[0]); })});
  setRecordingState_(getTargetFromOwner_(ownerKey), 'APPOINTMENT_DELETE_PICK');
  return '🗑️ เลือกนัดหมายที่ต้องการลบ โดยพิมพ์หมายเลข:\n' +
    rows.map(function(row, index) { return (index + 1) + '. ' + formatDateWeekday_(cellDateKey_(row[3])) + ' | ' + formatTask_(row); }).join('\n') +
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
  if (!pending || !pending.deleteIds) return 'ไม่พบรายการที่เลือกลบได้ค่ะ';
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.TASKS_SHEET);
  const today = todayDateKey_();
  const rows = sheet.getDataRange().getValues().slice(1)
    .filter(function(row) {
      const date = cellDateKey_(row[3]);
      return pending.deleteIds.indexOf(String(row[0])) >= 0 && String(row[1]) === ownerKey &&
        String(row[5]).toUpperCase() === 'OPEN' && date && date >= today;
    });
  if (rows.length) setPendingAppointment_(ownerKey, {deleteIds: rows.map(function(row) { return String(row[0]); })});
  return rows.length
    ? '🗑️ เลือกหมายเลขนัดหมายที่ต้องการลบ:\n' + rows.map(function(row, index) { return (index + 1) + '. ' + formatDateWeekday_(cellDateKey_(row[3])) + ' | ' + formatTask_(row); }).join('\n') + '\n0 — ยกเลิกและกลับเมนูนัดหมาย'
    : 'ไม่มีนัดหมายตั้งแต่วันนี้เป็นต้นไปให้ลบค่ะ';
}

function addSavings_(target, text) {
  const parts = text.split('|').map(s => s.trim());
  if (parts.length !== 4 || !isValidDate_(parts[0]) || !['ฝาก', 'ถอน'].includes(parts[1]) ||
      !/^\d+(\.\d{1,2})?$/.test(parts[2]) || Number(parts[2]) <= 0 || !parts[3]) {
    return 'รูปแบบไม่ถูกต้องค่ะ ใช้: YYYY-MM-DD | ฝาก/ถอน | จำนวนเงิน | หมายเหตุ\nตัวอย่าง: 2026-10-03 | ฝาก | 500 | เก็บเงินเดือน';
  }
  const id = Utilities.getUuid().slice(0, 8);
  SpreadsheetApp.getActive().getSheetByName(CONFIG.SAVINGS_SHEET)
    .appendRow([id, target.ownerKey, parts[0], parts[1], Number(parts[2]), parts[3], new Date()]);
  return '✅ บันทึกเงินเก็บแล้วค่ะ (รหัส ' + id + ')\n📅 วันที่: ' + parts[0] +
    '\n' + (parts[1] === 'ฝาก' ? '📥 ฝาก' : '📤 ถอน') + ' ' + formatMoney_(Number(parts[2])) + ' บาท' +
    '\n📝 หมายเหตุ: ' + parts[3];
}

function isBalanceCommand_(text) {
  return ['ยอดเงินเก็บ', 'ยอดคงเหลือ', 'ยอด balance', 'balance', '/balance'].includes(text.toLowerCase());
}

function getSavingsBalance_(ownerKey) {
  const rows = SpreadsheetApp.getActive().getSheetByName(CONFIG.SAVINGS_SHEET).getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) === ownerKey);
  const currentMonth = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM');
  let balanceDeposits = 0;
  let balanceWithdrawals = 0;
  let monthDeposits = 0;
  let monthWithdrawals = 0;
  rows.forEach(r => {
    const amount = Number(r[4]) || 0;
    if (r[3] === 'ฝาก') {
      balanceDeposits += amount;
      if (cellDateKey_(r[2]).slice(0, 7) === currentMonth) monthDeposits += amount;
    } else if (r[3] === 'ถอน') {
      balanceWithdrawals += amount;
      if (cellDateKey_(r[2]).slice(0, 7) === currentMonth) monthWithdrawals += amount;
    }
  });
  return '💵 Balance คงเหลือ: ' + formatMoney_(balanceDeposits - balanceWithdrawals) + ' บาท\n\n' +
    '📥 ยอดฝากสะสมเดือนนี้: ' + formatMoney_(monthDeposits) + ' บาท\n📤 ยอดถอนสะสมเดือนนี้: ' + formatMoney_(monthWithdrawals) + ' บาท';
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
  if (!recent.length) return 'ไม่มีรายการเงินเก็บในช่วง 15 วันล่าสุดค่ะ';
  const totalDeposit = recent.reduce((sum, item) => sum + (item.row[3] === 'ฝาก' ? Number(item.row[4]) || 0 : 0), 0);
  const totalWithdrawal = recent.reduce((sum, item) => sum + (item.row[3] === 'ถอน' ? Number(item.row[4]) || 0 : 0), 0);
  return 'รายการเงินเก็บย้อนหลัง 15 วัน (' + Utilities.formatDate(start, CONFIG.TIMEZONE, 'yyyy-MM-dd') +
    ' ถึง ' + Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM-dd') + ')\n\n' +
    recent.map(item => (item.row[3] === 'ฝาก' ? '📥 ฝาก' : '📤 ถอน') + ' | ' + item.dateKey + ' | ' +
      formatMoney_(Number(item.row[4]) || 0) + ' บาท | 📝 ' + item.row[5]).join('\n\n') +
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
  let matchedChats = 0;
  let sentChats = 0;
  const failures = [];
  chats.filter(row => row[1] === true || String(row[1]).toLowerCase() === 'true').forEach(row => {
    const ownerKey = String(row[0]);
    const items = tasks.filter(t => String(t[1]) === ownerKey && String(t[5]).toUpperCase() === 'OPEN' && reminderDates_(t).includes(today))
      .sort((a, b) => formatTaskTime_(a[4]).localeCompare(formatTaskTime_(b[4])));
    if (items.length) {
      const recipient = String(row[4] || row[0]);
      matchedChats++;
      try {
        pushLine_(recipient, '🔔 แจ้งเตือนนัดหมาย\n\n' + items.map(formatReminderTask_).join('\n\n'));
        sentChats++;
      } catch (err) {
        failures.push(ownerKey + ': ' + err.message);
        console.error('Daily reminder failed for ' + ownerKey + ': ' + err.message);
      }
    }
  });
  console.log('Daily reminders for ' + today + ': matched chats=' + matchedChats + ', sent chats=' + sentChats + ', opted-in chats=' + chats.filter(row => row[1] === true || String(row[1]).toLowerCase() === 'true').length);
  if (failures.length) throw new Error('Some LINE reminders failed: ' + failures.join(' | '));
}

function reminderDate_(taskRow) {
  return reminderDates_(taskRow)[0] || '';
}

function reminderDates_(taskRow) {
  const mode = String(taskRow[8] || 'ON_DATE');
  const scheduledDate = cellDateKey_(taskRow[3]);
  if (!scheduledDate) return [];
  if (mode === 'ON_DATE') return [scheduledDate];
  const dates = [];
  const date = new Date(scheduledDate + 'T00:00:00+07:00');
  date.setDate(date.getDate() - 1);
  dates.push(Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM-dd'));
  if (mode === 'BOTH') dates.push(scheduledDate);
  return dates;
}

function replyLine_(replyToken, message) {
  if (!replyToken) return;
  const lineMessage = {type: 'text', text: message};
  if (message === appointmentDatePrompt_()) {
    lineMessage.quickReply = appointmentDateTimeQuickReply_();
  } else if (message === appointmentSummaryDatePrompt_()) {
    lineMessage.quickReply = datePickerQuickReply_('selectSummaryDate');
  } else if (String(message).indexOf('📅 วันที่') === 0) {
    lineMessage.quickReply = transactionDateQuickReply_();
  }
  lineRequest_('https://api.line.me/v2/bot/message/reply', {
    replyToken: replyToken,
    messages: [lineMessage]
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
    throw new Error('LINE API error ' + response.getResponseCode() + ': ' + response.getContentText());
  }
}
