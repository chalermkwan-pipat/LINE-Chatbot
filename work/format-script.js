const fs = require('fs');
const p = 'C:/Users/chalermkwan.p/Documents/Codex/2026-10-02/new-chat/outputs/LineReminderBot.gs';
let s = fs.readFileSync(p, 'utf8');
const pairs = [
[`return '💰 ฝากเงิน — วันที่ปัจจุบันจะบันทึกให้อัตโนมัติ\\nกรอก: จำนวนเงิน | หมายเหตุ\\nตัวอย่าง: 500 | เก็บเงินเดือน\\n0 — ออกจากการกรอก';`,`return '💰 ฝากเงิน\\nวันที่ปัจจุบันจะบันทึกให้อัตโนมัติครับ\\n\\nกรอก: จำนวนเงิน | หมายเหตุ\\n\\nตัวอย่าง: 500 | เก็บเงินเดือน\\n\\n0 — ออกจากการกรอก';`],
[`return '💰 ถอนเงิน — วันที่ปัจจุบันจะบันทึกให้อัตโนมัติ\\nกรอก: จำนวนเงิน | หมายเหตุ\\nตัวอย่าง: 200 | ใช้จ่าย\\n0 — ออกจากการกรอก';`,`return '💰 ถอนเงิน\\nวันที่ปัจจุบันจะบันทึกให้อัตโนมัติครับ\\n\\nกรอก: จำนวนเงิน | หมายเหตุ\\n\\nตัวอย่าง: 200 | ใช้จ่าย\\n\\n0 — ออกจากการกรอก';`],
[`return '✅ บันทึกนัดหมายแล้วครับ (รหัส ' + id + ')\\n📅 วันที่: ' + date +\n    (time ? '\\n🕒 เวลา: ' + time : '') + '\\n📝 รายละเอียด: ' + description +\n    (location ? '\\n📍 สถานที่: ' + location : '');`,`return '✅ บันทึกนัดหมายแล้วครับ (รหัส ' + id + ')\\n\\n📅 วันที่: ' + date +\n    (time ? '\\n\\n🕒 เวลา: ' + time : '') + '\\n\\n📝 รายละเอียด: ' + description +\n    (location ? '\\n\\n📍 สถานที่: ' + location : '');`],
[`items.map(formatTask_).join('\\n')`,`items.map(formatTask_).join('\\n\\n')`],
[`recent.map(item => (item.row[3] === 'ฝาก' ? '📥 ฝาก' : '📤 ถอน') + ' | ' + item.dateKey + ' | ' +\n      formatMoney_(Number(item.row[4]) || 0) + ' บาท | 📝 ' + item.row[5]).join('\\n')`,`recent.map(item => (item.row[3] === 'ฝาก' ? '📥 ฝาก' : '📤 ถอน') + ' | ' + item.dateKey + ' | ' +\n      formatMoney_(Number(item.row[4]) || 0) + ' บาท | 📝 ' + item.row[5]).join('\\n\\n')`],
[`' ถึง ' + Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM-dd') + ')\\n' +`,`' ถึง ' + Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM-dd') + ')\\n\\n' +`],
[`return '💰 ยอดเงินเก็บของแชตนี้\\n💵 Balance คงเหลือ: ' + formatMoney_(deposits - withdrawals) + ' บาท\\n' +\n    '📥 ยอดฝากสะสม: ' + formatMoney_(deposits) + ' บาท\\n📤 ยอดถอนสะสม: ' + formatMoney_(withdrawals) + ' บาท';`,`return '💰 ยอดเงินเก็บของแชตนี้\\n\\n💵 Balance คงเหลือ: ' + formatMoney_(deposits - withdrawals) + ' บาท\\n\\n' +\n    '📥 ยอดฝากสะสม: ' + formatMoney_(deposits) + ' บาท\\n📤 ยอดถอนสะสม: ' + formatMoney_(withdrawals) + ' บาท';`]
];
for (const [oldText, newText] of pairs) {
  if (!s.includes(oldText)) console.log('NOT FOUND:', oldText.slice(0, 70));
  else s = s.replace(oldText, newText);
}
s = s.replace(`return items.length ? '📅 ' + label + ' (' + date + ')\\n' + items.map(formatTask_).join('\\n\\n')`, `return items.length ? '📅 ' + label + ' (' + date + ')\\n\\n' + items.map(formatTask_).join('\\n\\n')`);
s = s.replace(`'📅 สรุปงานวันที่ ' + date + ' (' + items.length + ' งาน)\\n' + items.map(formatTask_).join('\\n\\n')`, `'📅 สรุปงานวันที่ ' + date + ' (' + items.length + ' งาน)\\n\\n' + items.map(formatTask_).join('\\n\\n')`);
s = s.replace(`'📅 แจ้งเตือนนัดหมาย (' + today + ')\\n' + items.map(formatTask_).join('\\n\\n')`, `'📅 แจ้งเตือนนัดหมาย (' + today + ')\\n\\n' + items.map(formatTask_).join('\\n\\n')`);
fs.writeFileSync(p, s, 'utf8');
console.log('Formatting pass complete.');
