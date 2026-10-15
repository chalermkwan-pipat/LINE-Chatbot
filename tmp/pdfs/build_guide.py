from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.pdfgen.canvas import Canvas
from xml.sax.saxutils import escape
from pathlib import Path

ROOT = Path(r'C:\Users\chalermkwan.p\Documents\Codex\2026-10-02\new-chat')
OUT = ROOT / 'output' / 'pdf' / 'line_chatbot_google_sheets_setup_th.pdf'
FONT = r'C:\Windows\Fonts\LeelawUI.ttf'
FONT_B = r'C:\Windows\Fonts\Leelawdb.ttf'
pdfmetrics.registerFont(TTFont('LeelawUI', FONT))
pdfmetrics.registerFont(TTFont('LeelawUI-Bold', FONT_B))
pdfmetrics.registerFontFamily('LeelawUI', normal='LeelawUI', bold='LeelawUI-Bold', italic='LeelawUI', boldItalic='LeelawUI-Bold')

NAVY = colors.HexColor('#12304A')
TEAL = colors.HexColor('#0D766E')
PALE = colors.HexColor('#EAF4F4')
LIGHT = colors.HexColor('#F3F6F8')
INK = colors.HexColor('#203040')
MUTED = colors.HexColor('#5C6B77')
ORANGE = colors.HexColor('#FFF2DA')

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='TitleTH', fontName='LeelawUI-Bold', fontSize=25, leading=32, textColor=NAVY, alignment=TA_LEFT, spaceAfter=7*mm))
styles.add(ParagraphStyle(name='SubTH', fontName='LeelawUI', fontSize=12, leading=18, textColor=MUTED, spaceAfter=5*mm))
styles.add(ParagraphStyle(name='H1TH', fontName='LeelawUI-Bold', fontSize=17, leading=22, textColor=NAVY, spaceBefore=2*mm, spaceAfter=3*mm, keepWithNext=True))
styles.add(ParagraphStyle(name='H2TH', fontName='LeelawUI-Bold', fontSize=12, leading=16, textColor=TEAL, spaceBefore=2*mm, spaceAfter=1.5*mm, keepWithNext=True))
styles.add(ParagraphStyle(name='BodyTH', fontName='LeelawUI', fontSize=9.5, leading=14, textColor=INK, spaceAfter=2*mm, wordWrap='CJK'))
styles.add(ParagraphStyle(name='SmallTH', fontName='LeelawUI', fontSize=8, leading=11, textColor=MUTED, spaceAfter=1*mm, wordWrap='CJK'))
styles.add(ParagraphStyle(name='CellTH', fontName='LeelawUI', fontSize=8.5, leading=12, textColor=INK, wordWrap='CJK'))
styles.add(ParagraphStyle(name='CellHeadTH', fontName='LeelawUI-Bold', fontSize=8.5, leading=12, textColor=colors.white, wordWrap='CJK'))
styles.add(ParagraphStyle(name='CalloutTH', fontName='LeelawUI', fontSize=9, leading=13, textColor=NAVY, wordWrap='CJK'))
styles.add(ParagraphStyle(name='CodeTH', fontName='LeelawUI', fontSize=8.5, leading=12, textColor=NAVY, backColor=LIGHT, borderColor=colors.HexColor('#D6E0E5'), borderWidth=0.5, borderPadding=5, spaceBefore=1*mm, spaceAfter=3*mm, wordWrap='CJK'))
styles.add(ParagraphStyle(name='CenterTH', fontName='LeelawUI-Bold', fontSize=10, leading=14, textColor=NAVY, alignment=TA_CENTER, wordWrap='CJK'))


def P(text, style='BodyTH'):
    return Paragraph(text, styles[style])

def bullet(text):
    return P('- ' + text)

def numbered(n, title, text):
    return KeepTogether([P(f'<b>{n}. {title}</b>', 'H2TH'), P(text)])

def callout(text, bgcolor=PALE):
    t = Table([[P(text, 'CalloutTH')]], colWidths=[170*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),bgcolor),('BOX',(0,0),(-1,-1),0.5,colors.HexColor('#D5E2E6')),
        ('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),
        ('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7),
    ]))
    return t

def footer(canvas, doc):
    canvas.saveState()
    w,h=A4
    canvas.setStrokeColor(colors.HexColor('#D8E1E6'))
    canvas.setLineWidth(0.5)
    canvas.line(20*mm, 15*mm, w-20*mm, 15*mm)
    canvas.setFont('LeelawUI', 8)
    canvas.setFillColor(MUTED)
    canvas.drawString(20*mm, 9*mm, 'LINE Chat Bot + Google Sheets | คู่มือติดตั้งและใช้งาน')
    canvas.drawRightString(w-20*mm, 9*mm, f'{doc.page}')
    canvas.restoreState()


doc = SimpleDocTemplate(str(OUT), pagesize=A4, rightMargin=20*mm, leftMargin=20*mm, topMargin=17*mm, bottomMargin=21*mm,
                        title='คู่มือติดตั้ง LINE Chat Bot เชื่อม Google Sheets', author='OpenAI')
story=[]

# Page 1 - overview and current flow
story += [P('LINE Chat Bot<br/>เชื่อม Google Sheets', 'TitleTH'),
          P('คู่มือติดตั้ง ตั้งค่า และใช้งาน | อ้างอิง LineReminderBot.gs เวอร์ชันล่าสุดในโปรเจกต์', 'SubTH'),
          callout('<b>ภาพรวม</b><br/>LINE ส่งข้อความเข้า Webhook &gt; Google Apps Script ประมวลผล &gt; อ่าน/เขียนข้อมูลใน Google Sheets &gt; ตอบกลับหรือส่งแจ้งเตือนใน LINE'),
          Spacer(1,5*mm), P('1. เมนูหลัก', 'H1TH')]
menu_data = [
    [P('รหัส', 'CellHeadTH'), P('งาน', 'CellHeadTH')],
    [P('C1 / C2 / C3','CellTH'), P('จัดการนัดหมาย / บันทึกนัด / สรุปงานตามวัน','CellTH')],
    [P('M1-M4','CellTH'), P('เงินเก็บ / ฝาก / ถอน / ดูยอดคงเหลือ','CellTH')],
    [P('L1 / L2','CellTH'), P('จัดการเงินผ่อน / ดูยอดคงค้าง','CellTH')],
    [P('A1','CellTH'), P('กำหนดสิทธิ์ (admin ในแชตส่วนตัวเท่านั้น)','CellTH')],
    [P('X','CellTH'), P('ออกจากบอท (หรือพิมพ์ bye bot)','CellTH')],
]
t=Table(menu_data,colWidths=[34*mm,136*mm],repeatRows=1)
t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),NAVY),('GRID',(0,0),(-1,-1),0.35,colors.HexColor('#D5DEE4')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,LIGHT])]))
story += [t, Spacer(1,4*mm), P('2. Flow หลัก', 'H1TH'),
          P('<b>นัดหมาย:</b> C2 &gt; กรอกรายละเอียด | สถานที่ &gt; เลือกวันและเวลา หรือเลือกเฉพาะวัน &gt; เลือกแจ้งเตือนล่วงหน้า 1 วัน / วันนัด / ทั้งสองวัน &gt; บันทึกลงชีต Tasks'),
          P('<b>เงินเก็บ:</b> M1 &gt; ฝากหรือถอน &gt; กรอกจำนวน | หมายเหตุ &gt; เลือกวันที่หรือ “วันนี้” &gt; ดู balance หรือย้อนหลัง 15 วัน'),
          P('<b>เงินผ่อน:</b> L1 &gt; ตั้งสัญญา / จ่ายเงินต้น / จ่ายดอกเบี้ย / ตรวจยอดและประวัติ &gt; เลือกสัญญาก่อนจ่ายเมื่อมีหลายสัญญา'),
          P('<b>คำสั่งช่วยเหลือ:</b> /help ดูเมนูย่อ | /cancel ยกเลิกขั้นตอน | X หรือ bye bot ออกจากบอท'),
          PageBreak()]

# Page 2 - Google Sheets and Apps Script
story += [P('ติดตั้งระบบ', 'H1TH'), P('เตรียม Google Sheets และ Apps Script', 'SubTH'),
          numbered(1,'สร้าง Google Sheet','สร้าง Google Spreadsheet สำหรับเก็บข้อมูล (ถ้ามีชีตเดิมใช้ชีตนั้นได้) แล้วเปิด <b>Extensions &gt; Apps Script</b> เพื่อสร้าง Apps Script ที่ผูกกับชีตนี้'),
          numbered(2,'วางสคริปต์','เปิดไฟล์ <b>outputs/LineReminderBot.gs</b> ในชุดไฟล์นี้ คัดลอกโค้ดทั้งหมดไปวางใน Apps Script (เช่น Code.gs) แล้วกด Save'),
          numbered(3,'ตั้ง Script Properties','ใน Apps Script เปิด <b>Project Settings &gt; Script properties</b> แล้วเพิ่มค่าต่อไปนี้:'),]
props = [
    [P('Property','CellHeadTH'),P('ค่า','CellHeadTH'),P('ใช้ทำอะไร','CellHeadTH')],
    [P('LINE_CHANNEL_ACCESS_TOKEN','CellTH'),P('Channel access token จาก LINE Developers','CellTH'),P('ให้บอทตอบและส่งข้อความ LINE','CellTH')],
    [P('WEBHOOK_KEY','CellTH'),P('ข้อความลับที่สร้างเองแบบสุ่ม','CellTH'),P('ตรวจ key ใน URL ของ Webhook','CellTH')],
    [P('BOT_ADMIN_USER_IDS','CellTH'),P('LINE userId ของ admin คั่นด้วย comma','CellTH'),P('แสดงและอนุญาต A1 ในแชตส่วนตัว','CellTH')],
]
t=Table(props,colWidths=[53*mm,68*mm,49*mm],repeatRows=1)
t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),TEAL),('GRID',(0,0),(-1,-1),0.35,colors.HexColor('#D5DEE4')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,LIGHT])]))
story += [t, Spacer(1,3*mm),
          P('<b>หา LINE userId ของ admin:</b> ชั่วคราวเพิ่มบรรทัดนี้ใน handleEvent_(event) หลังบรรทัด getTarget_ แล้วส่งข้อความส่วนตัวถึงบอท จากนั้นดูค่าใน Apps Script &gt; Executions &gt; Logs นำ userId ไปใส่ BOT_ADMIN_USER_IDS คั่นหลายคนด้วย comma แล้วลบบรรทัด log ออกค่ะ'),
          P('console.log("LINE_USER_ID=" + event.source.userId);','CodeTH'),
          P('<b>WEBHOOK_KEY</b> ตั้งเป็นค่าสุ่มยาวและเก็บเป็นความลับ ห้ามใส่ token หรือ key จริงในเอกสาร/แชตสาธารณะค่ะ'),
          numbered(4,'เตรียม LINE Official Account และ Messaging API','ใน LINE Official Account Manager เปิดหรือสร้างบัญชีสำหรับบอท จากนั้นเชื่อม Messaging API channel ใน LINE Developers Console เตรียม Channel access token และคัดลอกมาใส่ Script Properties'),
          numbered(5,'ตั้งค่า LINE channel','ในแท็บ Messaging API เปิด <b>Use webhook</b> และเปิด <b>Allow bot to join group chats</b> หากจะใช้ในกลุ่ม ปิด Greeting messages/Auto-reply ที่ซ้ำกับคำตอบของสคริปต์'),
          PageBreak()]

# Page 3 - deploy and triggers
story += [P('Deploy และเชื่อม Webhook', 'H1TH'),
          numbered(6,'รัน setup()','กลับ Apps Script เลือกฟังก์ชัน <b>setup</b> แล้วกด Run ครั้งแรก จากนั้นอนุญาตสิทธิ์ Google Sheets และ Triggers ตามหน้าจอ ฟังก์ชันจะสร้างชีต Tasks, Savings, Installments, Users, Permissions และตั้ง Trigger ให้อัตโนมัติ'),
          callout('<b>Trigger ที่ setup สร้าง</b><br/>sendDailyReminders และ accrueAnnualInterest วันละครั้ง เวลา 06:00 น. ตาม Asia/Bangkok; expireIdleSessions ทุก 1 นาที การรัน setup ซ้ำจะลบและสร้าง Trigger ของสามฟังก์ชันนี้ใหม่ แต่ไม่ลบข้อมูลในชีต'),
          Spacer(1,4*mm),
          numbered(7,'Deploy เป็น Web app','กด <b>Deploy &gt; New deployment &gt; Web app</b><br/>- Execute as: <b>Me</b><br/>- Who has access: <b>Anyone</b> (หรือแบบ anonymous หากหน้าจอแสดงตัวเลือก)<br/>กด Deploy และอนุญาตสิทธิ์ จากนั้นคัดลอก URL ที่ลงท้ายด้วย <b>/exec</b>'),
          numbered(8,'สร้าง Webhook URL','นำ URL /exec มาต่อท้ายด้วย <b>?key=</b> และค่า WEBHOOK_KEY ที่ตั้งไว้ เช่น:'),
          P('https://script.google.com/macros/s/DEPLOYMENT_ID/exec?key=YOUR_WEBHOOK_KEY','CodeTH'),
          numbered(9,'ใส่ URL ใน LINE Developers','Messaging API channel &gt; Webhook settings &gt; วาง URL เต็ม &gt; Verify &gt; เมื่อสำเร็จ เปิด Use webhook'),
          numbered(10,'ทดสอบแชตส่วนตัว','เพิ่ม LINE Official Account เป็นเพื่อน แล้วพิมพ์ <b>hi bot</b> ควรเห็นเมนูหลัก ลอง /help, C2 หรือ M4'),
          numbered(11,'ทดสอบในกลุ่ม','เชิญ Official Account เข้ากลุ่ม แล้วพิมพ์ <b>hi bot</b> ในกลุ่มเพื่อให้บอทลงทะเบียนกลุ่ม ข้อมูลของแต่ละกลุ่มแยก scope กัน'),
          callout('<b>Deploy เวอร์ชันใหม่:</b> หลังแก้โค้ด เลือก Deploy &gt; Manage deployments &gt; Edit &gt; New version &gt; Deploy หากสร้าง deployment ใหม่แทน ต้องอัปเดต Webhook URL ใน LINE ด้วยค่ะ', ORANGE),
          PageBreak()]

# Page 4 - usage detail
story += [P('คู่มือใช้งาน', 'H1TH'), P('นัดหมาย', 'H2TH'),
          P('1) พิมพ์ C2 (หรือเลือกบันทึกจาก C1) &gt; 2) กรอก <b>รายละเอียด | สถานที่</b> &gt; 3) เลือกวันและเวลาจากปฏิทิน หรือเลือกวันอย่างเดียว &gt; 4) เลือกแจ้งเตือน 1 ล่วงหน้า 1 วัน, 2 วันนัด, 3 ทั้งสองวัน'),
          P('ผลสรุปใช้ชื่อวันและวันที่ เช่น <b>วันเสาร์ที่ 03/10/26</b> ถ้ามีเวลา ระบบแสดงเวลาหน้ารายละเอียดค่ะ'),
          P('C1 จัดการนัดหมาย: 1 บันทึก | 2 สรุปตามวัน | 3 ลบงาน', 'CodeTH'),
          P('เงินเก็บ', 'H2TH'),
          P('M1 เปิดเมนูเงินเก็บ: 1 ฝาก | 2 ถอน | 3 balance | 4 รายการย้อนหลัง 15 วัน. กรอกจำนวนและหมายเหตุ เช่น <b>500 | เก็บเงินเดือน</b> แล้วเลือกวันที่ หรือกด “วันนี้”'),
          P('M2 ฝากโดยตรง | M3 ถอนโดยตรง | M4 ดูยอดคงเหลือ', 'CodeTH'),
          P('เงินผ่อน', 'H2TH'),
          P('L1 เปิดเมนูเงินผ่อน | L2 ดูยอดคงค้าง รายการตั้งสัญญารับหัวเรื่อง เจ้าหนี้ ลูกหนี้ เงินต้น อัตราดอกเบี้ย หมายเหตุ และฐานดอกเบี้ย (ยอดตั้งต้น/ยอดคงเหลือ) จากนั้นเลือกวันที่'),
          P('เมื่อมีหลายสัญญา บอทให้เลือกสัญญาก่อนจ่ายเงินต้น/ดอกเบี้ย หรือคำนวณดอกเบี้ยจากยอดล่าสุด ดอกเบี้ยรายปีอัตโนมัติเมื่อครบปี; เมนู 7 ใช้คำนวณจากยอดคงเหลือล่าสุดตามคำสั่ง'),
          P('เมนูเงินผ่อน: 1 ตั้งยอด | 2 จ่ายเงินต้น | 3 จ่ายดอกเบี้ย | 4 ดูยอด | 5 ประวัติ 16 เดือน | 6 ดอกเบี้ยค้าง | 7 คำนวณจากยอดล่าสุด', 'CodeTH'),
          P('สิทธิ์ผู้ใช้', 'H2TH'),
          P('A1 ใช้ได้เฉพาะ admin ในแชตส่วนตัว: เลือกกลุ่ม &gt; เลือกกำหนดสิทธิ์ทั้งกลุ่มหรือสมาชิก &gt; ตอบ Y/N ตามลำดับ นัดหมาย, เงิน, เงินผ่อน (เช่น Y,N,Y). ในกลุ่ม admin เห็นเมนูตามสิทธิ์กลุ่มเช่นเดียวกับสมาชิก'),
          P('ออกจากระบบ: พิมพ์ X หรือ bye bot. หากไม่ส่งข้อความเกิน 1 นาที ระบบปิด session อัตโนมัติ', 'BodyTH'),
          PageBreak()]

# Page 5 - troubleshooting + references
story += [P('ตรวจสอบเมื่อใช้งานไม่ได้', 'H1TH')]
troubles = [
    ('บอทไม่ตอบข้อความ','ตรวจว่าได้พิมพ์ hi bot และ session ยังไม่หมดอายุ; Apps Script &gt; Executions ดู error; ตรวจ webhook URL ว่ามี ?key= ตรงกับ WEBHOOK_KEY และใช้ /exec'),
    ('LINE Verify ไม่ผ่าน','ตรวจ deployment access เป็น Anyone/anonymous, URL เป็น deployment ปัจจุบัน และ query key ถูกต้อง จากนั้นกด Verify อีกครั้ง'),
    ('ส่งข้อความออกไม่ได้','ตรวจ LINE_CHANNEL_ACCESS_TOKEN ว่ามาจาก Messaging API channel ที่ถูกต้องและยังใช้ได้; ตรวจ execution error จาก LINE API'),
    ('Trigger มี Last run แต่ไม่เห็นข้อความ','ตรวจ Tasks ว่าสถานะ OPEN, วันแจ้งเตือนตรงกับวันนี้/วันก่อนนัดตามตัวเลือก และ Users ของแชตนั้นเปิดรับแจ้งเตือน; ตรวจ Executions ของ sendDailyReminders'),
    ('ไม่เห็น A1','ตรวจว่าเป็น admin ใน BOT_ADMIN_USER_IDS และกำลังคุยกับบอทใน personal chat; A1 ไม่แสดงในกลุ่ม'),
    ('เมนูในกลุ่มไม่ครบ','ตรวจสิทธิ์กลุ่มผ่าน A1 ใน personal chat; admin ในกลุ่มใช้สิทธิ์ตามกลุ่ม ไม่ข้ามการตั้งค่า'),
]
data=[[P('อาการ','CellHeadTH'),P('จุดตรวจ','CellHeadTH')]]+[[P(a,'CellTH'),P(b,'CellTH')] for a,b in troubles]
t=Table(data,colWidths=[44*mm,126*mm],repeatRows=1)
t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),NAVY),('GRID',(0,0),(-1,-1),0.35,colors.HexColor('#D5DEE4')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,LIGHT])]))
story += [t, Spacer(1,4*mm), callout('<b>ข้อควรทราบด้านความปลอดภัย</b><br/>สคริปต์รุ่นนี้ตรวจ key จากพารามิเตอร์ URL (?key=...) ให้เก็บ URL นี้เป็นความลับและจำกัดผู้ที่เข้าถึงโปรเจกต์/ชีต เอกสาร LINE แนะนำให้ตรวจ X-Line-Signature ด้วย หากนำระบบไปใช้งานจริงในวงกว้าง ควรเพิ่มการตรวจลายเซ็นตามคู่มือ LINE', ORANGE),
          Spacer(1,4*mm), P('เอกสารอ้างอิงทางการ', 'H2TH'),
          P('<link href="https://developers.google.com/apps-script/guides/web" color="#0D766E">Google Apps Script: Deploy web apps</link>','SmallTH'),
          P('<link href="https://developers.line.biz/en/docs/messaging-api/building-bot/" color="#0D766E">LINE Developers: Build a bot</link>','SmallTH'),
          P('<link href="https://developers.line.biz/en/docs/messaging-api/receiving-messages/" color="#0D766E">LINE Developers: Receive messages (webhook)</link>','SmallTH'),
          P('<link href="https://developers.line.biz/en/docs/messaging-api/group-chats/" color="#0D766E">LINE Developers: Group chats</link>','SmallTH'),
          Spacer(1,2*mm), P('คู่มือนี้อ้างอิงสคริปต์ในโปรเจกต์ ณ 3 ตุลาคม 2026 การตั้งค่าบางรายการในหน้า Google/LINE อาจเปลี่ยนชื่อได้ตามเวอร์ชันหน้าจอ', 'SmallTH')]

doc.build(story, onFirstPage=footer, onLaterPages=footer)
print(OUT)
