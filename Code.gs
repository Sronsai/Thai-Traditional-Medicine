/**
 * PCH-SumYa AI — Google Apps Script Web App
 * โรงพยาบาลปากชม จังหวัดเลย | กลุ่มงานแพทย์แผนไทย & ฝ่ายสารสนเทศ
 *
 * วิธีติดตั้ง
 * 1) เปิด Google Sheet เป้าหมาย → Extensions → Apps Script → วางไฟล์นี้แทน Code.gs
 * 2) แก้ค่า CONFIG ด้านล่าง (SPREADSHEET_ID, DRIVE_FOLDER_ID, API_KEY)
 * 3) รันฟังก์ชัน setupSheets()  (ครั้งเดียว) → สร้างชีต sumary / tnss_daily / patients / lookup พร้อมหัวคอลัมน์
 * 4) Deploy → New deployment → Web app
 *      Execute as: Me
 *      Who has access: Anyone
 *    คัดลอก Web app URL (.../exec) ไปใส่ในหน้า "ตั้งค่า" ของแอป PCH-SumYa AI
 */

const CONFIG = {
  SPREADSHEET_ID: '1FAs-T5YqJPzKB1utX-bGuPhCFQOrE240Z0uKAoa1umg',
  DRIVE_FOLDER_ID: '',          // โฟลเดอร์เก็บไฟล์เสียงหายใจ (เว้นว่าง = สร้างให้อัตโนมัติ)
  API_KEY: 'pakchom-suma-2569',  // ต้องตรงกับค่าในแอป
  TIMEZONE: 'Asia/Bangkok'
};

/* ─────────── หัวคอลัมน์ 45 ช่อง ชีต sumary (1 แถว = 1 ครั้งการสุ่มยา) ─────────── */
const SUMARY_HEADERS = [
  'รหัสบันทึก',                 // A  SUMA-YYMMDD-xxx
  'วันที่-เวลาบันทึก',            // B  auto
  'ยินยอม PDPA',               // C  ยินยอม / ไม่ยินยอม
  'HN',                        // D
  'ชื่อ-นามสกุล',                // E
  'อายุ (ปี)',                  // F  ตัวเลข
  'เพศ',                       // G  ชาย/หญิง
  'ธาตุเจ้าเรือน',               // H  ไฟ/ลม/น้ำ/ดิน (ตามเดือนเกิด)
  'อุตุสมุฏฐาน (ฤดู)',            // I  คิมหันต/วสันต/เหมันตฤดู
  'กาลสมุฏฐาน (ช่วงเวลา)',        // J  ช่วงเวลาที่อาการกำเริบ
  'ตรีธาตุที่กระทำ',              // K  วาตะ/ปิตตะ/เสมหะเด่น
  'วันที่รับบริการ',              // L  yyyy-mm-dd
  'คลินิก/หน่วยบริการ',          // M
  'ผู้ทำหัตถการ (พท.)',          // N
  'รหัส ICD-10',               // O  J30.1 / J30.4 ฯลฯ
  'รหัสหัตถการ',                // P  เช่น 9999510 (สุ่มยาสมุนไพร)
  'ขิง (%)',                    // Q
  'ดอกปีบ (%)',                 // R
  'มะแขวน (%)',                 // S
  'ผิวมะกรูด (%)',              // T
  'ข่า (%)',                    // U  Q+R+S+T+U ต้องรวม = 100
  'น้ำหนักยารวม (กรัม)',         // V
  'อุณหภูมิไอน้ำ (°C)',          // W
  'ระยะเวลาสุ่มยา (นาที)',        // X
  'Patency ก่อน (%)',           // Y  AI Acoustic
  'Patency หลัง (%)',           // Z
  'ผลต่าง Patency (%)',         // AA  Z - Y
  'ลิงก์เสียงก่อน',              // AB  Drive URL
  'ลิงก์เสียงหลัง',              // AC
  'ระยะใกล้สุด (ซม.)',           // AD AR Safety
  'จำนวนแจ้งเตือนระยะ',          // AE
  'ทำครบเวลา (%)',              // AF
  'เหตุการณ์ไม่พึงประสงค์',        // AG ไม่มี/ระคายเคืองตา/ระคายเคืองผิวหนัง/ลวก
  'TNSS ก่อน (จาม/น้ำมูก/คัด/คัน)', // AH "2/3/3/1"
  'TNSS รวมก่อน',               // AI 0-12
  'TNSS หลัง (จาม/น้ำมูก/คัด/คัน)', // AJ (สคริปต์เขียนค่า)
  'TNSS รวมหลัง',               // AK (สคริปต์เขียนค่า)
  'TNSS ลดลง (%)',              // AL (สคริปต์เขียนค่า)
  'กลุ่มอาการเด่น',              // AM คัดจมูกเด่น/น้ำมูกเด่น/จามเด่น/คันจมูกเด่น
  'คำแนะนำปรับสูตรจาก AI',       // AN
  'สูตรแนะนำครั้งถัดไป',          // AO "ขิง35/มะแขวน30/ข่า20/มะกรูด10/ปีบ5"
  'คะแนน SUS (พท.)',            // AP 0-100
  'ความพึงพอใจผู้ป่วย (1-5)',     // AQ
  'ต้นทุนรวม/ครั้ง (บาท)',        // AR
  'อุปกรณ์',                    // AS
  'เวอร์ชันแอป',                // AT
  'หมายเหตุ',                   // AU
  'แปลผล TNSS ก่อน',             // AV  แอปส่งมา / สคริปต์เติม
  'แปลผล TNSS หลัง',             // AW  refreshTnssFormulas() เติม
  'RCQ-36 รวม (ก่อน)',           // AX  แอปส่งมา / linkRcqToSumary() เติม
  'RCQ-36 แปลผล (ก่อน)'          // AY
];

const TNSS_HEADERS = ['รหัสบันทึก', 'วันที่-เวลาบันทึก', 'HN', 'วันที่ประเมิน', 'วันที่ (ลำดับ)', 'จาม', 'น้ำมูกไหล', 'คัดจมูก', 'คันจมูก', 'TNSS รวม', 'สุ่มยาวันนี้', 'หมายเหตุ'];
const PATIENT_HEADERS = ['HN', 'วันที่ลงทะเบียน', 'ชื่อ-นามสกุล', 'อายุ (ปี)', 'เพศ', 'ธาตุเจ้าเรือน', 'เบอร์โทร/LINE', 'โรคร่วม/ข้อห้าม', 'ยินยอม PDPA', 'คลินิก/หน่วยบริการ'];
const LOOKUP_HEADERS = ['ประเภท', 'ค่า', 'รายละเอียด'];

/* ───── RCQ-36 · แบบสอบถามคุณภาพชีวิตเฉพาะโรคจมูกและตาอักเสบจากภูมิแพ้ ───── */
const RCQ_DOMAINS = [
  { id: 'd1', no: 1, title: 'อาการทางจมูก', items: ['น้ำมูกไหล', 'คันจมูก', 'คัดแน่นจมูก', 'จาม', 'ไอ', 'คอแห้งปากแห้ง', 'มีเสมหะ'] },
  { id: 'd2', no: 2, title: 'อาการทางตาและร่างกาย', items: ['คันตา', 'เคืองตา', 'น้ำตาไหล', 'ไม่สบายตา', 'สมองไม่โล่ง', 'อ่อนเพลีย', 'เหนื่อยง่าย', 'ปวดเมื่อยตามตัว', 'ปวดศีรษะ', 'ง่วงนอนตลอดเวลา'] },
  { id: 'd3', no: 3, title: 'ปัญหาในการทำงานหรือการเรียน', items: ['ต้องหยุดงาน หรือ หยุดเรียนหนังสือ', 'ไม่มีสมาธิในการทำงาน หรือ เรียนหนังสือ', 'เป็นอุปสรรคต่อการทำงาน'] },
  { id: 'd4', no: 4, title: 'ปัญหาในการทำกิจกรรม', items: ['ออกกำลังกายหนัก', 'ออกกำลังกายปานกลาง', 'เดินเป็นระยะทางครึ่งกิโลเมตร'] },
  { id: 'd5', no: 5, title: 'ปัญหาการนอนหลับ', items: ['ต้องตื่นกลางดึกบ่อย ๆ', 'นอนหลับยาก', 'นอนหลับไม่สนิท'] },
  { id: 'd6', no: 6, title: 'ปัญหาการเข้าสังคม', items: ['สูญเสียความมั่นใจในการพบปะคนอื่น', 'พบปะสังสรรค์น้อยลง', 'ไม่อยากออกไปไหน'] },
  { id: 'd7', no: 7, title: 'ด้านอารมณ์', items: ['รู้สึกรำคาญตนเอง', 'กังวลใจ', 'หงุดหงิด', 'ไม่แจ่มใส เบิกบาน', 'รำคาญที่ต้องพกกระดาษชำระมากกว่าปกติ'] }
];

function rcqHeaders_() {
  const h = ['รหัสบันทึก', 'วันที่-เวลาบันทึก', 'HN', 'ชื่อ-นามสกุล', 'ครั้งที่ประเมิน', 'รหัสบันทึก sumary'];
  RCQ_DOMAINS.forEach(d => d.items.forEach((it, i) => h.push(d.id + '_' + (i + 1) + ' ' + it)));
  h.push('สุขภาพโดยรวม', 'วันหยุดงาน/เรียน (วัน/เดือน)', 'คะแนนรวม (34-170)', 'เฉลี่ยต่อข้อ');
  RCQ_DOMAINS.forEach(d => h.push('คะแนนด้าน ' + d.no + ' ' + d.title));
  h.push('แปลผล RCQ-36', 'ดีขึ้น (%)', 'แปลผลการเปลี่ยนแปลง', 'หมายเหตุ');
  return h;
}
const RCQ_HEADERS = rcqHeaders_();

/** แปลผลค่าเฉลี่ยต่อข้อ (1-5) */
function rcqBand_(avg) {
  const n = Number(avg);
  if (!isFinite(n) || avg === '' || avg === null) return '';
  if (n <= 1.5) return 'กระทบคุณภาพชีวิตน้อยมาก';
  if (n <= 2.5) return 'กระทบคุณภาพชีวิตน้อย';
  if (n <= 3.5) return 'กระทบคุณภาพชีวิตปานกลาง';
  return 'กระทบคุณภาพชีวิตมาก';
}

const LOOKUP_SEED = [
  ['สูตรยามาตรฐาน', 'ขิง 30 / มะแขวน 25 / ข่า 20 / ผิวมะกรูด 15 / ดอกปีบ 10', 'สูตรตั้งต้น เน้นรสเผ็ดร้อนเปิดทางเดินหายใจ'],
  ['สูตรทางเลือก', 'ขิง 20 / มะแขวน 20 / ข่า 20 / ผิวมะกรูด 20 / ดอกปีบ 20', 'สัดส่วนเท่ากันทุกชนิด'],
  ['คลินิก/หน่วยบริการ', 'คลินิกแพทย์แผนไทย รพ.ปากชม', ''],
  ['ผู้ทำหัตถการ (พท.)', 'พท.ศุภกร หีบทอง', ''],
  ['ผู้ทำหัตถการ (พท.)', 'พท.วิไลวรรณ ไขกะวิง', ''],
  ['ธาตุเจ้าเรือน', 'ดิน|น้ำ|ลม|ไฟ', ''],
  ['รหัส ICD-10', 'J30.1|J30.2|J30.3|J30.4', 'Allergic rhinitis'],
  ['รหัสหัตถการ', '9999510', 'หัตถการสุ่มยาสมุนไพร'],
  ['อุณหภูมิไอน้ำ (°C)', '60', 'ค่าตั้งต้นที่แนะนำ'],
  ['ต้นทุนรวม/ครั้ง (บาท)', '35', 'ค่ายาสมุนไพร 5 ชนิด + วัสดุ'],
  ['ธาตุเจ้าเรือน', 'ไฟ', 'เกิดเดือน พฤษภาคม - กรกฎาคม'],
  ['ธาตุเจ้าเรือน', 'ลม', 'เกิดเดือน สิงหาคม - ตุลาคม'],
  ['ธาตุเจ้าเรือน', 'น้ำ', 'เกิดเดือน พฤศจิกายน - มกราคม'],
  ['ธาตุเจ้าเรือน', 'ดิน', 'เกิดเดือน กุมภาพันธ์ - เมษายน'],
  ['อุตุสมุฏฐาน', 'คิมหันตฤดู (ปิตตะสมุฏฐาน)', 'ฤดูร้อน ก.พ.–พ.ค.'],
  ['อุตุสมุฏฐาน', 'วสันตฤดู (วาตะสมุฏฐาน)', 'ฤดูฝน มิ.ย.–ก.ย.'],
  ['อุตุสมุฏฐาน', 'เหมันตฤดู (เสมหะสมุฏฐาน)', 'ฤดูหนาว ต.ค.–ม.ค.'],
  ['กาลสมุฏฐาน', '06.00-10.00 / 18.00-22.00', 'เสมหะกระทำ'],
  ['กาลสมุฏฐาน', '10.00-14.00 / 22.00-02.00', 'ปิตตะกระทำ'],
  ['กาลสมุฏฐาน', '14.00-18.00 / 02.00-06.00', 'วาตะกระทำ'],
  ['ตรีธาตุที่กระทำ', 'วาตะเด่น', 'แห้ง คัน จาม ลมพัดขึ้นสูง → เพิ่มขิง/ข่า'],
  ['ตรีธาตุที่กระทำ', 'ปิตตะเด่น', 'แสบ ร้อน บวม แดง ความร้อนสะสม → เพิ่มดอกปีบ/ผิวมะกรูด ลดขิง'],
  ['ตรีธาตุที่กระทำ', 'เสมหะเด่น', 'น้ำมูกมาก เหนียวข้น คัดจมูก → เพิ่มมะแขวน/ขิง']
];

/* ───────────────────────── Setup ───────────────────────── */
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  ensureSheet_(ss, 'sumary', SUMARY_HEADERS);
  ensureSheet_(ss, 'tnss_daily', TNSS_HEADERS);
  ensureSheet_(ss, 'patients', PATIENT_HEADERS);
  ensureSheet_(ss, 'rcq36', RCQ_HEADERS);
  const lk = ensureSheet_(ss, 'lookup', LOOKUP_HEADERS);
  if (lk.getLastRow() < 2) lk.getRange(2, 1, LOOKUP_SEED.length, 3).setValues(LOOKUP_SEED);
  clearValidation_(ss.getSheetByName('sumary'));
  return 'สร้างชีตเรียบร้อย (ไม่ใช้ dropdown ในชีต — ควบคุมค่าจากแอป)';
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  const cur = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  if (cur.join('|') !== headers.join('|')) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#123b2e').setFontColor('#f6f1e6')
    .setVerticalAlignment('middle').setWrap(true);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 46);
  return sh;
}

/** ล้าง Data Validation ทั้งหมดในชีต — ค่าที่ถูกต้องถูกบังคับจากฝั่งแอปแล้ว */
function clearValidation_(sh) {
  if (!sh || sh.getMaxRows() < 2) return;
  sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns()).clearDataValidations();
}

/** ล้าง dropdown ในทุกชีตด้วยมือ (รันเองได้) */
function clearAllValidations() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ['sumary', 'tnss_daily', 'patients', 'lookup'].forEach(n => clearValidation_(ss.getSheetByName(n)));
  return 'ล้าง dropdown ในชีตทั้งหมดแล้ว';
}

/* ───────────────────────── เมนูในชีต ───────────────────────── */
/** สร้างเมนู "PCH-SumYa AI" บนแถบเมนูของชีต (ทำงานเองทุกครั้งที่เปิดชีต) */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('PCH-SumYa AI')
    .addItem('อัพเดตผล TNSS หลัง (AJ:AL)', 'menuRefresh_')
    .addItem('เชื่อมคะแนน RCQ-36 (AX:AY)', 'menuLinkRcq_')
    .addItem('ตั้งค่าหัวคอลัมน์ทุกชีต', 'menuSetup_')
    .addItem('ล้าง dropdown ในชีต', 'menuClearValidation_')
    .addSeparator()
    .addItem('เปิดอัพเดตอัตโนมัติทุกชั่วโมง', 'menuInstallTrigger_')
    .addItem('ตรวจสอบการจับคู่ HN (log)', 'debugTnssMatch')
    .addToUi();
}

function menuRefresh_() { toast_(refreshTnssFormulas()); }
function menuLinkRcq_() { toast_(linkRcqToSumary()); }
function menuSetup_() { toast_(setupSheets()); }
function menuClearValidation_() { toast_(clearAllValidations()); }

function menuInstallTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshTnssFormulas')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshTnssFormulas').timeBased().everyHours(1).create();
  toast_('เปิดอัพเดตอัตโนมัติทุกชั่วโมงแล้ว');
}

function toast_(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(String(msg), 'PCH-SumYa AI', 8); }
  catch (e) { Logger.log(msg); }
}

/* ───────────────────────── Web API ───────────────────────── */
function doGet(e) {
  const p = (e && e.parameter) || {};

  // ไม่มี key = เปิดหน้าแอป (ต้องมีไฟล์ HTML ชื่อ Index ในโครงการ Apps Script)
  if (!p.key && !p.action) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('PCH-SumYa AI · รพ.ปากชม')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (p.key !== CONFIG.API_KEY) return json_({ ok: false, error: 'invalid key' });
  try {
    if (p.action === 'lookup') return json_({ ok: true, rows: readSheet_('lookup') });
    if (p.action === 'patients') return json_({ ok: true, rows: readSheet_('patients') });
    if (p.action === 'tnss') return json_({ ok: true, rows: readSheet_('tnss_daily') });
    if (p.action === 'rcq36') return json_({ ok: true, rows: readSheet_('rcq36') });
    return json_({ ok: true, rows: readSheet_('sumary') });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** แอปส่งด้วย text/plain เพื่อเลี่ยง CORS preflight */
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (body.key !== CONFIG.API_KEY) return json_({ ok: false, error: 'invalid key' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (body.action) {
      case 'session': return json_(saveSession_(body.data || {}));
      case 'tnss':    { const r = appendRow_('tnss_daily', TNSS_HEADERS, body.data || {}); refreshTnssFormulas(); return json_(r); }
      case 'patient': return json_(appendRow_('patients', PATIENT_HEADERS, body.data || {}));
      case 'rcq36':   { const r = appendRow_('rcq36', RCQ_HEADERS, body.data || {}); linkRcqToSumary(); return json_(r); }
      case 'audio':   return json_({ ok: true, url: saveAudio_(body) });
      case 'ping':    return json_({ ok: true, at: new Date().toISOString() });
      default:        return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function saveSession_(d) {
  if (d['ยินยอม PDPA'] !== 'ยินยอม') return { ok: false, error: 'ยังไม่ได้รับความยินยอม (PDPA)' };
  const herbs = ['ขิง (%)', 'ดอกปีบ (%)', 'มะแขวน (%)', 'ผิวมะกรูด (%)', 'ข่า (%)'];
  const sum = herbs.reduce((a, k) => a + (Number(d[k]) || 0), 0);
  if (sum && Math.abs(sum - 100) > 0.5) return { ok: false, error: 'สัดส่วนสมุนไพรรวมได้ ' + sum + '% (ต้องเท่ากับ 100%)' };
  const r = appendRow_('sumary', SUMARY_HEADERS, d);
  refreshTnssFormulas();
  return r;
}

/**
 * เติมสูตรดึง TNSS หลังจากชีต tnss_daily ลงคอลัมน์ AJ / AK / AL ของทุกแถวในชีต sumary
 * เฉลี่ยคะแนนวันที่ 7–14 หลังวันรับบริการ · รันซ้ำได้ ปลอดภัย
 * เรียกเองได้จากเมนู Apps Script เมื่อแก้ข้อมูลใน tnss_daily แล้วต้องการรีเฟรช
 */
const FU_FROM_DAY = 1;    // เริ่มนับวันที่เท่าไรหลังรับบริการ (1 = วันถัดไป)
const FU_TO_DAY = 999;    // ถึงวันที่เท่าไร (999 = ทุกวันหลังจากนั้น) · งานวิจัยจริงใช้ 7 กับ 14

/** แปลผลคะแนนรวม TNSS 0-12 ตามเกณฑ์มาตรฐาน */
function tnssBand_(total) {
  const n = Number(total);
  if (!isFinite(n) || total === '' || total === null) return '';
  if (n < 3) return 'อาการเล็กน้อยมาก';
  if (n <= 6) return 'อาการเล็กน้อย';
  if (n <= 9) return 'อาการปานกลาง';
  return 'อาการรุนแรง';
}

/** HN → ตัวเลขล้วน ตัดศูนย์หน้า/ช่องว่าง เทียบข้ามชีตได้แม้ชนิดข้อมูลต่างกัน */
function hnKey_(v) {
  const s = String(v == null ? '' : v).trim();
  const digits = s.replace(/[^0-9]/g, '');
  return (digits || s).replace(/^0+/, '');
}

/** รับได้ทั้ง Date, 2026-09-18, 18/09/2026, 18/09/2569 → จำนวนวันนับจาก epoch */
function dayNum_(v) {
  if (v instanceof Date && !isNaN(v)) return Math.floor(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()) / 86400000);
  const s = String(v == null ? '' : v).trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return Math.floor(Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000);
  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
  if (m) { let y = +m[3]; if (y > 2400) y -= 543; return Math.floor(Date.UTC(y, +m[2] - 1, +m[1]) / 86400000); }
  const d = new Date(s);
  return isNaN(d) ? null : Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000);
}

/**
 * คำนวณ TNSS หลังจากชีต tnss_daily แล้วเขียนเป็น "ค่า" ลงคอลัมน์ AG / AH / AI ของชีต sumary
 * ใช้การคำนวณในสคริปต์แทนสูตรชีต เพราะทนกับกรณี HN/วันที่ถูกเก็บเป็นข้อความ
 * เฉลี่ยคะแนนวันที่ 7–14 หลังวันรับบริการ · รันซ้ำได้ ปลอดภัย
 */
function refreshTnssFormulas() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName('sumary');
  const dh = ss.getSheetByName('tnss_daily');
  const last = sh.getLastRow();
  if (last < 2) return 'ชีต sumary ยังไม่มีข้อมูล';

  // อ่าน tnss_daily → จัดกลุ่มตาม HN
  const byHN = {};
  let dailyCount = 0;
  if (dh && dh.getLastRow() > 1) {
    const dv = dh.getRange(2, 1, dh.getLastRow() - 1, 10).getValues();
    dv.forEach(r => {
      const hn = hnKey_(r[2]);            // C  HN
      const day = dayNum_(r[3]);          // D  วันที่ประเมิน
      if (!hn || day === null) return;
      (byHN[hn] = byHN[hn] || []).push({
        day,
        v: [Number(r[5]) || 0, Number(r[6]) || 0, Number(r[7]) || 0, Number(r[8]) || 0],  // F–I
        total: r[9] === '' || r[9] == null ? null : Number(r[9])                           // J
      });
      dailyCount++;
    });
  }

  // อ่าน HN (D) / วันที่รับบริการ (I) / TNSS รวมก่อน (AF)
  const hnCol = sh.getRange(2, 4, last - 1, 1).getValues();
  const svcCol = sh.getRange(2, 12, last - 1, 1).getValues();
  const preCol = sh.getRange(2, 35, last - 1, 1).getValues();

  const out = [], bands = [], preBands = [];
  let filled = 0;
  for (let i = 0; i < last - 1; i++) {
    const hn = hnKey_(hnCol[i][0]);
    const svc = dayNum_(svcCol[i][0]);
    const list = (byHN[hn] || []).filter(x =>
      svc === null ? true : (x.day >= svc + FU_FROM_DAY && x.day <= svc + FU_TO_DAY));
    if (!hn || !list.length) { out.push(['', '', '']); bands.push(['']); continue; }

    const avg = k => list.reduce((a, x) => a + x.v[k], 0) / list.length;
    const parts = [0, 1, 2, 3].map(k => Math.round(avg(k)));
    const totals = list.map(x => x.total === null ? x.v.reduce((a, b) => a + b, 0) : x.total);
    const post = Math.round(totals.reduce((a, b) => a + b, 0) / totals.length * 10) / 10;
    const pre = Number(preCol[i][0]);
    const drop = pre > 0 ? Math.round((1 - post / pre) * 1000) / 10 : '';
    out.push([parts.join('/'), post, drop]);
    bands.push([tnssBand_(post)]);
    filled++;
  }
  // ล้างสูตร/ค่าเก่าในช่วง AJ:AL ก่อน (กันสูตรที่วางมือไว้ค้างอยู่)
  const target = sh.getRange(2, 36, last - 1, 3);
  const stale = target.getFormulas().reduce((a, r) => a + r.filter(x => x).length, 0);
  target.clearContent();

  sh.getRange(2, 36, out.length, 3).setValues(out);
  // AV แปลผลก่อน (จาก TNSS รวมก่อน) · AW แปลผลหลัง
  for (let i = 0; i < last - 1; i++) preBands.push([tnssBand_(preCol[i][0])]);
  sh.getRange(2, 48, preBands.length, 1).setValues(preBands);
  sh.getRange(2, 49, bands.length, 1).setValues(bands);
  SpreadsheetApp.flush();
  return 'อ่าน tnss_daily ' + dailyCount + ' แถว · เติมผลได้ ' + filled + ' จาก ' + (last - 1) + ' เคส'
    + (stale ? ' · ล้างสูตรเก่า ' + stale + ' ช่อง' : '');
}

/** ล้างสูตรที่วางมือไว้ทั้งบล็อก TNSS หลัง (AJ:AL) — ใช้เมื่อสูตรเก่าค้าง/เลื่อนตำแหน่ง */
function clearStaleTnssFormulas() {
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('sumary');
  const last = sh.getLastRow();
  if (last < 2) return 'ยังไม่มีข้อมูล';
  let found = [];
  // กวาดหาสูตรค้างในคอลัมน์ AH–AN (34–40) ยกเว้นคอลัมน์ที่เป็นข้อมูลจริง
  for (let c = 34; c <= 40; c++) {
    const fs = sh.getRange(2, c, last - 1, 1).getFormulas().filter(r => r[0]).length;
    if (fs) found.push(sh.getRange(1, c).getValue() + ' (คอลัมน์ ' + c + '): ' + fs + ' สูตร');
  }
  sh.getRange(2, 36, last - 1, 3).clearContent();
  Logger.log(found.length ? found.join('\n') : 'ไม่พบสูตรค้าง');
  return found.length ? ('พบสูตรค้าง: ' + found.join(' | ') + ' — ล้าง AJ:AL แล้ว รัน refreshTnssFormulas() ต่อ') : 'ล้าง AJ:AL แล้ว รัน refreshTnssFormulas() ต่อ';
}

/**
 * เติมคะแนน RCQ-36 (ก่อน) ลงคอลัมน์ AX / AY ของชีต sumary
 * จับคู่ด้วย "รหัสบันทึก sumary" ก่อน ถ้าไม่มีจึงใช้ HN (เอาชุดก่อนล่าสุด)
 */
function linkRcqToSumary() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName('sumary');
  const rh = ss.getSheetByName('rcq36');
  const last = sh.getLastRow();
  if (last < 2) return 'ชีต sumary ยังไม่มีข้อมูล';
  const byId = {}, byHN = {};
  if (rh && rh.getLastRow() > 1) {
    const head = rh.getRange(1, 1, 1, rh.getLastColumn()).getValues()[0];
    const iRef = head.indexOf('รหัสบันทึก sumary');
    const iHN = head.indexOf('HN');
    const iWhen = head.indexOf('ครั้งที่ประเมิน');
    const iTot = head.indexOf('คะแนนรวม (34-170)');
    const iAvg = head.indexOf('เฉลี่ยต่อข้อ');
    rh.getRange(2, 1, rh.getLastRow() - 1, rh.getLastColumn()).getValues().forEach(r => {
      if (String(r[iWhen]).indexOf('หลัง') >= 0) return;   // เก็บเฉพาะชุด "ก่อน"
      const rec = [r[iTot], rcqBand_(r[iAvg])];
      if (r[iRef]) byId[String(r[iRef]).trim()] = rec;
      if (r[iHN]) byHN[hnKey_(r[iHN])] = rec;
    });
  }
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  const hns = sh.getRange(2, 4, last - 1, 1).getValues();
  const out = ids.map((row, i) => {
    const hit = byId[String(row[0]).trim()] || byHN[hnKey_(hns[i][0])];
    return hit ? [hit[0], hit[1]] : ['', ''];
  });
  sh.getRange(2, 50, out.length, 2).setValues(out);
  SpreadsheetApp.flush();
  return 'เชื่อมคะแนน RCQ-36 แล้ว ' + out.filter(r => r[0] !== '').length + ' เคส';
}

/** วินิจฉัยว่าทำไมบางเคสไม่มีผล — ดูใน Execution log */
function debugTnssMatch() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName('sumary'), dh = ss.getSheetByName('tnss_daily');
  const s = sh.getRange(2, 1, Math.min(5, sh.getLastRow() - 1), 12).getValues();
  Logger.log('— sumary (HN | วันที่รับบริการ | ชนิดข้อมูลวันที่) —');
  s.forEach(r => Logger.log(hnKey_(r[3]) + ' | ' + r[11] + ' | ' + (r[11] instanceof Date ? 'Date' : typeof r[11]) + ' | day=' + dayNum_(r[11])));
  if (dh && dh.getLastRow() > 1) {
    const d = dh.getRange(2, 1, Math.min(8, dh.getLastRow() - 1), 10).getValues();
    Logger.log('— tnss_daily (HN | วันที่ประเมิน | รวม) —');
    d.forEach(r => Logger.log(hnKey_(r[2]) + ' | ' + r[3] + ' | ' + (r[3] instanceof Date ? 'Date' : typeof r[3]) + ' | day=' + dayNum_(r[3]) + ' | total=' + r[9]));
  } else Logger.log('tnss_daily ยังไม่มีข้อมูล');
  Logger.log('— หัวคอลัมน์จริงของ sumary —');
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  hdr.forEach((h, i) => { if (i >= 31 && i <= 40) Logger.log('คอลัมน์ ' + (i + 1) + ' = ' + h); });
  return 'ดูผลใน Execution log';
}

function appendRow_(name, headers, d) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(name) || ensureSheet_(ss, name, headers);
  if (!d['วันที่-เวลาบันทึก']) {
    d['วันที่-เวลาบันทึก'] = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  }
  if (!d['รหัสบันทึก'] && headers[0] === 'รหัสบันทึก') d['รหัสบันทึก'] = newId_(name);
  sh.appendRow(headers.map(h => (d[h] === undefined || d[h] === null) ? '' : d[h]));
  return { ok: true, id: d['รหัสบันทึก'] || '', row: sh.getLastRow() };
}

function newId_(name) {
  const stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd-HHmmss');
  return (name === 'tnss_daily' ? 'TNSS-' : 'SUMA-') + stamp;
}

function readSheet_(name) {
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getDisplayValues();
  const head = v.shift();
  return v.map(r => { const o = {}; head.forEach((h, i) => o[h] = r[i]); return o; });
}

/** body: {filename, mimeType, base64} → คืน URL ไฟล์ใน Drive */
function saveAudio_(body) {
  const folder = CONFIG.DRIVE_FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID)
    : getOrCreateFolder_('PCH-SumYa AI Breath Audio');
  const blob = Utilities.newBlob(Utilities.base64Decode(body.base64), body.mimeType || 'audio/webm', body.filename || ('breath-' + Date.now() + '.webm'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
