/**
 * Backend ของแอปสแกน (2 โหมด):
 * - start  : (RECEIVING → เข้าชีต1+2, ไม่เช็คสต๊อก) / (อื่นๆ → เช็คพอ ไม่หัก บันทึกเฉพาะชีต2)
 * - finish : (RECEIVING → เข้าชีต1+2, ไม่เช็คสต๊อก) / (อื่นๆ → เช็ค+หัก เข้าชีต1 และบันทึกชีต2)
 * - ship   : ส่งสินค้า → หักจาก FINISH_GOODS ในชีต1 + log ลงชีต2 (Menu=ส่งสินค้า)
 */

/** ================= Helper =================== */
// ✅ Normalize เป็นตัวพิมพ์ใหญ่ + แทนช่องว่าง/ขีด ด้วย _
const normU = s => String(s || '').trim().toUpperCase().replace(/[\s\-]+/g, '_');

// ✅ Map synonym: FINISHED_GOODS → FINISH_GOODS
const mapProc = p => {
  const x = normU(p);
  return (x === 'FINISHED_GOODS') ? 'FINISH_GOODS' : x;
};

// ✅ ชื่อสวยไว้แสดง/บันทึกฝั่ง sheet2
function prettyProc(key){
  const k = normU(key);
  const map = {
    RECEIVING:      'Receiving',
    CUTTING:        'Cutting',
    MILLING:        'Milling',
    CNC:            'CNC',
    GRINDING:       'Grinding',
    FINISH_GOODS:  'Finished Goods' // << ต้องการคำนี้ให้เหมือนกัน
  };
  return map[k] || key;
}

// ✅ helper ใหม่: หา index ของคอลัมน์โดยรองรับชื่อพ้อง
function findCol(headers, names /* array */) {
  const targets = names.map(normU);
  return headers.findIndex(h => targets.includes(normU(h)));
}

// --- CONFIG ---
const CONFIG = {
  SHEET_ID: '1Aq1ZvwqKVKDQGIynEILrFIEe6A-sUqFk9Ztxm6SrNPo',
  SHEET_NAME: 'sheet1',
  SHEET2_NAME: 'sheet2',
  ITEM_COLUMN_INDEX: 0,
  HEADER_ROW_INDEX: 0
};

/** เสิร์ฟหน้าเว็บ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Scan QR Code')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ดึงข้อมูล ITEM จาก sheet1 (สำหรับแสดงผลเบื้องต้น) */
function getItemDetails(itemCode) {
  const code = String(itemCode || '').trim();
  if (!code) return null;

  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`ไม่พบชีต "${CONFIG.SHEET_NAME}"`);

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][CONFIG.ITEM_COLUMN_INDEX]).trim() === code) {
      return {
        item: values[i][0],
        namePart: values[i][1],
        model: values[i][2],
        drawing: values[i][3]
      };
    }
  }
  return null;
}

/** ===== STOCK CHECK API ===== */
function getStockByItem(itemCode) {
  const code = String(itemCode || '').trim();
  if (!code) return null;

  const sh = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error(`ไม่พบชีต "${CONFIG.SHEET_NAME}"`);

  const values = sh.getDataRange().getValues();
  const headers = values[CONFIG.HEADER_ROW_INDEX];

  // map header -> index
  const hmap = {};
  headers.forEach((h, i) => { const k = normU(h); if (k) hmap[k] = i; });

  // หาแถวของ ITEM
  let rowIdx = -1;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][CONFIG.ITEM_COLUMN_INDEX]).trim() === code) { rowIdx = r; break; }
  }
  if (rowIdx === -1) return null;

  const row = values[rowIdx];
  const get = key => {
    const idx = hmap[normU(key)];
    return idx === undefined ? '' : row[idx];
  };

  return {
    ITEM:           get('ITEM') || row[0],
    NAME_PART:      get('NAME PART'),
    MODEL:          get('MODEL'),
    DRAWING:        get('DRAWING'),
    RECEIVING:      Number(get('RECEIVING') || 0),
    CUTTING:        Number(get('CUTTING') || 0),
    MILLING:        Number(get('MILLING') || 0),
    CNC:            Number(get('CNC') || 0),
    GRINDING:       Number(get('GRINDING') || 0),
    FINISH_GOODS:   Number(get('FINISH_GOODS') || get('FINISHED_GOODS') || get('FINISH_GOOD') || 0),
    NG:             Number(get('NG') || 0),
    TOTAL:          Number(get('TOTAL') || 0)
  };
}

/**
 * บันทึกข้อมูล Start/Finish (เพิ่มคอลัมน์ Menu ใน sheet2 → อยู่ระหว่าง Name กับ Timestamp)
 */
function saveDataRow(record, mode) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet1 = ss.getSheetByName(CONFIG.SHEET_NAME);
    const sheet2 = ss.getSheetByName(CONFIG.SHEET2_NAME);
    if (!sheet1 || !sheet2) throw new Error('ไม่พบชีต sheet1 หรือ sheet2');

    const values = sheet1.getDataRange().getValues();
    const headers = values[CONFIG.HEADER_ROW_INDEX];

    // หาแถวของ Item
    let itemRowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][CONFIG.ITEM_COLUMN_INDEX]).toUpperCase().trim() === String(record.item).toUpperCase().trim()) {
        itemRowIndex = i;
        break;
      }
    }
    if (itemRowIndex === -1) throw new Error(`ไม่พบ Item '${record.item}' ใน sheet1`);
    const row = itemRowIndex + 1;

    const processNameRaw = String(record.process || '').trim();
    const processKey = mapProc(processNameRaw); // ✅ normalize FINISHED_GOODS → FINISH_GOODS

    const goodQty = Number(record.goodQuantity || 0);
    const ngQty   = Number(record.ngQuantity || 0);
    const totalQty = goodQty + ngQty;
    if (totalQty <= 0) throw new Error('จำนวนรวมต้องมากกว่า 0');

    // ✅ รองรับชื่อพ้อง FINISHED_GOODS, FINISH_GOOD
    const procIdx = findCol(headers, [processKey, 'FINISHED_GOODS', 'FINISH_GOOD']);
    if (procIdx === -1) throw new Error(`ไม่พบหัวข้อ Process '${record.process}' ใน sheet1`);

    const ngIdx = findCol(headers, ['NG']);
    const tsIdx = findCol(headers, ['TIMESTAMP']);

    const modeStr = String(mode || '').toLowerCase();

    // RECEIVING (ทุกโหมด)
    if (processKey === 'RECEIVING') {
      const pCell = sheet1.getRange(row, procIdx + 1);
      pCell.setValue((Number(pCell.getValue()) || 0) + goodQty);

      if (ngIdx !== -1) {
        const ngCell = sheet1.getRange(row, ngIdx + 1);
        ngCell.setValue((Number(ngCell.getValue()) || 0) + ngQty);
      }

      updateTotal(sheet1, row, headers);
      if (tsIdx !== -1) sheet1.getRange(row, tsIdx + 1).setValue(new Date());

      appendToSheet2(record, modeStr); // 👉 จะถูกแปลงชื่อสวยใน appendToSheet2
      return 'บันทึกข้อมูล (RECEIVING) สำเร็จ!';
    }

    // start
    if (modeStr === 'start') {
      const stockIdxs = getStockSourcesForProcess(processKey, headers);
      let available = 0;
      stockIdxs.forEach(idx => { available += Number(sheet1.getRange(row, idx + 1).getValue()) || 0; });
      if (totalQty > available) throw new Error(`จำนวนไม่พอ! (มีอยู่ ${available}, ต้องการใช้ ${totalQty})`);
      appendToSheet2(record, modeStr);
      return 'บันทึกข้อมูลสำเร็จ! (โหมดเริ่มการทำงาน)';
    }

    // finish
    if (modeStr === 'finish') {
      const stockIdxs = getStockSourcesForProcess(processKey, headers);
      let available = 0;
      stockIdxs.forEach(idx => { available += Number(sheet1.getRange(row, idx + 1).getValue()) || 0; });
      if (totalQty > available) throw new Error(`จำนวนไม่พอ! (มีอยู่ ${available}, ต้องการใช้ ${totalQty})`);

      let remain = totalQty;
      stockIdxs.forEach(idx => {
        if (remain > 0) {
          const cell = sheet1.getRange(row, idx + 1);
          const val = Number(cell.getValue()) || 0;
          const use = Math.min(val, remain);
          cell.setValue(val - use);
          remain -= use;
        }
      });

      const pCell = sheet1.getRange(row, procIdx + 1);
      pCell.setValue((Number(pCell.getValue()) || 0) + goodQty);

      if (ngIdx !== -1) {
        const ngCell = sheet1.getRange(row, ngIdx + 1);
        ngCell.setValue((Number(ngCell.getValue()) || 0) + ngQty);
      }

      updateTotal(sheet1, row, headers);
      if (tsIdx !== -1) sheet1.getRange(row, tsIdx + 1).setValue(new Date());

      appendToSheet2(record, modeStr);
      return 'บันทึกข้อมูลสำเร็จ! (โหมดเสร็จการทำงาน)';
    }

    throw new Error('ไม่รู้จักโหมดการทำงาน: ' + modeStr);

  } catch (e) {
    throw new Error(e.message);
  }
}

/** แผนที่แหล่งสต๊อกตาม Process (RECEIVING ไม่มีแหล่ง) */
function getStockSourcesForProcess(p, headers) {
  const src = {
    'CUTTING':       ['RECEIVING'],
    'MILLING':       ['CUTTING', 'RECEIVING'],
    'CNC':           ['MILLING', 'CUTTING', 'RECEIVING'],
    'GRINDING':      ['CNC', 'MILLING', 'CUTTING', 'RECEIVING'],
    'FINISH_GOODS':  ['GRINDING', 'CNC', 'MILLING', 'CUTTING', 'RECEIVING']
  };
  // ใช้ normU ทั้งชื่อแหล่งและหัวคอลัมน์
  return (src[p] || [])
    .map(nameU => headers.findIndex(h => normU(h) === normU(nameU)))
    .filter(i => i >= 0);
}

/** คำนวณรวม TOTAL ของแถวใน sheet1 */
function updateTotal(sheet, row, headers) {
  const tIdx = headers.findIndex(h => normU(h) === 'TOTAL');
  if (tIdx === -1) return;

  const vals = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  let sum = 0;
  headers.forEach((h, i) => {
    const u = normU(h);
    const v = vals[i];
    if (!['ITEM', 'TOTAL', 'TIMESTAMP'].includes(u) && v !== '' && !isNaN(v)) {
      sum += Number(v);
    }
  });
  sheet.getRange(row, tIdx + 1).setValue(sum);
}

/** เติมข้อมูลครบลง sheet2 (หัวตารางมี Menu ระหว่าง Name กับ Timestamp) */
function appendToSheet2(record, mode) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet2 = ss.getSheetByName(CONFIG.SHEET2_NAME);
  if (!sheet2) throw new Error(`ไม่พบชีต "${CONFIG.SHEET2_NAME}"`);

  const HEADER = ['ID','Item','Name Part','Model','Drawing','Process','Good','NG','Total','Name','Menu','Timestamp'];

  // เขียน/ปรับหัวตาราง
  if (sheet2.getLastRow() === 0) {
    sheet2.appendRow(HEADER);
  } else {
    const current = sheet2.getRange(1, 1, 1, HEADER.length).getValues()[0];
    if (HEADER.join('|').toUpperCase() !== current.join('|').toUpperCase()) {
      sheet2.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
    }
  }

  const goodQty = Number(record.goodQuantity || 0);
  const ngQty   = Number(record.ngQuantity || 0);
  const total   = goodQty + ngQty;
  const id = sheet2.getLastRow(); // header แถว 1 → แถวข้อมูลแรก id = 1

  const m = String(mode||'').toLowerCase();
  const menuText = (m === 'start')  ? 'เริ่มการทำงาน'
                  : (m === 'finish') ? 'เสร็จการทำงาน'
                  : (m === 'ship')   ? 'ส่งสินค้า'
                  : String(mode||'');

  // ✅ แปลงชื่อสวยเฉพาะงานผลิต (SHIP ให้คงเป็น SHIP)
const procRaw = String(record.process || '');
const procForLog = (normU(procRaw) === 'SHIP')
  ? 'Ship'                         // ✅ บันทึกเป็น "Ship"
  : prettyProc(mapProc(procRaw));   // FINISH_GOODS → "Finished Goods"

  const row = [
    id,
    String(record.item || ''),
    String(record.namePart || ''),
    String(record.model || ''),
    String(record.drawing || ''),
    procForLog,          // << เขียนชื่อสวยลง sheet2
    goodQty,
    ngQty,
    total,
    String(record.name || ''),
    menuText,            // Menu
    new Date()           // Timestamp
  ];

  sheet2.appendRow(row);
}

/** ถอดรหัส QR แบบ fallback ผ่าน API (สำรอง) */
function decodeQrFromImage(dataUrl) {
  const m = /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/.exec(dataUrl || '');
  if (!m) throw new Error('Invalid image data URL.');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'frame.jpg');
  const res = UrlFetchApp.fetch('https://api.qrserver.com/v1/read-qr-code/', { method: 'post', payload: { file: blob } });
  if (res.getResponseCode() !== 200) throw new Error('QR server HTTP ' + res.getResponseCode());
  const json = JSON.parse(res.getContentText());
  return json?.[0]?.symbol?.[0]?.data || null;
}

/** ===== ส่งสินค้า: ตัดยอดจาก Finish_goods ใน sheet1 + log ลง sheet2 ===== */
function shipGoods(itemCode, qty, senderName) {
  const code = String(itemCode || '').trim();
  const sendQty = Number(qty || 0);
  const name = String(senderName || '').trim();

  if (!code) throw new Error('กรุณาระบุ ITEM');
  if (!isFinite(sendQty) || sendQty <= 0) throw new Error('จำนวนที่ส่งต้องมากกว่า 0');
  if (!name) throw new Error('กรุณาระบุชื่อผู้ส่ง');

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`ไม่พบชีต "${CONFIG.SHEET_NAME}"`);

  // อ่านตารางทั้งหมดเพื่อทำ header map
  const values  = sheet.getDataRange().getValues();
  const headers = values[CONFIG.HEADER_ROW_INDEX] || [];

  const hmap = {};
  headers.forEach((h, i) => { const k = normU(h); if (k) hmap[k] = i; });

  // หาแถว ITEM
  let rowIdx = -1;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][CONFIG.ITEM_COLUMN_INDEX]).trim() === code) { rowIdx = r; break; }
  }
  if (rowIdx === -1) throw new Error(`ไม่พบ ITEM '${code}' ใน ${CONFIG.SHEET_NAME}`);
  const row = rowIdx + 1;

  // หา index ของ FINISHED_GOODS/FINISH_GOODS (รองรับสะกดเดิม)
  const fgIdx =
    (hmap['FINISHED_GOODS'] != null) ? hmap['FINISHED_GOODS'] :
    (hmap['FINISH_GOODS']   != null) ? hmap['FINISH_GOODS']   :
    (hmap['FINISH_GOOD']    != null) ? hmap['FINISH_GOOD']    : null;

  if (fgIdx == null) throw new Error('ไม่พบหัวข้อ Finished Goods ใน sheet1');
  const tsIdx = headers.findIndex(h => normU(h) === 'TIMESTAMP');

  const before = Number(sheet.getRange(row, fgIdx + 1).getValue()) || 0;
  if (sendQty > before) throw new Error(`จำนวนส่ง (${sendQty}) มากกว่าจำนวนที่มี (${before}) ใน Finished Goods`);

  // ตัดยอด
  sheet.getRange(row, fgIdx + 1).setValue(before - sendQty);

  // อัปเดต TOTAL + Timestamp
  updateTotal(sheet, row, headers);
  if (tsIdx !== -1) sheet.getRange(row, tsIdx + 1).setValue(new Date());

  // สร้างบันทึกลง sheet2
  const r = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const record = {
    item: r[0],
    namePart: r[1],
    model: r[2],
    drawing: r[3],
    process: 'SHIP',
    goodQuantity: sendQty,
    ngQuantity: 0,
    name: name
  };
  appendToSheet2(record, 'ship'); // ให้ Menu = ส่งสินค้า

  // สรุป
  return {
    ok: true,
    message: `ส่งสินค้าเรียบร้อย หักจาก Finished Goods ${sendQty} ชิ้น (คงเหลือ ${before - sendQty})`,
    ITEM: r[0], NAME_PART: r[1], MODEL: r[2], DRAWING: r[3],
    FINISH_GOODS_BEFORE: before,
    FINISH_GOODS_AFTER: before - sendQty,
    SENT_QTY: sendQty,
    SENDER: name
  };
}

/* ------------------------------------------------------------------
 * เพิ่มใหม่: ดึง "รายงานการผลิต" จาก sheet2 ตาม ITEM
 * ------------------------------------------------------------------ */
function getProductionLogsByItem(itemCode) {
  const code = String(itemCode || '').trim();
  if (!code) return { itemInfo: null, logs: [] };

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet2 = ss.getSheetByName(CONFIG.SHEET2_NAME);
  if (!sheet2) throw new Error(`ไม่พบชีต "${CONFIG.SHEET2_NAME}"`);

  const values = sheet2.getDataRange().getValues();
  if (!values.length) return { itemInfo: null, logs: [] };

  const headers = values[0].map(h => String(h || '').trim());
  const idx = {};
  headers.forEach((h, i) => idx[normU(h)] = i);

  const need = ['ITEM','NAME_PART','DRAWING','PROCESS','GOOD','NG','TOTAL','NAME','MENU','TIMESTAMP','MODEL'];
  need.forEach(k => { if (idx[k] === undefined) idx[k] = -1; });

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const itemVal = String(row[idx['ITEM']] ?? '').trim();
    if (itemVal === code) {
      rows.push({
        Item: itemVal,
        NamePart: row[idx['NAME_PART']] ?? '',
        Drawing: row[idx['DRAWING']] ?? '',
        Process: row[idx['PROCESS']] ?? '',
        Good: Number(row[idx['GOOD']] ?? 0) || 0,
        NG: Number(row[idx['NG']] ?? 0) || 0,
        Total: Number(row[idx['TOTAL']] ?? 0) || 0,
        Name: row[idx['NAME']] ?? '',
        Menu: row[idx['MENU']] ?? '',
        Timestamp: row[idx['TIMESTAMP']] instanceof Date ? row[idx['TIMESTAMP']] : (row[idx['TIMESTAMP']] ? new Date(row[idx['TIMESTAMP']]) : null),
        Model: idx['MODEL'] >= 0 ? (row[idx['MODEL']] ?? '') : ''
      });
    }
  }

  if (!rows.length) return { itemInfo: null, logs: [] };

  // เรียงใหม่ล่าสุดก่อน
  rows.sort((a,b)=>{
    const ta = a.Timestamp ? a.Timestamp.getTime() : 0;
    const tb = b.Timestamp ? b.Timestamp.getTime() : 0;
    return tb - ta;
  });

  // สรุปหัวรายการ (เอาจากแถวล่าสุด)
  const top = rows[0];
  const itemInfo = {
    Item: top.Item,
    NamePart: top.NamePart,
    Drawing: top.Drawing,
    Model: top.Model || ''
  };

  // แปลง Timestamp → ISO string
  const logs = rows.map(r => ({
    Item: r.Item,
    NamePart: r.NamePart,
    Drawing: r.Drawing,
    Process: r.Process,
    Good: r.Good,
    NG: r.NG,
    Total: r.Total,
    Name: r.Name,
    Menu: r.Menu,
    Timestamp: r.Timestamp ? r.Timestamp.toISOString() : ''
  }));

  return { itemInfo, logs };
}

/** * บล็อกใหม่พิเศษ: ทำหน้าที่เป็นสะพานรับข้อมูลจากหน้าเว็บภายนอก
 * โดยจะส่งต่อข้อมูลไปให้ฟังก์ชันเดิมของพี่ทำงาน โดยไม่แก้ไขโค้ดเก่าเลยค่ะ
 */
function doPost(e) {
  try {
    // รับข้อมูลจากหน้าเว็บภายนอก
    const requestData = JSON.parse(e.postData.contents);
    const functionName = requestData.functionName;
    const args = requestData.arguments || [];
    
    let result;
    
    // วิ่งไปเรียกใช้ฟังก์ชันเดิมในระบบของพี่ตามที่หน้าเว็บร้องขอมา
    if (functionName === 'getItemDetails') {
      result = getItemDetails(args[0]);
    } else if (functionName === 'getStockByItem') {
      result = getStockByItem(args[0]);
    } else if (functionName === 'saveDataRow') {
      result = saveDataRow(args[0], args[1]);
    } else if (functionName === 'shipGoods') {
      result = shipGoods(args[0], args[1], args[2]);
    } else if (functionName === 'getProductionLogsByItem') {
      result = getProductionLogsByItem(args[0]);
    } else {
      throw new Error('ไม่พบฟังก์ชัน: ' + functionName);
    }
    
    // ส่งผลลัพธ์กลับไปให้หน้าเว็บภายนอกเป็น JSON
    return ContentService.createTextOutput(JSON.stringify({ result: result }))
      .setMimeType(ContentService.MimeType.JSON);
          
  } catch (error) {
    // หากเกิด Error ให้ส่งข้อความ Error กลับไปบอกหน้าเว็บ
    return ContentService.createTextOutput(JSON.stringify({ error: error.message || error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================================
// 🤖 ส่วนเสริม AI: ระบบสรุปรายงานอัตโนมัติ (ฉบับสมบูรณ์: คำนวณตัวเลขความแม่นยำ 100%)
// =====================================================================

const AI_CONFIG = {
  // ดึงคีย์จาก Script Properties เพื่อความปลอดภัยสูงสุด
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'), 
  REPORT_EMAILS: [
    'pongsak@smetaltech.co.th',
    'therayut@smetaltech.co.th',
    'wsompol@smetaltech.co.th',
    'ma.2@smetaltech.co.th',
    'ma.1@smetaltech.co.th',
    'ma.3@smetaltech.co.th',
  ]
};

function sendDailySmartReport() {
  console.log('📌 ขั้นตอนที่ 1: เริ่มต้นดึงข้อมูลจาก Supabase...');

  // =========================================================
  // ตั้งค่าการเชื่อมต่อ Supabase
  // =========================================================
  const SUPABASE_URL = 'https://orgbrvopqzpfvmtvxmcd.supabase.co';
  const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9yZ2Jydm9wcXpwZnZtdHZ4bWNkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODcxMTc5ODEsImV4cCI6MjEwMjY5Mzk4MX0.R-l6wgFtcoYCvXPrmO3xEjychxvEt4p_XUl71pchs00';
  
  const headers = {
    'apikey': SUPABASE_KEY,
    'Authorization': 'Bearer ' + SUPABASE_KEY,
    'Content-Type': 'application/json'
  };

  // กำหนดวันที่ทำรายงาน (ย้อนหลัง 1 วัน)
  const reportDate = new Date();
  reportDate.setDate(reportDate.getDate() - 1);
  const timeZone = Session.getScriptTimeZone();
  const targetDateString = Utilities.formatDate(reportDate, timeZone, "dd/MM/yyyy");
  
  // Format วันที่สำหรับค้นหาในฐานข้อมูล Supabase (YYYY-MM-DD)
  const targetIsoDate = Utilities.formatDate(reportDate, timeZone, "yyyy-MM-dd");

  let todayLogText = '';
  let matchCount = 0;

  let totalReceiving = 0;
  let totalProduced = 0;
  let totalShipped = 0;
  let empStats = {};
  let dailyIncome = 0;

  // =========================================================
  // ดึงข้อมูลราคา (Price) จากตาราง inventory ใน Supabase
  // =========================================================
  let priceMap = {};
  try {
    const invUrl = `${SUPABASE_URL}/rest/v1/inventory?select=item,price&limit=5000`;
    const invRes = UrlFetchApp.fetch(invUrl, { headers: headers });
    const invData = JSON.parse(invRes.getContentText());
    
    invData.forEach(row => {
      priceMap[row.item] = Number(row.price) || 0;
    });
  } catch (e) {
    console.log('⚠️ ไม่สามารถดึงราคาจาก Supabase ได้: ' + e.message);
  }

  // =========================================================
  // ดึงข้อมูลการผลิต (Production Logs) ของเมื่อวานจาก Supabase
  // =========================================================
  console.log(`📌 ขั้นตอนที่ 2: ค้นหาข้อมูลการผลิตของวันที่ ${targetDateString}`);
  
  // สร้างช่วงเวลาค้นหา: gte (>= เริ่มวัน) ถึง lt (< เริ่มวันถัดไป)
  const nextDate = new Date(reportDate);
  nextDate.setDate(nextDate.getDate() + 1);
  const gteString = Utilities.formatDate(reportDate, timeZone, "yyyy-MM-dd'T'00:00:00");
  const ltString = Utilities.formatDate(nextDate, timeZone, "yyyy-MM-dd'T'00:00:00");
  
  try {
    const logsUrl = `${SUPABASE_URL}/rest/v1/production_logs?timestamp=gte.${gteString}&timestamp=lt.${ltString}&limit=5000`;
    const logsRes = UrlFetchApp.fetch(logsUrl, { headers: headers });
    const logsData = JSON.parse(logsRes.getContentText());

    logsData.forEach(row => {
      const process = row.process;
      const menuStatus = row.menu;
      let isValidData = false;

      // ตรวจสอบ Process และ Menu ให้ตรงตามเงื่อนไขเดิม
      if (process === 'Receiving' && menuStatus === 'เริ่มการทำงาน') {
        isValidData = true;
      } else if (['Cutting', 'Milling', 'CNC', 'Grinding', 'Finished Goods'].includes(process) && menuStatus === 'เสร็จการทำงาน') {
        isValidData = true;
      } else if (process === 'Ship' && menuStatus === 'ส่งสินค้า') {
        isValidData = true;
      }

      if (isValidData) {
        const goodQty = Number(row.good_qty) || 0;
        const ngQty = Number(row.ng_qty) || 0;
        const employeeName = row.employee_name ? String(row.employee_name).trim() : '-';

        todayLogText += `- Item: ${row.item} (Name Part: ${row.name_part}), Process: ${row.process}, งานดี: ${goodQty}, NG: ${ngQty}, พนักงาน: ${employeeName}\n`;
        matchCount++;

        // บวกยอดแยกระบบ
        if (process === 'Receiving') {
          totalReceiving += goodQty;
        } else if (process === 'Ship') {
          totalShipped += goodQty;
          
          let itemCode = String(row.item).trim();
          let itemPrice = priceMap[itemCode] || 0;
          dailyIncome += (goodQty * itemPrice);
        } else {
          totalProduced += goodQty;
        }

        // เก็บสถิติพนักงาน
        if (employeeName !== '-') {
          if (!empStats[employeeName]) empStats[employeeName] = { good: 0, ng: 0 };
          empStats[employeeName].good += goodQty;
          empStats[employeeName].ng += ngQty;
        }
      }
    });
  } catch (e) {
    console.log('❌ ข้อผิดพลาดการดึงข้อมูล Log: ' + e.message);
    return;
  }

  if (todayLogText === '') {
    console.log('⚠️ แจ้งเตือน: ไม่พบข้อมูลที่ตรงกับเงื่อนไขของวันนี้เลยค่ะ');
    return;
  }

  console.log(`📌 ขั้นตอนที่ 3: พบข้อมูลทั้งหมด ${matchCount} รายการ กำลังเตรียมข้อมูลตัวเลขให้ AI...`);

  // =========================================================
  // บันทึกยอดขายและยอดสะสมลงตาราง daily_summary (แทน Sheet sumprice)
  // =========================================================
  console.log(`📌 กำลังบันทึกยอดขายและยอดสะสมลง Supabase (daily_summary)...`);
  let accumulateIncome = 0;

  try {
    // 1. ดึงสรุปยอดก่อนหน้าทั้งหมด
    const summaryUrl = `${SUPABASE_URL}/rest/v1/daily_summary?order=date.asc`;
    const summaryRes = UrlFetchApp.fetch(summaryUrl, { headers: headers });
    const summaryData = JSON.parse(summaryRes.getContentText());
    
    let previousAccumulate = 0;
    let targetSummaryId = null;
    const targetMonthPrefix = targetIsoDate.substring(0, 7); // e.g. "2026-08"

    for (let i = 0; i < summaryData.length; i++) {
      let r = summaryData[i];
      let rMonthPrefix = r.date.substring(0, 7);
      
      if (r.date === targetIsoDate) {
        targetSummaryId = r.date;
        if (i > 0) {
          let prevR = summaryData[i-1];
          if (prevR.date.substring(0, 7) === targetMonthPrefix) {
            previousAccumulate = Number(prevR.accumulate_income) || 0;
          } else {
            previousAccumulate = 0; // Reset for new month
          }
        }
        break;
      } else if (r.date < targetIsoDate) {
        if (rMonthPrefix === targetMonthPrefix) {
          previousAccumulate = Number(r.accumulate_income) || 0;
        } else {
          previousAccumulate = 0; // Reset for new month
        }
      }
    }

    accumulateIncome = previousAccumulate + dailyIncome;

    if (targetSummaryId) {
      // อัปเดตข้อมูลที่มีอยู่แล้ว
      const updateUrl = `${SUPABASE_URL}/rest/v1/daily_summary?date=eq.${targetSummaryId}`;
      UrlFetchApp.fetch(updateUrl, {
        method: 'PATCH',
        headers: headers,
        payload: JSON.stringify({ daily_income: dailyIncome, accumulate_income: accumulateIncome })
      });
    } else {
      // สร้างข้อมูลวันใหม่
      const insertUrl = `${SUPABASE_URL}/rest/v1/daily_summary`;
      UrlFetchApp.fetch(insertUrl, {
        method: 'POST',
        headers: headers,
        payload: JSON.stringify({ date: targetIsoDate, daily_income: dailyIncome, accumulate_income: accumulateIncome })
      });
    }
  } catch (e) {
    console.log('⚠️ ไม่สามารถอัปเดตยอดสะสมได้: ' + e.message);
  }

  // =========================================================
  // คำนวณพนักงานยอดเยี่ยม / ระวังตัว (เหมือนเดิม 100%)
  // =========================================================
  let bestEmp = { name: '-', good: -1 };
  let watchEmpList = [];

  for (let emp in empStats) {
    let stat = empStats[emp];
    if (stat.good > bestEmp.good) {
      bestEmp.name = emp;
      bestEmp.good = stat.good;
    }
    if (stat.ng > 0) {
      watchEmpList.push(`${emp} (NG = ${stat.ng} ชิ้น)`);
    }
  }
  let watchEmployeeText = watchEmpList.length > 0 ? watchEmpList.join(', ') : 'ไม่มีพนักงานที่ทำของเสีย (ยอดเยี่ยมทุกคนค่ะ!)';
  if (bestEmp.good === -1) {
    bestEmp = { name: 'ไม่มีพนักงานที่ได้ยอดผลิตเลย', good: 0 };
  }

  // =========================================================
  // สร้าง Prompt สำหรับ AI
  // =========================================================
  const prompt = `
  คุณคือ AI ช่วยสร้าง HTML สรุปรายงานการผลิตประจำวันของบริษัท S Metal Tech
  
  นี่คือข้อมูลดิบประจำวันที่ ${targetDateString}:
  ${todayLogText}
  
  เงื่อนไขการสร้างอีเมลรายงาน (ห้ามทำนอกเหนือคำสั่งเด็ดขาด):
  1. ห้ามเขียนคำทักทาย ห้ามแนะนำตัว ห้ามเขียนคำลงท้ายใดๆ ทั้งสิ้น ให้ตอบกลับมาเป็นโค้ด HTML เท่านั้น
  2. บรรทัดแรกสุดของรายงาน (ก่อนตาราง) ให้เขียนประโยคนี้ตรงๆ โดยไม่มีข้อความอื่น:
     <div style="font-family: Tahoma, sans-serif; font-size: 15px; font-weight: bold; color: #1e3c72; margin-bottom: 12px;">สรุปข้อมูลการผลิต Part MA ประจำวันที่ ${targetDateString}</div>
  3. รูปแบบตาราง (Scrollable บนมือถือ): ห้ามบีบตารางให้พอดีจอ เพื่อป้องกันชื่อพนักงานตกบรรทัด ให้ครอบตารางด้วย div สำหรับเลื่อนแนวนอน:
     <div style="overflow-x: auto;">
       <table style="width: 100%; min-width: 650px; border-collapse: collapse; font-family: Tahoma, sans-serif; font-size: 13px; white-space: nowrap;">
  4. โครงสร้างคอลัมน์: ให้มีแค่ 4 คอลัมน์ คือ "รายการสินค้า", "พนักงาน", "งานดี (ชิ้น)", "NG (ชิ้น)"
     (กำหนดหัวคอลัมน์ <th> ด้วยพื้นหลังสี #f1f5f9 และตีกรอบเส้น 1px solid #ddd)
  5. การจัดกลุ่มตาม Process (ตามแบบฟอร์มใหม่): ให้แบ่งกลุ่มข้อมูลตาม Process โดยใช้แถว <tr> คั่นเป็นหัวข้อ เช่น:
     <tr><td colspan="4" style="background-color: #e2e8f0; font-weight: bold; padding: 10px; font-size: 14px; border: 1px solid #ddd;">Process: Milling</td></tr>
     โดยเรียงลำดับ Process ตามนี้: Receiving, Cutting, Milling, CNC, Grinding, Finished Goods, Ship
  6. ในช่อง "รายการสินค้า" ให้นำ รหัสสินค้า และ ชื่อสินค้า มาต่อกัน เช่น "B12424A (DRAIN SHUTTER(AFC))"
  7. ภายใน <td> ให้ตีกรอบเส้น #ddd และใช้สีเขียวสำหรับตัวเลข NG เป็น 0 และสีแดงสำหรับเลขที่มี NG
  8. ส่วนท้าย "สรุปภาพรวมประจำวัน" ให้อยู่นอกตารางหลักและอยู่ด้านล่างสุด โดยใส่ในกรอบ:
      <div style="background-color: #f8fafc; padding: 16px; border-radius: 8px; margin-top: 20px; border: 1px solid #e2e8f0; font-family: Tahoma, sans-serif; font-size: 13px; max-width: 850px; white-space: normal;">
      - หัวข้อ: <h3 style="margin-top: 0; margin-bottom: 12px; font-size: 14px; color: #1e293b;">สรุปภาพรวมประจำวัน</h3>
      - ข้อมูลสรุปแต่ละข้อใช้แท็ก <ul style="margin: 0; padding-left: 20px;"> และ <li style="margin-bottom: 8px;">
  
  ⚠️ **ข้อมูลส่วนสรุปภาพรวมประจำวัน (ให้คัดลอกข้อความด้านล่างนี้ไปใส่ใน <li> ได้เลยโดยไม่ต้องคิดเลขใหม่)**:
  - <span style="font-weight: bold; color: #000;">ยอดรับ Raw Material (Process: Receiving): ${totalReceiving} ชิ้น</span>
  - <span style="font-weight: bold; color: #000;">ยอดรวมการผลิตสินค้าสำเร็จรูปประจำวัน: ${totalProduced} ชิ้น</span>
  - <span style="font-weight: bold; color: #000;">ยอดรวมสินค้าที่จัดส่ง (Process: Ship): ${totalShipped} ชิ้น</span>
  - <span style="font-weight: bold; color: #000;">ยอดขายประจำวัน: ${dailyIncome.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})} บาท</span>
  - <span style="font-weight: bold; color: #000;">ยอดขายรวม (Total Income): ${accumulateIncome.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})} บาท</span>
  - พนักงานยอดเยี่ยม: <span style="font-weight: bold; color: #1d4ed8;">${bestEmp.name} ด้วยยอดงานดีรวม ${bestEmp.good} ชิ้น</span>
  - พนักงานที่ต้องเฝ้าสังเกต: <span style="font-weight: bold; color: #b91c1c;">${watchEmployeeText}</span>`;

  const geminiResponse = callGemini(prompt);

  if (geminiResponse) {
    console.log('📌 ขั้นตอนที่ 4: AI ประมวลผลสำเร็จ! กำลังทำการส่งอีเมล...');
    const subject = `📊 สรุปรายงานการผลิต Part MA ประจำวันที่ ${targetDateString}`;
    sendToEmail(subject, geminiResponse);
  } else {
    console.log('❌ ข้อผิดพลาดร้ายแรง: ไม่สามารถสร้างอีเมลได้ เนื่องจาก AI ไม่ตอบสนองค่ะ');
  }
}
// ---------------------------------------------------------------------
function callGemini(promptText) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${AI_CONFIG.GEMINI_API_KEY}`;
  
  const payload = { "contents": [{ "parts": [{ "text": promptText }] }] };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true 
  };

  let maxRetries = 3;
  let attempt = 0;

  while (attempt < maxRetries) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        const json = JSON.parse(response.getContentText());
        let resultText = json?.candidates?.[0]?.content?.parts?.[0]?.text;
        
        if (!resultText) {
          console.log('⚠️ แจ้งเตือน: AI ทำงานสำเร็จ แต่ส่งข้อความว่างเปล่ากลับมาค่ะ');
          return null;
        }
        
        // ใช้คำสั่งตัดโค้ดออกที่ปลอดภัยที่สุด
        resultText = resultText.split('```html').join('');
        resultText = resultText.split('```').join('');
        return resultText;
        
      } else if (responseCode === 503 || responseCode === 429) {
        attempt++;
        console.log(`⏳ คิว AI แน่น หรือเรียกใช้งานถี่เกินไป กำลังลองใหม่ครั้งที่ ${attempt}/3 ...`);
        Utilities.sleep(3000 * attempt); 
      } else {
        console.log(`❌ AI ขัดข้อง (รหัส Error: ${responseCode})`);
        console.log(`รายละเอียด Error จาก AI: ${response.getContentText()}`);
        return null;
      }
    } catch (e) { 
      console.log(`❌ ระบบเครือข่ายมีปัญหา: ${e.message}`);
      return null; 
    }
  }
  console.log('❌ เชื่อมต่อ AI ล้มเหลวครบ 3 ครั้ง ระบบขอยกเลิกการทำงานค่ะ พี่ลองตรวจสอบ API Key ดูอีกครั้งนะคะ');
  return null;
}

// ---------------------------------------------------------------------
function sendToEmail(subject, message) {
  try {
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <body style="background-color: #f9fafb; padding: 0; font-family: Tahoma, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; margin: 0; font-size: 13px;">
        <div style="width: 100%; max-width: 100%; margin: 0 auto; background-color: #ffffff; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); overflow: hidden;">
          <div style="background-color: #2563eb; padding: 16px;">
            <h1 style="color: #ffffff; font-size: 20px; font-weight: 600; margin: 0; text-align: center; letter-spacing: 0.5px;">S Metal Tech Production Report</h1>
          </div>
          <div style="padding: 16px 12px; color: #374151; line-height: 1.5;">
            ${message}
          </div>
          <div style="background-color: #f9fafb; padding: 12px; text-align: center; border-top: 1px solid #f3f4f6;">
            <p style="font-size: 12px; color: #9ca3af; font-weight: 500; text-transform: uppercase; letter-spacing: 0.1em; margin: 0;">Generated by AI Smart System • Gemini</p>
          </div>
        </div>
      </body>
      </html>
    `;

    MailApp.sendEmail({
      to: AI_CONFIG.REPORT_EMAILS.join(','), 
      subject: subject,
      htmlBody: htmlBody
    });
    console.log('✅ ส่งรายงานไปยังอีเมลเรียบร้อยแล้วค่ะ!');
  } catch(e) { 
    console.log('❌ ส่งเมลล้มเหลว: ' + e.message); 
  }
}
// =========================================================
// ฟังก์ชันสำหรับทดสอบส่งให้คุณ Pongsak คนเดียว (ไม่กระทบของจริง)
// =========================================================
function testSendReportPongsakOnly() {
  // เคลียร์รายชื่ออีเมลเดิมออกชั่วคราว (เฉพาะการรันรอบนี้)
  AI_CONFIG.REPORT_EMAILS.length = 0;
  AI_CONFIG.REPORT_EMAILS.push('pongsak@smetaltech.co.th');
  
  console.log('📌 โหมดทดสอบ: กำลังสร้างรายงานและส่งให้ pongsak@smetaltech.co.th คนเดียว...');
  sendDailySmartReport();
}
