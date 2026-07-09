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
  // 1. ดึงคีย์จาก Script Properties เพื่อความปลอดภัย ไม่ให้คีย์หลุดไปบน GitHub
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
  console.log('📌 ขั้นตอนที่ 1: เริ่มต้นดึงข้อมูลจาก Google Sheet...');
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET2_NAME);
  if (!sheet) {
    console.log('❌ ข้อผิดพลาด: ไม่พบหน้า Sheet ที่ระบุค่ะ เช็คชื่อ Sheet2 อีกครั้งนะคะ');
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    console.log('❌ ข้อผิดพลาด: ตารางว่างเปล่า ไม่มีข้อมูลเลยค่ะ');
    return;
  }

  // ไนท์ปรับให้ดึงข้อมูลของ "เมื่อวาน" เสมอ ไม่ว่าจะรันสคริปต์ตอนกี่โมงก็ตามนะคะ
  const reportDate = new Date();
  reportDate.setDate(reportDate.getDate() - 1);
  
  const timeZone = Session.getScriptTimeZone();
  const targetDateString = Utilities.formatDate(reportDate, timeZone, "dd/MM/yyyy");
  
  let todayLogText = '';
  let matchCount = 0;

  // ตัวแปรสำหรับบวกเลขคำนวณ (เพื่อความเป๊ะ 100%)
  let totalReceiving = 0;
  let totalProduced = 0;
  let totalShipped = 0;
  let empStats = {};

// ==========================================
  // 🌟 ส่วนที่ 1: ดึงข้อมูลราคาเตรียมไว้คำนวณยอดขาย
  // ==========================================
  const INV_SHEET_ID = '1Aq1ZvwqKVKDQGIynEILrFIEe6A-sUqFk9Ztxm6SrNPo';
  const invSs = SpreadsheetApp.openById(INV_SHEET_ID);
  
  const priceSheet = invSs.getSheetByName('price');
  const priceData = priceSheet.getDataRange().getValues();
  let priceMap = {};
  for (let p = 1; p < priceData.length; p++) {
    let itemCode = String(priceData[p][0]).trim();
    let price = Number(priceData[p][4]) || 0; // คอลัมน์ E (Index 4) คือ Price
    if (itemCode !== "") priceMap[itemCode] = price;
  }

  let dailyIncome = 0;
  let accumulateIncome = 0;
  // ==========================================

  console.log(`📌 ขั้นตอนที่ 2: กำลังค้นหาข้อมูลการผลิตของวันที่ ${targetDateString}`);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[11] instanceof Date) {
      const timestampString = Utilities.formatDate(row[11], timeZone, "dd/MM/yyyy");
      
      if (timestampString === targetDateString) {
        const process = row[5];
        const menuStatus = row[10]; 
        let isValidData = false;

        // เช็คเงื่อนไขตาม Process และ Menu
        if (process === 'Receiving' && menuStatus === 'เริ่มการทำงาน') {
          isValidData = true;
        } else if (['Cutting', 'Milling', 'CNC', 'Grinding', 'Finished Goods'].includes(process) && menuStatus === 'เสร็จการทำงาน') {
          isValidData = true;
        } else if (process === 'Ship' && menuStatus === 'ส่งสินค้า') {
          isValidData = true;
        }

        if (isValidData) {
           const goodQty = (row[6] === '' || row[6] === undefined) ? 0 : Number(row[6]);
           const ngQty = (row[7] === '' || row[7] === undefined) ? 0 : Number(row[7]);
           const employeeName = (row[9] === '' || row[9] === undefined) ? '-' : String(row[9]).trim();

           todayLogText += `- Item: ${row[1]} (Name Part: ${row[2]}), Process: ${row[5]}, งานดี: ${goodQty}, NG: ${ngQty}, พนักงาน: ${employeeName}\n`;
           matchCount++;

           // --- สคริปต์บวกยอดแยกระบบ ---
           if (process === 'Receiving') {
             totalReceiving += goodQty;
           } else if (process === 'Ship') {
             totalShipped += goodQty;
             
             // 🌟 ไนท์เพิ่มส่วนนี้ค่ะ: เอาจำนวนที่ส่ง (goodQty) * ราคาต่อชิ้น
             let itemCode = String(row[1]).trim();
             let itemPrice = priceMap[itemCode] || 0;
             dailyIncome += (goodQty * itemPrice);
             
           } else {
             totalProduced += goodQty;
           }

           // --- เก็บสถิติพนักงานเพื่อหาพนักงานยอดเยี่ยม/ต้องระวัง ---
           if (employeeName !== '-') {
             if (!empStats[employeeName]) empStats[employeeName] = { good: 0, ng: 0 };
             empStats[employeeName].good += goodQty;
             empStats[employeeName].ng += ngQty;
           }
        }
      }
    }
  }

  if (todayLogText === '') {
    console.log('⚠️ แจ้งเตือน: ไม่พบข้อมูลที่ตรงกับเงื่อนไขของวันนี้เลยค่ะ');
    return;
  }

  console.log(`📌 ขั้นตอนที่ 3: พบข้อมูลทั้งหมด ${matchCount} รายการ กำลังเตรียมข้อมูลตัวเลขให้ AI...`);
// ==========================================
  // 🌟 ส่วนที่ 3: บันทึกยอดขายรายวัน & คำนวณยอดสะสม (เวอร์ชันแสนรู้: เพิ่มวันที่ให้อัตโนมัติ)
  // ==========================================
  console.log(`📌 กำลังบันทึกยอดขายและคำนวณยอดสะสมลง Sheet sumprice...`);
  const sumpriceSheet = invSs.getSheetByName('sumprice');
  let sumpriceData = sumpriceSheet.getDataRange().getValues();

  let targetRowIndex = -1;
  
  // 1. ค้นหาแถวของวันที่ทำรายงานก่อนค่ะ
  for (let r = 1; r < sumpriceData.length; r++) {
    let cellDate = sumpriceData[r][0];
    if (cellDate instanceof Date) {
      let sheetDateStr = Utilities.formatDate(cellDate, timeZone, "dd/MM/yyyy");
      if (sheetDateStr === targetDateString) {
        targetRowIndex = r + 1; 
        break;
      }
    }
  }

  // 🌟 ถ้าหาไม่เจอ ไนท์จะทำการเพิ่มแถวใหม่ให้พี่ทันทีค่ะ!
  if (targetRowIndex === -1) {
    console.log(`✨ ไม่พบวันที่ ${targetDateString} ในตาราง ไนท์กำลังเพิ่มแถวใหม่ให้พี่นะคะ...`);
    sumpriceSheet.appendRow([reportDate, 0, 0]); // เพิ่มวันที่, ยอดขาย 0, ยอดสะสม 0
    
    // จัดรูปแบบวันที่ในช่องใหม่ให้สวยงาม (d/M/yyyy)
    let lastRow = sumpriceSheet.getLastRow();
    sumpriceSheet.getRange(lastRow, 1).setNumberFormat('d/M/yyyy');
    
    // สั่งให้เรียงลำดับข้อมูลตามวันที่ (Column A) เผื่อมีการรันย้อนหลัง ข้อมูลจะได้ไม่กระโดดค่ะ
    const rangeToSort = sumpriceSheet.getRange(2, 1, sumpriceSheet.getLastRow() - 1, 3);
    rangeToSort.sort({column: 1, ascending: true});
    
    // หลังจากเพิ่มและเรียงแล้ว ให้หาแถวใหม่อีกครั้งเพื่อความเป๊ะค่ะ
    const refreshedData = sumpriceSheet.getDataRange().getValues();
    for (let r = 1; r < refreshedData.length; r++) {
      let cellDate = refreshedData[r][0];
      if (cellDate instanceof Date) {
        let sheetDateStr = Utilities.formatDate(cellDate, timeZone, "dd/MM/yyyy");
        if (sheetDateStr === targetDateString) {
          targetRowIndex = r + 1;
          break;
        }
      }
    }
  }

  // 2. เมื่อได้แถวที่ถูกต้องแล้ว (ไม่ว่าจะหาเจอเดิมหรือเพิ่มใหม่) ก็ทำการบันทึกยอดเลยค่ะ
  if (targetRowIndex !== -1) {
    sumpriceSheet.getRange(targetRowIndex, 2).setValue(dailyIncome);

    // 3. คำนวณยอดสะสมใหม่ของเดือนนี้
    let currentMonth = reportDate.getMonth();
    let currentYear = reportDate.getFullYear();
    let sumMonth = 0;

    const updatedData = sumpriceSheet.getRange(2, 1, sumpriceSheet.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < updatedData.length; i++) {
      let rDate = updatedData[i][0];
      if (rDate instanceof Date && rDate.getMonth() === currentMonth && rDate.getFullYear() === currentYear) {
        let rIncome = Number(updatedData[i][1]) || 0;
        sumMonth += rIncome;
        
        sumpriceSheet.getRange(i + 2, 3).setValue(sumMonth); 
        
        if (i + 2 === targetRowIndex) {
          accumulateIncome = sumMonth; 
        }
      }
    }
  }
  // ==========================================
  // --- สรุปหาคนเก่งและคนที่ต้องระวัง ---
  let bestEmp = { name: '-', good: -1 };
  let maxNg = 0;
  let badEmps = [];

  for (let name in empStats) {
    let st = empStats[name];
    if (st.ng === 0 && st.good > bestEmp.good) {
      bestEmp = { name: name, good: st.good };
    }
    if (st.ng > 0) {
      if (st.ng > maxNg) {
        maxNg = st.ng;
        badEmps = [name];
      } else if (st.ng === maxNg) {
        badEmps.push(name);
      }
    }
  }

  let watchEmployeeText = '';
  if (maxNg > 0) {
    watchEmployeeText = `${badEmps.join(', ')} มียอดของเสียรวม ${maxNg} ชิ้น`;
  } else {
    watchEmployeeText = 'ไม่มีพนักงานที่ทำของเสียเลยในวันนี้ค่ะ ยอดเยี่ยมมาก!';
  }

  if (bestEmp.name === '-') {
    bestEmp.name = 'ไม่มีพนักงานที่เข้าเงื่อนไข (ทุกคนมีของเสีย)';
    bestEmp.good = 0;
  }

  // --- กำหนด Prompt โดยโยนตัวเลขที่แม่นยำของเราให้ AI พูดตาม ---
  const prompt = `นี่คือข้อมูลการผลิตประจำวันที่ ${targetDateString}:\n${todayLogText}\n\n
  หน้าที่ของคุณ: สรุปข้อมูลเป็นภาษาไทยให้ผู้บริหาร โดยส่งผลลัพธ์เป็นโครงสร้าง HTML\n
  กำหนดให้มีหัวข้อใหญ่ด้านบนสุดของรายงานเขียนว่า: "<h2 style="font-size: 18px; margin-bottom: 12px; margin-top: 0;">สรุปข้อมูลการผลิต Part MA ประจำวันที่ ${targetDateString}</h2>"\n
  **ข้อควรระวังสำคัญ:** ห้ามใช้ <link> หรือ <style> เด็ดขาด ให้ใช้การตกแต่งแบบ Inline CSS (style="...") ในทุก Tag เท่านั้น (ใช้โทนสีและสไตล์คล้ายคลึงกับ Tailwind CSS)\n\n
  การจัดรูปแบบ UI ให้สวยงาม ทันสมัย เป็นระเบียบ และรองรับหน้าจอมือถือ:\n
  1. ให้สร้าง <div style="width: 100%;"> คลุมตารางเอาไว้ และสร้าง <table style="width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 11px; color: #374151; font-family: Tahoma, sans-serif;"> **เพียง 1 ตารางหลักเท่านั้น**\n
  2. หัวตาราง (th) ให้อยู่แถวแรกสุด มี 4 คอลัมน์ และปรับพื้นที่ให้สวยงาม ดังนี้:\n
      - <th style="width: 35%; background-color: #f3f4f6; padding: 10px 8px; border-bottom: 2px solid #e5e7eb; text-align: left; font-weight: 600; font-size: 11px;">Item (Name Part)</th>\n
      - <th style="width: 15%; background-color: #f3f4f6; padding: 10px 8px; border-bottom: 2px solid #e5e7eb; text-align: center; font-weight: 600; font-size: 11px; white-space: nowrap;">งานดี</th>\n
      - <th style="width: 15%; background-color: #f3f4f6; padding: 10px 8px; border-bottom: 2px solid #e5e7eb; text-align: center; font-weight: 600; font-size: 11px; white-space: nowrap;">NG</th>\n
      - <th style="width: 35%; background-color: #f3f4f6; padding: 10px 8px; border-bottom: 2px solid #e5e7eb; text-align: left; font-weight: 600; font-size: 11px;">พนักงาน</th>\n
  3. **การแบ่งกลุ่ม Process:** ห้ามสร้างตารางใหม่ ให้ใช้แทรกแถวคั่นกลางตาราง โดยใช้ <tr style="background-color: #e5e7eb;"><td colspan="4" style="padding: 10px 8px; font-weight: bold; font-size: 14px; color: #1f2937;">Process: [ชื่อ Process]</td></tr>\n
  4. **สำคัญมาก: ข้อมูลในตาราง (td) บังคับว่า 1 แถว (tr) จะต้องมี 4 คอลัมน์ (td) เสมอ** ได้แก่ Item, งานดี, NG, และพนักงาน ห้ามข้ามหรือยุบรวมคอลัมน์เด็ดขาด! (ต่อให้ข้อมูล NG จะเป็น 0 ก็ต้องสร้าง <td> ให้มันด้วย) โดยจัดให้อยู่ตำแหน่งเดียวกับหัวตาราง และใช้ (style="padding: 10px 8px; border-bottom: 1px solid #e5e7eb; word-break: break-word; font-size: 11px;")\n\n
  เงื่อนไขการจัดเนื้อหา:\n
  1. ระบุ "Name Part" ต่อท้ายชื่อ Item เสมอ\n
  2. **การเรียงลำดับ Process:** ให้แสดงผลโดยเรียงลำดับ Process ตามนี้อย่างเคร่งครัด (ถ้า Process ไหนไม่มีงานในวันนั้น ไม่ต้องโชว์แถวคั่นของ Process นั้น):\n
      ลำดับที่ 1: Process Receiving\n
      ลำดับที่ 2: Process Cutting\n
      ลำดับที่ 3: Process Milling\n
      ลำดับที่ 4: Process CNC\n
      ลำดับที่ 5: Process Grinding\n
      ลำดับที่ 6: Process Finished Goods\n
      ลำดับที่ 7: Process Ship\n
  3. ส่วนท้าย "สรุปภาพรวมประจำวัน" ให้อยู่นอกตารางหลัก โดยจัดรูปแบบให้อ่านง่ายบนจอมือถือ ดังนี้:\n
      - ให้นำเนื้อหาสรุปทั้งหมดไปใส่ในกรอบ <div style="background-color: #f8fafc; padding: 16px; border-radius: 8px; margin-top: 20px; border: 1px solid #e2e8f0; font-family: Tahoma, sans-serif; font-size: 13px;">\n
      - หัวข้อให้ใช้ <h3 style="margin-top: 0; margin-bottom: 12px; font-size: 15px; color: #1e293b;">สรุปภาพรวมประจำวัน</h3>\n
      - ข้อมูลสรุปแต่ละข้อ **บังคับให้ใช้แท็ก <ul style="margin: 0; padding-left: 20px;"> และ <li style="margin-bottom: 8px; line-height: 1.5; word-wrap: break-word;">...</li>** เพื่อที่เวลาข้อความยาวจนตกบรรทัด มันจะจัดย่อหน้าให้เป็นระเบียบ\n
  4. ใช้สีเขียวสำหรับตัวเลข NG เป็น 0 และใช้สีแดงสำหรับรายการที่มี NG\n\n
  ⚠️ **ข้อมูลส่วนสรุปภาพรวมประจำวัน (สำคัญมาก: ห้ามบวกเลขเอง ให้คัดลอกข้อความด้านล่างนี้ไปใส่ใน <li> ได้เลยเพื่อความถูกต้อง 100%)**:\n
  - <span style="font-weight: bold; color: #000;">ยอดรับ Raw Material (Process: Receiving): ${totalReceiving} ชิ้น</span>\n
  - <span style="font-weight: bold; color: #000;">ยอดรวมการผลิตสินค้าสำเร็จรูปประจำวัน: ${totalProduced} ชิ้น</span>\n
  - <span style="font-weight: bold; color: #000;">ยอดรวมสินค้าที่จัดส่ง (Process: Ship): ${totalShipped} ชิ้น</span>\n
  - <span style="font-weight: bold; color: #000;">ยอดขายประจำวัน: ${dailyIncome.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})} บาท</span>\n
  - <span style="font-weight: bold; color: #000;">ยอดขายรวม (Total Income): ${accumulateIncome.toLocaleString('th-TH', {minimumFractionDigits: 2, maximumFractionDigits: 2})} บาท</span>\n
  - พนักงานยอดเยี่ยม: <span style="font-weight: bold; color: #1d4ed8;">${bestEmp.name} ด้วยยอดงานดีรวม ${bestEmp.good} ชิ้น</span>\n
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