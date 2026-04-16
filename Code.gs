// ทดสอบอัปเดตโค้ดจาก VS Code
/**
 * Media EDU Club System - Backend (code.gs)
 * Features: 24h Thai Timestamp, File Uploads, 3-in-1 Tracking (Latest First)
 * แก้ไข: LINE OA integration, fix saveBooking dead code, fix notifyUserResult variable
 */

// ==========================================
// 1. การตั้งค่าเริ่มต้น & Helper Functions
// ==========================================

// ===== Spreadsheet Config =====
// เปลี่ยน ID ตรงนี้เมื่อต้องการเชื่อมกับฐานข้อมูลใหม่
var SPREADSHEET_ID = '1XdsJ9Ss4OTlEPxYx2t2bga9y70OxPSxIab_XHabbOSU';

/**
 * ดึง Spreadsheet ที่กำหนดไว้ใน SPREADSHEET_ID
 * ใช้แทน getDB() ทุกจุด
 */
function getDB() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ===== LINE OA Config =====
var LINE_TOKEN = '6wwUDIYGaDUdvZ5426QRojLeuD+IF8YZGPmU7Dfr5qIyBIEEFi+IhAsmAZLTZBQWvQJ5m1otfGyOoTQcTF7pdfZctMg3QeMuT/wZdwBSuIrQHUCnKzyihceOLT0RdrNMXfnh2+KbK8szMfDUV3FQhAdB04t89/1O/w1cDnyilFU=';

// ==========================================
// LINE Admin IDs — อ่าน/เขียนจาก PropertiesService
// (ไม่ต้องแก้ code.gs โดยตรงอีกต่อไป)
// ==========================================
function getLineAdminIdsFromProps() {
  var raw = PropertiesService.getScriptProperties().getProperty('LINE_ADMIN_IDS') || '';
  if (!raw) {
    var fallback = ['U8efbc8f350bbecec7e49b59f7562c5b7'];
    PropertiesService.getScriptProperties().setProperty('LINE_ADMIN_IDS', JSON.stringify(fallback));
    return fallback;
  }
  try {
    var parsed = JSON.parse(raw);
    // รองรับทั้ง array เดิม (string[]) และแบบใหม่ (object[])
    if (Array.isArray(parsed) && parsed.length > 0 && typeof parsed[0] === 'string') {
      // migrate รูปแบบเก่า → ใหม่
      var migrated = parsed.map(function(uid) { return { uid: uid, scope: 'both' }; });
      PropertiesService.getScriptProperties().setProperty('LINE_ADMIN_IDS', JSON.stringify(migrated));
      return parsed; // คืน array ของ uid เพื่อใช้กับ notifyAdmins เดิม
    }
    return parsed.map(function(a) { return a.uid || a; });
  } catch(e) { return []; }
}

// ดึง LINE Admins พร้อม scope
function getLineAdminsFull() {
  var raw = PropertiesService.getScriptProperties().getProperty('LINE_ADMIN_IDS') || '';
  var list = [];
  try {
    var parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      list = parsed.map(function(a) {
        return typeof a === 'string' ? { uid: a, scope: 'both' } : a;
      });
    }
  } catch(e) {}
  // เติมชื่อจาก LINE_Users sheet (ใช้ getLineUsersData เพื่อรองรับ header migration)
  var nameMap = {};
  try {
    var luRows = getLineUsersData();
    for (var i = 1; i < luRows.length; i++) {
      nameMap[String(luRows[i][1]||'').trim()] = String(luRows[i][2]||'').trim();
    }
  } catch(e) {}
  return list.map(function(a) {
    return { uid: a.uid, scope: a.scope || 'both', name: nameMap[a.uid] || '' };
  });
}

// notifyAdmins กรอง scope
function notifyAdminsByScope(message, scope) {
  // scope = 'media' | 'samo' | 'both' | null (null = ส่งทุกคน)
  var list = getLineAdminsFull();
  list.forEach(function(a) {
    var adminScope = a.scope || 'both';
    if (!scope || adminScope === 'both' || adminScope === scope) {
      sendLineMessage(a.uid, message);
    }
  });
}

// ดึงรายชื่อ LINE Users ทั้งหมดสำหรับ picker (UID + ชื่อ)
function getAllLineUsersForPicker() {
  var rows   = getLineUsersData();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var sid  = String(rows[i][0] || '').trim();
    var uid  = String(rows[i][1] || '').trim();
    var name = String(rows[i][2] || '').trim();
    if (uid) result.push({ sid: sid, uid: uid, name: name });
  }
  return result;
}

// ดึงรายชื่อ LINE Admin สำหรับหน้า Admin UI
function getLineAdmins() {
  return getLineAdminsFull();
}

// เพิ่ม LINE Admin
function addLineAdmin(uid, scope) {
  uid = String(uid || '').trim();
  if (!uid || uid.charAt(0) !== 'U' || uid.length < 10)
    return { ok: false, msg: 'UID ไม่ถูกต้อง (ต้องขึ้นต้นด้วย U และมีความยาวอย่างน้อย 10 ตัวอักษร)' };
  var list = getLineAdminsFull();
  if (list.find(function(a){ return a.uid === uid; }))
    return { ok: false, msg: 'UID นี้มีอยู่แล้วในรายการ' };
  list.push({ uid: uid, scope: scope || 'both' });
  PropertiesService.getScriptProperties().setProperty('LINE_ADMIN_IDS', JSON.stringify(list));
  return { ok: true, msg: 'เพิ่มสำเร็จ', admins: getLineAdmins() };
}

// ลบ LINE Admin
function removeLineAdmin(uid) {
  var list = getLineAdminsFull();
  var newList = list.filter(function(a){ return a.uid !== uid; });
  if (newList.length === list.length) return { ok: false, msg: 'ไม่พบ UID นี้' };
  PropertiesService.getScriptProperties().setProperty('LINE_ADMIN_IDS', JSON.stringify(newList));
  return { ok: true, msg: 'ลบสำเร็จ', admins: getLineAdmins() };
}

// อัปเดต scope ของ LINE Admin
function updateLineAdminScope(uid, scope) {
  var list = getLineAdminsFull();
  var found = false;
  list = list.map(function(a) {
    if (a.uid === uid) { found = true; return { uid: a.uid, scope: scope }; }
    return a;
  });
  if (!found) return { ok: false, msg: 'ไม่พบ UID นี้' };
  PropertiesService.getScriptProperties().setProperty('LINE_ADMIN_IDS', JSON.stringify(list));
  return { ok: true, admins: getLineAdmins() };
}

// ทดสอบส่ง LINE ให้ admin คนนั้น
function testSendLineToAdmin(uid) {
  try {
    sendLineMessage(uid, '✅ ทดสอบระบบแจ้งเตือน\nMedia EDU Club CMU — ตั้งค่า Admin สำเร็จแล้ว!');
    return { ok: true, msg: 'ส่ง LINE ทดสอบสำเร็จ' };
  } catch(e) {
    return { ok: false, msg: 'ส่งไม่สำเร็จ: ' + e.message };
  }
}

function doGet(e) {
  fixSheetColumnFormats();
  var output = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Media EDU Club System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

/**
 * ฟังก์ชัน include สำหรับ GAS Template
 * ใช้ใน Index.html ด้วย <?!= include('ชื่อไฟล์') ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// Web Settings — เปิด/ปิดเว็บ และระบบต่างๆ
// ==========================================
var WEB_SETTINGS_KEY = 'WEB_SETTINGS';

var DEFAULT_WEB_SETTINGS = {
  siteOpen:      true,   // เปิด/ปิดเว็บทั้งหมด
  closedMessage: 'ระบบปิดให้บริการชั่วคราว กรุณากลับมาใหม่ในภายหลัง',
  systems: {
    order:            true,   // ระบบสั่งซื้อสินค้า
    tracking:         true,   // ระบบติดตามสถานะ
    equipment:        true,   // ระบบยืมอุปกรณ์
    booking:          true,   // ระบบจองห้อง
    calendar:         true,   // ตารางการจองห้อง
    faq:              true,   // FAQ
    contact:          true,   // ติดต่อทีมงาน
    activityCalendar: true,   // ปฏิทินกิจกรรมสโมสร
    clubs:            true    // ลงทะเบียนชุมนุม/กีฬา
  }
};

/**
 * ดึงการตั้งค่าเว็บ โดยอ่านจาก Google Sheet เป็นหลัก
 * เพื่อให้เวลาแอดมินแก้ค่าใน Sheet เว็บไซต์จะเปลี่ยนตามทันที
 */
function getWebSettings() {
  try {
    var ss = getDB();
    var sheet = ss.getSheetByName('Web_Settings');
    var settings = JSON.parse(JSON.stringify(DEFAULT_WEB_SETTINGS)); // คัดลอกค่าเริ่มต้น

    // ถ้ายังไม่มี Sheet ให้อ่านจาก Properties ก่อน (ถ้ามี) หรือใช้ค่าเริ่มต้น
    if (!sheet) {
      var rawProps = PropertiesService.getScriptProperties().getProperty(WEB_SETTINGS_KEY);
      if (rawProps) return JSON.parse(rawProps);
      return settings;
    }

    // ดึงข้อมูลทั้งหมดใน Sheet มาตรวจสอบ
    var data = sheet.getDataRange().getValues();
    
    // ถ้ามีแค่หัวตาราง หรือไม่มีข้อมูล ให้ส่งค่า default กลับไป
    if (data.length > 1) {
      // วนลูปอ่านข้อมูลตั้งแต่บรรทัดที่ 2 เป็นต้นไป
      for (var i = 1; i < data.length; i++) {
        var systemName = String(data[i][0] || '').trim();
        var statusStr  = String(data[i][1] || '').trim();
        var message    = String(data[i][2] || '').trim();
        
        // เช็คว่ามีคำว่า "✅ เปิด" หรือไม่ (ถ้ามีถือว่า true, ถ้าเป็น "🔴 ปิด" ถือว่า false)
        var isOpen = (statusStr.indexOf('เปิด') !== -1);

        // จับคู่ชื่อระบบแล้วอัปเดตค่า
        if (systemName.indexOf('เว็บหลักทั้งหมด') !== -1) {
          settings.siteOpen = isOpen;
          if (message && message !== '-') settings.closedMessage = message;
        } 
        else if (systemName.indexOf('สั่งซื้อสินค้า') !== -1)        settings.systems.order = isOpen;
        else if (systemName.indexOf('ติดตามสถานะ') !== -1)           settings.systems.tracking = isOpen;
        else if (systemName.indexOf('ยืมอุปกรณ์') !== -1)            settings.systems.equipment = isOpen;
        else if (systemName.indexOf('จองห้อง') !== -1)               settings.systems.booking = isOpen;
        else if (systemName.indexOf('ตารางการจองห้อง') !== -1)       settings.systems.calendar = isOpen;
        else if (systemName.indexOf('คำถามที่พบบ่อย') !== -1)        settings.systems.faq = isOpen;
        else if (systemName.indexOf('ติดต่อทีมงาน') !== -1)          settings.systems.contact = isOpen;
        else if (systemName.indexOf('ปฏิทินกิจกรรม') !== -1)         settings.systems.activityCalendar = isOpen;
        else if (systemName.indexOf('ลงทะเบียนชุมนุม') !== -1)       settings.systems.clubs = isOpen;
      }
    }

    // บันทึกลง PropertiesService เพื่อเป็น Backup เผื่อกรณี Sheet มีปัญหา
    PropertiesService.getScriptProperties().setProperty(WEB_SETTINGS_KEY, JSON.stringify(settings));

    return settings;

  } catch(e) {
    Logger.log('getWebSettings error: ' + e.message);
    // กรณีเกิดข้อผิดพลาด ให้พยายามอ่านจาก PropertiesService เป็นตัวสำรอง
    var fallback = PropertiesService.getScriptProperties().getProperty(WEB_SETTINGS_KEY);
    if (fallback) return JSON.parse(fallback);
    return DEFAULT_WEB_SETTINGS;
  }
}

/**
 * บันทึกการตั้งค่า (ใช้เมื่อมีการสั่งบันทึกจากหน้าเว็บแอดมิน)
 */
function saveWebSettings(settings) {
  try {
    // 1. บันทึกลง PropertiesService
    PropertiesService.getScriptProperties().setProperty(WEB_SETTINGS_KEY, JSON.stringify(settings));

    // 2. บันทึกลง Google Sheet
    var ss = getDB();
    var sheetName = 'Web_Settings';
    var sheet = ss.getSheetByName(sheetName);

    // ถ้ายังไม่มีแท็บชีต ให้สร้างใหม่และทำหัวตาราง
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['ระบบ / ตั้งค่า', 'สถานะ', 'ข้อความแจ้งเตือนเมื่อปิด']);
      sheet.getRange(1, 1, 1, 3).setBackground('#1a3a6b').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setColumnWidth(1, 200);
      sheet.setColumnWidth(2, 100);
      sheet.setColumnWidth(3, 300);
    } else {
      // ล้างข้อมูลเก่าออก (ยกเว้นหัวตาราง)
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
      }
    }

    // เตรียมข้อมูลเพื่อเขียนลงชีต
    var rowsToSave = [
      ['เว็บหลักทั้งหมด (Site)', settings.siteOpen ? '✅ เปิด' : '🔴 ปิด', settings.closedMessage || '-'],
      ['สั่งซื้อสินค้า (Order)', settings.systems.order ? '✅ เปิด' : '🔴 ปิด', '-'],
      ['ติดตามสถานะ (Tracking)', settings.systems.tracking ? '✅ เปิด' : '🔴 ปิด', '-'],
      ['ยืมอุปกรณ์ (Equipment)', settings.systems.equipment ? '✅ เปิด' : '🔴 ปิด', '-'],
      ['จองห้อง (Booking)', settings.systems.booking ? '✅ เปิด' : '🔴 ปิด', '-'],
      ['ตารางการจองห้อง (Calendar)', settings.systems.calendar ? '✅ เปิด' : '🔴 ปิด', '-'],
      ['คำถามที่พบบ่อย (FAQ)', settings.systems.faq ? '✅ เปิด' : '🔴 ปิด', '-'],
      ['ติดต่อทีมงาน (Contact)', settings.systems.contact ? '✅ เปิด' : '🔴 ปิด', '-'],
      ['ปฏิทินกิจกรรมสโมสร (ActivityCalendar)', settings.systems.activityCalendar !== false ? '✅ เปิด' : '🔴 ปิด', '-'],
      ['ลงทะเบียนชุมนุม/กีฬา (Clubs)', settings.systems.clubs !== false ? '✅ เปิด' : '🔴 ปิด', '-']
    ];

    // เขียนข้อมูลลงไปในคราวเดียว
    sheet.getRange(2, 1, rowsToSave.length, 3).setValues(rowsToSave);

    return { ok: true };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

/**
 * ตั้งค่า format คอลัมน์เบอร์โทร / เวลา / วันที่ ทุก sheet เป็น Plain Text (@)
 */
function fixSheetColumnFormats() {
  var ss = getDB();
  var TEXT = '@';

  var er = ss.getSheetByName('Equipment_Requests');
  if (er) {
    er.getRange(2, 5, Math.max(er.getLastRow(), 2), 1).setNumberFormat(TEXT);
    er.getRange(2, 7, Math.max(er.getLastRow(), 2), 1).setNumberFormat(TEXT);
    er.getRange(2, 8, Math.max(er.getLastRow(), 2), 1).setNumberFormat(TEXT);
  }

  var bk = ss.getSheetByName('Bookings');
  if (bk) {
    bk.getRange(2, 5,  Math.max(bk.getLastRow(), 2), 1).setNumberFormat(TEXT);
    bk.getRange(2, 10, Math.max(bk.getLastRow(), 2), 1).setNumberFormat(TEXT);
    bk.getRange(2, 11, Math.max(bk.getLastRow(), 2), 1).setNumberFormat(TEXT);
    bk.getRange(2, 12, Math.max(bk.getLastRow(), 2), 1).setNumberFormat(TEXT);
  }

  var od = ss.getSheetByName('Orders');
  if (od) {
    od.getRange(2, 2, Math.max(od.getLastRow(), 2), 1).setNumberFormat(TEXT);
  }

  var dl = ss.getSheetByName('Delivery_Config');
  if (dl) {
    dl.getRange(2, 4, Math.max(dl.getLastRow(), 2), 1).setNumberFormat(TEXT);
    dl.getRange(2, 5, Math.max(dl.getLastRow(), 2), 1).setNumberFormat(TEXT);
  }
}

function getThaiTimestamp() {
  var now      = new Date();
  var datePart = Utilities.formatDate(now, 'GMT+7', 'dd/MM/');
  var yearPart = now.getFullYear() + 543;
  var timePart = Utilities.formatDate(now, 'GMT+7', ' HH:mm:ss');
  return datePart + yearPart + timePart;
}

function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

// ==========================================
// 2. MODULE: ระบบสั่งซื้อสินค้า (Orders)
// ==========================================

// ==========================================
// Batch loader — รวม call เดียวลด latency
// ==========================================

// ==========================================
// CacheService Helpers
// ==========================================

var CACHE = CacheService.getScriptCache();

var CACHE_TTL = {
  products: 300, rooms: 600, shipping: 600,
  slides: 600,   faq: 600,   delivery: 600,
  inventory: 120, dashboard: 60, webSettings: 300,
};

function withCache(key, ttlSeconds, loaderFn) {
  var cached = CACHE.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }
  var data = loaderFn();
  try {
    var json = JSON.stringify(data);
    if (json.length < 100000) CACHE.put(key, json, ttlSeconds);
  } catch(e) { Logger.log('Cache write error (' + key + '): ' + e.message); }
  return data;
}

function bustCache() {
  CACHE.removeAll(['products','rooms','shipping','slides','faq','delivery','inventory','dashboard','webSettings']);
}

function bustCacheKey(key) { CACHE.remove(key); }

function getInitialData() {
  try {
    return {
      webSettings: getWebSettings(),
      products:    getProductList(),
      rooms:       getRoomList(),
      inventory:   getDynamicInventory(),
      slides:      getSlidesConfig()
    };
  } catch(e) {
    Logger.log('getInitialData error: ' + e.message);
    return {
      webSettings: DEFAULT_WEB_SETTINGS,
      products: [], rooms: [], inventory: {}, slides: []
    };
  }
}

function getProductList() {
  return withCache('products', CACHE_TTL.products, getProductList_raw_);
}

function getProductList_raw_() {
  var sheet = getDB().getSheetByName('Products');
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var products = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== '') {
      var stock = data[i][6];
      var stockNum = (stock !== '' && !isNaN(Number(stock))) ? Number(stock) : null;
      var activeVal = data[i][12];
      var isActive  = (activeVal === '' || activeVal === undefined ||
                       String(activeVal).toUpperCase() === 'TRUE');
      if (!isActive) continue; // ข้ามสินค้าที่ปิดขาย
      products.push({
        id:               data[i][0],
        name:             data[i][1],
        price:            parseFloat(data[i][2]),
        image:            data[i][3] || '',
        details:          data[i][4] || '',
        sizes:            data[i][5] || '',
        stock:            stockNum,
        promptpay:        data[i][7] || '',
        productType:      data[i][8] || 'normal',
        preorderDeadline: data[i][9] || '',
        specialSizes:     data[i][10] || '',
        qrUrl:            data[i][11] || '',
        active:           'TRUE',
        deadlineDate:     data[i][13] || '',
        deadlineTime:     data[i][14] || ''
      });
    }
  }
  return products;
}

function getAdminProducts() {
  var sheet = getDB().getSheetByName('Products');
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var out  = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === '') continue;
    var st   = data[i][6];
    var stNum = (st !== '' && !isNaN(Number(st))) ? Number(st) : null;
    out.push({
      row:              i + 1,
      id:               data[i][0],
      name:             data[i][1],
      price:            data[i][2],
      image:            data[i][3] || '',
      details:          data[i][4] || '',
      sizes:            data[i][5] || '',
      stock:            stNum,
      promptpay:        data[i][7] || '',
      productType:      data[i][8] || 'normal',
      preorderDeadline: data[i][9] || '',
      specialSizes:     data[i][10] || ''
    });
  }
  return out;
}

// อัปโหลด QR image base64 → Drive folder "Product_QR" → คืน public URL
function uploadQrToDrive(base64Data, fileName) {
  try {
    var folder   = getOrCreateFolder('Product_QR');
    var mime     = (base64Data.match(/^data:([^;]+);base64,/) || [])[1] || 'image/png';
    var pure     = base64Data.replace(/^data:[^;]+;base64,/, '');
    var blob     = Utilities.newBlob(Utilities.base64Decode(pure), mime, fileName);
    var file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch(e) {
    Logger.log('uploadQrToDrive: ' + e.message);
    return '';
  }
}

function saveProduct(id, name, price, imgType, imgData, desc, sizes, stock, promptpay, productType, preorderDeadline, specialSizes, ppMode, qrData) {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Products');
  if (!sheet) {
    sheet = ss.insertSheet('Products');
    sheet.appendRow(['รหัส','ชื่อสินค้า','ราคา','รูปภาพ','รายละเอียด','ขนาด','สต็อก','PromptPay','ประเภท','วันปิดรับ','ไซส์พิเศษ','QR_URL','Active','DeadlineDate','DeadlineTime']);
  }
  // จัดการ QR image
  var finalQr = '';
  if (ppMode === 'qrfile' || ppMode === 'qrpaste') {
    if (qrData && qrData.indexOf('data:image') === 0) {
      finalQr = uploadQrToDrive(qrData, 'qr_' + id + '_' + Date.now() + '.png');
    }
  } else if (ppMode === 'qrurl') {
    finalQr = qrData || '';
  }
  sheet.appendRow([
    id, name, Number(price), imgData || '', desc || '', sizes || '',
    stock !== undefined && stock !== '' ? Number(stock) : '',
    promptpay || '',
    productType || 'normal',
    preorderDeadline || '',
    specialSizes || '',
    finalQr,
    (activeVal !== undefined && activeVal !== '') ? activeVal : 'TRUE',
    deadlineDate || '',
    deadlineTime || '23:59'
  ]);
}

function updateProduct(rowNum, id, name, price, imgType, imgData, desc, sizes, stock, promptpay, productType, preorderDeadline, specialSizes, ppMode, qrData, activeVal, deadlineDate, deadlineTime) {
  var sheet = getDB().getSheetByName('Products');
  if (!sheet || !rowNum) return false;
  // ถ้ามี QR image ใหม่ → upload
  var finalQr = '';
  if (ppMode === 'qrfile' || ppMode === 'qrpaste') {
    if (qrData && qrData.indexOf('data:image') === 0) {
      finalQr = uploadQrToDrive(qrData, 'qr_' + id + '_' + Date.now() + '.png');
    } else {
      // ใช้ค่าเดิม (qrData อาจเป็น URL เดิม)
      finalQr = qrData && qrData.indexOf('http') === 0 ? qrData : '';
    }
  } else if (ppMode === 'qrurl') {
    finalQr = qrData || '';
  }
  var stockVal = (stock !== undefined && stock !== '') ? Number(stock) : '';
  sheet.getRange(rowNum, 1, 1, 15).setValues([[
    id, name, Number(price), imgData || '', desc || '', sizes || '',
    stockVal,
    promptpay || '',
    productType || 'normal',
    preorderDeadline || '',
    specialSizes || '',
    finalQr,
    (activeVal !== undefined && activeVal !== '') ? activeVal : 'TRUE',
    deadlineDate || '',
    deadlineTime || '23:59'
  ]]);
  return true;
}

// toggle เปิด/ปิดขายสินค้า (col 13 = Active)
function toggleProductActive(prodId, newActive) {
  var sheet = getDB().getSheetByName('Products');
  if (!sheet) throw new Error('ไม่พบ Products sheet');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(prodId).trim()) {
      sheet.getRange(i + 1, 13).setValue(newActive ? 'TRUE' : 'FALSE');
      try { CacheService.getScriptCache().remove('products'); } catch(e) {}
      return { ok: true };
    }
  }
  throw new Error('ไม่พบสินค้า: ' + prodId);
}

// อัปเดต stock ของสินค้า (ลดลงตามจำนวนที่สั่ง)
function decrementProductStock(productId, qty) {
  var sheet = getDB().getSheetByName('Products');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(productId)) {
      var currentStock = data[i][6];
      if (currentStock !== '' && !isNaN(Number(currentStock))) {
        var newStock = Math.max(0, Number(currentStock) - qty);
        sheet.getRange(i + 1, 7).setValue(newStock);
      }
      break;
    }
  }
}

// อัปโหลดไฟล์ base64 ไปยัง Google Drive (folder: Order_Slips)
function uploadSlipToDrive(base64Data, fileName) {
  try {
    var folder = getOrCreateFolder('Order_Slips');
    var mimeMatch = base64Data.match(/^data:([^;]+);base64,/);
    var mimeType  = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    var pureBase64 = base64Data.replace(/^data:[^;]+;base64,/, '');
    var blob = Utilities.newBlob(Utilities.base64Decode(pureBase64), mimeType, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/file/d/' + file.getId() + '/view';
  } catch(e) {
    Logger.log('uploadSlipToDrive error: ' + e.message);
    return '';
  }
}

// ==========================================
// ระบบจัดการช่องทางจัดส่ง (ShippingOptions)
// ==========================================

function getShippingOptions() {
  return withCache('shipping', CACHE_TTL.shipping, getShippingOptions_raw_);
}

function getShippingOptions_raw_() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('ShippingOptions');
  if (!sheet) return [];
  var data    = sheet.getDataRange().getDisplayValues();
  var options = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== '') {
      options.push({
        row:            i + 1,
        id:             data[i][0],
        name:           data[i][1],
        fee:            parseFloat(data[i][2]) || 0,
        requireAddress: data[i][3] === 'TRUE' || data[i][3] === 'true' || data[i][3] === '1',
        active:         data[i][4] !== 'FALSE' && data[i][4] !== 'false' && data[i][4] !== '0',
        note:           data[i][5] || ''
      });
    }
  }
  return options.filter(function(o) { return o.active; });
}

function getShippingOptionsAdmin() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('ShippingOptions');
  if (!sheet) return [];
  var data    = sheet.getDataRange().getDisplayValues();
  var options = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== '') {
      options.push({
        row:            i + 1,
        id:             data[i][0],
        name:           data[i][1],
        fee:            data[i][2],
        requireAddress: data[i][3],
        active:         data[i][4],
        note:           data[i][5] || ''
      });
    }
  }
  return options;
}

function saveShippingOption(rowNum, id, name, fee, requireAddress, active, note) {
  var ss    = getDB();
  var sheet = ss.getSheetByName('ShippingOptions');
  if (!sheet) {
    sheet = ss.insertSheet('ShippingOptions');
    sheet.appendRow(['รหัส','ชื่อช่องทาง','ค่าจัดส่ง','ต้องการที่อยู่','เปิดใช้','หมายเหตุ']);
  }
  var row = [id, name, parseFloat(fee) || 0, requireAddress ? 'TRUE' : 'FALSE', active ? 'TRUE' : 'FALSE', note || ''];
  if (rowNum && rowNum > 1) {
    sheet.getRange(rowNum, 1, 1, 6).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return true;
}

function deleteShippingOption(rowNum) {
  var sheet = getDB().getSheetByName('ShippingOptions');
  if (!sheet || !rowNum) return false;
  sheet.deleteRow(rowNum);
  return true;
}

function placeOrder(customerData, cartItems) {
  var sheet     = getDB().getSheetByName('Orders');
  var orderId   = generateNextId('Orders', 'edushop');
  var timestamp = getThaiTimestamp();

  var orderDetails = cartItems.map(function(item) {
    return item.name + ' (x' + item.qty + ')';
  }).join('\n');

  var itemsTotal = cartItems.reduce(function(sum, item) {
    return sum + (item.price * item.qty);
  }, 0);
  var shippingFee = parseFloat(customerData.shippingFee) || 0;
  var totalPrice  = itemsTotal + shippingFee;

  // อัปโหลดสลิปไปยัง Drive
  var slipUrl = '';
  if (customerData.slipBase64 && customerData.slipBase64.length > 100) {
    slipUrl = uploadSlipToDrive(customerData.slipBase64, 'slip_' + orderId + '.jpg');
  }

  // สร้าง buyerInfo สรุป
  var buyerType = customerData.buyerType || 'นักศึกษาปัจจุบัน';
  var buyerInfo = customerData.buyerInfoStr || '';

  var newRowOd = sheet.getLastRow() + 1;
  sheet.getRange(newRowOd, 2).setNumberFormat('@');
  sheet.appendRow([
    orderId,
    customerData.phone,
    customerData.name,
    orderDetails,
    totalPrice,
    'รอตรวจสอบ',
    timestamp,
    '',                              // col 8: approver
    customerData.email || '',        // col 9
    customerData.address || '',      // col 10
    slipUrl,                         // col 11
    buyerType,                       // col 12
    buyerInfo,                       // col 13
    customerData.shippingName || '', // col 14
    shippingFee,                     // col 15
    customerData.contact || ''       // col 16: ช่องทางติดต่อ
  ]);

  // ลด stock ของแต่ละสินค้า
  cartItems.forEach(function(item) {
    if (item.id) decrementProductStock(item.id, item.qty);
  });

  if (customerData.email) {
    var subject = '[Media EDU Club] ยืนยันการสั่งซื้อ รหัส ' + orderId;
    var slipHtml = slipUrl
      ? '<div style="margin-top:12px;"><a href="' + slipUrl + '" target="_blank" style="color:#0369a1;">📎 ดูสลิปที่แนบ</a></div>'
      : '';
    var shippingRow = shippingFee > 0
      ? ['ค่าจัดส่ง', shippingFee.toLocaleString() + ' บาท']
      : null;
    var emailRows = [
      ['รหัสออเดอร์',     orderId],
      ['ประเภทผู้สั่ง',    buyerType],
      ['ชื่อ-นามสกุล',    customerData.name],
      ['เบอร์โทรศัพท์',   customerData.phone],
      ['ช่องทางติดต่อ',   customerData.contact || '-'],
      ['รายการสินค้า',    orderDetails.replace(/\n/g, '<br>')],
      ['ราคาสินค้า',      itemsTotal.toLocaleString() + ' บาท'],
      ['ช่องทางจัดส่ง',   customerData.shippingName || '-']
    ];
    if (shippingRow) emailRows.push(shippingRow);
    if (customerData.address) emailRows.push(['ที่อยู่จัดส่ง', customerData.address]);
    emailRows.push(['ยอดรวมทั้งหมด', totalPrice.toLocaleString() + ' บาท']);
    emailRows.push(['หมายเหตุ', customerData.note || '-']);
    emailRows.push(['สถานะ', 'รอตรวจสอบ']);
    emailRows.push(['วันเวลาที่สั่ง', timestamp]);

    var body = buildEmailHtml({
      title:      'ยืนยันการสั่งซื้อสินค้า',
      headerIcon: 'S',
      color:      '#1565c0',
      name:       customerData.name,
      intro:      'ระบบได้รับคำสั่งซื้อของคุณเรียบร้อยแล้ว รายละเอียดมีดังนี้:',
      rows:       emailRows,
      extraHtml:  slipHtml,
      note: 'ทีมงานจะตรวจสอบสลิปและแจ้งสถานะให้ทราบผ่านอีเมลนี้'
    });
    sendConfirmEmail(customerData.email, subject, body);
  }

  // แจ้งเตือน LINE Admin
  notifyNewOrder(customerData, orderId, orderDetails, itemsTotal, shippingFee, totalPrice);

  // แจ้งเตือน LINE ลูกค้า (ถ้ามี LINE User ID ใน LINE_Users)
  notifyOrderToCustomer(customerData, orderId, orderDetails, itemsTotal, shippingFee, totalPrice);

  return { orderId: orderId, slipUrl: slipUrl };
}

/** หา LINE User ID จาก phone หรือ email */
function findCustomerLineId(phone, email) {
  try {
    var data = getLineUsersData();
    phone = String(phone || '').replace(/[^0-9]/g, '');
    email = String(email || '').toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowPhone = String(row[5] || '').replace(/[^0-9]/g, '');
      var rowEmail = String(row[7] || '').toLowerCase().trim();
      if ((phone && rowPhone === phone) || (email && rowEmail === email)) {
        return String(row[1] || '').trim(); // LINE User ID
      }
    }
  } catch(e) { Logger.log('findCustomerLineId error: ' + e.message); }
  return null;
}

/** ส่ง LINE + อัปเดต notifyNewOrder ให้มีรายละเอียดสินค้า */
function notifyOrderToCustomer(data, orderId, orderDetails, itemsTotal, shippingFee, totalPrice) {
  var lineUid = findCustomerLineId(data.phone, data.email);
  if (!lineUid) return; // ไม่มี LINE → ข้าม

  var msg = '🎉 สั่งซื้อสำเร็จ!\n'
    + '━━━━━━━━━━━━━━━\n'
    + '🆔 รหัสออเดอร์: ' + orderId + '\n'
    + '👤 ชื่อ: ' + data.name + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '🛒 รายการที่สั่ง:\n'
    + orderDetails.split('\n').map(function(l){ return '  • ' + l; }).join('\n') + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '💰 ราคาสินค้า: ' + Number(itemsTotal).toLocaleString() + ' ฿\n'
    + (shippingFee > 0 ? '🚚 ค่าจัดส่ง: ' + Number(shippingFee).toLocaleString() + ' ฿\n' : '')
    + '💵 ยอดรวม: ' + Number(totalPrice).toLocaleString() + ' ฿\n'
    + '━━━━━━━━━━━━━━━\n'
    + '📍 ช่องทางจัดส่ง: ' + (data.shippingName || '-') + '\n'
    + '🔍 สถานะ: รอตรวจสอบ\n\n'
    + '💡 พิมพ์ "ติดตามสถานะ" เพื่อตรวจสอบรายการได้เลยครับ';

  sendLineMessage(lineUid, msg);
}

// ==========================================
// 3. MODULE: ระบบติดตามสถานะ (3-in-1)
// ==========================================

function searchTrackingData(keyword, typeFilter) {
  var ss      = getDB();
  var results = [];
  keyword = String(keyword || '').trim().toLowerCase();
  typeFilter = typeFilter || 'all';
  if (!keyword) return [];

  // helper: check if any field contains keyword (partial, case-insensitive)
  function match() {
    for (var a = 0; a < arguments.length; a++) {
      if (String(arguments[a] || '').toLowerCase().indexOf(keyword) !== -1) return true;
    }
    return false;
  }

  if (typeFilter === 'all' || typeFilter === 'order') {
    var orderSheet = ss.getSheetByName('Orders');
    if (orderSheet) {
      var data = orderSheet.getDataRange().getDisplayValues();
      for (var i = data.length - 1; i > 0; i--) {
        if (match(data[i][0], data[i][1], data[i][2], data[i][8], data[i][12])) {
          results.push({
            type:      'order',
            id:        data[i][0],
            name:      data[i][2],
            phone:     data[i][1],
            details:   data[i][3],
            total:     data[i][4],
            status:    data[i][5],
            date:      data[i][6],
            approver:  data[i][7] || '-',
            address:   data[i][9] || '',
            slipUrl:   data[i][10] || '',
            buyerType: data[i][11] || '',
            buyerInfo: data[i][12] || ''
          });
        }
      }
    }
  }

  if (typeFilter === 'all' || typeFilter === 'equipment') {
    var equipSheet = ss.getSheetByName('Equipment_Requests');
    if (equipSheet) {
      var data2 = equipSheet.getDataRange().getDisplayValues();
      for (var j = data2.length - 1; j > 0; j--) {
        if (match(data2[j][0], data2[j][2], data2[j][3], data2[j][4])) {
          results.push({
            type:       'equipment',
            id:         data2[j][0],
            name:       data2[j][2],
            phone:      data2[j][3],
            sid:        data2[j][4],
            items:      data2[j][8],
            status:     data2[j][11],
            date:       data2[j][1],
            returnTime: data2[j][7],
            fileUrl:    data2[j][10],
            approver:   data2[j][12] || '-'
          });
        }
      }
    }
  }

  if (typeFilter === 'all' || typeFilter === 'booking') {
    var bookSheet = ss.getSheetByName('Bookings');
    if (bookSheet) {
      var data3 = bookSheet.getDataRange().getDisplayValues();
      for (var k = data3.length - 1; k > 0; k--) {
        if (match(data3[k][0], data3[k][2], data3[k][4], data3[k][5])) {
          results.push({
            type:      'booking',
            id:        data3[k][0],
            name:      data3[k][2],
            phone:     data3[k][4],
            sid:       data3[k][5],
            room:      data3[k][8],
            status:    data3[k][14],
            date:      data3[k][9],
            timestamp: data3[k][1],
            fileUrl:   data3[k][15],
            approver:  data3[k][16] || '-'
          });
        }
      }
    }
  }

  // ==========================================
  // เพิ่มการดึงข้อมูลจากไฟล์ที่อยู่จัดส่ง (เชื่อมกับ EDUJERSEY และ การจัดส่ง)
  // ==========================================
  if (typeFilter === 'all' || typeFilter === 'jersey' || typeFilter === 'delivery') {
    try {
      var extSheetId = '11R1ls6JVQuqGj1WXisKLsZ3kuKKl0gTfRNUcywefcsk';
      var extSs = SpreadsheetApp.openById(extSheetId);
      var extSheet = extSs.getSheets()[0]; 
      
      if (extSheet) {
        var dataRange = extSheet.getDataRange();
        var extData = dataRange.getDisplayValues();
        var richText = dataRange.getRichTextValues(); 
        
        for (var l = extData.length - 1; l > 0; l--) { 
          var rowId = extData[l][0];             
          var rowDetails = extData[l][1] || '';  
          
          var rawTextUrl = extData[l][2]; 
          var hiddenLink = richText[l][2] ? richText[l][2].getLinkUrl() : null; 
          var rowUrl = hiddenLink || rawTextUrl || ''; 
          
          if (rowDetails && match(rowId, rowDetails, rowUrl)) {
            var cleanText = rowDetails.replace(/^\d+\.\s*/, '').trim(); 
            var words = cleanText.split(/\s+/); 
            var displayName = words[0]; 
            if (words.length > 1) {
              displayName += ' ' + words[1];
            }
            
            results.push({
              type:        'delivery', 
              id:          (rowId || '').replace(')', ''),
              name:        displayName,
              status:      'จัดส่งแล้ว',
              trackingUrl: rowUrl,
              date:        '-'
            });
          }
        }
      }
    } catch(e) {
      Logger.log('Error fetching delivery data: ' + e.message);
    }
  }

  return results; // ส่งผลลัพธ์กลับไปยังหน้าเว็บ
}

// ==========================================
// 4. MODULE: ระบบยืมครุภัณฑ์ (Full Flow)
// ==========================================

function submitComplexBorrow(payload) {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Equipment_Requests');
  if (!sheet) {
    sheet = ss.insertSheet('Equipment_Requests');
    sheet.appendRow([
      'รหัสรายการ', 'วันเวลาที่ทำรายการ', 'ผู้ยืม', 'รหัสนักศึกษา',
      'เบอร์โทร', 'คณะ', 'กำหนดรับ', 'กำหนดคืน',
      'รายการอุปกรณ์', 'วัตถุประสงค์', 'ลิงก์หลักฐาน', 'สถานะ', 'ผู้อนุมัติ', 'อีเมล', 'deliveryPersonId'
    ]);
  }

  var recordId  = generateNextId('Equipment_Requests', 'edurent');
  var timestamp = getThaiTimestamp();

  var folderUrl    = '';
  var parentFolder = getOrCreateFolder('Equipment_Evidence');
  var subFolder    = parentFolder.createFolder(recordId + '_' + payload.name);
  subFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  folderUrl = subFolder.getUrl();

  if (payload.evidenceFiles) {
    payload.evidenceFiles.forEach(function(file, idx) {
      var blob = Utilities.newBlob(
        Utilities.base64Decode(file.data),
        file.mimeType,
        'หลักฐาน_' + idx + '_' + file.name
      );
      subFolder.createFile(blob);
    });
  }

  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 5).setNumberFormat('@');
  sheet.getRange(newRow, 7).setNumberFormat('@');
  sheet.getRange(newRow, 8).setNumberFormat('@');

  // สร้าง Google Doc ใบยืมทันที
  var equipDocUrl = '';
  if (EQUIP_DOC_TEMPLATE_ID && EQUIP_DOC_TEMPLATE_ID !== 'YOUR_EQUIP_TEMPLATE_ID_HERE') {
    equipDocUrl = generateEquipDoc(payload, recordId);
  }

  sheet.appendRow([
    recordId, timestamp, payload.name, payload.sid, payload.phone,
    payload.faculty, payload.borrowTime, payload.returnTime,
    payload.items, payload.purpose, folderUrl, 'รออนุมัติ', '', payload.email || '',
    payload.deliveryPersonId || '', equipDocUrl
  ]);

  // ส่งอีเมลยืนยันให้ผู้ยืม
  if (payload.email) {
    var subject2 = '[Media EDU Club] ยืนยันคำขอยืมอุปกรณ์ รหัส ' + recordId;
    var emailRows = [
      ['รหัสรายการ',           recordId],
      ['ชื่อ-นามสกุล',         payload.name],
      ['รหัสนักศึกษา/บุคลากร', payload.sid],
      ['เบอร์โทรศัพท์',        payload.phone],
      ['คณะ/สังกัด',           payload.faculty],
      ['วัตถุประสงค์',          payload.purpose  || '-'],
      ['วันที่/เวลารับ',        payload.borrowTime],
      ['วันที่/เวลาคืน',        payload.returnTime],
      ['รายการอุปกรณ์',        payload.items.replace(/\n/g, '<br>')],
      ['สถานะ',                'รออนุมัติ'],
      ['วันเวลาที่ส่งคำขอ',    timestamp]
    ];
    if (equipDocUrl) {
      emailRows.push(['📄 ใบยืมอุปกรณ์ (PDF)', '<a href="' + equipDocUrl + '" style="color:#1565c0; font-weight:700;">ดาวน์โหลดเอกสาร PDF</a>']);
    }
    var body2 = buildEmailHtml({
      title:      'ยืนยันคำขอยืมอุปกรณ์',
      headerIcon: 'E',
      color:      '#6a1b9a',
      name:       payload.name,
      intro:      'ระบบได้รับคำขอยืมอุปกรณ์ของคุณแล้ว โปรดตรวจสอบรายละเอียดด้านล่าง:',
      rows:       emailRows,
      showTerms: true,
      note: 'กรุณาอ่านข้อกำหนดด้านบนให้ครบถ้วน ระบบจะแจ้งผลการอนุมัติผ่านอีเมลนี้'
    });
    sendConfirmEmail(payload.email, subject2, body2);
  }

  // ส่งอีเมลแจ้งหน่วยงาน + Admin
  sendOrgNotificationEmail(payload, recordId, payload.orgs || []);

  // แจ้งเตือน LINE Admin
  notifyNewEquipRequest(payload, recordId);

  // แจ้งเตือน LINE ผู้ยืม — ยืนยันรับคำขอ ✅
  notifyUserEquipSubmit(payload.sid, recordId, payload.name, payload);

  return JSON.stringify({ id: recordId, docUrl: equipDocUrl });
}

// ==========================================
// 5. MODULE: ระบบจองห้อง (Bookings)
// ==========================================

function saveBooking(data) {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Bookings');
  if (!sheet) {
    sheet = ss.insertSheet('Bookings');
    sheet.appendRow([
      'รหัสการจอง', 'วันเวลาที่จอง', 'ชื่อ-นามสกุล', 'อีเมล',
      'เบอร์โทร', 'รหัสนักศึกษา', 'หน่วยงาน/สาขา', 'คณะ/สังกัด',
      'ห้อง', 'วันที่ใช้งาน', 'เวลาเริ่ม', 'เวลาสิ้นสุด',
      'จำนวนผู้ใช้', 'วัตถุประสงค์', 'สถานะ', 'ไฟล์แนบ', 'ผู้อนุมัติ', 'ประเภทเอกสาร', 'ลิงค์เอกสาร'
    ]);
  }

  var bookingId = generateNextId('Bookings', 'edubook');
  var timestamp = getThaiTimestamp();

  var docUrl       = '';
  var docTypeLabel = '';
  if (data.docType !== undefined && data.docType >= 0) {
    try {
      docUrl       = generateBookingDoc(data, bookingId);
      docTypeLabel = data.docTypeLabel || '';
    } catch(e) {
      Logger.log('generateBookingDoc error: ' + e.message);
    }
  }

  var newRowBk = sheet.getLastRow() + 1;
  sheet.getRange(newRowBk, 5).setNumberFormat('@');
  sheet.getRange(newRowBk, 10).setNumberFormat('@');
  sheet.getRange(newRowBk, 11).setNumberFormat('@');
  sheet.getRange(newRowBk, 12).setNumberFormat('@');
  sheet.appendRow([
    bookingId, timestamp, data.name, data.email, data.phone,
    data.sid, data.major, data.faculty, data.room, data.date,
    data.startTime, data.endTime, data.userCount, data.purpose,
    'รออนุมัติ', '', '', docTypeLabel, docUrl
  ]);

  // ส่งอีเมลยืนยันการจองห้อง
  if (data.email) {
    var extraDocHtml = '';
    if (docUrl) {
      extraDocHtml = '<div style="margin:20px 0 0;">'
        + '<p style="margin:0 0 10px;font-weight:700;color:#1e293b;font-size:0.92em;border-left:4px solid #0369a1;padding-left:10px;">เอกสารประกอบการขอใช้สถานที่ (' + docTypeLabel + ')</p>'
        + '<a href="' + docUrl + '" target="_blank" style="display:inline-block;padding:11px 22px;background:#0369a1;color:white;border-radius:8px;text-decoration:none;font-weight:700;font-size:0.92em;">คลิกเพื่อเปิดเอกสาร</a>'
        + '<p style="margin:8px 0 0;font-size:0.82em;color:#64748b;">เอกสารถูกสร้างอัตโนมัติ กรุณาตรวจสอบและแก้ไขหากจำเป็นก่อนส่ง</p>'
        + '</div>';
    }
    var subject3 = '[Media EDU Club] ยืนยันการจองห้อง รหัส ' + bookingId;
    var body3    = buildEmailHtml({
      title:      'ยืนยันคำขอจองห้อง',
      headerIcon: 'R',
      color:      '#2e7d32',
      name:       data.name,
      intro:      'ระบบได้รับคำขอจองห้องของคุณแล้ว รายละเอียดมีดังนี้:',
      rows: [
        ['รหัสการจอง',           bookingId],
        ['ชื่อ-นามสกุล',         data.name],
        ['รหัสนักศึกษา/บุคลากร', data.sid],
        ['เบอร์โทรศัพท์',        data.phone],
        ['หน่วยงาน/สาขา',        data.major],
        ['คณะ/สังกัด',           data.faculty],
        ['ห้องที่จอง',            data.room],
        ['วันที่ใช้งาน',          data.date],
        ['เวลาเริ่ม – สิ้นสุด',   data.startTime + ' – ' + data.endTime],
        ['จำนวนผู้ใช้งาน',        data.userCount + ' คน'],
        ['ชื่อโครงการ/กิจกรรม',   data.purpose],
        ['วัตถุประสงค์การใช้ห้อง', data.purposeDetail || '-'],
        ['สถานะ',                'รออนุมัติ'],
        ['วันเวลาที่ส่งคำขอ',    timestamp]
      ],
      extraHtml: extraDocHtml,
      note: 'ระบบจะแจ้งผลการอนุมัติผ่านอีเมลนี้ กรุณาตรวจสอบและเตรียมพร้อมก่อนวันใช้งาน'
    });
    sendConfirmEmail(data.email, subject3, body3);
  }

  // แจ้งเตือน LINE Admin ✅ (แก้ไข: ย้ายมาก่อน return)
  notifyNewBooking(data, bookingId);

  return { bookingId: bookingId, message: 'จองห้องสำเร็จ!', docUrl: docUrl };
}

// ==========================================
// สร้าง Google Doc จาก Template
// ==========================================

// Template IDs สำหรับแต่ละประเภทเอกสารจองห้อง
// 0 = ภายในคณะ, 1 = นอกคณะ (ในมหาวิทยาลัย), 2 = นอกมหาวิทยาลัย
var BOOKING_DOC_TEMPLATES = [
  '1SuQ1qTpF0pIYFCHvXHh2lEO4j7R7rMXpKfhKvffI9d8',  // 0: ภายในคณะ
  '1pC09-8B6sj9LmU2hlFvz6ptQKHSw97lQRjTeRaaizXE',   // 1: นอกคณะ (ในมหาวิทยาลัย)
  ''                                                   // 2: นอกมหาวิทยาลัย (รอ ID)
];

// ==========================================
// 4b. สร้าง Google Doc ใบยืมอุปกรณ์
// ==========================================

var EQUIP_DOC_TEMPLATE_ID = '1j7Dofq6jUlV0_nBVt1Om0UliYLitH5lMXprIeNvrasQ';

// แปลงวันที่ YYYY-MM-DD → "วันจันทร์ ที่ 25 มกราคม 2568"
function formatThaiDateFull(dateStr) {
  if (!dateStr || dateStr === '-') return '-';
  var THAI_DAYS   = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
  var THAI_MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                     'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  try {
    var p = dateStr.trim().split('-');
    if (p.length !== 3) return dateStr;
    var y = parseInt(p[0]), m = parseInt(p[1]) - 1, d = parseInt(p[2]);
    if (isNaN(y) || isNaN(m) || isNaN(d)) return dateStr;
    var dt = new Date(y, m, d);
    return 'วัน' + THAI_DAYS[dt.getDay()] + ' ที่ ' + d + ' ' + THAI_MONTHS[m] + ' ' + (y + 543);
  } catch(e) { return dateStr; }
}

function formatThaiDateNow() {
  var THAI_DAYS   = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
  var THAI_MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                     'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  var now = new Date();
  return 'วัน' + THAI_DAYS[now.getDay()] + ' ที่ ' + now.getDate() + ' ' + THAI_MONTHS[now.getMonth()] + ' ' + (now.getFullYear() + 543);
}

function generateEquipDoc(payload, recordId) {
  try {
    var templateFile = DriveApp.getFileById(EQUIP_DOC_TEMPLATE_ID);
    var folder       = getOrCreateFolder('Equip_Documents');
    var newFile      = templateFile.makeCopy('ใบยืม_' + recordId + '_' + payload.name, folder);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var docId = newFile.getId();

    // แยกวันที่ / เวลา
    var borrowRaw = String(payload.borrowTime || '').trim().replace('T', ' ');
    var returnRaw = String(payload.returnTime || '').trim().replace('T', ' ');
    var bSplit = borrowRaw.split(' ');
    var rSplit = returnRaw.split(' ');
    var borrowDateThai = formatThaiDateFull(bSplit[0] || '');
    var returnDateThai = formatThaiDateFull(rSplit[0] || '');
    var borrowTimeFmt  = ((bSplit[1] || '').substring(0, 5) || '-') + ' น.';
    var returnTimeFmt  = ((rSplit[1] || '').substring(0, 5) || '-') + ' น.';
    var activityDateThai = (payload.activityDate && payload.activityDate !== '')
      ? formatThaiDateFull(String(payload.activityDate))
      : borrowDateThai;

    // แยก items ตัด org prefix
    var itemLines = String(payload.items || '').split('\n').filter(function(l){ return l.trim(); });
    if (itemLines.length > 0 && itemLines[0].indexOf('[ยืม:') === 0) { itemLines.shift(); }

    // สร้าง pairs placeholder→value ใช้ Docs API replaceAllText (ไม่ต้องใช้ regex เลย)
    var pairs = [
      ['{DocNo}',            recordId],
      ['{BorrowerName}',     payload.name      || '-'],
      ['{BorrowerSid}',      payload.sid       || '-'],
      ['{BorrowerFaculty}',  payload.faculty   || '-'],
      ['{BorrowerPhone}',    payload.phone     || '-'],
      ['{BorrowerClubName}', payload.clubName  || '-'],
      ['{BorrowerFor}',      payload.borrowFor || payload.purpose || '-'],
      ['{Purpose}',          payload.purpose   || '-'],
      ['{BorrowDate}',       borrowDateThai],
      ['{BorrowTime}',       borrowTimeFmt],
      ['{ReturnDate}',       returnDateThai],
      ['{ReturenDate}',      returnDateThai],
      ['{ReturnTime}',       returnTimeFmt],
      ['{DDay}',             activityDateThai],
      ['{RequestDate}',      formatThaiDateNow()],
    ];

    // ★ อุปกรณ์ — ใช้ {I1} {R1} {Q1} เลขอาราบิก (ตรงกับ template)
    for (var i = 0; i < 10; i++) {
      var idx  = String(i + 1);
      var line = itemLines[i] || '';
      var qtyM = line.match(/\(x(\d+)\)\s*$/);
      var qty  = qtyM ? qtyM[1] : (line ? '1' : '');
      var name = line ? line.replace(/\s*\(x\d+\)\s*$/, '').replace(/^\[.*?\]\s*/, '') : '';
      pairs.push(['{I' + idx + '}', line ? idx : '___EMPTY_ROW___']);
      pairs.push(['{R' + idx + '}', name]);
      pairs.push(['{Q' + idx + '}', qty]);
    }


    // ★ batchUpdate: replaceAllText (ใช้ Arabic numerals ตรงกับ template)
    var replaceRequests = pairs.map(function(p) {
      return {
        replaceAllText: {
          containsText: { text: p[0], matchCase: true },
          replaceText: String(p[1])
        }
      };
    });

    try {
      Docs.Documents.batchUpdate({ requests: replaceRequests }, docId);
    } catch(e2) {
      Logger.log('batchUpdate error: ' + e2.message);
      var doc2  = DocumentApp.openById(docId);
      var body2 = doc2.getBody();
      pairs.forEach(function(p) {
        try {
          var escaped = p[0].replace(/[.+?()[\]{}^$|*\\]/g, '\\$&');
          body2.replaceText(escaped, String(p[1]));
        } catch(ex) { Logger.log('fallback rep[' + p[0] + ']: ' + ex.message); }
      });
      body2.editAsText().setFontFamily('Sarabun').setFontSize(14);
      doc2.saveAndClose();
      return 'https://docs.google.com/document/d/' + docId + '/preview';
    }

    // ★ เปิด DocumentApp เพื่อ: (1) ลบแถวว่าง (2) ตั้งฟอนต์
    try {
      var docD  = DocumentApp.openById(docId);
      var bodyD = docD.getBody();

      // ลบแถวตาราง ที่ cell แรกมี ___EMPTY_ROW___ (หมายถึงไม่มีอุปกรณ์)
      var tables = bodyD.getTables();
      tables.forEach(function(tbl) {
        // วนจากท้ายขึ้นบน เพื่อ index ไม่เลื่อน
        for (var r = tbl.getNumRows() - 1; r >= 1; r--) {
          var cellText = tbl.getCell(r, 0).getText().trim();
          if (cellText === '___EMPTY_ROW___' || cellText === '') {
            // ลบแถว — แต่ต้องเหลืออย่างน้อย 1 แถว (header)
            if (tbl.getNumRows() > 1) {
              tbl.removeRow(r);
            }
          }
        }
      });

      // ตั้งฟอนต์ + ขนาด 14
      bodyD.editAsText().setFontFamily('Sarabun').setFontSize(14);
      docD.saveAndClose();
    } catch(e3) {
      Logger.log('post-process error: ' + e3.message);
    }

    return 'https://docs.google.com/document/d/' + docId + '/preview';
  } catch(e) {
    Logger.log('generateEquipDoc error: ' + e.message);
    return '';
  }
}


function toThaiNumerals(str) {
  var thaiDigits = ['๐','๑','๒','๓','๔','๕','๖','๗','๘','๙'];
  return String(str).replace(/[0-9]/g, function(d) {
    return thaiDigits[parseInt(d)];
  });
}

var DOC_HEAD_MAP = {
  0: 'คณบดีคณะศึกษาศาสตร์',
  1: 'อธิการบดีมหาวิทยาลัยเชียงใหม่',
  2: 'ผู้อำนวยการ / ผู้บริหารหน่วยงาน'
};

function generateBookingDoc(data, bookingId) {
  try {
    var docType = (data.docType !== undefined && data.docType >= 0) ? data.docType : 0;
    var templateId = BOOKING_DOC_TEMPLATES[docType] || BOOKING_DOC_TEMPLATES[0];
    if (!templateId) {
      Logger.log('generateBookingDoc: ไม่มี template ID สำหรับ docType ' + docType);
      return '';
    }
    var templateFile = DriveApp.getFileById(templateId);
    var folder       = getOrCreateFolder('Booking_Documents');
    var newFile      = templateFile.makeCopy(
      'เอกสารขอใช้ห้อง_' + bookingId + '_' + (data.doc_activity || data.purpose),
      folder
    );
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    var docId = newFile.getId();

    var THAI_MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                       'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

    function toThaiDate(dateStr) {
      if (!dateStr) return '-';
      var dp = String(dateStr).split('-');
      if (dp.length === 3) {
        return parseInt(dp[2]) + ' ' + THAI_MONTHS[parseInt(dp[1]) - 1] + ' ' + (parseInt(dp[0]) + 543);
      }
      return dateStr;
    }

    var now = new Date();
    var requestDate = now.getDate() + ' ' + THAI_MONTHS[now.getMonth()] + ' ' + (now.getFullYear() + 543);

    var clubName    = data.doc_club        || data.major   || '-';
    var purpose     = data.doc_activity    || data.purpose || '-';
    var borrowFor   = data.purposeDetail   || data.doc_for || data.purpose || '-';
    var yearEdu     = data.doc_year        || '-';
    var contactName = data.doc_contactName || data.name    || '-';
    var contactTel  = data.doc_contactTel  || data.phone   || '-';
    var proLevel    = data.doc_proLevel    || '';
    var proName     = data.doc_proName     || '';
    var proFullName = proLevel ? proLevel + ' ' + proName : proName;
    var levelStr    = data.doc_level       || '';
    var headStr     = data.doc_head        || '';
    var dDayStr     = toThaiDate(data.date);

    var rooms = [];
    if (data.rooms && Array.isArray(data.rooms) && data.rooms.length > 0) {
      rooms = data.rooms;
    } else {
      rooms = [{
        room:      data.room || '-',
        date:      data.date || '',
        startTime: data.startTime || '-',
        endTime:   data.endTime   || '-',
        activity:  data.doc_activity || data.purpose || '-'
      }];
    }

    var pairs = [
      ['{RequestDate}',      requestDate],
      ['{ReequestDate}',     requestDate],
      ['{BorrowerClubName}', clubName],
      ['{Purpose}',          purpose],
      ['{DDay}',             dDayStr],
      ['{BorrowerFor}',      borrowFor],
      ['{Year}',             yearEdu],
      ['{BorrowerPhone}',    contactTel],
      ['{BorrowerName}',     contactName],
      ['{ProLevel}',         proLevel],
      ['{ProName}',          proFullName],
      ['{Level}',            levelStr],
      ['{Head}',             headStr]
    ];

    var MAX_ROWS = 10;
    for (var i = 0; i < MAX_ROWS; i++) {
      var idx = String(i + 1);
      if (i < rooms.length) {
        var rm = rooms[i];
        var timeThai = (rm.startTime || '-') + ' - ' + (rm.endTime || '-') + ' น.';
        pairs.push(['{l'   + idx + '}', idx]);
        pairs.push(['{I'   + idx + '}', idx]);
        pairs.push(['{R'   + idx + '}', rm.room     || '-']);
        pairs.push(['{T'   + idx + '}', timeThai]);
        pairs.push(['{For' + idx + '}', rm.activity || purpose]);
        if (i === 0) { pairs.push(['{Date}', toThaiDate(rm.date)]); }
        pairs.push(['{Date' + idx + '}', toThaiDate(rm.date)]);
      } else {
        pairs.push(['{l'   + idx + '}', '___DEL___']);
        pairs.push(['{I'   + idx + '}', '___DEL___']);
        pairs.push(['{R'   + idx + '}', '']);
        pairs.push(['{T'   + idx + '}', '']);
        pairs.push(['{For' + idx + '}', '']);
        pairs.push(['{Date' + idx + '}', '']);
      }
    }

    var replaceRequests = pairs.map(function(p) {
      return {
        replaceAllText: {
          containsText: { text: p[0], matchCase: true },
          replaceText: String(p[1])
        }
      };
    });

    try {
      Docs.Documents.batchUpdate({ requests: replaceRequests }, docId);
    } catch(e2) {
      Logger.log('batchUpdate error: ' + e2.message);
      var doc2  = DocumentApp.openById(docId);
      var body2 = doc2.getBody();
      pairs.forEach(function(p) {
        try {
          body2.replaceText(p[0].replace(/[{}\[\]\\^$.|?*+()]/g, '\\$&'), String(p[1]));
        } catch(ex) { Logger.log('rep err[' + p[0] + ']: ' + ex.message); }
      });
      body2.editAsText().setFontFamily('Sarabun').setFontSize(docType === 1 ? 16 : 15);
      doc2.saveAndClose();
      return 'https://docs.google.com/document/d/' + docId + '/preview';
    }

    try {
      var docD  = DocumentApp.openById(docId);
      var bodyD = docD.getBody();
      var tables = bodyD.getTables();
      tables.forEach(function(tbl) {
        for (var r = tbl.getNumRows() - 1; r >= 1; r--) {
          var cellText = tbl.getCell(r, 0).getText().trim();
          if (cellText === '___DEL___') {
            if (tbl.getNumRows() > 1) tbl.removeRow(r);
          }
        }
      });
      bodyD.editAsText().setFontFamily('Sarabun').setFontSize(docType === 1 ? 16 : 15);

      // --- จัดรูปแบบพิเศษสำหรับเอกสาร นอกคณะ (ในมหาวิทยาลัย) ---
      if (docType === 1) {
        var numChildren = bodyD.getNumChildren();
        for (var ci = 0; ci < numChildren; ci++) {
          var child = bodyD.getChild(ci);
          if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
          var para = child.asParagraph();
          var text = para.getText().trim();

          // "บันทึกข้อความ" → ตัวหนา ขนาด 29
          if (text === 'บันทึกข้อความ') {
            para.editAsText()
              .setBold(true)
              .setFontSize(29);
            continue;
          }

          // ส่วนงาน / ที่ / เรื่อง / วันที่ → ตัวหนา ขนาด 20
          // จับ pattern: ขึ้นต้นด้วยคำเหล่านี้ตามด้วย (space หรือ tab หรือ :)
          if (/^(ส่วนงาน|ที่|เรื่อง|วันที่)[\s\t:]/.test(text) || text === 'ส่วนงาน' || text === 'ที่' || text === 'เรื่อง' || text === 'วันที่') {
            // หา index ของคำ label ใน paragraph แล้วทำ bold เฉพาะส่วน label
            var labelWords = ['ส่วนงาน','เรื่อง','วันที่','ที่'];
            var fullText = para.getText();
            for (var li = 0; li < labelWords.length; li++) {
              var lw = labelWords[li];
              if (fullText.indexOf(lw) === 0) {
                para.editAsText()
                  .setBold(true)
                  .setFontSize(20);
                break;
              }
            }
          }
        }
      }

      docD.saveAndClose();
    } catch(e3) {
      Logger.log('post-process error: ' + e3.message);
    }

    return 'https://docs.google.com/document/d/' + docId + '/preview';
  } catch(e) {
    Logger.log('generateBookingDoc error: ' + e.message);
    return '';
  }
}

// ==========================================
// 6. MODULE: ระบบ Admin (Backoffice)
// ==========================================

// ==========================================
// Admin Security — Hash + Rate Limiting
// ==========================================

/** SHA-256 hash ด้วย GAS Utilities */
function hashPassword(password, salt) {
  var raw = salt ? (salt + ':' + password) : password;
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    raw,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

/** สุ่ม salt */
function generateSalt() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}

/** Hash password แล้วเก็บในรูป "salt:hash" */
function encodePassword(plainPassword) {
  var salt = generateSalt();
  var hash = hashPassword(plainPassword, salt);
  return salt + ':' + hash;
}

/** ตรวจสอบ password กับ stored hash */
function verifyPassword(plainPassword, stored) {
  if (!stored || stored.indexOf(':') === -1) {
    // Plaintext fallback (สำหรับ account เก่า — ครั้งเดียว)
    return plainPassword === stored;
  }
  var parts = stored.split(':');
  var salt  = parts[0];
  var hash  = parts[1];
  return hashPassword(plainPassword, salt) === hash;
}

// ── Rate limiting keys ──
var LOGIN_ATTEMPTS_PREFIX = 'LOGIN_FAIL_';
var LOGIN_LOCKOUT_PREFIX  = 'LOGIN_LOCK_';
var MAX_ATTEMPTS = 5;       // ผิดได้สูงสุด 5 ครั้ง
var LOCKOUT_MINUTES = 15;   // ล็อก 15 นาที
var SESSION_MINUTES = 120;  // session หมดอายุใน 2 ชั่วโมง

function getRateLimitKey(username) {
  return LOGIN_ATTEMPTS_PREFIX + username.toLowerCase().replace(/[^a-z0-9]/g, '_');
}
function getLockoutKey(username) {
  return LOGIN_LOCKOUT_PREFIX + username.toLowerCase().replace(/[^a-z0-9]/g, '_');
}

/** ตรวจว่าถูก lock อยู่ไหม — คืน {locked, remainingMin} */
function checkLockout(username) {
  var props   = PropertiesService.getScriptProperties();
  var lockKey = getLockoutKey(username);
  var lockVal = props.getProperty(lockKey);
  if (!lockVal) return { locked: false };
  var lockTime = parseInt(lockVal);
  var elapsed  = (Date.now() - lockTime) / 60000;
  if (elapsed < LOCKOUT_MINUTES) {
    return { locked: true, remainingMin: Math.ceil(LOCKOUT_MINUTES - elapsed) };
  }
  // Lock หมดอายุแล้ว — ล้าง
  props.deleteProperty(lockKey);
  props.deleteProperty(getRateLimitKey(username));
  return { locked: false };
}

/** บันทึก login fail และ lock ถ้าเกิน limit */
function recordLoginFail(username) {
  var props   = PropertiesService.getScriptProperties();
  var attKey  = getRateLimitKey(username);
  var current = parseInt(props.getProperty(attKey) || '0') + 1;
  if (current >= MAX_ATTEMPTS) {
    props.setProperty(getLockoutKey(username), String(Date.now()));
    props.deleteProperty(attKey);
    Logger.log('[SECURITY] Account locked: ' + username);
  } else {
    props.setProperty(attKey, String(current));
  }
  return current;
}

/** ล้าง fail counter หลัง login สำเร็จ */
function clearLoginFail(username) {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(getRateLimitKey(username));
  props.deleteProperty(getLockoutKey(username));
}

/** สร้าง session token */
function createSessionToken(username, role) {
  var token   = Utilities.getUuid();
  var expires = Date.now() + SESSION_MINUTES * 60 * 1000;
  var data    = JSON.stringify({ username: username, role: role, expires: expires });
  PropertiesService.getScriptProperties()
    .setProperty('SESSION_' + token, data);
  return token;
}

/** ตรวจ session token — คืน {valid, username, role} */
function verifySessionToken(token) {
  if (!token) return { valid: false };
  var raw = PropertiesService.getScriptProperties()
    .getProperty('SESSION_' + token);
  if (!raw) return { valid: false };
  try {
    var data = JSON.parse(raw);
    if (Date.now() > data.expires) {
      PropertiesService.getScriptProperties().deleteProperty('SESSION_' + token);
      return { valid: false, expired: true };
    }
    return { valid: true, username: data.username, role: data.role };
  } catch(e) {
    return { valid: false };
  }
}

/** Invalidate session (logout) */
function invalidateSession(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('SESSION_' + token);
}

/** Validate password strength */
function validatePasswordStrength(password) {
  if (!password || password.length < 8)
    return { ok: false, msg: 'Password ต้องมีอย่างน้อย 8 ตัวอักษร' };
  if (!/[A-Za-z]/.test(password))
    return { ok: false, msg: 'Password ต้องมีตัวอักษรภาษาอังกฤษ' };
  if (!/[0-9]/.test(password))
    return { ok: false, msg: 'Password ต้องมีตัวเลข' };
  var weak = ['password','admin1234','12345678','admin','1234567890','pass1234'];
  if (weak.indexOf(password.toLowerCase()) !== -1)
    return { ok: false, msg: 'Password นี้ไม่ปลอดภัย กรุณาเปลี่ยน' };
  return { ok: true };
}

/** Login หลัก — ใช้แทน verifyAdmin เดิม */
function verifyAdmin(username, password) {
  if (!username || !password)
    return { success: false, message: 'กรุณากรอก Username และ Password' };

  username = String(username).trim();

  // 1. ตรวจ lockout ก่อน
  var lockStatus = checkLockout(username);
  if (lockStatus.locked) {
    Logger.log('[SECURITY] Login blocked (locked): ' + username);
    return {
      success: false,
      message: 'บัญชีถูกล็อกชั่วคราว กรุณารอ ' + lockStatus.remainingMin + ' นาที',
      locked: true
    };
  }

  // 2. ค้นหา account
  var sheet = getDB().getSheetByName('Admins');
  if (!sheet) return { success: false, message: 'ไม่พบชีต Admins' };
  var data = sheet.getDataRange().getDisplayValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== username) continue;

    var storedPass = String(data[i][1] || '');
    var isValid    = verifyPassword(password, storedPass);

    if (isValid) {
      // Login สำเร็จ
      clearLoginFail(username);

      // Auto-upgrade plaintext password → hashed (migration)
      if (storedPass.indexOf(':') === -1) {
        sheet.getRange(i + 1, 2).setValue(encodePassword(password));
        Logger.log('[SECURITY] Password upgraded to hash: ' + username);
      }

      var role  = String(data[i][3] || 'admin').trim().toLowerCase();
      var token = createSessionToken(username, role);

      Logger.log('[AUTH] Login success: ' + username + ' role=' + role);
      return {
        success:  true,
        name:     data[i][2],
        role:     role,
        username: username,
        token:    token,
        sessionMinutes: SESSION_MINUTES
      };
    }

    // Password ผิด
    var attempts = recordLoginFail(username);
    var remaining = MAX_ATTEMPTS - attempts;
    Logger.log('[SECURITY] Login fail: ' + username + ' attempts=' + attempts);
    logAuditEvent('auth', username, username, 'login', 'failed', 'system', '');

    return {
      success: false,
      message: remaining > 0
        ? 'Username/Password ผิด! เหลือโอกาสอีก ' + remaining + ' ครั้ง'
        : 'บัญชีถูกล็อก ' + LOCKOUT_MINUTES + ' นาที',
      attemptsLeft: remaining
    };
  }

  // ไม่พบ username
  recordLoginFail(username);
  return { success: false, message: 'Username/Password ผิด!' };
}

/** Admin logout */
function adminLogout(token) {
  invalidateSession(token);
  return { ok: true };
}

/** Superadmin: unlock account ที่ถูก lock */
function unlockAdminAccount(username) {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(getLockoutKey(username));
  props.deleteProperty(getRateLimitKey(username));
  return { ok: true, msg: 'ปลดล็อก ' + username + ' แล้ว' };
}

// ==========================================
// Admin Account Management (CRUD)
// ==========================================
var ADMIN_ROLES = ['superadmin','equipment','booking','order','content'];

function getAdminAccounts() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Admins');
  if (!sheet) {
    sheet = ss.insertSheet('Admins');
    sheet.appendRow(['Username','Password','ชื่อ-นามสกุล','Role']);
    sheet.getRange(1,1,1,4).setBackground('#1a3a6b').setFontColor('#fff').setFontWeight('bold');
    // สร้าง superadmin ตั้งต้น ด้วย password แบบ hash
    var initPass = encodePassword('Admin@1234');
    sheet.appendRow(['admin', initPass, 'ผู้ดูแลระบบ', 'superadmin']);
    Logger.log('[SECURITY] Default admin created. Password: Admin@1234 (เปลี่ยนทันที!)');
  }
  var rows = sheet.getDataRange().getDisplayValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({ row: i+1, username: rows[i][0], name: rows[i][2], role: rows[i][3] || 'superadmin' });
  }
  return result;
}

function addAdminAccount(data) {
  var sheet = getDB().getSheetByName('Admins');
  if (!sheet) return { ok: false, msg: 'ไม่พบชีต Admins' };
  // ตรวจ duplicate username
  var rows = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.username) return { ok: false, msg: 'Username นี้มีอยู่แล้ว' };
  }
  // Validate password strength
  var pwCheck = validatePasswordStrength(data.password || '');
  if (!pwCheck.ok) return { ok: false, msg: pwCheck.msg };
  // Hash password ก่อน save
  var hashed = encodePassword(data.password);
  sheet.appendRow([data.username, hashed, data.name, data.role || 'admin']);
  return { ok: true, msg: 'เพิ่มบัญชีสำเร็จ', accounts: getAdminAccounts() };
}

function updateAdminAccount(data) {
  var sheet = getDB().getSheetByName('Admins');
  if (!sheet) return { ok: false, msg: 'ไม่พบชีต Admins' };
  var row = parseInt(data.row);
  if (isNaN(row) || row < 2) return { ok: false, msg: 'Row ไม่ถูกต้อง' };
  sheet.getRange(row, 3).setValue(data.name);
  sheet.getRange(row, 4).setValue(data.role);
  if (data.password) {
    var pwCheck2 = validatePasswordStrength(data.password);
    if (!pwCheck2.ok) return { ok: false, msg: pwCheck2.msg };
    sheet.getRange(row, 2).setValue(encodePassword(data.password));
  }
  return { ok: true, msg: 'อัปเดตสำเร็จ', accounts: getAdminAccounts() };
}

function deleteAdminAccount(row) {
  var sheet = getDB().getSheetByName('Admins');
  if (!sheet) return { ok: false, msg: 'ไม่พบชีต Admins' };
  var r = parseInt(row);
  if (isNaN(r) || r < 2) return { ok: false, msg: 'Row ไม่ถูกต้อง' };
  // ป้องกันลบ superadmin คนสุดท้าย
  var rows = sheet.getDataRange().getDisplayValues();
  var superCount = 0;
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][3] || '').toLowerCase() === 'superadmin') superCount++;
  }
  if ((rows[r-1][3] || '').toLowerCase() === 'superadmin' && superCount <= 1)
    return { ok: false, msg: 'ไม่สามารถลบ Superadmin คนสุดท้ายได้' };
  sheet.deleteRow(r);
  return { ok: true, msg: 'ลบสำเร็จ', accounts: getAdminAccounts() };
}

function getAdminOrders() {
  var sheet = getDB().getSheetByName('Orders');
  if (!sheet) return [];
  var data   = sheet.getDataRange().getDisplayValues();
  var orders = [];
  for (var i = data.length - 1; i > 0; i--) {
    orders.push({
      row:      i + 1,
      id:       data[i][0],
      phone:    data[i][1],
      name:     data[i][2],
      details:  data[i][3],
      total:    data[i][4],
      status:   data[i][5],
      date:     data[i][6],
      approver: data[i][7] || '-',
      email:    data[i][8] || '',
      address:  data[i][9] || '',
      slipUrl:  data[i][10] || '',
      buyerType:    data[i][11] || '',
      buyerInfo:    data[i][12] || '',
      shippingName: data[i][13] || '',
      shippingFee:  data[i][14] || '',
      contact:      data[i][15] || ''
    });
  }
  return orders;
}

function getAdminEquipRequests() {
  var sheet = getDB().getSheetByName('Equipment_Requests');
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var reqs = [];
  for (var i = data.length - 1; i > 0; i--) {
    reqs.push({
      row:              i + 1,
      id:               data[i][0],
      timestamp:        data[i][1],
      name:             data[i][2],
      sid:              data[i][3],
      phone:            data[i][4],
      faculty:          data[i][5],
      borrowTime:       data[i][6],
      returnTime:       data[i][7],
      items:            data[i][8],
      purpose:          data[i][9],
      evidenceUrl:      data[i][10],
      status:           data[i][11],
      approver:         data[i][12] || '-',
      email:            data[i][13] || '',
      deliveryPersonId: data[i][14] || '',
      docUrl:           data[i][15] || ''
    });
  }
  return reqs;
}

function updateEquipRequestFull(row, payload) {
  var sheet = getDB().getSheetByName('Equipment_Requests');
  if (!sheet) return 'ไม่พบชีต';

  sheet.getRange(row, 3).setValue(payload.name);
  sheet.getRange(row, 4).setValue(payload.sid);
  sheet.getRange(row, 5).setValue(payload.phone);
  sheet.getRange(row, 6).setValue(payload.faculty);
  sheet.getRange(row, 7).setValue(payload.borrowTime);
  sheet.getRange(row, 8).setValue(payload.returnTime);
  sheet.getRange(row, 9).setValue(payload.items);
  sheet.getRange(row, 10).setValue(payload.purpose);
  sheet.getRange(row, 12).setValue(payload.status);
  var adminName = payload.adminName || 'Admin';
  sheet.getRange(row, 13).setValue(adminName);
  if (payload.deliveryPersonId !== undefined) {
    sheet.getRange(row, 15).setValue(payload.deliveryPersonId || '');
  }

  var email = sheet.getRange(row, 14).getValue();
  if (email && payload.status && payload.status.includes('อนุมัติ') && !payload.status.includes('ไม่')) {
    var deliveryInfo = null;
    if (payload.deliveryPersonId) {
      var dlvList = getDeliveryConfigList();
      for (var d = 0; d < dlvList.length; d++) {
        if (dlvList[d].id === payload.deliveryPersonId) { deliveryInfo = dlvList[d]; break; }
      }
    }

    var baseRows = [
      ['รหัสรายการ',           sheet.getRange(row, 1).getValue()],
      ['ชื่อ-นามสกุล',         payload.name],
      ['รหัสนักศึกษา/บุคลากร', payload.sid    || '-'],
      ['เบอร์โทรศัพท์',        payload.phone  || '-'],
      ['คณะ/สังกัด',           payload.faculty || '-'],
      ['รายการอุปกรณ์',        String(payload.items || '').replace(/\n/g, '<br>')],
      ['วันที่/เวลารับ',        payload.borrowTime],
      ['วันที่/เวลาคืน',        payload.returnTime],
      ['สถานะ',                payload.status],
      ['ดำเนินการโดย',          adminName]
    ];

    var deliveryHtml = '';
    if (deliveryInfo) {
      deliveryHtml = '<div style="margin:20px 0 0;">'
        + '<p style="margin:0 0 10px;font-weight:700;color:#1e293b;font-size:0.92em;border-left:4px solid #0891b2;padding-left:10px;">ข้อมูลการรับอุปกรณ์</p>'
        + '<table style="width:100%;border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;font-size:0.9em;">'
        + _dvRow('ชื่อผู้จ่ายของ', deliveryInfo.name     || '-')
        + _dvRow('หน่วยงาน',       deliveryInfo.org      || '-')
        + _dvRow('เบอร์โทรศัพท์', deliveryInfo.phone    || '-')
        + _dvRow('เวลานัดรับ',     deliveryInfo.time     || '-')
        + _dvRow('สถานที่รับ',     deliveryInfo.location || '-')
        + '</table></div>';
    }

    var subjectE = '[Media EDU Club] ✅ อนุมัติการยืมอุปกรณ์ ' + sheet.getRange(row, 1).getValue();
    var bodyE    = buildEmailHtml({
      title:      'อนุมัติการยืมอุปกรณ์',
      headerIcon: 'E',
      color:      '#27ae60',
      name:       payload.name,
      intro:      'ยินดีด้วย! คำขอยืมอุปกรณ์ของคุณได้รับการอนุมัติแล้ว กรุณาติดต่อรับอุปกรณ์ตามข้อมูลด้านล่าง',
      rows:       baseRows,
      extraHtml:  deliveryHtml,
      showTerms:  true,
      note:       'กรุณาติดต่อผู้จ่ายของตามข้อมูลที่ระบุ หากมีข้อสงสัยกรุณาติดต่อทีมงาน Media EDU Club CMU.'
    });
    sendConfirmEmail(email, subjectE, bodyE);
  }

  return 'อัปเดตสำเร็จ';
}

function _dvRow(label, val) {
  return '<tr>'
    + '<td style="padding:8px 14px;font-weight:600;color:#475569;background:#f8fafc;width:40%;border-bottom:1px solid #f1f5f9;">' + label + '</td>'
    + '<td style="padding:8px 14px;color:#2d3748;border-bottom:1px solid #f1f5f9;">' + val + '</td>'
    + '</tr>';
}

// ดึงการจองทั้งหมดในวันที่ระบุ (ใช้ check conflict ฝั่งผู้ใช้)
function getBookingsByDate(date) {
  var sheet = getDB().getSheetByName('Bookings');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var rowDate   = String(r[9]  || '').trim();   // col 10: date
    var rowStatus = String(r[14] || '').trim();   // col 15: status
    if (rowDate === date && !rowStatus.includes('ยกเลิก')) {
      result.push({
        id:        String(r[0]  || ''),
        name:      String(r[2]  || ''),
        room:      String(r[8]  || ''),
        date:      rowDate,
        startTime: String(r[10] || ''),
        endTime:   String(r[11] || ''),
        purpose:   String(r[13] || ''),
        status:    rowStatus
      });
    }
  }
  return result;
}

function getAdminBookings() {
  var sheet = getDB().getSheetByName('Bookings');
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var bks  = [];
  for (var i = data.length - 1; i > 0; i--) {
    bks.push({
      row:       i + 1,
      id:        data[i][0],
      timestamp: data[i][1],
      name:      data[i][2],
      email:     data[i][3],
      phone:     data[i][4],
      sid:       data[i][5],
      major:     data[i][6],
      faculty:   data[i][7],
      room:      data[i][8],
      date:      data[i][9],
      startTime: data[i][10],
      endTime:   data[i][11],
      userCount: data[i][12],
      purpose:   data[i][13],
      status:    data[i][14],
      approver:  data[i][16] || '-',
      docTypeLabel: data[i][17] || '',
      docUrl:    data[i][18] || ''
    });
  }
  return bks;
}

function updateOrderStatus(row, status, adminName) {
  var sheet = getDB().getSheetByName('Orders');
  var oldStatus = sheet.getRange(row, 6).getValue();
  sheet.getRange(row, 6).setValue(status);
  sheet.getRange(row, 8).setValue(adminName);

  var rowData = sheet.getRange(row, 1, 1, 9).getValues()[0];
  var email   = rowData[8];
  logAuditEvent('order', rowData[0], rowData[2], oldStatus, status, adminName, '');
  if (email) {
    var orderId = rowData[0], name = rowData[2], details = rowData[3], total = rowData[4];
    var subjectO = '[Media EDU Club] อัปเดตสถานะออเดอร์ ' + orderId + ' - ' + status;
    var bodyO    = buildEmailHtml({
      title:      getStatusTitle(status),
      headerIcon: 'S',
      color:      getStatusColor(status),
      name:       name,
      intro:      'ทีมงานได้ดำเนินการอัปเดตสถานะคำสั่งซื้อของคุณแล้ว รายละเอียดมีดังนี้:',
      rows: [
        ['รหัสออเดอร์',   orderId],
        ['ชื่อลูกค้า',    name],
        ['เบอร์โทรศัพท์', rowData[1]],
        ['รายการสินค้า',  String(details).replace(/\n/g, '<br>')],
        ['ยอดรวม',        Number(total).toLocaleString() + ' บาท'],
        ['สถานะใหม่',     status],
        ['ดำเนินการโดย',  adminName]
      ],
      note: 'หากมีข้อสงสัย กรุณาติดต่อทีมงาน Media EDU Club CMU.'
    });
    sendConfirmEmail(email, subjectO, bodyO);
  }
}

// ==========================================
// ระบบตรวจสอบสลิป (Admin)
// ==========================================

/** ดึงรายการออเดอร์ที่รอตรวจสลิป (สถานะ = รอตรวจสอบ + มีสลิป) */
function getOrdersForSlipVerify() {
  var sheet = getDB().getSheetByName('Orders');
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var result = [];
  for (var i = data.length - 1; i > 0; i--) {
    var status  = String(data[i][5] || '').trim();
    var slipUrl = String(data[i][10] || '').trim();
    if (status === 'รอตรวจสอบ' && slipUrl) {
      result.push({
        row:          i + 1,
        id:           data[i][0],
        name:         data[i][2],
        phone:        data[i][1],
        details:      data[i][3],
        total:        data[i][4],
        status:       status,
        date:         data[i][6],
        email:        data[i][8] || '',
        slipUrl:      slipUrl,
        buyerType:    data[i][11] || '',
        shippingName: data[i][13] || '',
        shippingFee:  data[i][14] || 0
      });
    }
  }
  return result;
}

/** Admin อนุมัติสลิป → เปลี่ยนสถานะ + แจ้งลูกค้า */
function approveSlip(row, adminName) {
  var sheet = getDB().getSheetByName('Orders');
  if (!sheet) throw new Error('ไม่พบ Orders sheet');
  var rowData = sheet.getRange(row, 1, 1, 16).getValues()[0];
  var oldStatus = rowData[5];
  var newStatus = 'กำลังเตรียมจัดส่ง';

  sheet.getRange(row, 6).setValue(newStatus);
  sheet.getRange(row, 8).setValue(adminName);
  logAuditEvent('order', rowData[0], rowData[2], oldStatus, newStatus, adminName, 'อนุมัติสลิป');

  // แจ้ง LINE + Email ลูกค้า
  _notifySlipResult(rowData, newStatus, adminName, true);
  return { ok: true };
}

/** Admin ปฏิเสธสลิป → เปลี่ยนสถานะ + แจ้งลูกค้า */
function rejectSlip(row, adminName, reason) {
  var sheet = getDB().getSheetByName('Orders');
  if (!sheet) throw new Error('ไม่พบ Orders sheet');
  var rowData = sheet.getRange(row, 1, 1, 16).getValues()[0];
  var oldStatus = rowData[5];
  var newStatus = 'รอตรวจสอบ'; // คืนสถานะ ให้ลูกค้าส่งสลิปใหม่

  logAuditEvent('order', rowData[0], rowData[2], oldStatus, 'ปฏิเสธสลิป', adminName, reason || '');
  _notifySlipResult(rowData, 'ปฏิเสธสลิป', adminName, false, reason);
  return { ok: true };
}

function _notifySlipResult(rowData, status, adminName, isApproved, reason) {
  var orderId  = rowData[0];
  var name     = rowData[2];
  var phone    = rowData[1];
  var email    = rowData[8];
  var details  = rowData[3];
  var total    = rowData[4];

  // LINE ลูกค้า
  var lineUid = findCustomerLineId(phone, email);
  if (lineUid) {
    var lineMsg = isApproved
      ? '✅ สลิปได้รับการยืนยันแล้ว!\n'
        + '━━━━━━━━━━━━━━━\n'
        + '🆔 รหัส: ' + orderId + '\n'
        + '📦 สถานะ: กำลังเตรียมจัดส่ง\n'
        + '💡 พิมพ์ "ติดตามสถานะ" เพื่อติดตามออเดอร์ครับ'
      : '❌ สลิปไม่ผ่านการตรวจสอบ\n'
        + '━━━━━━━━━━━━━━━\n'
        + '🆔 รหัส: ' + orderId + '\n'
        + (reason ? '📝 เหตุผล: ' + reason + '\n' : '')
        + '💡 กรุณาติดต่อทีมงานเพื่อชำระเงินใหม่ครับ';
    sendLineMessage(lineUid, lineMsg);
  }

  // Email ลูกค้า
  if (email) {
    var subject = isApproved
      ? '[Media EDU Club] ✅ ยืนยันการรับสลิป — ' + orderId
      : '[Media EDU Club] ❌ สลิปไม่ผ่านการตรวจสอบ — ' + orderId;
    var emailRows = [
      ['รหัสออเดอร์', orderId],
      ['รายการสินค้า', String(details).replace(/\n/g, '<br>')],
      ['ยอดรวม', Number(total).toLocaleString() + ' บาท'],
      ['สถานะ', isApproved ? 'กำลังเตรียมจัดส่ง' : 'ปฏิเสธสลิป'],
      ['ดำเนินการโดย', adminName]
    ];
    if (!isApproved && reason) emailRows.push(['เหตุผล', reason]);
    var body = buildEmailHtml({
      title:      isApproved ? 'สลิปได้รับการยืนยัน' : 'สลิปไม่ผ่านการตรวจสอบ',
      headerIcon: 'S',
      color:      isApproved ? '#166534' : '#b91c1c',
      name:       name,
      intro:      isApproved
        ? 'ทีมงานได้ตรวจสอบสลิปโอนเงินของคุณเรียบร้อยแล้ว และกำลังดำเนินการจัดส่ง'
        : 'ทีมงานพบปัญหาเกี่ยวกับสลิปโอนเงินที่คุณแนบมา กรุณาติดต่อทีมงานเพื่อดำเนินการต่อ',
      rows: emailRows,
      note: 'หากมีข้อสงสัย กรุณาติดต่อทีมงาน Media EDU Club CMU.'
    });
    sendConfirmEmail(email, subject, body);
  }
}

// ==========================================
// Edit / Delete Order (Admin)
// ==========================================

function editOrderData(row, data, adminName) {
  var sheet = getDB().getSheetByName('Orders');
  if (!sheet) throw new Error('ไม่พบ Orders sheet');
  var oldData = sheet.getRange(row, 1, 1, 16).getValues()[0];
  var oldStatus = oldData[5];

  // อัปเดตทีละ column
  sheet.getRange(row, 3).setValue(data.name    || oldData[2]);  // ชื่อ
  sheet.getRange(row, 2).setValue(data.phone   || oldData[1]);  // เบอร์
  sheet.getRange(row, 4).setValue(data.details || oldData[3]);  // รายการ
  sheet.getRange(row, 5).setValue(data.total   || oldData[4]);  // ยอดรวม
  sheet.getRange(row, 6).setValue(data.status  || oldData[5]);  // สถานะ
  sheet.getRange(row, 8).setValue(adminName);                   // approver
  sheet.getRange(row, 9).setValue(data.email   || oldData[8]);  // email
  sheet.getRange(row, 10).setValue(data.address || oldData[9]); // address
  sheet.getRange(row, 16).setValue(data.note   || oldData[15]); // note

  // Log
  logAuditEvent('order', oldData[0], data.name || oldData[2],
    oldStatus, data.status || oldStatus, adminName, 'แก้ไขข้อมูล');

  // ถ้าสถานะเปลี่ยน → แจ้งลูกค้า
  if (data.status && data.status !== oldStatus) {
    var email = data.email || oldData[8];
    var phone = data.phone || oldData[1];
    var lineUid = findCustomerLineId(phone, email);
    if (lineUid) {
      var lineMsg = '📦 สถานะออเดอร์อัปเดต\n'
        + '━━━━━━━━━━━━━━━\n'
        + '🆔 รหัส: ' + oldData[0] + '\n'
        + '📊 สถานะใหม่: ' + data.status + '\n'
        + '💡 พิมพ์ "ติดตามสถานะ" เพื่อดูรายละเอียดครับ';
      sendLineMessage(lineUid, lineMsg);
    }
    if (email) {
      var subject = '[Media EDU Club] อัปเดตสถานะออเดอร์ ' + oldData[0] + ' - ' + data.status;
      var body = buildEmailHtml({
        title: getStatusTitle(data.status),
        headerIcon: 'S', color: getStatusColor(data.status),
        name: data.name || oldData[2],
        intro: 'ทีมงานได้อัปเดตข้อมูลคำสั่งซื้อของคุณ',
        rows: [
          ['รหัสออเดอร์', oldData[0]],
          ['สถานะใหม่', data.status],
          ['ดำเนินการโดย', adminName]
        ],
        note: 'หากมีข้อสงสัย กรุณาติดต่อทีมงาน Media EDU Club CMU.'
      });
      sendConfirmEmail(email, subject, body);
    }
  }
  return { ok: true };
}

function deleteOrderRow(row, orderId, adminName) {
  var sheet = getDB().getSheetByName('Orders');
  if (!sheet) throw new Error('ไม่พบ Orders sheet');
  var rowData = sheet.getRange(row, 1, 1, 16).getValues()[0];
  // ตรวจว่า Order ID ตรงกัน (ป้องกัน row shift)
  if (String(rowData[0]).trim() !== String(orderId).trim()) {
    throw new Error('Order ID ไม่ตรง (' + rowData[0] + ' vs ' + orderId + ') กรุณา refresh หน้าก่อนลบ');
  }
  logAuditEvent('order', orderId, rowData[2], rowData[5], 'ลบ', adminName, 'ลบออเดอร์');
  sheet.deleteRow(row);
  return { ok: true };
}

// ==========================================
// Batch update order status (Frontend bulk save)
// ==========================================
function batchUpdateOrderStatus(rows, status, adminName) {
  var sheet = getDB().getSheetByName('Orders');
  if (!sheet) throw new Error('ไม่พบ Orders sheet');
  if (!rows || !rows.length) return { ok: true, updated: 0 };

  for (var i = 0; i < rows.length; i++) {
    var row = Number(rows[i]);
    if (!row || row < 2) continue;
    var oldStatus = sheet.getRange(row, 6).getValue();
    sheet.getRange(row, 6).setValue(status);
    sheet.getRange(row, 8).setValue(adminName);

    var rowData = sheet.getRange(row, 1, 1, 16).getValues()[0];
    logAuditEvent('order', rowData[0], rowData[2], oldStatus, status, adminName, '');

    // ส่งอีเมลแจ้งเตือน (ถ้ามี)
    var email = rowData[8];
    if (email) {
      var orderId = rowData[0], name = rowData[2], details = rowData[3], total = rowData[4];
      var subject = '[Media EDU Club] อัปเดตสถานะออเดอร์ ' + orderId + ' - ' + status;
      var body = buildEmailHtml({
        title:      getStatusTitle(status),
        headerIcon: 'S',
        color:      getStatusColor(status),
        name:       name,
        intro:      'ทีมงานได้ดำเนินการอัปเดตสถานะคำสั่งซื้อของคุณแล้ว',
        rows: [
          ['รหัสออเดอร์',  orderId],
          ['รายการสินค้า', String(details).replace(/\n/g, '<br>')],
          ['ยอดรวม',       Number(total).toLocaleString() + ' บาท'],
          ['สถานะใหม่',    status],
          ['ดำเนินการโดย', adminName]
        ],
        note: 'หากมีข้อสงสัย กรุณาติดต่อทีมงาน Media EDU Club CMU.'
      });
      sendConfirmEmail(email, subject, body);
    }
    Utilities.sleep(50); // ป้องกัน quota exceeded
  }
  return { ok: true, updated: rows.length };
}

// ==========================================
// Export orders to new Google Sheet
// ==========================================
function exportOrdersToSheet(headers, rows2d, label, sizeSummary) {
  var ss    = getDB();
  var title = 'Orders Export — ' + label + ' — ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
  var newSs = SpreadsheetApp.create(title);

  // ── Sheet 1: รายการออเดอร์ ──
  var sheet = newSs.getActiveSheet();
  sheet.setName('รายการออเดอร์');

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#0d2454').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);

  if (rows2d && rows2d.length > 0) {
    sheet.getRange(2, 1, rows2d.length, headers.length).setValues(rows2d);
    // alternate rows
    for (var i = 2; i <= rows2d.length + 1; i++) {
      if (i % 2 === 0) sheet.getRange(i, 1, 1, headers.length).setBackground('#f8faff');
    }
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  // Total row
  var totalRow = rows2d.length + 2;
  var totalVal = rows2d.reduce(function(s, r) { return s + (Number(r[8]) || 0); }, 0);
  var totalData = new Array(headers.length).fill('');
  totalData[0] = 'รวมทั้งหมด (' + rows2d.length + ' รายการ)';
  totalData[8] = totalVal;
  sheet.getRange(totalRow, 1, 1, headers.length).setValues([totalData]);
  sheet.getRange(totalRow, 1, 1, headers.length).setBackground('#e0e7ff').setFontWeight('bold');

  // ── Sheet 2: สรุปจำนวนสินค้า/ไซส์ ──
  if (sizeSummary && typeof sizeSummary === 'object') {
    var summarySheet = newSs.insertSheet('สรุปสินค้า/ไซส์');
    summarySheet.getRange(1, 1, 1, 2).setValues([['รายการ / ไซส์', 'จำนวน (ชิ้น)']]);
    summarySheet.getRange(1, 1, 1, 2)
      .setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);

    var keys = Object.keys(sizeSummary).sort();
    var summaryData = keys.map(function(k) { return [k, sizeSummary[k]]; });
    if (summaryData.length > 0) {
      summarySheet.getRange(2, 1, summaryData.length, 2).setValues(summaryData);
      for (var j = 2; j <= summaryData.length + 1; j++) {
        if (j % 2 === 0) summarySheet.getRange(j, 1, 1, 2).setBackground('#eff6ff');
      }
    }

    var totalPcs = keys.reduce(function(s, k) { return s + summaryData[keys.indexOf(k)][1]; }, 0);
    var lastRow = summaryData.length + 2;
    summarySheet.getRange(lastRow, 1, 1, 2).setValues([['รวมทั้งหมด', totalPcs]]);
    summarySheet.getRange(lastRow, 1, 1, 2).setBackground('#e0e7ff').setFontWeight('bold');
    summarySheet.autoResizeColumns(1, 2);
    summarySheet.setFrozenRows(1);
  }

  return newSs.getUrl();
}


function updateEquipRequestStatus(row, status, adminName, deliveryPersonId, rejectReason) {
  var sheet = getDB().getSheetByName('Equipment_Requests');
  var oldStatus = sheet.getRange(row, 12).getValue();
  sheet.getRange(row, 12).setValue(status);
  sheet.getRange(row, 13).setValue(adminName);

  if (deliveryPersonId !== undefined) {
    sheet.getRange(row, 15).setValue(deliveryPersonId || '');
  }

  var rowData    = sheet.getRange(row, 1, 1, 14).getValues()[0];
  logAuditEvent('equipment', rowData[0], rowData[2], oldStatus, status, adminName, rejectReason || '');
  var email      = rowData[13];
  var recordId   = rowData[0];
  var name       = rowData[2];
  var items      = rowData[8];
  var borrowTime = rowData[6];
  var returnTime = rowData[7];

  if (email) {
    var deliveryHtml = '';
    if (status.includes('อนุมัติ') && !status.includes('ไม่') && deliveryPersonId) {
      var dlvList = getDeliveryConfigList();
      for (var d = 0; d < dlvList.length; d++) {
        if (dlvList[d].id === deliveryPersonId) {
          var dv = dlvList[d];
          deliveryHtml = '<div style="margin:20px 0 0;">'
            + '<p style="margin:0 0 10px;font-weight:700;color:#1e293b;font-size:0.92em;border-left:4px solid #0891b2;padding-left:10px;">ข้อมูลการรับอุปกรณ์</p>'
            + '<table style="width:100%;border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;font-size:0.9em;">'
            + _dvRow('ชื่อผู้จ่ายของ', dv.name     || '-')
            + _dvRow('หน่วยงาน',       dv.org      || '-')
            + _dvRow('เบอร์โทรศัพท์', dv.phone    || '-')
            + _dvRow('เวลานัดรับ',     dv.time     || '-')
            + _dvRow('สถานที่รับ',     dv.location || '-')
            + '</table></div>';
          break;
        }
      }
    }

    var subjectE = '[Media EDU Club] อัปเดตสถานะการยืมอุปกรณ์ ' + recordId + ' - ' + status;

    // สร้าง QR block ถ้าสถานะเป็นอนุมัติ
    var qrHtml = '';
    var isApproved = status.includes('อนุมัติ') && !status.includes('ไม่');
    if (isApproved) {
      var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=200x200&data='
        + encodeURIComponent(recordId) + '&bgcolor=ffffff&color=1e293b&margin=8';
      var qrActionLabel = status === 'อนุมัติ/กำลังยืม'
        ? '[คืนอุปกรณ์] แสดง QR นี้ให้ Admin เพื่อยืนยันการคืนอุปกรณ์'
        : '[รับอุปกรณ์] แสดง QR นี้ให้ Admin เพื่อยืนยันการรับอุปกรณ์';
      qrHtml = '<div style="margin:22px 0 0; text-align:center;">'
        + '<p style="margin:0 0 12px; font-weight:700; color:#1e293b; font-size:0.92em; '
        + 'border-left:4px solid #6d28d9; padding-left:10px; text-align:left;">QR Code สำหรับยืนยันการรับอุปกรณ์</p>'
        + '<div style="display:inline-block; background:#faf5ff; border:2px solid #ddd6fe; '
        + 'border-radius:16px; padding:18px 22px;">'
        + '<img src="' + qrUrl + '" width="160" height="160" '
        + 'style="border-radius:8px; display:block; margin:0 auto 10px;" '
        + 'alt="QR Code ' + recordId + '">'
        + '<div style="font-family:monospace; font-size:1em; font-weight:700; color:#6d28d9; '
        + 'background:white; border:1px solid #ede9fe; border-radius:6px; padding:4px 14px; '
        + 'display:inline-block; margin-bottom:6px;">' + recordId + '</div>'
        + '<p style="margin:6px 0 0; font-size:0.82em; color:#7c3aed;">' + qrActionLabel + '</p>'
        + '</div>'
        + '<p style="margin:10px 0 0; font-size:0.78em; color:#94a3b8;">'
        + 'Admin จะสแกน QR นี้ผ่านระบบเพื่อยืนยันการรับ/คืนอุปกรณ์</p>'
        + '</div>';
    }

    var bodyE    = buildEmailHtml({
      title:      getStatusTitle(status),
      headerIcon: 'E',
      color:      getStatusColor(status),
      name:       name,
      intro:      'ทีมงานได้ดำเนินการอัปเดตสถานะคำขอยืมอุปกรณ์ของคุณแล้ว รายละเอียดมีดังนี้:',
      rows: [
        ['รหัสรายการ',    recordId],
        ['ชื่อ-นามสกุล',  name],
        ['รหัสนักศึกษา',  rowData[3]],
        ['เบอร์โทรศัพท์', rowData[4]],
        ['คณะ/สังกัด',    rowData[5]],
        ['รายการอุปกรณ์', String(items).replace(/\n/g, '<br>')],
        ['วันที่/เวลารับ', borrowTime],
        ['วันที่/เวลาคืน', returnTime],
        ['สถานะใหม่',     status],
        ['ดำเนินการโดย',  adminName]
      ].concat(rejectReason ? [['เหตุผล', rejectReason]] : []),
      extraHtml: deliveryHtml + qrHtml,
      note: 'กรุณาติดต่อผู้จ่ายของตามข้อมูลที่ระบุ หากมีข้อสงสัยกรุณาติดต่อทีมงาน Media EDU Club CMU.'
    });
    sendConfirmEmail(email, subjectE, bodyE);
  }

  // แจ้งเตือน LINE ผู้ยืม ค้นหาจากรหัสนักศึกษา
  notifyUserBySid(rowData[3], status, 'equip', recordId, adminName, deliveryPersonId);
}

function updateBookingStatus(row, status, adminName, rejectReason) {
  var sheet = getDB().getSheetByName('Bookings');
  var oldStatus = sheet.getRange(row, 15).getValue();
  sheet.getRange(row, 15).setValue(status);
  sheet.getRange(row, 17).setValue(adminName);

  var rowData = sheet.getRange(row, 1, 1, 17).getValues()[0];
  logAuditEvent('booking', rowData[0], rowData[2], oldStatus, status, adminName, rejectReason || '');
  var email   = rowData[3];
  if (email) {
    var bookingId = rowData[0], name = rowData[2], room = rowData[8];
    var date = rowData[9], startTime = rowData[10], endTime = rowData[11];
    var subjectB = '[Media EDU Club] อัปเดตสถานะการจองห้อง ' + bookingId + ' - ' + status;
    var bodyB    = buildEmailHtml({
      title:      getStatusTitle(status),
      headerIcon: 'R',
      color:      getStatusColor(status),
      name:       name,
      intro:      'ทีมงานได้ดำเนินการอัปเดตสถานะการจองห้องของคุณแล้ว รายละเอียดมีดังนี้:',
      rows: [
        ['รหัสการจอง',         bookingId],
        ['ชื่อ-นามสกุล',       name],
        ['รหัสนักศึกษา',       rowData[5]],
        ['เบอร์โทรศัพท์',      rowData[4]],
        ['ห้องที่จอง',          room],
        ['วันที่ใช้งาน',        date],
        ['เวลาเริ่ม – สิ้นสุด', startTime + ' – ' + endTime],
        ['จำนวนผู้ใช้งาน',     rowData[12] + ' คน'],
        ['วัตถุประสงค์',        rowData[13]],
        ['สถานะใหม่',          status],
        ['ดำเนินการโดย',        adminName]
      ].concat(rejectReason ? [['เหตุผล', rejectReason]] : []),
      note: 'กรุณาติดต่อทีมงานหากมีข้อสงสัยเกี่ยวกับการจองห้อง'
    });
    sendConfirmEmail(email, subjectB, bodyB);
  }

  // แจ้งเตือน LINE ผู้จอง ค้นหาจากรหัสนักศึกษา
  var bookingSid = rowData[5];
  notifyUserBySid(bookingSid, status, 'booking', rowData[0]);
}

function deleteRow(sheetName, id) {
  var sheet = getDB().getSheetByName(sheetName);
  if (!sheet) return 'ไม่พบชีต';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id) || String(data[i][1]) === String(id)) {
      sheet.deleteRow(i + 1);
      return 'ลบสำเร็จ';
    }
  }
  return 'ไม่พบรายการ';
}

// ==========================================
// 7. Helper: สร้างรหัสรันอัตโนมัติ
// ==========================================

function generateNextId(sheetName, prefix) {
  var sheet = getDB().getSheetByName(sheetName);
  if (!sheet) return prefix + '0001';
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return prefix + '0001';
  // วนหาจากบนลงล่างเพื่อหา ID ล่าสุดที่ตรงรูปแบบ
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxNum = 0;
  for (var i = 0; i < data.length; i++) {
    var id = String(data[i][0] || '').trim();
    if (id.indexOf(prefix) === 0) {
      var numPart = id.replace(prefix, '');
      var n = parseInt(numPart, 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
  }
  return prefix + (maxNum + 1).toString().padStart(4, '0');
}

// ==========================================
// 8. MODULE: ปฏิทินจองห้อง
// ==========================================

function getCalendarBookings() {
  var sheet = getDB().getSheetByName('Bookings');
  if (!sheet) return [];
  var data     = sheet.getDataRange().getDisplayValues();
  var bookings = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][14] === 'รออนุมัติ' || data[i][14] === 'อนุมัติแล้ว') {
      bookings.push({
        room:      data[i][8],
        date:      data[i][9],
        startTime: data[i][10],
        endTime:   data[i][11],
        name:      data[i][2],
        purpose:   data[i][13],
        status:    data[i][14],
        approver:  data[i][16] || '-'
      });
    }
  }
  return bookings;
}

// ==========================================
// 9. MODULE: จัดการคลังอุปกรณ์ (Equipment_List)
// ==========================================

function getDynamicInventory() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Equipment_List');

  if (!sheet) {
    sheet = ss.insertSheet('Equipment_List');
    sheet.appendRow(['รหัส', 'สังกัด (media/samo)', 'หมวดหมู่', 'ชื่ออุปกรณ์', 'สถานะ', 'จำนวน']);
    sheet.appendRow(['EQ001', 'media', 'หมวดไฟสตูดิโอ',      'Nanlite Pavotube II 15C LED', 'พร้อมใช้', 2]);
    sheet.appendRow(['EQ002', 'media', 'หมวดอุปกรณ์สตูดิโอ', 'สาย HDMI',                   'พร้อมใช้', 5]);
    sheet.appendRow(['EQ003', 'samo',  'หมวดเครื่องเสียง',   'ลำโพงบลูทูธ (เล็ก)',          'พร้อมใช้', 1]);
  }

  var data      = sheet.getDataRange().getDisplayValues();
  var inventory = {
    'media': { title: 'Media EDU Club CMU.',                   items: {} },
    'samo':  { title: 'สโมสรนักศึกษาคณะศึกษาศาสตร์ มช.', items: {} }
  };

  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    if (!row[0]) continue;
    var org    = (row[1] || '').toLowerCase().trim();
    var cat    = row[2] || '';
    var name   = row[3] || '';
    var status = row[4] || 'พร้อมใช้';
    var qty    = parseInt(row[5]) > 0 ? parseInt(row[5]) : 1;

    if (status !== 'พร้อมใช้') continue;
    if (!cat || !name) continue;

    if (inventory[org]) {
      if (!inventory[org].items[cat]) inventory[org].items[cat] = [];
      inventory[org].items[cat].push({ name: name, qty: qty });
    }
  }
  return inventory;
}

function getAdminInventory() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Equipment_List');

  if (!sheet) {
    sheet = ss.insertSheet('Equipment_List');
    sheet.appendRow(['รหัส', 'สังกัด (media/samo)', 'หมวดหมู่', 'ชื่ออุปกรณ์', 'สถานะ', 'จำนวน']);
  } else {
    var lastCol = sheet.getLastColumn();
    var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (lastCol < 5 || !headers[4]) sheet.getRange(1, 5).setValue('สถานะ');
    if (lastCol < 6 || !headers[5]) sheet.getRange(1, 6).setValue('จำนวน');

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var data2 = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = 0; i < data2.length; i++) {
        if (!data2[i][0]) continue;
        if (!data2[i][4]) sheet.getRange(i + 2, 5).setValue('พร้อมใช้');
        if (!data2[i][5] || isNaN(parseInt(data2[i][5]))) sheet.getRange(i + 2, 6).setValue(1);
      }
    }
  }

  var data  = sheet.getDataRange().getDisplayValues();
  var items = [];
  for (var j = data.length - 1; j > 0; j--) {
    var row = data[j];
    if (!row[0]) continue;
    items.push({
      id:     row[0] || '',
      org:    row[1] === 'media' ? 'Media EDU' : 'สโมสรฯ',
      orgRaw: row[1] || 'media',
      cat:    row[2] || '',
      name:   row[3] || '',
      status: row[4] || 'พร้อมใช้',
      qty:    parseInt(row[5]) > 0 ? parseInt(row[5]) : 1
    });
  }
  return items;
}

function debugInventory() {
  var sheet = getDB().getSheetByName('Equipment_List');
  if (!sheet) { Logger.log('❌ ไม่พบ Sheet Equipment_List'); return; }
  Logger.log('✅ พบ Sheet: ' + sheet.getName());
  Logger.log('Rows: ' + sheet.getLastRow() + ', Cols: ' + sheet.getLastColumn());
  var data = sheet.getDataRange().getValues();
  Logger.log('Header row: ' + JSON.stringify(data[0]));
  if (data.length > 1) Logger.log('Row 2: ' + JSON.stringify(data[1]));
  Logger.log('getAdminInventory result: ' + JSON.stringify(getAdminInventory().slice(0, 3)));
}

function addEquipmentItem(org, category, name, status, qty) {
  var sheet = getDB().getSheetByName('Equipment_List');
  if (!sheet) return 'ไม่พบชีต Equipment_List';
  var id = generateNextId('Equipment_List', 'EQ');
  sheet.appendRow([id, org, category, name, status || 'พร้อมใช้', parseInt(qty) || 1]);
  return 'เพิ่มอุปกรณ์สำเร็จ: ' + name;
}

function updateEquipmentItem(id, org, category, name, status, qty) {
  var sheet = getDB().getSheetByName('Equipment_List');
  if (!sheet) return 'ไม่พบชีต Equipment_List';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      if (org      !== null && org      !== undefined) sheet.getRange(i + 1, 2).setValue(org);
      if (category !== null && category !== undefined) sheet.getRange(i + 1, 3).setValue(category);
      if (name     !== null && name     !== undefined) sheet.getRange(i + 1, 4).setValue(name);
      if (status   !== null && status   !== undefined) sheet.getRange(i + 1, 5).setValue(status);
      if (qty      !== null && qty      !== undefined) sheet.getRange(i + 1, 6).setValue(parseInt(qty) || 1);
      return 'แก้ไขสำเร็จ';
    }
  }
  return 'ไม่พบรหัส ' + id;
}

// ==========================================
// 10. MODULE: ระบบส่งอีเมล (Email Notifications)
// ==========================================

function sendConfirmEmail(to, subject, htmlBody) {
  try {
    GmailApp.sendEmail(to, subject, '', {
      htmlBody: htmlBody,
      from:     'media.edcmu@gmail.com',
      name:     'Media EDU Club CMU.'
    });
  } catch (e) {
    Logger.log('sendConfirmEmail error: ' + e.message);
  }
}

var BORROW_TERMS = [
  'ผู้ยืมต้องรับผิดชอบต่อความเสียหายหรือสูญหายของอุปกรณ์ทุกชิ้นที่ยืมไป',
  'ต้องคืนอุปกรณ์ในสภาพสมบูรณ์ภายในวันและเวลาที่กำหนด',
  'หากคืนเกินกำหนด จะถูกระงับสิทธิ์การยืมในครั้งถัดไป',
  'ห้ามนำอุปกรณ์ไปให้บุคคลอื่นใช้โดยไม่ได้รับอนุญาต',
  'กรณีเกิดความเสียหาย ผู้ยืมต้องแจ้งเจ้าหน้าที่ทันที',
  'อุปกรณ์ที่ยืมต้องใช้เพื่อวัตถุประสงค์ที่ระบุไว้เท่านั้น'
];

function buildTermsHtml(color) {
  var items = BORROW_TERMS.map(function(t, i) {
    return '<li style="padding:5px 0; color:#475569; font-size:0.88em;">' + (i + 1) + '. ' + t + '</li>';
  }).join('');
  return '<div style="margin-top:24px; border:1px solid #e2e8f0; border-radius:10px; overflow:hidden;">'
    + '<div style="background:' + color + '; padding:10px 16px;">'
    + '<strong style="color:white; font-size:0.9em;">ข้อกำหนดและเงื่อนไขการยืมอุปกรณ์</strong></div>'
    + '<div style="padding:14px 16px; background:#fffbf0;">'
    + '<ul style="margin:0; padding-left:0; list-style:none;">' + items + '</ul>'
    + '</div></div>';
}

function buildEmailHtml(opts) {
  var rows = (opts.rows || []).map(function(r) {
    return '<tr>'
      + '<td style="padding:10px 14px; font-weight:600; color:#475569; background:#f8fafc; width:38%; border-bottom:1px solid #f1f5f9; white-space:nowrap;">' + r[0] + '</td>'
      + '<td style="padding:10px 14px; color:#2d3748; border-bottom:1px solid #f1f5f9;">' + r[1] + '</td>'
      + '</tr>';
  }).join('');

  var termsSection = opts.showTerms ? buildTermsHtml(opts.color || '#007bff') : '';
  var noteSection  = opts.note
    ? '<div style="margin-top:20px; padding:14px 16px; background:#f0f6ff; border-left:4px solid ' + (opts.color || '#007bff') + '; border-radius:0 8px 8px 0; font-size:0.88em; color:#475569;">หมายเหตุ: ' + opts.note + '</div>'
    : '';
  var headerIcon = opts.headerIcon || 'M';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>'
    + '<body style="margin:0; padding:0; background:#f4f8fb; font-family: Arial, sans-serif;">'
    + '<div style="max-width:580px; margin:30px auto; background:white; border-radius:16px; overflow:hidden; box-shadow:0 4px 20px rgba(0,0,0,0.08);">'
    + '<div style="background:linear-gradient(135deg, ' + (opts.color || '#007bff') + ', #1a4480); padding:32px 30px; text-align:center;">'
    + '<div style="width:60px; height:60px; background:rgba(255,255,255,0.2); border-radius:50%; display:inline-flex; align-items:center; justify-content:center; margin-bottom:14px;">'
    + '<span style="color:white; font-size:1.6em; font-weight:bold;">' + headerIcon + '</span></div>'
    + '<h1 style="margin:0; color:white; font-size:1.3em; font-weight:700; line-height:1.4;">' + opts.title + '</h1>'
    + '<p style="margin:8px 0 0 0; color:rgba(255,255,255,0.8); font-size:0.88em;">Media EDU Club CMU.</p>'
    + '</div>'
    + '<div style="padding:28px 30px;">'
    + '<p style="margin:0 0 20px 0; color:#475569; font-size:0.95em; line-height:1.6;">เรียน <strong>' + (opts.name || 'ผู้ใช้งาน') + '</strong>,<br>' + (opts.intro || 'ระบบได้รับข้อมูลของคุณเรียบร้อยแล้ว') + '</p>'
    + '<table style="width:100%; border-collapse:collapse; border-radius:10px; overflow:hidden; border:1px solid #e2e8f0; font-size:0.9em;">' + rows + '</table>'
    + (opts.extraHtml || '') + termsSection + noteSection
    + '</div>'
    + '<div style="background:#f8fafc; padding:20px 30px; text-align:center; border-top:1px solid #e2e8f0;">'
    + '<p style="margin:0 0 6px 0; font-size:0.82em; color:#64748b;">Media EDU Club CMU. | คณะศึกษาศาสตร์ มหาวิทยาลัยเชียงใหม่</p>'
    + '<p style="margin:0; font-size:0.78em; color:#94a3b8;">อีเมลนี้ถูกส่งโดยอัตโนมัติ กรุณาอย่าตอบกลับ</p>'
    + '</div>'
    + '</div></body></html>';
}

function getStatusTitle(status) {
  var s = (status || '').trim();
  if (s.includes('อนุมัติแล้ว') || s.includes('อนุมัติ/กำลัง') || s.includes('กำลังเตรียม')) return 'ได้รับการอนุมัติแล้ว';
  if (s.includes('จัดส่งแล้ว')) return 'จัดส่งเรียบร้อยแล้ว';
  if (s.includes('คืนแล้ว'))    return 'รับคืนอุปกรณ์เรียบร้อย';
  if (s.includes('ยกเลิก') || s.includes('ไม่อนุมัติ')) return 'ไม่ผ่านการอนุมัติ';
  return 'อัปเดตสถานะ: ' + status;
}

function getStatusColor(status) {
  var s = (status || '').trim();
  if (s.includes('อนุมัติ') && !s.includes('ไม่')) return '#2e7d32';
  if (s.includes('จัดส่ง') || s.includes('คืนแล้ว')) return '#1565c0';
  if (s.includes('ยกเลิก') || s.includes('ไม่อนุมัติ')) return '#c62828';
  return '#007bff';
}

// ==========================================
// 11. MODULE: ระบบจัดการห้อง (Room_List)
// ==========================================

function getRoomList() {
  return withCache('rooms', CACHE_TTL.rooms, getRoomList_raw_);
}

function getRoomList_raw_() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Room_List');
  if (!sheet) {
    sheet = ss.insertSheet('Room_List');
    sheet.appendRow(['รหัสห้อง', 'ชื่อห้อง', 'ความจุ (คน)', 'อุปกรณ์/สเปค', 'หมายเหตุ', 'สถานะ']);
    sheet.appendRow(['R001', 'Studio Room',      '6',  'กล้อง DSLR, ฉากหลัง, ไฟสตูดิโอ, ไมโครโฟน',      'ห้องบันทึกภาพ/เสียง', 'เปิดให้จอง']);
    sheet.appendRow(['R002', 'Meeting Room A',   '6',  'โปรเจคเตอร์, กระดานไวท์บอร์ด, WiFi',             'ห้องประชุมขนาดเล็ก',  'เปิดให้จอง']);
    sheet.appendRow(['R003', 'Meeting Room B',   '10', 'โปรเจคเตอร์, กระดานไวท์บอร์ด, WiFi, ระบบเสียง',  'ห้องประชุมขนาดกลาง', 'เปิดให้จอง']);
    sheet.appendRow(['R004', 'Co-Working Space', '20', 'โต๊ะทำงาน, WiFi, ปลั๊กไฟ',                       'โซนทำงานแบบเปิด',    'เปิดให้จอง']);
  }
  var data  = sheet.getDataRange().getDisplayValues();
  var rooms = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) rooms.push({
      row:      i + 1,
      id:       data[i][0],
      name:     data[i][1],
      capacity: data[i][2],
      specs:    data[i][3],
      note:     data[i][4],
      status:   data[i][5] || 'เปิดให้จอง'
    });
  }
  return rooms;
}

function saveRoom(payload) {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Room_List');
  if (!sheet) { getRoomList(); sheet = ss.getSheetByName('Room_List'); }

  if (payload.row) {
    var r = parseInt(payload.row);
    sheet.getRange(r, 2).setValue(payload.name);
    sheet.getRange(r, 3).setValue(payload.capacity);
    sheet.getRange(r, 4).setValue(payload.specs);
    sheet.getRange(r, 5).setValue(payload.note);
    sheet.getRange(r, 6).setValue(payload.status);
    return 'แก้ไขห้องสำเร็จ: ' + payload.name;
  } else {
    var id = generateNextId('Room_List', 'R');
    sheet.appendRow([id, payload.name, payload.capacity, payload.specs, payload.note, payload.status || 'เปิดให้จอง']);
    return 'เพิ่มห้องสำเร็จ: ' + payload.name;
  }
}

function deleteRoom(row) {
  var sheet = getDB().getSheetByName('Room_List');
  if (!sheet) return 'ไม่พบชีต';
  sheet.deleteRow(parseInt(row));
  return 'ลบสำเร็จ';
}

function updateBookingFull(row, payload) {
  var sheet = getDB().getSheetByName('Bookings');
  if (!sheet) return 'ไม่พบชีต';
  sheet.getRange(row, 3).setValue(payload.name);
  sheet.getRange(row, 4).setValue(payload.email);
  sheet.getRange(row, 5).setValue(payload.phone);
  sheet.getRange(row, 6).setValue(payload.sid);
  sheet.getRange(row, 7).setValue(payload.major);
  sheet.getRange(row, 8).setValue(payload.faculty);
  sheet.getRange(row, 9).setValue(payload.room);
  sheet.getRange(row, 10).setValue(payload.date);
  sheet.getRange(row, 11).setValue(payload.startTime);
  sheet.getRange(row, 12).setValue(payload.endTime);
  sheet.getRange(row, 13).setValue(payload.userCount);
  sheet.getRange(row, 14).setValue(payload.purpose);
  sheet.getRange(row, 15).setValue(payload.status);
  sheet.getRange(row, 17).setValue(payload.adminName || 'Admin');
  return 'แก้ไขการจองสำเร็จ';
}

// ==========================================
// 12. MODULE: Admin Dashboard Stats
// ==========================================

function getDashboardStats() {
  return withCache('dashboard', CACHE_TTL.dashboard, getDashboardStats_raw_);
}

function getDashboardStats_raw_() {
  var ss        = getDB();
  var now       = new Date();
  var thisMonth = now.getMonth();
  var thisYear  = now.getFullYear();

  function countSheet(name, statusCol, pendingVals) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return { pending: 0, total: 0, thisMonth: 0 };
    var data = sheet.getDataRange().getValues();
    var pending = 0, total = 0, month = 0;
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      total++;
      var st = String(data[i][statusCol] || '').trim();
      for (var j = 0; j < pendingVals.length; j++) {
        if (st === pendingVals[j]) { pending++; break; }
      }
      try {
        var ts = data[i][1];
        var d  = (ts instanceof Date) ? ts : new Date(ts);
        if (!isNaN(d.getTime()) && d.getMonth() === thisMonth && d.getFullYear() === thisYear) month++;
      } catch(e) {}
    }
    return { pending: pending, total: total, thisMonth: month };
  }

  var orders   = countSheet('Orders',             5,  ['รอดำเนินการ', 'รอตรวจสอบ']);
  var equip    = countSheet('Equipment_Requests', 11, ['รอดำเนินการ', 'รอตรวจสอบ', 'รออนุมัติ']);
  var bookings = countSheet('Bookings',           14, ['รออนุมัติ']);

  var brokenCount = 0;
  var eqSheet     = ss.getSheetByName('Equipment_List');
  if (eqSheet) {
    var ed = eqSheet.getDataRange().getValues();
    for (var i = 1; i < ed.length; i++) {
      if (!ed[i][0]) continue;
      var st = String(ed[i][4] || '').trim();
      if (st === 'ชำรุด' || st === 'งดใช้') brokenCount++;
    }
  }

  var msgPending = 0, msgTotal = 0;
  var msgSheet = ss.getSheetByName('Contact_Messages');
  if (msgSheet) {
    var md = msgSheet.getDataRange().getValues();
    for (var i = 1; i < md.length; i++) {
      if (!md[i][0]) continue;
      msgTotal++;
      var mst = String(md[i][6] || 'ใหม่').trim();
      if (mst === 'ใหม่' || mst === '') msgPending++;
    }
  }

  return { orders: orders, equip: equip, bookings: bookings, brokenCount: brokenCount, msgPending: msgPending, msgTotal: msgTotal };
}

// ==========================================
// 13. MODULE: ระบบอีเมลหน่วยงาน (Email_Config)
// ==========================================

function getOrCreateEmailConfigSheet() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Email_Config');
  if (!sheet) {
    sheet = ss.insertSheet('Email_Config');
    sheet.appendRow(['รหัสหน่วยงาน', 'ชื่อหน่วยงาน', 'อีเมล']);
    sheet.appendRow(['admin', 'Admin ระบบ',                           '']);
    sheet.appendRow(['media', 'Media EDU Club CMU.',                  '']);
    sheet.appendRow(['samo',  'สโมสรนักศึกษาคณะศึกษาศาสตร์ มช.', '']);
    var header = sheet.getRange(1, 1, 1, 3);
    header.setBackground('#1a3a6b').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 280);
    sheet.setColumnWidth(3, 260);
    Logger.log('สร้าง Email_Config sheet เรียบร้อยแล้ว กรุณากรอกอีเมลในช่อง col C');
  }
  return sheet;
}

function getEmailConfigMap() {
  var sheet = getOrCreateEmailConfigSheet();
  var data  = sheet.getDataRange().getValues();
  var map   = {};
  for (var i = 1; i < data.length; i++) {
    var key   = String(data[i][0] || '').trim().toLowerCase();
    var name  = String(data[i][1] || '').trim();
    var email = String(data[i][2] || '').trim();
    if (key) map[key] = { name: name, email: email };
  }
  return map;
}

function getEmailConfigRows() {
  var sheet = getOrCreateEmailConfigSheet();
  var data  = sheet.getDataRange().getValues();
  var rows  = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push([
      String(data[i][0] || '').trim(),
      String(data[i][1] || '').trim(),
      String(data[i][2] || '').trim()
    ]);
  }
  return rows;
}

function updateEmailConfigRow(key, email) {
  var sheet = getOrCreateEmailConfigSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === key.toLowerCase()) {
      sheet.getRange(i + 1, 3).setValue(email);
      return true;
    }
  }
  sheet.appendRow([key, key, email]);
  return true;
}

function sendOrgNotificationEmail(payload, recordId, orgs) {
  var emailMap  = getEmailConfigMap();
  var adminInfo = emailMap['admin'] || { name: 'Admin', email: '' };

  var itemLines = payload.items.split('\n');
  var itemsHtml = itemLines.map(function(line) {
    var l = line.trim();
    if (!l) return '';
    if (l.startsWith('[ยืม:')) {
      return '<tr><td colspan="2" style="padding:8px 14px; background:#eef4ff; font-weight:700; color:#1a3a6b; font-size:0.88em;">' + l + '</td></tr>';
    }
    return '<tr><td style="padding:7px 14px; color:#2d3748; border-bottom:1px solid #f1f5f9;">• ' + l + '</td></tr>';
  }).filter(Boolean).join('');

  var itemsTable  = '<table style="width:100%; border-collapse:collapse; border:1px solid #e2e8f0; border-radius:8px; overflow:hidden; font-size:0.9em;">' + itemsHtml + '</table>';
  var commonRows  = [
    ['รหัสรายการ',           recordId],
    ['ชื่อ-นามสกุลผู้ยืม',   payload.name],
    ['รหัสนักศึกษา/บุคลากร', payload.sid     || '-'],
    ['เบอร์โทรศัพท์',        payload.phone   || '-'],
    ['คณะ/สังกัด',           payload.faculty || '-'],
    ['วัตถุประสงค์',          payload.purpose || '-'],
    ['วันที่/เวลารับ',        payload.borrowTime],
    ['วันที่/เวลาคืน',        payload.returnTime],
    ['อีเมลผู้ยืม',           payload.email   || '-'],
    ['วันเวลาที่ส่งคำขอ',     getThaiTimestamp()]
  ];

  var rowsHtml = commonRows.map(function(r) {
    return '<tr>'
      + '<td style="padding:9px 14px; font-weight:600; color:#475569; background:#f8fafc; width:40%; border-bottom:1px solid #f1f5f9; white-space:nowrap;">' + r[0] + '</td>'
      + '<td style="padding:9px 14px; color:#2d3748; border-bottom:1px solid #f1f5f9;">' + r[1] + '</td>'
      + '</tr>';
  }).join('');

  var detailTable     = '<table style="width:100%; border-collapse:collapse; border-radius:10px; overflow:hidden; border:1px solid #e2e8f0; font-size:0.9em;">' + rowsHtml + '</table>';
  var deliveryList    = getActiveDeliveryInfo();
  var deliverySection = '';
  if (deliveryList.length > 0) {
    var dvRows = deliveryList.map(function(d, idx) {
      var badge = '<span style="display:inline-block;background:#1a3a6b;color:white;border-radius:20px;padding:1px 10px;font-size:0.78em;margin-right:6px;">' + (idx + 1) + '</span>';
      return '<tr style="border-bottom:1px solid #f1f5f9;">'
        + '<td style="padding:10px 14px; vertical-align:top; width:32px;">' + badge + '</td>'
        + '<td style="padding:10px 14px;">'
        + '<div style="font-weight:700; color:#1e293b; font-size:0.93em;">' + (d.name || '-') + '</div>'
        + '<div style="color:#64748b; font-size:0.84em; margin-top:3px;">'
        + '<b>หน่วยงาน:</b> ' + (d.org || '-') + ' &nbsp;|&nbsp; '
        + '<b>โทร:</b> ' + (d.phone || '-') + ' &nbsp;|&nbsp; '
        + '<b>สถานที่:</b> ' + (d.location || '-')
        + '</div></td></tr>';
    }).join('');
    deliverySection = '<div style="margin-bottom:20px;">'
      + '<p style="margin:0 0 10px;font-weight:700;color:#1e293b;font-size:0.92em;border-left:4px solid #0891b2;padding-left:10px;">ข้อมูลผู้จ่ายของ / นัดรับอุปกรณ์</p>'
      + '<table style="width:100%;border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;font-size:0.9em;">'
      + dvRows + '</table></div>';
  }

  var recipients = [];
  if (orgs && orgs.length > 0) {
    orgs.forEach(function(orgKey) {
      var info = emailMap[orgKey];
      if (info && info.email) recipients.push({ email: info.email, name: info.name, isAdmin: false });
    });
  }
  if (adminInfo.email) recipients.push({ email: adminInfo.email, name: adminInfo.name, isAdmin: true });

  if (recipients.length === 0) {
    Logger.log('sendOrgNotificationEmail: ไม่มีอีเมลหน่วยงาน/admin ที่กำหนดไว้ใน Email_Config');
    return;
  }

  recipients.forEach(function(rec) {
    var orgLabel = rec.isAdmin ? 'มีคำขอยืมอุปกรณ์เข้ามาใหม่' : 'มีการยืมอุปกรณ์ของ ' + rec.name;
    var subject  = '[Media EDU Club] ' + orgLabel + ' | รหัส ' + recordId;

    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
      + '<body style="margin:0;padding:0;background:#f4f8fb;font-family:Arial,sans-serif;">'
      + '<div style="max-width:600px;margin:30px auto;background:white;border-radius:16px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);">'
      + '<div style="background:linear-gradient(135deg,#1a3a6b,#0a5eb0);padding:30px;text-align:center;">'
      + '<h1 style="margin:0;color:white;font-size:1.2em;font-weight:700;">แจ้งเตือน: คำขอยืมอุปกรณ์ใหม่</h1>'
      + '<p style="margin:6px 0 0;color:rgba(255,255,255,0.75);font-size:0.85em;">Media EDU Club CMU. | ระบบแจ้งเตือนอัตโนมัติ</p>'
      + '</div>'
      + '<div style="padding:28px 28px 20px;">'
      + '<p style="margin:0 0 18px;color:#475569;font-size:0.94em;line-height:1.7;">เรียน <strong>' + rec.name + '</strong>,<br>'
      + orgLabel + ' รหัสรายการ <strong style="color:#1a3a6b;">' + recordId + '</strong> โปรดตรวจสอบและดำเนินการตามขั้นตอน</p>'
      + '<div style="margin-bottom:20px;"><p style="margin:0 0 10px;font-weight:700;color:#1e293b;font-size:0.92em;border-left:4px solid #1a3a6b;padding-left:10px;">ข้อมูลผู้ยืม</p>' + detailTable + '</div>'
      + '<div style="margin-bottom:20px;"><p style="margin:0 0 10px;font-weight:700;color:#1e293b;font-size:0.92em;border-left:4px solid #6a1b9a;padding-left:10px;">รายการอุปกรณ์ที่ขอยืม</p>' + itemsTable + '</div>'
      + deliverySection
      + '<div style="padding:14px 16px;background:#fff8e1;border-left:4px solid #f59e0b;border-radius:0 8px 8px 0;font-size:0.87em;color:#78350f;">&#9888; กรุณาเข้าสู่ระบบ Admin เพื่ออนุมัติหรือปฏิเสธคำขอนี้</div>'
      + '</div>'
      + '<div style="background:#f8fafc;padding:18px 28px;text-align:center;border-top:1px solid #e2e8f0;">'
      + '<p style="margin:0 0 4px;font-size:0.8em;color:#64748b;">Media EDU Club CMU. | คณะศึกษาศาสตร์ มหาวิทยาลัยเชียงใหม่</p>'
      + '<p style="margin:0;font-size:0.76em;color:#94a3b8;">อีเมลนี้ถูกส่งโดยอัตโนมัติ กรุณาอย่าตอบกลับ</p>'
      + '</div></div></body></html>';

    sendConfirmEmail(rec.email, subject, html);
  });
}

// ==========================================
// 14. MODULE: ผู้จ่ายของ (Delivery_Config)
// ==========================================

function getOrCreateDeliverySheet() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Delivery_Config');
  if (!sheet) {
    sheet = ss.insertSheet('Delivery_Config');
    sheet.appendRow(['รหัส', 'ชื่อผู้จ่ายของ', 'หน่วยงาน', 'เวลานัด', 'เบอร์โทร', 'สถานที่']);
    var header = sheet.getRange(1, 1, 1, 6);
    header.setBackground('#1a3a6b').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidths(1, 6, 160);
    Logger.log('สร้าง Delivery_Config sheet เรียบร้อยแล้ว');
  }
  return sheet;
}

function getDeliveryConfigList() {
  return withCache('delivery', CACHE_TTL.delivery, getDeliveryConfigList_raw_);
}

function getDeliveryConfigList_raw_() {
  var sheet = getOrCreateDeliverySheet();
  var data  = sheet.getDataRange().getValues();
  var list  = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      id:       String(data[i][0] || '').trim(),
      name:     String(data[i][1] || '').trim(),
      org:      String(data[i][2] || '').trim(),
      time:     String(data[i][3] || '').trim(),
      phone:    String(data[i][4] || '').trim(),
      location: String(data[i][5] || '').trim(),
      row:      i + 1
    });
  }
  return list;
}

function addDeliveryConfig(d) {
  var sheet     = getOrCreateDeliverySheet();
  var id        = generateNextId('Delivery_Config', 'DLV');
  var newRowDlv = sheet.getLastRow() + 1;
  sheet.getRange(newRowDlv, 4).setNumberFormat('@');
  sheet.getRange(newRowDlv, 5).setNumberFormat('@');
  sheet.appendRow([id, d.name, d.org, d.time, d.phone, d.location]);
  return id;
}

function updateDeliveryConfig(row, d) {
  var sheet = getOrCreateDeliverySheet();
  sheet.getRange(row, 2).setValue(d.name);
  sheet.getRange(row, 3).setValue(d.org);
  sheet.getRange(row, 4).setValue(d.time);
  sheet.getRange(row, 5).setValue(d.phone);
  sheet.getRange(row, 6).setValue(d.location);
  return true;
}

function deleteDeliveryConfig(row) {
  var sheet = getOrCreateDeliverySheet();
  sheet.deleteRow(row);
  return true;
}

function getActiveDeliveryInfo() {
  var sheet = getOrCreateDeliverySheet();
  var data  = sheet.getDataRange().getValues();
  var list  = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      name:     String(data[i][1] || '').trim(),
      org:      String(data[i][2] || '').trim(),
      time:     String(data[i][3] || '').trim(),
      phone:    String(data[i][4] || '').trim(),
      location: String(data[i][5] || '').trim()
    });
  }
  return list;
}

// ==========================================
// 15. MODULE: Slides_Config
// ==========================================

function getOrCreateSlidesSheet() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Slides_Config');
  if (!sheet) {
    sheet = ss.insertSheet('Slides_Config');
    sheet.appendRow(['ลำดับ', 'imageUrl', 'หัวข้อ', 'คำอธิบาย', 'แสดง']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSlidesConfig() {
  return withCache('slides', CACHE_TTL.slides, getSlidesConfig_raw_);
}

function getSlidesConfig_raw_() {
  var sheet = getOrCreateSlidesSheet();
  var data  = sheet.getDataRange().getDisplayValues();
  var list  = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[1]) continue;
    list.push({
      order:    parseInt(row[0]) || i,
      imageUrl: String(row[1]).trim(),
      title:    String(row[2]).trim(),
      caption:  String(row[3]).trim(),
      active:   String(row[4]).trim().toUpperCase() !== 'FALSE'
    });
  }
  list.sort(function(a, b) { return a.order - b.order; });
  return list;
}

function addSlideConfig(d) {
  var sheet  = getOrCreateSlidesSheet();
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, 5).setValues([[
    d.order || newRow - 1,
    d.imageUrl || '',
    d.title    || '',
    d.caption  || '',
    d.active !== false ? 'TRUE' : 'FALSE'
  ]]);
  bustCacheKey('slides'); // ล้างแคชเพื่อให้เว็บอัปเดตทันที
  return true;
}

function updateSlideConfig(rowIndex, d) {
  var sheet    = getOrCreateSlidesSheet();
  var sheetRow = rowIndex + 2;
  if (sheetRow > sheet.getLastRow()) throw new Error('ไม่พบแถวที่ต้องการแก้ไข');
  sheet.getRange(sheetRow, 1, 1, 5).setValues([[
    d.order    || rowIndex + 1,
    d.imageUrl || '',
    d.title    || '',
    d.caption  || '',
    d.active !== false ? 'TRUE' : 'FALSE'
  ]]);
  bustCacheKey('slides'); // ล้างแคชเพื่อให้เว็บอัปเดตทันที
  return true;
}

function deleteSlideConfig(rowIndex) {
  var sheet    = getOrCreateSlidesSheet();
  var sheetRow = rowIndex + 2;
  if (sheetRow > sheet.getLastRow()) throw new Error('ไม่พบแถวที่ต้องการลบ');
  sheet.deleteRow(sheetRow);
  bustCacheKey('slides'); // ล้างแคชเพื่อให้เว็บอัปเดตทันที
  return true;
}

// ==========================================
// 16. MODULE: Equipment Availability Calendar
// ==========================================

function getEquipCalendarData() {
  var sheet = getDB().getSheetByName('Equipment_Requests');
  if (!sheet) return [];
  var data   = sheet.getDataRange().getDisplayValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var status     = String(row[11] || '').trim();
    var borrowRaw  = String(row[6]  || '').trim();
    var returnRaw  = String(row[7]  || '').trim();
    var borrowDate = borrowRaw.split(' ')[0] || borrowRaw;
    var returnDate = returnRaw.split(' ')[0] || returnRaw;
    if (!borrowDate) continue;
    result.push({
      id:          String(row[0]).trim(),
      name:        String(row[2]).trim(),
      sid:         String(row[3]).trim(),
      phone:       String(row[4]).trim(),
      faculty:     String(row[5]).trim(),
      borrowDate:  borrowDate,
      returnDate:  returnDate || borrowDate,
      borrowTime:  borrowRaw,
      returnTime:  returnRaw,
      items:       String(row[8]).trim(),
      purpose:     String(row[9]).trim(),
      status:      status,
      isCancelled: (status === 'ยกเลิก' || status === 'ปฏิเสธ' || status === 'ไม่อนุมัติ')
    });
  }
  return result;
}

function checkEquipAvailability(startDate, endDate) {
  if (!startDate) return { busy: [], records: [] };
  var end     = endDate || startDate;
  var allRaw  = getEquipCalendarData();
  var all     = allRaw.filter(function(r) { return !r.isCancelled; });
  var busySet = {};
  var records = [];
  all.forEach(function(r) {
    if (r.borrowDate <= end && r.returnDate >= startDate) {
      records.push(r);
      r.items.split('\n').forEach(function(line) {
        var name = line.replace(/^\d+\.\s*/, '').replace(/\s*x\d+.*/i, '').replace(/\[.*?\]/g, '').trim();
        if (name) busySet[name] = true;
      });
    }
  });
  return { busy: Object.keys(busySet), records: records };
}

// ==========================================
// 17. MODULE: LINE Notification Functions
// ==========================================

/** ส่งข้อความหาคนๆ เดียว */
function sendLineMessage(userId, message) {
  if (!userId || userId === 'Uxxxxxxxxxx') return; // ข้าม placeholder
  var url     = 'https://api.line.me/v2/bot/message/push';
  var options = {
    method:             'post',
    contentType:        'application/json',
    headers:            { 'Authorization': 'Bearer ' + LINE_TOKEN },
    payload:            JSON.stringify({ to: userId, messages: [{ type: 'text', text: message }] }),
    muteHttpExceptions: true
  };
  try {
    var res = UrlFetchApp.fetch(url, options);
    Logger.log('LINE push to ' + userId + ': HTTP ' + res.getResponseCode());
  } catch(e) {
    Logger.log('sendLineMessage error: ' + e.message);
  }
}

/** ส่งหา Admin ทุกคนพร้อมกัน */
function notifyAdmins(message) {
  notifyAdminsByScope(message, null);
}

/** แจ้ง Admin เมื่อมีคำขอยืมอุปกรณ์ใหม่ */
function notifyNewEquipRequest(data, recordId) {
  var msg = '🎥 มีคำขอยืมอุปกรณ์ใหม่!\n'
    + '━━━━━━━━━━━━━━━\n'
    + '🆔 รหัส: '    + recordId        + '\n'
    + '👤 ผู้ยืม: '   + data.name       + '\n'
    + '🏛️ สังกัด: '  + data.faculty     + '\n'
    + '📅 ยืม: '     + data.borrowTime  + '\n'
    + '📅 คืน: '     + data.returnTime  + '\n'
    + '📦 อุปกรณ์:\n' + data.items      + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '👉 กรุณาเข้าระบบเพื่ออนุมัติ';
  notifyAdminsByScope(msg, 'media'); // ส่งเฉพาะ admin ที่ดูแล media
}

/** แจ้ง Admin เมื่อมีคำขอจองห้องใหม่ */
function notifyNewBooking(data, recordId) {
  var msg = '🏫 มีคำขอจองห้องใหม่!\n'
    + '━━━━━━━━━━━━━━━\n'
    + '🆔 รหัส: '         + recordId       + '\n'
    + '👤 ผู้จอง: '        + data.name      + '\n'
    + '🚪 ห้อง: '          + data.room      + '\n'
    + '📅 วันที่: '         + data.date      + '\n'
    + '🕐 เวลา: '          + data.startTime + ' - ' + data.endTime + '\n'
    + '📝 วัตถุประสงค์: '  + data.purpose   + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '👉 กรุณาเข้าระบบเพื่ออนุมัติ';
  notifyAdmins(msg); // ห้อง = ส่งทุกคน
}

/** แจ้ง Admin เมื่อมีออเดอร์ใหม่ */
function notifyNewOrder(data, recordId, orderDetails, itemsTotal, shippingFee, totalPrice) {
  var msg = '🛒 มีออเดอร์ใหม่!\n'
    + '━━━━━━━━━━━━━━━\n'
    + '🆔 รหัส: '    + recordId   + '\n'
    + '👤 ลูกค้า: '  + data.name  + '\n'
    + '📱 โทร: '     + data.phone + '\n'
    + '🏷️ ประเภท: '  + (data.buyerType || '-') + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '🛍️ รายการ:\n'
    + (orderDetails || '').split('\n').map(function(l){ return '  • ' + l; }).join('\n') + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '💵 ยอดรวม: ' + Number(totalPrice || 0).toLocaleString() + ' ฿\n'
    + '🚚 จัดส่ง: '  + (data.shippingName || '-') + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '👉 กรุณาเข้าระบบเพื่อตรวจสอบสลิปและดำเนินการ';
  notifyAdmins(msg);
}

/** แจ้งผู้ใช้เมื่อผลอนุมัติออก โดย lookup LINE User ID จาก sid */
function notifyUserResult(userId, status, type, recordId, adminName, deliveryPersonId) {
  if (!userId) return;

  var isApproved = status.includes('อนุมัติ') && !status.includes('ไม่');
  var isRejected = status.includes('ไม่อนุมัติ') || status.includes('ยกเลิก') || status.includes('ปฏิเสธ');
  var isReturned = status.includes('คืนแล้ว') || status.includes('เสร็จสิ้น');
  var statusIcon = isApproved ? '✅' : isReturned ? '📦' : '❌';

  // ข้อมูลผู้จ่ายของ
  var dlvText = '';
  if (isApproved && deliveryPersonId) {
    var dlvList = getDeliveryConfigList();
    for (var d = 0; d < dlvList.length; d++) {
      if (dlvList[d].id === deliveryPersonId) {
        var dv = dlvList[d];
        dlvText = '━━━━━━━━━━━━━━━\n'
          + '📦 ข้อมูลการรับอุปกรณ์\n'
          + '👷 ผู้จ่ายของ: '   + (dv.name     || '-') + '\n'
          + '🏛️ หน่วยงาน: '    + (dv.org      || '-') + '\n'
          + '📞 โทร: '         + (dv.phone    || '-') + '\n'
          + '🕐 เวลานัดรับ: '  + (dv.time     || '-') + '\n'
          + '📍 สถานที่รับ: '  + (dv.location || '-') + '\n';
        break;
      }
    }
  }

  var approverText = adminName ? '👨‍💼 ผู้อนุมัติ: ' + adminName + '\n' : '';
  var msg = '';

  if (type === 'equip') {
    var data = getEquipRequestById(recordId);
    var itemLines = data ? String(data.items || '').split('\n').filter(function(l){
      return l.trim() && l.indexOf('[ยืม:') !== 0;
    }) : [];
    var itemsText = itemLines.length
      ? itemLines.map(function(l, i){ return '  ' + (i+1) + '. ' + l.replace(/^\[.*?\]\s*/,''); }).join('\n')
      : '-';

    var statusDetail = isApproved ? '🎉 คำขอได้รับการอนุมัติแล้ว!\nกรุณารับอุปกรณ์ตามข้อมูลด้านล่าง'
                     : isReturned ? '✅ บันทึกการคืนอุปกรณ์เรียบร้อยแล้ว ขอบคุณที่ใช้บริการ'
                     : '😔 คำขอไม่ผ่านการอนุมัติ\nหากมีข้อสงสัยกรุณาติดต่อทีมงาน';

    msg = statusIcon + ' อัปเดตสถานะ: ยืมอุปกรณ์\n'
      + '━━━━━━━━━━━━━━━\n'
      + '🆔 รหัสรายการ: '    + recordId                         + '\n'
      + (data ? '👤 ชื่อ: '            + data.name                  + '\n' : '')
      + (data ? '🎯 โครงการ: '         + (data.purpose || '-')      + '\n' : '')
      + (data ? '📅 วันที่ยืม: '      + data.borrowTime            + '\n' : '')
      + (data ? '📅 วันที่คืน: '      + data.returnTime            + '\n' : '')
      + (data ? '📦 รายการอุปกรณ์:\n' + itemsText                  + '\n' : '')
      + '━━━━━━━━━━━━━━━\n'
      + '📌 สถานะ: '         + status                              + '\n'
      + approverText
      + dlvText
      + statusDetail;

  } else if (type === 'booking') {
    var bdata = getBookingById(recordId);
    var statusDetail = isApproved ? '🎉 การจองได้รับการอนุมัติแล้ว!\nกรุณาเตรียมตัวใช้งานตามวันเวลาที่จอง'
                     : isReturned ? '✅ บันทึกการใช้งานเสร็จสิ้น ขอบคุณที่ใช้บริการ'
                     : '😔 การจองไม่ผ่านการอนุมัติ\nหากมีข้อสงสัยกรุณาติดต่อทีมงาน';

    msg = statusIcon + ' อัปเดตสถานะ: จองห้อง\n'
      + '━━━━━━━━━━━━━━━\n'
      + '🆔 รหัสรายการ: '   + recordId                                      + '\n'
      + (bdata ? '👤 ชื่อ: '          + bdata.name                           + '\n' : '')
      + (bdata ? '🚪 ห้อง: '          + bdata.room                           + '\n' : '')
      + (bdata ? '📅 วันที่: '         + bdata.date                           + '\n' : '')
      + (bdata ? '🕐 เวลา: '          + bdata.startTime + ' – ' + bdata.endTime + ' น.\n' : '')
      + (bdata ? '🎯 กิจกรรม: '       + (bdata.purpose || '-')               + '\n' : '')
      + (bdata ? '👥 จำนวน: '         + bdata.userCount + ' คน\n'            : '')
      + '━━━━━━━━━━━━━━━\n'
      + '📌 สถานะ: '         + status                                        + '\n'
      + approverText
      + statusDetail;
  }

  if (msg) sendLineMessage(userId, msg);
}

// helper: หาข้อมูล Equipment_Requests จาก recordId

// helper: หาข้อมูล Bookings จาก recordId
function getBookingById(recordId) {
  var sheet = getDB().getSheetByName('Bookings');
  if (!sheet) return null;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(recordId)) {
      return {
        name:      String(rows[i][2]  || ''),
        room:      String(rows[i][8]  || ''),
        date:      String(rows[i][9]  || ''),
        startTime: String(rows[i][10] || ''),
        endTime:   String(rows[i][11] || ''),
        userCount: String(rows[i][12] || ''),
        purpose:   String(rows[i][13] || ''),
        status:    String(rows[i][14] || '')
      };
    }
  }
  return null;
}

/** แจ้งผู้ยืมเมื่อส่งคำขอสำเร็จ (ค้นหา User ID จาก sid) */
function notifyUserEquipSubmit(sid, recordId, name, payload) {
  var userId = getLineUserIdBySid(sid);
  if (!userId) return;

  // แยก items ตัด org prefix
  var itemLines = payload ? String(payload.items || '').split('\n').filter(function(l){
    return l.trim() && l.indexOf('[ยืม:') !== 0;
  }) : [];
  var itemsText = itemLines.length
    ? itemLines.map(function(l, i){ return '  ' + (i+1) + '. ' + l.replace(/^\[.*?\]\s*/,''); }).join('\n')
    : '-';

  var msg = '📋 ได้รับคำขอยืมอุปกรณ์แล้ว!\n'
    + '━━━━━━━━━━━━━━━\n'
    + '🆔 รหัสรายการ: '   + recordId                        + '\n'
    + '👤 ชื่อ: '          + name                             + '\n'
    + '🎓 รหัสนักศึกษา: '  + (payload ? payload.sid     : '') + '\n'
    + '🏛️ คณะ/สังกัด: '   + (payload ? payload.faculty : '') + '\n'
    + '🎯 โครงการ/กิจกรรม: '+ (payload ? payload.purpose : '-') + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '📅 วันที่ยืม: '   + (payload ? payload.borrowTime : '') + '\n'
    + '📅 วันที่คืน: '   + (payload ? payload.returnTime : '') + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '📦 รายการอุปกรณ์:\n' + itemsText + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '📌 สถานะ: รออนุมัติ\n'
    + '⏳ ทีมงานจะตรวจสอบและแจ้งผลให้ทราบภายใน 1-2 วันทำการ';
  sendLineMessage(userId, msg);
}

/** แจ้งผู้ยืมเมื่อสถานะเปลี่ยน โดย lookup จาก sid */
function notifyUserBySid(sid, status, type, recordId, adminName, deliveryPersonId) {
  var userId = getLineUserIdBySid(sid);
  if (!userId) return;
  notifyUserResult(userId, status, type, recordId, adminName, deliveryPersonId);
}

// ==========================================
// 18. MODULE: LINE_Users — จับคู่รหัสนักศึกษา ↔ LINE User ID
// ==========================================

// ==========================================
// Helper: ตรวจและ migrate LINE_Users Sheet
// ==========================================

var LINE_USERS_HEADERS = ['รหัสนักศึกษา/บุคลากร', 'LINE User ID', 'ชื่อใน LINE', 'วันที่ลงทะเบียน', 'ชื่อ-นามสกุล', 'เบอร์โทร', 'ประเภท', 'อีเมล'];

/**
 * ตรวจว่า row แรกเป็น header หรือข้อมูล
 * header = คอลัม A มีคำว่า "รหัส" หรือ "LINE"
 */
/**
 * ตรวจว่าแถวแรกของ LINE_Users Sheet เป็น header จริง
 * ใช้วิธีตรวจว่าคอลัม B (LINE User ID) แถว 1 ขึ้นต้นด้วย "U"
 * เพราะ LINE User ID จริงขึ้นต้นด้วย U เสมอ และ header text ขึ้นต้นด้วยคำไทย
 */
function lineUsersHasHeader(sheet) {
  if (sheet.getLastRow() === 0) return false;
  var colA = String(sheet.getRange(1, 1).getValue() || '').trim();
  var colB = String(sheet.getRange(1, 2).getValue() || '').trim();
  // Header: คอลัม A คือ "รหัสนักศึกษา/บุคลากร", คอลัม B คือ "LINE User ID"
  // ข้อมูลจริง: คอลัม B ขึ้นต้นด้วย "U" (LINE User ID format: Uxxxxxxxx)
  if (colB.indexOf('U') === 0 && colB.length > 10) return false; // แถวแรกเป็นข้อมูลจริง
  if (colA === LINE_USERS_HEADERS[0]) return true;                 // แถวแรกเป็น header
  if (colA === '' && colB === '') return false;                    // Sheet ว่าง
  // fallback: ถ้าคอลัม A เป็นตัวเลขล้วน = SID = ข้อมูล ไม่ใช่ header
  return !/^[0-9]+$/.test(colA);
}

/**
 * เพิ่ม header ให้ Sheet เก่าที่ไม่มี — ตรวจก่อนว่ายังไม่มีจริง
 */
function migrateLineUsersSheet(sheet) {
  // ตรวจซ้ำก่อนแทรก — ป้องกัน double-migrate
  var colA = String(sheet.getRange(1, 1).getValue() || '').trim();
  if (colA === LINE_USERS_HEADERS[0]) {
    Logger.log('LINE_Users: มี header แล้ว ข้ามการ migrate');
    return;
  }
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, LINE_USERS_HEADERS.length).setValues([LINE_USERS_HEADERS]);
  sheet.getRange(1, 1, 1, LINE_USERS_HEADERS.length)
    .setBackground('#1a3a6b').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setColumnWidth(1, 180); sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 180); sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 180); sheet.setColumnWidth(6, 140); sheet.setColumnWidth(7, 120);
  Logger.log('LINE_Users: migrate header สำเร็จ');
}

function getOrCreateLineUsersSheet() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('LINE_Users');
  if (!sheet) {
    sheet = ss.insertSheet('LINE_Users');
    sheet.appendRow(LINE_USERS_HEADERS);
    sheet.getRange(1, 1, 1, LINE_USERS_HEADERS.length)
      .setBackground('#1a3a6b').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(1, 180); sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 180); sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 180); sheet.setColumnWidth(6, 140); sheet.setColumnWidth(7, 120);
    Logger.log('สร้าง LINE_Users sheet ใหม่');
    return sheet;
  }
  // Sheet มีอยู่แล้ว — ตรวจ header ด้วย logic ที่แม่นยำ
  if (!lineUsersHasHeader(sheet)) {
    migrateLineUsersSheet(sheet);
  }
  return sheet;
}

/** ดึงข้อมูลทั้งหมด (แถว 0 = header, 1+ = ข้อมูล) */
function getLineUsersData() {
  var sheet = getOrCreateLineUsersSheet();
  var data  = sheet.getDataRange().getValues();
  // safety: ถ้าแถวแรกเป็น header ซ้ำ (migrate bug) ให้ข้ามแถวซ้ำ
  var startRow = 1;
  if (data.length > 1 && String(data[1][0]).trim() === LINE_USERS_HEADERS[0]) {
    Logger.log('LINE_Users: พบ double header — ข้ามแถว header ที่ 2');
    startRow = 2;
    // คืน array ที่ตัด double header ออก
    return [data[0]].concat(data.slice(startRow));
  }
  return data;
}

/** ฟังก์ชัน debug — รันจาก GAS Editor เพื่อดูสถานะ Sheet */
function debugLineUsersSheet() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('LINE_Users');
  if (!sheet) { Logger.log('ไม่พบ LINE_Users sheet!'); return; }
  var data  = sheet.getDataRange().getValues();
  Logger.log('LINE_Users sheet:');
  Logger.log('  lastRow=' + sheet.getLastRow() + ', lastCol=' + sheet.getLastColumn());
  Logger.log('  hasHeader=' + lineUsersHasHeader(sheet));
  for (var i = 0; i < Math.min(data.length, 5); i++) {
    Logger.log('  row[' + i + ']: ' + JSON.stringify(data[i].slice(0,4)));
  }
}

/** ค้นหา LINE User ID จากรหัสนักศึกษา */
function getLineUserIdBySid(sid) {
  if (!sid) return null;
  var data   = getLineUsersData();
  var sidStr = String(sid).trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === sidStr) {
      return String(data[i][1]).trim() || null;
    }
  }
  return null;
}

/** ค้นหารหัสนักศึกษาจาก LINE User ID (reverse lookup) */
function getRegisteredSidByLineId(lineUserId) {
  if (!lineUserId) return null;
  var data = getLineUsersData();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === lineUserId) {
      return String(data[i][0]).trim() || null;
    }
  }
  return null;
}

/** ดึงโปรไฟล์เต็มจาก LINE User ID */
function getRegisteredProfileByLineId(lineUserId) {
  if (!lineUserId) return null;
  var data = getLineUsersData();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === lineUserId) {
      return {
        sid:      String(data[i][0] || '').trim(),
        lineUid:  String(data[i][1] || '').trim(),
        lineName: String(data[i][2] || '').trim(),
        regDate:  String(data[i][3] || '').trim(),
        name:     String(data[i][4] || '').trim(),
        phone:    String(data[i][5] || '').trim(),
        type:     String(data[i][6] || '').trim(),
        email:    String(data[i][7] || '').trim()
      };
    }
  }
  return null;
}

/** บันทึก/อัปเดตข้อมูลผู้ใช้ LINE */
function registerLineUser(sid, lineUserId, displayName, fullName, phone, userType, email) {
  var sheet  = getOrCreateLineUsersSheet();
  var data   = sheet.getDataRange().getValues();
  var sidStr = String(sid      || '-').trim();
  fullName   = String(fullName  || '').trim();
  phone      = String(phone     || '').trim();
  userType   = String(userType  || 'นักศึกษา').trim();
  email      = String(email     || '').trim();

  // ค้นหาจาก LINE UID ก่อน (รองรับกรณีเปลี่ยน SID หรือข้อมูลเก่า)
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === lineUserId) {
      sheet.getRange(i + 1, 1).setValue(sidStr);
      if (displayName) sheet.getRange(i + 1, 3).setValue(displayName);
      sheet.getRange(i + 1, 4).setValue(getThaiTimestamp());
      if (fullName) sheet.getRange(i + 1, 5).setValue(fullName);
      if (phone)    sheet.getRange(i + 1, 6).setValue(phone);
      if (userType) sheet.getRange(i + 1, 7).setValue(userType);
      if (email)    sheet.getRange(i + 1, 8).setValue(email);
      return 'updated';
    }
  }
  // ไม่พบ → เพิ่มแถวใหม่
  sheet.appendRow([sidStr, lineUserId, displayName || '', getThaiTimestamp(), fullName, phone, userType, email]);
  return 'created';
}

/** ดึงรายชื่อ LINE Users ทั้งหมด (Admin) */
function getLineUsersList() {
  try {
    var data = getLineUsersData();
    if (!data || data.length <= 1) return [];
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || (!row[0] && !row[1])) continue;
      // ข้ามถ้าแถวนี้คือ header อีกชั้น (ป้องกัน double-header bug)
      if (String(row[0]).trim() === LINE_USERS_HEADERS[0]) continue;
      list.push({
        sid:          String(row[0] || '').trim(),
        lineUserId:   String(row[1] || '').trim(),
        displayName:  String(row[2] || '').trim(),
        registeredAt: String(row[3] || '').trim(),
        name:         String(row[4] || '').trim(),
        phone:        String(row[5] || '').trim(),
        type:         String(row[6] || '').trim()
      });
    }
    return list;
  } catch(e) {
    Logger.log('getLineUsersList error: ' + e.message);
    return [];
  }
}

// ==========================================
// 19. LINE Webhook — doPost (รับข้อความจาก LINE)
// ==========================================
// แก้ปัญหา Timeout: doPost ตอบ 200 ทันที แล้วใช้ Trigger ประมวลผล
// ==========================================

/**
 * doPost: รับ Webhook จาก LINE -> บันทึก queue -> ตอบ 200 ทันที
 * ห้ามทำ sheet/HTTP ใน doPost เพราะ LINE timeout 3 วินาที
 */
function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
    var events   = contents.events || [];

    var toQueue = [];
    events.forEach(function(event) {
      if (event.type === 'message' && event.message.type === 'text') {
        toQueue.push({
          type:       'message',
          userId:     event.source.userId,
          replyToken: event.replyToken,
          text:       (event.message.text || '').trim()
        });
      } else if (event.type === 'follow') {
        // ผู้ใช้ add เพื่อนหรือ unblock
        toQueue.push({
          type:       'follow',
          userId:     event.source.userId,
          replyToken: event.replyToken,
          text:       ''
        });
      } else if (event.type === 'unfollow') {
        // ผู้ใช้ block — log ไว้
        Logger.log('LINE unfollow: ' + event.source.userId);
      }
    });

    if (toQueue.length > 0) {
      // รัน processLineQueue ตรงๆ เลย — ตอบกลับทันที ไม่ต้องรอ trigger
      toQueue.forEach(function(item) {
        try {
          var userId = item.userId;
          var text   = (item.text || '').trim();
          var textLower = text.toLowerCase();
          var isSidPattern = /^\d{5,12}$/.test(text) || /^[A-Za-z]\d{5,12}$/.test(text);

           if (item.type === 'follow') {
             var fp = getRegisteredProfileByLineId(userId);
             if (!fp || !fp.sid) {
               var oldSid = getRegisteredSidByLineId(userId);
               if (oldSid) fp = { sid: oldSid, name: '', phone: '' };
             }
             if (fp && fp.sid) {
               sendLineMessage(userId,
                 '👋 ยินดีต้อนรับกลับ' + (fp.name ? ', ' + fp.name + '!' : '!') + '\n'
                 + '━━━━━━━━━━━━━━━\n'
                 + '📋 ข้อมูลของคุณ:\n'
                 + '🎓 รหัส: '  + fp.sid + '\n'
                 + (fp.name  ? '👤 ชื่อ: '  + fp.name  + '\n' : '')
                 + (fp.phone ? '📞 เบอร์: ' + fp.phone + '\n' : '')
                 + '━━━━━━━━━━━━━━━\n'
                 + '💬 พิมพ์ "ติดตามสถานะ" เพื่อดูรายการ\n'
                 + '📝 พิมพ์ "แก้ไขข้อมูล" เพื่อเปลี่ยนข้อมูล'
               );
             } else {
               sendLineMessage(userId,
                 '👋 สวัสดีครับ! ยินดีต้อนรับสู่\n'
                 + '🎬 ชมรมสื่อสร้างสรรค์ Media EDU Club CMU.\n'
                 + '━━━━━━━━━━━━━━━\n'
                 + 'กรุณาลงทะเบียนเพื่อรับการแจ้งเตือนผ่าน LINE\n'
                 + 'ใช้เวลาเพียง 1 นาทีครับ!'
               );
               startReg(userId);
             }
             return;
           }




          var session = getUserSession(userId);

          // ── Registration ──
          if (session === 'reg:sid') {
            regGetSid(userId, text);
          } else if (session && session.indexOf('reg:name|') === 0) {
            regGetName(userId, text);
          } else if (session && session.indexOf('reg:phone|') === 0) {
            regGetPhone(userId, text);
          } else if (session && session.indexOf('reg:email|') === 0) {
            regGetEmail(userId, text);

          // ── ติดตามสถานะ ──

          } else if (text === 'ติดตามสถานะ' || text === 'ติดตาม' || text === 'tracking') {
            handleTrackingCommand(userId);

          // ── ลงทะเบียน / แก้ไขข้อมูล ──
          } else if (text === 'ลงทะเบียน' || text === 'register') {
            startReg(userId);
          } else if (text === 'แก้ไขข้อมูล' || text === 'แก้ไข' || text === 'edit') {
            startReg(userId);

          } else if (text === 'ยืมของ' || text === 'ยืมอุปกรณ์' || text === 'borrow') {
            var webUrl = (PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '') + getUserSidSuffix(userId);
            sendLineMessage(userId, '💻 ระบบยืมอุปกรณ์\n━━━━━━━━━━━━━━━\nสามารถยืมอุปกรณ์ได้ที่:\n' + webUrl + '\n\n📌 เลือกเมนู "ระบบยืมของ"\n⏰ เปิดให้บริการ จ–ศ 08:30–16:30 น.\n\nหลังส่งคำขอแล้วระบบจะแจ้งผลทาง LINE นี้ครับ');
          } else if (text === 'จองห้อง' || text === 'booking') {
            var webUrl2 = (PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '') + getUserSidSuffix(userId);
            sendLineMessage(userId, '🏫 ระบบจองห้อง\n━━━━━━━━━━━━━━━\nสามารถจองห้องได้ที่:\n' + webUrl2 + '\n\n📌 เลือกเมนู "ระบบจองห้อง"\n⏰ เปิดให้บริการ จ–ศ 08:30–16:30 น.\n\nหลังส่งคำขอแล้วระบบจะแจ้งผลทาง LINE นี้ครับ');
          } else if (text === 'ติดต่อ' || text === 'contact' || text === 'ติดต่อสอบถาม') {
            sendLineMessage(userId, getTpl('msg_contact'));
          } else if (text === 'สถานะ' || text === 'status' || text === 'ข้อมูลของฉัน') {
            var prof = getRegisteredProfileByLineId(userId);
            if (prof && (prof.name || prof.sid)) {
              sendLineMessage(userId,
                '📋 ข้อมูลของคุณ\n'
                + '━━━━━━━━━━━━━━━\n'
                + '👤 ชื่อ: '   + (prof.name  || '(ยังไม่ได้กรอก)') + '\n'
                + '🎓 รหัส: '   + (prof.sid   || '-') + '\n'
                + '📞 เบอร์: '  + (prof.phone || '(ยังไม่ได้กรอก)') + '\n'
                + '🏷️ ประเภท: ' + (prof.type  || '-') + '\n'
                + '━━━━━━━━━━━━━━━\n'
                + '📝 พิมพ์ "แก้ไขข้อมูล" เพื่ออัปเดต'
              );
            } else {
              startReg(userId);
            }
          } else if (text === 'ยกเลิก' || text === 'cancel') {
            clearUserSession(userId);
            sendLineMessage(userId, '↩️ ยกเลิกแล้วครับ');
          } else {
            var defProf = getRegisteredProfileByLineId(userId);
            var greeting = '👋 สวัสดีครับ' + (defProf && defProf.name ? ', ' + defProf.name + '!' : '!') + '\n\n';
            sendLineMessage(userId, greeting + getTpl('msg_default_menu')
              + (defProf ? '' : '\n✍️ ลงทะเบียน'));
          }
        } catch(eItem) {
          Logger.log('doPost item error: ' + eItem.message);
        }
      });
    }

  } catch(err) {
    Logger.log('doPost error: ' + err.message);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * processLineQueue: ทำงานหลัง doPost ผ่าน Trigger
 */
function processLineQueue() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processLineQueue') {
      ScriptApp.deleteTrigger(t);
    }
  });

  var props = PropertiesService.getScriptProperties();
  var raw   = props.getProperty('LINE_QUEUE');
  if (!raw) return;

  var queue = [];
  try { queue = JSON.parse(raw); } catch(e) { props.deleteProperty('LINE_QUEUE'); return; }
  props.deleteProperty('LINE_QUEUE');

  queue.forEach(function(item) {
    try {
      var userId = item.userId;
      var text   = (item.text || '').trim();

       // ---- Follow event ----
       if (item.type === 'follow') {
         var fp = getRegisteredProfileByLineId(userId);
         if (!fp || !fp.sid) {
           var oldSid = getRegisteredSidByLineId(userId);
           if (oldSid) fp = { sid: oldSid, name: '', phone: '' };
         }
         if (fp && fp.sid) {
           sendLineMessage(userId,
             '👋 ยินดีต้อนรับกลับ' + (fp.name ? ', ' + fp.name + '!' : '!') + '\n'
             + '━━━━━━━━━━━━━━━\n'
             + '📋 ข้อมูลของคุณ:\n'
             + '🎓 รหัส: '  + fp.sid + '\n'
             + (fp.name  ? '👤 ชื่อ: '  + fp.name  + '\n' : '')
             + (fp.phone ? '📞 เบอร์: ' + fp.phone + '\n' : '')
             + '━━━━━━━━━━━━━━━\n'
             + '💬 พิมพ์ "ติดตามสถานะ" เพื่อดูรายการ\n'
             + '📝 พิมพ์ "แก้ไขข้อมูล" เพื่อเปลี่ยนข้อมูล'
           );
         } else {
           sendLineMessage(userId,
             '👋 สวัสดีครับ! ยินดีต้อนรับสู่\n'
             + '🎬 ชมรมสื่อสร้างสรรค์ Media EDU Club CMU.\n'
             + '━━━━━━━━━━━━━━━\n'
             + 'กรุณาลงทะเบียนเพื่อรับการแจ้งเตือนผ่าน LINE\n'
             + 'ใช้เวลาเพียง 1 นาทีครับ!'
           );
           startReg(userId);
         }
         return;
       }




      // ---- Message event ----
      var textLower = text.toLowerCase();
      var isSidPattern = /^\d{5,12}$/.test(text) || /^[A-Za-z]\d{5,12}$/.test(text);

      var qSession = getUserSession(userId);

      // ── Registration ──
      if (qSession === 'reg:sid') {
        regGetSid(userId, text);
      } else if (qSession && qSession.indexOf('reg:name|') === 0) {
        regGetName(userId, text);
      } else if (qSession && qSession.indexOf('reg:phone|') === 0) {
        regGetPhone(userId, text);
      } else if (qSession && qSession.indexOf('reg:email|') === 0) {
        regGetEmail(userId, text);

      // ── ติดตามสถานะ ──

      } else if (text === 'ติดตามสถานะ' || text === 'ติดตาม' || text === 'tracking') {
        handleTrackingCommand(userId);

      // ── ลงทะเบียน / แก้ไข ──
      } else if (text === 'ลงทะเบียน' || text === 'register') {
        startReg(userId);
      } else if (text === 'แก้ไขข้อมูล' || text === 'แก้ไข' || text === 'edit') {
        startReg(userId);

      // ── ระบบยืมของ ──
      } else if (text === 'ยืมของ' || text === 'ยืมอุปกรณ์' || text === 'borrow') {
        var webUrl = (PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '') + getUserSidSuffix(userId);
        sendLineMessage(userId,
          '💻 ระบบยืมอุปกรณ์\n'
          + '━━━━━━━━━━━━━━━\n'
          + 'สามารถยืมอุปกรณ์ได้ที่:\n'
          + webUrl + '\n\n'
          + '📌 เลือกเมนู "ระบบยืมของ"\n'
          + '⏰ เปิดให้บริการ จ–ศ 08:30–16:30 น.\n\n'
          + 'หลังส่งคำขอแล้วระบบจะแจ้งผลทาง LINE นี้ครับ'
        );

      // ── ระบบจองห้อง ──
      } else if (text === 'จองห้อง' || text === 'booking') {
        var webUrl2 = (PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '') + getUserSidSuffix(userId);
        sendLineMessage(userId,
          '🏫 ระบบจองห้อง\n'
          + '━━━━━━━━━━━━━━━\n'
          + 'สามารถจองห้องได้ที่:\n'
          + webUrl2 + '\n\n'
          + '📌 เลือกเมนู "ระบบจองห้อง"\n'
          + '⏰ เปิดให้บริการ จ–ศ 08:30–16:30 น.\n\n'
          + 'หลังส่งคำขอแล้วระบบจะแจ้งผลทาง LINE นี้ครับ'
        );

      // ── ติดต่อสอบถาม ──
      } else if (text === 'ติดต่อ' || text === 'contact' || text === 'ติดต่อสอบถาม') {
        sendLineMessage(userId,
          '📞 ช่องทางติดต่อ\n'
          + '━━━━━━━━━━━━━━━\n'
          + '📘 Facebook: Media EDU Club CMU.\n'
          + 'https://www.facebook.com/media.ed.cmu\n\n'
          + '🕐 เวลาทำการ:\n'
          + 'จันทร์ – ศุกร์  08:30–16:30 น.\n\n'
          + '📍 ที่ตั้ง: คณะศึกษาศาสตร์ มหาวิทยาลัยเชียงใหม่\n\n'
          + '💬 หรือพิมพ์ข้อความมาได้เลยครับ ทีมงานจะตอบกลับโดยเร็ว'
        );

      // ── ลงทะเบียน (รหัสนักศึกษา) ──
      } else if (isSidPattern) {
        clearUserSession(userId);
        var result  = registerLineUser(text.toUpperCase(), userId, '');
        var pushMsg = result === 'updated'
          ? '✅ อัปเดตข้อมูลสำเร็จ!\n\n'
            + '📋 รหัสใหม่: ' + text.toUpperCase() + '\n'
            + 'ระบบจะส่งการแจ้งเตือนมาที่ LINE นี้ต่อไปครับ\n\n'
            + '📌 หากต้องการเปลี่ยนรหัส ส่งรหัสใหม่มาได้เลย'
          : '🎉 ลงทะเบียนสำเร็จ!\n\n'
            + '📋 รหัสของคุณ: ' + text.toUpperCase() + '\n\n'
            + 'ตอนนี้คุณจะได้รับการแจ้งเตือนผ่าน LINE เมื่อ:\n'
            + '📋 ส่งคำขอสำเร็จ\n'
            + '✅ อนุมัติ / ❌ ไม่อนุมัติ\n'
            + '📦 คืนอุปกรณ์แล้ว\n'
            + '⏰ ใกล้ถึงกำหนดคืน\n\n'
            + '📌 หากต้องการเปลี่ยนรหัส ส่งรหัสใหม่มาได้เลยครับ';
        sendLineMessage(userId, pushMsg);

      } else if (text === 'สถานะ' || text === 'status') {
        var sid = getRegisteredSidByLineId(userId);
        sendLineMessage(userId, sid
          ? '📋 ข้อมูลของคุณ\n━━━━━━━━━━━━━━━\nรหัส: ' + sid + '\n✅ ลงทะเบียนแล้ว รับการแจ้งเตือน\n\nพิมพ์รหัสใหม่ถ้าต้องการเปลี่ยน'
          : '❌ ยังไม่ได้ลงทะเบียน\nส่งรหัสนักศึกษา/บุคลากรมาเพื่อเริ่มรับการแจ้งเตือน');

      } else if (text === 'ยกเลิก' || text === 'cancel') {
        clearUserSession(userId);
        sendLineMessage(userId, '📱 ส่งรหัสนักศึกษา/บุคลากรของคุณเพื่อลงทะเบียนรับการแจ้งเตือน\nเช่น: 650210343');

      } else {
        sendLineMessage(userId,
          '👋 สวัสดีครับ!\n\n'
          + '📌 เลือกเมนูด้านล่าง หรือพิมพ์คำสั่งได้เลย:\n\n'
          + '🔍 ติดตามสถานะ — ดูสถานะรายการของคุณ\n'
          + '💻 ยืมอุปกรณ์ — ข้อมูลระบบยืมของ\n'
          + '🏫 จองห้อง — ข้อมูลระบบจองห้อง\n'
          + '📞 ติดต่อ — ช่องทางติดต่อทีมงาน\n\n'
          + '💡 ส่งรหัสนักศึกษาเพื่อลงทะเบียนรับแจ้งเตือน'
        );
      }
    } catch(err) {
      Logger.log('processLineQueue item error: ' + err.message);
    }
  });
}

/**
 * setupLineQueueTrigger: รันครั้งเดียวเพื่อเคลียร์ trigger เก่า
 */
function setupLineQueueTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processLineQueue') {
      ScriptApp.deleteTrigger(t);
    }
  });
  Logger.log('Trigger cleared');
}

// ==========================================
// 20. Contact Form — รับข้อความแจ้งปัญหา
// ==========================================
function submitContactMessage(payload) {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Contact_Messages');
  if (!sheet) {
    sheet = ss.insertSheet('Contact_Messages');
    sheet.appendRow(['Timestamp','ชื่อ','เบอร์โทร','อีเมล','ประเภท','ข้อความ','สถานะ']);
    sheet.getRange(1,1,1,7).setBackground('#1a3a6b').setFontColor('#fff').setFontWeight('bold');
  }
  var id = getThaiTimestamp();
  sheet.appendRow([id, payload.name, payload.phone, payload.email || '', payload.type || '', payload.message, 'ใหม่']);

  // แจ้ง Admin ทาง LINE
  var msg = '📨 มีข้อความใหม่จากผู้ใช้!\n'
    + '━━━━━━━━━━━━━━━\n'
    + '👤 ชื่อ: '     + payload.name    + '\n'
    + '📞 เบอร์: '    + payload.phone   + '\n'
    + '🏷 ประเภท: '  + (payload.type || '-') + '\n'
    + '💬 ข้อความ: ' + (payload.message || '').substring(0, 100)
    + (payload.message && payload.message.length > 100 ? '...' : '');
  notifyAdmins(msg);

  // ส่งอีเมลตอบรับกลับผู้ใช้ (ถ้ามีอีเมล)
  if (payload.email) {
    var subject = '[Media EDU Club] รับข้อความของคุณแล้ว — ' + (payload.type || 'ติดต่อทีมงาน');
    var body = buildEmailHtml({
      title:      'รับข้อความของคุณแล้ว!',
      headerIcon: 'S',
      color:      '#1565c0',
      name:       payload.name,
      intro:      'ทีมงาน Media EDU Club CMU. ได้รับข้อความของคุณแล้ว จะติดต่อกลับภายใน 1-2 วันทำการ',
      rows: [
        ['ชื่อ',         payload.name],
        ['เบอร์โทร',    payload.phone],
        ['ประเภท',      payload.type  || '-'],
        ['ข้อความ',     payload.message],
      ],
      note: 'หากมีคำถามเพิ่มเติมสามารถติดต่อผ่าน LINE OA: @005wrfsu ได้เลยครับ'
    });
    sendConfirmEmail(payload.email, subject, body);
  }

  return 'ok';
}


// ==========================================
// 22. Contact Messages — Admin Inbox
// ==========================================

function getContactMessages() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Contact_Messages');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    // [timestamp, name, phone, email, type, message, status]
    rows.push([
      String(data[i][0] || ''),
      String(data[i][1] || ''),
      String(data[i][2] || ''),
      String(data[i][3] || ''),
      String(data[i][4] || ''),
      String(data[i][5] || ''),
      String(data[i][6] || 'ใหม่')
    ]);
  }
  return rows;
}

function updateContactMessageStatus(listIndex, newStatus) {
  var ss    = getDB();
  var sheet = ss.getSheetByName('Contact_Messages');
  if (!sheet) return 'no_sheet';
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (count === listIndex) {
      sheet.getRange(i + 1, 7).setValue(newStatus);
      return 'ok';
    }
    count++;
  }
  return 'not_found';
}

function sendAdminReplyEmail(payload) {
  var toEmail   = payload.toEmail   || '';
  var toName    = payload.toName    || 'ผู้ใช้งาน';
  var replyBody = payload.replyBody || '';
  var msgType   = payload.msgType   || 'ติดต่อทีมงาน';
  var rowIdx    = payload.rowIdx;

  if (!toEmail) throw new Error('ไม่พบอีเมลผู้รับ');
  if (!replyBody) throw new Error('ข้อความว่างเปล่า');

  var subject = '[Media EDU Club CMU] ตอบกลับ: ' + msgType;

  var body = buildEmailHtml({
    title:      'ข้อความตอบกลับจากทีมงาน',
    headerIcon: 'M',
    color:      '#1565c0',
    name:       toName,
    intro:      replyBody,
    rows:       [],
    note:       'หากมีคำถามเพิ่มเติมสามารถติดต่อกลับได้เลยครับ/ค่ะ\nFacebook: Media EDU Club CMU | LINE OA: @005wrfsu'
  });

  sendConfirmEmail(toEmail, subject, body);

  // อัปเดตสถานะใน Sheet เป็น "ตอบแล้ว"
  if (typeof rowIdx !== 'undefined') {
    updateContactMessageStatus(rowIdx, 'ตอบแล้ว');
  }

  // แจ้ง Admin LINE
  notifyAdmins('✅ ส่งอีเมลตอบกลับ\n👤 ' + toName + '\n📧 ' + toEmail + '\n\n' + replyBody.substring(0, 100) + (replyBody.length > 100 ? '...' : ''));

  return 'ok';
}




function getOrCreateFaqSheet() {
  var ss    = getDB();
  var sheet = ss.getSheetByName('FAQ_Config');
  if (!sheet) {
    sheet = ss.insertSheet('FAQ_Config');
    sheet.appendRow(['หมวดหมู่', 'คำถาม', 'คำตอบ', 'ลำดับ']);
    sheet.getRange(1,1,1,4).setBackground('#1a3a6b').setFontColor('#fff').setFontWeight('bold');
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 280);
    sheet.setColumnWidth(3, 380);
    sheet.setColumnWidth(4, 80);
    // ใส่ข้อมูล default
    var defaults = [
      ['borrow','ยืมอุปกรณ์ได้สูงสุดกี่วัน?','ยืมได้สูงสุด <b>7 วัน</b> ต่อครั้ง หากต้องการยืมนานกว่านั้น กรุณาติดต่อทีมงานโดยตรงครับ',1],
      ['borrow','ต้องนำอะไรไปแสดงตอนรับอุปกรณ์?','กรุณานำ <b>บัตรนักศึกษา</b> หรือบัตรประชาชนไปแสดงต่อผู้จ่ายของ และแจ้งรหัสรายการที่ได้รับทางอีเมลครับ',2],
      ['borrow','ส่งคำขอแล้ว ต้องรอนานแค่ไหน?','ทีมงานจะตรวจสอบและอนุมัติภายใน <b>1-2 วันทำการ</b> คุณจะได้รับแจ้งผลทางอีเมลครับ',3],
      ['borrow','คืนอุปกรณ์ที่ไหน?','คืนได้ที่ห้องชมรม Media EDU คณะศึกษาศาสตร์ ในเวลาทำการ จ–ศ 08:30–16:30 น.',4],
      ['booking','จองห้องล่วงหน้าได้นานแค่ไหน?','สามารถจองล่วงหน้าได้ไม่เกิน <b>30 วัน</b> และต้องจองอย่างน้อย <b>1 วันทำการ</b> ก่อนวันใช้งานครับ',1],
      ['booking','ยกเลิกการจองห้องได้ไหม?','ได้ครับ กรุณาแจ้งยกเลิกก่อน <b>1 วันทำการ</b> ผ่านช่องทางการติดต่อของชมรม',2],
      ['order','ชำระเงินยังไง?','รอรับการติดต่อจากทีมงานหลังส่งคำสั่งซื้อ ทีมงานจะแจ้งช่องทางชำระเงินทางอีเมลครับ',1],
      ['order','รับสินค้าได้ที่ไหน?','รับได้ที่ห้องชมรม Media EDU คณะศึกษาศาสตร์ ในเวลาทำการ หรือตามที่ทีมงานนัดหมายครับ',2],
      ['general','ใครสามารถใช้บริการได้บ้าง?','บริการเปิดสำหรับ <b>นักศึกษาและบุคลากร</b> คณะศึกษาศาสตร์ มหาวิทยาลัยเชียงใหม่ครับ',1],
      ['general','ลืมรหัสรายการ ทำยังไง?','กรอกเบอร์โทรศัพท์หรือรหัสนักศึกษาในหน้า "ติดตามสถานะ" ระบบจะแสดงรายการทั้งหมดของคุณครับ',2],
    ];
    defaults.forEach(function(row) { sheet.appendRow(row); });
    Logger.log('สร้าง FAQ_Config sheet พร้อม default data แล้ว');
  }
  return sheet;
}

/** ดึง FAQ ทั้งหมด (สำหรับผู้ใช้และ Admin) */
function getFaqList() {
  return withCache('faq', CACHE_TTL.faq, getFaqList_raw_);
}

function getFaqList_raw_() {
  var sheet = getOrCreateFaqSheet();
  var data  = sheet.getDataRange().getValues();
  var list  = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][1]) continue; // ข้ามแถวว่าง
    list.push({
      cat:   String(data[i][0] || 'general').trim(),
      q:     String(data[i][1] || '').trim(),
      a:     String(data[i][2] || '').trim(),
      order: parseInt(data[i][3]) || 0
    });
  }
  // เรียงตาม order
  list.sort(function(a, b) { return a.order - b.order; });
  return list;
}

/** เพิ่ม FAQ ใหม่ */
function addFaqItem(item) {
  var sheet = getOrCreateFaqSheet();
  sheet.appendRow([
    item.cat   || 'general',
    item.q     || '',
    item.a     || '',
    item.order || 0
  ]);
  return 'ok';
}

/** แก้ไข FAQ (ตาม index ของ list ที่ไม่นับ header) */
function updateFaqItem(idx, item) {
  var sheet = getOrCreateFaqSheet();
  var data  = sheet.getDataRange().getValues();
  // idx คือ index ใน list (ไม่นับ header row=0 และแถวว่าง)
  var count = -1;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    count++;
    if (count === idx) {
      sheet.getRange(i+1, 1).setValue(item.cat   || 'general');
      sheet.getRange(i+1, 2).setValue(item.q     || '');
      sheet.getRange(i+1, 3).setValue(item.a     || '');
      sheet.getRange(i+1, 4).setValue(item.order || 0);
      return 'ok';
    }
  }
  return 'not_found';
}

/** ลบ FAQ (ตาม index ของ list) */
function deleteFaqItem(idx) {
  var sheet = getOrCreateFaqSheet();
  var data  = sheet.getDataRange().getValues();
  var count = -1;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    count++;
    if (count === idx) {
      sheet.deleteRow(i+1);
      return 'ok';
    }
  }
  return 'not_found';
}

// ==========================================
// TEST FUNCTION — รันใน Apps Script Editor
// ==========================================
function testLine() {
  notifyAdmins('✅ ทดสอบระบบแจ้งเตือน LINE สำเร็จ!\nMedia EDU Club CMU System');
}

function testLineUser() {
  // ทดสอบแจ้งเตือนผู้ยืม — ใส่รหัสนักศึกษาจริงที่ลงทะเบียนแล้ว
  var testSid = '650000001'; // ← เปลี่ยนเป็นรหัสที่ทดสอบ
  var userId  = getLineUserIdBySid(testSid);
  if (userId) {
    notifyUserResult(userId, 'อนุมัติแล้ว', 'equip', 'edurent001');
    Logger.log('✅ ส่งสำเร็จ → ' + userId);
  } else {
    Logger.log('❌ ไม่พบ LINE User ID สำหรับรหัส ' + testSid + ' กรุณาส่งรหัสหา Bot ก่อน');
  }
}

// ==========================================
// FACULTY API INTEGRATION — ตรวจสอบห้องว่าง
// ==========================================
// วิธีตั้งค่า: ไปที่ Apps Script → Project Settings → Script Properties
// เพิ่ม property:
//   FACULTY_API_URL  = https://api.example.cmu.ac.th/v1
//   FACULTY_API_KEY  = your_api_key_here

var FACULTY_API_ENABLED = false; // ← เปลี่ยนเป็น true เมื่อได้ API จริง

/**
 * ดึงการตั้งค่า API จาก Script Properties
 */
function getFacultyApiConfig() {
  var props = PropertiesService.getScriptProperties().getProperties();
  return {
    baseUrl: props['FACULTY_API_URL'] || '',
    apiKey:  props['FACULTY_API_KEY'] || '',
    enabled: props['FACULTY_API_ENABLED'] === 'true' || FACULTY_API_ENABLED
  };
}

/**
 * ดึงรายการการจองห้องจากระบบคณะตามวันที่
 * @param {string} date  - วันที่ในรูปแบบ YYYY-MM-DD
 * @param {string} roomId - รหัสห้อง (optional)
 * @returns {Array} รายการการจองที่มีสถานะ approved/pending
 */
function getFacultyBookingsByDate(date, roomId) {
  var cfg = getFacultyApiConfig();

  // ถ้ายังไม่ได้เปิดใช้ หรือไม่มี URL → คืน array ว่าง
  if (!cfg.enabled || !cfg.baseUrl) {
    Logger.log('[FacultyAPI] ปิดอยู่หรือยังไม่ได้ตั้งค่า FACULTY_API_URL');
    return [];
  }

  try {
    var url = cfg.baseUrl + '/room-bookings?date=' + date;
    if (roomId) url += '&room_id=' + encodeURIComponent(roomId);

    var options = {
      method:            'GET',
      headers:           { 'X-API-Key': cfg.apiKey, 'Accept': 'application/json' },
      muteHttpExceptions: true,
      followRedirects:   true
    };

    var response = UrlFetchApp.fetch(url, options);
    var code     = response.getResponseCode();

    if (code === 200) {
      var json = JSON.parse(response.getContentText());
      if (json.success && Array.isArray(json.bookings)) {
        // กรองเฉพาะ approved และ pending (ไม่รวม cancelled)
        return json.bookings.filter(function(b) {
          return b.status !== 'cancelled';
        });
      }
      Logger.log('[FacultyAPI] Response ไม่ตรง format: ' + response.getContentText().substring(0, 200));
      return [];
    } else if (code === 401) {
      Logger.log('[FacultyAPI] API Key ไม่ถูกต้อง (401)');
      return [];
    } else {
      Logger.log('[FacultyAPI] Error HTTP ' + code + ': ' + response.getContentText().substring(0, 200));
      return [];
    }
  } catch (e) {
    Logger.log('[FacultyAPI] Exception: ' + e.message);
    return [];
  }
}

/**
 * ตรวจสอบว่าห้องว่างหรือไม่ในช่วงเวลาที่กำหนด
 * รวมการเช็คทั้งจาก Bookings Sheet ของเราและจาก Faculty API
 *
 * @param {string} roomName  - ชื่อห้อง (ตรง col ใน Bookings sheet)
 * @param {string} date      - วันที่ YYYY-MM-DD
 * @param {string} startTime - เวลาเริ่ม HH:MM
 * @param {string} endTime   - เวลาสิ้นสุด HH:MM
 * @returns {Object} { available: bool, conflicts: Array }
 */
function checkRoomAvailability(roomName, date, startTime, endTime) {
  var conflicts = [];

  // --- ส่วนที่ 1: เช็ค Bookings Sheet ของระบบเรา ---
  var localBookings = getBookingsByDate(date);
  localBookings.forEach(function(b) {
    if (!b.room) return;
    // เช็คว่าเป็นห้องเดียวกัน (case-insensitive, trim)
    if (b.room.trim().toLowerCase() !== roomName.trim().toLowerCase()) return;
    // เช็ค time overlap
    if (timesOverlap(startTime, endTime, b.startTime, b.endTime)) {
      conflicts.push({
        source:    'ระบบชมรม',
        bookingId: b.id,
        room:      b.room,
        date:      b.date,
        startTime: b.startTime,
        endTime:   b.endTime,
        status:    b.status
      });
    }
  });

  // --- ส่วนที่ 2: เช็ค Faculty API ---
  var cfg = getFacultyApiConfig();
  if (cfg.enabled && cfg.baseUrl) {
    // หา room_id ที่ map กับชื่อห้อง
    var facultyRoomId = getFacultyRoomId(roomName);
    var facultyBookings = getFacultyBookingsByDate(date, facultyRoomId || undefined);

    facultyBookings.forEach(function(b) {
      // ถ้าไม่มี room_id mapping ให้เช็คทุกห้อง
      if (facultyRoomId && b.room_id && b.room_id !== facultyRoomId) return;
      if (timesOverlap(startTime, endTime, b.start_time, b.end_time)) {
        conflicts.push({
          source:    'ระบบคณะ',
          bookingId: b.booking_id,
          room:      b.room_name || b.room_id,
          date:      b.date,
          startTime: b.start_time,
          endTime:   b.end_time,
          status:    b.status
        });
      }
    });
  }

  return {
    available: conflicts.length === 0,
    conflicts: conflicts
  };
}

/**
 * ดึง room_id ของคณะจากตาราง mapping ใน Script Properties
 * format: FACULTY_ROOM_MAP = {"Studio Room":"STUDIO_01","ห้องประชุม":"CONF_01"}
 */
function getFacultyRoomId(roomName) {
  try {
    var props   = PropertiesService.getScriptProperties().getProperties();
    var mapJson = props['FACULTY_ROOM_MAP'];
    if (!mapJson) return null;
    var map = JSON.parse(mapJson);
    return map[roomName] || null;
  } catch(e) {
    return null;
  }
}

/**
 * Helper: เช็ค time overlap
 * คืน true ถ้าช่วงเวลา [s1,e1] ทับกับ [s2,e2]
 */
function timesOverlap(s1, e1, s2, e2) {
  if (!s1 || !e1 || !s2 || !e2) return false;
  var start1 = timeToMinutes(s1), end1 = timeToMinutes(e1);
  var start2 = timeToMinutes(s2), end2 = timeToMinutes(e2);
  return start1 < end2 && end1 > start2;
}

function timeToMinutes(t) {
  if (!t) return 0;
  var parts = String(t).replace(' น.', '').replace('น.', '').trim().split(':');
  return parseInt(parts[0] || 0) * 60 + parseInt(parts[1] || 0);
}

/**
 * เรียกจาก Frontend: เช็คห้องว่างก่อนแสดงฟอร์มจอง
 * @param {string} date - YYYY-MM-DD
 * @returns {Object} { localBookings, facultyBookings, roomAvailability }
 */
function getBookingAvailabilityForDate(date) {
  var cfg  = getFacultyApiConfig();
  var result = {
    date:             date,
    localBookings:    [],
    facultyBookings:  [],
    facultyEnabled:   cfg.enabled && !!cfg.baseUrl,
    facultyError:     null
  };

  // โหลด local bookings
  result.localBookings = getBookingsByDate(date);

  // โหลด faculty bookings
  if (result.facultyEnabled) {
    try {
      result.facultyBookings = getFacultyBookingsByDate(date);
    } catch(e) {
      result.facultyError = e.message;
    }
  }

  return result;
}

/**
 * Admin: บันทึกการตั้งค่า Faculty API
 * (เรียกจากหน้า Admin → ตั้งค่า Integration)
 */
function saveFacultyApiSettings(apiUrl, apiKey, roomMapJson, enabled) {
  var props = PropertiesService.getScriptProperties();
  if (apiUrl)      props.setProperty('FACULTY_API_URL',      apiUrl.trim());
  if (apiKey)      props.setProperty('FACULTY_API_KEY',       apiKey.trim());
  if (roomMapJson) props.setProperty('FACULTY_ROOM_MAP',      roomMapJson.trim());
  props.setProperty('FACULTY_API_ENABLED', enabled ? 'true' : 'false');
  return true;
}

/**
 * Admin: ดึงการตั้งค่าปัจจุบัน (ซ่อน API key ไว้บางส่วน)
 */
function getFacultyApiSettings() {
  var cfg = getFacultyApiConfig();
  var props = PropertiesService.getScriptProperties().getProperties();
  var key = cfg.apiKey;
  return {
    apiUrl:     cfg.baseUrl,
    apiKeyMask: key ? key.substring(0,4) + '****' + key.substring(key.length - 4) : '',
    hasKey:     !!key,
    roomMap:    props['FACULTY_ROOM_MAP'] || '{}',
    enabled:    cfg.enabled
  };
}

/**
 * Admin: ทดสอบการเชื่อมต่อ Faculty API
 */
function testFacultyApiConnection() {
  var cfg = getFacultyApiConfig();
  if (!cfg.baseUrl) return { success: false, message: 'ยังไม่ได้ตั้งค่า FACULTY_API_URL' };
  if (!cfg.apiKey)  return { success: false, message: 'ยังไม่ได้ตั้งค่า FACULTY_API_KEY' };

  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  try {
    var url  = cfg.baseUrl + '/room-bookings?date=' + today;
    var resp = UrlFetchApp.fetch(url, {
      method:            'GET',
      headers:           { 'X-API-Key': cfg.apiKey },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText().substring(0, 300);
    if (code === 200) {
      return { success: true, message: 'เชื่อมต่อสำเร็จ ✅ (HTTP 200)', preview: body };
    } else {
      return { success: false, message: 'HTTP ' + code, preview: body };
    }
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ==========================================
// QR CODE — ระบบยืนยันรับ/คืนอุปกรณ์
// ==========================================

/**
 * ดึงข้อมูลคำขอยืมจากรหัส (สำหรับ QR scan)
 */
function getEquipRequestById(requestId) {
  var sheet = getDB().getSheetByName('Equipment_Requests');
  if (!sheet) return null;
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(requestId).trim()) {
      var status = String(data[i][11] || '').trim();
      return {
        row:        i + 1,
        id:         data[i][0],
        timestamp:  data[i][1],
        name:       data[i][2],
        phone:      data[i][3],
        sid:        data[i][4],
        faculty:    data[i][5],
        startDate:  data[i][6],
        returnDate: data[i][7],
        items:      data[i][8],
        purpose:    data[i][9] || '',
        clubName:   data[i][10] || '',
        status:     status,
        approver:   data[i][12] || '',
        email:      data[i][13] || '',
        // action ที่ทำได้ตาม status
        canPickup:  status === 'อนุมัติแล้ว',
        canReturn:  status === 'อนุมัติ/กำลังยืม'
      };
    }
  }
  return null;
}

/**
 * ยืนยันการรับอุปกรณ์ (อนุมัติแล้ว → อนุมัติ/กำลังยืม)
 */
function confirmEquipPickup(requestId, adminName) {
  var req = getEquipRequestById(requestId);
  if (!req) return { success: false, message: 'ไม่พบรหัสคำขอ: ' + requestId };
  if (!req.canPickup) {
    return { success: false, message: 'สถานะปัจจุบัน "' + req.status + '" ไม่สามารถยืนยันการรับได้' };
  }
  updateEquipRequestStatus(req.row, 'อนุมัติ/กำลังยืม', adminName || 'QR Scan', undefined, null);
  return {
    success: true,
    action:  'pickup',
    message: 'ยืนยันการรับอุปกรณ์เรียบร้อย ✅',
    name:    req.name,
    items:   req.items,
    newStatus: 'อนุมัติ/กำลังยืม'
  };
}

/**
 * ยืนยันการคืนอุปกรณ์ (อนุมัติ/กำลังยืม → คืนแล้ว)
 */
function confirmEquipReturn(requestId, adminName) {
  var req = getEquipRequestById(requestId);
  if (!req) return { success: false, message: 'ไม่พบรหัสคำขอ: ' + requestId };
  if (!req.canReturn) {
    return { success: false, message: 'สถานะปัจจุบัน "' + req.status + '" ไม่สามารถยืนยันการคืนได้' };
  }
  updateEquipRequestStatus(req.row, 'คืนแล้ว', adminName || 'QR Scan', undefined, null);
  return {
    success: true,
    action:  'return',
    message: 'บันทึกการคืนอุปกรณ์เรียบร้อย ✅',
    name:    req.name,
    items:   req.items,
    newStatus: 'คืนแล้ว'
  };
}

// ==========================================
// AUDIT LOG — บันทึกประวัติการแก้ไขสถานะ
// ==========================================

/**
 * บันทึก log การเปลี่ยนสถานะลง Sheet 'AuditLog'
 * @param {string} type       - 'order' | 'equipment' | 'booking'
 * @param {string} recordId   - รหัสรายการ
 * @param {string} targetName - ชื่อผู้ใช้ (ผู้ยืม/ผู้จอง/ผู้สั่ง)
 * @param {string} oldStatus  - สถานะก่อนเปลี่ยน
 * @param {string} newStatus  - สถานะใหม่
 * @param {string} adminName  - Admin ที่แก้ไข
 * @param {string} note       - หมายเหตุเพิ่มเติม (เช่น เหตุผลปฏิเสธ)
 */
function logAuditEvent(type, recordId, targetName, oldStatus, newStatus, adminName, note) {
  try {
    var ss    = getDB();
    var sheet = ss.getSheetByName('AuditLog');
    if (!sheet) {
      sheet = ss.insertSheet('AuditLog');
      sheet.appendRow(['เวลา', 'ประเภท', 'รหัสรายการ', 'ชื่อ', 'สถานะเดิม', 'สถานะใหม่', 'Admin', 'หมายเหตุ']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 8).setBackground('#1a3a6b').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 100);
      sheet.setColumnWidth(3, 160);
      sheet.setColumnWidth(4, 160);
      sheet.setColumnWidth(5, 160);
      sheet.setColumnWidth(6, 160);
      sheet.setColumnWidth(7, 120);
      sheet.setColumnWidth(8, 200);
    }
    var timestamp = getThaiTimestamp();
    var typeLabel = type === 'order' ? 'ออเดอร์' : type === 'equipment' ? 'ยืมอุปกรณ์' : 'จองห้อง';
    sheet.appendRow([timestamp, typeLabel, recordId, targetName || '', oldStatus || '', newStatus || '', adminName || '', note || '']);
  } catch(e) {
    Logger.log('logAuditEvent error: ' + e.message);
  }
}

/**
 * ดึง Audit Log สำหรับ Admin
 * @param {Object} filters - { type, adminName, dateFrom, dateTo, recordId, keyword }
 */
function getAuditLog(filters) {
  var sheet = getDB().getSheetByName('AuditLog');
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  filters = filters || {};
  var rows = [];
  for (var i = data.length - 1; i >= 1; i--) { // ล่าสุดขึ้นก่อน
    var row = data[i];
    var timestamp  = row[0] || '';
    var type       = row[1] || '';
    var recordId   = row[2] || '';
    var name       = row[3] || '';
    var oldStatus  = row[4] || '';
    var newStatus  = row[5] || '';
    var admin      = row[6] || '';
    var note       = row[7] || '';

    // Filter
    if (filters.type      && type     !== filters.type)                     continue;
    if (filters.adminName && admin.toLowerCase().indexOf(filters.adminName.toLowerCase()) === -1) continue;
    if (filters.recordId  && recordId.toLowerCase().indexOf(filters.recordId.toLowerCase()) === -1) continue;
    if (filters.keyword) {
      var kw = filters.keyword.toLowerCase();
      if ([recordId, name, oldStatus, newStatus, admin, note].join(' ').toLowerCase().indexOf(kw) === -1) continue;
    }

    rows.push({ timestamp, type, recordId, name, oldStatus, newStatus, admin, note });
    if (rows.length >= 200) break; // จำกัด 200 รายการ
  }
  return rows;
}

/**
 * ดึง Audit Log ของรายการเดียว (สำหรับดูประวัติของ record นั้น)
 */
function getAuditLogByRecord(recordId) {
  return getAuditLog({ recordId: recordId });
}

// ==========================================
// OVERDUE NOTIFICATIONS — แจ้งเตือนเกินกำหนดคืน
// ==========================================

/**
 * ดึงการตั้งค่า overdue จาก Script Properties
 */
function getOverdueConfig() {
  var props = PropertiesService.getScriptProperties().getProperties();
  return {
    enabled:       props['OVERDUE_ENABLED']      === 'true',
    graceDays:     parseInt(props['OVERDUE_GRACE_DAYS']  || '0'),  // วันผ่อนผันหลังครบกำหนด
    reminderDays:  props['OVERDUE_REMINDER_DAYS'] ? props['OVERDUE_REMINDER_DAYS'].split(',').map(Number) : [0, 1, 3], // แจ้งก่อน N วัน (0=วันครบกำหนด, ลบ=เกินกำหนด)
    sendLine:      props['OVERDUE_SEND_LINE']     !== 'false',
    sendEmail:     props['OVERDUE_SEND_EMAIL']    !== 'false',
  };
}

function saveOverdueConfig(enabled, graceDays, reminderDays, sendLine, sendEmail) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('OVERDUE_ENABLED',       enabled ? 'true' : 'false');
  props.setProperty('OVERDUE_GRACE_DAYS',    String(parseInt(graceDays) || 0));
  props.setProperty('OVERDUE_REMINDER_DAYS', (reminderDays || '0,1,3'));
  props.setProperty('OVERDUE_SEND_LINE',     sendLine  ? 'true' : 'false');
  props.setProperty('OVERDUE_SEND_EMAIL',    sendEmail ? 'true' : 'false');
  return true;
}

/**
 * ค้นหารายการที่ใกล้ครบกำหนด / เกินกำหนด
 * @returns {Array} รายการพร้อม field: daysOverdue (+ = เกิน, - = ยังไม่ถึง, 0 = วันนี้)
 */
function findOverdueEquipment() {
  var sheet = getDB().getSheetByName('Equipment_Requests');
  if (!sheet) return [];

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var data = sheet.getDataRange().getDisplayValues();
  var results = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var status = String(row[11] || '').trim();
    // เฉพาะที่กำลังยืมอยู่ (อนุมัติแล้ว หรือ กำลังยืม)
    var isActive = status.includes('อนุมัติ') && !status.includes('ไม่')
      || status.includes('กำลังยืม');
    if (!isActive) continue;

    // parse returnDate (col 7) รองรับหลาย format
    var returnDateStr = String(row[7] || '').trim();
    if (!returnDateStr) continue;

    var returnDate = parseThaiDate(returnDateStr);
    if (!returnDate) continue;
    returnDate.setHours(0, 0, 0, 0);

    var diffMs   = today - returnDate;
    var daysOverdue = Math.round(diffMs / (1000 * 60 * 60 * 24)); // + = เกิน, - = ยังไม่ถึง

    results.push({
      row:           i + 1,
      id:            row[0],
      name:          row[2],
      sid:           row[3],
      phone:         row[4],
      email:         row[13] || '',
      items:         row[8],
      borrowDate:    row[6],
      returnDate:    returnDateStr,
      status:        status,
      daysOverdue:   daysOverdue,   // จำนวนวันที่เกิน (บวก=เกิน, ลบ=ยังไม่ถึง)
      returnDateObj: returnDate
    });
  }

  // เรียงจากเกินมากสุด
  results.sort(function(a, b) { return b.daysOverdue - a.daysOverdue; });
  return results;
}

/**
 * parse วันที่จากหลาย format (YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY, Thai date string)
 */
function parseThaiDate(str) {
  if (!str) return null;
  str = str.trim().split(' ')[0]; // เอาแค่ส่วนวันที่
  var d;

  // YYYY-MM-DD
  var m1 = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m1) {
    var y = parseInt(m1[1]); if (y > 2500) y -= 543;
    return new Date(y, parseInt(m1[2]) - 1, parseInt(m1[3]));
  }
  // DD/MM/YYYY or DD-MM-YYYY
  var m2 = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m2) {
    var y2 = parseInt(m2[3]); if (y2 > 2500) y2 -= 543; if (y2 < 100) y2 += 2000;
    return new Date(y2, parseInt(m2[2]) - 1, parseInt(m2[1]));
  }
  return null;
}

/**
 * ส่งการแจ้งเตือน overdue (เรียกจาก Time Trigger หรือ Manual)
 * @param {boolean} dryRun - ถ้า true จะไม่ส่งจริง แค่คืนรายการ
 */
function runOverdueNotifications(dryRun) {
  var cfg  = getOverdueConfig();
  var list = findOverdueEquipment();
  var sent = [], skipped = [], errors = [];

  // reminderDays: 0=วันครบกำหนด, 1=เกิน 1 วัน, -1=แจ้งล่วงหน้า 1 วัน
  var reminderDays = cfg.reminderDays || [0, 1, 3];

  list.forEach(function(item) {
    var match = reminderDays.some(function(d) {
      return item.daysOverdue === d;
    });
    if (!match) { skipped.push(item.id); return; }
    if (dryRun) { sent.push(item); return; }

    try {
      var label = item.daysOverdue === 0 ? 'ครบกำหนดคืนวันนี้'
        : item.daysOverdue > 0 ? 'เกินกำหนดคืน ' + item.daysOverdue + ' วัน'
        : 'ครบกำหนดในอีก ' + Math.abs(item.daysOverdue) + ' วัน';

      // LINE
      if (cfg.sendLine) {
        var lineUserId = getLineUserIdBySid(item.sid);
        if (lineUserId) {
          var lineMsg = '⏰ แจ้งเตือนการคืนอุปกรณ์\n'
            + '━━━━━━━━━━━━━━━\n'
            + '🆔 รหัสรายการ: ' + item.id + '\n'
            + '📦 ' + label + '\n'
            + '📅 กำหนดคืน: ' + item.returnDate + '\n'
            + '━━━━━━━━━━━━━━━\n'
            + 'รายการ:\n' + item.items + '\n'
            + '━━━━━━━━━━━━━━━\n'
            + 'กรุณาติดต่อทีมงานเพื่อนัดคืนอุปกรณ์ด่วน ขอบคุณครับ';
          sendLineMessage(lineUserId, lineMsg);
        }
      }

      // Email
      if (cfg.sendEmail && item.email) {
        var urgencyColor = item.daysOverdue > 3 ? '#c62828' : item.daysOverdue > 0 ? '#e65100' : '#1565c0';
        var emailBody = buildEmailHtml({
          title:      '⏰ แจ้งเตือน: ' + label,
          headerIcon: '!',
          color:      urgencyColor,
          name:       item.name,
          intro:      'ระบบตรวจพบว่ามีรายการยืมอุปกรณ์ที่ ' + label + ' กรุณาติดต่อทีมงานเพื่อดำเนินการคืน',
          rows: [
            ['รหัสรายการ',    item.id],
            ['ชื่อ-นามสกุล',  item.name],
            ['กำหนดคืน',      item.returnDate],
            ['รายการอุปกรณ์', String(item.items).replace(/\n/g, '<br>')],
            ['สถานะ',         item.status]
          ],
          note: 'หากคืนอุปกรณ์ไปแล้ว กรุณาแจ้ง Admin เพื่ออัปเดตสถานะในระบบ'
        });
        sendConfirmEmail(item.email, '[Media EDU Club] แจ้งเตือนคืนอุปกรณ์ — ' + label, emailBody);
        logAuditEvent('equipment', item.id, item.name, item.status, 'OVERDUE_NOTIFIED', 'System', label);
      }

      sent.push(item);
    } catch(e) {
      errors.push({ id: item.id, error: e.message });
      Logger.log('Overdue notify error for ' + item.id + ': ' + e.message);
    }
  });

  return {
    total:       list.length,
    matched:     sent.length,
    skipped:     skipped.length,
    errors:      errors.length,
    sentItems:   sent,
    errorDetail: errors,
    timestamp:   getThaiTimestamp()
  };
}

/**
 * สร้าง Time-based Trigger รันทุกวัน 08:00
 */
function setupOverdueTrigger() {
  // ลบ trigger เก่าก่อน
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyOverdueCheck') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyOverdueCheck')
    .timeBased().everyDays(1).atHour(8).create();
  return true;
}

function deleteTrigger(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(t);
  });
  return true;
}

function checkTriggerStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var found = triggers.filter(function(t) {
    return t.getHandlerFunction() === 'dailyOverdueCheck';
  });
  return { active: found.length > 0, count: found.length };
}

/** ฟังก์ชันที่ Trigger จะเรียก */
function dailyOverdueCheck() {
  var cfg = getOverdueConfig();
  if (!cfg.enabled) return;
  var result = runOverdueNotifications(false);
  Logger.log('dailyOverdueCheck: sent=' + result.matched + ', errors=' + result.errors);
}


/** รวม config + trigger status สำหรับ Admin UI */
function getOverdueSettingsAndStatus() {
  return {
    config:  getOverdueConfig(),
    trigger: checkTriggerStatus()
  };
}

// ==========================================
// LINE Admin Helpers
// ==========================================

function sendTestLineMessage(sid, message) {
  var userId = getLineUserIdBySid(sid);
  if (!userId) return false;
  sendLineMessage(userId, message || '🧪 ทดสอบระบบแจ้งเตือน Media EDU Club ✅');
  return true;
}

function testLineAdminNotify() {
  notifyAdmins('🧪 ทดสอบระบบแจ้งเตือน Admin\nMedia EDU Club CMU. — ระบบทำงานปกติ ✅');
  return true;
}

function removeLineUser(sid) {
  var sheet = getOrCreateLineUsersSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(sid).trim()) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

// ==========================================
// LINE — Session & Tracking Lookup
// ==========================================

/** เก็บ session ชั่วคราวว่า user กำลังรออะไร (หมดอายุ 10 นาที) */
function setUserSession(userId, state) {
  var key = 'LINE_SESSION_' + userId;
  var data = JSON.stringify({ state: state, ts: Date.now() });
  PropertiesService.getScriptProperties().setProperty(key, data);
}

function getUserSession(userId) {
  var key = 'LINE_SESSION_' + userId;
  var raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return null;
  try {
    var data = JSON.parse(raw);
    // หมดอายุ 10 นาที
    if (Date.now() - data.ts > 10 * 60 * 1000) {
      PropertiesService.getScriptProperties().deleteProperty(key);
      return null;
    }
    return data.state;
  } catch(e) { return null; }
}

function clearUserSession(userId) {
  PropertiesService.getScriptProperties().deleteProperty('LINE_SESSION_' + userId);
}

/** ผู้ใช้กด "ติดตามสถานะ" */
function handleTrackingCommand(userId) {
  var fp = getRegisteredProfileByLineId(userId);
  if (!fp || !fp.sid) {
    var oldSid = getRegisteredSidByLineId(userId);
    if (oldSid) fp = { sid: oldSid, name: '', phone: '' };
  }
  if (fp && fp.sid && fp.sid !== '-') {
    // ลงทะเบียนแล้ว → ค้นจาก SID ทันที
    handleTrackingLookup(userId, fp.sid);
    return;
  }
  // ยังไม่ลงทะเบียน
  sendLineMessage(userId,
    '🔍 ติดตามสถานะรายการ\n'
    + '━━━━━━━━━━━━━━━\n'
    + '⚠️ คุณยังไม่ได้ลงทะเบียนครับ\n\n'
    + 'กรุณาลงทะเบียนก่อนครับ\n'
    + '💡 พิมพ์ "ลงทะเบียน" เพื่อเริ่มเลย'
  );
}

// ==========================================
// Registration Flow (clean rebuild)
// sessions: reg:sid | reg:name|{sid} | reg:phone|{sid}|{name}
// ==========================================

function startReg(userId) {
  setUserSession(userId, 'reg:sid');
  sendLineMessage(userId,
    '✍️ ลงทะเบียนรับการแจ้งเตือน\n'
    + '━━━━━━━━━━━━━━━\n'
    + 'ขั้นตอนที่ 1/4\n\n'
    + '🎓 กรุณาส่งรหัสนักศึกษา\n'
    + 'เช่น: 650210343\n\n'
    + '(ศิษย์เก่า/บุคลากร ส่งรหัสบุคลากร\n'
    + 'หรือพิมพ์ "-" ถ้าไม่มีรหัส)\n\n'
    + '💡 พิมพ์ "ยกเลิก" เพื่อออก'
  );
}

function regGetSid(userId, text) {
  text = text.trim();
  if (text === 'ยกเลิก') { clearUserSession(userId); sendLineMessage(userId, '↩️ ยกเลิกแล้วครับ'); return; }
  var sid = text.toUpperCase();
  if (sid !== '-' && !/^[A-Z0-9]{3,15}$/.test(sid)) {
    sendLineMessage(userId, '⚠️ รหัสไม่ถูกต้องครับ\nกรุณาส่งรหัสนักศึกษา หรือพิมพ์ "-" ถ้าไม่มี');
    return;
  }
  setUserSession(userId, 'reg:name|' + sid);
  sendLineMessage(userId,
    'ขั้นตอนที่ 2/4\n━━━━━━━━━━━━━━━\n\n'
    + '👤 กรุณาส่งชื่อ-นามสกุล\n'
    + 'เช่น: สมชาย ใจดี\n\n'
    + '💡 พิมพ์ "ยกเลิก" เพื่อออก'
  );
}

function regGetName(userId, text) {
  text = text.trim();
  if (text === 'ยกเลิก') { clearUserSession(userId); sendLineMessage(userId, '↩️ ยกเลิกแล้วครับ'); return; }
  if (text.length < 2) { sendLineMessage(userId, '⚠️ กรุณาส่งชื่อ-นามสกุลให้ครบครับ'); return; }
  var sid = (getUserSession(userId) || '').replace('reg:name|', '');
  setUserSession(userId, 'reg:phone|' + sid + '|' + text);
  sendLineMessage(userId,
    'ขั้นตอนที่ 3/4\n━━━━━━━━━━━━━━━\n\n'
    + '📞 กรุณาส่งเบอร์โทรศัพท์\n'
    + 'เช่น: 0812345678\n\n'
    + 'หรือพิมพ์ "-" ถ้าไม่ต้องการระบุ\n'
    + '💡 พิมพ์ "ยกเลิก" เพื่อออก'
  );
}

function regGetPhone(userId, text) {
  text = text.trim();
  if (text === 'ยกเลิก') { clearUserSession(userId); sendLineMessage(userId, '↩️ ยกเลิกแล้วครับ'); return; }
  var parts = (getUserSession(userId) || '').replace('reg:phone|', '').split('|');
  var sid   = parts[0] || '-';
  var name  = parts.slice(1).join('|');
  var phone = text === '-' ? '' : text.replace(/[^0-9+]/g, '');
  if (phone && !/^[0-9+]{8,12}$/.test(phone)) {
    sendLineMessage(userId, '⚠️ เบอร์โทรไม่ถูกต้องครับ\nกรุณาส่งเบอร์ 8-12 หลัก หรือพิมพ์ "-" เพื่อข้าม');
    return;
  }
  // ต่อไปถามอีเมล — เก็บ phone ใน session
  setUserSession(userId, 'reg:email|' + sid + '|' + name + '|' + phone);
  sendLineMessage(userId,
    'ขั้นตอนที่ 4/4\n━━━━━━━━━━━━━━━\n\n'
    + '📧 กรุณาส่งอีเมลของคุณ\n'
    + 'เช่น: student@cmu.ac.th\n\n'
    + 'หรือพิมพ์ "-" ถ้าไม่ต้องการระบุ\n'
    + '💡 พิมพ์ "ยกเลิก" เพื่อออก'
  );
}

function regGetEmail(userId, text) {
  text = text.trim();
  if (text === 'ยกเลิก') { clearUserSession(userId); sendLineMessage(userId, '↩️ ยกเลิกแล้วครับ'); return; }
  var parts = (getUserSession(userId) || '').replace('reg:email|', '').split('|');
  var sid   = parts[0] || '-';
  var phone = parts[parts.length - 1] || '';
  var name  = parts.slice(1, parts.length - 1).join('|');
  var email = text === '-' ? '' : text.toLowerCase().trim();
  // validate email
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    sendLineMessage(userId, '⚠️ รูปแบบอีเมลไม่ถูกต้องครับ\nเช่น student@cmu.ac.th\nหรือพิมพ์ "-" เพื่อข้าม');
    return;
  }
  clearUserSession(userId);
  var result = registerLineUser(sid, userId, '', name, phone, '', email);
  sendLineMessage(userId,
    '🎉 ' + (result === 'updated' ? 'อัปเดตข้อมูลสำเร็จ!' : 'ลงทะเบียนสำเร็จ!') + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + '🎓 รหัส: '  + sid              + '\n'
    + '👤 ชื่อ: '  + name             + '\n'
    + '📞 เบอร์: ' + (phone || '-')   + '\n'
    + '📧 อีเมล: ' + (email || '-')   + '\n'
    + '━━━━━━━━━━━━━━━\n'
    + 'ตอนนี้คุณจะได้รับการแจ้งเตือนผ่าน LINE ครับ!\n\n'
    + '💡 พิมพ์ "ติดตามสถานะ" เพื่อตรวจสอบรายการของคุณ'
  );
}

/** ดึงและแสดงสถานะรายการจาก keyword — ใช้ Templates */
function handleTrackingLookup(userId, keyword) {
  clearUserSession(userId);
  keyword = String(keyword || '').trim();

  if (keyword.length < 2) {
    sendLineMessage(userId, '⚠️ กรุณากรอกข้อมูลอย่างน้อย 2 ตัวอักษรครับ');
    return;
  }

  var results = [];
  try {
    results = searchTrackingDataAll(keyword, 'all');
  } catch(e) {
    results = searchTrackingData(keyword, 'all') || [];
  }

  if (!results || results.length === 0) {
    sendLineMessage(userId, fillTpl('track_not_found', { keyword: keyword }));
    return;
  }

  var msgLines = [fillTpl('track_header', { keyword: keyword }), '━━━━━━━━━━━━━━━'];
  var shown = results.slice(0, 5);

  shown.forEach(function(r) {
    var typeLabel;
    if      (r.type === 'order')     typeLabel = getTpl('label_order');
    else if (r.type === 'equipment') typeLabel = getTpl('label_equip');
    else if (r.type === 'jersey')    typeLabel = getTpl('label_jersey');
    else                              typeLabel = getTpl('label_booking');

    msgLines.push('');
    msgLines.push(typeLabel);
    msgLines.push(fillTpl('track_name_line',   { name:   r.name   || '—' }));
    msgLines.push(fillTpl('track_status_line', { status: r.status || '—' }));

    if (r.type === 'equipment') {
      msgLines.push(fillTpl('track_equip_return', { date: r.returnTime || r.returnDate || '—' }));
    } else if (r.type === 'booking') {
      msgLines.push(fillTpl('track_book_room', { room: r.room || '—' }));
      msgLines.push(fillTpl('track_book_date', { date: r.date || '—' }));
    } else if (r.type === 'order') {
      msgLines.push(fillTpl('track_order_id',    { id:    r.id    || '—' }));
      msgLines.push(fillTpl('track_order_total', { total: r.total || '—' }));
    } else if (r.type === 'jersey') {
      var itemsSummary = String(r.items || '').split('\n')
        .filter(function(l){ return l.trim(); }).slice(0, 3).join(' | ');
      if (itemsSummary) msgLines.push(getTpl('label_jersey') + ' ' + itemsSummary);
      msgLines.push((r.shipping || '').indexOf('ไปรษณีย์') !== -1
        ? getTpl('track_jersey_postal') : getTpl('track_jersey_pickup'));
    }
  });

  if (results.length > 5) {
    msgLines.push('');
    msgLines.push(fillTpl('track_more', { n: String(results.length - 5) }));
  }

  msgLines.push('');
  msgLines.push('━━━━━━━━━━━━━━━');
  msgLines.push(getTpl('track_footer'));

  sendLineMessage(userId, msgLines.join('\n'));
}

/** บันทึก Webapp URL สำหรับแสดงในข้อความ */

/** ดึง SID ของผู้ใช้ LINE (ใช้ต่อท้าย URL) */
function getUserSidSuffix(userId) {
  var fp = getRegisteredProfileByLineId(userId);
  var sid = fp && fp.sid ? fp.sid : null;
  if (!sid) sid = getRegisteredSidByLineId(userId);
  return sid ? ('?sid=' + encodeURIComponent(sid)) : '';
}

// ==========================================
// Auto-fill: ดึงข้อมูลผู้ใช้จาก SID (สำหรับ web form)
// ==========================================
function getUserProfileBySid(sid) {
  try {
    if (!sid) return null;
    sid = String(sid).trim().toUpperCase();
    var data = getLineUsersData();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || !row[0]) continue;
      if (String(row[0]).trim().toUpperCase() === sid) {
        return {
          sid:   String(row[0] || '').trim(),
          name:  String(row[4] || '').trim(),   // col E = ชื่อ-นามสกุล
          phone: String(row[5] || '').trim(),   // col F = เบอร์โทร
          type:  String(row[6] || '').trim(),   // col G = ประเภท
          email: String(row[7] || '').trim()    // col H = อีเมล
        };
      }
    }
    return null;
  } catch(e) {
    Logger.log('getUserProfileBySid error: ' + e.message);
    return null;
  }
}


function setWebAppUrl(url) {
  PropertiesService.getScriptProperties().setProperty('WEBAPP_URL', url);
  return true;
}

// ==========================================
// ==========================================
// MODULE: ระบบจุดรับสินค้า + EDUJERSEY
// (เพิ่มเติมโดยระบบ — ต่อจาก Code.gs เดิม)
// ==========================================
// ==========================================
// ==========================================
// MODULE: ระบบจุดรับสินค้า v2
// (รองรับทั้ง Orders sheet และ EDUJERSEY Google Form)
// ==========================================
// วิธีใช้: คัดลอกทั้งหมดต่อท้าย Code.gs เดิม
// ==========================================

// ===== ตั้งค่า Sheet ID ของ EDUJERSEY =====
// → เปลี่ยน ID ตรงนี้ให้ตรงกับ Spreadsheet ของคุณ
var EDUJERSEY_SHEET_ID   = '1RyG4abD8GJOln97H4IjB-NVXzyWSIc5_wNx6YdVVZB8';
var EDUJERSEY_SHEET_NAME = 'การตอบกลับของแบบฟอร์ม 1'; // ชื่อ Sheet tab

// คอลัมน์ที่ใช้ (0-indexed)
// [0]  ประทับเวลา
// [1]  สถานะผู้สั่ง (buyerType)
// [2]  ชื่อ-นามสกุล
// [3]  ชั้นปี
// [4]  รหัสนักศึกษา
// [5]  สาขาวิชา
// [6]  เบอร์โทร
// [7]  ช่องทางติดต่อ
// [8]  รุ่นเสื้อ #1, [9] ไซซ์ #1, [10] ไซซ์พิเศษ #1
// [11] ช่องทางจัดส่ง
// [12] ที่อยู่ (กรณีไปรษณีย์)
// [13] จำนวนเงิน
// [14] URL สลิป
// [15] เวลาชำระ
// [16] จำนวน #1
// [17] รุ่น #2, [18] ต้องการเพิ่ม?, [19] ไซซ์ #2, [20] ไซซ์พิเศษ #2, [21] จำนวน #2
// [22] รุ่น #3, [23] ไซซ์ #3, [24] ไซซ์พิเศษ #3, [25] จำนวน #3
// [26] รุ่น #4, [27] ไซซ์ #4, [28] ไซซ์พิเศษ #4, [29] จำนวน #4
// [33] สถานะการมอบ (เพิ่มเองในคอลัมน์ AH)

var EDUJERSEY_STATUS_COL = 34; // คอลัมน์ที่ 34 (AH) — เพิ่มโดยระบบอัตโนมัติ

// ==========================================
// 1. ค้นหา EDUJERSEY
// ==========================================

// ==========================================
// 1. ค้นหา EDUJERSEY (อัปเดตให้รองรับ Filter จากเซิร์ฟเวอร์)
// ==========================================
// ==========================================
// 1. ค้นหา EDUJERSEY (Backend เหมาข้อมูลส่งให้ Frontend)
// ==========================================
function searchEduJerseyForPickup(keyword) {
  keyword = String(keyword || '').trim().toLowerCase();

  var isPhone = /^0\d{4,}/.test(keyword);
  var isSid   = /^\d{4,}$/.test(keyword) && !isPhone;
  var kwDigits = keyword.replace(/[^0-9]/g, '');

  try {
    var ss    = SpreadsheetApp.openById(EDUJERSEY_SHEET_ID);
    var sheet = ss.getSheetByName(EDUJERSEY_SHEET_NAME);
    if (!sheet) sheet = ss.getSheets()[0];
    if (!sheet) return [];

    var data    = sheet.getDataRange().getDisplayValues();
    var results = [];
    
    // ถ้าไม่ได้พิมพ์ค้นหา (กดตัวกรองเฉยๆ) ให้ดึง 150 รายการล่าสุดมาเผื่อไว้กรอง
    var fetchAll = (keyword.length === 0);

    for (var i = data.length - 1; i > 0; i--) {
      var row      = data[i];
      var name     = String(row[2] || '');
      var sid      = String(row[4] || '').replace(/\s/g, '');
      var phoneRaw = String(row[6] || '');
      var phoneD   = phoneRaw.replace(/[^0-9]/g, '');
      var contact  = String(row[7] || '');
      var shippingName = String(row[11] || '');
      var statusIdx = EDUJERSEY_STATUS_COL - 1;
      var status    = (row.length > statusIdx && row[statusIdx]) ? String(row[statusIdx]).trim() : 'รอมอบ';

      var matched = false;

      if (fetchAll) {
        matched = true; // เหมาหมดให้ Frontend จัดการ
      } else {
        if (isPhone) {
          matched = phoneD === kwDigits;
        } else if (isSid) {
          matched = sid === keyword || sid.startsWith(keyword);
        } else {
          var searchText = [name, sid, phoneRaw, contact].join(' ').toLowerCase();
          matched = searchText.indexOf(keyword) !== -1;
        }
      }

      if (!matched) continue;

      var items = [];
      var groups = [
        { gen: row[8], size: row[9], special: row[10], qty: row[16] },
        { gen: row[17], size: row[19], special: row[20], qty: row[21] },
        { gen: row[22], size: row[23], special: row[24], qty: row[25] },
        { gen: row[26], size: row[27], special: row[28], qty: row[29] }
      ];
      groups.forEach(function(g) {
        var gGen = String(g.gen || '').trim();
        var gSz  = String(g.size || '').trim();
        if (gGen || gSz) items.push({ gen: gGen, size: gSz, special: String(g.special || '').trim(), qty: parseInt(g.qty || '1') || 1 });
      });

      results.push({
        row:         i + 1,
        timestamp:   String(row[0]  || ''),
        buyerType:   String(row[1]  || ''),
        name:        name,
        year:        String(row[3]  || ''),
        sid:         sid,
        major:       String(row[5]  || ''),
        phone:       phoneRaw,
        contact:     contact,
        items:       items,
        shipping:    shippingName,
        address:     String(row[12] || ''),
        total:       String(row[13] || '0'),
        slipUrl:     String(row[14] || ''),
        paymentTime: String(row[15] || ''),
        status:      status,
        isPostal:    shippingName.indexOf('ไปรษณีย์') !== -1
      });

      // ลิมิตแค่ 150 ออเดอร์ล่าสุดเพื่อความรวดเร็วของระบบ
      if (results.length >= 150) break; 
    }
    return results;
  } catch(e) {
    throw new Error('เข้าถึง EDUJERSEY Sheet ไม่ได้: ' + e.message);
  }
}

// ==========================================
// 2. อัปเดตสถานะ EDUJERSEY
// ==========================================

function updateEduJerseyStatus(row, status, adminName) {
  try {
    var ss    = SpreadsheetApp.openById(EDUJERSEY_SHEET_ID);
    var sheet = ss.getSheetByName(EDUJERSEY_SHEET_NAME);
    if (!sheet) sheet = ss.getSheets()[0];
    if (!sheet) return { ok: false, msg: 'ไม่พบ Sheet' };

    // ตรวจสอบว่ามีหัวคอลัมน์สถานะหรือยัง
    var lastCol    = sheet.getLastColumn();
    var statusCol  = EDUJERSEY_STATUS_COL;

    if (lastCol < statusCol) {
      // เพิ่ม header ถ้ายังไม่มี
      sheet.getRange(1, statusCol).setValue('สถานะการมอบ');
      sheet.getRange(1, statusCol).setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    }

    // เขียนสถานะ + ผู้ดำเนินการ + timestamp
    sheet.getRange(row, statusCol).setValue(status);

    // บันทึกว่าใครอัปเดต + เมื่อไร (คอลัมน์ถัดไป)
    var metaCol = statusCol + 1;
    if (lastCol < metaCol) {
      sheet.getRange(1, metaCol).setValue('ผู้อัปเดต / วันเวลา');
      sheet.getRange(1, metaCol).setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    }
    sheet.getRange(row, metaCol).setValue((adminName || 'Admin') + ' | ' + getThaiTimestamp());

    return { ok: true, status: status };

  } catch(e) {
    Logger.log('updateEduJerseyStatus error: ' + e.message);
    return { ok: false, msg: e.message };
  }
}

// ==========================================
// 3. ค้นหาออเดอร์ทั่วไป (Orders sheet)
// ==========================================
function searchOrdersForPickup(keyword) {
  var sheet = getDB().getSheetByName('Orders');
  if (!sheet) return [];

  keyword = String(keyword || '').trim().toLowerCase();

  var isPhone  = /^0\d{4,}/.test(keyword);
  var isSid    = /^\d{4,}$/.test(keyword) && !isPhone;
  var kwDigits = keyword.replace(/[^0-9]/g, '');
  var fetchAll = (keyword.length === 0);

  var data    = sheet.getDataRange().getDisplayValues();
  var results = [];

  for (var i = data.length - 1; i > 0; i--) {
    var id        = String(data[i][0]  || '');
    var phoneRaw  = String(data[i][1]  || '');
    var phoneD    = phoneRaw.replace(/[^0-9]/g, '');
    var name      = String(data[i][2]  || '');
    var details   = String(data[i][3]  || '');
    var total     = String(data[i][4]  || '0');
    var status    = String(data[i][5]  || '');
    var date      = String(data[i][6]  || '');
    var buyerInfo = String(data[i][12] || '').replace(/\s/g, '');

    var matched = false;
    if (fetchAll) {
      matched = true;
    } else {
      if (isPhone) {
        matched = phoneD === kwDigits;
      } else if (isSid) {
        matched = buyerInfo === keyword || buyerInfo.startsWith(keyword);
      } else {
        var searchText = [id, phoneRaw, name, buyerInfo].join(' ').toLowerCase();
        matched = searchText.indexOf(keyword) !== -1;
      }
    }

    if (!matched) continue;

    results.push({
      row: i+1, id: id, phone: phoneRaw, name: name,
      details: details, total: total, status: status, date: date,
      approver: String(data[i][7] || '') || '-', email: String(data[i][8] || ''), 
      address: String(data[i][9] || ''), slipUrl: String(data[i][10] || ''), 
      buyerType: String(data[i][11] || ''), buyerInfo: buyerInfo,
      shippingName: String(data[i][13] || ''), shippingFee: String(data[i][14] || '0'), 
      contact: String(data[i][15] || '')
    });
    
    if (results.length >= 150) break;
  }
  return results;
}

// ==========================================
// 4. อัปเดตสถานะออเดอร์ทั่วไป + แจ้ง LINE/Email
// ==========================================

function getLineUserIdByPhone(phone) {
  if (!phone) return null;
  var phoneClean = String(phone).replace(/[^0-9]/g, '');
  var orderSheet = getDB().getSheetByName('Orders');
  if (!orderSheet) return null;
  var orderData = orderSheet.getDataRange().getValues();
  for (var i = 1; i < orderData.length; i++) {
    var rowPhone = String(orderData[i][1] || '').replace(/[^0-9]/g, '');
    if (rowPhone !== phoneClean) continue;
    var buyerInfo = String(orderData[i][12] || '');
    var sidMatch  = buyerInfo.match(/\b(\d{9,11})\b/);
    if (sidMatch) {
      var uid = getLineUserIdBySid(sidMatch[1]);
      if (uid) return uid;
    }
    break;
  }
  return null;
}

function updateOrderStatusPickup(row, status, adminName) {
  var sheet = getDB().getSheetByName('Orders');
  if (!sheet) return { ok: false, msg: 'ไม่พบชีต Orders' };

  try {
    var oldStatus = sheet.getRange(row, 6).getValue();
    sheet.getRange(row, 6).setValue(status);
    sheet.getRange(row, 8).setValue(adminName || 'Admin');

    var rowData      = sheet.getRange(row, 1, 1, 16).getValues()[0];
    var orderId      = String(rowData[0]  || '');
    var phone        = String(rowData[1]  || '');
    var name         = String(rowData[2]  || '');
    var details      = String(rowData[3]  || '');
    var total        = rowData[4];
    var email        = String(rowData[8]  || '');
    var buyerInfo    = String(rowData[12] || '');
    var shippingName = String(rowData[13] || '');

    logAuditEvent('order', orderId, name, oldStatus, status, adminName || 'Admin', 'pickup-counter');

    // ---- Email ----
    var emailSent = false;
    if (email) {
      var subjectO = '[Media EDU Club] อัปเดตสถานะออเดอร์ ' + orderId + ' — ' + status;
      var bodyO = buildEmailHtml({
        title:      getStatusTitle(status),
        headerIcon: 'S',
        color:      getStatusColor(status),
        name:       name,
        intro:      'ทีมงานได้ดำเนินการอัปเดตสถานะคำสั่งซื้อของคุณแล้ว รายละเอียดมีดังนี้:',
        rows: [
          ['รหัสออเดอร์',   orderId],
          ['ชื่อลูกค้า',    name],
          ['เบอร์โทรศัพท์', phone],
          ['รายการสินค้า',  String(details).replace(/\n/g, '<br>')],
          ['ยอดรวม',        Number(total || 0).toLocaleString() + ' บาท'],
          ['ช่องทางจัดส่ง', shippingName || '-'],
          ['สถานะใหม่',     status],
          ['ดำเนินการโดย',  adminName || 'Admin']
        ],
        note: 'หากมีข้อสงสัย กรุณาติดต่อทีมงาน Media EDU Club CMU.'
      });
      sendConfirmEmail(email, subjectO, bodyO);
      emailSent = true;
    }

    // ---- LINE ----
    var lineNotified = false;
    var lineUserId   = null;
    var sidMatch2    = buyerInfo.match(/\b(\d{9,11})\b/);
    if (sidMatch2) lineUserId = getLineUserIdBySid(sidMatch2[1]);
    if (!lineUserId) lineUserId = getLineUserIdByPhone(phone);

    if (lineUserId) {
      var isCancelled = status.indexOf('ยกเลิก') !== -1;
      var isPickedUp  = status === 'รับแล้ว' || status.indexOf('รับแล้ว') !== -1;
      var statusIcon  = isCancelled ? '❌' : isPickedUp ? '🎉' : '🔔';
      var closingMsg  = isCancelled ? 'หากมีข้อสงสัยกรุณาติดต่อทีมงานครับ'
                      : isPickedUp  ? '🎉 ขอบคุณที่ใช้บริการ Media EDU Club CMU!'
                      : 'ทีมงานกำลังดำเนินการ จะแจ้งให้ทราบอีกครั้งครับ';
      var itemLines   = String(details).split('\n').filter(function(l) { return l.trim(); });
      var itemsText   = itemLines.map(function(l, i) { return '  ' + (i+1) + '. ' + l.trim(); }).join('\n');
      var lineMsg = statusIcon + ' อัปเดตสถานะออเดอร์\n'
        + '━━━━━━━━━━━━━━━\n'
        + '🆔 รหัสออเดอร์: '   + orderId + '\n'
        + '👤 ชื่อ: '           + name    + '\n'
        + '━━━━━━━━━━━━━━━\n'
        + '🛒 รายการสินค้า:\n' + itemsText + '\n'
        + '━━━━━━━━━━━━━━━\n'
        + '💰 ยอดรวม: '        + Number(total || 0).toLocaleString() + ' บาท\n'
        + '📌 สถานะ: '         + status + '\n'
        + '👨‍💼 ดำเนินการโดย: '  + (adminName || 'Admin') + '\n'
        + '━━━━━━━━━━━━━━━\n'
        + closingMsg;
      sendLineMessage(lineUserId, lineMsg);
      lineNotified = true;
    }

    return { ok: true, lineNotified: lineNotified, emailSent: emailSent, status: status };

  } catch(e) {
    Logger.log('updateOrderStatusPickup error: ' + e.message);
    return { ok: false, msg: e.message };
  }
}

// ==========================================
// 5. Bulk Update EDUJERSEY Status
// ==========================================

/**
 * อัปเดตสถานะ EDUJERSEY หลายแถวพร้อมกัน
 * @param {Array} updates - [{row: number, status: string}, ...]
 * @param {string} adminName
 */
function updateEduJerseyStatusBulk(updates, adminName) {
  if (!updates || updates.length === 0) return { ok: false, msg: 'ไม่มีรายการที่จะอัปเดต' };

  try {
    var ss    = SpreadsheetApp.openById(EDUJERSEY_SHEET_ID);
    var sheet = ss.getSheetByName(EDUJERSEY_SHEET_NAME);
    if (!sheet) sheet = ss.getSheets()[0];
    if (!sheet) return { ok: false, msg: 'ไม่พบ Sheet' };

    var lastCol   = sheet.getLastColumn();
    var statusCol = EDUJERSEY_STATUS_COL;
    var metaCol   = statusCol + 1;

    // เพิ่ม header ถ้ายังไม่มี
    if (lastCol < statusCol) {
      sheet.getRange(1, statusCol).setValue('สถานะการมอบ');
      sheet.getRange(1, statusCol).setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    }
    if (lastCol < metaCol) {
      sheet.getRange(1, metaCol).setValue('ผู้อัปเดต / วันเวลา');
      sheet.getRange(1, metaCol).setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    }

    var timestamp = getThaiTimestamp();
    var meta      = (adminName || 'Admin') + ' | ' + timestamp;
    var updated   = 0;

    updates.forEach(function(u) {
      try {
        sheet.getRange(u.row, statusCol).setValue(u.status);
        sheet.getRange(u.row, metaCol).setValue(meta);
        updated++;
      } catch(e) {
        Logger.log('bulk update row ' + u.row + ' error: ' + e.message);
      }
    });

    return { ok: true, updated: updated, total: updates.length, timestamp: timestamp };

  } catch(e) {
    Logger.log('updateEduJerseyStatusBulk error: ' + e.message);
    return { ok: false, msg: e.message };
  }
}

// ==========================================
// 6. ติดตามสถานะ EDUJERSEY (สำหรับลูกค้า)
// ==========================================

/**
 * ค้นหาแบบรวม: Orders + Equipment + Bookings + EDUJERSEY
 * ใช้แทน searchTrackingData เดิม เพื่อรองรับ jersey type
 */
function searchTrackingDataAll(keyword, typeFilter) {
  keyword    = String(keyword || '').trim().toLowerCase();
  typeFilter = typeFilter || 'all';
  if (!keyword || keyword.length < 2) return [];

  var results = [];

  // ดึงจากระบบเดิม (order / equipment / booking)
  if (typeFilter !== 'jersey') {
    var existing = searchTrackingData(keyword, typeFilter === 'jersey' ? 'none' : typeFilter);
    results = results.concat(existing || []);
  }

  // ดึงจาก EDUJERSEY
  if (typeFilter === 'all' || typeFilter === 'jersey') {
    var jerseyResults = searchEduJerseyTracking(keyword);
    results = results.concat(jerseyResults || []);
  }

  return results;
}

/**
 * ค้นหาใน EDUJERSEY sheet สำหรับหน้าติดตามสถานะ
 */
function searchEduJerseyTracking(keyword) {
  keyword = String(keyword || '').trim().toLowerCase();
  if (!keyword || keyword.length < 2) return [];

  try {
    var ss    = SpreadsheetApp.openById(EDUJERSEY_SHEET_ID);
    var sheet = ss.getSheetByName(EDUJERSEY_SHEET_NAME);
    if (!sheet) sheet = ss.getSheets()[0];
    if (!sheet) return [];

    var data    = sheet.getDataRange().getDisplayValues();
    var results = [];

    for (var i = data.length - 1; i > 0; i--) {
      var row   = data[i];
      var name  = String(row[2]  || '');
      var sid   = String(row[4]  || '');
      var phone = String(row[6]  || '');

      var searchText = [name, sid, phone].join(' ').toLowerCase();
      if (searchText.indexOf(keyword) === -1) continue;

      // ดึงรายการเสื้อทั้งหมด (ชุดที่ 1-4)
      var itemParts = [];
      var groups = [
        { gen: row[8], size: row[9], special: row[10], qty: row[16] },
        { gen: row[17], size: row[19], special: row[20], qty: row[21] },
        { gen: row[22], size: row[23], special: row[24], qty: row[25] },
        { gen: row[26], size: row[27], special: row[28], qty: row[29] }
      ];
      groups.forEach(function(g) {
        var gen  = String(g.gen  || '').trim();
        var size = String(g.size || '').trim();
        if (!gen && !size) return;
        var qty  = parseInt(g.qty || '1') || 1;
        var sp   = String(g.special || '').trim();
        var line = 'รุ่น ' + (gen||'—') + ' ไซซ์ ' + (size||'—') + ' ×' + qty;
        if (sp && sp !== '-') line += ' (' + sp + ')';
        itemParts.push(line);
      });

      var statusIdx = EDUJERSEY_STATUS_COL - 1;
      var status    = (row.length > statusIdx && row[statusIdx]) ? String(row[statusIdx]).trim() : 'รอมอบ';
      var isPostal  = String(row[11] || '').indexOf('ไปรษณีย์') !== -1;

      results.push({
        type:        'jersey',
        id:          'JERSEY-' + i,
        row:         i + 1,
        name:        name,
        phone:       phone,
        sid:         sid,
        major:       String(row[5]  || ''),
        contact:     String(row[7]  || ''),
        items:       itemParts.join('\n'),
        total:       String(row[13] || '0'),
        slipUrl:     String(row[14] || ''),
        paymentTime: String(row[15] || ''),
        shipping:    String(row[11] || ''),
        isPostal:    isPostal,
        address:     String(row[12] || ''),
        buyerType:   String(row[1]  || ''),
        status:      status,
        date:        String(row[0]  || ''),
        approver:    '-'
      });

      if (results.length >= 20) break;
    }

    return results;

  } catch(e) {
    Logger.log('searchEduJerseyTracking error: ' + e.message);
    return [];
  }
}

// ==========================================
// 7. สรุปภาพรวมการรับเสื้อ EDUJERSEY
// ==========================================

/**
 * ดึงข้อมูลสรุป: ใครรับแล้ว / ยังไม่รับ พร้อมไซซ์
 * @returns {{done, pending, gens, totalSizes}}
 */
function getEduJerseySummary() {
  try {
    var ss    = SpreadsheetApp.openById(EDUJERSEY_SHEET_ID);
    var sheet = ss.getSheetByName(EDUJERSEY_SHEET_NAME);
    if (!sheet) sheet = ss.getSheets()[0];
    if (!sheet) return { done: [], pending: [], gens: [], sizes: [], history: [] };

    var data      = sheet.getDataRange().getDisplayValues();
    var statusIdx = EDUJERSEY_STATUS_COL - 1;
    var metaIdx   = statusIdx + 1; // คอลัมน์ที่เก็บชื่อคนจ่ายของ (Admin)
    var done      = [];
    var pending   = [];
    var history   = [];
    var genSet    = {};
    var sizeSet   = {};

    var doneStatuses = ['มอบแล้ว', 'รับแล้ว', 'จัดส่งแล้ว'];

    for (var i = 1; i < data.length; i++) {
      var row  = data[i];
      var name = String(row[2] || '').trim();
      if (!name) continue;

      var phone  = String(row[6] || '').trim();
      var status = (row.length > statusIdx && row[statusIdx]) ? String(row[statusIdx]).trim() : 'รอมอบ';
      var meta   = (row.length > metaIdx && row[metaIdx]) ? String(row[metaIdx]).trim() : '';

      // ดึงชื่อคนจ่ายของ และเวลา
      var updater = 'Admin', updateTime = '';
      if (meta) {
        var parts = meta.split('|');
        updater = (parts[0] || 'Admin').trim();
        updateTime = (parts[1] || '').trim();
      }

      var sizes = [];
      var gens  = [];
      var groups = [
        { gen: row[8], size: row[9], qty: row[16] },
        { gen: row[17], size: row[19], qty: row[21] },
        { gen: row[22], size: row[23], qty: row[25] },
        { gen: row[26], size: row[27], qty: row[29] }
      ];
      
      groups.forEach(function(g) {
        var gen  = String(g.gen  || '').trim().toUpperCase();
        var size = String(g.size || '').trim().toUpperCase();
        var qty  = parseInt(g.qty || '1') || 1;
        if (!size || size === '-') return;
        for (var q = 0; q < Math.min(qty, 5); q++) sizes.push(size);
        sizeSet[size] = true;
        if (gen && gen !== '-') {
          gens.push(gen);
          genSet[gen] = true;
        }
      });
      
      var sizeDisplay = dedupeWithCount(sizes);
      var isPostal = String(row[11] || '').indexOf('ไปรษณีย์') !== -1;
      
      var person = {
        name: name, phone: phone, sizes: sizeDisplay, gens: gens,
        rawSizes: sizes, isPostal: isPostal, status: status,
        updater: updater, updateTime: updateTime, row: i+1
      };

      var isDone = doneStatuses.indexOf(status) !== -1;
      if (isDone) {
        done.push(person);
        if (updateTime) history.push(person); // ถ้ามีเวลาทำรายการ ให้เก็บลงประวัติ
      } else {
        pending.push(person);
      }
    }

    // เรียงประวัติจากเวลาล่าสุดขึ้นก่อน (ดักจับและแปลงเวลาให้แม่นยำที่สุด)
    history.sort(function(a, b) {
      function parseTime(tStr) {
        if (!tStr) return 0;
        // เผื่อเวลามาเป็นรูปแบบมาตรฐาน
        var d = new Date(tStr);
        if (!isNaN(d.getTime())) return d.getTime();
        
        // กรณีเวลาเป็นฟอร์แมตไทย เช่น "26/03/2567 14:30:00" หรือ "26/03/2024"
        var parts = tStr.split(' ');
        var dateParts = (parts[0] || '').split('/');
        var timePart = parts[1] || '00:00:00';
        
        if (dateParts.length === 3) {
          var day = dateParts[0];
          var month = dateParts[1];
          var year = parseInt(dateParts[2]);
          // ถ้าเป็น พ.ศ. ให้แปลงเป็น ค.ศ. เพื่อคำนวณ
          if (year > 2500) year -= 543; 
          
          var newDate = new Date(year + '/' + month + '/' + day + ' ' + timePart);
          if (!isNaN(newDate.getTime())) return newDate.getTime();
        }
        return 0;
      }
      // เรียงจากมากไปน้อย (เวลาล่าสุดอยู่บนสุด)
      return parseTime(b.updateTime) - parseTime(a.updateTime);
    });

    return {
      done:    done,
      pending: pending,
      history: history.slice(0, 100), // ดึงมาแค่ 100 รายการล่าสุดกันเว็บค้าง
      gens:    Object.keys(genSet).sort(),
      sizes:   Object.keys(sizeSet).sort(),
      timestamp: getThaiTimestamp()
    };

  } catch(e) {
    Logger.log('getEduJerseySummary error: ' + e.message);
    throw new Error('โหลดข้อมูลไม่ได้: ' + e.message);
  }
}

/**
 * แปลง array ของไซซ์ → array ที่รวมแสดง count เช่น ['M','M','L'] → ['M×2','L']
 */
function dedupeWithCount(arr) {
  var counts = {};
  arr.forEach(function(s) { counts[s] = (counts[s] || 0) + 1; });
  return Object.keys(counts).sort().map(function(s) {
    return counts[s] > 1 ? s + '×' + counts[s] : s;
  });
}

// ==========================================
// MODULE: LINE Message Templates
// แก้ไขข้อความ bot ได้จากหน้า Admin
// ==========================================

var LINE_TEMPLATES_SHEET = 'LINE_Templates';

/**
 * ค่า default ของทุก template
 * key → { label, default, group, hint }
 */
var LINE_TEMPLATE_DEFAULTS = {
  // ── Tracking: หัวข้อ & ท้าย ──
  'track_header':        { label:'หัวข้อผลค้นหา',         group:'ผลค้นหา',    hint:'ใช้ {keyword} แทนคำที่ค้นหา',         val:'📋 ผลค้นหา: "{keyword}"' },
  'track_footer':        { label:'ข้อความท้ายผลค้นหา',    group:'ผลค้นหา',    hint:'แสดงต่อท้ายผลทุกครั้ง',               val:'💡 ค้นหาใหม่: พิมพ์ "ติดตามสถานะ"' },
  'track_more':          { label:'มีรายการเพิ่มเติม',      group:'ผลค้นหา',    hint:'ใช้ {n} แทนจำนวนรายการที่เหลือ',      val:'... และอีก {n} รายการ\nดูทั้งหมดที่เว็บระบบได้เลยครับ' },
  'track_not_found':     { label:'ไม่พบรายการ',            group:'ผลค้นหา',    hint:'ใช้ {keyword} แทนคำที่ค้นหา',         val:'🔍 ไม่พบรายการ\n━━━━━━━━━━━━━━━\nไม่พบข้อมูลสำหรับ: "{keyword}"\n\n💡 ลองค้นหาด้วย:\n• ชื่อ-นามสกุล\n• รหัสนักศึกษา\n• เบอร์โทรศัพท์\n• รหัสรายการ\n\nพิมพ์ "ติดตามสถานะ" เพื่อค้นหาใหม่' },

  // ── Tracking: ป้ายชื่อประเภท ──
  'label_order':         { label:'ป้ายชื่อ ออเดอร์',      group:'ป้ายชื่อ',   hint:'แสดงเป็นหัวแต่ละรายการ',              val:'🛒 ออเดอร์' },
  'label_equip':         { label:'ป้ายชื่อ ยืมอุปกรณ์',   group:'ป้ายชื่อ',   hint:'แสดงเป็นหัวแต่ละรายการ',              val:'💻 ยืมอุปกรณ์' },
  'label_booking':       { label:'ป้ายชื่อ จองห้อง',       group:'ป้ายชื่อ',   hint:'แสดงเป็นหัวแต่ละรายการ',              val:'🏫 จองห้อง' },
  'label_jersey':        { label:'ป้ายชื่อ EDUJERSEY',     group:'ป้ายชื่อ',   hint:'แสดงเป็นหัวแต่ละรายการ',              val:'👕 EDUJERSEY' },

  // ── Tracking: รายละเอียดต่อรายการ ──
  'track_name_line':     { label:'บรรทัด ชื่อ',            group:'รายละเอียด', hint:'ใช้ {name}',                           val:'👤 {name}' },
  'track_status_line':   { label:'บรรทัด สถานะ',           group:'รายละเอียด', hint:'ใช้ {status}',                         val:'📌 สถานะ: {status}' },
  'track_equip_return':  { label:'บรรทัด วันคืนอุปกรณ์',  group:'รายละเอียด', hint:'ใช้ {date}',                           val:'📅 คืน: {date}' },
  'track_book_room':     { label:'บรรทัด ห้องที่จอง',      group:'รายละเอียด', hint:'ใช้ {room}',                           val:'🏫 ห้อง: {room}' },
  'track_book_date':     { label:'บรรทัด วันที่จอง',       group:'รายละเอียด', hint:'ใช้ {date}',                           val:'📅 วันที่: {date}' },
  'track_order_id':      { label:'บรรทัด รหัสออเดอร์',     group:'รายละเอียด', hint:'ใช้ {id}',                             val:'🆔 {id}' },
  'track_order_total':   { label:'บรรทัด ยอดชำระ',         group:'รายละเอียด', hint:'ใช้ {total}',                          val:'💰 ยอด: {total}' },
  'track_jersey_postal': { label:'ป้าย ส่งไปรษณีย์',       group:'รายละเอียด', hint:'แสดงเมื่อเลือกส่งไปรษณีย์',           val:'📮 ส่งไปรษณีย์' },
  'track_jersey_pickup': { label:'ป้าย รับที่สโมสร',        group:'รายละเอียด', hint:'แสดงเมื่อรับด้วยตัวเอง',              val:'🏫 รับที่สโมสร' },

  // ── ข้อความทั่วไปของ bot ──
  'msg_welcome_new':     { label:'ข้อความต้อนรับ (ใหม่)',  group:'ข้อความ bot', hint:'แสดงเมื่อ Add เพื่อนครั้งแรก',       val:'👋 สวัสดีครับ! ยินดีต้อนรับสู่\n🎬 ชมรมสื่อสร้างสรรค์ Media EDU Club CMU.\n━━━━━━━━━━━━━━━\nกรุณาลงทะเบียนเพื่อรับการแจ้งเตือนผ่าน LINE\nใช้เวลาเพียง 1 นาที!' },
  'msg_default_menu':    { label:'เมนูหลัก (default)',      group:'ข้อความ bot', hint:'แสดงเมื่อพิมพ์ข้อความที่ไม่รู้จัก', val:'📌 คำสั่งที่ใช้ได้:\n\n🔍 ติดตามสถานะ\n💻 ยืมอุปกรณ์\n🏫 จองห้อง\n📞 ติดต่อ\n📋 ข้อมูลของฉัน\n📝 แก้ไขข้อมูล' },
  'msg_contact':         { label:'ข้อความติดต่อ',           group:'ข้อความ bot', hint:'แสดงเมื่อพิมพ์ "ติดต่อ"',            val:'📞 ช่องทางติดต่อ\n━━━━━━━━━━━━━━━\n📘 Facebook: Media EDU Club CMU.\nhttps://www.facebook.com/media.ed.cmu\n\n🕐 เวลาทำการ:\nจันทร์ – ศุกร์  08:30–16:30 น.\n\n📍 ที่ตั้ง: คณะศึกษาศาสตร์ มหาวิทยาลัยเชียงใหม่' },
};

/** โหลด templates ทั้งหมด (cache 1 ครั้งต่อการ call) */
var _tplCache = null;
function loadTemplateCache() {
  if (_tplCache) return _tplCache;
  var ss    = getDB();
  var sheet = ss.getSheetByName(LINE_TEMPLATES_SHEET);
  _tplCache = {};
  if (!sheet) return _tplCache;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][0] || '').trim();
    var v = String(data[i][1] || '');
    if (k) _tplCache[k] = v;
  }
  return _tplCache;
}

/** ดึง template พร้อม fallback default */
function getTpl(key) {
  var cache = loadTemplateCache();
  if (cache[key] !== undefined && cache[key] !== '') return cache[key];
  var def = LINE_TEMPLATE_DEFAULTS[key];
  return def ? def.val : key;
}

/** แทนค่า variables ใน template  เช่น {name}, {status} */
function fillTpl(key, vars) {
  var text = getTpl(key);
  if (vars) {
    Object.keys(vars).forEach(function(k) {
      text = text.replace(new RegExp('\{' + k + '\}', 'g'), vars[k] || '—');
    });
  }
  return text;
}

/** อ่าน templates ทั้งหมด (สำหรับ Admin UI) */
function getLineTemplates() {
  var ss    = getDB();
  var sheet = ss.getSheetByName(LINE_TEMPLATES_SHEET);
  var saved = {};
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var k = String(data[i][0] || '').trim();
      if (k) saved[k] = String(data[i][1] || '');
    }
  }

  // รวม defaults กับ saved values
  var result = [];
  var groups = {};
  Object.keys(LINE_TEMPLATE_DEFAULTS).forEach(function(key) {
    var def = LINE_TEMPLATE_DEFAULTS[key];
    var row = {
      key:     key,
      label:   def.label,
      group:   def.group,
      hint:    def.hint,
      default: def.val,
      value:   saved[key] !== undefined ? saved[key] : def.val
    };
    if (!groups[def.group]) groups[def.group] = [];
    groups[def.group].push(row);
  });

  // จัดเป็น array ตาม group
  Object.keys(groups).forEach(function(g) {
    result.push({ group: g, items: groups[g] });
  });
  return result;
}

/** บันทึก templates (Admin save) */
function saveLineTemplates(updates) {
  // updates = [{key, value}, ...]
  _tplCache = null; // clear cache
  var ss    = getDB();
  var sheet = ss.getSheetByName(LINE_TEMPLATES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LINE_TEMPLATES_SHEET);
    sheet.appendRow(['key', 'value', 'description', 'วันที่อัปเดต']);
    sheet.getRange(1,1,1,4).setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold');
    sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 400); sheet.setColumnWidth(3, 250); sheet.setColumnWidth(4, 160);
  }

  var data = sheet.getDataRange().getValues();
  var keyRowMap = {};
  for (var i = 1; i < data.length; i++) {
    keyRowMap[String(data[i][0]).trim()] = i + 1;
  }

  var timestamp = getThaiTimestamp();
  updates.forEach(function(u) {
    var key = String(u.key || '').trim();
    if (!key) return;
    var def = LINE_TEMPLATE_DEFAULTS[key] || {};
    if (keyRowMap[key]) {
      sheet.getRange(keyRowMap[key], 2).setValue(u.value);
      sheet.getRange(keyRowMap[key], 4).setValue(timestamp);
    } else {
      sheet.appendRow([key, u.value, def.label || '', timestamp]);
    }
  });

  return { ok: true, updated: updates.length };
}

/** รีเซ็ต template กลับค่า default */
function resetLineTemplate(key) {
  _tplCache = null;
  var def = LINE_TEMPLATE_DEFAULTS[key];
  if (!def) return { ok: false, msg: 'ไม่พบ key: ' + key };
  return saveLineTemplates([{ key: key, value: def.val }]);
}

/** รีเซ็ตทุก template กลับค่า default */
function resetAllLineTemplates() {
  _tplCache = null;
  var ss    = getDB();
  var sheet = ss.getSheetByName(LINE_TEMPLATES_SHEET);
  if (sheet) ss.deleteSheet(sheet);
  return { ok: true };
}


// ============================================================
//  ปฏิทินกิจกรรมสโมสร — Activity Events
//  Sheet "กิจกรรม_สโมสร"
//  Columns: ชื่อกิจกรรม | วันที่ | สถานที่ | รายละเอียด | ประเภท(activity/camp/academic) | สถานะ(แสดง/ซ่อน)
// ============================================================

var ACTIVITY_EVENTS_SHEET = 'กิจกรรม_สโมสร';

function getActivityEvents() {
  try {
    var ss    = getDB();
    var sheet = ss.getSheetByName(ACTIVITY_EVENTS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { events: [] };

    var rows = sheet.getDataRange().getValues().slice(1);
    var events = rows
      .filter(function(r) { return r[0] && r[5] !== 'ซ่อน'; })
      .map(function(r) {
        var d = r[1] ? new Date(r[1]) : null;
        return {
          title:    String(r[0] || ''),
          date:     d ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd') : '',
          location: String(r[2] || ''),
          desc:     String(r[3] || ''),
          type:     String(r[4] || 'activity'),
        };
      })
      .filter(function(e) { return e.date; });

    return { events: events };
  } catch(e) {
    Logger.log('getActivityEvents error: ' + e);
    return { events: [] };
  }
}

function setupActivityEventsSheet() {
  var ss    = getDB();
  var sheet = ss.getSheetByName(ACTIVITY_EVENTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ACTIVITY_EVENTS_SHEET);
    var headers = ['ชื่อกิจกรรม','วันที่','สถานที่','รายละเอียด',
                   'ประเภท (activity/camp/academic)','สถานะ (แสดง/ซ่อน)'];
    sheet.appendRow(headers);
    var hr = sheet.getRange(1,1,1,headers.length);
    hr.setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,220); sheet.setColumnWidth(2,120);
    sheet.setColumnWidth(3,160); sheet.setColumnWidth(4,220);
    sheet.setColumnWidth(5,160); sheet.setColumnWidth(6,120);

    // ข้อมูลตัวอย่าง
    var y = new Date().getFullYear();
    var samples = [
      ['ค่ายอาสาพัฒนาชุมชน ครั้งที่ 15', new Date(y, 4, 20), 'จังหวัดน่าน', 'ค่ายอาสาพัฒนาโรงเรียนในชนบท', 'camp', 'แสดง'],
      ['กีฬาสีคณะศึกษาศาสตร์',            new Date(y, 5, 10), 'สนามกีฬามหาวิทยาลัย', '', 'activity', 'แสดง'],
      ['สัมมนาครูยุคใหม่ใจดิจิทัล',        new Date(y, 5, 25), 'ห้องประชุมใหญ่', 'นวัตกรรมการสอน', 'academic', 'แสดง'],
      ['วันไหว้ครู',                         new Date(y, 5, 12), 'อาคารพลศึกษา', '', 'activity', 'แสดง'],
      ['ค่ายน้องใหม่ ED',                    new Date(y, 6,  5), 'ค่ายลูกเสือ', 'ปฐมนิเทศนักศึกษาใหม่', 'camp', 'แสดง'],
    ];
    samples.forEach(function(r) { sheet.appendRow(r); });
  }
  return 'สร้าง Sheet ' + ACTIVITY_EVENTS_SHEET + ' สำเร็จ';
}

// ============================================================
//  ลงทะเบียนชุมนุม — Club Registration
//  Sheet "ชุมนุม"  Columns: ชื่อชุมนุม | ประเภท | ไอคอน | จำนวน | จำนวนสูงสุด
//  Sheet "ลงทะเบียนชุมนุม"  Columns: รหัสนักศึกษา | สาขา | ชุมนุม | วันที่ | สถานะ
// ============================================================

// ============================================================
// ACTIVITY REGISTRATION SYSTEM
// ============================================================
var ACTIVITIES_SHEET = 'กิจกรรม';

// สร้างชื่อ Sheet จากชื่อกิจกรรม (ไม่เกิน 30 ตัว ปลอดภัยสำหรับ GSheet)
function _actSheetName(name) {
  var clean = name.replace(/[\/:*?\[\]]/g, '').replace(/\s+/g, '_').substring(0, 28);
  return 'REG_' + clean;
}

// สร้าง/get Sheet ลงทะเบียนของกิจกรรม
function _ensureActRegSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var h = ['รหัสนักศึกษา','ชื่อ-นามสกุล','ชื่อเล่น','ชั้นปี','เบอร์โทร','สาขาวิชา','คณะ',
             'ประเภทอาหาร','อาหารที่แพ้','ยาที่แพ้','โรคประจำตัว','ยาประจำตัว',
             'วันที่ลงทะเบียน','สถานะ'];
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length)
         .setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, h.length, 140);
  }
  return sheet;
}

// แปลงวันที่ (YYYY-MM-DD) + เวลา (HH:MM) เป็นข้อความไทย
function _formatThaiDateTime(dateStr, timeStr) {
  if (!dateStr) return '';
  try {
    var parts = dateStr.split('-');
    var thaiMonths = ['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
    var d = parseInt(parts[2]), m = parseInt(parts[1]), y = parseInt(parts[0]) + 543;
    var result = d + ' ' + thaiMonths[m] + ' ' + y;
    if (timeStr) result += ' เวลา ' + timeStr + ' น.';
    return result;
  } catch(e) { return dateStr + (timeStr ? ' ' + timeStr : ''); }
}

/**
 * ดึงกิจกรรมสำหรับนักศึกษา (เฉพาะที่เปิดและยังไม่หมดเวลา)
 */
function getActivitiesList() {
  try {
    var ss    = getDB();
    var sheet = ss.getSheetByName(ACTIVITIES_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { activities: [] };
    var now   = new Date();
    var rows  = sheet.getDataRange().getValues().slice(1);
    var activities = rows
      .filter(function(r) { return r[0]; })
      .map(function(r, i) {
        var sheetName = String(r[14] || '');
        var cnt = sheetName ? _getRegCount(ss, sheetName) : (Number(r[12]) || 0);
        var max = Number(r[13]) || 0;
        // ตรวจวันปิดรับสมัคร
        var isOpen = String(r[9] || 'TRUE').toUpperCase() !== 'FALSE';
        if (isOpen && r[10]) {
          try {
            var deadline = new Date(String(r[10]) + 'T' + (String(r[11]||'23:59')) + ':00');
            if (!isNaN(deadline.getTime()) && now > deadline) isOpen = false;
          } catch(e) {}
        }
        return {
          id:           i + 1,
          name:         String(r[0] || ''),
          type:         String(r[1] || 'general'),
          icon:         String(r[2] || '🎉'),
          desc:         String(r[3] || ''),
          dateStart:    String(r[4] || ''),
          timeStart:    String(r[5] || ''),
          dateEnd:      String(r[6] || ''),
          timeEnd:      String(r[7] || ''),
          location:     String(r[8] || ''),
          deadlineDate: String(r[10] || ''),
          deadlineTime: String(r[11] || ''),
          count:        cnt,
          maxCount:     max,
          isOpen:       isOpen,
          sheetName:    sheetName,
          // แสดงผลวันที่สวยงาม
          dateStr:      _buildDateStr(r[4], r[5], r[6], r[7]),
          deadline:     _formatThaiDateTime(String(r[10]||''), String(r[11]||'')),
        };
      });
    return { activities: activities };
  } catch(e) {
    Logger.log('getActivitiesList error: ' + e);
    return { activities: [] };
  }
}

function _buildDateStr(ds, ts, de, te) {
  if (!ds) return '';
  var start = _formatThaiDateTime(String(ds||''), String(ts||''));
  if (de && de !== ds) {
    var end = _formatThaiDateTime(String(de||''), String(te||''));
    return start + ' — ' + end;
  }
  return start;
}

function _getRegCount(ss, sheetName) {
  try {
    var s = ss.getSheetByName(sheetName);
    return s ? Math.max(0, s.getLastRow() - 1) : 0;
  } catch(e) { return 0; }
}

/**
 * ดึงกิจกรรมทั้งหมดสำหรับ Admin (รวมที่ปิด)
 */
function getActivitiesListAdmin() {
  try {
    var ss    = getDB();
    var sheet = ss.getSheetByName(ACTIVITIES_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { activities: [] };
    var rows  = sheet.getDataRange().getValues().slice(1);
    var activities = rows
      .filter(function(r) { return r[0]; })
      .map(function(r, i) {
        var sheetName = String(r[14] || '');
        var cnt = sheetName ? _getRegCount(ss, sheetName) : (Number(r[12]) || 0);
        return {
          id:           i + 1,
          name:         String(r[0] || ''),
          type:         String(r[1] || 'general'),
          icon:         String(r[2] || '🎉'),
          desc:         String(r[3] || ''),
          dateStart:    String(r[4] || ''),
          timeStart:    String(r[5] || ''),
          dateEnd:      String(r[6] || ''),
          timeEnd:      String(r[7] || ''),
          location:     String(r[8] || ''),
          deadlineDate: String(r[10] || ''),
          deadlineTime: String(r[11] || ''),
          count:        cnt,
          maxCount:     Number(r[13]) || 0,
          isOpen:       String(r[9] || 'TRUE').toUpperCase() !== 'FALSE',
          sheetName:    sheetName,
          dateStr:      _buildDateStr(r[4], r[5], r[6], r[7]),
        };
      });
    return { activities: activities };
  } catch(e) { return { activities: [] }; }
}

/**
 * เพิ่ม / แก้ไขกิจกรรม + สร้าง Sheet ลงทะเบียน
 * Columns: A=ชื่อ B=ประเภท C=ไอคอน D=รายละเอียด
 *          E=dateStart F=timeStart G=dateEnd H=timeEnd
 *          I=สถานที่ J=isOpen K=deadlineDate L=deadlineTime
 *          M=count(legacy) N=maxCount O=sheetName
 */
function saveActivity(data) {
  try {
    var ss    = getDB();
    var sheet = ss.getSheetByName(ACTIVITIES_SHEET);
    if (!sheet) { setupActivitiesSheet(); sheet = ss.getSheetByName(ACTIVITIES_SHEET); }

    var name      = String(data.name || '').trim();
    var rowIndex  = Number(data.rowIndex) || 0;
    var sheetName = String(data.sheetName || '').trim();

    // สร้าง sheetName ใหม่ถ้ายังไม่มี
    if (!sheetName) sheetName = _actSheetName(name);

    // สร้าง Sheet ลงทะเบียนถ้ายังไม่มี
    _ensureActRegSheet(ss, sheetName);

    var rowData = [
      name,                                          // A ชื่อกิจกรรม
      String(data.type     || 'general'),            // B ประเภท
      String(data.icon     || '🎉'),                 // C ไอคอน
      String(data.desc     || ''),                   // D รายละเอียด
      String(data.dateStart    || ''),               // E วันที่เริ่ม
      String(data.timeStart    || ''),               // F เวลาเริ่ม
      String(data.dateEnd      || ''),               // G วันที่สิ้นสุด
      String(data.timeEnd      || ''),               // H เวลาสิ้นสุด
      String(data.location     || ''),               // I สถานที่
      data.isOpen !== false ? 'TRUE' : 'FALSE',      // J เปิดรับสมัคร
      String(data.deadlineDate || ''),               // K วันปิดรับสมัคร
      String(data.deadlineTime || ''),               // L เวลาปิดรับสมัคร
      0,                                             // M count (legacy)
      Number(data.maxCount) || 0,                    // N จำนวนสูงสุด
      sheetName,                                     // O ชื่อ Sheet
    ];

    if (rowIndex > 0) {
      sheet.getRange(rowIndex + 1, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return { ok: true, sheetName: sheetName };
  } catch(e) {
    Logger.log('saveActivity error: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

/**
 * ลบกิจกรรม (Admin) — ลบแถวจาก Master Sheet (ไม่ลบ Sheet ย่อย)
 */
function deleteActivity(actId) {
  try {
    var ss    = getDB();
    var sheet = ss.getSheetByName(ACTIVITIES_SHEET);
    if (!sheet) return { ok: false, msg: 'ไม่พบ Sheet' };
    sheet.deleteRow(Number(actId) + 1);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

/**
 * บันทึกการลงทะเบียนกิจกรรม → บันทึกลง Sheet ของกิจกรรมนั้น
 */
function registerActivity(data) {
  try {
    var actId  = Number(data.activityId);
    var sid    = String(data.sid   || '').trim();
    var fname  = String(data.fname || '').trim();
    var lname  = String(data.lname || '').trim();
    if (!actId || !sid || !fname || !lname) return { ok: false, msg: 'ข้อมูลไม่ครบ' };

    var ss       = getDB();
    var actSheet = ss.getSheetByName(ACTIVITIES_SHEET);
    if (!actSheet) return { ok: false, msg: 'ไม่พบ Sheet กิจกรรม' };

    var actRows = actSheet.getDataRange().getValues().slice(1);
    var actRow  = actRows[actId - 1];
    if (!actRow) return { ok: false, msg: 'ไม่พบกิจกรรม ID ' + actId };
    if (String(actRow[9] || 'TRUE').toUpperCase() === 'FALSE') return { ok: false, msg: 'กิจกรรมนี้ปิดรับสมัครแล้ว' };

    // ตรวจ deadline
    if (actRow[10]) {
      try {
        var dl = new Date(String(actRow[10]) + 'T' + (String(actRow[11]||'23:59')) + ':00');
        if (!isNaN(dl.getTime()) && new Date() > dl) return { ok: false, msg: 'หมดเวลารับสมัครแล้ว' };
      } catch(e) {}
    }

    // ดึง Sheet ย่อยของกิจกรรม
    var sheetName = String(actRow[14] || _actSheetName(String(actRow[0]||'')));
    var regSheet  = _ensureActRegSheet(ss, sheetName);

    // ตรวจสอบ max
    var cnt = Math.max(0, regSheet.getLastRow() - 1);
    var max = Number(actRow[13]) || 0;
    if (max > 0 && cnt >= max) return { ok: false, msg: 'กิจกรรมนี้รับสมัครครบแล้ว (' + cnt + '/' + max + ' คน)' };

    // ตรวจซ้ำ
    if (regSheet.getLastRow() > 1) {
      var existing = regSheet.getDataRange().getValues().slice(1)
        .find(function(r) { return String(r[0]).trim() === sid; });
      if (existing) return { ok: false, msg: 'รหัสนักศึกษา ' + sid + ' ลงทะเบียนกิจกรรมนี้แล้ว' };
    }

    var actName   = String(actRow[0] || '');
    var foodLabel = { normal:'ปกติ', veg:'เจ/มังสวิรัติ', halal:'อิสลาม (ฮาลาล)' };

    regSheet.appendRow([
      sid,
      fname + ' ' + lname,
      String(data.nickname    || ''),
      String(data.year        || ''),
      String(data.phone       || ''),
      String(data.major       || ''),
      String(data.faculty     || ''),
      foodLabel[data.foodType] || String(data.foodType || 'ปกติ'),
      String(data.foodAllergy || '-'),
      String(data.drugAllergy || '-'),
      String(data.disease     || '-'),
      String(data.medicine    || '-'),
      getThaiTimestamp(),
      'รอยืนยัน',
    ]);

    // แจ้ง LINE Admin
    try {
      notifyAdminsByScope(
        '📋 ลงทะเบียนกิจกรรมใหม่\n' +
        '━━━━━━━━━━━━━━━\n' +
        '🎉 ' + actName + '\n' +
        '👤 ' + fname + ' ' + lname + ' (' + sid + ')\n' +
        '📱 ' + String(data.phone||'') + '\n' +
        '📅 ' + getThaiTimestamp(), 'samo'
      );
    } catch(le) {}

    return { ok: true, msg: 'ลงทะเบียนสำเร็จ' };
  } catch(e) {
    Logger.log('registerActivity error: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

/**
 * ดึงรายชื่อผู้สมัครของกิจกรรม (Admin)
 */
function getActivityRegistrations(actId) {
  try {
    var ss       = getDB();
    var actSheet = ss.getSheetByName(ACTIVITIES_SHEET);
    if (!actSheet || actSheet.getLastRow() < 2) return [];
    var actRow   = actSheet.getDataRange().getValues().slice(1)[Number(actId) - 1];
    if (!actRow) return [];
    var sheetName = String(actRow[14] || '');
    if (!sheetName) return [];
    var regSheet = ss.getSheetByName(sheetName);
    if (!regSheet || regSheet.getLastRow() < 2) return [];
    // Return rows compatible with old format (add actId at index 1 for JS rendering)
    return regSheet.getDataRange().getValues().slice(1).map(function(r) {
      return [r[0], actId, String(actRow[0]||''), r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[13]];
    });
  } catch(e) { return []; }
}

/**
 * Setup Sheet กิจกรรม
 */
function setupActivitiesSheet() {
  var ss    = getDB();
  var sheet = ss.getSheetByName(ACTIVITIES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ACTIVITIES_SHEET);
    var h = ['ชื่อกิจกรรม','ประเภท','ไอคอน','รายละเอียด',
             'วันที่เริ่ม','เวลาเริ่ม','วันที่สิ้นสุด','เวลาสิ้นสุด',
             'สถานที่','เปิดรับสมัคร',
             'วันปิดรับสมัคร','เวลาปิดรับสมัคร',
             'จำนวน(legacy)','จำนวนรับสูงสุด','ชื่อ Sheet'];
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length)
         .setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, h.length, 140);
    return 'สร้าง Sheet กิจกรรมสำเร็จ';
  }
  return 'Sheet มีอยู่แล้ว';
}

// เก็บ functions เก่าไว้เพื่อ backward-compat (ชุมนุม/กีฬา)
var CLUBS_SHEET     = 'ชุมนุม';
var CLUB_REG_SHEET  = 'ลงทะเบียนชุมนุม';


function getClubsList() {
  try {
    var ss    = getDB();
    var sheet = ss.getSheetByName(CLUBS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { clubs: [] };

    var clubs = sheet.getDataRange().getValues().slice(1)
      .filter(function(r) { return r[0]; })
      .map(function(r, i) {
        var cnt = Number(r[3]) || 0, max = Number(r[4]) || 0;
        return {
          id: i + 1,
          name:     String(r[0] || ''),
          type:     String(r[1] || 'ชุมนุม'),
          icon:     String(r[2] || '🎯'),
          count:    cnt,
          maxCount: max,
          full:     max > 0 && cnt >= max,
        };
      });
    return { clubs: clubs };
  } catch(e) {
    Logger.log('getClubsList error: ' + e);
    return { clubs: [] };
  }
}

function getStudentClubInfo(stdid) {
  try {
    var ss    = getDB();
    var sheet = ss.getSheetByName(CLUB_REG_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { currentClub: '' };
    var found = sheet.getDataRange().getValues().slice(1)
                  .find(function(r) { return String(r[0]).trim() === String(stdid).trim(); });
    return { currentClub: found ? String(found[2] || '') : '' };
  } catch(e) {
    return { currentClub: '' };
  }
}

function registerClub(data) {
  try {
    var stdid  = String(data.stdid  || '').trim();
    var clubId = Number(data.clubId);
    if (!stdid || !clubId) return { ok: false, msg: 'ข้อมูลไม่ครบ' };

    var ss        = getDB();
    var clubSheet = ss.getSheetByName(CLUBS_SHEET);
    if (!clubSheet) return { ok: false, msg: 'ไม่พบ Sheet ชุมนุม' };

    var clubRows  = clubSheet.getDataRange().getValues().slice(1);
    var clubRow   = clubRows[clubId - 1];
    if (!clubRow) return { ok: false, msg: 'ไม่พบชุมนุม ID ' + clubId };

    var clubName = String(clubRow[0] || '');
    var cnt      = Number(clubRow[3]) || 0;
    var max      = Number(clubRow[4]) || 0;
    if (max > 0 && cnt >= max) return { ok: false, msg: 'ชุมนุมนี้เต็มแล้ว (' + cnt + '/' + max + ')' };

    // ตรวจสอบลงทะเบียนซ้ำ
    var regSheet = _ensureClubRegSheet(ss);
    var existing = regSheet.getDataRange().getValues().slice(1)
                    .find(function(r) { return String(r[0]).trim() === stdid; });
    if (existing) return { ok: false, msg: 'รหัสนักศึกษา ' + stdid + ' ลงทะเบียนชุมนุม "' + String(existing[2]) + '" แล้ว' };

    // บันทึก
    regSheet.appendRow([stdid, data.major || '', clubName, new Date(), 'รอยืนยัน']);

    // อัปเดตจำนวน
    clubSheet.getRange(clubId + 1, 4).setValue(cnt + 1);

    // แจ้ง LINE Admin (ถ้ามี)
    try {
      notifyAdminsByScope(
        '🎽 ลงทะเบียนชุมนุมใหม่\n' +
        '━━━━━━━━━━━━━━━\n' +
        '👤 รหัส: ' + stdid + '\n' +
        '🏫 ชุมนุม: ' + clubName + '\n' +
        '📅 ' + getThaiTimestamp(),
        'samo'
      );
    } catch(le) {}

    return { ok: true, msg: 'ลงทะเบียนชุมนุม "' + clubName + '" สำเร็จ' };
  } catch(e) {
    Logger.log('registerClub error: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

function _ensureClubRegSheet(ss) {
  var sheet = ss.getSheetByName(CLUB_REG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CLUB_REG_SHEET);
    var h = ['รหัสนักศึกษา','สาขา','ชุมนุม','วันที่ลงทะเบียน','สถานะ'];
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length)
         .setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function setupClubsSheet() {
  var ss    = getDB();
  var sheet = ss.getSheetByName(CLUBS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CLUBS_SHEET);
    var h = ['ชื่อชุมนุม/กีฬา','ประเภท (ชุมนุม/กีฬา)','ไอคอน (Emoji)','จำนวนสมาชิก','จำนวนสูงสุด'];
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length)
         .setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
    var samples = [
      ['ชุมนุมคอมพิวเตอร์','ชุมนุม','💻',0,60],
      ['ชุมนุมดนตรี','ชุมนุม','🎵',0,40],
      ['กีฬาฟุตบอล','กีฬา','⚽',0,30],
      ['กีฬาบาสเกตบอล','กีฬา','🏀',0,24],
      ['ชุมนุมศิลปะ','ชุมนุม','🎨',0,35],
      ['กีฬาวอลเลย์บอล','กีฬา','🏐',0,24],
      ['ชุมนุมถ่ายภาพ','ชุมนุม','📷',0,20],
      ['ชุมนุมอาสาพัฒนา','ชุมนุม','🌱',0,50],
      ['สื่อสร้างสรรค์ Media EDU','ชุมนุม','🎬',0,40],
    ];
    samples.forEach(function(r) { sheet.appendRow(r); });
  }
  _ensureClubRegSheet(ss);
  return 'สร้าง Sheet ชุมนุม และ ลงทะเบียนชุมนุม สำเร็จ';
}

// เรียกรวดเดียวเพื่อ setup ทุก Sheet ใหม่
function setupAllNewSheets() {
  Logger.log(setupActivityEventsSheet());
  Logger.log(setupClubsSheet());
  return 'Setup ครบทุก Sheet แล้ว ✅';
}// ============================================================
//  📌 วิธีใช้ไฟล์นี้:
//  เพิ่มโค้ดทั้งหมดด้านล่างนี้ ต่อท้ายไฟล์ Code.gs เดิมของคุณ
//  (วางหลัง function setupAllNewSheets() ที่บรรทัดสุดท้าย)
// ============================================================

// ============================================================
//  ทำเนียบองค์กร — Org Directory
//  Sheet "ทำเนียบองค์กร"
//  Columns:
//    A = ID        B = ชื่อ-สกุล    C = ตำแหน่ง      D = ฝ่าย
//    E = URL รูปภาพ  F = เบอร์โทร   G = อีเมล         H = ห้องทำงาน
//    I = LINE ID    J = ลำดับ       K = สถานะ (แสดง/ซ่อน)
// ============================================================

var ORG_DIRECTORY_SHEET = 'ทำเนียบองค์กร';

/**
 * ดึงข้อมูลสมาชิกทำเนียบองค์กร (สำหรับหน้าสาธารณะ)
 * กรองเฉพาะที่สถานะ = "แสดง" และจัดกลุ่มตามฝ่าย
 */
/**
 * Helper: สร้าง map จากชื่อ header → index column (0-based)
 * อ่านจาก row แรกของ sheet เพื่อรองรับ sheet ที่มี column ต่างกัน
 */
function _orgBuildColMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  headers.forEach(function(h, i) { if (h) map[String(h).trim()] = i; });
  return map;
}

/** Helper: อ่านค่าจาก row ด้วยชื่อ header */
function _orgGet(row, map, colName, def) {
  var i = map[colName];
  return (i !== undefined && row[i] !== undefined && row[i] !== null)
    ? String(row[i]).trim()
    : (def || '');
}

function getOrgDirectory() {
  try {
    var ss    = getDB();
    var sheet = ss.getSheetByName(ORG_DIRECTORY_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { departments: [], members: [] };

    var allData = sheet.getDataRange().getValues();
    var map     = _orgBuildColMap(sheet);
    var rows    = allData.slice(1);

    var members = rows
      .filter(function(r) {
        return _orgGet(r, map, 'ชื่อ') && _orgGet(r, map, 'สถานะ', 'แสดง') !== 'ซ่อน';
      })
      .map(function(r) {
        var g = function(col, def) { return _orgGet(r, map, col, def); };
        var firstName = g('ชื่อ');
        var lastName  = g('นามสกุล');
        var fullName  = lastName ? firstName + ' ' + lastName : firstName;
        return {
          id:       g('ID'),
          name:     fullName,
          major:    g('สาขาวิชา'),
          position: g('ตำแหน่ง'),
          dept:     g('ฝ่าย') || 'ไม่ระบุฝ่าย',
          photoUrl: g('URL รูปภาพ'),
          phone:    g('เบอร์โทร'),
          email:    g('อีเมล'),
          office:   g('ห้องทำงาน'),
          lineId:   g('LINE ID'),
          order:    Number(g('ลำดับ')) || 99,
        };
      })
      .sort(function(a, b) {
        if (a.dept < b.dept) return -1;
        if (a.dept > b.dept) return 1;
        return a.order - b.order;
      });

    var deptSet = [];
    members.forEach(function(m) {
      if (deptSet.indexOf(m.dept) === -1) deptSet.push(m.dept);
    });

    return { departments: deptSet, members: members };
  } catch(e) {
    Logger.log('getOrgDirectory error: ' + e);
    return { departments: [], members: [] };
  }
}

/**
 * ดึงข้อมูลสมาชิกทั้งหมด (สำหรับหน้า Admin รวมที่ซ่อน)
 */
function getOrgDirectoryAdmin() {
  try {
    var ss    = getDB();
    var sheet = ss.getSheetByName(ORG_DIRECTORY_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { members: [] };

    var allData = sheet.getDataRange().getValues();
    var map     = _orgBuildColMap(sheet);
    var rows    = allData.slice(1);

    var members = rows
      .filter(function(r) { return _orgGet(r, map, 'ชื่อ'); })
      .map(function(r, i) {
        var g = function(col, def) { return _orgGet(r, map, col, def); };
        var firstName = g('ชื่อ');
        var lastName  = g('นามสกุล');
        var fullName  = lastName ? firstName + ' ' + lastName : firstName;
        return {
          rowIdx:    i + 2,
          id:        g('ID'),
          firstName: firstName,
          lastName:  lastName,
          name:      fullName,
          major:     g('สาขาวิชา'),
          position:  g('ตำแหน่ง'),
          dept:      g('ฝ่าย'),
          photoUrl:  g('URL รูปภาพ'),
          phone:     g('เบอร์โทร'),
          email:     g('อีเมล'),
          office:    g('ห้องทำงาน'),
          lineId:    g('LINE ID'),
          order:     Number(g('ลำดับ')) || 99,
          status:    g('สถานะ', 'แสดง'),
        };
      });

    return { members: members };
  } catch(e) {
    Logger.log('getOrgDirectoryAdmin error: ' + e);
    return { members: [] };
  }
}

/**
 * เพิ่ม / แก้ไขสมาชิก — เขียนตามชื่อ header ของ sheet จริง
 */
function saveOrgMember(data) {
  try {
    if (!data.firstName && !data.name)
      return { ok: false, msg: 'กรุณากรอกชื่อให้ครบ' };
    if (!data.dept)
      return { ok: false, msg: 'กรุณาระบุฝ่าย' };

    var ss    = getDB();
    var sheet = _ensureOrgSheet(ss);
    var map   = _orgBuildColMap(sheet);
    var numCols = sheet.getLastColumn();

    // สร้าง array ขนาดเท่ากับ header ปัจจุบัน
    var row = new Array(numCols).fill('');
    function set(col, val) {
      var i = map[col];
      if (i !== undefined) row[i] = (val !== null && val !== undefined) ? val : '';
    }

    set('ID',           data.id        || '');
    set('ชื่อ',          data.firstName || data.name || '');
    set('นามสกุล',       data.lastName  || '');
    set('สาขาวิชา',      data.major     || '');
    set('ตำแหน่ง',       data.position  || '');
    set('ฝ่าย',          data.dept      || '');
    set('URL รูปภาพ',   data.photoUrl  || '');
    set('เบอร์โทร',     data.phone     || '');
    set('อีเมล',         data.email     || '');
    set('ห้องทำงาน',    data.office    || '');
    set('LINE ID',      data.lineId    || '');
    set('ลำดับ',         Number(data.order) || 99);
    set('สถานะ',         data.status    || 'แสดง');

    var rowIdx = Number(data.rowIdx) || 0;
    if (rowIdx >= 2) {
      sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
      return { ok: true, msg: 'แก้ไขข้อมูลสำเร็จ' };
    } else {
      var idIdx = map['ID'];
      if (idIdx !== undefined && !row[idIdx]) row[idIdx] = 'ORG' + new Date().getTime();
      sheet.appendRow(row);
      return { ok: true, msg: 'เพิ่มสมาชิกสำเร็จ' };
    }
  } catch(e) {
    Logger.log('saveOrgMember error: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

/**
 * ลบสมาชิกตาม rowIdx
 */
function deleteOrgMember(rowIdx) {
  try {
    rowIdx = Number(rowIdx);
    if (!rowIdx || rowIdx < 2) return { ok: false, msg: 'rowIdx ไม่ถูกต้อง' };
    var ss    = getDB();
    var sheet = ss.getSheetByName(ORG_DIRECTORY_SHEET);
    if (!sheet) return { ok: false, msg: 'ไม่พบ Sheet ทำเนียบองค์กร' };
    sheet.deleteRow(rowIdx);
    return { ok: true, msg: 'ลบสมาชิกสำเร็จ' };
  } catch(e) {
    Logger.log('deleteOrgMember error: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

/**
 * Setup Sheet ทำเนียบองค์กร พร้อมข้อมูลตัวอย่าง
 */
function setupOrgDirectorySheet() {
  var ss    = getDB();
  var sheet = _ensureOrgSheet(ss);
  if (sheet.getLastRow() > 1) return 'Sheet ทำเนียบองค์กร มีอยู่แล้ว ✅';

  // [ID, ชื่อ, นามสกุล, สาขาวิชา, ตำแหน่ง, ฝ่าย, URL รูปภาพ, เบอร์โทร, อีเมล, ห้องทำงาน, LINE ID, ลำดับ, สถานะ]
  var samples = [
    ['ORG001','สมชาย',    'ใจดี',      '',           'นายกสโมสร',           'คณะกรรมการบริหาร',    '','081-000-0001','somchai@cmu.ac.th',   'อาคาร A ห้อง 101','',1,'แสดง'],
    ['ORG002','สมหญิง',   'รักเรียน',  '',           'อุปนายกฝ่ายวิชาการ', 'คณะกรรมการบริหาร',    '','081-000-0002','somying@cmu.ac.th',   'อาคาร A ห้อง 101','',2,'แสดง'],
    ['ORG003','วิชาญ',    'เก่งกาจ',   '',           'อุปนายกฝ่ายกิจการ',  'คณะกรรมการบริหาร',    '','081-000-0003','wichan@cmu.ac.th',    'อาคาร A ห้อง 101','',3,'แสดง'],
    ['ORG004','มาลี',     'ดอกไม้',    'คณิตศาสตร์', 'อาจารย์ที่ปรึกษา',  'อาจารย์ที่ปรึกษา',    '','081-000-0004','malee@cmu.ac.th',     'อาคาร B ห้อง 201','',1,'แสดง'],
    ['ORG005','ธนา',      'มั่งมี',    '',           'กรรมการ',             'ฝ่ายสวัสดิการ',       '','081-000-0005','thana@cmu.ac.th',     'อาคาร B ห้อง 201','',1,'แสดง'],
    ['ORG006','กนกวรรณ',  'พิมลรัตน์', '',           'กรรมการ',             'ฝ่ายสวัสดิการ',       '','081-000-0006','kanok@cmu.ac.th',     'อาคาร B ห้อง 201','',2,'แสดง'],
    ['ORG007','ธนพล',     'สดใส',      '',           'หัวหน้าฝ่ายสื่อสาร', 'ฝ่ายสื่อสาร',         '','081-000-0007','thanapol@cmu.ac.th',  'อาคาร C ห้อง 301','',1,'แสดง'],
    ['ORG008','จิราพร',   'มีสุข',     '',           'กรรมการ',             'ฝ่ายสื่อสาร',         '','081-000-0008','jiraporn@cmu.ac.th',  'อาคาร C ห้อง 301','',2,'แสดง'],
    ['ORG009','วิทยา',    'ใจงาม',     '',           'เจ้าหน้าที่',         'เจ้าหน้าที่',          '','081-000-0009','wittaya@cmu.ac.th',   'อาคาร A ห้อง 102','',1,'แสดง'],
    ['ORG010','อรุณ',     'แสงใส',     'ฟิสิกส์',   'ที่ปรึกษาพิเศษ',     'ที่ปรึกษาพิเศษ',       '','081-000-0010','arun@cmu.ac.th',      'อาคาร D ห้อง 401','',1,'แสดง'],
  ];
  samples.forEach(function(r) { sheet.appendRow(r); });
  return 'สร้าง Sheet ทำเนียบองค์กร พร้อมข้อมูลตัวอย่างสำเร็จ ✅';
}

/**
 * Helper: สร้าง Sheet ถ้ายังไม่มี (schema มาตรฐาน)
 */
function _ensureOrgSheet(ss) {
  var sheet = ss.getSheetByName(ORG_DIRECTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ORG_DIRECTORY_SHEET);
    var headers = ['ID','ชื่อ','นามสกุล','สาขาวิชา','ตำแหน่ง','ฝ่าย','URL รูปภาพ','เบอร์โทร','อีเมล','ห้องทำงาน','LINE ID','ลำดับ','สถานะ'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1565c0').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
    // กำหนดความกว้างคอลัมน์
    [110,140,140,160,180,160,240,130,200,200,120,80,100].forEach(function(w, i) {
      sheet.setColumnWidth(i + 1, w);
    });
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['แสดง','ซ่อน'], true).build();
    sheet.getRange(2, 13, 500, 1).setDataValidation(rule);
  }
  return sheet;
}

function setupAllSheetsIncludingOrg() {
  Logger.log(setupActivityEventsSheet());
  Logger.log(setupClubsSheet());
  Logger.log(setupOrgDirectorySheet());
  return 'Setup ครบทุก Sheet รวมทำเนียบองค์กร ✅';
}
