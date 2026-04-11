// ================================================================
// 組合單生成器 - Google Apps Script
// ================================================================

// SHEET_ID 和 API 密碼存在 PropertiesService，不寫在程式碼裡
// 第一次請執行 initProperties() 設定
function getSheetId() {
  return PropertiesService.getScriptProperties().getProperty('SHEET_ID');
}
function getApiPassword() {
  return PropertiesService.getScriptProperties().getProperty('API_PASSWORD');
}

// ── 讀取所有設定 ─────────────────────────────────────
function getSettings() {
  var ss = SpreadsheetApp.openById(getSheetId());

  // 機型設定
  var mmSheet = ss.getSheetByName('機型設定');
  var mmRows  = mmSheet.getDataRange().getValues();
  var MM = {};
  for (var i = 1; i < mmRows.length; i++) {
    var r = mmRows[i];
    if (!r[0]) continue;
    MM[r[0]] = {
      c14:      r[1] || '',
      c15:      r[2] || '',
      ribbonSz: r[3] || '',
      c21:      r[4] || '',
      c22:      r[5] || '',
      c24:      r[6] || '',
      c25:      r[7] || '',
      c27:      r[8] || '',
      c28:      r[9] || '',
    };
  }

  // 人員名單
  var staffSheet = ss.getSheetByName('人員名單');
  var staffRows  = staffSheet.getDataRange().getValues();
  var sales = [], engineers = [];
  for (var i = 1; i < staffRows.length; i++) {
    if (staffRows[i][0]) sales.push(staffRows[i][0]);
    if (staffRows[i][1]) engineers.push(staffRows[i][1]);
  }

  // 碳帶庫存
  var invSheet = ss.getSheetByName('碳帶庫存');
  var invRows  = invSheet.getDataRange().getValues();
  var INV = [];
  for (var i = 1; i < invRows.length; i++) {
    if (!invRows[i][0]) continue;
    INV.push({ '品名': invRows[i][0], '料號': invRows[i][1] });
  }

  return { MM: MM, sales: sales, engineers: engineers, INV: INV };
}

// ── 初始化設定（第一次或重設用）────────────────────────
function initSettings() {
  var ss = SpreadsheetApp.openById(getSheetId());

  // 機型設定
  var mmSheet = ss.getSheetByName('機型設定');
  mmSheet.clearContents();
  mmSheet.getRange(1,1,1,10).setValues([[
    '機型','C14品名','C15料號','碳帶尺寸',
    'C21配件1','C22料號','C24配件2','C25料號','C27配件3','C28料號'
  ]]);
  var VJ = [
    '編碼器 KOYO TRD-2T3600BF','QQENCODERKYO2T3600BF',
    '407968-06','QQTTOLANGUAGEVJ63530',
    'CLARiSUITE ADVANCED 軟體','QQTTOCLARISUITEADVAN'
  ];
  var MP = [
    '編碼器 32/64 位元軟體USB KEY','QQASOFTER32BITEUSB00',
    '手動機載物台','QQAAAOPSMPPANEL00000',
    '','',
  ];
  var AP = [
    '編碼器 32/64 位元軟體USB KEY','QQASOFTER32BITEUSB00',
    '幫浦220V (日本製) OFS-610 (EF500A)用','QQASPUMP220JAPAN0000',
    'OFS-610A (EF-500A半成品)','QQAOFS610ASEMI000000'
  ];
  var mmData = [
    ['6330 1inch LH','TTO 32MM VJ6330 LEFT HAND', 'QQTTO32VJ6330LH00000','34*600 OUT'].concat(VJ),
    ['6330 1inch RH','TTO 32MM VJ6330 RIGHT HAND','QQTTO32VJ6330RH00000','34*600 OUT'].concat(VJ),
    ['6330 2inch LH','TTO 53MM VJ6330 LEFT HAND', 'QQTTO53VJ6330LH00000','55*600 OUT'].concat(VJ),
    ['6330 2inch RH','TTO 53MM VJ6330 RIGHT HAND','QQTTO53VJ6330RH00000','55*600 OUT'].concat(VJ),
    ['6530 LH',      'TTO 53MM VJ6530 LEFT HAND', 'QQTTO53VJ6530LH00000','55*600 OUT'].concat(VJ),
    ['6530 RH',      'TTO 53MM VJ6530 RIGHT HAND','QQTTO53VJ6530RH00000','55*600 OUT'].concat(VJ),
    ['MP2',   'OPS-MP 2吋 手動機','QQAOPSMP200000000000','57*300 OUT'].concat(MP),
    ['MP3',   'OPS-MP 3吋 手動機','QQAOPSMP300000000000','82*300 OUT'].concat(MP),
    ['MP4',   'OPS-MP 4吋 手動機','QQAOPSMP400000000000','110*280 IN 置中'].concat(MP),
    ['610AP2','OPS-MP 2吋 手動機','QQAOPSMP200000000000','57*300 OUT'].concat(AP),
    ['610AP3','OPS-MP 3吋 手動機','QQAOPSMP300000000000','82*300 OUT'].concat(AP),
    ['610AP4','OPS-MP4 吋 手動機','QQAOPSMP400000000000','110*280 IN 置中'].concat(AP),
  ];
  mmSheet.getRange(2,1,mmData.length,10).setValues(mmData);

  // 人員名單
  var staffSheet = ss.getSheetByName('人員名單');
  staffSheet.clearContents();
  staffSheet.getRange(1,1,1,2).setValues([['業務員','工程師']]);
  var staffData = [
    ['呂佳朋','陳威豪'],['朱祐慶','翁靖其'],['陳冠華','蔡昀祐'],
    ['謝孟揚','王俊傑'],['包英傑','張正泰'],['蘇士傑','林炯岳'],
    ['','林頡銘'],['','張崇維'],
  ];
  staffSheet.getRange(2,1,staffData.length,2).setValues(staffData);

  // 碳帶庫存
  var invSheet = ss.getSheetByName('碳帶庫存');
  invSheet.clearContents();
  invSheet.getRange(1,1,1,2).setValues([['品名','料號']]);
  var invData = [["AN-011(R)黑 57*300 OUT", "RRAN0116010005700300"], ["AN-011(R)黑 82*300 OUT", "RRAN0116010008200300"], ["AN-210(R)黑 34*600 OUT", "RRAN2106010003400600"], ["AN-510+ 黑 110*600 OUT", "RSAN51P6010011000600"], ["AN-510+ 黑 34*50 OUT", "RSAN51P6010003400050"], ["AN-510+ 黑 34*600 OUT", "RSAN51P6010003400600"], ["AN-510+ 黑 55*600 OUT", "RSAN51P6010005500600"], ["AN-510+ 黑 57*300 OUT", "RSAN51P6010005700300"], ["AN-510+ 黑 82*300 OUT", "RSAN51P6010008200300"], ["AR-710黑 82*300 OUT", "RRAR7106010008200300"], ["AS-900黑 82*50 OUT", "RSAS9006010008200050"], ["DN-320(R) 黑 34*600 OUT", "RRDN3206010003400600"], ["DN-320(R) 黑 55*600 OUT", "RRDN3206010005500600"], ["DN-320(R) 黑 57*300 OUT", "RRDN3206010005700300"], ["DN-320(R) 黑 82*300 OUT", "RRDN3206010008200300"], ["DN-320(R) 黑 110MM*280M IN 置中", "RRDN3206M11011000280"], ["DN-421(R) 黑 55*600 OUT", "RRDN4216010005500600"], ["DN-421(R) 黑 57*300 OUT", "RRDN4216010005700300"], ["DN-421(R) 黑 82*300 OUT", "RRDN4216010008200300"], ["DN-421(R) 黑 110*280 IN 置中", "RRDN4216M11011000280"], ["DN-480(R) 黑 55*600 OUT", "RRDN4806010005500600"], ["DN-561 (R) 白 55*600 OUT", "RRDN5617010005500600"], ["DN-561 (R) 白 57*300 OUT", "RRDN5617010005700300"], ["DN-561 (R) 黑 55*600 OUT", "RRDN5616010005500600"], ["DN-561 (R) 黑 57*300 OUT", "RRDN5616010005700300"], ["DN-561黑 110*280 IN 置中", "RRDN5616M11011000280"], ["DN-707(R)黑 34*600 OUT", "RRDN7A76010003400600"], ["DN-707(R)黑 57*300 OUT", "RRDN7A76010005700300"], ["DN-707(R)黑 82*300 OUT", "RRDN7A76010008200300"], ["DN-707(R)黑 110*600 OUT", "RRDN7A76010011000600"], ["DN-808白 34*600 OUT", "RSDN8087010003400600"], ["DN-808白 82*300 OUT", "RSDN8087010008200300"], ["DN-808黑 34*600 OUT", "RSDN8086010003400600"], ["DN-808黑 55*600 OUT", "RSDN8086010005500600"], ["DN-808黑 82*300 OUT", "RSDN8086010008200300"], ["EN-890 (R)黑 110*280 IN 置中", "RREN8906M11011000280"], ["GN-314黑 34*600 OUT", "RSGN3146010003400600"], ["GN-314黑 55*600 OUT", "RSGN3146010005500600"], ["GN-331黑 34*600 OUT", "RRGN3316010003400600"], ["GN-331黑 57*300 OUT", "RRGN3316010005700300"], ["GN-331(R) 白 55MM*600M OUT", "RRGN3317010005500600"], ["GN-331(R) 白 82MM*300M OUT", "RRGN3317010008200300"], ["GN-512黑 34*600 OUT", "RSGN5126010003400600"], ["GN-512黑 55*600 OUT", "RSGN5126010005500600"], ["GN-811(R) 黑 110MM*280M IN 置中", "RRGN8116M11011000280"], ["HL-35白 34*600 OUT", "RRHL3507010003400600"], ["HL-35白 57*300 OUT", "RRHL3507010005700300"], ["KN-830(R) 黑 34*600 OUT", "RRKN8306010003400600"], ["KN-830(R) 黑 55*600 OUT", "RRKN8306010005500600"], ["KN-830(R) 黑 57*300 OUT", "RRKN8306010005700300"], ["KN-830(R) 黑 82*300 OUT", "RRKN8306010008200300"], ["ON-222白 82MM*300M OUT", "RSON2227010008200300"], ["ZR-810金 36*300 OUT", "RRZR8101010003600300"], ["ZR-810金 82*300 OUT", "RRZR8101010008200300"]];
  invSheet.getRange(2,1,invData.length,2).setValues(invData);

  Logger.log('initSettings 完成');
}

// ── GET ──────────────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : '';
  if (action === 'getSettings') {
    return ContentService
      .createTextOutput(JSON.stringify(getSettings()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST ─────────────────────────────────────────────
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    // 後端密碼驗證
    if (d.action === 'saveSettings') {
      if (d.password !== getApiPassword()) {
        return ContentService
          .createTextOutput(JSON.stringify({error: '密碼錯誤，拒絕存取'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return saveSettings(d);
    }
    var pdf = generatePDF(d);
    return ContentService
      .createTextOutput(JSON.stringify({
        pdf: Utilities.base64Encode(pdf.getBytes()),
        filename: '組合單-' + d.model + (d.customer ? '-' + d.customer : '') + '-' + d.date + '.pdf'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 儲存設定 ─────────────────────────────────────────
function saveSettings(d) {
  var ss = SpreadsheetApp.openById(getSheetId());

  if (d.MM) {
    var mmSheet = ss.getSheetByName('機型設定');
    mmSheet.clearContents();
    mmSheet.getRange(1,1,1,10).setValues([[
      '機型','C14品名','C15料號','碳帶尺寸',
      'C21配件1','C22料號','C24配件2','C25料號','C27配件3','C28料號'
    ]]);
    var rows = d.MM.map(function(r) {
      return [r.model,r.c14,r.c15,r.ribbonSz||'',
              r.c21||'',r.c22||'',r.c24||'',r.c25||'',r.c27||'',r.c28||''];
    });
    if (rows.length) mmSheet.getRange(2,1,rows.length,10).setValues(rows);
  }

  if (d.staff) {
    var staffSheet = ss.getSheetByName('人員名單');
    staffSheet.clearContents();
    staffSheet.getRange(1,1,1,2).setValues([['業務員','工程師']]);
    var max = Math.max(d.staff.sales.length, d.staff.engineers.length);
    var rows = [];
    for (var i = 0; i < max; i++) {
      rows.push([d.staff.sales[i]||'', d.staff.engineers[i]||'']);
    }
    if (rows.length) staffSheet.getRange(2,1,rows.length,2).setValues(rows);
  }

  if (d.INV) {
    var invSheet = ss.getSheetByName('碳帶庫存');
    invSheet.clearContents();
    invSheet.getRange(1,1,1,2).setValues([['品名','料號']]);
    var rows = d.INV.map(function(r) { return [r['品名'], r['料號']]; });
    if (rows.length) invSheet.getRange(2,1,rows.length,2).setValues(rows);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 生成 PDF ─────────────────────────────────────────
function generatePDF(d) {
  var copy = null;
  try {
    var cfg = getSettings();
    var MM  = cfg.MM;
    var INV = cfg.INV;

    var model     = d.model     || '';
    var customer  = d.customer  || '';
    var partno    = (d.partno   || '0000').padStart(4,'0');
    var dateVal   = d.date      || '';
    var ribbon    = d.ribbon    || '';
    var bracket   = d.bracket   || '';
    var sales     = d.sales     || '';
    var engineer  = d.engineer  || '';
    var orderno   = d.orderno   || '';
    var ribbonQty = d.ribbon_qty || 1;

    // 複製範本
    var src  = DriveApp.getFileById(getSheetId());
    copy     = src.makeCopy('tmp_' + new Date().getTime());
    var ss2  = SpreadsheetApp.open(copy);
    var ws   = ss2.getSheets()[0];

    function s(addr, val) {
      if (val !== '' && val !== null && val !== undefined) {
        ws.getRange(addr).setValue(val);
      }
    }

    // 日期 / 客戶
    s('F4', dateVal);
    if (customer) ws.getRange('C5').setValue('客戶名稱：' + customer);

    // 機器本體
    var m = MM[model];
    if (m) {
      s('C8',  m.c14 + (customer ? ' (' + customer + ')' : ''));
      s('D8',  'TQ'); s('E8', 1);
      s('C9',  m.c15.substring(0,16) + partno);
      s('C14', m.c14); s('D14','TA'); s('E14', 1);
      s('C15', m.c15);
      // 配件
      if (m.c21) { s('C21', m.c21); s('D21','TA'); s('E21', 1); }
      if (m.c22) { s('C22', m.c22); }
      if (m.c24) { s('C24', m.c24); s('D24','TA'); s('E24', 1); }
      if (m.c25) { s('C25', m.c25); }
      if (m.c27) { s('C27', m.c27); s('D27','TA'); s('E27', 1); }
      if (m.c28) { s('C28', m.c28); }
    }

    // 搭贈固定
    s('C32','酒精筆'); s('D32','TA'); s('E32', 1);
    s('C33','QQTHERMALHEADCLEANER');

    // 碳帶比對
    if (ribbon && m && m.ribbonSz) {
      var rc  = ribbon.toUpperCase().replace(/[-_+ ]/g,'');
      var sz  = m.ribbonSz;
      var isWhite  = /白|white/i.test(ribbon);
      var colorKey = isWhite ? '白' : '黑';

      var hit = INV.find(function(i) {
        var clean = i['品名'].toUpperCase().replace(/[-_+ (R)\s]/g,'');
        return clean.indexOf(rc) >= 0
            && i['品名'].indexOf(sz) >= 0
            && i['品名'].indexOf(colorKey) >= 0;
      });
      if (!hit) {
        hit = INV.find(function(i) {
          var clean = i['品名'].toUpperCase().replace(/[-_+ (R)\s]/g,'');
          return clean.indexOf(rc) >= 0 && i['品名'].indexOf(sz) >= 0;
        });
      }
      if (hit) {
        s('C35', hit['品名']); s('D35','TA'); s('E35', ribbonQty);
        s('C36', hit['料號']);
      }
    }

    // 支架
    if (bracket) {
      var bp = bracket.padStart(4,'0');
      s('C38','支架' + bp + (customer ? ' (' + customer + ')' : ''));
      s('D38','TQ'); s('E38', 1);
      s('C39',('QQTTT00000000000' + bp).substring(0,20));
    }

    // 人員
    if (sales)    ws.getRange('E41').setValue('業   務   員 ：' + sales);
    if (engineer) ws.getRange('E43').setValue('技 術 人 員 ：' + engineer);
    if (orderno)  ws.getRange('C45').setValue('訂單編號:' + orderno);

    // 輸出 PDF
    SpreadsheetApp.flush();
    var sheetId = ws.getSheetId();
    var url = 'https://docs.google.com/spreadsheets/d/' + copy.getId()
      + '/export?exportFormat=pdf&format=pdf&size=A4&portrait=true'
      + '&fitw=true&fith=true'
      + '&top_margin=0.2&bottom_margin=0.2&left_margin=0.2&right_margin=0.2'
      + '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false'
      + '&gid=' + sheetId + '&r1=0&c1=0&r2=45&c2=8';

    var resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
    });
    var pdfBlob = resp.getBlob();

    // 自動存到 Google Drive：組合單/年份/月份/
    try {
      var now      = new Date();
      var year     = String(now.getFullYear());
      var month    = String(now.getMonth() + 1).padStart(2, '0');
      var dateStr  = d.date || (year + '.' + month + '.' + String(now.getDate()).padStart(2, '0'));
      var root     = getOrCreateFolder(DriveApp.getRootFolder(), '組合單');
      var yearDir  = getOrCreateFolder(root, year);
      var monthDir = getOrCreateFolder(yearDir, month);
      var fname    = '組合單-' + model + (customer ? '-' + customer : '') + '-' + dateStr + '.pdf';
      var backupBlob = pdfBlob.copyBlob().setName(fname);
      monthDir.createFile(backupBlob);
    } catch(e) {
      Logger.log('Drive 存檔失敗: ' + e.message);
    }

    return pdfBlob;

  } finally {
    if (copy) { try { copy.setTrashed(true); } catch(e) {} }
  }
}

// ── 取得或建立資料夾 ────────────────────────────────────
function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

// ── 初始化 Properties（第一次設定時執行）────────────────
function initProperties() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID',     '11fuPx2YcZ9Fo0dKUD44y7fcczTU3EWs_Hb4zcv764Y0');
  props.setProperty('API_PASSWORD', '198310');
  Logger.log('Properties 設定完成');
}

// ── 清除暫存檔 ───────────────────────────────────────
function cleanupTempFiles() {
  var all = DriveApp.getFiles(), count = 0;
  while (all.hasNext()) {
    var f = all.next();
    if (f.getName().indexOf('tmp_') === 0) { f.setTrashed(true); count++; }
  }
  Logger.log('已刪除 ' + count + ' 個暫存檔');
}

// ── 授權用（執行一次即可）───────────────────────────
function testAuth() {
  UrlFetchApp.fetch('https://docs.google.com/', {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  Logger.log('授權正常');
}
