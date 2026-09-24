/**
 * 防護室耗材管理系統 - Google Apps Script 後端處理程式
 * 維持原工作表名稱：庫存總表、領用紀錄
 * 升級項目：
 *  1. 支援批次贊助者 (相容 數量*效期 以及 數量*效期*贊助者)
 *  2. 新增第 13 欄 (N 欄)：採購數量
 */

const SHEET_INVENTORY = '庫存總表';
const SHEET_LOGS = '領用紀錄';

// 欄位標題：包含原本 13 欄，並在最後擴充「採購數量」
const INVENTORY_HEADERS = [
  '品項ID',        // 0 (A)
  '耗材名稱',      // 1 (B)
  '分類',          // 2 (C)
  '倉庫庫存',      // 3 (D)
  '外區現貨',      // 4 (E)
  '總計庫存',      // 5 (F)
  '單位',          // 6 (G)
  '外區警戒值',    // 7 (H)
  '總庫存警戒值',  // 8 (I)
  '倉庫批次明細',  // 9 (J) (格式: 數量*效期 或 數量*效期*贊助者)
  '外區批次明細',  // 10 (K)
  '是否採購中',    // 11 (L)
  '採購備註',      // 12 (M)
  '採購數量'       // 13 (N) [新增支援快速採購數量]
];
const LOG_HEADERS = ['時間', '登記人', '用途', '領用明細', '紀錄ID'];

// 找不到分頁時自動建立，並補齊標題列
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // 檢查標題欄位是否需要補上新欄位（例如新增的採購數量）
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (existingHeaders.length < headers.length || !existingHeaders[0]) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}

// 把領用明細陣列組成「品名*數量, 品名*數量」的純文字
function formatItemsUsed(itemsUsed) {
  return (itemsUsed || [])
    .map(function (i) { return i.name + '*' + i.qty; })
    .join(', ');
}

// 把儲存格內容還原成領用明細陣列
function parseItemsUsed(raw) {
  raw = String(raw || '').trim();
  if (!raw) return [];
  if (raw.charAt(0) === '[') {
    try { return JSON.parse(raw); } catch (err) { return []; }
  }
  return raw.split(',').map(function (s) {
    var parts = s.trim().split('*');
    return { name: (parts[0] || '').trim(), qty: Number(parts[1]) || 0 };
  }).filter(function (u) { return u.name; });
}

// 把批次陣列組成純文字：支援贊助者寫入
// 格式: 10*2027-06-01*林老師 提供 (若無贊助者則為 10*2027-06-01)
function formatBatches(batches, location) {
  return (batches || [])
    .filter(function (b) { return b.location === location && b.qty > 0; })
    .map(function (b) { 
      var base = b.qty + '*' + (b.expiryDate || '');
      if (b.sponsor && String(b.sponsor).trim() !== '') {
        return base + '*' + String(b.sponsor).trim();
      }
      return base;
    })
    .join(', ');
}

// 把倉庫批次欄與外區批次欄還原成 batches 陣列，完美相容舊格式 (2段) 與新格式 (3段含贊助者)
function parseBatches(warehouseRaw, outerRaw, rowIndex) {
  var batches = [];
  var seq = 0;
  function parseOne(raw, location) {
    raw = String(raw || '').trim();
    if (!raw) return;
    if (raw.charAt(0) === '[') {
      try {
        var parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          parsed.forEach(function (b) {
            seq++;
            batches.push({
              id: b.id || ('batch-r' + rowIndex + '-' + seq),
              qty: Number(b.qty) || 0,
              expiryDate: String(b.expiryDate || ''),
              location: location,
              sponsor: String(b.sponsor || '')
            });
          });
          return;
        }
      } catch (e) {}
    }
    raw.split(',').forEach(function (s) {
      var parts = s.trim().split('*');
      var qty = Number(parts[0]) || 0;
      var expiryDate = (parts[1] || '').trim();
      var sponsor = (parts[2] || '').trim(); // 解析第 3 個欄位：贊助者 (選填)
      
      if (qty > 0 && expiryDate) {
        seq++;
        batches.push({ 
          id: 'batch-r' + rowIndex + '-' + seq, 
          qty: qty, 
          expiryDate: expiryDate, 
          location: location,
          sponsor: sponsor
        });
      }
    });
  }
  parseOne(warehouseRaw, 'warehouse');
  parseOne(outerRaw, 'outer');
  return batches;
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = getOrCreateSheet(ss, SHEET_INVENTORY, INVENTORY_HEADERS);
    var logSheet = getOrCreateSheet(ss, SHEET_LOGS, LOG_HEADERS);

    var items = [];
    var invData = invSheet.getDataRange().getValues();
    for (var i = 1; i < invData.length; i++) {
      var r = invData[i];
      if (!r[0]) continue;
      
      var batches = parseBatches(r[9], r[10], i);
      var hasWarehouseBatches = String(r[9] || '').trim() !== '';
      var hasOuterBatches = String(r[10] || '').trim() !== '';
      
      var warehouseFromBatches = batches.filter(function (b) { return b.location === 'warehouse'; })
        .reduce(function (s, b) { return s + b.qty; }, 0);
      var outerFromBatches = batches.filter(function (b) { return b.location === 'outer'; })
        .reduce(function (s, b) { return s + b.qty; }, 0);

      // 解析 L 欄 (第 11 索引)：是否採購中
      var isPurchasing = r[11] === true || String(r[11]).toLowerCase() === 'true' || r[11] === '是';
      // 解析 M 欄 (第 12 索引)：採購備註
      var purchasingNote = String(r[12] || '');
      // 解析 N 欄 (第 13 索引)：採購數量
      var purchasingQty = Number(r[13]) || 0;

      items.push({
        id: String(r[0]), 
        name: String(r[1] || ''), 
        category: String(r[2] || '未分類'),
        warehouseStock: hasWarehouseBatches ? warehouseFromBatches : (Number(r[3]) || 0),
        outerStock: hasOuterBatches ? outerFromBatches : (Number(r[4]) || 0),
        unit: String(r[6] || '件'), 
        minOuter: Number(r[7]) !== undefined && r[7] !== '' ? Number(r[7]) : 2,
        minTotal: Number(r[8]) !== undefined && r[8] !== '' ? Number(r[8]) : 5,
        batches: batches,
        isPurchasing: isPurchasing,
        purchasingQty: purchasingQty,
        purchasingNote: purchasingNote
      });
    }

    var logs = [];
    var logData = logSheet.getDataRange().getValues();
    for (var j = 1; j < logData.length; j++) {
      var lr = logData[j];
      if (!lr[0]) continue;
      logs.push({
        id: String(lr[4] || ('log-' + j)),
        timestamp: String(lr[0] || ''),
        userName: String(lr[1] || ''),
        purpose: String(lr[2] || ''),
        itemsUsed: parseItemsUsed(lr[3])
      });
    }
    logs.reverse();

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', items: items, logs: logs }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = getOrCreateSheet(ss, SHEET_INVENTORY, INVENTORY_HEADERS);
    var logSheet = getOrCreateSheet(ss, SHEET_LOGS, LOG_HEADERS);

    var invData = invSheet.getDataRange().getValues();
    var idRowMap = {};
    for (var i = 1; i < invData.length; i++) {
      if (invData[i][0]) idRowMap[String(invData[i][0])] = i + 1;
    }

    (data.items || []).forEach(function (item) {
      var rowNum = idRowMap[String(item.id)];
      var whStock = Number(item.warehouseStock) || 0;
      var outStock = Number(item.outerStock) || 0;
      
      var rowValues = [
        item.id,                                              // A: 品項ID
        item.name,                                            // B: 耗材名稱
        item.category || '未分類',                            // C: 分類
        whStock,                                              // D: 倉庫庫存
        outStock,                                             // E: 外區現貨
        whStock + outStock,                                   // F: 總計庫存
        item.unit || '件',                                    // G: 單位
        item.minOuter !== undefined ? item.minOuter : 2,      // H: 外區警戒值
        item.minTotal !== undefined ? item.minTotal : 5,      // I: 總庫存警戒值
        formatBatches(item.batches, 'warehouse'),             // J: 倉庫批次明細 (含贊助者)
        formatBatches(item.batches, 'outer'),                 // K: 外區批次明細 (含贊助者)
        Boolean(item.isPurchasing),                           // L: 是否採購中 (TRUE/FALSE)
        String(item.purchasingNote || ''),                    // M: 採購備註
        Number(item.purchasingQty) || 0                       // N: 採購數量 [新增]
      ];
      
      if (rowNum) {
        invSheet.getRange(rowNum, 1, 1, rowValues.length).setValues([rowValues]);
      } else {
        invSheet.appendRow(rowValues);
      }
    });

    // 支援寫入訪客領用紀錄，以及管理者調撥／採購到貨入庫等
    if ((data.action === 'VISITOR_CHECKOUT' || data.action === 'TRANSFER' || data.action === 'UPDATE' || data.action === 'MANUAL_SYNC' || data.action === 'ARRIVE_RESTOCK') && data.log) {
      logSheet.appendRow([
        data.log.timestamp,
        data.log.userName,
        data.log.purpose,
        formatItemsUsed(data.log.itemsUsed),
        data.log.id
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}