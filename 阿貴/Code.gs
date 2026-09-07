/**
 * 阿貴麵店訂餐網頁 - Apps Script 後端
 * 請將本檔貼到「綁定 Google 試算表」的 Apps Script 專案中。
 */

const ORDERS_SHEET = '訂單';
const ORDER_ITEMS_SHEET = '訂單明細';
const TIME_ZONE = 'Asia/Taipei';
// 請填入「要接收訂單」的 Google 試算表 ID（網址中 /d/ 與 /edit 之間的字串）。
// 留白時會使用此 Apps Script 所綁定的試算表。
const SPREADSHEET_ID = '';

const MENU = {
  restaurant: '合作新村 阿貴麵店',
  currency: 'TWD',
  categories: [
    { name: '麵類', items: [
      { name: '白麵', style: '乾', prices: { '小': 40, '大': 60 } },
      { name: '白麵', style: '湯', prices: { '小': 40, '大': 60 } },
      { name: '黃麵', style: '乾', prices: { '小': 40, '大': 60 } },
      { name: '黃麵', style: '湯', prices: { '小': 40, '大': 60 } },
      { name: '米粉', style: '乾', prices: { '小': 40, '大': 60 } },
      { name: '米粉', style: '湯', prices: { '小': 40, '大': 60 } },
      { name: '餛飩麵', style: '乾', prices: { '小': 65, '大': 85 } },
      { name: '餛飩麵', style: '湯', prices: { '小': 65, '大': 85 } },
      { name: '雞絲麵', prices: { '小': 40, '大': 60 } },
      { name: '粿仔湯', prices: { '小': 40, '大': 60 } },
      { name: '乾粿仔', prices: { '小': 65, '大': 85 } },
      { name: '麻醬麵', prices: { '小': 55, '大': 75 } }
    ]},
    { name: '湯類', items: [
      { name: '貢丸湯', price: 30 }, { name: '蛋花湯', price: 30 }, { name: '餛飩湯', price: 50 }
    ]},
    { name: '小菜類', items: [
      { name: '小豆干', unit: '個', price: 5 }, { name: '大豆干', unit: '個', price: 15 },
      { name: '豆皮', unit: '片', price: 15 }, { name: '海帶', unit: '條', price: 15 }, { name: '滷蛋', unit: '顆', price: 15 }
    ]}
  ],
  addOns: [{ name: '餛飩', quantity: 6, price: 25, condition: '任意餐點加購' }],
  notes: ['特殊要求請先告知', '飲料請自取']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(`${MENU.restaurant}｜線上點餐`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('訂餐系統')
    .addItem('建立／檢查訂單工作表', 'setupOrderSheets')
    .addToUi();
}

function setupOrderSheets() {
  const sheets = ensureOrderSheets_();
  SpreadsheetApp.getUi().alert(`已準備完成：${sheets.orders.getName()}、${sheets.items.getName()}\n\n目標試算表：${sheets.spreadsheet.getUrl()}`);
}

function getMenu() {
  return MENU;
}

function submitOrder(order) {
  const normalized = validateAndNormalizeOrder_(order);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  try {
    const sheets = ensureOrderSheets_();
    const timestamp = new Date();
    const orderId = makeOrderId_(timestamp);
    const itemSummary = normalized.items.map(item => `${item.name} ×${item.quantity}`).join('、');

    sheets.orders.appendRow([
      orderId,
      timestamp,
      normalized.orderType,
      safeCell_(normalized.customerName),
      safeCell_(normalized.tableNumber),
      safeCell_(normalized.phone),
      safeCell_(normalized.note),
      normalized.total,
      normalized.items.length,
      safeCell_(itemSummary),
      '新訂單'
    ]);

    const itemRows = normalized.items.map(item => [
      orderId,
      timestamp,
      item.category,
      safeCell_(item.name),
      safeCell_(item.size || ''),
      item.unitPrice,
      item.quantity,
      item.subtotal
    ]);
    sheets.items.getRange(sheets.items.getLastRow() + 1, 1, itemRows.length, itemRows[0].length).setValues(itemRows);

    return {
      ok: true,
      orderId,
      total: normalized.total,
      createdAt: Utilities.formatDate(timestamp, TIME_ZONE, 'yyyy/MM/dd HH:mm:ss'),
      ordersSheetName: sheets.orders.getName(),
      ordersSheetUrl: `${sheets.spreadsheet.getUrl()}#gid=${sheets.orders.getSheetId()}`
    };
  } finally {
    lock.releaseLock();
  }
}

function ensureOrderSheets_() {
  const spreadsheet = getTargetSpreadsheet_();
  const orders = getOrCreateSheet_(spreadsheet, ORDERS_SHEET, [
    '訂單編號', '下單時間', '取餐方式', '訂購人', '桌號', '電話', '訂單備註', '總金額', '品項數', '訂單摘要', '狀態'
  ]);
  const items = getOrCreateSheet_(spreadsheet, ORDER_ITEMS_SHEET, [
    '訂單編號', '下單時間', '分類', '品項', '規格', '單價', '數量', '小計'
  ]);
  return { spreadsheet, orders, items };
}

function getTargetSpreadsheet_() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== '請貼上試算表ID') {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('找不到目標試算表。請在 Code.gs 的 SPREADSHEET_ID 填入試算表 ID，然後重新部署。');
  }
  return spreadsheet;
}

function getOrCreateSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#c1121f').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function validateAndNormalizeOrder_(order) {
  if (!order || typeof order !== 'object' || !Array.isArray(order.items) || order.items.length === 0) {
    throw new Error('請至少選擇一項餐點。');
  }
  const orderType = order.orderType === '內用' ? '內用' : '外帶';
  const customerName = cleanText_(order.customerName, 40);
  const tableNumber = cleanText_(order.tableNumber, 20);
  const phone = cleanText_(order.phone, 30);
  const note = cleanText_(order.note, 300);
  if (orderType === '內用' && !tableNumber) throw new Error('內用請填寫桌號。');
  if (orderType === '外帶' && !customerName) throw new Error('外帶請填寫訂購人姓名。');

  const menuIndex = buildMenuIndex_();
  const combined = {};
  order.items.forEach(rawItem => {
    const key = String(rawItem.key || '');
    const quantity = Number(rawItem.quantity);
    if (!Number.isInteger(quantity) || quantity < 1 || quantity > 99 || !menuIndex[key]) {
      throw new Error('訂單資料不正確，請重新整理後再送出。');
    }
    combined[key] = (combined[key] || 0) + quantity;
    if (combined[key] > 99) throw new Error('單一品項最多 99 份。');
  });

  const items = Object.keys(combined).map(key => {
    const item = menuIndex[key];
    const quantity = combined[key];
    return Object.assign({}, item, { quantity, subtotal: item.unitPrice * quantity });
  });
  return { orderType, customerName, tableNumber, phone, note, items, total: items.reduce((sum, item) => sum + item.subtotal, 0) };
}

function buildMenuIndex_() {
  const index = {};
  MENU.categories.forEach(category => category.items.forEach(item => {
    const displayName = `${item.style || ''}${item.name}`;
    if (item.prices) {
      Object.keys(item.prices).forEach(size => {
        const key = `${category.name}|${item.name}|${item.style || ''}|${size}`;
        index[key] = { category: category.name, name: displayName, size, unitPrice: item.prices[size] };
      });
    } else {
      const key = `${category.name}|${item.name}||`;
      index[key] = { category: category.name, name: item.name, size: '', unitPrice: item.price };
    }
  }));
  MENU.addOns.forEach(item => {
    const key = `加購|${item.name}||`;
    index[key] = { category: '加購', name: `${item.name}（${item.quantity}顆）`, size: '', unitPrice: item.price };
  });
  return index;
}

function cleanText_(value, maxLength) {
  return String(value || '').trim().replace(/\s+/g, ' ').slice(0, maxLength);
}

function safeCell_(value) {
  const text = String(value || '');
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function makeOrderId_(date) {
  return `AG-${Utilities.formatDate(date, TIME_ZONE, 'yyyyMMdd-HHmmss')}-${Math.floor(Math.random() * 900 + 100)}`;
}
