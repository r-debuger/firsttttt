/**
 * 餃個朋友｜線上訂單 → Google Sheet（Google Apps Script 網頁應用程式）
 *
 * 使用方式請見 gas/README.md。重點：
 *   1. 在 Google 試算表中開啟「擴充功能 → Apps Script」，貼上本檔。
 *   2. 執行一次 setup()，自動建立「訂單」「訂單明細」「商品」「設定」工作表。
 *   3. 部署為網頁應用程式（執行身分：我；存取權：所有人），把網址填入 index.html 的 ORDER_API_URL。
 *
 * 安全設計：
 *   - 價格一律以「商品」工作表為準，在伺服器重新計算，不採用前端傳來的金額。
 *   - 使用 LockService 避免同時下單時訂單編號重複。
 *   - 同一支電話 10 分鐘內最多 5 筆訂單；含隱藏欄位（honeypot）的機器人請求一律拒絕。
 *   - 所有寫入儲存格的文字都會移除開頭的 = + - @，避免試算表公式注入。
 */

const SHEETS = {
  orders: '訂單',
  items: '訂單明細',
  products: '商品',
  settings: '設定',
};

const ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '姓名', '電話', 'Email', '宅配地址', '備註', '品項摘要', '總件數', '小計', '運費', '總金額', '來源頁面'];
const ITEM_HEADERS = ['訂單編號', '商品ID', '品名', '單價', '數量', '小計'];
const PRODUCT_HEADERS = ['商品ID', '分類', '品名', '價格', '供應中'];
const STATUS_OPTIONS = ['新訂單', '製作中', '可取餐', '已出貨', '已完成', '已取消'];
const DELIVERY_OPTIONS = ['外帶', '內用', '宅配'];

// 與網站 index.html 的商品清單一致；setup() 會寫入「商品」工作表，之後請直接在試算表調整價格或停售。
const DEFAULT_PRODUCTS = [
  ['p1', '鍋貼', '經典高麗菜鍋貼', 85],
  ['p2', '鍋貼', '香蔥爆汁鍋貼', 95],
  ['p3', '鍋貼', '椒麻香辣鍋貼', 100],
  ['p4', '鍋貼', '鮮蝦韭菜鍋貼', 110],
  ['p5', '水餃', '鮮肉高麗菜水餃', 90],
  ['p6', '水餃', '韭黃鮮蝦水餃', 125],
  ['p7', '湯飲', '暖心酸辣湯', 45],
  ['p8', '湯飲', '元氣蔬菜湯', 45],
  ['p9', '湯飲', '冰釀烏龍茶', 35],
  ['p10', '小菜', '秘製涼拌小黃瓜', 45],
  ['p11', '小菜', '香滷豆干拼盤', 55],
  ['p12', '小菜', '芝麻時蔬', 50],
  ['g1', '禮盒', '招牌鍋貼家庭禮盒 50 入', 599],
];

// 「設定」工作表的預設值（可直接在試算表修改，不必改程式）
const DEFAULT_SETTINGS = [
  ['店家通知信箱', '', '有新訂單時寄信通知；留空則寄給部署者本人，填「無」則不寄'],
  ['寄送顧客確認信', '是', '顧客有填 Email 時寄送訂單確認信（是／否）'],
  ['宅配運費', 0, '宅配訂單加收的運費；0 表示不收運費'],
  ['宅配免運門檻', 0, '小計達此金額免運費；0 表示不提供免運'],
  ['暫停接單', '否', '填「是」時網站送出訂單會收到暫停接單訊息'],
];

const MAX_QTY_PER_ITEM = 99;
const RATE_LIMIT = { max: 5, seconds: 600 };

/* ------------------------------------------------------------------ */
/* 初始化                                                               */
/* ------------------------------------------------------------------ */

/** 第一次使用時手動執行一次：建立工作表、標題列、下拉選單與預設資料。可重複執行，不會覆蓋既有資料。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orders = ensureSheet_(ss, SHEETS.orders, ORDER_HEADERS);
  ensureSheet_(ss, SHEETS.items, ITEM_HEADERS);
  const products = ensureSheet_(ss, SHEETS.products, PRODUCT_HEADERS);
  const settings = ensureSheet_(ss, SHEETS.settings, ['項目', '值', '說明']);

  if (products.getLastRow() < 2) {
    products.getRange(2, 1, DEFAULT_PRODUCTS.length, 5).setValues(DEFAULT_PRODUCTS.map(p => p.concat('是')));
  }
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
  }

  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS, true).build();
  orders.getRange(2, 3, orders.getMaxRows() - 1, 1).setDataValidation(statusRule);
  const yesNo = SpreadsheetApp.newDataValidation().requireValueInList(['是', '否'], true).build();
  products.getRange(2, 5, products.getMaxRows() - 1, 1).setDataValidation(yesNo);
  orders.getRange('L:N').setNumberFormat('#,##0');
  orders.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  return '設定完成';
}

function ensureSheet_(ss, name, headers) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#3c2920').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* ------------------------------------------------------------------ */
/* 網頁應用程式進入點                                                    */
/* ------------------------------------------------------------------ */

/** 健康檢查：瀏覽器直接打開部署網址會看到 {"ok":true,...} */
function doGet() {
  return json_({ ok: true, service: 'gyo-friends-orders', open: !isPaused_() });
}

/** 網站送出訂單。前端以 text/plain 傳送 JSON（避免 CORS 預檢）。 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action && body.action !== 'createOrder') return json_({ ok: false, error: '不支援的操作' });
    return json_(createOrder_(body));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: err && err.userMessage ? err.userMessage : '系統忙碌中，請稍後再試或直接致電門市' });
  }
}

/* ------------------------------------------------------------------ */
/* 建立訂單                                                             */
/* ------------------------------------------------------------------ */

function createOrder_(body) {
  if (body.website) return { ok: false, error: '無法送出訂單' }; // honeypot：真人看不到這個欄位
  if (isPaused_()) return { ok: false, error: '目前暫停線上接單，請直接致電門市' };

  const c = body.customer || {};
  const delivery = String(body.delivery || '');
  const name = clean_(c.name, 60);
  const phone = clean_(c.phone, 20);
  const email = clean_(c.email, 120);
  const address = clean_(c.address, 200);
  const note = clean_(c.note, 100);

  if (DELIVERY_OPTIONS.indexOf(delivery) < 0) return fail_('請選擇取餐方式');
  if (!name) return fail_('請填寫聯絡姓名');
  if (!/^[0-9+()\-\s]{7,20}$/.test(phone)) return fail_('請填寫有效的聯絡電話');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return fail_('電子郵件格式不正確');
  if (delivery === '宅配' && !address) return fail_('宅配訂單請填寫收件地址');

  const catalog = loadProducts_();
  const items = [];
  (Array.isArray(body.items) ? body.items : []).forEach(it => {
    const id = String(it && it.id || '');
    const qty = Math.floor(Number(it && it.qty));
    if (!id || !(qty > 0)) return;
    const p = catalog[id];
    if (!p) throw userError_('部分商品已下架，請重新整理頁面後再試');
    if (qty > MAX_QTY_PER_ITEM) throw userError_('單一商品數量上限為 ' + MAX_QTY_PER_ITEM + ' 份');
    items.push({ id: id, name: p.name, price: p.price, qty: qty, subtotal: p.price * qty, cat: p.cat });
  });
  if (!items.length) return fail_('購物車是空的');
  if (delivery === '內用' && items.some(it => it.cat === '禮盒')) return fail_('禮盒為冷凍商品，請改選外帶或宅配');

  if (!rateLimitOk_(phone)) return fail_('下單次數過多，請稍後再試或直接致電門市');

  const subtotal = items.reduce((s, it) => s + it.subtotal, 0);
  const count = items.reduce((s, it) => s + it.qty, 0);
  const settings = loadSettings_();
  const fee = Number(settings['宅配運費']) || 0;
  const freeOver = Number(settings['宅配免運門檻']) || 0;
  const shipping = delivery === '宅配' && !(freeOver > 0 && subtotal >= freeOver) ? fee : 0;
  const total = subtotal + shipping;
  const summary = items.map(it => it.name + ' ×' + it.qty).join('、');
  const now = new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let orderId;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    orderId = nextOrderId_(now);
    ss.getSheetByName(SHEETS.orders).appendRow([
      orderId, now, '新訂單', delivery, safeCell_(name), "'" + phone, safeCell_(email), safeCell_(address), safeCell_(note),
      summary, count, subtotal, shipping, total, safeCell_(clean_(body.page, 200)),
    ]);
    const itemSheet = ss.getSheetByName(SHEETS.items);
    const rows = items.map(it => [orderId, it.id, it.name, it.price, it.qty, it.subtotal]);
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  const order = { orderId: orderId, createdAt: now, delivery: delivery, name: name, phone: phone, email: email, address: address, note: note, items: items, count: count, subtotal: subtotal, shipping: shipping, total: total };
  notify_(order, settings);
  return { ok: true, orderId: orderId, subtotal: subtotal, shipping: shipping, total: total, count: count };
}

/** 訂單編號：GF-年月日-當日流水號，例如 GF-20261127-007。須在 lock 內呼叫。 */
function nextOrderId_(now) {
  const day = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = 'seq_' + day;
  const seq = (Number(props.getProperty(key)) || 0) + 1;
  props.setProperty(key, String(seq));
  return 'GF-' + day + '-' + ('00' + seq).slice(-3);
}

/* ------------------------------------------------------------------ */
/* 通知信                                                               */
/* ------------------------------------------------------------------ */

function notify_(order, settings) {
  const lines = order.items.map(it => '・' + it.name + ' ×' + it.qty + '　NT$' + it.subtotal);
  const contact = [
    '訂單編號：' + order.orderId,
    '取餐方式：' + order.delivery,
    '姓名：' + order.name,
    '電話：' + order.phone,
    order.email && 'Email：' + order.email,
    order.address && '宅配地址：' + order.address,
    order.note && '備註：' + order.note,
  ].filter(Boolean).join('\n');
  const money = ['小計：NT$' + order.subtotal, order.shipping && '運費：NT$' + order.shipping, '總金額：NT$' + order.total].filter(Boolean).join('\n');
  const detail = [contact, lines.join('\n'), money].join('\n\n');

  try {
    let shopTo = String(settings['店家通知信箱'] || '').trim();
    if (!shopTo) shopTo = Session.getEffectiveUser().getEmail();
    if (shopTo && shopTo !== '無') {
      MailApp.sendEmail({ to: shopTo, subject: '【新訂單】' + order.orderId + '｜' + order.delivery + '｜NT$' + order.total, body: detail + '\n\n試算表：' + SpreadsheetApp.getActiveSpreadsheet().getUrl() });
    }
  } catch (err) { console.warn('店家通知信寄送失敗', err); }

  try {
    if (order.email && String(settings['寄送顧客確認信'] || '是') !== '否') {
      MailApp.sendEmail({
        to: order.email,
        name: '餃個朋友',
        subject: '餃個朋友｜已收到您的訂單 ' + order.orderId,
        body: order.name + ' 您好：\n\n感謝您的訂購，我們已收到以下訂單，製作完成後將與您聯繫。\n\n' + detail + '\n\n如需修改或取消，請直接致電門市並告知訂單編號。\n餃個朋友 敬上',
      });
    }
  } catch (err) { console.warn('顧客確認信寄送失敗', err); }
}

/* ------------------------------------------------------------------ */
/* 工具函式                                                             */
/* ------------------------------------------------------------------ */

function loadProducts_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.products);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) {
    DEFAULT_PRODUCTS.forEach(p => { map[p[0]] = { cat: p[1], name: p[2], price: p[3] }; });
    return map;
  }
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().forEach(r => {
    const id = String(r[0]).trim();
    const price = Number(r[3]);
    if (id && price >= 0 && String(r[4]).trim() !== '否') map[id] = { cat: String(r[1]), name: String(r[2]), price: price };
  });
  return map;
}

function loadSettings_() {
  const out = {};
  DEFAULT_SETTINGS.forEach(r => { out[r[0]] = r[1]; });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.settings);
  if (sheet && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(r => { if (String(r[0]).trim()) out[String(r[0]).trim()] = r[1]; });
  }
  return out;
}

function isPaused_() {
  return String(loadSettings_()['暫停接單'] || '否').trim() === '是';
}

function rateLimitOk_(phone) {
  const cache = CacheService.getScriptCache();
  const key = 'rl_' + phone.replace(/\D/g, '');
  const n = Number(cache.get(key)) || 0;
  if (n >= RATE_LIMIT.max) return false;
  cache.put(key, String(n + 1), RATE_LIMIT.seconds);
  return true;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max);
}

/** 防止試算表公式注入：以 = + - @ 開頭的文字前面加上單引號。 */
function safeCell_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function fail_(msg) { return { ok: false, error: msg }; }
function userError_(msg) { const e = new Error(msg); e.userMessage = msg; return e; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
