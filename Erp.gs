/**
 * 蓁便誼 ERP：後台（Google Apps Script，獨立專案）
 *
 * 功能：商品（國際條碼）、進貨、門市掃碼結帳、盤點、收支記帳、月報
 * 資料存在一份獨立的試算表「蓁便誼ERP」（setupErp 會自動建立）。
 *
 * 權限：
 *   - 用 LINE 登入。第一個登入的人自動成為「老闆」，之後的人登記為「員工」並等老闆啟用。
 *   - 員工：結帳、進貨（不填進價）、盤點、新增商品。
 *   - 老闆／總會計：全部功能，另外看得到成本、毛利、收支與月報。
 *     要給會計權限：在「使用者」工作表把那個人的「角色」改成「總會計」並勾選啟用。
 *
 * 第一次使用：選 setupErp 按「執行」一次，執行記錄會印出試算表網址。
 *
 * 指令碼屬性：
 *   LINE_LOGIN_CHANNEL_ID  預設 2011973016（和訂購系統同一個 LINE Login 頻道）
 *   ERP_SHEET_ID           setupErp 自動填入
 */

const E_SHEET = {
  USERS: '使用者', PRODUCTS: '商品', PURCHASES: '進貨', SALES: '銷貨',
  COUNTS: '盤點', LEDGER: '收支', AGENT: '代購',
};
const E_HEAD = {
  USERS:     ['LINE用戶ID', 'LINE名稱', '姓名', '角色', '啟用', '加入時間'],
  PRODUCTS:  ['條碼', '品名', '分類', '單位', '售價', '成本', '庫存', '安全庫存', '啟用', '更新時間', '效期'],
  PURCHASES: ['單號', '時間', '供應商', '條碼', '品名', '數量', '進價', '小計', '付款方式', '經手人'],
  SALES:     ['單號', '時間', '條碼', '品名', '數量', '售價', '小計', '成本小計', '付款方式', '經手人'],
  COUNTS:    ['時間', '條碼', '品名', '系統庫存', '實際庫存', '差異', '原因', '經手人', '進貨價(成本)', '售價', '效期'],
  LEDGER:    ['日期', '類型', '分類', '金額', '說明', '付款方式', '單號', '經手人'],
  AGENT:     ['單號', '建立時間', '客戶', '電話', '平台', '商品內容', '商品金額', '服務費', '應收合計',
              '預計到貨', '狀態', '物流單號', '備註', '經手人', '更新時間'],
};
const AGENT_STATUS = ['待下單', '已下單', '已到貨', '已取貨', '已取消'];
const AGENT_OPEN = ['待下單', '已下單', '已到貨'];
const AGENT_PLATFORMS = ['蝦皮', '淘寶', '1688', 'momo', 'PChome', '日本網購', '其他'];
/** 代購服務費：商品金額 5%，最低 35 元 */
function eAgentFee_(amount) { return amount > 0 ? Math.max(35, Math.round(amount * 0.05)) : 0; }
/** 線上訂購系統的試算表（蓁便誼訂單系統），儀表板用 */
const ORDER_SHEET_DEFAULT = '1JRF09OnZ1M62OIWr3t0BQc-AhmHR7kOk1vR3k_ZDb04';
const OWNER = '老闆', ACCOUNTANT = '總會計', STAFF = '員工';
const ROLES = [OWNER, ACCOUNTANT, STAFF];
/** 老闆和總會計看得到成本、帳務與月報 */
function eFull_(me) { return me.role === OWNER || me.role === ACCOUNTANT; }
const PAY_SALE = ['現金', 'LINE Pay', '轉帳', '其他'];
const PAY_BUY = ['現金', '轉帳', '月結（未付）'];
const LEDGER_CATS = {
  '收入': ['銷貨收入', '代購服務費', '代收貨服務費', '旅遊代訂', '其他收入'],
  '支出': ['進貨', '房租', '水電瓦斯', '薪資', '運費', '包材耗材', '廣告行銷', '設備', '雜支'],
};
const E_TZ = 'Asia/Taipei';

/* ---------------- 第一次使用 ---------------- */

function setupErp() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const id = props.getProperty('ERP_SHEET_ID');
  if (id) ss = SpreadsheetApp.openById(id);
  else {
    ss = SpreadsheetApp.create('蓁便誼ERP');
    props.setProperty('ERP_SHEET_ID', ss.getId());
  }
  if (!props.getProperty('LINE_LOGIN_CHANNEL_ID')) props.setProperty('LINE_LOGIN_CHANNEL_ID', '2011973016');

  Object.keys(E_SHEET).forEach(k => {
    const name = E_SHEET[k], head = E_HEAD[k];
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(head);
      sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#FCE4EC');
      sh.setFrozenRows(1);
    }
    // 條碼一律存成文字，避免開頭的 0 被吃掉
    const col = head.indexOf('條碼') + 1;
    if (col) sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
  });
  const blank = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);
  Logger.log('完成！ERP 試算表：' + ss.getUrl());
}

/* ---------------- API 入口 ---------------- */

function doGet() {
  return eOut_({ ok: true, msg: '蓁便誼 ERP API 運作中' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const me = eUser_(eVerify_(req.idToken));
    if (req.action === 'init') return eOut_({ ok: true, ...eInit_(me) });
    if (!me.enabled) throw new Error('帳號尚未啟用，請老闆在「使用者」工作表勾選啟用');
    const owner = eFull_(me);
    const needOwner = () => { if (!owner) throw new Error('這個功能只有老闆和總會計可以使用'); };
    switch (req.action) {
      case 'saveProduct': return eOut_({ ok: true, product: eSaveProduct_(me, req.product) });
      case 'purchase':    return eOut_({ ok: true, ...ePurchase_(me, req.purchase) });
      case 'sale':        return eOut_({ ok: true, ...eSale_(me, req.sale) });
      case 'count':       return eOut_({ ok: true, ...eCount_(me, req.count) });
      case 'today':       return eOut_({ ok: true, ...eToday_(me) });
      case 'addLedger':   needOwner(); return eOut_({ ok: true, ...eAddLedger_(me, req.entry) });
      case 'report':      needOwner(); return eOut_({ ok: true, ...eReport_(req.month) });
      case 'agentList':   return eOut_({ ok: true, ...eAgentList_() });
      case 'agentSave':   return eOut_({ ok: true, ...eAgentSave_(me, req.agent) });
      case 'dashboard':   needOwner(); return eOut_({ ok: true, ...eDashboard_() });
      default: throw new Error('未知的動作');
    }
  } catch (err) {
    return eOut_({ ok: false, error: String(err && err.message || err) });
  }
}

function eOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function eProp_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v === null || v === '' ? fallback : v;
}

/* ---------------- 試算表小工具 ---------------- */

function eSS_() {
  const id = eProp_('ERP_SHEET_ID');
  if (!id) throw new Error('後台尚未初始化，請先執行 setupErp()');
  return SpreadsheetApp.openById(id);
}

function eTable_(name) {
  let sh = eSS_().getSheetByName(name);
  if (!sh) {
    // 新版加的工作表（例如「代購」）自動建立，不用重跑 setupErp
    const key = Object.keys(E_SHEET).find(k => E_SHEET[k] === name);
    if (!key) throw new Error('找不到工作表「' + name + '」，請先執行 setupErp()');
    sh = eSS_().insertSheet(name);
    sh.appendRow(E_HEAD[key]);
    sh.getRange(1, 1, 1, E_HEAD[key].length).setFontWeight('bold').setBackground('#FCE4EC');
    sh.setFrozenRows(1);
  }
  // 新版加的欄位（例如「效期」）自動補在表頭最後面，不用重跑 setupErp
  const tkey = Object.keys(E_SHEET).find(k => E_SHEET[k] === name);
  if (tkey && sh.getLastRow() > 0) {
    const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    const miss = E_HEAD[tkey].filter(h => cur.indexOf(h) < 0);
    if (miss.length) {
      const start = cur.filter(h => h !== '').length + 1;
      sh.getRange(1, start, 1, miss.length).setValues([miss]).setFontWeight('bold').setBackground('#FCE4EC');
      if (miss.indexOf('效期') >= 0) sh.getRange(2, start + miss.indexOf('效期'), Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd');
    }
  }
  const values = sh.getDataRange().getValues();
  const head = values.shift().map(String);
  const col = (h) => head.indexOf(h);
  const rows = [];
  values.forEach((r, i) => {
    if (r.every(c => c === '' || c === false)) return;
    const o = { _row: i + 2 };
    head.forEach((h, j) => { o[h] = r[j]; });
    rows.push(o);
  });
  return { sh, head, col, rows };
}

/** 寫入一列（依表頭順序），條碼欄強制當文字 */
function eAppend_(t, obj) {
  const vals = t.head.map(h => (obj[h] === undefined ? '' : obj[h]));
  const row = t.sh.getLastRow() + 1;
  const range = t.sh.getRange(row, 1, 1, t.head.length);
  const bc = t.col('條碼');
  if (bc >= 0) t.sh.getRange(row, bc + 1).setNumberFormat('@');
  range.setValues([vals]);
  return row;
}

function eAppendMany_(t, objs) {
  if (!objs.length) return;
  const row = t.sh.getLastRow() + 1;
  const bc = t.col('條碼');
  if (bc >= 0) t.sh.getRange(row, bc + 1, objs.length, 1).setNumberFormat('@');
  t.sh.getRange(row, 1, objs.length, t.head.length)
    .setValues(objs.map(o => t.head.map(h => (o[h] === undefined ? '' : o[h]))));
}

function eSet_(t, rowObj, field, value) {
  t.sh.getRange(rowObj._row, t.col(field) + 1).setValue(value);
  rowObj[field] = value;
}

function eNow_() { return new Date(); }
function eFmt_(d, p) { return Utilities.formatDate(d instanceof Date ? d : new Date(d), E_TZ, p); }
/** 日期欄位轉成 yyyy-MM-dd 文字（空白就回空字串） */
function eDay_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return eFmt_(v, 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}
function eNum_(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function eRound_(n) { return Math.round(n * 100) / 100; }
function eLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function eId_(prefix) { return prefix + eFmt_(eNow_(), 'yyMMdd-HHmmss') + String(Math.floor(Math.random() * 90 + 10)); }

/* ---------------- LINE 身分與使用者 ---------------- */

function eVerify_(idToken) {
  if (!idToken) throw new Error('請先用 LINE 登入');
  const cache = CacheService.getScriptCache();
  const key = 'etok_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: idToken, client_id: eProp_('LINE_LOGIN_CHANNEL_ID', '2011973016') },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('LINE 登入已過期，請重新整理頁面');
  const body = JSON.parse(res.getContentText());
  const user = { userId: body.sub, displayName: body.name || '' };
  cache.put(key, JSON.stringify(user), 600);
  return user;
}

/** 找使用者；沒有的話自動登記。第一個人是老闆（直接啟用），其他人是員工（等啟用）。 */
function eUser_(u) {
  const t = eTable_(E_SHEET.USERS);
  const hit = t.rows.find(r => r['LINE用戶ID'] === u.userId);
  if (hit) {
    return { userId: u.userId, name: String(hit['姓名'] || hit['LINE名稱'] || u.displayName),
             role: ROLES.indexOf(String(hit['角色']).trim()) >= 0 ? String(hit['角色']).trim() : STAFF, enabled: hit['啟用'] === true };
  }
  return eLock_(() => {
    const first = eTable_(E_SHEET.USERS).rows.length === 0;
    const role = first ? OWNER : STAFF;
    const row = eAppend_(t, { 'LINE用戶ID': u.userId, 'LINE名稱': u.displayName, '姓名': u.displayName,
      '角色': role, '啟用': first, '加入時間': eNow_() });
    t.sh.getRange(row, t.col('啟用') + 1).insertCheckboxes();
    t.sh.getRange(row, t.col('角色') + 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(ROLES, true).build());
    if (first) t.sh.getRange(row, t.col('啟用') + 1).setValue(true);
    return { userId: u.userId, name: u.displayName, role, enabled: first, isNew: true };
  });
}

/* ---------------- 讀取 ---------------- */

function eProducts_(owner) {
  return eTable_(E_SHEET.PRODUCTS).rows.map(r => {
    const p = {
      code: String(r['條碼']), name: String(r['品名']), cat: String(r['分類'] || ''), unit: String(r['單位'] || ''),
      price: eNum_(r['售價']), stock: eNum_(r['庫存']), safety: eNum_(r['安全庫存']),
      active: r['啟用'] !== false, exp: eDay_(r['效期']),
    };
    if (owner) p.cost = eNum_(r['成本']);
    return p;
  });
}

function eInit_(me) {
  const base = { me: { name: me.name, role: me.role, enabled: me.enabled, isNew: !!me.isNew } };
  if (!me.enabled) return base;
  const owner = eFull_(me);
  return {
    ...base,
    products: eProducts_(owner),
    paySale: PAY_SALE, payBuy: PAY_BUY,
    agentStatus: AGENT_STATUS, agentPlatforms: AGENT_PLATFORMS,
    ledgerCats: owner ? LEDGER_CATS : null,
    today: eToday_(me),
  };
}

function eToday_(me) {
  const d = eFmt_(eNow_(), 'yyyy-MM-dd');
  const t = eTable_(E_SHEET.SALES);
  const orders = {};
  t.rows.forEach(r => {
    if (eFmt_(r['時間'], 'yyyy-MM-dd') !== d) return;
    const id = String(r['單號']);
    const o = orders[id] || (orders[id] = { id, time: eFmt_(r['時間'], 'HH:mm'), total: 0, pay: r['付款方式'], by: r['經手人'], lines: [] });
    o.total += eNum_(r['小計']);
    o.lines.push(r['品名'] + (r['條碼'] === '折扣' ? '' : ' x' + r['數量']));
  });
  const list = Object.keys(orders).map(k => orders[k]).reverse();
  const byPay = {};
  list.forEach(o => { byPay[o.pay] = (byPay[o.pay] || 0) + o.total; });
  return { count: list.length, total: list.reduce((s, o) => s + o.total, 0), byPay, orders: list.slice(0, 30) };
}

/* ---------------- 商品 ---------------- */

function eSaveProduct_(me, p) {
  if (!p || !String(p.name || '').trim()) throw new Error('請填品名');
  const owner = eFull_(me);
  return eLock_(() => {
    const t = eTable_(E_SHEET.PRODUCTS);
    let code = String(p.code || '').trim();
    if (!code) {
      // 沒有國際條碼的商品，自動給店內編號 ZBY00001…
      const max = t.rows.reduce((m, r) => {
        const x = /^ZBY(\d+)$/.exec(String(r['條碼'])); return x ? Math.max(m, Number(x[1])) : m;
      }, 0);
      code = 'ZBY' + String(max + 1).padStart(5, '0');
    }
    const hit = t.rows.find(r => String(r['條碼']) === code);
    const fields = {
      '品名': String(p.name).trim(), '分類': String(p.cat || '').trim(), '單位': String(p.unit || '').trim(),
      '售價': eNum_(p.price), '安全庫存': eNum_(p.safety), '啟用': p.active !== false, '更新時間': eNow_(),
    };
    if (owner && p.cost !== undefined && p.cost !== '') fields['成本'] = eNum_(p.cost);
    if (hit) {
      if (p.isNew) throw new Error('條碼 ' + code + ' 已經有商品了：' + hit['品名']);
      Object.keys(fields).forEach(k => eSet_(t, hit, k, fields[k]));
    } else {
      const row = eAppend_(t, { '條碼': code, '庫存': 0, '成本': fields['成本'] || 0, ...fields });
      t.sh.getRange(row, t.col('啟用') + 1).insertCheckboxes();
      t.sh.getRange(row, t.col('啟用') + 1).setValue(fields['啟用']);
    }
    return eProducts_(owner).find(x => x.code === code);
  });
}

function eFind_(t, code) {
  const hit = t.rows.find(r => String(r['條碼']) === String(code));
  if (!hit) throw new Error('找不到商品條碼 ' + code);
  return hit;
}

/* ---------------- 進貨 ---------------- */

function ePurchase_(me, pu) {
  const owner = eFull_(me);
  const items = (pu && pu.items || []).filter(i => eNum_(i.qty) > 0);
  if (!items.length) throw new Error('請至少加入一項商品');
  const pay = PAY_BUY.indexOf(pu.pay) >= 0 ? pu.pay : PAY_BUY[0];
  return eLock_(() => {
    const t = eTable_(E_SHEET.PRODUCTS);
    const id = eId_('P'), now = eNow_();
    let total = 0;
    const lines = items.map(i => {
      const p = eFind_(t, i.code);
      const qty = eNum_(i.qty);
      const price = owner && i.price !== '' && i.price !== undefined && i.price !== null ? eNum_(i.price) : '';
      const oldStock = eNum_(p['庫存']), oldCost = eNum_(p['成本']);
      if (price !== '') {
        // 移動平均成本：(舊庫存×舊成本＋本次數量×進價) ÷ 新庫存
        const base = Math.max(oldStock, 0);
        eSet_(t, p, '成本', eRound_(base > 0 ? (base * oldCost + qty * price) / (base + qty) : price));
        total += qty * price;
      }
      eSet_(t, p, '庫存', oldStock + qty);
      eSet_(t, p, '更新時間', now);
      return { '單號': id, '時間': now, '供應商': String(pu.supplier || ''), '條碼': String(p['條碼']), '品名': p['品名'],
        '數量': qty, '進價': price, '小計': price === '' ? '' : eRound_(qty * price), '付款方式': pay, '經手人': me.name };
    });
    eAppendMany_(eTable_(E_SHEET.PURCHASES), lines);
    if (owner && total > 0 && pay !== '月結（未付）') {
      eAppend_(eTable_(E_SHEET.LEDGER), { '日期': now, '類型': '支出', '分類': '進貨', '金額': eRound_(total),
        '說明': (pu.supplier ? pu.supplier + ' ' : '') + lines.length + ' 項', '付款方式': pay, '單號': id, '經手人': me.name });
    }
    return { id, count: lines.length, total: owner ? eRound_(total) : undefined, products: eProducts_(owner) };
  });
}

/* ---------------- 結帳（銷貨） ---------------- */

function eSale_(me, s) {
  const owner = eFull_(me);
  const items = (s && s.items || []).filter(i => eNum_(i.qty) > 0);
  if (!items.length) throw new Error('購物車是空的');
  const pay = PAY_SALE.indexOf(s.pay) >= 0 ? s.pay : PAY_SALE[0];
  const discount = Math.max(0, eNum_(s.discount));
  return eLock_(() => {
    const t = eTable_(E_SHEET.PRODUCTS);
    const id = eId_('S'), now = eNow_();
    const warn = [];
    const lines = items.map(i => {
      const p = eFind_(t, i.code);
      const qty = eNum_(i.qty), price = eNum_(p['售價']);
      const stock = eNum_(p['庫存']) - qty;
      if (stock < 0) warn.push(p['品名'] + ' 庫存變成 ' + stock + '，記得盤點');
      eSet_(t, p, '庫存', stock);
      return { '單號': id, '時間': now, '條碼': String(p['條碼']), '品名': p['品名'], '數量': qty, '售價': price,
        '小計': eRound_(qty * price), '成本小計': eRound_(qty * eNum_(p['成本'])), '付款方式': pay, '經手人': me.name };
    });
    let total = lines.reduce((a, l) => a + l['小計'], 0);
    if (discount > 0) {
      const d = Math.min(discount, total);
      lines.push({ '單號': id, '時間': now, '條碼': '折扣', '品名': '折扣', '數量': 1, '售價': -d, '小計': -d,
        '成本小計': 0, '付款方式': pay, '經手人': me.name });
      total -= d;
    }
    eAppendMany_(eTable_(E_SHEET.SALES), lines);
    eAppend_(eTable_(E_SHEET.LEDGER), { '日期': now, '類型': '收入', '分類': '銷貨收入', '金額': eRound_(total),
      '說明': '門市結帳 ' + items.length + ' 項', '付款方式': pay, '單號': id, '經手人': me.name });
    const paid = eNum_(s.paid);
    return { id, total: eRound_(total), change: paid > total ? eRound_(paid - total) : 0, warn,
      products: eProducts_(owner), today: eToday_(me) };
  });
}

/* ---------------- 盤點 ---------------- */

function eCount_(me, c) {
  if (!c || c.actual === '' || c.actual === undefined) throw new Error('請輸入實際數量');
  const owner = eFull_(me);
  let exp = '';
  if (c.exp) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(c.exp))) throw new Error('效期格式要像 2026-12-31');
    exp = new Date(c.exp + 'T12:00:00+08:00');
  }
  return eLock_(() => {
    const t = eTable_(E_SHEET.PRODUCTS);
    const p = eFind_(t, c.code);
    const before = eNum_(p['庫存']), actual = eNum_(c.actual);
    eSet_(t, p, '庫存', actual);
    // 盤點時順便更新售價、效期；進貨價（成本）只有老闆／總會計可以改
    if (c.price !== undefined && c.price !== '' && c.price !== null) eSet_(t, p, '售價', eNum_(c.price));
    if (owner && c.cost !== undefined && c.cost !== '' && c.cost !== null) eSet_(t, p, '成本', eNum_(c.cost));
    if (exp) eSet_(t, p, '效期', exp);
    eSet_(t, p, '更新時間', eNow_());
    eAppend_(eTable_(E_SHEET.COUNTS), { '時間': eNow_(), '條碼': String(p['條碼']), '品名': p['品名'],
      '系統庫存': before, '實際庫存': actual, '差異': actual - before, '原因': String(c.reason || ''), '經手人': me.name,
      '進貨價(成本)': eNum_(p['成本']), '售價': eNum_(p['售價']), '效期': exp || p['效期'] || '' });
    return { diff: actual - before, products: eProducts_(owner) };
  });
}

/* ---------------- 收支與月報（老闆） ---------------- */

function eAddLedger_(me, x) {
  if (!x || !LEDGER_CATS[x.type]) throw new Error('請選收入或支出');
  const amount = eNum_(x.amount);
  if (amount <= 0) throw new Error('金額要大於 0');
  const date = x.date ? new Date(x.date + 'T12:00:00+08:00') : eNow_();
  eAppend_(eTable_(E_SHEET.LEDGER), { '日期': date, '類型': x.type, '分類': String(x.cat || '其他'), '金額': amount,
    '說明': String(x.note || ''), '付款方式': String(x.pay || ''), '單號': '', '經手人': me.name });
  return eReport_(eFmt_(date, 'yyyy-MM'));
}

function eReport_(month) {
  month = /^\d{4}-\d{2}$/.test(month || '') ? month : eFmt_(eNow_(), 'yyyy-MM');
  const inMonth = (d) => d !== '' && eFmt_(d, 'yyyy-MM') === month;

  let revenue = 0, cogs = 0, orders = {};
  const top = {};
  eTable_(E_SHEET.SALES).rows.forEach(r => {
    if (!inMonth(r['時間'])) return;
    revenue += eNum_(r['小計']); cogs += eNum_(r['成本小計']); orders[r['單號']] = 1;
    if (r['條碼'] !== '折扣') {
      const k = String(r['品名']);
      top[k] = top[k] || { name: k, qty: 0, amount: 0 };
      top[k].qty += eNum_(r['數量']); top[k].amount += eNum_(r['小計']);
    }
  });

  const income = {}, expense = {}, entries = [];
  eTable_(E_SHEET.LEDGER).rows.forEach(r => {
    if (!inMonth(r['日期'])) return;
    const bucket = r['類型'] === '收入' ? income : expense;
    bucket[r['分類']] = (bucket[r['分類']] || 0) + eNum_(r['金額']);
    entries.push({ date: eFmt_(r['日期'], 'MM/dd'), type: r['類型'], cat: r['分類'], amount: eNum_(r['金額']),
      note: r['說明'], pay: r['付款方式'], by: r['經手人'] });
  });
  const sum = (o) => Object.keys(o).reduce((a, k) => a + o[k], 0);
  const otherIncome = sum(income) - (income['銷貨收入'] || 0);
  const opex = sum(expense) - (expense['進貨'] || 0);
  const gross = revenue - cogs;
  return {
    month,
    pnl: { revenue: eRound_(revenue), cogs: eRound_(cogs), gross: eRound_(gross), otherIncome: eRound_(otherIncome),
           opex: eRound_(opex), net: eRound_(gross + otherIncome - opex), orders: Object.keys(orders).length },
    cash: { in: eRound_(sum(income)), out: eRound_(sum(expense)), net: eRound_(sum(income) - sum(expense)) },
    income, expense,
    top: Object.keys(top).map(k => top[k]).sort((a, b) => b.amount - a.amount).slice(0, 10),
    entries: entries.reverse().slice(0, 100),
  };
}

/* ---------------- 代購 ---------------- */

function eAgentRow_(r) {
  const d = (v) => (v instanceof Date ? eFmt_(v, 'yyyy-MM-dd') : String(v || ''));
  return {
    id: String(r['單號']), created: r['建立時間'] ? eFmt_(r['建立時間'], 'MM/dd HH:mm') : '',
    customer: String(r['客戶'] || ''), phone: String(r['電話'] || ''), platform: String(r['平台'] || ''),
    item: String(r['商品內容'] || ''), amount: eNum_(r['商品金額']), fee: eNum_(r['服務費']), total: eNum_(r['應收合計']),
    eta: d(r['預計到貨']), status: String(r['狀態'] || '待下單'), tracking: String(r['物流單號'] || ''),
    note: String(r['備註'] || ''), by: String(r['經手人'] || ''),
  };
}

function eAgentList_() {
  const rows = eTable_(E_SHEET.AGENT).rows.map(eAgentRow_);
  const open = rows.filter(a => AGENT_OPEN.indexOf(a.status) >= 0);
  const done = rows.filter(a => AGENT_OPEN.indexOf(a.status) < 0).reverse().slice(0, 30);
  return { agents: open.concat(done) };
}

/** 新增或更新代購單。狀態改成「已取貨」時，自動記一筆「代購服務費」收入。 */
function eAgentSave_(me, a) {
  if (!a) throw new Error('沒有資料');
  return eLock_(() => {
    const t = eTable_(E_SHEET.AGENT);
    const now = eNow_();
    const status = AGENT_STATUS.indexOf(a.status) >= 0 ? a.status : '待下單';
    let hit = a.id ? t.rows.find(r => String(r['單號']) === String(a.id)) : null;
    const amount = a.amount === undefined ? (hit ? eNum_(hit['商品金額']) : 0) : eNum_(a.amount);
    const fee = a.fee === '' || a.fee === undefined || a.fee === null ? eAgentFee_(amount) : eNum_(a.fee);
    const fields = {
      '客戶': String(a.customer || '').trim(), '電話': String(a.phone || '').trim(), '平台': String(a.platform || '').trim(),
      '商品內容': String(a.item || '').trim(), '商品金額': amount, '服務費': fee, '應收合計': amount + fee,
      '預計到貨': String(a.eta || ''), '狀態': status, '物流單號': String(a.tracking || '').trim(),
      '備註': String(a.note || '').trim(), '更新時間': now,
    };
    if (!fields['客戶']) throw new Error('請填客戶名稱');
    if (!fields['商品內容']) throw new Error('請填代購商品');
    const before = hit ? String(hit['狀態']) : '';
    let id;
    if (hit) {
      id = String(hit['單號']);
      Object.keys(fields).forEach(k => eSet_(t, hit, k, fields[k]));
    } else {
      id = eId_('A');
      eAppend_(t, { '單號': id, '建立時間': now, '經手人': me.name, ...fields });
    }
    if (status === '已取貨' && before !== '已取貨' && fee > 0) {
      eAppend_(eTable_(E_SHEET.LEDGER), { '日期': now, '類型': '收入', '分類': '代購服務費', '金額': fee,
        '說明': fields['客戶'] + '｜' + fields['商品內容'].slice(0, 30), '付款方式': String(a.pay || '現金'), '單號': id, '經手人': me.name });
    }
    return { id, ...eAgentList_() };
  });
}

/* ---------------- 老闆儀表板 ---------------- */

function eDashboard_() {
  const today = eFmt_(eNow_(), 'yyyy-MM-dd');
  const month = today.slice(0, 7);
  const day = (v) => (v === '' || v === null || v === undefined ? '' : (v instanceof Date ? eFmt_(v, 'yyyy-MM-dd') : String(v).slice(0, 10)));

  // 門市結帳（ERP 銷貨）
  const store = { revenue: 0, cost: 0, orders: 0, monthRevenue: 0, monthCost: 0 };
  const ids = {};
  eTable_(E_SHEET.SALES).rows.forEach(r => {
    const d = day(r['時間']);
    if (d.slice(0, 7) === month) { store.monthRevenue += eNum_(r['小計']); store.monthCost += eNum_(r['成本小計']); }
    if (d !== today) return;
    store.revenue += eNum_(r['小計']); store.cost += eNum_(r['成本小計']); ids[r['單號']] = 1;
  });
  store.orders = Object.keys(ids).length;

  // 代購
  const agents = eTable_(E_SHEET.AGENT).rows.map(eAgentRow_);
  const agentOpen = agents.filter(a => AGENT_OPEN.indexOf(a.status) >= 0);
  const agentCount = {}; AGENT_OPEN.forEach(s => { agentCount[s] = agentOpen.filter(a => a.status === s).length; });
  const agentFeeToday = eTable_(E_SHEET.LEDGER).rows
    .filter(r => r['分類'] === '代購服務費' && day(r['日期']) === today).reduce((s, r) => s + eNum_(r['金額']), 0);

  // 庫存不足
  const low = eTable_(E_SHEET.PRODUCTS).rows
    .filter(r => r['啟用'] !== false && eNum_(r['安全庫存']) > 0 && eNum_(r['庫存']) <= eNum_(r['安全庫存']))
    .map(r => ({ name: String(r['品名']), stock: eNum_(r['庫存']), safety: eNum_(r['安全庫存']) }));

  // LINE 線上訂購（蓁便誼訂單系統）
  const online = { ok: true, newToday: 0, newTodayAmount: 0, pickedToday: 0, pickedTodayAmount: 0,
                   toConfirm: 0, ready: [], overdue: [] };
  try {
    const sh = SpreadsheetApp.openById(eProp_('ORDER_SHEET_ID', ORDER_SHEET_DEFAULT)).getSheetByName('訂單');
    const v = sh.getDataRange().getValues();
    const h = v.shift().map(String);
    const c = (name) => h.indexOf(name);
    v.forEach(r => {
      const st = String(r[c('狀態')] || ''), amt = eNum_(r[c('總金額')]);
      const placed = day(r[c('下單時間')]), pick = day(r[c('取貨日期')]);
      const o = { id: String(r[c('訂單編號')]), name: String(r[c('姓名')] || ''), amount: amt, pickup: pick, status: st,
                  items: String(r[c('品項')] || '').split('\n').join('、') };
      if (!o.id) return;
      if (placed === today && st !== '已取消') { online.newToday++; online.newTodayAmount += amt; }
      if (st === '已取貨' && pick === today) { online.pickedToday++; online.pickedTodayAmount += amt; }
      if (st === '待確認') online.toConfirm++;
      const open = st !== '已取貨' && st !== '已取消';
      if (open && pick && pick < today) online.overdue.push(o);
      else if (st === '可取貨') online.ready.push(o);
    });
  } catch (err) {
    online.ok = false; online.error = String(err && err.message || err);
  }

  return {
    date: today,
    store: { revenue: eRound_(store.revenue), gross: eRound_(store.revenue - store.cost), orders: store.orders,
             monthRevenue: eRound_(store.monthRevenue), monthGross: eRound_(store.monthRevenue - store.monthCost) },
    online,
    total: { revenue: eRound_(store.revenue + online.pickedTodayAmount + agentFeeToday), orders: store.orders + online.newToday },
    agent: { open: agentOpen.length, count: agentCount, feeToday: eRound_(agentFeeToday), list: agentOpen.slice(0, 15) },
    low,
  };
}
