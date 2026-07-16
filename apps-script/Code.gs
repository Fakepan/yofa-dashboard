/**************************************************************
 * 佑發空調 — 獨立施工案件 成本利潤管理 + 報價單系統
 * 後端 Code.gs
 *
 * 資料庫：同一個 Google 試算表內四張工作表
 *   案件、成本明細、報價項目、常用項目
 *
 * 部署：Apps Script 編輯器 → 部署 → 新增部署 → 網頁應用程式
 *   執行身分：我自己    誰可以存取：僅限自己（或指定給老闆的 Google 帳號）
 *
 * ── 想調整的地方都集中在下面 CONFIG 區塊 ──
 **************************************************************/

/* ==================== 可調整設定 ==================== */
var TAX_RATE = 0.05;               // 營業稅 5%（單一稅率）

// 報價單抬頭 / 承辦資訊（改這裡就好，不用動 makeQuotePdf）
var COMPANY = {
  name:    '佑發空調工程服務有限公司',
  nameEn:  'Yo-Fa Air Conditioning Engineering Services Co., Ltd.',
  contact: '費可潘',
  phone:   '0903-757-908',
  email:   'momoqu0810@yoshg.com',
  quoteValidDays: 30,
  footNote: '報價含施工及既有設備測試。'
};
/* =================================================== */

var SS_ID_PROP = 'YOFA_SS_ID';     // 記住資料庫試算表 ID 的屬性鍵

var SHEETS = {
  CASE:  '案件',
  COST:  '成本明細',
  QUOTE: '報價項目',
  ITEM:  '常用項目'
};

/* ---------- Web App 入口 ---------- */
function doGet() {
  ensureDatabase();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('佑發空調 案件成本利潤管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ---------- 並行寫入保護 ----------
 * 兩台電腦同時操作時，用 ScriptLock 避免資料互相覆蓋。
 */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // 最多等 20 秒
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 資料庫初始化 ---------- */
function getSS() {
  var id = PropertiesService.getScriptProperties().getProperty(SS_ID_PROP);
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }
  var ss = SpreadsheetApp.create('佑發空調_案件資料庫');
  PropertiesService.getScriptProperties().setProperty(SS_ID_PROP, ss.getId());
  return ss;
}

function ensureDatabase() {
  var ss = getSS();
  var defs = {};
  defs[SHEETS.CASE]  = ['案號','案名','客戶','類型','狀態','建立日期','備註','預留成本率','預留基準'];
  defs[SHEETS.COST]  = ['案號','項目','分類','數量','單價','小計','記錄日期'];
  defs[SHEETS.QUOTE] = ['案號','品名規格','單位','數量','單價','淨額'];
  defs[SHEETS.ITEM]  = ['品名規格','單位','參考單價','分類'];

  Object.keys(defs).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      var head = defs[name];
      sh.getRange(1, 1, 1, head.length).setValues([head])
        .setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
    }
  });
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('工作表1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  // 相容：舊案件表若缺新欄位，補上表頭
  var cs = ss.getSheetByName(SHEETS.CASE);
  var chead = cs.getRange(1, 1, 1, cs.getLastColumn()).getValues()[0];
  ['預留成本率','預留基準'].forEach(function (col) {
    if (chead.indexOf(col) < 0) {
      var c = cs.getLastColumn() + 1;
      cs.getRange(1, c).setValue(col)
        .setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
      chead.push(col);
    }
  });

  // 首次建立時塞幾筆常用維修項目當範例
  var item = ss.getSheetByName(SHEETS.ITEM);
  if (item.getLastRow() < 2) {
    item.getRange(2, 1, 6, 4).setValues([
      ['R410A 冷媒 補充','kg', 1200, '耗材'],
      ['壓縮機 更換工資','式', 3500, '工資'],
      ['排水管 疏通','式', 1500, '工資'],
      ['室外機 保養清洗','台', 1800, '工資'],
      ['電容 更換（含料）','個',  900, '料工'],
      ['冷媒管 銅管','m',    450, '材料']
    ]);
  }
  return ss.getId();
}

/* ---------- 讀取 ---------- */
function rowsOf(name) {
  var sh = getSS().getSheetByName(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return data.map(function (r) {
    var o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

function num(v) {
  var n = Number(String(v == null ? '' : v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function getDashboard() {
  var cases  = rowsOf(SHEETS.CASE);
  var costs  = rowsOf(SHEETS.COST);
  var quotes = rowsOf(SHEETS.QUOTE);

  // 先把成本/淨額依案號彙總，避免逐案 filter 造成 O(n²)
  var costByNo = {}, netByNo = {};
  costs.forEach(function (x) {
    var no = String(x['案號']);
    costByNo[no] = (costByNo[no] || 0) + num(x['小計']);
  });
  quotes.forEach(function (x) {
    var no = String(x['案號']);
    netByNo[no] = (netByNo[no] || 0) + num(x['淨額']);
  });

  return cases.map(function (c) {
    var no = String(c['案號']);
    var cost = costByNo[no] || 0;
    var net  = netByNo[no] || 0;
    var profit = net - cost;                          // 帳面利潤
    var rate = num(c['預留成本率']);                  // 例：10 代表 10%
    var base = c['預留基準'] === '報價淨額' ? net : cost;
    var reserve = Math.round(base * rate / 100);      // 預留成本金額
    var netProfit = profit - reserve;                 // 純利潤（可發獎金）
    var margin = net > 0 ? profit / net : 0;
    var netMargin = net > 0 ? netProfit / net : 0;
    return {
      caseNo: no, name: c['案名'], client: c['客戶'], type: c['類型'],
      status: c['狀態'], date: fmtDate(c['建立日期']),
      cost: cost, net: net, profit: profit, margin: margin,
      reserveRate: rate, reserveBase: c['預留基準'] || '成本',
      reserve: reserve, netProfit: netProfit, netMargin: netMargin
    };
  });
}

function getCaseDetail(caseNo) {
  caseNo = String(caseNo);
  var c = rowsOf(SHEETS.CASE).filter(function (x) { return String(x['案號']) === caseNo; })[0] || {};
  return {
    info:   c,
    costs:  rowsOf(SHEETS.COST).filter(function (x) { return String(x['案號']) === caseNo; }),
    quotes: rowsOf(SHEETS.QUOTE).filter(function (x) { return String(x['案號']) === caseNo; }),
    items:  rowsOf(SHEETS.ITEM)
  };
}

/* ---------- 寫入 ---------- */
function addCase(o) {
  return withLock_(function () {
    var no = String(o.caseNo || '').trim();
    if (!no) throw new Error('案號不可空白');
    var sh = getSS().getSheetByName(SHEETS.CASE);
    // 案號重複則擋下，避免兩台同時新增同一案
    var existing = rowsOf(SHEETS.CASE).some(function (x) { return String(x['案號']) === no; });
    if (existing) throw new Error('案號「' + no + '」已存在');
    sh.appendRow([no, o.name || '', o.client || '', o.type || '', o.status || '進行中',
                  new Date(), o.note || '', '', '']);
    return true;
  });
}

function addCost(o) {
  return withLock_(function () {
    var sh = getSS().getSheetByName(SHEETS.COST);
    var qty = num(o.qty), price = num(o.price);
    sh.appendRow([String(o.caseNo), o.item, o.category || '', qty, price, qty * price, new Date()]);
    return true;
  });
}

function saveQuoteItems(caseNo, items) {
  return withLock_(function () {
    caseNo = String(caseNo);
    var sh = getSS().getSheetByName(SHEETS.QUOTE);
    var last = sh.getLastRow();
    // 先刪掉此案舊的報價列（一次算好要保留的列，整批重寫，避免逐列 deleteRow）
    if (last >= 2) {
      var width = sh.getLastColumn();
      var data = sh.getRange(2, 1, last - 1, width).getValues();
      var keep = data.filter(function (r) { return String(r[0]) !== caseNo; });
      sh.getRange(2, 1, last - 1, width).clearContent();
      if (keep.length) sh.getRange(2, 1, keep.length, width).setValues(keep);
    }
    // 追加本案新的報價列
    var rows = (items || []).map(function (it) {
      var qty = num(it.qty), price = num(it.price);
      return [caseNo, it.name, it.unit || '式', qty, price, qty * price];
    });
    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    }
    return true;
  });
}

function saveReserve(caseNo, rate, base) {
  return withLock_(function () {
    caseNo = String(caseNo);
    var sh = getSS().getSheetByName(SHEETS.CASE);
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var colRate = head.indexOf('預留成本率') + 1;
    var colBase = head.indexOf('預留基準') + 1;
    var last = sh.getLastRow();
    for (var r = 2; r <= last; r++) {
      if (String(sh.getRange(r, 1).getValue()) === caseNo) {
        if (colRate > 0) sh.getRange(r, colRate).setValue(num(rate));
        if (colBase > 0) sh.getRange(r, colBase).setValue(base || '成本');
        return true;
      }
    }
    return false;
  });
}

function deleteCase(caseNo) {
  return withLock_(function () {
    caseNo = String(caseNo);
    [SHEETS.CASE, SHEETS.COST, SHEETS.QUOTE].forEach(function (name) {
      var sh = getSS().getSheetByName(name);
      var last = sh.getLastRow();
      if (last < 2) return;
      var width = sh.getLastColumn();
      var data = sh.getRange(2, 1, last - 1, width).getValues();
      var keep = data.filter(function (r) { return String(r[0]) !== caseNo; });
      sh.getRange(2, 1, last - 1, width).clearContent();
      if (keep.length) sh.getRange(2, 1, keep.length, width).setValues(keep);
    });
    return true;
  });
}

/* ---------- 歷史單價查詢 ---------- */
function searchHistoryPrice(keyword) {
  keyword = String(keyword || '').toLowerCase();
  if (!keyword) return [];
  var out = [];
  rowsOf(SHEETS.QUOTE).forEach(function (q) {
    if (String(q['品名規格']).toLowerCase().indexOf(keyword) >= 0) {
      out.push({ name: q['品名規格'], unit: q['單位'],
                 price: num(q['單價']), caseNo: q['案號'] });
    }
  });
  return out.slice(0, 20);
}

/* ---------- PDF 報價單解析 ---------- */
/**
 * 前端傳來 base64 的 PDF，透過 Drive OCR 轉文字，
 * 依佑發報價單格式解析出項目列，回傳給前端預覽（不直接入庫）。
 */
function parseQuotePdf(b64, fileName) {
  var text = pdfToText_(b64, fileName);
  return { text: text, rows: extractQuoteRows_(text) };
}

function pdfToText_(b64, fileName) {
  var bytes = Utilities.base64Decode(b64);
  var blob = Utilities.newBlob(bytes, 'application/pdf', fileName || 'upload.pdf');
  // 用 Drive 進階服務把 PDF 轉成 Google Doc（會做文字辨識），讀完即刪
  var resource = { title: (fileName || 'tmp') + '_ocr', mimeType: 'application/vnd.google-apps.document' };
  var file = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: 'zh-TW' });
  var text = '';
  try {
    text = DocumentApp.openById(file.id).getBody().getText();
  } finally {
    try { Drive.Files.remove(file.id); } catch (e) {}
  }
  return text;
}

/**
 * 從文字抽出報價列。策略：逐行掃描，抓「品名 … 單位 數量 單價 淨額」樣式，
 * 以「行內至少有兩個數字、且最後一個數字≈前兩個相乘」判斷是報價明細列。
 */
function extractQuoteRows_(text) {
  var lines = String(text || '').split(/\r?\n/);
  var units = ['式','台','組','個','支','只','m','M','公尺','米','kg','KG','套','座','片','桶','捲','箱','年','次'];
  var out = [];
  lines.forEach(function (line) {
    var raw = line.trim();
    if (!raw) return;
    if (/合計|小計|營業稅|總計|報價單|承辦|電話|信箱|項次|品名/.test(raw)) return;

    // 抓出行內所有數字（去逗號）
    var nums = (raw.match(/[\d,]+(?:\.\d+)?/g) || [])
      .map(function (s) { return Number(s.replace(/,/g, '')); })
      .filter(function (n) { return !isNaN(n); });
    if (nums.length < 2) return;

    // 找單位
    var unit = '式';
    for (var i = 0; i < units.length; i++) {
      if (raw.indexOf(units[i]) >= 0) { unit = units[i]; break; }
    }

    // 品名 = 行首到第一個數字/單位前的中文與英文
    var nameMatch = raw.match(/^[^\d]+/);
    var name = nameMatch ? nameMatch[0].replace(/[\s　]+$/,'').trim() : raw;
    // 去掉品名尾端可能夾到的單位字
    units.forEach(function (u) {
      if (name.slice(-u.length) === u) name = name.slice(0, -u.length).trim();
    });
    if (!name) return;

    // 判斷數量/單價/淨額：取最後三個數字嘗試 qty*price≈net
    var qty = 1, price = 0, net = 0;
    if (nums.length >= 3) {
      var a = nums[nums.length - 3], b = nums[nums.length - 2], c = nums[nums.length - 1];
      if (Math.abs(a * b - c) <= Math.max(1, c * 0.02)) { qty = a; price = b; net = c; }
      else { qty = 1; price = b; net = c; }
    } else {
      // 兩個數字：視為 單價 淨額 或 數量 單價
      price = nums[0]; net = nums[1];
      if (Math.abs(price - net) < 0.01) { qty = 1; }
    }
    out.push({ name: name, unit: unit, qty: qty, price: price, net: net });
  });
  return out;
}

/** 把校對後的項目寫進歷史（常用項目表），供日後查詢/帶入 */
function importToHistory(items) {
  return withLock_(function () {
    var sh = getSS().getSheetByName(SHEETS.ITEM);
    var rows = (items || []).filter(function (it) { return it && it.name; })
      .map(function (it) { return [it.name, it.unit || '式', num(it.price), it.category || '匯入']; });
    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    }
    return true;
  });
}

/* ---------- 產生 PDF 報價單 ---------- */
/**
 * 依案號組出報價單 HTML → 轉 PDF，回傳 base64 給前端下載。
 * （抬頭/承辦資訊改上方 COMPANY 設定即可。）
 */
function makeQuotePdf(caseNo) {
  caseNo = String(caseNo);
  var d = getCaseDetail(caseNo);
  var net = d.quotes.reduce(function (s, x) { return s + num(x['淨額']); }, 0);
  var tax = Math.round(net * TAX_RATE);
  var total = net + tax;

  var rows = d.quotes.map(function (q, i) {
    return '<tr>' +
      '<td class="c">' + (i + 1) + '</td>' +
      '<td>' + esc(q['品名規格']) + '</td>' +
      '<td class="c">' + esc(q['單位']) + '</td>' +
      '<td class="r">' + q['數量'] + '</td>' +
      '<td class="r">' + money(q['單價']) + '</td>' +
      '<td class="r">' + money(q['淨額']) + '</td></tr>';
  }).join('');

  var taxPct = Math.round(TAX_RATE * 100);
  var html =
    '<style>' +
    'body{font-family:"Microsoft JhengHei",sans-serif;color:#222;font-size:12px;}' +
    '.hd{background:#1F4E78;color:#fff;padding:14px 18px;}' +
    '.hd h1{margin:0;font-size:20px;letter-spacing:2px;}' +
    '.hd .sub{font-size:11px;opacity:.9;}' +
    '.meta{width:100%;margin:12px 0;font-size:12px;}' +
    '.meta td{padding:2px 4px;}' +
    'table.items{width:100%;border-collapse:collapse;margin-top:8px;}' +
    'table.items th{background:#D9E1F2;border:1px solid #9bb0d0;padding:6px;}' +
    'table.items td{border:1px solid #cbd5e6;padding:6px;}' +
    '.r{text-align:right;} .c{text-align:center;}' +
    '.sum td{padding:4px 8px;}' +
    '.total{font-weight:bold;background:#FFF2CC;}' +
    '.foot{margin-top:24px;font-size:11px;color:#555;}' +
    '</style>' +
    '<div class="hd"><h1>' + esc(COMPANY.name) + '</h1>' +
    '<div class="sub">' + esc(COMPANY.nameEn) + ' ｜ 報 價 單</div></div>' +
    '<table class="meta"><tr>' +
    '<td>案　號：' + esc(caseNo) + '</td>' +
    '<td>日　期：' + fmtDate(new Date()) + '</td></tr>' +
    '<tr><td>客　戶：' + esc(d.info['客戶'] || '') + '</td>' +
    '<td>案　名：' + esc(d.info['案名'] || '') + '</td></tr></table>' +
    '<table class="items"><tr>' +
    '<th>項次</th><th>品名規格</th><th>單位</th><th>數量</th><th>單價</th><th>淨額</th></tr>' +
    rows +
    '</table>' +
    '<table class="meta sum" align="right" style="width:300px;margin-top:10px;">' +
    '<tr><td class="r">未稅小計</td><td class="r">' + money(net) + '</td></tr>' +
    '<tr><td class="r">營業稅 ' + taxPct + '%</td><td class="r">' + money(tax) + '</td></tr>' +
    '<tr class="total"><td class="r">含稅總計</td><td class="r">' + money(total) + '</td></tr>' +
    '</table>' +
    '<div class="foot" style="clear:both;">' +
    '承辦：' + esc(COMPANY.contact) + '　電話 ' + esc(COMPANY.phone) + '　信箱 ' + esc(COMPANY.email) + '<br>' +
    '本報價單有效期 ' + COMPANY.quoteValidDays + ' 天，' + esc(COMPANY.footNote) +
    '</div>';

  var blob = Utilities.newBlob(html, 'text/html', 'quote.html')
                      .getAs('application/pdf')
                      .setName('佑發報價單_' + caseNo + '.pdf');
  return {
    b64: Utilities.base64Encode(blob.getBytes()),
    name: blob.getName()
  };
}

/* ---------- 小工具 ---------- */
function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function money(n) {
  n = Math.round(Number(n || 0));
  return n.toLocaleString('en-US');
}
function fmtDate(d) {
  if (!d) return '';
  if (!(d instanceof Date)) d = new Date(d);
  if (isNaN(d.getTime())) return '';
  var y = d.getFullYear() - 1911; // 民國
  return y + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + ('0' + d.getDate()).slice(-2);
}
