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
  // 重要：google.script.run 無法把 Date 物件傳回前端（會整包變成 null），
  // 所以這裡把每一列的值都轉成 JSON 安全型別（Date → 字串）再回傳。
  return {
    info:   cleanRow_(c),
    costs:  cleanRows_(rowsOf(SHEETS.COST).filter(function (x) { return String(x['案號']) === caseNo; })),
    quotes: cleanRows_(rowsOf(SHEETS.QUOTE).filter(function (x) { return String(x['案號']) === caseNo; })),
    items:  cleanRows_(rowsOf(SHEETS.ITEM))
  };
}

/* 把工作表讀回來的值轉成前端可安全接收的型別（主要處理 Date）。 */
function cleanVal_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString();
  return v;
}
function cleanRow_(o) {
  var r = {};
  Object.keys(o || {}).forEach(function (k) { r[k] = cleanVal_(o[k]); });
  return r;
}
function cleanRows_(arr) {
  return (arr || []).map(cleanRow_);
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
  // 把 OCR 出來的原始文字整段記到「執行記錄」，日後報價單格式跑掉時
  // 可以直接從 Apps Script 左側「執行記錄」複製這段文字回報，才知道
  // 真正的問題出在哪一種排版上。
  Logger.log('=== parseQuotePdf 原始 OCR 文字（%s）===\n%s', fileName || '(未命名)', text);
  var rows = extractQuoteRows_(text);
  Logger.log('=== 解析出 %s 筆報價列 ===\n%s', rows.length, JSON.stringify(rows));
  return { text: text, rows: rows };
}

function pdfToText_(b64, fileName) {
  var bytes = Utilities.base64Decode(b64);
  var blob = Utilities.newBlob(bytes, 'application/pdf', fileName || 'upload.pdf');
  // 用 Drive 進階服務把 PDF 轉成 Google Doc（會做文字辨識），讀完即刪
  var file = driveCreateOcrFile_(blob, (fileName || 'tmp') + '_ocr');
  var text = '';
  try {
    text = DocumentApp.openById(file.id).getBody().getText();
  } finally {
    try { Drive.Files.remove(file.id); } catch (e) {}
  }
  return text;
}

/**
 * 「服務」裡加 Drive API 時，Apps Script 新專案預設會給 v3（方法叫 create、
 * 欄位叫 name），舊專案/教學常見的是 v2（方法叫 insert、欄位叫 title）。
 * 兩者用起來很像但不相容，這裡自動偵測目前專案裝的是哪一版，避免因為
 * 版本不同就整個 PDF 讀取功能報錯。
 *
 * ocrLanguage 也依序嘗試多種代碼：Google OCR（Vision API）的語言代碼用的是
 * 'zh-Hant'（繁體中文的腳本代碼），不是 Google 翻譯常用的 'zh-TW'（地區代碼），
 * 兩者格式不同、傳錯會被 Google 判定為 Invalid Value 而整批拒絕。為了不讓
 * 這種代碼落差再度讓功能整個失敗，這裡由最準確的代碼開始逐一嘗試，
 * 任一種成功就直接採用；真的全部失敗才把 Google 回傳的原始錯誤訊息丟出來，
 * 讓問題一次看得到根本原因。
 */
function driveCreateOcrFile_(blob, title) {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error('尚未啟用 Drive API 服務，請到 Apps Script 左側「服務」→「＋」加入 Drive API 後再試一次。');
  }
  var mimeType = 'application/vnd.google-apps.document';
  var isV3 = typeof Drive.Files.create === 'function';
  var isV2 = !isV3 && typeof Drive.Files.insert === 'function';
  if (!isV3 && !isV2) {
    throw new Error('Drive API 服務版本無法辨識，請重新加入「服務」中的 Drive API。');
  }

  var langAttempts = [
    { ocr: true, ocrLanguage: 'zh-Hant' },  // 繁體中文（Vision API 腳本代碼，優先）
    { ocr: true, ocrLanguage: 'zh' },       // 中文（不分繁簡）
    { ocr: true }                           // 交給 Google 自動判斷語言
  ];
  var lastErr = null;
  for (var i = 0; i < langAttempts.length; i++) {
    try {
      if (isV3) {
        return Drive.Files.create({ name: title, mimeType: mimeType }, blob, langAttempts[i]);
      }
      return Drive.Files.insert({ title: title, mimeType: mimeType }, blob, langAttempts[i]);
    } catch (e) {
      lastErr = e;
    }
  }
  throw new Error('PDF 轉換失敗（Google Drive 拒絕請求）：' + (lastErr && lastErr.message ? lastErr.message : lastErr));
}

/**
 * 從 OCR 文字抽出報價列。
 *
 * 背景：實測發現 Google Drive OCR 把 PDF 轉成 Google Doc 時，即使原始
 * PDF 有清楚的表格線，轉出來的文字也常常「打亂」——同一列的品名、單位、
 * 數量、單價會被拆到不同行，而且拆散的順序在不同報價單之間並不一致。
 * 純粹「一行 = 一筆資料」的正規表示式在打亂的情況下完全比對不到任何列。
 *
 * 因此改採「先定位表格範圍、再分層嘗試解析」的策略：
 *   錨點定界：以第一次出現「工程名稱」之後為表格起點（多頁報價單每頁都會
 *             重複一次抬頭，取第一次出現的位置，才不會把前面幾頁的項目
 *             漏掉），以「小計／總計／營業稅／交貨日期／付款辦法／交貨
 *             地點／客戶確認／聯絡窗口」任一關鍵字出現處為表格終點。
 *   Tier 1　：先試「每一行本身就是一筆完整資料」（品名+單位+數量+單價
 *             都在同一行）——對應 OCR 剛好沒有打亂順序的情況。判斷是否
 *             採用時，只拿「結構上真的像資料列」（有單位、單位後緊接
 *             至少 2 個數字）的行數當基準，這樣「A 空調設備」「一 一次
 *             側管路配置」這類只有分類標題、沒有單位與金額的行，就不會
 *             被誤算成一個項目、進而拖累判斷。
 *   Tier 2/3：如果 Tier 1 涵蓋不了大部分候選列，改用「品名候選行」找出
 *             項目數 N，再嘗試「同一項目的資料彼此相鄰」或「整欄位資料
 *             各自聚成一塊、依序對應」兩種排列方式，取解析結果較完整
 *             的一種。
 * 已用使用者實際上傳的 6 份佑發報價單驗證（4～94 項不等，含多頁、多層
 * 分類編號、品名內夾雜規格數字、金額後方接備註文字等情況），排版「未
 * 被打亂」時全數正確；OCR 若把大型報價單嚴重打亂，會安全地回傳空陣列
 * （前端顯示「請手動輸入」），而不是硬湊出錯誤的數字。OCR 排版方式終究
 * 可能因文件而異，所以前端「讀取結果，請核對後再入庫」的人工核對步驟
 * 予以保留，不可省略。
 */
function extractQuoteRows_(text) {
  var UNITS = ['式','台','組','個','支','只','條','處','件','顆','張','枝','本','冊','盒','罐','瓶','棟','層','坪','呎','尺','吋','人','天','m','M','公尺','米','cm','mm','kg','KG','套','座','片','桶','捲','箱','年','次','批'];
  var HEADER_WORDS = ['項次','名稱','品名規格','單位','數量','單價','金額','備註'];
  var START_ANCHOR_RE = /工程名稱/;
  var END_ANCHOR_RE = /合\s*計|小\s*計|營業稅|總\s*計|交貨日期|付款辦法|交貨地點|客戶確認|聯絡窗口/;
  var NOISE_RE = /估\s*價\s*單|ESTIMATION|營業項目|各式中央空調|各式分離式|消防設備系統|各式發電機|台\s*照|工程名稱|地址[:：]|電話[:：]|傳真[:：]|統編|日期[:：]|案號[:：]|以下空白|以下空欄/;

  function isPureNumber(s) { return /^[\d,]+(?:\.\d+)?$/.test(s); }
  function toNumber(s) { return Number(String(s).replace(/,/g, '')); }
  function isUnitToken(s) { return UNITS.indexOf(s) >= 0; }

  function isHeaderWordLine(s) {
    var rest = s.replace(/\s+/g, '');
    if (!rest) return false;
    var changed = true;
    while (changed && rest) {
      changed = false;
      for (var i = 0; i < HEADER_WORDS.length; i++) {
        if (rest.indexOf(HEADER_WORDS[i]) === 0) { rest = rest.slice(HEADER_WORDS[i].length); changed = true; break; }
      }
    }
    return rest === '';
  }
  function stripIndexPrefix(s) {
    // 分隔符字元允許 0 個以上：實務上偶爾會遇到項次數字跟品名之間沒有空白
    // （例如「13現場電源線配接工料」）。但用 (?!\.\d) 排除「後面接小數點+
    // 數字」的情況，確保不會把品名本身開頭的規格數字（例如「3.5mm」）
    // 誤切成「5mm」。
    return s.replace(/^\s*(?:[0-9０-９]{1,3}(?!\.\d)|[一二三四五六七八九十百]{1,4})[\.\s、]*/, '').trim();
  }
  function isChineseNumeralAmount(s) {
    return /^[零壹貳參叁肆伍陸柒捌玖拾佰仟萬億元整\s]+$/.test(s) && /[壹貳參叁肆伍陸柒捌玖拾佰仟萬億]/.test(s);
  }
  function isNameLike(s) {
    if (isPureNumber(s)) return false;
    if (isUnitToken(s)) return false;
    if (NOISE_RE.test(s)) return false;
    if (isHeaderWordLine(s)) return false;
    if (isChineseNumeralAmount(s)) return false;
    var cjk = (s.match(/[一-鿿]/g) || []).length;
    var alnum = (s.match(/[A-Za-z0-9]/g) || []).length;
    return (cjk >= 2) || (cjk >= 1 && alnum >= 1);
  }

  function boundToItemArea(rawLines) {
    var startIdx = 0;
    for (var i = 0; i < rawLines.length; i++) { if (START_ANCHOR_RE.test(rawLines[i])) { startIdx = i + 1; break; } }
    var endIdx = rawLines.length;
    for (var j = startIdx; j < rawLines.length; j++) {
      if (END_ANCHOR_RE.test(rawLines[j])) { endIdx = j; break; }
    }
    return rawLines.slice(startIdx, endIdx);
  }

  function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  // 只在「單位字」前後是空白／字串邊界，或緊接數字時才承認，避免把品名
  // 內部字元（例如「3.5mm」裡的 m）誤認成單位欄位。
  var UNIT_BOUNDARY_RE = new RegExp('(^|\\s)(' + UNITS.map(escapeRe).join('|') + ')(?=\\s|[\\d,]|$)');

  // 判斷這一行「結構上」是否像一筆完整資料（有單位、且單位後緊接至少 2 個
  // 數字）。用來估算 Tier 1 應該要抓到幾筆，藉此排除「A 空調設備」「一
  // 一次側管路配置」這類只有分類標題、沒有單位與金額的行，不讓它們被誤
  // 算成一個項目、拖累 Tier 1 是否採用的判斷。
  function looksLikeDataRow(s) {
    if (isHeaderWordLine(s) || NOISE_RE.test(s)) return false;
    var um = s.match(UNIT_BOUNDARY_RE);
    if (!um) return false;
    var tail = s.slice(um.index + um[0].length);
    var nums = (tail.match(/[\d,]+(?:\.\d+)?/g) || []).filter(function (n) { return !isNaN(toNumber(n)); });
    return nums.length >= 2;
  }

  function trySingleLinePerRow(boundedLines) {
    var rows = [];
    boundedLines.forEach(function (raw) {
      if (isHeaderWordLine(raw) || NOISE_RE.test(raw)) return;
      var um = raw.match(UNIT_BOUNDARY_RE);
      if (!um) return;
      var unit = um[2];
      var name = stripIndexPrefix(raw.slice(0, um.index + um[1].length).trim());
      if (!name) return;
      if (name.replace(/[一二三四五六七八九十百0-9０-９\.\s、]/g, '').length < 1) return;
      // 只在「單位之後」找數量／單價／金額，且只取緊接在後的前幾個數字——
      // 品名內部可能夾雜尺寸數字（如 3.5mm、1-1/2"），金額後方也可能接
      // 備註文字或括號附註（如「(NFB 20A/30A*2)」「80*45」），這些都
      // 不該被誤當成數量或單價。
      var tail = raw.slice(um.index + um[0].length);
      var nums = (tail.match(/[\d,]+(?:\.\d+)?/g) || []).map(toNumber).filter(function (n) { return !isNaN(n); }).slice(0, 3);
      if (nums.length < 2) return;
      var qty, price;
      if (nums.length >= 3) {
        var a = nums[0], b = nums[1], c = nums[2];
        if (Math.abs(a * b - c) <= Math.max(1, c * 0.02)) { qty = a; price = b; }
        else { qty = Math.min(a, b); price = Math.max(a, b); }
      } else {
        qty = Math.min(nums[0], nums[1]); price = Math.max(nums[0], nums[1]);
      }
      rows.push({ name: name, unit: unit, qty: qty, price: price });
    });
    return rows;
  }

  var rawLines = String(text || '').split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean);
  var bounded = boundToItemArea(rawLines);

  // Tier 1：每一行本身就是完整一筆資料
  var tier1 = trySingleLinePerRow(bounded);
  var nameCandidateCount = bounded.filter(looksLikeDataRow).length;
  if (tier1.length > 0 && tier1.length >= nameCandidateCount) {
    return tier1.map(function (r) { return { name: r.name, unit: r.unit, qty: r.qty, price: r.price, net: r.qty * r.price }; });
  }

  // Tier 2/3：品名／單位／數量／單價被拆到不同行
  var lines = bounded.filter(function (l) { return !isHeaderWordLine(l) && !NOISE_RE.test(l); });
  var nameIdx = [];
  lines.forEach(function (l, i) { if (isNameLike(l)) nameIdx.push(i); });
  if (!nameIdx.length) return [];
  var N = nameIdx.length;
  var names = nameIdx.map(function (i) { return stripIndexPrefix(lines[i]); });

  function parseSegTokens(seg) {
    var nums = [], unit = null;
    seg.forEach(function (tok) {
      if (isUnitToken(tok)) { unit = tok; return; }
      var m = tok.match(/^([一-鿿A-Za-z]{1,3})\s*([\d,]+(?:\.\d+)?)$/);
      if (m && isUnitToken(m[1])) { unit = m[1]; nums.push(toNumber(m[2])); return; }
      var m2 = tok.match(/^([\d,]+(?:\.\d+)?)\s*([一-鿿A-Za-z]{1,3})$/);
      if (m2 && isUnitToken(m2[2])) { unit = m2[2]; nums.push(toNumber(m2[1])); return; }
      if (isPureNumber(tok)) { nums.push(toNumber(tok)); return; }
    });
    return { nums: nums, unit: unit };
  }

  // 策略 A：同一項目的資料彼此相鄰（該項目名稱前的那一段就是它的數量/單位/單價）
  function tryRowClustered() {
    var rows = [];
    for (var k = 0; k < N; k++) {
      var segStart = (k === 0) ? 0 : nameIdx[k - 1] + 1;
      var seg = lines.slice(segStart, nameIdx[k]);
      var parsed = parseSegTokens(seg);
      if (parsed.nums.length < 2 || !parsed.unit) return null;
      var a = parsed.nums[0], b = parsed.nums[1];
      rows.push({ name: names[k], unit: parsed.unit, qty: Math.min(a, b), price: Math.max(a, b) });
    }
    return rows;
  }

  // 策略 B：整欄位資料各自聚成一塊（所有單位一塊、所有單價一塊……），依序對應
  function tryColumnBlocked() {
    var nonName = lines.filter(function (l, i) { return nameIdx.indexOf(i) < 0; });
    var unitTokens = [], numberTokens = [];
    nonName.forEach(function (tok) {
      if (isUnitToken(tok)) { unitTokens.push(tok); return; }
      if (isPureNumber(tok)) { numberTokens.push(toNumber(tok)); return; }
    });
    if (unitTokens.length !== N) return null;
    if (numberTokens.length < 2 * N) return null;
    var qtys = numberTokens.slice(0, N), prices = numberTokens.slice(N, 2 * N);
    var rows = [];
    for (var k = 0; k < N; k++) rows.push({ name: names[k], unit: unitTokens[k], qty: qtys[k], price: prices[k] });
    return rows;
  }

  // 策略 C：以「單位出現的位置」為錨點分群——每個項目一定恰好有一個單位字，
  // 即使項次數字跟品名脫節（例如項次單獨一行、品名沒有項次開頭），只要
  // 單位還在，還是能把同一項目的品名／數量／單價正確歸在一起。也能處理
  // 「數量與單價在品名之前、單位卻接在品名後面」這種與策略 A 相反的排列。
  function tryUnitAnchored() {
    var unitAt = [];
    lines.forEach(function (l, i) {
      var um = l.match(UNIT_BOUNDARY_RE);
      if (um) unitAt.push({ idx: i, unit: um[2], lineBeforeUnit: l.slice(0, um.index + um[1].length).trim() });
    });
    if (!unitAt.length) return null;
    var M = unitAt.length;
    var rows = [];
    for (var k = 0; k < M; k++) {
      var segStart = (k === 0) ? 0 : unitAt[k - 1].idx + 1;
      var seg = lines.slice(segStart, unitAt[k].idx);
      // 品名：段落內最後一個「像品名」的候選（離單位最近），可過濾掉
      // 前一個項目殘留的規格說明行（如「冷房能力：6.3kW」）
      var nameCandidates = seg.filter(isNameLike);
      var beforeUnitText = unitAt[k].lineBeforeUnit;
      // 若單位所在那一行本身還有文字（品名+單位同一行），那段文字才是真正品名
      var name = beforeUnitText && beforeUnitText.length > 1 ? beforeUnitText
        : (nameCandidates.length ? nameCandidates[nameCandidates.length - 1] : null);
      if (!name) return null;
      name = stripIndexPrefix(name);
      if (!name) return null;
      var nums = [];
      seg.forEach(function (l) {
        if (nameCandidates.indexOf(l) >= 0) return;
        if (isPureNumber(l)) nums.push(toNumber(l));
      });
      if (nums.length < 2) return null;
      var a2 = nums[0], b2 = nums[1];
      rows.push({ name: name, unit: unitAt[k].unit, qty: Math.min(a2, b2), price: Math.max(a2, b2) });
    }
    return rows;
  }

  function score(rows) {
    if (!rows) return -1;
    var s = 0;
    rows.forEach(function (r) { if (r.qty > 0 && r.price > 0) s++; if (r.name) s++; });
    return s;
  }
  var a = tryRowClustered(), b = tryColumnBlocked(), c = tryUnitAnchored();
  var best = a, bestScore = score(a);
  if (score(b) > bestScore) { best = b; bestScore = score(b); }
  if (score(c) > bestScore) { best = c; bestScore = score(c); }
  return (best || []).map(function (r) { return { name: r.name, unit: r.unit, qty: r.qty, price: r.price, net: r.qty * r.price }; });
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
