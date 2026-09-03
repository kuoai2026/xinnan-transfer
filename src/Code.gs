/**
 * 信男藥局 · 倉庫調貨單 —— 後端 Web App
 * 綁在「藥局調貨表單」試算表 (1em9eYPdK8ErWPDH_kVJwx8cT9zk-qiBjIW2_qt4TogY)
 *
 * 部署：部署 → 新增部署作業 → 類型「網頁應用程式」
 *        執行身分：我
 *        誰可以存取：任何人
 *   複製 /exec 網址，填進 transfer_page/config.json 的 "endpoint"
 *
 * 首次請先在編輯器手動執行一次 setup()（會建立兩個分頁並授權寄信）。
 *
 * 舊的「表單回應」分頁完全不會被動到。
 */

var TOKEN = 'phY_sOCvZQ-aX2kaKwEQS4UHjHoOHyzV';   // 與 config.json 一致
var MAIL_TO = '174kuo@gmail.com';
var TZ = 'Asia/Taipei';

var SH_MAIN = '調貨單_主檔';
var SH_LINE = '調貨單_明細';
var HEAD_MAIN = ['單號','時間戳記','填寫日期','填寫人員','品項數','總數量','客訂項數',
                 '整單備註','文字摘要','系統品項數','手填品項數','clientId','狀態'];
var HEAD_LINE = ['單號','時間戳記','填寫人員','類型','貨號','品名','系列','數量','客訂','備註'];

function token_() {
  var p = PropertiesService.getScriptProperties().getProperty('TOKEN');
  return p || TOKEN;
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SH_MAIN, HEAD_MAIN);
  ensureSheet_(ss, SH_LINE, HEAD_LINE);
  // 觸發寄信授權
  try { MailApp.getRemainingDailyQuota(); } catch (e) {}
  Logger.log('setup done: ' + SH_MAIN + ' / ' + SH_LINE);
}

function ensureSheet_(ss, name, head) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
  return sh;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- POST：門市送出一張調貨單 ---------------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents || '{}');
    if (body.token !== token_()) return jsonOut_({ ok: false, error: '驗證失敗' });

    var person = String(body.person || '').trim();
    if (!person) return jsonOut_({ ok: false, error: '缺少填寫人員' });

    var items = Array.isArray(body.items) ? body.items : [];
    var manual = Array.isArray(body.manualItems) ? body.manualItems : [];
    if (!items.length && !manual.length) return jsonOut_({ ok: false, error: '清單是空的' });

    var ss = SpreadsheetApp.getActive();
    var main = ensureSheet_(ss, SH_MAIN, HEAD_MAIN);
    var line = ensureSheet_(ss, SH_LINE, HEAD_LINE);

    var now = new Date();
    var ts = Utilities.formatDate(now, TZ, 'yyyy/MM/dd HH:mm:ss');
    var day = Utilities.formatDate(now, TZ, 'yyyy/MM/dd');
    var id = nextId_(main);

    var totalQty = 0, customCount = 0;
    var lineRows = [];
    items.forEach(function (it) {
      var q = Number(it.qty) || 0; totalQty += q;
      if (it.custom) customCount++;
      lineRows.push([id, ts, person, '系統', it.sku || '', it.name || '', it.series || '',
                     q, it.custom ? '客訂' : '', it.note || '']);
    });
    manual.forEach(function (it) {
      var q = Number(it.qty) || 0; totalQty += q;
      if (it.custom) customCount++;
      var nm = (it.name || '') + (it.unit ? '（' + it.unit + '）' : '');
      lineRows.push([id, ts, person, '手填', '', nm, '', q, it.custom ? '客訂' : '', it.note || '']);
    });

    var summary = String(body.summary || '');
    main.appendRow([id, ts, day, person, items.length + manual.length, totalQty, customCount,
                    String(body.ordNote || ''), summary, items.length, manual.length,
                    String(body.clientId || ''), '']);
    if (lineRows.length)
      line.getRange(line.getLastRow() + 1, 1, lineRows.length, HEAD_LINE.length).setValues(lineRows);

    sendMail_(id, person, ts, items, manual, summary, String(body.ordNote || ''), totalQty);

    return jsonOut_({ ok: true, id: id, ts: now.getTime() });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function nextId_(main) {
  var last = main.getLastRow();
  if (last < 2) return 1001;
  var v = main.getRange(last, 1).getValue();
  var n = parseInt(String(v).replace(/[^0-9]/g, ''), 10);
  return (isNaN(n) ? 1000 : n) + 1;
}

function sendMail_(id, person, ts, items, manual, summary, ordNote, totalQty) {
  try {
    var ssUrl = SpreadsheetApp.getActive().getUrl();
    var n = items.length + manual.length;
    var subject = '🚚 調貨單 #' + id + ' · ' + person + ' · ' + n + ' 項';
    var lines = [];
    lines.push('填寫人員：' + person);
    lines.push('時間：' + ts);
    lines.push('項目數：' + n + '　總數量：' + totalQty);
    if (ordNote) lines.push('整單備註：' + ordNote);
    lines.push('');
    lines.push('──── 撿貨摘要 ────');
    lines.push(summary || '(無)');
    lines.push('');
    lines.push('──── 系統品項（含貨號）────');
    items.forEach(function (it) {
      lines.push('• ' + (it.name || '') + '  ×' + (it.qty || 0)
        + (it.custom ? '  [客訂]' : '') + (it.note ? '  (' + it.note + ')' : '')
        + '   ' + (it.sku || ''));
    });
    if (manual.length) {
      lines.push('');
      lines.push('──── 系統沒有的商品（門市手填）────');
      manual.forEach(function (it) {
        lines.push('• ' + (it.name || '') + '  ×' + (it.qty || 0) + (it.unit ? ' ' + it.unit : '')
          + (it.custom ? '  [客訂]' : '') + (it.note ? '  (' + it.note + ')' : ''));
      });
    }
    lines.push('');
    lines.push('明細表：' + ssUrl);
    MailApp.sendEmail(MAIL_TO, subject, lines.join('\n'));
  } catch (err) {
    // 寄信失敗不影響送單
    console.error('sendMail_ failed: ' + err);
  }
}

/* ---------------- GET：門市查自己最近送出的單 ---------------- */
function doGet(e) {
  var p = e.parameter || {};
  if (p.token !== token_()) return jsonOut_({ ok: false, error: '驗證失敗' });
  if (!p.recent) return jsonOut_({ ok: true, alive: true });

  var main = SpreadsheetApp.getActive().getSheetByName(SH_MAIN);
  if (!main || main.getLastRow() < 2) return jsonOut_({ ok: true, submissions: [] });

  var cid = String(p.clientId || '');
  var last = main.getLastRow();
  var take = Math.min(200, last - 1);
  var rows = main.getRange(last - take + 1, 1, take, HEAD_MAIN.length).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0 && out.length < 8; i--) {
    var r = rows[i];
    if (cid && String(r[11]) !== cid) continue;
    out.push({
      id: r[0],
      time: Utilities.formatDate(new Date(r[1]), TZ, 'MM/dd HH:mm'),
      person: r[3],
      count: r[4],
      summary: r[8],
      status: r[12] || ''
    });
  }
  return jsonOut_({ ok: true, submissions: out });
}
