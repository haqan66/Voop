/**
 * Barkod Kontrol — ortak liste (Google E-Tablosu + Apps Script)
 *
 * Kurulum (bir kez):
 *  1. Yeni bir Google E-Tablosu açın (ör. "Barkod Kontrol Ortak").
 *  2. Uzantılar → Apps Script. Açılan dosyadaki her şeyi silip bu dosyanın tamamını yapıştırın, kaydedin.
 *  3. Dağıt → Yeni dağıtım → tür: "Web uygulaması"
 *       Şu kullanıcı olarak yürüt: Ben
 *       Erişimi olanlar: Herkes
 *     → Dağıt → (izinleri onaylayın) → "Web uygulaması URL'si"ni kopyalayın (…/exec ile biter).
 *  4. Barkod Kontrol uygulamasında Kayıtlar → Ortak Çalışma → adresi yapıştırıp "Bağlan".
 *
 * Tüm okutmalar "Olaylar" sayfasına satır satır yazılır. Bu sayfadaki satırları elle silmeyin;
 * yeni bir sayım için menüden "Barkod Kontrol → Yeni sayım başlat" kullanın.
 */

var SHEET_NAME = 'Olaylar';
var HEADERS = ['Tarih', 'Kişi', 'İşlem', 'Barkod', 'Ürün Adı', 'Stok Kodu', 'Marka', 'Not', 'Olay ID', 'Veri'];
var COL_ID = 9;     // Olay ID
var COL_JSON = 10;  // Veri (uygulamanın okuduğu sütun)
var PAGE = 1000;    // bir istekte gönderilecek en fazla olay

var LABELS = {
  scan: 'Okutma',
  newbc: 'Yeni Barkod Eklendi',
  undo_newbc: 'Barkod Geri Alındı',
  notfound: 'Bulunamadı',
  newprod: 'Yeni Ürün',
  del_newprod: 'Kayıt Silindi'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Barkod Kontrol')
    .addItem('Yeni sayım başlat (mevcut kayıtları arşivle)', 'yeniSayimBaslat')
    .addToUi();
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME, 0);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(5, 320);
    sh.hideColumns(COL_JSON);
  }
  return sh;
}

function epoch_() {
  var props = PropertiesService.getScriptProperties();
  var e = props.getProperty('epoch');
  if (!e) {
    e = Utilities.getUuid();
    props.setProperty('epoch', e);
  }
  return e;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Uygulama yeni olayları okur: ?since=<kaçıncı olaydan sonra> */
function doGet(e) {
  try {
    var sh = getSheet_();
    var since = parseInt((e && e.parameter && e.parameter.since) || '0', 10);
    if (!(since >= 0)) since = 0;
    var total = Math.max(0, sh.getLastRow() - 1);
    var events = [];
    var next = total;
    if (total > since) {
      var n = Math.min(PAGE, total - since);
      var vals = sh.getRange(2 + since, COL_JSON, n, 1).getValues();
      for (var i = 0; i < vals.length; i++) {
        try { events.push(JSON.parse(vals[i][0])); } catch (err) { /* bozuk satırı atla */ }
      }
      next = since + n;
    }
    return json_({ ok: true, epoch: epoch_(), next: next, more: next < total, events: events });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Uygulama yeni olayları gönderir: { epoch, events: [...] } */
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  var ep = epoch_();
  if (body.epoch && body.epoch !== ep) return json_({ ok: false, error: 'epoch', epoch: ep });
  var incoming = (body.events || []).slice(0, 500);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = getSheet_();
    var lastRow = sh.getLastRow();
    var seen = {};
    if (lastRow > 1) {
      // aynı olay tekrar gönderilirse (bağlantı kopması vb.) iki kez yazılmasın
      var start = Math.max(2, lastRow - 4999);
      var ids = sh.getRange(start, COL_ID, lastRow - start + 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) seen[ids[i][0]] = true;
    }
    var rows = [];
    for (var j = 0; j < incoming.length; j++) {
      var ev = incoming[j];
      if (!ev || !ev.id || seen[ev.id]) continue;
      seen[ev.id] = true;
      rows.push(row_(ev));
    }
    if (rows.length) sh.getRange(lastRow + 1, 1, rows.length, HEADERS.length).setValues(rows);
    return json_({ ok: true, epoch: ep, added: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function row_(ev) {
  var label = LABELS[ev.type] || ev.type;
  if (ev.type === 'scan') label = ev.res || label;
  var d = new Date(ev.t);
  return [
    isNaN(d.getTime()) ? new Date() : d,
    ev.by || '',
    label,
    ev.code ? "'" + ev.code : '',
    ev.name || '',
    ev.stok ? "'" + ev.stok : '',
    ev.brand || '',
    ev.note || '',
    ev.id,
    JSON.stringify(ev)
  ];
}

/** Menü: mevcut kayıtları arşivler, tüm cihazlar boş bir sayımla devam eder. */
function yeniSayimBaslat() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Yeni sayım başlatılsın mı?',
    'Mevcut "Olaylar" sayfası arşivlenir ve tüm cihazlardaki kayıtlar sıfırlanır. Önce uygulamadan Excel indirmeniz önerilir.',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (sh) sh.setName('Arşiv ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH.mm'));
    PropertiesService.getScriptProperties().setProperty('epoch', Utilities.getUuid());
    getSheet_();
  } finally {
    lock.releaseLock();
  }
  ui.alert('Yeni sayım başlatıldı. Uygulamalar birkaç saniye içinde sıfırlanır.');
}
