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
 *  4. Bu adresi uygulamanın config.js dosyasına yazın (veya geliştiriciye iletin).
 *
 * Sayfalar:
 *  - Ürün Listesi : sabit Excel ürün listesi; Yeni Barkod / Kontrol Durumu / Kontrol Tarihi / Kontrol Eden
 *                   sütunları her okutmada anında güncellenir. (İlk bağlantıda uygulama listeyi kendisi yükler.)
 *  - Yeni Ürünler : listede olmayan / bulunamayan ürünler.
 *  - Olaylar      : tüm işlemler satır satır (uygulama buradan okur). Satırları elle silmeyin.
 * Yeni bir sayım için menüden "Barkod Kontrol → Yeni sayım başlat" kullanın.
 */

var SHEET_NAME = 'Olaylar';
var PRODUCTS = 'Ürün Listesi';
var NEWPRODS = 'Yeni Ürünler';
var ORIG = '_ilk_yeni_barkod';
var HEADERS = ['Tarih', 'Kişi', 'İşlem', 'Barkod', 'Ürün Adı', 'Stok Kodu', 'Marka', 'Not', 'Olay ID', 'Veri'];
var COL_ID = 9;     // Olay ID
var COL_JSON = 10;  // Veri (uygulamanın okuduğu sütun)
var PAGE = 1000;    // bir istekte gönderilecek en fazla olay

var S = {
  MATCH: 'Eşleşti', NEWBC: 'Yeni Barkod Eklendi', NOTFOUND: 'Bulunamadı', NEWPROD: 'Yeni Ürün', NOMATCH: 'Eşleşmedi'
};
var LABELS = {
  scan: 'Okutma', newbc: 'Yeni Barkod Eklendi', undo_newbc: 'Barkod Geri Alındı',
  notfound: 'Bulunamadı', newprod: 'Yeni Ürün', del_newprod: 'Kayıt Silindi'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Barkod Kontrol')
    .addItem('Yeni sayım başlat (mevcut kayıtları arşivle)', 'yeniSayimBaslat')
    .addItem('Ürün Listesi sayfasını yeniden hesapla', 'yenidenHesapla')
    .addToUi();
}

// ---------------------------------------------------------------------------
// Yardımcılar
// ---------------------------------------------------------------------------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
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

function norm_(s) {
  return String(s == null ? '' : s).trim().toLocaleLowerCase('tr');
}

function cleanCode_(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/\s+/g, '').trim().toUpperCase();
}

/** Uygulamadaki eşleştirme anahtarının aynısı (baştaki sıfırlar önemsiz). */
function codeKey_(v) {
  var s = cleanCode_(v);
  if (/^\d+$/.test(s)) return s.replace(/^0+(?=\d)/, '');
  return s;
}

function splitCodes_(v) {
  if (v === null || v === undefined || v === '') return [];
  return String(v).split(/[,;\n/|]+/).map(cleanCode_).filter(function (x) { return x; });
}

function uniq_(arr) {
  var seen = {}, out = [];
  for (var i = 0; i < arr.length; i++) if (!seen[arr[i]]) { seen[arr[i]] = true; out.push(arr[i]); }
  return out;
}

function readEvents_() {
  var sh = getSheet_();
  var total = Math.max(0, sh.getLastRow() - 1);
  if (!total) return [];
  var vals = sh.getRange(2, COL_JSON, total, 1).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    try { out.push(JSON.parse(vals[i][0])); } catch (err) { /* atla */ }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Web uygulaması
// ---------------------------------------------------------------------------

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
    return json_({ ok: true, epoch: epoch_(), next: next, more: next < total, events: events,
      products: !!ss_().getSheetByName(PRODUCTS) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/**
 * Uygulama gönderir:
 *   { epoch, events: [...] }                        yeni olaylar
 *   { action: 'products', headerRow, rows: [[...]] } ürün listesinin ilk yüklenmesi
 */
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (body.action === 'products') return json_(uploadProducts_(body));

    var ep = epoch_();
    if (body.epoch && body.epoch !== ep) return json_({ ok: false, error: 'epoch', epoch: ep });
    var incoming = (body.events || []).slice(0, 500);
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
    if (rows.length) {
      sh.getRange(lastRow + 1, 1, rows.length, HEADERS.length).setValues(rows);
      rebuild_();
    }
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

// ---------------------------------------------------------------------------
// Ürün Listesi sayfası
// ---------------------------------------------------------------------------

/** Uygulama, ilk bağlandığında sabit Excel listesini buraya yükler (yalnızca bir kez). */
function uploadProducts_(body) {
  var ss = ss_();
  if (ss.getSheetByName(PRODUCTS)) return { ok: true, existed: true };
  var rows = body.rows || [];
  var hr = parseInt(body.headerRow, 10) || 0;
  if (!rows.length || !rows[hr]) return { ok: false, error: 'bos_liste' };
  var width = 0;
  for (var i = 0; i < rows.length; i++) width = Math.max(width, rows[i].length);
  var data = rows.map(function (r) {
    var out = [];
    for (var c = 0; c < width; c++) out.push(r[c] === null || r[c] === undefined ? '' : r[c]);
    return out;
  });
  var sh = ss.insertSheet(PRODUCTS, 0);
  sh.getRange(1, 1, data.length, width).setValues(data);
  sh.getRange(hr + 1, 1, 1, width).setFontWeight('bold');
  sh.setFrozenRows(hr + 1);
  PropertiesService.getScriptProperties().setProperty('headerRow', String(hr));

  // Excel'deki mevcut "Yeni Barkod" değerlerini sakla (sonradan eklenenlerle birleştirilir)
  var cols = productCols_(sh, hr);
  var orig = ss.insertSheet(ORIG);
  var vals = data.map(function (r) { return [cols.newbc >= 0 ? String(r[cols.newbc] || '') : '']; });
  orig.getRange(1, 1, vals.length, 1).setValues(vals);
  orig.hideSheet();

  rebuild_();
  return { ok: true, created: true };
}

/** Ürün Listesi'nde güncellenen sütunları bulur, yoksa sona ekler (0 tabanlı indeks). */
function productCols_(sh, hr) {
  var lastCol = sh.getLastColumn();
  var header = sh.getRange(hr + 1, 1, 1, lastCol).getValues()[0].map(norm_);
  function col(name) {
    var i = header.indexOf(norm_(name));
    if (i >= 0) return i;
    lastCol += 1;
    sh.getRange(hr + 1, lastCol, 1, 1).setValues([[name]]).setFontWeight('bold');
    header.push(norm_(name));
    return lastCol - 1;
  }
  return { newbc: col('Yeni Barkod'), status: col('Kontrol Durumu'), date: col('Kontrol Tarihi'), by: col('Kontrol Eden') };
}

/** Tüm olayları sırayla uygulayıp Ürün Listesi ve Yeni Ürünler sayfalarını günceller (uygulamadaki hesabın aynısı). */
function rebuild_() {
  var ss = ss_();
  var pSh = ss.getSheetByName(PRODUCTS);
  var events = readEvents_();
  events.sort(function (a, b) { return a.t < b.t ? -1 : a.t > b.t ? 1 : (a.id < b.id ? -1 : 1); });

  var added = {}, checked = {}, np = [];
  function findNp(pid, code) {
    var i;
    if (pid) for (i = 0; i < np.length; i++) if (np[i].id === pid) return np[i];
    if (code) for (i = 0; i < np.length; i++) if (codeKey_(np[i].barcode) === codeKey_(code)) return np[i];
    return null;
  }
  events.forEach(function (ev) {
    var by = ev.by || '';
    switch (ev.type) {
      case 'scan':
        if (ev.res === S.MATCH) (ev.rows || []).forEach(function (r) { checked[r] = { s: ev.via === 'new' ? S.NEWBC : S.MATCH, t: ev.t, by: by }; });
        break;
      case 'newbc':
        added[ev.r] = uniq_((added[ev.r] || []).concat([ev.code]));
        checked[ev.r] = { s: S.NEWBC, t: ev.t, by: by };
        break;
      case 'undo_newbc': {
        var left = (added[ev.r] || []).filter(function (c) { return c !== ev.code; });
        if (left.length) added[ev.r] = left;
        else {
          delete added[ev.r];
          if (checked[ev.r] && checked[ev.r].s === S.NEWBC) delete checked[ev.r];
        }
        break;
      }
      case 'notfound':
        if (!findNp(null, ev.code)) np.push({ id: ev.pid || ev.id, barcode: ev.code, name: '', brand: '', note: '', status: S.NOTFOUND, t: ev.t, by: by });
        break;
      case 'newprod': {
        var n = findNp(ev.pid, ev.prev || ev.code);
        if (!n) { n = { id: ev.pid || ev.id, t: ev.t }; np.push(n); }
        n.barcode = ev.code; n.name = ev.name || ''; n.brand = ev.brand || ''; n.note = ev.note || ''; n.status = S.NEWPROD; n.by = by;
        break;
      }
      case 'del_newprod': {
        var d = findNp(ev.pid, ev.code);
        if (d) np = np.filter(function (x) { return x !== d; });
        break;
      }
    }
  });
  var addedKeys = {};
  Object.keys(added).forEach(function (r) { added[r].forEach(function (c) { addedKeys[codeKey_(c)] = true; }); });
  np = np.filter(function (n) { return !(n.status === S.NOTFOUND && addedKeys[codeKey_(n.barcode)]); });

  // --- Ürün Listesi
  if (pSh) {
    var hr = parseInt(PropertiesService.getScriptProperties().getProperty('headerRow') || '0', 10);
    var cols = productCols_(pSh, hr);
    var last = pSh.getLastRow();
    var n = last - (hr + 1);
    if (n > 0) {
      var origSh = ss.getSheetByName(ORIG);
      var orig = origSh ? origSh.getRange(1, 1, Math.max(1, origSh.getLastRow()), 1).getValues() : [];
      var newbc = [], status = [], date = [], who = [];
      for (var i = 0; i < n; i++) {
        var r = hr + 1 + i; // olaylardaki satır no (0 tabanlı)
        var base = splitCodes_(orig[r] ? orig[r][0] : '');
        var codes = uniq_(base.concat(added[r] || []));
        newbc.push([codes.join(', ')]);
        var c = checked[r];
        status.push([c ? c.s : '']);
        date.push([c ? new Date(c.t) : '']);
        who.push([c ? c.by : '']);
      }
      var top = hr + 2;
      pSh.getRange(top, cols.newbc + 1, n, 1).setNumberFormat('@').setValues(newbc);
      pSh.getRange(top, cols.status + 1, n, 1).setValues(status);
      pSh.getRange(top, cols.date + 1, n, 1).setNumberFormat('dd.mm.yyyy hh:mm').setValues(date);
      pSh.getRange(top, cols.by + 1, n, 1).setValues(who);
    }
  }

  // --- Yeni Ürünler
  var nSh = ss.getSheetByName(NEWPRODS);
  if (!nSh) {
    if (!np.length) return;
    nSh = ss.insertSheet(NEWPRODS);
  }
  nSh.clearContents();
  var head = ['Barkod', 'Ürün Adı', 'Marka', 'Durum', 'Not', 'Tarih', 'Ekleyen'];
  var out = [head].concat(np.map(function (x) {
    return ["'" + x.barcode, x.name, x.brand, x.status, x.note, new Date(x.t), x.by || ''];
  }));
  nSh.getRange(1, 1, out.length, head.length).setValues(out);
  nSh.getRange(1, 1, 1, head.length).setFontWeight('bold');
}

// ---------------------------------------------------------------------------
// Menü
// ---------------------------------------------------------------------------

/** Mevcut kayıtları arşivler, tüm cihazlar boş bir sayımla devam eder. */
function yeniSayimBaslat() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Yeni sayım başlatılsın mı?',
    'Mevcut "Olaylar" ve "Ürün Listesi" sayfalarının kopyası arşivlenir; ürün listesindeki kontrol bilgileri temizlenir ve tüm cihazlar sıfırlanır.',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = ss_();
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH.mm');
    var p = ss.getSheetByName(PRODUCTS);
    if (p) p.copyTo(ss).setName('Arşiv Liste ' + stamp);
    var sh = ss.getSheetByName(SHEET_NAME);
    if (sh) sh.setName('Arşiv Olaylar ' + stamp);
    PropertiesService.getScriptProperties().setProperty('epoch', Utilities.getUuid());
    getSheet_();
    rebuild_();
  } finally {
    lock.releaseLock();
  }
  ui.alert('Yeni sayım başlatıldı. Uygulamalar birkaç saniye içinde sıfırlanır.');
}

function yenidenHesapla() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { rebuild_(); } finally { lock.releaseLock(); }
}
