/* Barkod Kontrol — mobil barkod eşleştirme uygulaması
 * Akış: Excel ürün listesini oku → kamerayla barkod okut → listede var mı kontrol et
 *   → yoksa ürün adıyla ara → bulunursa barkodu "Yeni Barkod" sütununa ekle
 *   → bulunmazsa "Bulunamadı" işaretle ve yeni ürün olarak kaydet.
 */
(function () {
  'use strict';

  const DEFAULT_FILE = 'data/urun-listesi.xlsx';
  const SHEET_NEW = 'Yeni Ürünler';
  const SHEET_LOG = 'Tarama Geçmişi';
  const COL_NEW = 'Yeni Barkod';
  const COL_STATUS = 'Kontrol Durumu';
  const COL_DATE = 'Kontrol Tarihi';
  const SAME_CODE_COOLDOWN = 2500; // aynı barkodu tekrar saymamak için (ms)

  const S = { // durum etiketleri (Excel'e de bu metinler yazılır)
    MATCH: 'Eşleşti',
    NEWBC: 'Yeni Barkod Eklendi',
    NOTFOUND: 'Bulunamadı',
    NEWPROD: 'Yeni Ürün',
    NOMATCH: 'Eşleşmedi',
  };

  // ------------------------------------------------------------------
  // Yardımcılar
  // ------------------------------------------------------------------
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nowIso = () => new Date().toISOString();
  const fmtDate = (iso) => {
    if (!iso) return '';
    const d = new Date(iso);
    const p = (n) => String(n).padStart(2, '0');
    return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
  };
  const fmtTime = (iso) => {
    const d = new Date(iso);
    return d.toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' });
  };

  /** Türkçe karakterleri sadeleştirip arama için normalleştirir. */
  function normText(s) {
    return String(s ?? '')
      .toLocaleLowerCase('tr')
      .replace(/ı/g, 'i').replace(/ş/g, 's').replace(/ğ/g, 'g')
      .replace(/ü/g, 'u').replace(/ö/g, 'o').replace(/ç/g, 'c')
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
  }

  /** Barkod metnini temizler (boşluk vb. atar). */
  function cleanCode(v) {
    if (v === null || v === undefined) return '';
    let s = typeof v === 'number' ? (Number.isInteger(v) ? String(v) : v.toFixed(0)) : String(v);
    return s.replace(/\s+/g, '').trim().toUpperCase();
  }

  /** Eşleştirme anahtarı: sadece rakamsa baştaki sıfırlar atılır (UPC-A ↔ EAN-13 uyumu). */
  function codeKey(v) {
    const s = cleanCode(v);
    if (!s) return '';
    if (/^\d+$/.test(s)) return s.replace(/^0+(?=\d)/, '');
    return s;
  }

  /** "8681812443401-1" veya "8681812443036T" gibi kodlar için temel barkod. */
  function baseKey(v) {
    const s = cleanCode(v);
    const m = s.match(/^0*(\d{8,14})(?:[-_./][A-Z0-9]{1,3}|[A-Z]{1,2})$/);
    return m ? m[1] : '';
  }

  /** Bir hücredeki birden fazla barkodu ayırır. */
  function splitCodes(v) {
    if (v === null || v === undefined || v === '') return [];
    if (typeof v === 'number') return [cleanCode(v)];
    return String(v).split(/[,;\n/|]+/).map(cleanCode).filter(Boolean);
  }

  function uniq(arr) { return [...new Set(arr)]; }

  // ------------------------------------------------------------------
  // Kalıcı depolama (IndexedDB)
  // ------------------------------------------------------------------
  const DB = (() => {
    let dbp = null;
    function open() {
      if (dbp) return dbp;
      dbp = new Promise((res, rej) => {
        const r = indexedDB.open('barkod-kontrol', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('kv');
        r.onsuccess = () => res(r.result);
        r.onerror = () => rej(r.error);
      });
      return dbp;
    }
    async function op(mode, fn) {
      const db = await open();
      return new Promise((res, rej) => {
        const tx = db.transaction('kv', mode);
        const req = fn(tx.objectStore('kv'));
        tx.oncomplete = () => res(req && req.result);
        tx.onerror = () => rej(tx.error);
      });
    }
    return {
      get: (k) => op('readonly', (s) => s.get(k)).catch(() => undefined),
      set: (k, v) => op('readwrite', (s) => s.put(v, k)).catch((e) => console.warn('DB yazılamadı', e)),
      del: (k) => op('readwrite', (s) => s.delete(k)).catch(() => {}),
    };
  })();

  // ------------------------------------------------------------------
  // Uygulama durumu
  // ------------------------------------------------------------------
  let fileBuf = null;          // orijinal Excel dosyası (ArrayBuffer)
  let book = null;             // ayrıştırılmış ürün listesi
  let state = emptyState();    // kullanıcının yaptığı işlemler
  let index = new Map();       // barkod anahtarı → eşleşmeler
  let baseIndex = new Map();   // temel barkod → eşleşmeler

  function emptyState() {
    return {
      v: 1,
      fileName: '',
      added: {},       // { satırNo: [barkod, ...] }  bu oturumda eklenen yeni barkodlar
      checked: {},     // { satırNo: { s: durum, t: tarih } }
      newProducts: [], // [{ id, barcode, name, brand, note, status, t }]
      log: [],         // [{ t, code, result, name, ref }]
      dirty: 0,        // son dışa aktarımdan beri yapılan değişiklik sayısı
    };
  }

  function saveState() { return DB.set('state', state); }
  function markDirty() { state.dirty++; saveState(); renderDirty(); }

  // ------------------------------------------------------------------
  // Excel okuma
  // ------------------------------------------------------------------
  function findHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const n = (rows[i] || []).map(normText);
      if (n.includes('urun adi') && n.some((h) => h === 'barkod' || h === 'barcode')) return i;
    }
    return -1;
  }

  function parseWorkbook(buf) {
    const wb = XLSX.read(buf, { type: 'array' });
    let sheetName = null, rows = null, h = -1;
    for (const name of wb.SheetNames) {
      if (name === SHEET_NEW || name === SHEET_LOG) continue;
      const r = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: true });
      const hi = findHeader(r);
      if (hi >= 0) { sheetName = name; rows = r; h = hi; break; }
    }
    if (!sheetName) throw new Error('Excel içinde "Barkod" ve "Ürün Adı" sütunları bulunan bir sayfa bulunamadı.');

    const header = rows[h].map((x) => (x == null ? '' : String(x).trim()));
    const nh = header.map(normText);
    const idx = (name) => nh.indexOf(normText(name));
    const all = (name) => nh.reduce((a, v, i) => (v === normText(name) ? a.concat(i) : a), []);

    const cols = {
      barcodes: all('Barkod').concat(all('Barcode')),
      name: idx('Ürün Adı'),
      brand: idx('Marka'),
      code: idx('Satıcı Stok Kodu'),
      other: idx('Diğer Ürün Kodları'),
      cat: idx('Trendyol Ana Kategori'),
      sub: idx('Trendyol Alt Kategori'),
      qty: idx('Miktar'),
      unit: idx('Birim'),
      newbc: idx(COL_NEW),
      status: idx(COL_STATUS),
      date: idx(COL_DATE),
    };

    const products = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const name = r[cols.name];
      const codes = uniq(cols.barcodes.flatMap((c) => splitCodes(r[c])));
      if ((name == null || String(name).trim() === '') && codes.length === 0) continue;
      const get = (c) => (c >= 0 && r[c] != null ? String(r[c]).trim() : '');
      products.push({
        r: i,
        name: get(cols.name),
        brand: get(cols.brand),
        code: get(cols.code),
        other: get(cols.other),
        cat: [get(cols.cat), get(cols.sub)].filter(Boolean).join(' › '),
        qty: [get(cols.qty), get(cols.unit)].filter(Boolean).join(' '),
        codes,
        fileNew: cols.newbc >= 0 ? splitCodes(r[cols.newbc]) : [],
        fileStatus: get(cols.status),
        fileDate: get(cols.date),
      });
    }

    // Daha önce dışa aktarılmış dosyadaki "Yeni Ürünler" ve "Tarama Geçmişi" sayfalarını geri oku
    const newProducts = [];
    if (wb.Sheets[SHEET_NEW]) {
      const r = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_NEW], { raw: false, defval: '' });
      r.forEach((o, k) => {
        const bc = cleanCode(o['Barkod']);
        if (!bc && !o['Ürün Adı']) return;
        newProducts.push({
          id: 'f' + k + '_' + Date.now().toString(36),
          barcode: bc, name: String(o['Ürün Adı'] || '').trim(), brand: String(o['Marka'] || '').trim(),
          note: String(o['Not'] || '').trim(), status: String(o['Durum'] || S.NEWPROD).trim(), t: '', tText: String(o['Tarih'] || ''),
        });
      });
    }
    const log = [];
    if (wb.Sheets[SHEET_LOG]) {
      const r = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_LOG], { raw: false, defval: '' });
      r.forEach((o) => log.push({ t: '', tText: String(o['Tarih'] || ''), code: cleanCode(o['Okunan Barkod']), result: String(o['Sonuç'] || ''), name: String(o['Ürün Adı'] || '') }));
    }

    const brands = uniq(products.map((p) => p.brand).filter(Boolean)).sort((a, b) => a.localeCompare(b, 'tr'));
    products.forEach((p) => { p.hay = normText([p.name, p.brand, p.code, p.other, p.cat].join(' ')); p.nName = normText(p.name); });
    return { sheetName, headerRow: h, header, cols, products, byRow: new Map(products.map((p) => [p.r, p])), brands, newProducts, log };
  }

  // ------------------------------------------------------------------
  // Barkod indeksi
  // ------------------------------------------------------------------
  function productCodes(p) {
    return {
      main: p.codes,
      extra: uniq([...p.fileNew, ...(state.added[p.r] || [])]),
    };
  }

  function rebuildIndex() {
    index = new Map();
    baseIndex = new Map();
    const put = (map, key, entry) => {
      if (!key) return;
      const arr = map.get(key) || [];
      if (!arr.some((e) => e.kind === entry.kind && e.ref === entry.ref)) arr.push(entry);
      map.set(key, arr);
    };
    if (!book) return;
    for (const p of book.products) {
      const { main, extra } = productCodes(p);
      main.forEach((c) => { put(index, codeKey(c), { kind: 'p', ref: p.r, via: 'main' }); put(baseIndex, baseKey(c), { kind: 'p', ref: p.r, via: 'main' }); });
      extra.forEach((c) => { put(index, codeKey(c), { kind: 'p', ref: p.r, via: 'new' }); put(baseIndex, baseKey(c), { kind: 'p', ref: p.r, via: 'new' }); });
    }
    for (const n of state.newProducts) {
      if (n.status === S.NEWPROD && n.barcode) put(index, codeKey(n.barcode), { kind: 'n', ref: n.id, via: 'newprod' });
    }
  }

  function lookup(code) {
    const k = codeKey(code);
    if (!k) return [];
    let hits = index.get(k);
    if (!hits || !hits.length) {
      // okunan barkod "XXXX" ama listede "XXXX-1" / "XXXXT" gibi kayıtlı olabilir (veya tersi)
      hits = baseIndex.get(k) || (baseKey(code) ? index.get(baseKey(code)) : null) || [];
    }
    return hits;
  }

  // ------------------------------------------------------------------
  // İsimle arama
  // ------------------------------------------------------------------
  /** Kelime içinde arar; 1-2 harfli kelimeler (ör. "c", "d3") sadece kelime başında eşleşir. */
  function tokenPos(hay, t, from = 0) {
    let i = hay.indexOf(t, from);
    if (t.length > 2) return i;
    while (i > 0 && hay[i - 1] !== ' ') i = hay.indexOf(t, i + 1);
    return i;
  }

  function searchProducts(q, limit = 60) {
    const nq = normText(q);
    if (!nq) return { exact: [], similar: [] };
    const toks = nq.split(' ').filter(Boolean);
    const qKey = codeKey(q);
    const exact = [], similar = [];
    for (const p of book.products) {
      let hit = 0, score = 0;
      for (const t of toks) {
        const pos = tokenPos(p.hay, t);
        if (pos >= 0) {
          hit++;
          score += (pos === 0 || p.hay[pos - 1] === ' ') ? 3 : 1;
        }
      }
      if (qKey && /^\d{4,}$/.test(qKey) && p.codes.concat(productCodes(p).extra).some((c) => codeKey(c).includes(qKey))) { hit = toks.length; score += 5; }
      if (!hit) continue;
      if (p.nName.startsWith(nq)) score += 4;
      if (p.nName.includes(nq)) score += 2;
      const item = { p, score: score - p.name.length / 400, hit };
      if (hit === toks.length) exact.push(item); else if (toks.length > 1 && hit >= Math.ceil(toks.length / 2)) similar.push(item);
    }
    exact.sort((a, b) => b.score - a.score);
    similar.sort((a, b) => b.hit - a.hit || b.score - a.score);
    return { exact: exact.slice(0, limit).map((x) => x.p), similar: similar.slice(0, 20).map((x) => x.p) };
  }

  function highlight(text, q) {
    const toks = normText(q).split(' ').filter((t) => t.length > 0);
    if (!toks.length) return esc(text);
    // normalleştirilmiş metindeki konumları orijinal metne taşı (karakter sayısı korunuyor)
    const src = String(text);
    const norm = [...src].map((ch) => normText(ch) || ' ').map((c) => c[0]).join('');
    const marks = new Array(src.length).fill(false);
    for (const t of toks) {
      let i = 0;
      while ((i = tokenPos(norm, t, i)) >= 0) { for (let k = i; k < i + t.length; k++) marks[k] = true; i += t.length; }
    }
    let out = '', open = false;
    [...src].forEach((ch, i) => {
      if (marks[i] && !open) { out += '<mark>'; open = true; }
      if (!marks[i] && open) { out += '</mark>'; open = false; }
      out += esc(ch);
    });
    if (open) out += '</mark>';
    return out;
  }

  // ------------------------------------------------------------------
  // Ses / titreşim geri bildirimi
  // ------------------------------------------------------------------
  let audioCtx = null;
  function beep(kind) {
    try {
      audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
      const seq = kind === 'ok' ? [[1320, 0, 0.12]] : kind === 'info' ? [[990, 0, 0.1], [1320, 0.12, 0.1]] : [[300, 0, 0.16], [220, 0.2, 0.22]];
      seq.forEach(([f, at, dur]) => {
        const o = audioCtx.createOscillator(), g = audioCtx.createGain();
        o.type = kind === 'bad' ? 'square' : 'sine';
        o.frequency.value = f;
        g.gain.setValueAtTime(0.0001, audioCtx.currentTime + at);
        g.gain.exponentialRampToValueAtTime(0.25, audioCtx.currentTime + at + 0.01);
        g.gain.exponentialRampToValueAtTime(0.0001, audioCtx.currentTime + at + dur);
        o.connect(g).connect(audioCtx.destination);
        o.start(audioCtx.currentTime + at);
        o.stop(audioCtx.currentTime + at + dur + 0.02);
      });
    } catch (e) { /* ses desteklenmiyor */ }
    if (navigator.vibrate) navigator.vibrate(kind === 'ok' ? 60 : kind === 'info' ? [40, 60, 40] : [180, 80, 180]);
  }

  // ------------------------------------------------------------------
  // Kamera
  // ------------------------------------------------------------------
  let scanner = null;
  let camState = 'off'; // off | starting | on | paused
  let lastCode = '', lastAt = 0;
  let torchOn = false;

  function scannerFormats() {
    const F = window.Html5QrcodeSupportedFormats;
    return [F.EAN_13, F.EAN_8, F.UPC_A, F.UPC_E, F.CODE_128, F.CODE_39, F.CODE_93, F.ITF, F.CODABAR, F.QR_CODE].filter((x) => x !== undefined);
  }

  async function startCamera() {
    if (camState === 'on' || camState === 'starting') return;
    if (!window.Html5Qrcode) { $('camMsg').textContent = 'Tarayıcı kütüphanesi yüklenemedi.'; return; }
    if (!window.isSecureContext) {
      $('camMsg').textContent = 'Kamera için sayfa HTTPS üzerinden açılmalıdır. Barkodu aşağıya elle girebilirsiniz.';
      return;
    }
    camState = 'starting';
    $('btnStartCam').disabled = true;
    $('camMsg').textContent = 'Kamera açılıyor…';
    try {
      scanner = scanner || new Html5Qrcode('reader', {
        formatsToSupport: scannerFormats(),
        experimentalFeatures: { useBarCodeDetectorIfSupported: true },
        verbose: false,
      });
      await scanner.start(
        { facingMode: 'environment' },
        {
          fps: 15,
          qrbox: (w, h) => {
            const width = Math.floor(w * 0.86);
            const height = Math.floor(Math.min(h * 0.6, w * 0.5));
            return { width: Math.max(width, 50), height: Math.max(height, 50) };
          },
          aspectRatio: 4 / 3,
          disableFlip: true,
        },
        onScan,
        () => {}
      );
      camState = 'on';
      $('cameraIdle').hidden = true;
      $('cameraTools').hidden = false;
      document.querySelector('.camera-wrap').classList.add('running');
      addScanline();
      setupTorch();
      try { localStorage.setItem('bk-cam', '1'); } catch (e) {}
    } catch (err) {
      camState = 'off';
      console.warn(err);
      const msg = String(err && (err.name || err.message || err));
      $('camMsg').textContent = /NotAllowed|Permission/i.test(msg)
        ? 'Kamera izni verilmedi. Tarayıcı ayarlarından kameraya izin verin.'
        : 'Kamera açılamadı: ' + msg;
    } finally {
      $('btnStartCam').disabled = false;
    }
  }

  async function stopCamera() {
    if (!scanner || camState === 'off') return;
    try { await scanner.stop(); } catch (e) {}
    camState = 'off';
    torchOn = false;
    $('cameraIdle').hidden = false;
    $('cameraTools').hidden = true;
    document.querySelector('.camera-wrap').classList.remove('running');
    $('camMsg').textContent = '';
    try { localStorage.setItem('bk-cam', '0'); } catch (e) {}
  }

  function pauseCamera() {
    if (scanner && camState === 'on') { try { scanner.pause(true); camState = 'paused'; } catch (e) {} }
  }
  function resumeCamera() {
    if (scanner && camState === 'paused') { try { scanner.resume(); camState = 'on'; } catch (e) {} }
  }

  function addScanline() {
    if (!document.querySelector('#reader .scanline')) {
      const l = document.createElement('div');
      l.className = 'scanline';
      $('reader').appendChild(l);
    }
  }

  function setupTorch() {
    const btn = $('btnTorch');
    btn.hidden = true;
    try {
      const caps = scanner.getRunningTrackCapabilities && scanner.getRunningTrackCapabilities();
      if (caps && caps.torch) btn.hidden = false;
    } catch (e) {}
  }

  async function toggleTorch() {
    try {
      torchOn = !torchOn;
      await scanner.applyVideoConstraints({ advanced: [{ torch: torchOn }] });
      $('btnTorch').classList.toggle('on', torchOn);
    } catch (e) { torchOn = false; toast('Fener açılamadı'); }
  }

  function onScan(text) {
    const code = cleanCode(text);
    if (!code) return;
    const t = Date.now();
    // Aynı barkod kamerada durduğu sürece tekrar sayılmaz; çekilip tekrar okutulursa sayılır
    if (code === lastCode && t - lastAt < SAME_CODE_COOLDOWN) { lastAt = t; return; }
    lastCode = code; lastAt = t;
    handleCode(code);
  }

  // ------------------------------------------------------------------
  // Barkod kontrolü
  // ------------------------------------------------------------------
  function handleCode(code) {
    if (!book) { toast('Önce ürün listesi yüklenmeli'); return; }
    const hits = lookup(code);
    if (hits.length) {
      const first = hits[0];
      if (first.kind === 'p') {
        const p = book.byRow.get(first.ref);
        const viaNew = hits.every((h) => h.via === 'new');
        hits.filter((h) => h.kind === 'p').forEach((h) => { state.checked[h.ref] = { s: viaNew ? S.NEWBC : S.MATCH, t: nowIso() }; });
        addLog(code, S.MATCH, p.name, first.ref);
        const extra = hits.length > 1 ? `<div class="small">⚠ Bu barkod listede ${hits.length} üründe kayıtlı</div>` : '';
        showResult('ok', 'EŞLEŞİYOR ✓', p.name,
          `${esc(code)} · Stok: ${esc(p.code)}${p.brand ? ' · ' + esc(p.brand) : ''}${viaNew ? ' · <b>Yeni Barkod ile</b>' : ''}${extra}`);
        beep('ok');
      } else {
        const n = state.newProducts.find((x) => x.id === first.ref);
        addLog(code, S.MATCH, n ? n.name : '', 'n:' + first.ref);
        showResult('info', 'YENİ ÜRÜN LİSTESİNDE', n ? n.name : '', `${esc(code)} · Daha önce yeni ürün olarak eklendi`);
        beep('info');
      }
      markDirty();
      renderAfterChange();
      return;
    }
    // Eşleşme yok → isimle arama paneli
    showResult('bad', 'EŞLEŞME YOK ✗', '', `${esc(code)} · Ürünü adıyla arayın`);
    beep('bad');
    openNoMatch(code);
  }

  function showResult(kind, title, name, metaHtml) {
    const card = $('resultCard');
    card.className = 'result ' + kind;
    void card.offsetWidth;
    card.classList.add('flash');
    const icons = {
      ok: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
      bad: '<path d="M6 6l12 12M18 6 6 18"/>',
      info: '<path d="M12 5v14M5 12h14"/>',
      idle: '<path d="M3 7V5a2 2 0 0 1 2-2h2M17 3h2a2 2 0 0 1 2 2v2M21 17v2a2 2 0 0 1-2 2h-2M7 21H5a2 2 0 0 1-2-2v-2M7 8v8M10 8v8M13 8v8M17 8v8"/>',
      warn: '<path d="M12 8v5M12 16.5v.01M10.3 3.9 2.4 18a2 2 0 0 0 1.7 3h15.8a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/>',
    };
    $('resultIcon').innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true">${icons[kind] || icons.info}</svg>`;
    $('resultTitle').textContent = title;
    $('resultName').textContent = name || '';
    $('resultMeta').innerHTML = metaHtml || '';
  }

  function addLog(code, result, name, ref) {
    state.log.push({ t: nowIso(), code, result, name: name || '', ref: ref ?? null });
    if (state.log.length > 20000) state.log.splice(0, state.log.length - 20000);
  }

  // ------------------------------------------------------------------
  // Eşleşme yok paneli
  // ------------------------------------------------------------------
  let sheetCode = '';
  let sheetMode = ''; // nomatch | newprod | detail
  let editingNewId = null;

  function openSheet(mode) {
    sheetMode = mode;
    $('sheet').hidden = false;
    document.body.style.overflow = 'hidden';
    $('stepSearch').hidden = mode !== 'nomatch';
    $('footSearch').hidden = mode !== 'nomatch';
    $('stepNew').hidden = mode !== 'newprod';
    $('stepDetail').hidden = mode !== 'detail';
  }

  function closeSheet(logNoMatch = true) {
    if (sheetMode === 'nomatch' && logNoMatch && sheetCode) {
      addLog(sheetCode, S.NOMATCH, '', null);
      markDirty();
      renderAfterChange();
    }
    $('sheet').hidden = true;
    document.body.style.overflow = '';
    sheetMode = '';
    sheetCode = '';
    editingNewId = null;
    $('nameSearch').blur();
    resumeCamera();
  }

  function openNoMatch(code) {
    pauseCamera();
    sheetCode = code;
    const prev = state.newProducts.find((n) => codeKey(n.barcode) === codeKey(code));
    $('sheetHead').className = 'sheet-head';
    $('sheetKicker').textContent = 'Eşleşme yok';
    $('sheetTitle').textContent = code;
    $('searchHint').innerHTML = (prev && prev.status === S.NOTFOUND
      ? '<b>Bu barkod daha önce “Bulunamadı” olarak işaretlendi.</b> '
      : '') + 'Ürünü adıyla arayın. Bulduğunuz ürüne dokunun, okutulan barkod <b>Yeni Barkod</b> olarak eklenir.';
    $('nameSearch').value = '';
    renderSearchResults('');
    openSheet('nomatch');
    setTimeout(() => $('nameSearch').focus(), 150);
  }

  function renderSearchResults(q) {
    const ul = $('searchResults');
    if (!normText(q)) {
      ul.innerHTML = '<li class="empty">Aramak için ürün adından birkaç harf yazın.<br>Örn: <b>vitamin c</b>, <b>ocean d3</b>, <b>serum 30</b></li>';
      return;
    }
    const { exact, similar } = searchProducts(q);
    if (!exact.length && !similar.length) {
      ul.innerHTML = `<li class="empty">“${esc(q)}” ile eşleşen ürün yok.<br>Farklı bir kelime deneyin veya aşağıdan <b>Bulunamadı</b> işaretleyin.</li>`;
      return;
    }
    const li = (p) => {
      const { main, extra } = productCodes(p);
      return `<li class="item" data-r="${p.r}">
        <div class="item-main">
          <div class="item-title">${highlight(p.name, q)}</div>
          <div class="item-sub"><span>${esc(p.brand)}</span><span>Stok: ${esc(p.code)}</span>
          <span class="mono">${main.length ? esc(main.join(', ')) : '<span class="tag muted">Barkodsuz</span>'}</span>
          ${extra.length ? `<span class="tag info">Yeni: ${esc(extra.join(', '))}</span>` : ''}</div>
        </div>
        <div class="item-side"><span class="tag ok">Seç</span></div>
      </li>`;
    };
    let html = exact.map(li).join('');
    if (similar.length) html += `<li class="section-title">Benzer ürünler</li>` + similar.map(li).join('');
    ul.innerHTML = html;
  }

  function assignBarcode(r, code, fromSheet) {
    const p = book.byRow.get(r);
    const { extra } = productCodes(p);
    const existing = extra.length ? `<p class="small muted">Bu üründe zaten yeni barkod var: <span class="mono">${esc(extra.join(', '))}</span>. Yeni barkod yanına eklenecek.</p>` : '';
    confirmBox('Yeni barkod eklensin mi?',
      `<p><b>${esc(p.name)}</b></p><p>Yeni Barkod: <span class="mono">${esc(code)}</span></p>${existing}`,
      'Ekle').then((ok) => {
      if (!ok) return;
      state.added[r] = uniq([...(state.added[r] || []), code]);
      state.checked[r] = { s: S.NEWBC, t: nowIso() };
      // daha önce "bulunamadı" olarak işaretlendiyse o kaydı kaldır
      state.newProducts = state.newProducts.filter((n) => !(n.status === S.NOTFOUND && codeKey(n.barcode) === codeKey(code)));
      addLog(code, S.NEWBC, p.name, r);
      rebuildIndex();
      markDirty();
      renderAfterChange();
      showResult('ok', 'YENİ BARKOD EKLENDİ ✓', p.name, `${esc(code)} → Yeni Barkod sütununa yazıldı`);
      beep('ok');
      toast('Yeni barkod eklendi');
      if (fromSheet) closeSheet(false);
      else openDetail(r);
    });
  }

  function markNotFound() {
    const code = sheetCode;
    const q = $('nameSearch').value.trim();
    let n = state.newProducts.find((x) => codeKey(x.barcode) === codeKey(code));
    if (!n) {
      n = { id: 'n' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), barcode: code, name: '', brand: '', note: '', status: S.NOTFOUND, t: nowIso() };
      state.newProducts.push(n);
      addLog(code, S.NOTFOUND, '', 'n:' + n.id);
      markDirty();
      renderAfterChange();
    }
    showResult('warn', 'BULUNAMADI', '', `${esc(code)} · Yeni ürün olarak ekleyebilirsiniz`);
    openNewProductForm(n, q);
  }

  function openNewProductForm(n, nameGuess) {
    editingNewId = n.id;
    sheetCode = n.barcode;
    $('sheetHead').className = 'sheet-head warn';
    $('sheetKicker').textContent = n.status === S.NEWPROD ? 'Yeni ürünü düzenle' : 'Bulunamadı · Yeni ürün ekle';
    $('sheetTitle').textContent = n.barcode;
    $('notFoundNotice').hidden = n.status === S.NEWPROD;
    $('npName').value = n.name || (nameGuess ? nameGuess.toLocaleUpperCase('tr') : '');
    $('npBarcode').value = n.barcode;
    $('npBrand').value = n.brand || '';
    $('npNote').value = n.note || '';
    $('btnKeepNotFound').textContent = n.status === S.NEWPROD ? 'Vazgeç' : 'Sadece “Bulunamadı” olarak bırak';
    openSheet('newprod');
    setTimeout(() => $('npName').focus(), 150);
  }

  function saveNewProduct(e) {
    e.preventDefault();
    const n = state.newProducts.find((x) => x.id === editingNewId);
    if (!n) { closeSheet(false); return; }
    const name = $('npName').value.trim();
    const bc = cleanCode($('npBarcode').value);
    if (!name || !bc) { toast('Ürün adı ve barkod gerekli'); return; }
    const clash = lookup(bc).filter((h) => !(h.kind === 'n' && h.ref === n.id));
    if (clash.length) {
      const c = clash[0];
      const cname = c.kind === 'p' ? book.byRow.get(c.ref).name : (state.newProducts.find((x) => x.id === c.ref) || {}).name;
      toast('Bu barkod zaten kayıtlı: ' + cname);
      return;
    }
    const wasNew = n.status === S.NEWPROD;
    Object.assign(n, { name, barcode: bc, brand: $('npBrand').value.trim(), note: $('npNote').value.trim(), status: S.NEWPROD, t: n.t || nowIso() });
    if (!wasNew) addLog(bc, S.NEWPROD, name, 'n:' + n.id);
    rebuildIndex();
    markDirty();
    renderAfterChange();
    showResult('info', wasNew ? 'YENİ ÜRÜN GÜNCELLENDİ' : 'YENİ ÜRÜN EKLENDİ', name, `${esc(bc)} · “${SHEET_NEW}” sayfasına yazılacak`);
    if (!wasNew) beep('info');
    toast(wasNew ? 'Yeni ürün güncellendi' : 'Yeni ürün eklendi');
    closeSheet(false);
  }

  // ------------------------------------------------------------------
  // Ürün detayı (Ürünler sekmesinden)
  // ------------------------------------------------------------------
  function openDetail(r) {
    const p = book.byRow.get(r);
    if (!p) return;
    const { main, extra } = productCodes(p);
    const chk = state.checked[r];
    $('sheetHead').className = 'sheet-head info';
    $('sheetKicker').textContent = 'Ürün detayı';
    $('sheetTitle').textContent = p.code || '—';
    $('stepDetail').innerHTML = `
      <p class="detail-name">${esc(p.name)}</p>
      <dl class="kv">
        <dt>Marka</dt><dd>${esc(p.brand) || '—'}</dd>
        <dt>Stok kodu</dt><dd>${esc(p.code) || '—'}</dd>
        ${p.other ? `<dt>Diğer kod</dt><dd>${esc(p.other)}</dd>` : ''}
        <dt>Barkod</dt><dd class="mono">${main.length ? esc(main.join(', ')) : '<span class="tag muted">Yok</span>'}</dd>
        <dt>Yeni barkod</dt><dd class="mono">${extra.length ? esc(extra.join(', ')) : '—'}</dd>
        ${p.cat ? `<dt>Kategori</dt><dd>${esc(p.cat)}</dd>` : ''}
        ${p.qty ? `<dt>Miktar</dt><dd>${esc(p.qty)}</dd>` : ''}
        <dt>Kontrol</dt><dd>${chk ? `<span class="tag ${chk.s === S.MATCH ? 'ok' : 'info'}">${esc(chk.s)}</span> <span class="muted small">${fmtDate(chk.t)}</span>` : (p.fileStatus ? `<span class="tag muted">${esc(p.fileStatus)}</span> <span class="muted small">${esc(p.fileDate)}</span>` : '<span class="tag muted">Kontrol edilmedi</span>')}</dd>
      </dl>
      <form id="detailForm" class="manual" autocomplete="off">
        <input id="detailCode" type="text" inputmode="numeric" placeholder="Bu ürüne yeni barkod ekle" aria-label="Yeni barkod">
        <button class="btn btn-primary" type="submit">Ekle</button>
      </form>`;
    $('detailForm').addEventListener('submit', (e) => {
      e.preventDefault();
      const code = cleanCode($('detailCode').value);
      if (!code) return;
      const hits = lookup(code);
      if (hits.length) {
        const h = hits[0];
        const nm = h.kind === 'p' ? book.byRow.get(h.ref).name : (state.newProducts.find((x) => x.id === h.ref) || {}).name;
        toast(h.kind === 'p' && h.ref === r ? 'Bu barkod zaten bu ürüne ait' : 'Bu barkod başka üründe kayıtlı: ' + nm);
        return;
      }
      assignBarcode(r, code, false);
    });
    openSheet('detail');
  }

  // ------------------------------------------------------------------
  // Onay kutusu ve bildirim
  // ------------------------------------------------------------------
  function confirmBox(title, html, yesText = 'Tamam', danger = false) {
    return new Promise((resolve) => {
      $('confirmTitle').textContent = title;
      $('confirmText').innerHTML = html;
      const yes = $('confirmYes'), no = $('confirmNo');
      yes.textContent = yesText;
      yes.className = 'btn grow ' + (danger ? 'btn-warn' : 'btn-primary');
      $('confirm').hidden = false;
      const done = (v) => { $('confirm').hidden = true; yes.onclick = no.onclick = null; resolve(v); };
      yes.onclick = () => done(true);
      no.onclick = () => done(false);
      setTimeout(() => yes.focus(), 50);
    });
  }

  let toastTimer = null;
  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 2600);
  }

  // ------------------------------------------------------------------
  // Ekran çizimleri
  // ------------------------------------------------------------------
  function counts() {
    const scans = state.log.filter((l) => l.t);
    return {
      match: scans.filter((l) => l.result === S.MATCH).length,
      newbc: Object.values(state.added).reduce((a, v) => a + v.length, 0),
      notfound: state.newProducts.filter((n) => n.status === S.NOTFOUND).length,
      newprod: state.newProducts.filter((n) => n.status === S.NEWPROD).length,
      checked: book ? book.products.filter((p) => state.checked[p.r]).length : 0,
    };
  }

  function renderStats() {
    const c = counts();
    $('stats').innerHTML = `
      <div class="stat ok"><b>${c.match}</b><span>Eşleşen okutma</span></div>
      <div class="stat info"><b>${c.newbc}</b><span>Yeni barkod</span></div>
      <div class="stat warn"><b>${c.notfound}</b><span>Bulunamadı</span></div>
      <div class="stat"><b>${c.newprod}</b><span>Yeni ürün</span></div>`;
  }

  const RESULT_CLASS = { [S.MATCH]: 'ok', [S.NEWBC]: 'info', [S.NOTFOUND]: 'warn', [S.NEWPROD]: 'info', [S.NOMATCH]: 'bad' };

  function renderRecent() {
    const items = state.log.filter((l) => l.t).slice(-8).reverse();
    $('recentList').innerHTML = items.length ? items.map((l) => `
      <li class="item" data-code="${esc(l.code)}">
        <span class="dot ${RESULT_CLASS[l.result] || 'none'}"></span>
        <div class="item-main">
          <div class="item-title">${esc(l.name || l.code)}</div>
          <div class="item-sub"><span class="mono">${esc(l.code)}</span><span>${fmtTime(l.t)}</span></div>
        </div>
        <span class="tag ${RESULT_CLASS[l.result] || 'muted'}">${esc(l.result)}</span>
      </li>`).join('') : '<li class="empty">Henüz okutma yok. Kamerayı başlatıp barkodu okutun.</li>';
  }

  let listFilter = 'all', listLimit = 100;
  function renderList() {
    if (!book) return;
    const q = $('listSearch').value;
    let items = normText(q) ? searchProducts(q, 100000).exact : book.products.slice();
    items = items.filter((p) => {
      const { main, extra } = productCodes(p);
      switch (listFilter) {
        case 'checked': return !!state.checked[p.r];
        case 'unchecked': return !state.checked[p.r];
        case 'newbc': return extra.length > 0;
        case 'nobc': return main.length === 0 && extra.length === 0;
        default: return true;
      }
    });
    $('listCount').textContent = `${items.length} ürün`;
    const shown = items.slice(0, listLimit);
    $('productList').innerHTML = shown.length ? shown.map((p) => {
      const { main, extra } = productCodes(p);
      const chk = state.checked[p.r];
      const tag = chk ? `<span class="tag ${chk.s === S.MATCH ? 'ok' : 'info'}">${chk.s === S.MATCH ? '✓' : 'Yeni'}</span>` : '';
      return `<li class="item" data-r="${p.r}">
        <span class="dot ${chk ? (chk.s === S.MATCH ? 'ok' : 'info') : 'none'}"></span>
        <div class="item-main">
          <div class="item-title">${normText(q) ? highlight(p.name, q) : esc(p.name)}</div>
          <div class="item-sub"><span>${esc(p.brand)}</span><span class="mono">${main.length ? esc(main[0]) : 'Barkodsuz'}</span>${extra.length ? `<span class="tag info">+${extra.length} yeni</span>` : ''}</div>
        </div>
        <div class="item-side">${tag}</div>
      </li>`;
    }).join('') : '<li class="empty">Sonuç yok</li>';
    $('listMore').hidden = items.length <= listLimit;
  }

  let logFilter = 'all';
  function renderLog() {
    const ul = $('logList');
    let html = '';
    const trash = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h16M10 11v6M14 11v6M5 7l1 12a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2l1-12M9 7V4h6v3"/></svg>';
    if (logFilter === 'all' || logFilter === 'newbc') {
      const rows = Object.entries(state.added).flatMap(([r, codes]) => codes.map((c) => ({ r: +r, c })));
      if (logFilter === 'newbc' || rows.length) html += `<li class="section-title">Eklenen yeni barkodlar (${rows.length})</li>`;
      html += rows.map(({ r, c }) => {
        const p = book && book.byRow.get(r);
        return `<li class="item" data-r="${r}">
          <span class="dot info"></span>
          <div class="item-main"><div class="item-title">${esc(p ? p.name : 'Satır ' + (r + 1))}</div>
          <div class="item-sub"><span>Yeni Barkod:</span><span class="mono">${esc(c)}</span></div></div>
          <button class="del-btn" data-undo-bc="${r}|${esc(c)}" aria-label="Geri al">${trash}</button>
        </li>`;
      }).join('');
      if (logFilter === 'newbc' && !rows.length) html += '<li class="empty">Henüz yeni barkod eklenmedi.</li>';
    }
    if (logFilter === 'all' || logFilter === 'newprod') {
      const rows = state.newProducts;
      if (logFilter === 'newprod' || rows.length) html += `<li class="section-title">Yeni ürünler / Bulunamayanlar (${rows.length})</li>`;
      html += rows.slice().reverse().map((n) => `
        <li class="item" data-new="${n.id}">
          <span class="dot ${n.status === S.NEWPROD ? 'info' : 'warn'}"></span>
          <div class="item-main"><div class="item-title">${esc(n.name || '(isimsiz)')}</div>
          <div class="item-sub"><span class="mono">${esc(n.barcode)}</span>${n.brand ? `<span>${esc(n.brand)}</span>` : ''}<span class="tag ${n.status === S.NEWPROD ? 'info' : 'warn'}">${esc(n.status)}</span></div></div>
          <button class="del-btn" data-del-new="${n.id}" aria-label="Sil">${trash}</button>
        </li>`).join('');
      if (logFilter === 'newprod' && !rows.length) html += '<li class="empty">Kayıt yok.</li>';
    }
    if (logFilter === 'all' || logFilter === 'scan') {
      const rows = state.log.slice().reverse();
      const lim = logFilter === 'scan' ? 300 : 30;
      html += `<li class="section-title">Tarama geçmişi (${rows.length})</li>`;
      html += rows.slice(0, lim).map((l) => `
        <li class="item" data-code="${esc(l.code)}">
          <span class="dot ${RESULT_CLASS[l.result] || 'none'}"></span>
          <div class="item-main"><div class="item-title">${esc(l.name || l.code)}</div>
          <div class="item-sub"><span class="mono">${esc(l.code)}</span><span>${l.t ? fmtDate(l.t) : esc(l.tText || '')}</span></div></div>
          <span class="tag ${RESULT_CLASS[l.result] || 'muted'}">${esc(l.result)}</span>
        </li>`).join('');
      if (!rows.length) html += '<li class="empty">Tarama geçmişi boş.</li>';
      else if (rows.length > lim) html += `<li class="empty">… ve ${rows.length - lim} kayıt daha (Excel çıktısında hepsi var)</li>`;
    }
    ul.innerHTML = html;
  }

  function renderDirty() {
    const b = $('dirtyBadge');
    b.textContent = state.dirty > 99 ? '99+' : String(state.dirty);
    b.hidden = !state.dirty;
  }

  function renderFileInfo() {
    if (!book) return;
    const withBc = book.products.filter((p) => p.codes.length).length;
    const c = counts();
    $('fileInfo').textContent = `${book.products.length} ürün · ${c.checked} kontrol edildi`;
    $('fileDetail').innerHTML = `<b>${esc(state.fileName)}</b><br>Sayfa: ${esc(book.sheetName)} · ${book.products.length} ürün · ${withBc} barkodlu`;
    $('brandList').innerHTML = book.brands.map((b) => `<option value="${esc(b)}">`).join('');
  }

  function renderAfterChange() {
    renderStats();
    renderRecent();
    renderFileInfo();
    if ($('tab-list').classList.contains('active')) renderList();
    if ($('tab-log').classList.contains('active')) renderLog();
  }

  // ------------------------------------------------------------------
  // Excel'e yazma
  // ------------------------------------------------------------------
  function codeCell(v) {
    const s = String(v);
    if (/^[1-9]\d{0,14}$/.test(s)) return { t: 'n', v: Number(s), z: '0' };
    return { t: 's', v: s };
  }

  function buildWorkbook() {
    const wb = XLSX.read(fileBuf, { type: 'array', cellStyles: true });
    const ws = wb.Sheets[book.sheetName];
    const range = XLSX.utils.decode_range(ws['!ref']);
    const hr = book.headerRow;
    let lastCol = Math.max(range.e.c, book.header.length - 1);
    const ensureCol = (idx, title) => {
      if (idx >= 0) return idx;
      lastCol += 1;
      ws[XLSX.utils.encode_cell({ r: hr, c: lastCol })] = { t: 's', v: title };
      return lastCol;
    };
    const cNew = ensureCol(book.cols.newbc, COL_NEW);
    const cStat = ensureCol(book.cols.status, COL_STATUS);
    const cDate = ensureCol(book.cols.date, COL_DATE);

    for (const p of book.products) {
      const added = state.added[p.r] || [];
      if (added.length) {
        const all = uniq([...p.fileNew, ...added]);
        ws[XLSX.utils.encode_cell({ r: p.r, c: cNew })] = all.length === 1 ? codeCell(all[0]) : { t: 's', v: all.join(', ') };
      }
      const chk = state.checked[p.r];
      if (chk) {
        ws[XLSX.utils.encode_cell({ r: p.r, c: cStat })] = { t: 's', v: chk.s };
        ws[XLSX.utils.encode_cell({ r: p.r, c: cDate })] = { t: 's', v: fmtDate(chk.t) };
      }
    }
    range.e.c = Math.max(range.e.c, lastCol);
    ws['!ref'] = XLSX.utils.encode_range(range);
    const cols = ws['!cols'] || [];
    [cNew, cStat, cDate].forEach((c, i) => { if (!cols[c] || !cols[c].wch) cols[c] = { wch: [18, 20, 17][i] }; });
    ws['!cols'] = cols;
    ws['!autofilter'] = ws['!autofilter'] || { ref: XLSX.utils.encode_range({ s: { r: hr, c: range.s.c }, e: { r: range.e.r, c: range.e.c } }) };

    // Yeni Ürünler sayfası
    const npRows = [['Barkod', 'Ürün Adı', 'Marka', 'Durum', 'Not', 'Tarih']];
    state.newProducts.forEach((n) => npRows.push([n.barcode, n.name, n.brand, n.status, n.note, n.t ? fmtDate(n.t) : (n.tText || '')]));
    const wsNew = XLSX.utils.aoa_to_sheet(npRows);
    for (let i = 1; i < npRows.length; i++) wsNew[XLSX.utils.encode_cell({ r: i, c: 0 })] = codeCell(npRows[i][0]);
    wsNew['!cols'] = [{ wch: 16 }, { wch: 50 }, { wch: 16 }, { wch: 14 }, { wch: 24 }, { wch: 17 }];
    putSheet(wb, SHEET_NEW, wsNew);

    // Tarama Geçmişi sayfası
    const lgRows = [['Tarih', 'Okunan Barkod', 'Sonuç', 'Ürün Adı']];
    state.log.forEach((l) => lgRows.push([l.t ? fmtDate(l.t) : (l.tText || ''), l.code, l.result, l.name]));
    const wsLog = XLSX.utils.aoa_to_sheet(lgRows);
    for (let i = 1; i < lgRows.length; i++) wsLog[XLSX.utils.encode_cell({ r: i, c: 1 })] = codeCell(lgRows[i][1]);
    wsLog['!cols'] = [{ wch: 17 }, { wch: 16 }, { wch: 20 }, { wch: 50 }];
    putSheet(wb, SHEET_LOG, wsLog);

    return XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
  }

  function putSheet(wb, name, ws) {
    if (wb.Sheets[name]) wb.Sheets[name] = ws;
    else XLSX.utils.book_append_sheet(wb, ws, name);
  }

  function exportName() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    const base = (state.fileName || 'Barkod-Kontrol').replace(/\.(xlsx|xls)$/i, '').replace(/_kontrol_\d{8}-\d{4}$/, '');
    return `${base}_kontrol_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}.xlsx`;
  }

  function exportBlob() {
    const out = buildWorkbook();
    return new File([out], exportName(), { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function afterExport() {
    state.dirty = 0;
    saveState();
    renderDirty();
  }

  async function doDownload() {
    if (!book) return;
    try {
      const f = exportBlob();
      const url = URL.createObjectURL(f);
      const a = document.createElement('a');
      a.href = url; a.download = f.name;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      afterExport();
      toast('Excel indirildi: ' + f.name);
    } catch (e) { console.error(e); toast('Excel oluşturulamadı: ' + e.message); }
  }

  async function doShare() {
    if (!book) return;
    try {
      const f = exportBlob();
      if (navigator.canShare && navigator.canShare({ files: [f] })) {
        await navigator.share({ files: [f], title: 'Barkod Kontrol', text: 'Barkod kontrol sonuçları' });
        afterExport();
      } else {
        doDownload();
      }
    } catch (e) {
      if (e && e.name === 'AbortError') return;
      console.error(e); toast('Paylaşılamadı: ' + e.message);
    }
  }

  // ------------------------------------------------------------------
  // Dosya yükleme
  // ------------------------------------------------------------------
  async function loadBuffer(buf, fileName, { keepState } = { keepState: false }) {
    const parsed = parseWorkbook(buf);
    fileBuf = buf;
    book = parsed;
    if (!keepState) {
      state = emptyState();
      state.fileName = fileName;
      state.newProducts = parsed.newProducts;
      state.log = parsed.log;
      await DB.set('file', buf);
      await saveState();
    }
    rebuildIndex();
    renderAfterChange();
    renderDirty();
    if ($('tab-list').classList.contains('active')) renderList();
  }

  async function onFileChosen(e) {
    const f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!f) return;
    if (state.dirty) {
      const ok = await confirmBox('Yeni dosya yüklensin mi?',
        `<p>Dışa aktarılmamış <b>${state.dirty}</b> değişiklik var. Yeni dosya yüklenirse bu kayıtlar silinir.</p><p class="small muted">Önce “Excel İndir” ile kaydetmeniz önerilir.</p>`, 'Yükle', true);
      if (!ok) return;
    }
    try {
      const buf = await f.arrayBuffer();
      await loadBuffer(buf, f.name);
      toast(`${book.products.length} ürün yüklendi`);
    } catch (err) { console.error(err); toast('Dosya okunamadı: ' + err.message); }
  }

  async function resetAll() {
    const ok = await confirmBox('Kayıtlar sıfırlansın mı?',
      '<p>Eklenen yeni barkodlar, yeni ürünler ve tarama geçmişi silinir. Ürün listesi korunur.</p>', 'Sıfırla', true);
    if (!ok) return;
    const name = state.fileName;
    state = emptyState();
    state.fileName = name;
    await saveState();
    rebuildIndex();
    renderAfterChange();
    renderDirty();
    showResult('idle', 'Barkodu kameraya gösterin', '', '');
    toast('Kayıtlar sıfırlandı');
  }

  // ------------------------------------------------------------------
  // Olaylar
  // ------------------------------------------------------------------
  function switchTab(name) {
    document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.id === 'tab-' + name));
    document.querySelectorAll('.tabbtn').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
    if (name === 'list') renderList();
    if (name === 'log') renderLog();
    if (name === 'scan') resumeCamera(); else pauseCamera();
    window.scrollTo(0, 0);
  }

  function bind() {
    document.querySelectorAll('.tabbtn').forEach((b) => b.addEventListener('click', () => switchTab(b.dataset.tab)));
    $('btnStartCam').addEventListener('click', startCamera);
    $('btnStopCam').addEventListener('click', stopCamera);
    $('btnTorch').addEventListener('click', toggleTorch);

    $('manualForm').addEventListener('submit', (e) => {
      e.preventDefault();
      const code = cleanCode($('manualInput').value);
      if (!code) return;
      $('manualInput').value = '';
      $('manualInput').blur();
      lastCode = code; lastAt = Date.now();
      handleCode(code);
    });

    // Eşleşme yok paneli
    let st = null;
    $('nameSearch').addEventListener('input', (e) => { clearTimeout(st); st = setTimeout(() => renderSearchResults(e.target.value), 80); });
    $('nameSearch').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); } });
    $('searchResults').addEventListener('click', (e) => {
      const li = e.target.closest('[data-r]');
      if (li) assignBarcode(+li.dataset.r, sheetCode, true);
    });
    $('btnNotFound').addEventListener('click', markNotFound);
    $('stepNew').addEventListener('submit', saveNewProduct);
    $('btnKeepNotFound').addEventListener('click', () => {
      const n = state.newProducts.find((x) => x.id === editingNewId);
      if (n && n.status === S.NOTFOUND) toast('“Bulunamadı” olarak kaydedildi');
      closeSheet(false);
    });
    $('sheetClose').addEventListener('click', () => closeSheet(true));
    $('sheet').addEventListener('click', (e) => { if (e.target === $('sheet')) closeSheet(true); });

    // Ürünler sekmesi
    let lt = null;
    $('listSearch').addEventListener('input', () => { clearTimeout(lt); lt = setTimeout(() => { listLimit = 100; renderList(); }, 100); });
    $('listFilters').addEventListener('click', (e) => {
      const c = e.target.closest('.chip'); if (!c) return;
      listFilter = c.dataset.f; listLimit = 100;
      $('listFilters').querySelectorAll('.chip').forEach((x) => x.classList.toggle('active', x === c));
      renderList();
    });
    $('listMore').addEventListener('click', () => { listLimit += 200; renderList(); });
    $('productList').addEventListener('click', (e) => { const li = e.target.closest('[data-r]'); if (li) openDetail(+li.dataset.r); });

    // Kayıtlar sekmesi
    $('logFilters').addEventListener('click', (e) => {
      const c = e.target.closest('.chip'); if (!c) return;
      logFilter = c.dataset.f;
      $('logFilters').querySelectorAll('.chip').forEach((x) => x.classList.toggle('active', x === c));
      renderLog();
    });
    $('logList').addEventListener('click', async (e) => {
      const undo = e.target.closest('[data-undo-bc]');
      if (undo) {
        const [r, code] = undo.dataset.undoBc.split('|');
        const p = book.byRow.get(+r);
        if (!(await confirmBox('Yeni barkod geri alınsın mı?', `<p><b>${esc(p ? p.name : '')}</b></p><p class="mono">${esc(code)}</p>`, 'Geri al', true))) return;
        state.added[r] = (state.added[r] || []).filter((c) => c !== code);
        if (!state.added[r].length) { delete state.added[r]; if (state.checked[r] && state.checked[r].s === S.NEWBC) delete state.checked[r]; }
        rebuildIndex(); markDirty(); renderAfterChange(); renderLog();
        toast('Geri alındı');
        return;
      }
      const del = e.target.closest('[data-del-new]');
      if (del) {
        const n = state.newProducts.find((x) => x.id === del.dataset.delNew);
        if (!n || !(await confirmBox('Kayıt silinsin mi?', `<p><b>${esc(n.name || '(isimsiz)')}</b></p><p class="mono">${esc(n.barcode)}</p>`, 'Sil', true))) return;
        state.newProducts = state.newProducts.filter((x) => x !== n);
        rebuildIndex(); markDirty(); renderAfterChange(); renderLog();
        toast('Silindi');
        return;
      }
      const nn = e.target.closest('[data-new]');
      if (nn) { const n = state.newProducts.find((x) => x.id === nn.dataset.new); if (n) openNewProductForm(n, ''); return; }
      const pr = e.target.closest('[data-r]');
      if (pr) openDetail(+pr.dataset.r);
    });

    $('btnDownload').addEventListener('click', doDownload);
    $('btnShare').addEventListener('click', doShare);
    $('btnExportTop').addEventListener('click', () => switchTab('log'));
    $('fileInput').addEventListener('change', onFileChosen);
    $('btnReset').addEventListener('click', resetAll);

    try {
      const probe = new File([''], 'a.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      $('btnShare').hidden = !(navigator.canShare && navigator.canShare({ files: [probe] }));
    } catch (e) { /* paylaşım desteklenmiyor */ }

    document.addEventListener('visibilitychange', () => {
      if (document.hidden) pauseCamera();
      else if (!sheetMode && $('tab-scan').classList.contains('active')) resumeCamera();
    });
  }

  // ------------------------------------------------------------------
  // Başlangıç
  // ------------------------------------------------------------------
  /** Varsayılan ürün listesi: sunucudan, olmazsa (file:// ile açıldığında) gömülü kopyadan. */
  async function defaultFile() {
    if (location.protocol !== 'file:') {
      try {
        const res = await fetch(DEFAULT_FILE, { cache: 'no-cache' });
        if (res.ok) return await res.arrayBuffer();
      } catch (e) { /* gömülü kopyaya geç */ }
    }
    if (!window.DEFAULT_XLSX_B64) throw new Error('Varsayılan ürün listesi bulunamadı');
    const bin = atob(window.DEFAULT_XLSX_B64);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8.buffer;
  }

  async function init() {
    bind();
    renderStats();
    renderRecent();
    try {
      const [buf, saved] = await Promise.all([DB.get('file'), DB.get('state')]);
      if (buf && saved) {
        state = Object.assign(emptyState(), saved);
        await loadBuffer(buf, state.fileName, { keepState: true });
      } else {
        await loadBuffer(await defaultFile(), 'Barkod-Kontrol.xlsx');
      }
    } catch (e) {
      console.error(e);
      $('fileInfo').textContent = 'Ürün listesi yüklenmedi';
      showResult('warn', 'Ürün listesi yok', '', 'Kayıtlar sekmesinden “Excel Yükle” ile ürün listesini seçin.');
    }
    let wantCam = '1';
    try { wantCam = localStorage.getItem('bk-cam') ?? '1'; } catch (e) {}
    if (wantCam === '1' && window.isSecureContext) startCamera();
    else if (!window.isSecureContext) $('camMsg').textContent = 'Kamera için sayfa HTTPS üzerinden açılmalıdır.';

    if ('serviceWorker' in navigator && location.protocol === 'https:') {
      navigator.serviceWorker.register('sw.js').catch(() => {});
    }
  }

  // test/hata ayıklama için
  window.BarkodKontrol = { handleCode, lookup, searchProducts, get state() { return state; }, get book() { return book; }, buildWorkbook };

  init();
})();
