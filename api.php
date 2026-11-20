<?php
/**
 * Barkod Kontrol — ortak sayım sunucusu
 *
 * Sitenin bulunduğu klasöre yüklenir; ek kurulum veya veritabanı gerekmez (PHP 7.0+).
 * Kayıtlar "veri" klasöründe tutulur (klasör otomatik oluşturulur ve dışarıya kapatılır).
 *
 *   GET  api.php?since=<konum>            → yeni kayıtları döndürür
 *   POST api.php  {"epoch":..,"events":[..]} → yeni kayıtları ekler
 *   POST api.php  {"action":"reset"}       → yeni sayım: mevcut kayıtlar veri/arsiv klasörüne taşınır
 *   POST api.php  {"action":"list","name":..,"by":..,"data":"<base64>"} → ürün listesini (Excel) yükler
 *   POST api.php  {"action":"list_reset"}  → yüklenen listeyi kaldırır (sitedeki varsayılan liste kullanılır)
 *   GET  api.php?liste=1                   → yüklenen ürün listesini (Excel) indirir
 */

header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store, no-cache, must-revalidate');
header('X-Content-Type-Options: nosniff');

$DIR = __DIR__ . '/veri';
$LOG = $DIR . '/olaylar.jsonl';
$EPOCH = $DIR . '/sayim.txt';
$LOCK = $DIR . '/kilit';
$LIST = $DIR . '/urun-listesi.xlsx';
$LIST_META = $DIR . '/liste.json';
$PAGE = 1048576;          // bir yanıtta en fazla 1 MB kayıt
$MAX_BODY = 2097152;      // kayıt gönderiminde en fazla 2 MB
$MAX_LIST = 6291456;      // ürün listesi (Excel) en fazla 6 MB

function cevap($data, $code = 200)
{
    http_response_code($code);
    echo json_encode($data, JSON_UNESCAPED_UNICODE);
    exit;
}

if (!is_dir($DIR) && !@mkdir($DIR, 0775, true)) {
    cevap(array('ok' => false, 'error' => 'veri klasörü oluşturulamadı (yazma izni verin)'), 500);
}
if (!is_writable($DIR)) {
    cevap(array('ok' => false, 'error' => 'veri klasörüne yazma izni yok'), 500);
}
if (!file_exists($DIR . '/.htaccess')) {
    @file_put_contents($DIR . '/.htaccess',
        "<IfModule mod_authz_core.c>\n  Require all denied\n</IfModule>\n<IfModule !mod_authz_core.c>\n  Deny from all\n</IfModule>\n");
    @file_put_contents($DIR . '/index.html', '');
}

function sayim_no($file)
{
    $e = @file_get_contents($file);
    if ($e === false || trim($e) === '') {
        $e = bin2hex(random_bytes(8));
        file_put_contents($file, $e);
    }
    return trim($e);
}

function liste_bilgi($file)
{
    $j = @file_get_contents($file);
    if ($j === false) return null;
    $d = json_decode($j, true);
    return is_array($d) ? $d : null;
}

function kisa($s, $n)
{
    $s = preg_replace('/[\x00-\x1F\x7F]/u', '', (string)$s);
    if ($s === null) return '';
    return function_exists('mb_substr') ? mb_substr($s, 0, $n, 'UTF-8') : substr($s, 0, $n);
}

$lock = fopen($LOCK, 'c');
if (!$lock) cevap(array('ok' => false, 'error' => 'kilit dosyası açılamadı'), 500);

// ---------------------------------------------------------------- okuma
if ($_SERVER['REQUEST_METHOD'] === 'GET' && isset($_GET['liste'])) {
    flock($lock, LOCK_SH);
    if (!is_file($LIST)) {
        flock($lock, LOCK_UN);
        cevap(array('ok' => false, 'error' => 'yüklenmiş liste yok'), 404);
    }
    $bin = file_get_contents($LIST);
    flock($lock, LOCK_UN);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="urun-listesi.xlsx"');
    header('Content-Length: ' . strlen($bin));
    echo $bin;
    exit;
}

if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    flock($lock, LOCK_SH);
    $epoch = sayim_no($EPOCH);
    clearstatcache();
    $size = file_exists($LOG) ? filesize($LOG) : 0;
    $since = isset($_GET['since']) ? max(0, (int)$_GET['since']) : 0;
    $lines = array();
    $next = $size;
    if ($since < $size) {
        $h = fopen($LOG, 'rb');
        fseek($h, $since);
        $data = fread($h, min($PAGE, $size - $since));
        fclose($h);
        $cut = strrpos($data, "\n");
        $data = $cut === false ? '' : substr($data, 0, $cut + 1);
        $next = $since + strlen($data);
        foreach (explode("\n", $data) as $line) {
            if ($line !== '') $lines[] = $line; // satırlar yazılırken doğrulanmış JSON
        }
    }
    $list = liste_bilgi($LIST_META);
    flock($lock, LOCK_UN);
    echo '{"ok":true,"epoch":' . json_encode($epoch) . ',"next":' . $next . ',"more":' . ($next < $size ? 'true' : 'false')
        . ',"list":' . json_encode($list, JSON_UNESCAPED_UNICODE)
        . ',"events":[' . implode(',', $lines) . ']}';
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') cevap(array('ok' => false, 'error' => 'yöntem desteklenmiyor'), 405);

// ---------------------------------------------------------------- yazma
$limit = (int)ceil($MAX_LIST * 4 / 3) + 4096;
$raw = file_get_contents('php://input', false, null, 0, $limit + 1);
if (strlen($raw) > $limit) cevap(array('ok' => false, 'error' => 'dosya çok büyük (en fazla 6 MB)'), 413);
$body = json_decode($raw, true);
if (!is_array($body)) cevap(array('ok' => false, 'error' => 'bad_json'), 400);
$action = isset($body['action']) ? (string)$body['action'] : '';
if ($action !== 'list' && strlen($raw) > $MAX_BODY) cevap(array('ok' => false, 'error' => 'istek çok büyük'), 413);

flock($lock, LOCK_EX);
$epoch = sayim_no($EPOCH);

// Ürün listesi (Excel) yükleme: herkes bu listeyi kullanır
if ($action === 'list') {
    $bin = base64_decode(isset($body['data']) ? (string)$body['data'] : '', true);
    $ok = $bin !== false && strlen($bin) > 100 && strlen($bin) <= $MAX_LIST
        && (substr($bin, 0, 4) === "PK\x03\x04" || substr($bin, 0, 8) === "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1");
    if (!$ok) {
        flock($lock, LOCK_UN);
        cevap(array('ok' => false, 'error' => 'geçerli bir Excel dosyası değil'), 400);
    }
    file_put_contents($LIST . '.tmp', $bin);
    rename($LIST . '.tmp', $LIST);
    $meta = array(
        'version' => bin2hex(random_bytes(6)),
        'name' => kisa(basename(isset($body['name']) ? (string)$body['name'] : 'urun-listesi.xlsx'), 120),
        'by' => kisa(isset($body['by']) ? $body['by'] : '', 60),
        't' => date('c'),
        'size' => strlen($bin),
    );
    file_put_contents($LIST_META, json_encode($meta, JSON_UNESCAPED_UNICODE));
    flock($lock, LOCK_UN);
    cevap(array('ok' => true, 'list' => $meta));
}

if ($action === 'list_reset') {
    @unlink($LIST);
    @unlink($LIST_META);
    flock($lock, LOCK_UN);
    cevap(array('ok' => true, 'list' => null));
}

// Yeni sayım: mevcut kayıtlar arşive taşınır, tüm cihazlar sıfırlanır
if ($action === 'reset') {
    clearstatcache();
    if (file_exists($LOG) && filesize($LOG) > 0) {
        if (!is_dir($DIR . '/arsiv')) @mkdir($DIR . '/arsiv', 0775, true);
        rename($LOG, $DIR . '/arsiv/olaylar-' . date('Y-m-d-His') . '.jsonl');
    }
    $epoch = bin2hex(random_bytes(8));
    file_put_contents($EPOCH, $epoch);
    flock($lock, LOCK_UN);
    cevap(array('ok' => true, 'epoch' => $epoch));
}

if (!empty($body['epoch']) && $body['epoch'] !== $epoch) {
    flock($lock, LOCK_UN);
    cevap(array('ok' => false, 'error' => 'epoch', 'epoch' => $epoch));
}

$incoming = isset($body['events']) && is_array($body['events']) ? array_slice($body['events'], 0, 500) : array();

// Aynı kayıt tekrar gönderilirse (bağlantı kopması vb.) iki kez yazılmasın: son 2 MB'taki kimlikler
$seen = array();
clearstatcache();
$size = file_exists($LOG) ? filesize($LOG) : 0;
if ($size > 0) {
    $h = fopen($LOG, 'rb');
    $from = max(0, $size - 2097152);
    fseek($h, $from);
    $tail = fread($h, $size - $from);
    fclose($h);
    if (preg_match_all('/"id":"([A-Za-z0-9_-]{1,64})"/', $tail, $m)) {
        foreach ($m[1] as $id) $seen[$id] = true;
    }
}

$out = '';
$added = 0;
foreach ($incoming as $ev) {
    if (!is_array($ev) || !isset($ev['id'], $ev['t'], $ev['type'])) continue;
    $id = (string)$ev['id'];
    if (!preg_match('/^[A-Za-z0-9_-]{1,64}$/', $id) || isset($seen[$id])) continue;
    if (!is_string($ev['type']) || !preg_match('/^[a-z_]{1,24}$/', $ev['type'])) continue;
    $line = json_encode(array('id' => $id) + $ev, JSON_UNESCAPED_UNICODE);
    if ($line === false || strlen($line) > 8000) continue;
    $seen[$id] = true;
    $out .= $line . "\n";
    $added++;
}
if ($out !== '') file_put_contents($LOG, $out, FILE_APPEND);
flock($lock, LOCK_UN);
cevap(array('ok' => true, 'epoch' => $epoch, 'added' => $added));
