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
 */

header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store, no-cache, must-revalidate');
header('X-Content-Type-Options: nosniff');

$DIR = __DIR__ . '/veri';
$LOG = $DIR . '/olaylar.jsonl';
$EPOCH = $DIR . '/sayim.txt';
$LOCK = $DIR . '/kilit';
$PAGE = 1048576;          // bir yanıtta en fazla 1 MB kayıt
$MAX_BODY = 2097152;      // en fazla 2 MB istek
$TYPES = array('scan', 'newbc', 'undo_newbc', 'notfound', 'newprod', 'del_newprod');

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

$lock = fopen($LOCK, 'c');
if (!$lock) cevap(array('ok' => false, 'error' => 'kilit dosyası açılamadı'), 500);

// ---------------------------------------------------------------- okuma
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
    flock($lock, LOCK_UN);
    echo '{"ok":true,"epoch":' . json_encode($epoch) . ',"next":' . $next . ',"more":' . ($next < $size ? 'true' : 'false')
        . ',"events":[' . implode(',', $lines) . ']}';
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') cevap(array('ok' => false, 'error' => 'yöntem desteklenmiyor'), 405);

// ---------------------------------------------------------------- yazma
$raw = file_get_contents('php://input', false, null, 0, $MAX_BODY + 1);
if (strlen($raw) > $MAX_BODY) cevap(array('ok' => false, 'error' => 'istek çok büyük'), 413);
$body = json_decode($raw, true);
if (!is_array($body)) cevap(array('ok' => false, 'error' => 'bad_json'), 400);

flock($lock, LOCK_EX);
$epoch = sayim_no($EPOCH);

// Yeni sayım: mevcut kayıtlar arşive taşınır, tüm cihazlar sıfırlanır
if (isset($body['action']) && $body['action'] === 'reset') {
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
    if (!in_array($ev['type'], $TYPES, true)) continue;
    $line = json_encode(array('id' => $id) + $ev, JSON_UNESCAPED_UNICODE);
    if ($line === false || strlen($line) > 8000) continue;
    $seen[$id] = true;
    $out .= $line . "\n";
    $added++;
}
if ($out !== '') file_put_contents($LOG, $out, FILE_APPEND);
flock($lock, LOCK_UN);
cevap(array('ok' => true, 'epoch' => $epoch, 'added' => $added));
