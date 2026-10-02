# Barkod Kontrol

Ürün barkodunu okutup sabit Excel ürün listesiyle eşleştiren web uygulaması. Kendi hostinginize yüklenir;
siteye giren kişi **sadece adını yazıp** okutmaya başlar. **2–3 kişi aynı anda** okutabilir, herkesin okuttuğu
birkaç saniye içinde diğer ekranlara ve Excel çıktısına yansır. Telefonda ve bilgisayarda çalışır.

## Nasıl çalışır?

1. **Ürün listesi sabittir** – `data/urun-listesi.xlsx` sitenin içindedir.
2. **Adınızı yazın** – siteye ilk girişte bir kez sorulur; okuttuğunuz her üründe bu isim görünür.
3. **Barkod okutun** – telefon kamerası, bilgisayar kamerası, USB / el barkod okuyucu veya elle giriş.
   - Listede varsa: yeşil **EŞLEŞİYOR ✓**, ürün adı, bip sesi ve titreşim.
   - 12 haneli UPC barkodlar (ör. Ocean ürünleri) 13 haneli `0…` okumasıyla da eşleşir.
4. **Eşleşme yoksa** – kırmızı **EŞLEŞME YOK** paneli açılır, ürün adıyla aranır. Bulunan ürüne dokunup onaylayınca
   barkod o ürünün **Yeni Barkod** sütununa eklenir; ekipteki herkeste hemen eşleşir.
5. **İsimle de bulunamazsa** – **Bulunamadı** işaretlenir, ürün adı + barkod girilerek **Yeni Ürün** eklenir.

## Hostinge yükleme

Gereken: **PHP** destekli bir hosting (cPanel / Plesk vb. paylaşımlı hostinglerin hepsinde vardır) ve **SSL (https)**.
Telefon kamerası yalnızca `https://` adreslerde açılır; hostinginizin ücretsiz SSL’ini (Let’s Encrypt) açın. Veritabanı gerekmez.

1. Arşivdeki **tüm dosya ve klasörleri** hostingde sitenin klasörüne yükleyin (ör. `public_html/barkod/`).
   cPanel’de *Dosya Yöneticisi → Yükle* ile RAR/ZIP’i yükleyip *Çıkart* diyebilirsiniz.
2. Tarayıcıda `https://alanadiniz.com/barkod/` adresini açın, adınızı yazın.
   Üst barda **Ortak · 1 kişi** görünüyorsa kurulum tamamdır.
3. Bu linki ekibe gönderin. Telefonda Chrome: menü → **Ana ekrana ekle** · Safari: Paylaş → **Ana Ekrana Ekle**.

Kayıtlar sunucuda otomatik oluşturulan **`veri/`** klasöründe tutulur ve bu klasör dışarıya kapatılır
(`.htaccess`). Üst bar **Bağlantı yok** veya “yazma izni yok” diyorsa, `veri` klasörüne (yoksa site klasörüne)
yazma izni verin (cPanel → izinler → 755).

> Nginx kullanan bir sunucudaysanız `.htaccess` çalışmaz; `veri/` klasörüne erişimi sunucu ayarından kapatın.

## Birlikte çalışma

- Üst bardaki gösterge: **Ortak · 3 kişi** (bağlı), **Gönderiliyor** (sarı), **Bağlantı yok** (kırmızı).
- İnternet kesilse de okutmaya devam edilir; kayıtlar bağlantı gelince otomatik gönderilir.
- Bir kişinin eklediği yeni barkod diğerlerinde de hemen eşleşir; aynı barkodu iki kişi aynı anda “Bulunamadı”
  işaretlese de tek kayıt oluşur.
- **Excel İndir** herhangi bir cihazdan yapılabilir; ekibin tüm kayıtlarını içerir.
- **Yeni sayım:** *Kayıtlar → Yeni Sayım Başlat*. Herkesin kayıtları temizlenir, tüm cihazlar sıfırlanır;
  eski kayıtlar sunucuda `veri/arsiv/` klasörüne yedeklenir.
- Site linkini bilen herkes okutma yapabilir; linki yalnızca ekiple paylaşın.

## Excel çıktısı

*Kayıtlar → Excel İndir* (telefonda *Paylaş* ile WhatsApp / e-posta / Drive’a gönderilebilir):

| Sayfa | İçerik |
|---|---|
| Ürün Stok Kartı Ana Dosya | Sabit liste + **Yeni Barkod**, **Kontrol Durumu**, **Kontrol Tarihi**, **Kontrol Eden** |
| Yeni Ürünler | Barkod, Ürün Adı, Marka, Durum (*Yeni Ürün* / *Bulunamadı*), Not, Tarih, Ekleyen |
| Tarama Geçmişi | Her işlemin tarihi, barkodu, sonucu, ürün adı ve okutan kişi |

Yanlış eklenen barkod veya ürün *Kayıtlar* sekmesinden çöp kutusu ile geri alınabilir (ekipte herkes için geri alınır).

## Bilgisayar kullanımı

Geniş ekranda masaüstü düzeni açılır: sol menü, iki sütunlu tarama ekranı, sıralanabilir ürün tablosu.

- **USB / el barkod okuyucu:** imleç her zaman barkod kutusunda durur; art arda okutmanız yeterli, hiçbir yere
  tıklamanız gerekmez. Okuyucu barkod sonuna Enter veya Tab eklese de çalışır. Eşleşme yok penceresi kapanınca
  imleç kendiliğinden barkod kutusuna döner.
- **Bilgisayar kamerası:** “Kamerayı Başlat” ile açılır; birden fazla kamera varsa değiştirme düğmesi çıkar.
- **Kısayollar:** `/` arama kutusuna gider, `Esc` açık pencereyi kapatır.

`index.html` dosyasına çift tıklayarak da açılabilir; bu durumda sunucu olmadığı için tek cihaz olarak çalışır.

## Ürün listesini güncellemek

`data/urun-listesi.xlsx` dosyasını yenisiyle değiştirin ve gömülü kopyayı yeniden üretin:

```
python3 -c "import base64;d=base64.b64encode(open('data/urun-listesi.xlsx','rb').read()).decode();open('data/urun-listesi.js','w').write(\"window.DEFAULT_XLSX_B64 = '\"+d+\"';\n\")"
```

Kayıtlar listedeki satır sırasına göre tutulduğundan listeyi yeni bir sayımın başında güncelleyin.

## Dosyalar

```
index.html               arayüz
app.css                  mobil + masaüstü tasarım
app.js                   okuma, eşleştirme, arama, ortak çalışma, Excel yazma
api.php                  ortak kayıt sunucusu (PHP, veritabanı gerekmez)
sw.js                    çevrimdışı çalışma
data/urun-listesi.xlsx   sabit ürün listesi
data/urun-listesi.js     aynı listenin gömülü kopyası (çift tıklayarak açmak için)
vendor/                  SheetJS (Excel) ve html5-qrcode (barkod okuma) kütüphaneleri
veri/                    (sunucuda otomatik oluşur) ortak kayıtlar ve arşiv
```
