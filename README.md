# Barkod Kontrol

Telefon kamerasıyla ürün barkodu okutup Excel ürün listesiyle eşleştiren mobil web uygulaması.
Kurulum gerektirmez; telefonda tarayıcıdan açılır, "Ana ekrana ekle" ile uygulama gibi kullanılır ve internet olmadan da çalışır.

## Nasıl çalışır?

1. **Ürün listesi okunur** – `data/urun-listesi.xlsx` otomatik yüklenir. Başka bir liste için *Kayıtlar → Excel Yükle*.
   Sayfada `Barkod` ve `Ürün Adı` başlıkları olması yeterli; `Yeni Barkod` sütunu yoksa çıktıda eklenir.
2. **Barkod okutulur** – Kamera açılır, barkod çerçeveye tutulur.
   - Listede varsa: yeşil **EŞLEŞİYOR ✓** ekranı, ürün adı, bip sesi ve titreşim. Kamera açık kalır, sıradaki ürüne geçilir.
   - 12 haneli UPC barkodlar (ör. Ocean ürünleri) telefonun 13 haneli `0…` okumasıyla da eşleşir.
3. **Eşleşme yoksa** – Kırmızı **EŞLEŞME YOK** paneli açılır ve ürün adıyla arama yapılır
   (Türkçe karakter duyarsız: `vitamin` = `VİTAMİN`, birden fazla kelime yazılabilir).
   Bulunan ürüne dokunup onaylayınca okutulan barkod o ürünün **Yeni Barkod** sütununa yazılır.
   Bu barkod bir dahaki okutmada otomatik eşleşir.
4. **İsimle de bulunamazsa** – *“Ürün listede yok — Bulunamadı işaretle”* ile ürün **Bulunamadı** olarak işaretlenir,
   ardından ürün adı + barkod (marka/not isteğe bağlı) girilerek **Yeni Ürün** eklenir.

Barkod elle de girilebilir (kamera açılmazsa veya barkod okunmazsa).

## Excel çıktısı

*Kayıtlar → Excel İndir* (veya *Paylaş* ile WhatsApp / e-posta / Drive'a gönder):

| Sayfa | İçerik |
|---|---|
| Ürün Stok Kartı Ana Dosya | Orijinal liste + doldurulan **Yeni Barkod**, ayrıca **Kontrol Durumu** ve **Kontrol Tarihi** sütunları |
| Yeni Ürünler | Barkod, Ürün Adı, Marka, Durum (*Yeni Ürün* / *Bulunamadı*), Not, Tarih |
| Tarama Geçmişi | Her okutmanın tarihi, barkodu, sonucu ve ürün adı |

Dışa aktarılan dosya tekrar yüklendiğinde yeni barkodlar, yeni ürünler ve geçmiş geri okunur, kaldığınız yerden devam edilir.

Yapılan işlemler telefonda saklanır (sayfa kapansa da kaybolmaz). Üstteki indirme ikonundaki sayı, henüz Excel'e
aktarılmamış değişiklik sayısıdır. Yanlış eklenen barkod veya ürün *Kayıtlar* sekmesinden çöp kutusu ile geri alınabilir.

*Ürünler* sekmesinde tüm liste aranabilir; *Kontrol edilmedi / Kontrol edildi / Yeni barkodlu / Barkodsuz* filtreleriyle
sayım takibi yapılabilir, ürüne dokunarak elle barkod eklenebilir.

## Bilgisayarda (yerel) çalıştırmak

Klasördeki **`index.html`** dosyasına çift tıklayın (Chrome veya Edge önerilir). Sunucu veya kurulum gerekmez;
ürün listesi `data/urun-listesi.js` içindeki gömülü kopyadan yüklenir. Bilgisayara bağlı / dahili kamera ile okutma yapılabilir,
USB barkod okuyucu kullanılıyorsa “Barkodu elle girin” kutusuna tıklayıp okutmanız yeterli (okuyucu Enter gönderir).

> Not: Telefondan aynı Wi-Fi üzerinden `http://192.168…` adresiyle açarsanız tarayıcı kamerayı engeller; telefon için aşağıdaki HTTPS yöntemini kullanın.

## Telefonda açmak

Kamera izni için sayfanın **HTTPS** üzerinden açılması gerekir. En kolay yol GitHub Pages:

1. GitHub'da depo → **Settings → Pages**
2. *Source*: **Deploy from a branch**, branch olarak bu dalı ve `/ (root)` klasörünü seçip **Save**
3. Birkaç dakika sonra verilen `https://<kullanıcı>.github.io/<depo>/` adresini telefonda açın
4. Chrome: menü → **Ana ekrana ekle** · Safari: Paylaş → **Ana Ekrana Ekle**

Alternatif: klasörü [Netlify Drop](https://app.netlify.com/drop) sayfasına sürükleyip bırakmak da HTTPS adres verir.

Bilgisayarda denemek için: `python3 -m http.server 8000` → `http://localhost:8000`

## Dosyalar

```
index.html             arayüz
app.css                mobil tasarım (açık / koyu tema)
app.js                 okuma, eşleştirme, arama, Excel yazma
sw.js                  çevrimdışı çalışma
data/urun-listesi.xlsx varsayılan ürün listesi
data/urun-listesi.js   aynı listenin gömülü kopyası (çift tıklayarak açmak için)
vendor/                SheetJS (Excel) ve html5-qrcode (barkod okuma) kütüphaneleri
```
