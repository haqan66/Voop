# Barkod Kontrol

Ürün barkodunu okutup sabit Excel ürün listesiyle eşleştiren web uygulaması. Aynı uygulama hem **telefonda** hem
**bilgisayarda (web tarayıcısında)** çalışır; **2–3 kişi aynı anda** okutabilir, herkes diğerlerinin okuttuklarını birkaç saniye içinde görür.

**Web adresi (GitHub Pages açıldıktan sonra):** https://haqan66.github.io/Voop/

## Nasıl çalışır?

1. **Ürün listesi sabittir** – `data/urun-listesi.xlsx` uygulamanın içindedir ve uygulamadan değiştirilemez.
   Listeyi güncellemek gerekirse bu dosya (ve `data/urun-listesi.js`, aşağıya bakın) depoda değiştirilir.
2. **Barkod okutulur** – telefon kamerası, bilgisayar kamerası, USB / el barkod okuyucu veya elle giriş.
   - Listede varsa: yeşil **EŞLEŞİYOR ✓**, ürün adı, bip sesi ve titreşim.
   - 12 haneli UPC barkodlar (ör. Ocean ürünleri) 13 haneli `0…` okumasıyla da eşleşir.
3. **Eşleşme yoksa** – kırmızı **EŞLEŞME YOK** paneli açılır, ürün adıyla aranır (Türkçe karakter duyarsız).
   Bulunan ürüne dokunup onaylayınca barkod o ürünün **Yeni Barkod** sütununa eklenir; ekipteki herkeste hemen eşleşir.
4. **İsimle de bulunamazsa** – **Bulunamadı** işaretlenir, ürün adı + barkod girilerek **Yeni Ürün** eklenir.

## Ortak çalışma (2–3 kişi aynı anda)

Siteye giren herkes **sadece adını yazıp** okutmaya başlar; ayar yapmaz. Okutmalar ücretsiz bir **Google E-Tablosu**
üzerinden ekipteki tüm cihazlara birkaç saniyede yansır ve E-Tablo’daki sabit ürün listesi **anında güncellenir**:

| E-Tablo sayfası | İçerik |
|---|---|
| **Ürün Listesi** | Sabit Excel listesi; **Yeni Barkod**, **Kontrol Durumu**, **Kontrol Tarihi**, **Kontrol Eden** sütunları her okutmada güncellenir |
| **Yeni Ürünler** | Listede olmayan / bulunamayan ürünler, ekleyen kişiyle |
| **Olaylar** | Tüm işlemler satır satır (tarih, kişi, işlem, barkod, ürün) |

Güncel Excel’i E-Tablo’dan **Dosya → İndir → Microsoft Excel (.xlsx)** ile ya da uygulamadan **Excel İndir** ile alabilirsiniz.

### Kurulum (bir kez, ~5 dakika — yalnızca yönetici)

1. https://sheets.new ile yeni bir Google E-Tablosu açın, adını örn. *Barkod Kontrol Ortak* yapın.
2. **Uzantılar → Apps Script**. Açılan dosyadaki her şeyi silin, depodaki
   [`google-apps-script/Kod.gs`](google-apps-script/Kod.gs) dosyasının tamamını yapıştırıp **Kaydet**’e basın.
3. **Dağıt → Yeni dağıtım** → tür olarak **Web uygulaması** seçin:
   - *Şu kullanıcı olarak yürüt*: **Ben**
   - *Erişimi olanlar*: **Herkes**
   → **Dağıt** → Google’ın istediği izinleri onaylayın (“Gelişmiş → … sayfasına git” adımı çıkabilir; betik sizin hesabınızda çalışır).
4. Çıkan **Web uygulaması URL’sini** (`https://script.google.com/macros/s/…/exec`) depodaki **`config.js`** dosyasına yazın:
   ```js
   ortakListeAdresi: 'https://script.google.com/macros/s/…/exec',
   ```
   (GitHub’da dosyayı açıp kalem simgesiyle düzenleyebilirsiniz.) Siteyi ilk açan kişide ürün listesi E-Tablo’ya otomatik yüklenir.

Bundan sonra site linkini ekibe göndermeniz yeterli. Üst bardaki gösterge durumu gösterir: **Ortak · 3 kişi** (bağlı),
**Gönderiliyor** (sarı), **Bağlantı yok** (kırmızı). İnternet kesilse de okutmaya devam edilir; kayıtlar bağlantı gelince gönderilir.

- Bir kişinin eklediği yeni barkod diğerlerinde de hemen eşleşir; aynı barkodu iki kişi aynı anda “Bulunamadı” işaretlese de tek kayıt oluşur.
- **Yeni sayım** için E-Tablo’da **Barkod Kontrol → Yeni sayım başlat**: mevcut liste ve olaylar arşiv sayfalarına kopyalanır,
  kontrol bilgileri temizlenir, tüm cihazlar birkaç saniye içinde sıfırlanır. *Olaylar* sayfasındaki satırları elle silmeyin.
- Site linkini bilen herkes okutma yapabilir (istenen davranış); linki yalnızca ekiple paylaşın.
- `Kod.gs` güncellenirse Apps Script’te **Dağıt → Dağıtımları yönet → Düzenle → Sürüm: Yeni sürüm** ile yeniden dağıtın (adres değişmez).

`config.js` boşsa uygulama tek cihazda çalışır; adres istenirse *Kayıtlar → Ortak Çalışma* bölümünden de girilebilir.

## Excel çıktısı

*Kayıtlar → Excel İndir* (telefonda *Paylaş* ile WhatsApp / e-posta / Drive’a gönderilebilir):

| Sayfa | İçerik |
|---|---|
| Ürün Stok Kartı Ana Dosya | Sabit liste + **Yeni Barkod**, **Kontrol Durumu**, **Kontrol Tarihi**, **Kontrol Eden** |
| Yeni Ürünler | Barkod, Ürün Adı, Marka, Durum (*Yeni Ürün* / *Bulunamadı*), Not, Tarih, Ekleyen |
| Tarama Geçmişi | Her işlemin tarihi, barkodu, sonucu, ürün adı ve okutan kişi |

Üstteki indirme ikonundaki sayı, son Excel indirmeden bu yana gelen değişiklik sayısıdır.
Yanlış eklenen barkod veya ürün *Kayıtlar* sekmesinden çöp kutusu ile geri alınabilir (ekipte herkes için geri alınır).

## Bilgisayar / web kullanımı

Geniş ekranda (1024 px ve üzeri) masaüstü düzeni açılır: sol menü, iki sütunlu tarama ekranı, sıralanabilir ürün tablosu.

- **USB / el barkod okuyucu:** hiçbir kutuya tıklamadan, sayfa açıkken okutmanız yeterli.
- **Bilgisayar kamerası:** “Kamerayı Başlat” ile açılır; birden fazla kamera varsa değiştirme düğmesi çıkar.
- **Kısayollar:** `/` arama kutusuna gider, `Esc` açık pencereyi kapatır.

## Web’de yayınlamak

Kamera izni için sayfanın **HTTPS** üzerinden açılması gerekir. GitHub Pages (ücretsiz):

1. https://github.com/haqan66/Voop/settings/pages
2. *Source*: **Deploy from a branch** → Branch: `claude/stoic-noether-81erq6`, klasör: `/ (root)` → **Save**
3. 1–2 dakika sonra uygulama **https://haqan66.github.io/Voop/** adresinde açılır.
4. Telefonda Chrome: menü → **Ana ekrana ekle** · Safari: Paylaş → **Ana Ekrana Ekle**

> Depo herkese açık (public) olduğu için `data/` klasöründeki ürün listesi de herkese açıktır.

Bilgisayarda yerel deneme: `index.html` dosyasına çift tıklayın veya `python3 -m http.server 8000` → `http://localhost:8000`.

## Ürün listesini güncellemek

`data/urun-listesi.xlsx` dosyasını yenisiyle değiştirin ve gömülü kopyayı yeniden üretin:

```
python3 -c "import base64;d=base64.b64encode(open('data/urun-listesi.xlsx','rb').read()).decode();open('data/urun-listesi.js','w').write(\"window.DEFAULT_XLSX_B64 = '\"+d+\"';\n\")"
```

Sayfada `Barkod` ve `Ürün Adı` başlıkları olması yeterlidir. Satır sırası değişirse önceki sayımın kayıtları yanlış
satırlara düşebileceği için listeyi yeni bir sayımın başında güncelleyin.

## Dosyalar

```
index.html                 arayüz
config.js                  ortak liste adresi (Google Apps Script)
app.css                    mobil + masaüstü tasarım (açık / koyu tema)
app.js                     okuma, eşleştirme, arama, ortak çalışma, Excel yazma
sw.js                      çevrimdışı çalışma
google-apps-script/Kod.gs  ortak liste için Google Apps Script
data/urun-listesi.xlsx     sabit ürün listesi
data/urun-listesi.js       aynı listenin gömülü kopyası (çift tıklayarak açmak için)
vendor/                    SheetJS (Excel) ve html5-qrcode (barkod okuma) kütüphaneleri
```
