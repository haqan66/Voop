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

Okutmalar ücretsiz bir **Google E-Tablosu** üzerinden ekipteki tüm cihazlara dağıtılır. Sunucu kurmaya gerek yoktur.
Excel ürün listesi değişmez; tüm okutmalar ayrıca E-Tablo’daki **Olaylar** sayfasında satır satır görülebilir.

### Kurulum (bir kez, ~5 dakika — bir kişi yapar)

1. https://sheets.new ile yeni bir Google E-Tablosu açın, adını örn. *Barkod Kontrol Ortak* yapın.
2. **Uzantılar → Apps Script**. Açılan dosyadaki her şeyi silin, depodaki
   [`google-apps-script/Kod.gs`](google-apps-script/Kod.gs) dosyasının tamamını yapıştırıp **Kaydet**’e basın.
3. **Dağıt → Yeni dağıtım** → tür olarak **Web uygulaması** seçin:
   - *Şu kullanıcı olarak yürüt*: **Ben**
   - *Erişimi olanlar*: **Herkes**
   → **Dağıt** → Google’ın istediği izinleri onaylayın (“Gelişmiş → … sayfasına git (güvenli değil)” adımı çıkabilir; betik sizin hesabınızda çalışır).
4. Çıkan **Web uygulaması URL’sini** (`https://script.google.com/macros/s/…/exec`) kopyalayın.
5. Uygulamada **Kayıtlar → Ortak Çalışma**: adınızı yazın, adresi yapıştırın, **Bağlan**.
6. **Ekip linkini paylaş** ile linki WhatsApp vb. ile ekip arkadaşlarınıza gönderin. Linki açan telefon/bilgisayar
   otomatik bağlanır ve bir kez adını sorar.

Üst bardaki gösterge durumu gösterir: **Ortak · 3 kişi** (bağlı), **Gönderiliyor** (sarı), **Bağlantı yok** (kırmızı).
İnternet kesilse de okutmaya devam edilir; kayıtlar bağlantı gelince otomatik gönderilir.

- Herkes aynı listede çalışır: bir kişinin eklediği yeni barkod diğerlerinde de eşleşir, aynı barkodu iki kişi aynı anda
  “Bulunamadı” işaretlese de tek kayıt oluşur.
- **Excel İndir** herhangi bir cihazdan yapılabilir; ekibin tüm kayıtlarını içerir.
- **Yeni sayım başlatmak** için E-Tablo’da **Barkod Kontrol → Yeni sayım başlat** menüsünü kullanın: mevcut kayıtlar
  arşiv sayfasına taşınır, tüm cihazlar birkaç saniye içinde sıfırlanır. *Olaylar* sayfasındaki satırları elle silmeyin.
- Bağlantı adresini bilen herkes kayıt ekleyebilir; linki yalnızca ekiple paylaşın.
- `Kod.gs` güncellenirse Apps Script’te **Dağıt → Dağıtımları yönet → Düzenle → Yeni sürüm** ile yeniden dağıtın (adres değişmez).

Ortak çalışma kurulmazsa uygulama tek cihazda çalışır; kayıtlar o cihazın tarayıcısında saklanır.

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
app.css                    mobil + masaüstü tasarım (açık / koyu tema)
app.js                     okuma, eşleştirme, arama, ortak çalışma, Excel yazma
sw.js                      çevrimdışı çalışma
google-apps-script/Kod.gs  ortak liste için Google Apps Script
data/urun-listesi.xlsx     sabit ürün listesi
data/urun-listesi.js       aynı listenin gömülü kopyası (çift tıklayarak açmak için)
vendor/                    SheetJS (Excel) ve html5-qrcode (barkod okuma) kütüphaneleri
```
