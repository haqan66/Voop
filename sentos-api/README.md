# Sentos API

Kaynak: https://api.sentos.com.tr/docs (OpenAPI 3.0, sürüm 1.5.2). Tam şartname: [`openapi.yaml`](openapi.yaml).

## Temel bilgiler

| | |
|---|---|
| Adres | `https://firma-adi.sentos.com.tr/api` (dokümandaki sunucu: `https://api.sentos.com.tr/api`) |
| Kimlik doğrulama | Basic Auth — API anahtarı/şifresi Sentos paneli → **Hesabım** sayfasında |
| Header | `Content-Type: application/json` |
| Rate limit | Aynı gövdeyle GET: **2/dk**, POST: **12/dk** |
| Sayfalama | `page`, `size` parametreleri |

> Not: `GET /orders` için doküman "tüm parametreler opsiyonel ve **Request Body içinde** gönderilmeli" diyor.

## Sipariş durumları

`1` Onay Bekliyor · `2` Onaylandı · `3` Tedarik Sürecinde · `4` Hazırlanıyor · `5` Kargoya Verildi · `6` İptal Edildi · `99` Teslim Edildi

## Uç noktalar

### Ürünler
- `GET /products` — listele (`page`, `size`, `sku`, `barcode`, `orderby_id`)
- `GET /products/{id}` · `POST /products` · `PUT /products/{id}`
- `PUT /products/{id}/platform-variants` — varyant güncelle
- `POST /products/{id}/upload-image` — resim yükle
- `PUT /products/stocks` — **SKU/barkod ile toplu stok** (istek başına en fazla 100 kayıt; aynı anda tek istek, yoksa 429)

```json
[
  { "sku": "1357912-R7", "stocks": [{ "warehouse": 1, "stock": 150 }, { "warehouse": 2, "stock": 20 }] },
  { "barcode": "8690000000001", "stocks": { "warehouse": 1, "stock": 5 } }
]
```
`sku` önce varyant kodunda, sonra ürün kodunda aranır; ikisi birden gelirse `sku` kullanılır. Bulunamayanlar `errors` alanında döner, gövde hatalıysa tüm istek 422 ile reddedilir.

### Siparişler
- `GET /orders` — filtreler: `status`, `status_ids`, `start_date`/`end_date`, `created_*`, `updated_*`, `ship_due_*`, `order_id`, `order_code`, `orderby_field`, `orderby_direction`, `erp_transfer`, `erp_check_changed`, `is_updated`, `has_invoice`
- `GET /orders/{id}` · `POST /orders`
- `PUT /orders/{id}` — API ile oluşturulan siparişin durumu (5 → `cargo_company`, `tracking_number`, `tracking_url`; 6 → `add_stock`, `cancel_reason`, `cancel_date`; 99 → `delivery_date`)
- `PUT /orders/change-status` — toplu durum/kargo (en fazla 25, pazaryeri siparişlerinde çalışmaz, 5 ve 6 gönderilemez)
- Fatura: `POST|PUT|DELETE /orders/invoice/{id}`, `POST /orders/invoice-file/{id}`, `POST /orders/upload-invoice/{id}`, `POST|PUT /orders/invoice-binary/{id}` (base64 PDF)
- Kargo: `POST /orders/update-tracking/{id}`, `POST /orders/change-carrier/{id}`, `POST /orders/alternative-delivery/{id}`, `POST /orders/alternative-delivered/{id}`
- `GET /carriers` — platformlara göre kargo firması kodları

### Kategoriler
- `GET /categories` · `POST /categories` · `GET /categories/{id}` · `PUT /categories/{id}`

### Depolar
- `GET /warehouses` · `GET /warehouses/{id}`

### ERP
- `GET|POST|DELETE /erp/temp-products` — geçici ürünler
- `PUT /erp/order/{id}` — ERP aktarım bilgisi
- `GET /erp/matches?erp_code=` — eşleşmiş ürünler
- `PUT /erp/products` — `erp_code` ile `vat_rate`, `prices`, `stocks` güncelle
- `PUT /erp/products/price-stock` — entegrasyon fiyat/stok
- `POST /erp/erp-settings`

### Platformlar, Perakende, Satış Kanalları, İadeler
- `GET /platforms` · `GET|PUT /platforms/product/{id}`
- `GET /retails` (`page`, `size`, `sort`, `dir`) · `GET /retails/{id}`
- `GET|POST /sales-channels` · `PUT /sales-channels/{id}`
- `GET /claims` (tarih, sipariş, `platform`, `shop_id` filtreleri) · `GET /claims/{id}`
