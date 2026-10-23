# Sipariş ve Tahsilat Defteri

Şarj aleti ve kablo siparişlerini, tahsilatları ve müşteri bakiyelerini takip eden web sitesi.

- **Canlı site:** https://claude.ai/artifact/CSYLd5oJoDAhDp9bwiopUh (claude.ai üzerinde çalışır, kayıtlar orada saklanır)
- **`index.html`:** Sitenin kodu. Tek dosya.
- **`araclar/`:** Excel aktarımı ve Excel raporları için Python betikleri.
  - `excel_aktar.py`: Eski cari hesap Excel'ini site kayıtlarına dönüştürür.
  - `musteri_listesi.py`: Müşteri listesi ve bakiyeler raporu.
  - `tahsilat_yillik.py`: Tahsilatları yıllara bölen rapor.

Müşteri, sipariş ve tahsilat verileri bu depoda **yoktur** (`.gitignore`). Veriler sitenin veritabanında durur.

## Sitenin bölümleri

Özet, Siparişler, Tahsilatlar (yıllara göre döküm dahil), Müşteriler (hesap ekstresi), Ürünler.
Para birimi USD; TL ödemeler kurla dolara çevrilerek girilir.
