# Kupa — proje notları

Statik tek sayfalık sunum sitesi. **GitHub Pages** + custom domain
**kupa.crossovertalks.com**. Backend yok.

## Yapı (modüler — Python build)

```
kupa/
├── build.py                    ← templates → HTML üretir + PDF/Excel
├── export_pdf_excel.py         ← reportlab + openpyxl
├── .nojekyll                   ← GitHub Pages Jekyll'i kapat
├── index.html                  ← BUILT artifact (commit'lenir)
├── detay/index.html            ← BUILT artifact, içerik index ile aynı
├── templates/
│   ├── pages/
│   │   ├── index.html          ← {{INCLUDE x.html}} listesi
│   │   └── detay.html          ← index ile birebir aynı içerik
│   └── partials/
│       ├── _head.html, _tail.html
│       ├── hero.html           ← TV noise + YouTube video sequence
│       ├── show-cards.html
│       ├── timeline.html       ← Kupa Duo + Masası akordiyon
│       ├── summary-totals-detay.html
│       ├── channel-stats.html
│       ├── integration-types.html
│       ├── extra-pricing-detay.html  ← Diğer Bütçeler + Paketler
│       ├── grand-total-detay.html
│       └── footer-logo.html    ← PDF + Excel download butonları
├── assets/
│   ├── styles.css              ← tek paylaşımlı CSS
│   ├── scripts.js              ← hero TV-noise sequence
│   ├── fonts/AllRoundGothic-*  ← D20 marka fontu (PDF için)
│   ├── img/                    ← entegrasyon thumbnail'ları + d20 logo
│   ├── kupa-sunum.pdf          ← BUILT (10 sayfa, 4:3 PowerPoint klasik)
│   └── kupa-paketler.xlsx      ← BUILT (Diğer Bütçeler + Paketler tabları)
├── header-logo.png             ← 6 kareli orijinal Crossover logo
├── Mavi-Crossover-Sports-Logo.png  ← PDF/Excel için tek logo
├── Kırmızı-Transparan.png      ← Excel'de D20 yerine
├── kupaduo1-son-sunum.jpeg     ← FAZ 1 show card thumb
└── kupamasasi1-son-sunum.jpeg  ← FAZ 2 show card thumb
```

## Workflow

```bash
# Tek partial veya CSS düzenle
nano templates/partials/timeline.html

# Derle (HTML + PDF + Excel hepsini günceller)
python build.py

# Her iki sayfa da aynı içerik — sadece PATH_PREFIX farkı
# (index: kök, detay: ../)

git add . && git commit -m "..." && git push
```

GitHub Pages 1-3 dk içinde deploy alır.

## Önemli noktalar

- **Index = Detay:** Müşteriye ikisi de gönderildiği için artık aynı içerik.
  Eğer ayrılması gerekirse `templates/pages/index.html`'i basitleştir.
- **`.gitignore`** `__pycache__/` ve `detay_index.html` (deliverable cruft)
  exclude eder.
- **Mobil TV-noise devre dışı:** `scripts.js` `(max-width: 900px)` altında
  return; CSS de `display: none !important` ile garanti. Mobil'de sadece
  logo. Desktop'ta hero'da TV noise → YouTube video (Bp88d4ioe5s) sequence
  oynuyor.
- **Bütçeler:** Hepsi "+KDV". Genel Toplam'da "ENTEGRASYON FORMATI" badge.
- **Paketler:** 4 paket (Bronz/Altın/Özel/Elmas). Gümüş + Advertorial silindi.
- **PDF/Excel butonları:** `footer-logo.html`'de, build her seferinde yenisini
  üretir.

## Path konvansiyonu

Tüm asset path'leri `{{PATH_PREFIX}}` token'ı ile yazılır. `build.py`
substitusyonu yapar:
- `index.html` → `PATH_PREFIX=""` (kök)
- `detay/index.html` → `PATH_PREFIX="../"`

YouTube/Google linkleri ve harici URL'ler token kullanmaz.

## Bilinen kararlar

- **Jekyll yerine Python build** seçildi: lokal `file://` preview'da Jekyll
  template'leri broken render olurdu, custom build script HTML üretip
  preview'da da doğru görünmesini sağlıyor.
- **weasyprint yerine reportlab:** Windows'ta GTK runtime gerektirmiyor.
- **AllRoundGothic'te ₺ glyph yok:** PDF'te ₺ kaldırıldı, sadece sayı + KDV.
  HTML/Excel'de ₺ korundu.
- **YouTube video Error 153 fix:** `referrerpolicy="strict-origin-when-cross-origin"`
  iframe'e ekli (YouTube 2025 sonu yeni şart).

## Last commit / state

Modüler yapı `7a90e36`'dan beri kararlı. Mobil video disable `d080fd8`,
mobil logo crop fix `061cd40`. Detayları `git log --oneline` ile gör.

## Local dev — http server

Mobil preview ve `file://` çözmediği path'ler için **port 8000'de Python
http.server** açık tutulmuş olabilir:

```bash
cd kupa
python -m http.server 8000
# → http://localhost:8000        (ana sayfa)
# → http://localhost:8000/detay/  (detay)
# → http://localhost:8000/assets/kupa-sunum.pdf
```

Production'da gereksiz; sadece geliştirme sırasında. Açıksa `Ctrl+C` ile
veya işlemi kill ederek kapat. Kapatmadan repo committable.

GitHub Pages canlı URL: **https://kupa.crossovertalks.com**
