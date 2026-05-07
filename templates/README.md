# Kupa — modüler yapı

## Klasörler

- `templates/pages/` — sayfa template'leri (sadece include direktifi listesi)
- `templates/partials/` — section'lar (hero, show-cards, timeline, ...)
- `assets/styles.css` — paylaşımlı CSS (her iki sayfa için tek kaynak)
- `assets/img/` — paylaşımlı görseller
- `index.html`, `detay/index.html` — **derlenmiş çıktı** (commit edilir, GitHub Pages bunları serve eder)
- `build.py` — template'leri çıktıya dönüştürür
- `.nojekyll` — GitHub Pages'in Jekyll çalıştırmasını engeller (built dosyalar olduğu gibi servis edilsin)

## Workflow

```
# template/partial düzenle
nano templates/partials/duo-accordion.html

# derle (index.html + detay/index.html güncellenir)
python build.py

# commit + push
git add . && git commit -m "..." && git push
```

## Differ dosyalar

Bu section'lar iki sayfada FARKLI olduğu için her birinin `-index.html` ve
`-detay.html` varyantı var:

        - `templates/partials/footer-logo-{index,detay}.html`
        - `templates/partials/grand-total-{index,detay}.html`
        - `templates/partials/hero-{index,detay}.html`
        - `templates/partials/integration-types-{index,detay}.html`
        - `templates/partials/show-cards-{index,detay}.html`
        - `templates/partials/summary-totals-{index,detay}.html`
        - `templates/partials/timeline-{index,detay}.html`
