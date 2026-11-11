# TruthLens AI

Brauzerda ishlaydigan raqamli media skrining vositasi (rasm, video, audio). Barcha tahlil mijoz tomonida bajariladi — fayllar serverga yuborilmaydi.

## Vercel'ga deploy

1. [vercel.com/new](https://vercel.com/new) → **Import Git Repository** → `billsgm78-blip/trustai`.
2. **Framework Preset:** `Other`. Build Command va Output Directory bo'sh qoldiring.
3. **Deploy** tugmasini bosing.

Bu statik sayt (`index.html`), hech qanday build kerak emas. `main` branchga har bir push avtomatik qayta deploy qilinadi.

## Lokal ishga tushirish

```bash
npx serve .
```
