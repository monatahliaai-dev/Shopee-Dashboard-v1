# Shopee Bookkeeping Dashboard

Dashboard pembukuan multi-toko Shopee dengan reconciliation HPP, chart analitik, dan upload data via web UI.

## Fitur

- **Multi-store support**: Hiban Store, Royal Abiya, Hiban Signature
- **Dashboard KPI**: Dana masuk, HPP, biaya iklan, laba bersih estimasi
- **Trend bulanan**: Chart interaktif 6 bulan terakhir
- **SKU belum HPP**: Form inline untuk update HPP langsung dari dashboard
- **Upload data**: Upload file order/income/ads Excel/CSV via UI
- **Export**: Consolidated workbook dengan ringkasan bulanan

## Tech Stack

- **Backend**: Python 3.9+ (pandas, openpyxl)
- **Frontend**: Bootstrap 5, Chart.js, vanilla JS
- **Server**: Python HTTP server dengan custom handler untuk writeback

## Setup

### 1. Install dependencies

```bash
python3 -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate   # Windows

pip install pandas openpyxl
```

### 2. Struktur file

```
hiban_shopee_bookkeeping/
├── build_bookkeeping.py   # Konsolidasi data → Excel + JSON
├── server.py              # HTTP server dengan HPP/upload API
├── hpp.xlsx               # Master HPP (SKU Internal + HPP)
├── web/
│   ├── index.html         # Dashboard UI
│   ├── data.json          # Generated data untuk frontend
│   └── assets/            # CSS/JS/fonts
├── order_*.xlsx           # Export order Shopee (Hiban Store)
├── income_*.xlsx          # Export income (dana dilepas)
├── ads_*.csv              # Export iklan Shopee
├── royal_abiya/           # Data Royal Abiya
└── hiban_signature/       # Data Hiban Signature
```

### 3. Jalankan

```bash
# Generate data
python build_bookkeeping.py

# Start server
python server.py
```

Buka: **http://127.0.0.1:8787/index.html**

## Workflow

1. **Upload data** via dashboard → pilih toko → upload file Excel/CSV
2. **HPP reconciliation** → klik SKU di tab "SKU Belum HPP" → isi HPP → Save
3. **Auto rebuild** → setiap save/upload otomatis rebuild workbook + regenerate `data.json`

## API Endpoints

- `GET /api/upload-log` → Riwayat upload (200 terakhir)
- `POST /api/hpp` → Upsert HPP rows ke `hpp.xlsx`
- `POST /api/upload` → Upload files ke folder toko

## Deploy ke Vercel

Dashboard ini **static frontend + Python backend**. Untuk Vercel:

1. **Frontend** → deploy `web/` folder sebagai static site
2. **Backend** → convert Python ke Vercel serverless function (Python runtime)

Atau alternatif:
- **Frontend** → Vercel
- **Backend** → Railway / Render / Fly.io (Python support)

## File yang di-ignore

File data source (.xlsx, .csv) **tidak** di-commit ke GitHub (kecuali `hpp.xlsx`).

Lihat `.gitignore` untuk daftar lengkap.

## License

MIT
