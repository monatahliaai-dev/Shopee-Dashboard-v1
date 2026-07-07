import json, os, subprocess, sys
from pathlib import Path
from flask import Flask, request, jsonify
import pandas as pd
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
BASE = Path(__file__).parent.parent
HPP = BASE / 'hpp.xlsx'
UPLOAD_LOG = BASE / 'upload_log.json'
COLS = ['SKU Internal','Nama Produk','Varian','Supplier','Kategori','HPP','Harga Jual Normal','Harga Minimum Aman','Berat Gram','Status','Catatan']
UPLOAD_MAP = {
    'Hiban Store': BASE,
    'Royal Abiya': BASE / 'royal_abiya' / 'uploaded',
    'Hiban Signature': BASE / 'hiban_signature' / 'uploaded',
}

def is_vercel():
    return os.environ.get('VERCEL') == '1'

def load_upload_log():
    if not UPLOAD_LOG.exists(): return []
    try: return json.loads(UPLOAD_LOG.read_text(encoding='utf-8'))
    except Exception: return []

def append_upload_log(store, saved):
    if is_vercel(): return load_upload_log()
    rows = load_upload_log()
    now = datetime.now().isoformat(timespec='seconds')
    for path in saved:
        rows.insert(0, {'waktu': now, 'toko': store, 'file': path})
    UPLOAD_LOG.write_text(json.dumps(rows[:1000], ensure_ascii=False, indent=2), encoding='utf-8')
    return rows[:1000]

def upsert_hpp_rows(path, rows):
    if is_vercel(): return {'updated': 0, 'added': 0, 'error': 'Read-only on Vercel. Please run locally to update HPP.'}
    df = pd.read_excel(path, sheet_name='Utama') if Path(path).exists() else pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns: df[c] = None
    added = updated = 0
    for r in rows:
        sku = str(r.get('sku','')).strip()
        if not sku: continue
        hpp = float(r.get('hpp') or 0)
        if hpp <= 0: continue
        vals = {'SKU Internal':sku,'Nama Produk':r.get('nama',''),'Varian':r.get('varian',''),'HPP':hpp,'Status':'Aktif'}
        hit = df['SKU Internal'].astype(str).eq(sku)
        if hit.any():
            for k,v in vals.items(): df.loc[hit,k] = v
            updated += 1
        else:
            df = pd.concat([df, pd.DataFrame([{**{c:None for c in COLS}, **vals}])], ignore_index=True)
            added += 1
    with pd.ExcelWriter(path, engine='openpyxl') as w:
        df[COLS].to_excel(w, sheet_name='Utama', index=False)
    return {'updated': updated, 'added': added}

def rebuild():
    if is_vercel(): return
    subprocess.check_call([sys.executable, 'build_bookkeeping.py'], cwd=BASE)

@app.route('/api/upload-log', methods=['GET'])
def api_upload_log():
    return jsonify({'ok': True, 'rows': load_upload_log()[:200]})

@app.route('/api/hpp', methods=['POST'])
def api_hpp():
    try:
        data = request.json or {}
        rows = data.get('rows', [])
        result = upsert_hpp_rows(HPP, rows)
        if 'error' in result:
            return jsonify({'ok': False, 'error': result['error']}), 400
        rebuild()
        return jsonify({'ok': True, **result})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/api/upload', methods=['POST'])
def api_upload():
    try:
        if is_vercel():
            return jsonify({'ok': False, 'error': 'Read-only on Vercel. Please run locally to upload files.'}), 400
        store = request.form.get('store', '')
        target = UPLOAD_MAP.get(store)
        if not target: return jsonify({'ok': False, 'error': 'Toko teu valid'}), 400
        target.mkdir(parents=True, exist_ok=True)
        saved = []
        files = request.files.getlist('files')
        for file in files:
            if not file.filename: continue
            name = secure_filename(file.filename) or 'upload.bin'
            dest = target / name
            file.save(dest)
            saved.append(str(dest.relative_to(BASE)))
        if not saved: return jsonify({'ok': False, 'error': 'File kosong'}), 400
        rebuild()
        log = append_upload_log(store, saved)
        return jsonify({'ok': True, 'saved': saved, 'log': log[:200]})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

# Endpoint fallback (for Vercel serverless)
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def catch_all(path):
    return jsonify({'error': 'Not Found'}), 404
