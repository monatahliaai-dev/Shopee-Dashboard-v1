import cgi, json, shutil, subprocess, sys
from datetime import datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import pandas as pd

BASE = Path(__file__).parent
WEB = BASE / 'web'
HPP = BASE / 'hpp.xlsx'
UPLOAD_LOG = BASE / 'upload_log.json'
COLS = ['SKU Internal','Nama Produk','Varian','Supplier','Kategori','HPP','Harga Jual Normal','Harga Minimum Aman','Berat Gram','Status','Catatan']
UPLOAD_MAP = {
    'Hiban Store': BASE,
    'Royal Abiya': BASE / 'royal_abiya' / 'uploaded',
    'Hiban Signature': BASE / 'hiban_signature' / 'uploaded',
}

def load_upload_log():
    if not UPLOAD_LOG.exists(): return []
    try: return json.loads(UPLOAD_LOG.read_text(encoding='utf-8'))
    except Exception: return []

def append_upload_log(store, saved):
    rows = load_upload_log()
    now = datetime.now().isoformat(timespec='seconds')
    for path in saved:
        rows.insert(0, {'waktu': now, 'toko': store, 'file': path})
    UPLOAD_LOG.write_text(json.dumps(rows[:1000], ensure_ascii=False, indent=2), encoding='utf-8')
    return rows[:1000]

def upsert_hpp_rows(path, rows):
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

def safe_name(name):
    return ''.join(c for c in Path(name).name if c.isalnum() or c in ' ._+-()[]').strip() or 'upload.bin'

def save_uploads(form):
    store = form.getfirst('store', '')
    target = UPLOAD_MAP.get(store)
    if not target: raise ValueError('Toko teu valid')
    target.mkdir(parents=True, exist_ok=True)
    saved = []
    items = form['files'] if 'files' in form else []
    if not isinstance(items, list): items = [items]
    for item in items:
        if not getattr(item, 'filename', None): continue
        name = safe_name(item.filename)
        dest = target / name
        with open(dest, 'wb') as f: shutil.copyfileobj(item.file, f)
        saved.append(str(dest.relative_to(BASE)))
    if not saved: raise ValueError('File kosong')
    rebuild()
    log = append_upload_log(store, saved)
    return {'saved': saved, 'log': log[:200]}

def rebuild():
    subprocess.check_call([sys.executable, 'build_bookkeeping.py'], cwd=BASE)

class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw): super().__init__(*a, directory=str(WEB), **kw)
    def send_json(self, code, obj):
        b=json.dumps(obj).encode(); self.send_response(code); self.send_header('Content-Type','application/json'); self.send_header('Content-Length',str(len(b))); self.end_headers(); self.wfile.write(b)
    def do_GET(self):
        if self.path.startswith('/api/upload-log'):
            return self.send_json(200, {'ok': True, 'rows': load_upload_log()[:200]})
        return super().do_GET()
    def do_POST(self):
        try:
            if self.path == '/api/hpp':
                body = self.rfile.read(int(self.headers.get('Content-Length','0')))
                rows = json.loads(body or b'{}').get('rows', [])
                result = upsert_hpp_rows(HPP, rows); rebuild()
                return self.send_json(200, {'ok': True, **result})
            if self.path == '/api/upload':
                form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={'REQUEST_METHOD':'POST','CONTENT_TYPE':self.headers.get('Content-Type')})
                return self.send_json(200, {'ok': True, **save_uploads(form)})
            return self.send_error(404)
        except Exception as e:
            self.send_json(500, {'ok': False, 'error': str(e)})

if __name__ == '__main__':
    ThreadingHTTPServer(('127.0.0.1', 8787), Handler).serve_forever()
