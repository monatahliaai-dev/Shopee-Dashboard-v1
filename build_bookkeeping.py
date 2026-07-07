import csv, glob, json, os, re
from pathlib import Path
import pandas as pd

BASE = Path('.')
OUT = 'Pembukuan_Shopee_Multi_Toko_2026.xlsx'

STORES = [
    {'name':'Hiban Store', 'root':'.', 'orders':['order_2026*.xlsx'], 'income':['income_jan_mayplus.xlsx','income_202606.xlsx'], 'ads':['ads_2026*.csv']},
    {'name':'Royal Abiya', 'root':'royal_abiya', 'orders':['**/Order*.xlsx'], 'income':['**/Income*.xlsx'], 'ads':['**/*.csv']},
    {'name':'Hiban Signature', 'root':'hiban_signature', 'orders':['**/Order*.xlsx'], 'income':['**/Income*.xlsx'], 'ads':['**/*.csv']},
]

def money(x):
    if pd.isna(x): return 0.0
    if isinstance(x,(int,float)): return float(x)
    s=str(x).strip().replace('Rp','').replace(' ','')
    if s in ('','-','nan','None'): return 0.0
    if re.fullmatch(r'-?\d{1,3}(\.\d{3})+(,\d+)?', s): s=s.replace('.','').replace(',','.')
    else: s=s.replace(',','')
    try: return float(s)
    except Exception: return 0.0

def month_from_name(fn):
    s=str(fn)
    m=re.search(r'2026(\d{2})', s)
    if not m: m=re.search(r'\d{2}_(\d{2})_2026', s)
    return f"2026-{m.group(1)}" if m else 'UNKNOWN'

def files(root, patterns):
    out=[]
    for pat in patterns:
        out += glob.glob(str(Path(root)/pat), recursive=True)
    return sorted(set(out))

def read_hpp(base=BASE):
    df=pd.read_excel(Path(base)/'hpp.xlsx', sheet_name='Utama', header=0)
    out={}
    for _,r in df.iterrows():
        sku=str(r.get('SKU Internal','')).strip()
        if sku and sku!='nan': out[sku]=money(r.get('HPP'))
    return out, df

def read_orders(store, hpp_map):
    rows=[]
    for fn in files(store['root'], store['orders']):
        mo=month_from_name(fn)
        df=pd.read_excel(fn, sheet_name='orders', header=0, dtype=str)
        for _,r in df.iterrows():
            sku=str(r.get('Nomor Referensi SKU') or r.get('SKU Induk') or '').strip()
            if not sku or sku=='nan': sku='(SKU kosong)'
            qty=money(r.get('Jumlah')); ret=money(r.get('Returned quantity'))
            hpp=hpp_map.get(sku, 0)
            rows.append({
                'Toko': store['name'], 'Bulan': mo, 'No Pesanan': r.get('No. Pesanan'),
                'Status': str(r.get('Status Pesanan','')).strip(), 'SKU': sku,
                'Nama Produk': r.get('Nama Produk'), 'Varian': r.get('Nama Variasi'),
                'Qty': qty, 'Qty Retur': ret,
                'Dibayar Pembeli': money(r.get('Dibayar Pembeli') if 'Dibayar Pembeli' in r else r.get('Total Pembayaran')),
                'Total Pembayaran': money(r.get('Total Pembayaran')),
                'HPP Satuan': hpp, 'HPP Total': hpp * max(qty-ret, 0), 'SKU Match': sku in hpp_map,
            })
    return pd.DataFrame(rows)

def read_income(store):
    dfs=[]
    for fn in files(store['root'], store['income']):
        try:
            xl=pd.ExcelFile(fn)
            sheet_names=[s for s in xl.sheet_names if s == 'Income' or s.startswith('Income -')]
        except Exception:
            continue
        for sheet in sheet_names:
            df=pd.read_excel(fn, sheet_name=sheet, header=5)
            if 'No. Pesanan' not in df: continue
            df=df[df['No. Pesanan'].notna()].copy()
            df['Toko']=store['name']
            df['Bulan']=pd.to_datetime(df.get('Tanggal Dana Dilepaskan'), errors='coerce').dt.strftime('%Y-%m')
            for c in ['Total Penghasilan','Harga Asli Produk','Total Diskon Produk','Biaya Administrasi','Biaya Layanan','Biaya Proses Pesanan','Biaya Komisi AMS']:
                if c not in df.columns: df[c]=0
            dfs.append(df)
    df=pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    if len(df): df=df.drop_duplicates(subset=['Toko','No. Pesanan','Tanggal Dana Dilepaskan','Total Penghasilan'])
    return df

def read_ads(store):
    rows=[]
    for fn in files(store['root'], store['ads']):
        if not fn.lower().endswith('.csv'): continue
        mo=month_from_name(fn)
        with open(fn, encoding='utf-8-sig', newline='') as f: data=list(csv.reader(f))
        if len(data)<8: continue
        header=data[7]; idx={h:i for i,h in enumerate(header)}
        for r in data[8:]:
            if len(r)<len(header): continue
            rows.append({
                'Toko': store['name'], 'Bulan': mo, 'Nama Iklan': r[idx.get('Nama Iklan',1)],
                'Biaya Iklan': money(r[idx.get('Biaya')]), 'Omzet Iklan': money(r[idx.get('Omzet Penjualan')]),
                'Produk Terjual Ads': money(r[idx.get('Produk Terjual')]), 'Klik': money(r[idx.get('Jumlah Klik')]),
            })
    return pd.DataFrame(rows)

def store_dataset(store, hpp_map):
    orders=read_orders(store, hpp_map)
    income=read_income(store)
    ads=read_ads(store)
    orders_selesai=orders[orders['Status'].str.lower().eq('selesai')].copy() if len(orders) else pd.DataFrame()
    inc_sum=income.groupby(['Toko','Bulan'], dropna=False).agg(
        Pesanan_Cair=('No. Pesanan','nunique'), Dana_Masuk=('Total Penghasilan','sum'), Harga_Asli=('Harga Asli Produk','sum'),
        Diskon_Produk=('Total Diskon Produk','sum'), Biaya_Admin=('Biaya Administrasi','sum'), Biaya_Layanan=('Biaya Layanan','sum'),
        Biaya_Proses=('Biaya Proses Pesanan','sum'), Biaya_Komisi_AMS=('Biaya Komisi AMS','sum'),
    ).reset_index() if len(income) else pd.DataFrame(columns=['Toko','Bulan'])
    ord_sum=orders_selesai.groupby(['Toko','Bulan']).agg(
        Order_Selesai=('No Pesanan','nunique'), Item_Terjual=('Qty','sum'), Omzet_Order=('Dibayar Pembeli','sum'),
        HPP=('HPP Total','sum'), SKU_Tidak_Match=('SKU Match', lambda s: int((~s).sum())),
    ).reset_index() if len(orders_selesai) else pd.DataFrame(columns=['Toko','Bulan'])
    ads_sum=ads.groupby(['Toko','Bulan']).agg(
        Biaya_Iklan=('Biaya Iklan','sum'), Omzet_Ads=('Omzet Iklan','sum'), Klik_Ads=('Klik','sum'), Produk_Terjual_Ads=('Produk Terjual Ads','sum'),
    ).reset_index() if len(ads) else pd.DataFrame(columns=['Toko','Bulan'])
    months=sorted(set(ord_sum.get('Bulan',[]))|set(inc_sum.get('Bulan',[]))|set(ads_sum.get('Bulan',[])))
    summary=pd.DataFrame({'Toko':store['name'], 'Bulan':months})
    for df in [ord_sum,inc_sum,ads_sum]: summary=summary.merge(df,on=['Toko','Bulan'],how='left')
    summary=summary.fillna(0)
    summary['Laba Kotor Setelah HPP']=summary['Dana_Masuk']-summary['HPP']
    summary['Laba Bersih Estimasi']=summary['Dana_Masuk']-summary['HPP']-summary['Biaya_Iklan']
    summary['Margin Bersih %']=summary.apply(lambda r: (r['Laba Bersih Estimasi']/r['Dana_Masuk']*100) if r['Dana_Masuk'] else 0, axis=1)
    summary['ROAS Ads']=summary.apply(lambda r: (r['Omzet_Ads']/r['Biaya_Iklan']) if r['Biaya_Iklan'] else 0, axis=1)
    unmatched=orders_selesai[~orders_selesai['SKU Match']].groupby(['Toko','SKU','Nama Produk','Varian'], dropna=False).agg(Qty=('Qty','sum'), Omzet=('Dibayar Pembeli','sum')).reset_index().sort_values('Qty', ascending=False) if len(orders_selesai) else pd.DataFrame()
    return {'summary':summary, 'orders':orders_selesai, 'income':income, 'ads':ads, 'unmatched':unmatched}

def build_dataset(base='.'):
    os.chdir(base)
    hpp_map, hpp_sheet = read_hpp('.')
    parts=[store_dataset(s, hpp_map) for s in STORES if Path(s['root']).exists()]
    return {k: pd.concat([p[k] for p in parts], ignore_index=True) if parts else pd.DataFrame() for k in ['summary','orders','income','ads','unmatched']} | {'hpp_sheet':hpp_sheet}

def clean(df): return json.loads(df.to_json(orient='records', force_ascii=False, date_format='iso'))

def write_outputs(data, out=OUT):
    orders=data['orders']; summary=data['summary']; unmatched=data['unmatched']
    prod=(orders.groupby(['Toko','SKU','Nama Produk'],dropna=False).agg(Qty=('Qty','sum'),Omzet=('Dibayar Pembeli','sum'),HPP=('HPP Total','sum')).reset_index())
    prod['Laba Kotor']=prod['Omzet']-prod['HPP']
    prod=prod.sort_values('Omzet',ascending=False).head(200)
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        summary.to_excel(w, sheet_name='Ringkasan Bulanan', index=False)
        orders.to_excel(w, sheet_name='Detail Order + HPP', index=False)
        data['income'].to_excel(w, sheet_name='Income Dana Dilepas', index=False)
        data['ads'].to_excel(w, sheet_name='Iklan Shopee', index=False)
        unmatched.to_excel(w, sheet_name='SKU Belum Ada HPP', index=False)
        data['hpp_sheet'].to_excel(w, sheet_name='Master HPP Utama', index=False)
    Path('web').mkdir(exist_ok=True)
    Path('web/data.json').write_text(json.dumps({'summary':clean(summary),'unmatched':clean(unmatched.head(500)),'products':clean(prod),'orders':clean(orders.tail(1000))}, ensure_ascii=False), encoding='utf-8')
    return out

if __name__ == '__main__':
    data=build_dataset('.')
    out=write_outputs(data)
    print(out)
    print(data['summary'].to_string(index=False))
    print('\nSKU unmatched rows:', len(data['unmatched']), 'qty:', data['unmatched']['Qty'].sum() if len(data['unmatched']) else 0)
