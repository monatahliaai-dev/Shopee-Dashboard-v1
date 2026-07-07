import tempfile
from pathlib import Path
import pandas as pd
from server import upsert_hpp_rows

def test_upsert_hpp_rows_updates_existing_and_adds_missing_sku():
    p = Path(tempfile.mkdtemp())/'hpp.xlsx'
    pd.DataFrame([{'SKU Internal':'A','Nama Produk':'old','Varian':'v','HPP':100}]).to_excel(p, sheet_name='Utama', index=False)
    result = upsert_hpp_rows(p, [
        {'sku':'A','nama':'new','varian':'v2','hpp':150},
        {'sku':'B','nama':'prod b','varian':'x','hpp':200},
    ])
    df = pd.read_excel(p, sheet_name='Utama')
    assert result == {'updated': 1, 'added': 1}
    assert dict(zip(df['SKU Internal'], df['HPP'])) == {'A':150, 'B':200}
    assert df.loc[df['SKU Internal'].eq('B'), 'Nama Produk'].iloc[0] == 'prod b'
