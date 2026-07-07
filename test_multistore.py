from build_bookkeeping import build_dataset


def test_build_dataset_contains_hiban_and_royal_abiya_store_rows():
    data = build_dataset('.')
    summary = data['summary']
    orders = data['orders']
    stores = set(summary['Toko'])
    assert {'Hiban Store', 'Royal Abiya', 'Hiban Signature'} <= stores
    assert {'Hiban Store', 'Royal Abiya', 'Hiban Signature'} <= set(orders['Toko'])
    assert len(summary[summary['Toko'] == 'Royal Abiya']) == 6
    assert len(summary[summary['Toko'] == 'Hiban Store']) == 6
