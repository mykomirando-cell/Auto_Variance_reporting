import pandas as pd

files = {
    'prev': r'sample data\sample_previous_template 08102026.xlsx',
    'curr': r'sample data\sample_current_template 08102026.xlsx',
    'iss':  r'sample data\sample_issuance_template 08102026.xlsx',
    'recv': r'sample data\sample_received_template 08102026.xlsx',
    'conv': r'sample data\uom_conversion_lookup_template 08142026.csv',
}
for k, p in files.items():
    if p.endswith('.csv'):
        df = pd.read_csv(p)
    else:
        df = pd.read_excel(p)
    print('===', k, '===')
    print('rows:', len(df), '| cols:', list(df.columns))
    print(df.head(2).to_string(index=False))
    print()
