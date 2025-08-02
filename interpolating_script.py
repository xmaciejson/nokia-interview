import pandas as pd

df = pd.read_excel('Data.xlsx')
df['Data'] = pd.to_datetime(df['Data'])
df.set_index('Data', inplace=True)
arpu_columns = ['ARPU_M1', 'ARPU_M2', 'ARPU_M3', 'ARPU_M4', 'ARPU_M5', 'ARPU_M6', 'ARPU_M7']
for col in arpu_columns:
    df[col] = df[col].interpolate(method='linear', limit_direction='both')

df.to_excel('arpu_data_filled.xlsx')