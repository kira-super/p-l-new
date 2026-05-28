import pandas as pd
df = pd.read_parquet(r'oefof_pl/out/pl_results.parquet')
sk = df[df['Stock Name'].str.contains('SK SQUARE', case=False, na=False)].iloc[0]
print('=== SK Square ===')
for col in ['Starting Units','Ending Units','Units Sold','Market Value Start (EUR)','Market Value End (EUR)','Cost Basis (EUR)','Realised P&L (EUR)','Unrealised P&L (EUR)','Total P&L (EUR)','Total P&L (%)','Realised P&L (%)','Unrealised P&L (%)']:
    print(f'  {col:35s} {sk[col]}')
print()
print('=== HEADLINE ===')
print(f'Total P&L:  {df["Total P&L (EUR)"].sum():>15,.0f}')
print(f'Realised:   {df["Realised P&L (EUR)"].sum():>15,.0f}')
print(f'Unrealised: {df["Unrealised P&L (EUR)"].sum():>15,.0f}')
print(f'Income:     {df["Income (EUR)"].sum():>15,.0f}')
print(f'Cost Basis: {df["Cost Basis (EUR)"].sum():>15,.0f}')
print(f'MV End:     {df["Market Value End (EUR)"].sum():>15,.0f}')
print(f'MV Start:   {df["Market Value Start (EUR)"].sum():>15,.0f}')
print()
print('=== TOP 10 by Total P&L ===')
top = df.nlargest(10, 'Total P&L (EUR)')[['Stock Name','Total P&L (EUR)','Total P&L (%)','Realised P&L (EUR)','Unrealised P&L (EUR)']]
print(top.to_string(index=False))
