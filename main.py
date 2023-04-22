import pandas as pd

table = pd.read_excel('v1 Example Warehouse Inventory.xlsx')

table['Total Value'] = table['Quantity'] * table['Price']

table = table.sort_values(by='Total Value', ascending=False)
table['Accumulated Value'] = table['Total Value'].cumsum()
table['Percentage'] = table['Accumulated Value'] / table['Total Value'].sum()

A = int(input("A = percentage less than: "))
C = int(input("C = percentage greater than: "))

table.loc[table['Percentage'] <= A/100, 'Class'] = 'A'
table.loc[(table['Percentage'] > A/100) & (table['Percentage'] <= C/100), 'Class'] = 'B'
table.loc[table['Percentage'] > C/100, 'Class'] = 'C'

proportion_sku = table.groupby('Class')['Material'].nunique() / table['Material'].nunique()
proportion_value = table.groupby('Class')['Total Value'].sum() / table['Total Value'].sum()

with pd.ExcelWriter('ABC_Table_Proportions.xlsx') as writer:
    table.to_excel(writer, sheet_name='ABC Table', index=False)
    proportion_sku.to_frame('SKU Proportion').to_excel(writer, sheet_name='Proportions', index=True)
    proportion_value.to_frame('Value Proportion').to_excel(writer, sheet_name='Proportions', index=True, startcol=2)

print('The ABC classification table and proportions have been written to the ABC_Table_Proportions.xlsx file.')
