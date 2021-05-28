# This automation project was created and developed by Enzo Ferreira
# This project use both pandas and openpyxl libraries
import pandas as pd

# Read .xlsx files
tab1 = pd.read_excel('/home/user/Documentos/GitHub/Proj3/gen_ins.xlsx')
tab2 = pd.read_excel('/home/user/Documentos/GitHub/Proj3/conf.xlsx')
tab3 = pd.read_excel('/home/user/Documentos/GitHub/Proj3/inventory1_.xlsx')
tab4 = pd.read_excel('/home/user/Documentos/GitHub/Proj3/inventory2_.xlsx')

# Print DataFrames
print('\n============= DATAFRAMES =============')
print(tab1, '\n')
print(tab2, '\n')
print(tab3, '\n')
print(tab4, '\n')

conf = pd.DataFrame({
    'P/O #': [],
    'Item #': [],
    'Description': [],
    'Qty': [],
    'Lot #/Pallet #': [],
    'Shpmnt #': [],
    'Closed Date': [],
    'PLANT': [],
    'SHIP TO': [],
    'ORDER': [],
    'INVOICE': []
})

plants = []
shipto = []

# CPCs lists
print('\n\n\n============= CPCs LISTS =============')
cpc_t = tab1['p1'].tolist()
cpc_m = tab1['p2'].tolist()
print('CPCs 1:', cpc_t)
print('CPCs 2:', cpc_m)

a_city = tab3['Package id'].tolist()
bc_cities = tab4['Package id'].tolist()
pallet = tab2['Lot #/Pallet #']

# CPCs Wh
print('\n\n\n============= CPCs WAREHOUSES =============')
for i in tab2['Item #']:
    a = str(i)
    if int(a[0:4]) in cpc_t:
        print(i, '1')
        plants.append('1')
    elif int(a[0:4]) in cpc_m:
        print(i, '2')
        plants.append('2')
    else:
        print(f'CPC {i} not found in both inventories. Check manually!')
        plants.append('?')

# CPCs Ship to
print('\n\n\n============= CPCs DESTINATION =============')
for z in pallet:
    if z in a_city:
        print(z, 'A')
        shipto.append('A')
    elif z in bc_cities:
        x = tab4.loc[tab4['Package id'] == z]
        k = str(x['CustProdName'].item())
        if k == '22' or 'B' in k:
            print(z, 'B')
            shipto.append('B')
        else:
            print(z, 'C')
            shipto.append('C')
    else:
        print(f'Package {z} not found in both inventories. Check manually!')
        shipto.append('?')

# Invoice DataFrame
itera = ['P/O #', 'Item #', 'Description', 'Qty', 'Lot #/Pallet #', 'Shpmnt #', 'Closed Date']
for i in itera:
    conf[i] = tab2[i]

for b in tab2['Item #']:
    conf['ORDER'] = ''
    conf['INVOICE'] = ''

conf['PLANT'] = plants
conf['SHIP TO'] = shipto

print('\n\n\n', conf)
conf.to_excel("/home/user/Documentos/GitHub/Proj3/invoice.xlsx", index=False)
print('\nSpreadsheet "invoice.xlsx" successfully generated!')
