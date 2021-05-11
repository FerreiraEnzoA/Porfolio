# This automation project was created and developed by Enzo Ferreira
# This project use both pandas and openpyxl libraries
import pandas as pd

# Read .xlsx files
tab1 = pd.read_excel('/home/user/Área de trabalho/GitHub/Proj2/conf_cs.xlsx', usecols='A:C')
tab2 = pd.read_excel('/home/user/Área de trabalho/GitHub/Proj2/conf_customer.xlsx', usecols='A:C')

# Print DataFrames
print('\n============= DATAFRAMES =============')
print(tab1, '\n')
print(tab2, '\n')

conf = pd.DataFrame({
    'CPC': [],
    'QUANT': [],
    'ARRIVE': []
})

warn_data = pd.DataFrame({
    'CPC': [],
    'QUANT': [],
    'ARRIVE': []
})

warn_cs = pd.DataFrame({
    'CPC': []
})

warn_customer = pd.DataFrame({
    'CPC': []
})

# CPCs lists
print('\n\n\n============= LISTS CPC =============')
cpcs1 = tab1['CPC'].to_numpy()
cpcs2 = tab2['CPC'].to_numpy()
print('CPCs1: ', cpcs1)
print('CPCs2: ', cpcs2)

# CPC in conf_cs, but not in conf_customer (missing)
print('\n\n\n============= CONFIRMATION CS, BUT NOT CUSTOMER =============')
for i in cpcs1:
    if i not in cpcs2:
        print(i, ' is confirmed for cs, but not for customer (missing item)')
        warn_cs = warn_cs.append({'CPC': i}, ignore_index=True)

# CPC in conf_customer, but not in conf_cs (new items)
print('\n\n\n============= CONFIRMATION CUSTOMER, BUT NOT CS =============')
for i in cpcs2:
    if i not in cpcs1:
        print(i, ' is confirmed for customer, but not for cs (new item)')
        warn_customer = warn_customer.append({'CPC': i}, ignore_index=True)

# CPC in both conf
print('\n\n\n============= CPCs CONFIRMED IN BOTH =============')
for index1, row1 in tab1.iterrows():
    for index2, row2 in tab2.iterrows():
        if row1['CPC'] == row2['CPC']:
            # All matched
            if row1['QUANT'] == row2['QUANT'] and row1['ARRIVE'] == row2['ARRIVE']:
                print('Confirmed: ', row1['CPC'], row1['QUANT'], row1['ARRIVE'])
                conf = conf.append(
                    {
                        'CPC': row1['CPC'], 'QUANT': row1['QUANT'], 'ARRIVE': row1['ARRIVE']
                    },
                    ignore_index=True
                )
            # Mismatch in quantity
            elif row1['QUANT'] != row2['QUANT'] and row1['ARRIVE'] == row2['ARRIVE']:
                print(row1['CPC'], ' mismatch in quantity!')
                warn_data = warn_data.append(
                    {
                        'CPC': row1['CPC'], 'QUANT': row1['QUANT'], 'ARRIVE': row1['ARRIVE']
                    },
                    ignore_index=True)
            # Mismatch in arrive
            elif row1['QUANT'] == row2['QUANT'] and row1['ARRIVE'] != row2['ARRIVE']:
                print(row1['CPC'], ' mismatch in arrival date!')
                warn_data = warn_data.append(
                    {
                        'CPC': row1['CPC'], 'QUANT': row1['QUANT'], 'ARRIVE': row1['ARRIVE']
                    },
                    ignore_index=True)
            # Mismatch in both quantity and arrive
            else:
                print(row1['CPC'], ' mismatch in both quantity and arrival date!')
                warn_data = warn_data.append(
                    {
                        'CPC': row1['CPC'], 'QUANT': row1['QUANT'], 'ARRIVE': row1['ARRIVE']
                    },
                    ignore_index=True)

# Final conf .xlsx file
print('\n\n\n============= RECONFIRMATION =============')
print(conf)
conf.to_excel("/home/user/Área de trabalho/GitHub/Proj2/reconfirmation.xlsx", index=False)
print('Spreadsheet "reconfirmation.xlsx" successfully generated!')

# Warnings DataFrames
print('\n\n\n============= WARNINGS DURING RECONFIRMATION =============')
print(' *** Error 1: missing item in conf_customer\n', warn_cs, '\n')
print(' *** Error 2: missing item in conf_cs\n', warn_customer, '\n')
print(' *** Error 3: mismatch in quantity or/and arrival date\n', warn_data)
