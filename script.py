import pandas as pd
import openpyxl
import os
from copy import copy

print('Program running.')

# Get file
for f in os.listdir():
    if '.xlsx' in f:
        file_name = f

# Creates output folder
try:
    os.makedirs('output')
except:
    pass

# Read file
xl_file = pd.ExcelFile(file_name)
sheet_names = [sheet_name for sheet_name in xl_file.sheet_names]
dfs = [xl_file.parse(sheet_name) for sheet_name in xl_file.sheet_names]

sheet0 = dfs[0]
sheet1 = dfs[1]
sheet2 = dfs[2]
sheet3 = dfs[3]
sheet4 = dfs[4]

# Get list of unique companies
delimiter1 = b" " # Normal space
delimiter2 = b"\x81@" # Japanese space
companies = []
current_company = None
for sheet in [sheet0, sheet1, sheet2]:
    for i, company in enumerate(sheet[sheet.columns[2]][1:]):
        company_encode = company.encode('cp932')
        company_encode_splitted_1 = company_encode.split(delimiter1)
        company_encode_splitted_2 = company_encode.split(delimiter2)
        if company_encode_splitted_1[0] == company_encode_splitted_2[0]:
            company_final = company_encode_splitted_1[0]
        else:
            if len(company_encode_splitted_1) > 1:
                company_final = company_encode_splitted_1[0]
            elif len(company_encode_splitted_2) > 1:
                company_final = company_encode_splitted_2[0]
        companies.append(company_final)
companies = list(set(companies))

# Find index of seperation
data = {}
for company_ in companies:
    data[company_] = {}
    for sheet_n, sheet in enumerate([sheet0, sheet1, sheet2]):
        indexes = []
        for i, company in enumerate(sheet[sheet.columns[2]][1:]):
            company_encode = company.encode('cp932')
            company_encode_splitted_1 = company_encode.split(delimiter1)
            company_encode_splitted_2 = company_encode.split(delimiter2)
            if company_encode_splitted_1[0] == company_encode_splitted_2[0]:
                company_final = company_encode_splitted_1[0]
            else:
                if len(company_encode_splitted_1) > 1:
                    company_final = company_encode_splitted_1[0]
                elif len(company_encode_splitted_2) > 1:
                    company_final = company_encode_splitted_2[0]

            if company_final == company_:
                indexes.append(i+1)
        try:
            data[company_][sheet_n] = [min(indexes), max(indexes)]
        except:
            data[company_][sheet_n] = None

# Create new excels
def copy_cell(source_cell, dest_coord, tgt):
    tgt[dest_coord].value = source_cell.value
    if source_cell.has_style:
        tgt[dest_coord]._style = copy(source_cell._style)
    return tgt[dest_coord]

for company, values in data.items():
    wb = openpyxl.load_workbook(file_name)

    # Delete not needed rows
    for i in range(3):
        indexes = data[company][i]
        if indexes != None:
            total_row = wb[sheet_names[i]].max_row
            start, end = indexes
            if start != 1:
                wb[sheet_names[i]].delete_rows(3, start-1)
                n_deletion = start-1-3+1
                wb[sheet_names[i]].delete_rows(end+1-n_deletion, total_row)
            else:
                wb[sheet_names[i]].delete_rows(end+3, total_row)
        else:
            wb[sheet_names[i]].delete_rows(3, total_row)

    wb.save(f'output/{company.decode("cp932")}.xlsx')

print('Program complete.')
