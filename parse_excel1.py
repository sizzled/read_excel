import xlrd
import pprint
book = xlrd.open_workbook('SOWC 2014 Stat Tables_Table 9.xlsx')
sheet = book.sheet_by_name('Table 9 ')
"""
print(sheet)
print(dir(sheet))

print(sheet.ncols)
"""
"""
for i in range(sheet.nrows):
    print(sheet.row_values(i))

for i in range(sheet.ncols):
    print(sheet.col_values(i))

count = 0
for i in range(sheet.nrows):
    if count < 20:
        if i <= 14:
            row = sheet.row_values(i)
            print(i,row)
        count += 1;

count = 0
data = {}
for i in range(sheet.nrows):
    if count < 20:
        if i >= 14:
            row = sheet.row_values(i)
            country = row[1]
            data[country] = {}
            count += 1
print(data)
"""
data = {}
for i in range(14,sheet.nrows):
    row = sheet.row_values(i)
    country = row[1]

    data[country] = {
            'child_labor':{
                'total':[row[4],row[5]],
                'male':[row[6],row[7]],
                'female':[row[8],row[9]],
                },
            'child_marriage':{
                'married_by_15':[row[10],row[11]],
                'married_by_18':[row[12],row[13]],
                }
            }
    if country == 'Zimbabwe':
        break


print('************Test 1 for Afghanistan********')
print(data['Afghanistan'])
print('************Test 2 for all Tables*********')
pprint.pprint(data)

