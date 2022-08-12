import matplotlib.pyplot as plt
import numpy as np
import openpyxl
print('Автор: Василь Онуфрійчук') 
print('Виділення неспівпадаючих елементів списку із файлу ексель.')
print('')
a=list()
b=list()
c=list()
table_name='you_table.XLSX'
def build_list(a,b,h):
    k=2
    for i in range(h):
        x=sheet_wb[b+str(k)].value
        a.append(x)
        k+=1
def write_list(a,b):
    ws = wb.active
    k=2
    for i in a:
        ws[b+str(k)]=i
        k+=1
    wb.save(table_name)


wb = openpyxl.load_workbook(table_name)
sheet_wb=wb.active
b_wb=sheet_wb['B2'].value
d_wb=sheet_wb['D2'].value
build_list(a,'A',b_wb)
build_list(b,'C',d_wb)

for i in b:
    n=0
    for j in a:
        if i==j:
            n+=1
    if n==0:
        c.append(i)

write_list(c,'E')
