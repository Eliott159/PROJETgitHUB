#usr!/bin/env python3

import xlrd


civilité=[]
age=[]



document = xlrd.open_workbook("exportADOC_2021-2022.xls")
sheet=document.sheet_by_name("Détaillé")



civilité=sheet.col_values(3)

m=civilité.count("Monsieur")
ma=civilité.count("Madame")
print('nombre femmes : ',ma)
print('nombre hommes : ',m)


age=sheet.col_values(3)

mini1=age.count('1')
mini2=age.count('2')
mini3=age.count('3')
mini4=age.count('4')
mini5=age.count('5')
mini6=age.count('6')
mini7=age.count('7')
mini8=age.count('8')
mini=mini1+mini2+mini3+mini4+mini5+mini6+mini7+mini8

print('nombre mini : ',mini)





"""
print(sheet.cell_value(2, 2))
"""


"""civilité, age, pratique licence, niveau de jeu"""
