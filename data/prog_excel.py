#usr!/bin/env python3

import xlrd


civilité=[]
age=[]
licence=[]
cp=[]

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


licence=sheet.col_values(3)

c=licence.count("Compétition")
l=licence.count("Loisir")
print('nombre de licence compétition : ',c)
print('nombre de licence loisir : ',l)


cp=sheet.col_values(3)

v=cp.count('VOUNEUIL SOUS BIARD')
mo=cp.count('MONTREUIL BONNIN')
p=cp.count('POITIERS')
b=cp.count('BERUGES')
mi=cp.count('MIGNE AUXANCES')
print('nombre de licencié provenant de Vouneuil Sous Biard : ',v)
print('nombre de licencié provenant de Montreuil Bonin : ',mo)
print('nombre de licencié provenant de Poitiers : ',p)
print('nombre de licencié provenant de Beruges : ',b)
print('nombre de licencié provenant de Migné Auxances : ',mi)



"""civilité, age, pratique licence, niveau de jeu"""
