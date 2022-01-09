#usr!/bin/env python3

import xlrd
import matplotlib.pyplot as plt

civilité=[]
age=[]
licence=[]
cp=[]
mini=0
poussin=0
junior=0
adulte=0
document = xlrd.open_workbook("exportADOC_2021-2022.xls")
sheet=document.sheet_by_name("Détaillé")
nbligne=sheet.nrows


civilité=sheet.col_values(3)

m=civilité.count("Monsieur")
ma=civilité.count("Madame")
print('nombre femmes : ',ma)
print('nombre hommes : ',m)
print('')

labels = 'Hommes', 'Femmes'
sizes = [m, ma]
colors = ['lightskyblue', 'lightcoral']
plt.pie(sizes, labels=labels, colors=colors, 
        autopct='%1.1f%%', shadow=True, startangle=90)
plt.axis('equal')
plt.savefig('graphsexe.png')
plt.show()




age=sheet.col_values(5)
age.pop(0)
age.pop(0)
x=0
a=len(age)
while a>=0:
	age.count(x)
	if 0<x<7:
		mini=mini+1
	elif 7<=x<11:
		poussin=poussin+1
	elif 11<=x<=17:
		junior=junior+1
	elif 18<=x<=150:
		adulte=adulte+1
	else:
		pass
	n=0
	x=x+1
	a=a-1
print('nombre mini : ',mini)
print('nombre poussin : ',poussin)
print('nombre juniors : ',junior)
print('nombre adultes : ',adulte)
print('')

labels = 'Mini', 'Poussin', 'Junior', 'Adulte'
sizes = [mini, poussin, junior, adulte]
colors = ['yellowgreen', 'gold', 'lightskyblue', 'lightcoral']
plt.pie(sizes, labels=labels, colors=colors, 
        autopct='%1.1f%%', shadow=True, startangle=90)
plt.axis('equal')
plt.savefig('graphage.png')
plt.show()




licence=sheet.col_values(50)

c=licence.count('Compétition')
l=licence.count("Loisir")
nonr=licence.count('N')
print('nombre de licence compétition : ',c)
print('nombre de licence loisir : ',l)
print('non renseigé : ',nonr)
print('')

labels = 'Compétition', 'Loisir', 'Non renseigné'
sizes = [c, l, nonr]
colors = ['yellowgreen', 'gold', 'lightskyblue']
plt.pie(sizes, labels=labels, colors=colors, 
        autopct='%1.1f%%', shadow=True, startangle=90)
plt.axis('equal')
plt.savefig('graphlic.png')
plt.show()




cp=sheet.col_values(11)

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

labels = 'Vouneuille sous Biard', 'Montreuil Bonin', 'Poitiers', 'Beruges', 'Migne Auxances'
sizes = [v, mo, p, b, mi]
colors = ['yellowgreen', 'gold', 'lightskyblue', 'lightcoral', 'red']
plt.pie(sizes, labels=labels, colors=colors, 
        autopct='%1.1f%%', shadow=True, startangle=90)
plt.axis('equal')
plt.savefig('graphville.png')
plt.show()