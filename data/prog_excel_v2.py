#usr!/bin/env python3

import xlrd
import matplotlib.pyplot as plt
import webbrowser
import os

effectif=[]
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

"""
effectif=sheet.col_values(23)

n=effectif.split("/")
b=n[3]
e=b.count("2022")
print('Effectifs : ',e)
"""


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

"""
x1 = [ma]
x2 = [m]
bins = [0, 1, 2]
plt.hist([x1, x2], bins = bins, color = ['pink', 'blue'],
            edgecolor = 'red', hatch = '/', label = ['Madame', 'Monsieur'],
            histtype = 'bar') # bar est le defaut
plt.ylabel('Nombres')
plt.title('Histogramme par sexe')
plt.savefig('histsexe.png')
plt.legend()



x1 = [1, 2, 2, 3, 4, 4, 4, 4, 4, 5, 5]
x2 = [1, 1, 1, 2, 2, 3, 3, 3, 3, 4, 5, 5, 5]
bins = [x + 0.5 for x in range(0, 6)]
pyplot.hist([x1, x2], bins = bins, color = ['yellow', 'green'],
            edgecolor = 'red', hatch = '/', label = ['x1', 'x2'],
            histtype = 'bar') # bar est le defaut
pyplot.ylabel('valeurs')
pyplot.xlabel('nombres')
pyplot.title('2 series')
pyplot.legend()
"""


age=sheet.col_values(5)
age.pop(0)
age.pop(0)
"""y=[]"""
x=0
a=len(age)
for i in range (a):
	x=age[i]
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

data = [age]
plt.hist(data, bins = [0,5,10,15,20,25,30,35,40,45,50,55,60,65,70])
plt.title('Histogramme du nombre de personnes par tranche d\'âge de 5 ans', fontsize=10)
plt.xlabel('Age')
plt.ylabel('Nombre')
plt.savefig("hist_age1.png")
plt.show()
data = [age]
plt.hist(data, bins = [0,5,10,15,20,25,30,35,40,45,50,55,60,65,70], cumulative = -1)
plt.title('Histogramme cumulatif par tranche d\'âge de 5 ans', fontsize=10)
plt.savefig("hist_age2.png")
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

"""
x = [mini, poussin, junior, adulte]
plt.hist(x, range = (0, 5), bins = 5, color = 'yellow',
            edgecolor = 'red')
plt.xlabel(1)
plt.ylabel('mini')
plt.title('Histogramme simple')
plt.show()
"""




webbrowser.open(os.getcwd()+"/../html/html/tennis.html")





