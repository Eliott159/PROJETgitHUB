#usr!/bin/env python3

import xlrd 
"""Module xlrd.

Permet de:
-importer les données du document excel
-lire les données excel
"""
import matplotlib.pyplot as plt
"""Module pyplot de la bibilothèque matplotlib

Permet de:
-creer des graphiques
"""
import webbrowser
import os
"""Module webbrowser et os

os:
Fournit une façon portable d'utiliser les fonctionnalités dépendantes du système d'exploitation.
webbrowser:
Fournit une interface de haut niveau pour permettre l'affichage de documents Web aux utilisateurs.
"""

effectif=[]
civilité=[]
age=[]
licence=[]
cp=[]
mah=[]
mh=[]
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

"""
x1 = [ma]*1
x2 = [m]*1
bins = [0, 1, 2]
plt.hist([x1, x2], bins = bins, color = ['pink', 'blue'],
            edgecolor = 'red', hatch = '/', label = ['Madame', 'Monsieur'],
            histtype = 'bar') # bar est le defaut
plt.ylabel('Nombres')
plt.title('Histogramme par sexe')
plt.savefig('histsexe.png')
plt.legend()
"""


for loop in range (ma):
	mah.append(1)
	"""loop permettant de creer une liste mah utilisable par le module PYPLOT afin de creer un histogramme."""
for loop in range (m):
	mh.append(2)
	"""loop permettant de creer une liste mh utilisable par le module PYPLOT afin de creer un histogramme."""
plt.hist([mah, mh], bins = [1,2], color = ['pink', 'blue'],
            label = ['Madame', 'Monsieur'], histtype = 'bar') # Création du graphique: avec les valeurs des mah et mh, qui va de 1 à 2 en ordonnées, de coleur rose pour mah et bleu pour mh, avec une 									  légende (Madame=rose et Monsieur=bleu), sélection du type d'gistogramme classic (bar)
plt.ylabel('Nombre')
plt.xlabel('Valeurs à ignorer')
plt.title('Histogramme des effectifs par sexe')
plt.savefig('histsexe')
plt.legend()
plt.show()



age=sheet.col_values(5)
age.pop(0)
age.pop(0)
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



webbrowser.open(os.getcwd()+"/../html/html/tennis.html")




