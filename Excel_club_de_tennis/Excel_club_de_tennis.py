#usr!/bin/env python3

"""
Created on Wed Nov  23 010:58:20 2021
@authors: Mazzolini Tim, Marceau Eliott
"""


import xlrd 
"""
.. module:: xlrd
   :platform: Unix, windows, java, macOS
   :synopsis: Module excel permet d'extraire des données excel et de les traités

.. moduleauthor:: Marceau Eliott <eliott.marceau@univ-poitiers.fr>

"""

import matplotlib.pyplot as plt
"""
.. module:: pyplot
   :platform: Unix, windows, java, macOS
   :synopsis: Module graphique permet de créer des graphiques

.. moduleauthor:: Marceau Eliott <eliott.marceau@univ-poitiers.fr>

"""
import webbrowser
"""
.. module:: webbrowser
   :platform: Tous les système d'exploitation
   :synopsis: Fournit une interface de haut niveau pour permettre l'affichage de documents Web aux utilisateurs.

.. moduleauthor:: Mazzolini Tim <tim.mazzolini@univ-poitiers.fr>

"""
import os
"""
.. module:: os
   :platform: Tous les système d'exploitation
   :synopsis: Fournit une façon portable d'utiliser les fonctionnalités dépendantes du système d'exploitation.

.. moduleauthor:: Mazzolini Tim <tim.mazzolini@univ-poitiers.fr>

"""

effectif=[]
civilité=[]
age=[]
licence=[]
cp=[]
mah=[]
mh=[]
nch=[]
ch40=[]
ch305=[]
ch304=[]
ch303=[]
ch302=[]
ch301=[]
ch30=[]
ch155=[]
ch154=[]
ch153=[]
ch152=[]
ch151=[]
ch15=[]
ch56=[]
ch46=[]
ch36=[]
ch26=[]
ch16=[]
ch0=[]
mini=0
poussin=0
junior=0
adulte=0
document = xlrd.open_workbook("../data/exportADOC_2021-2022.xls")
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
"""Fonction qui creer un graphique camembert à partir de données.

Parameters
----------
sizes : liste [m ,ma]
   taille des différentes parties
colors : "lightskyblue" et "lightcoral"
   coleurs des différentes parties
labels : 'Hommes', 'Femmes' 
   noms des différentes parties
autopct : %1.1f%%
   calcul du pourcentage de chaque partie
shadow : True
   ajout d'une ombre au graphique
startangle : 90
   angle du graphique
pre: m et ma de type int
     m, ma >= 0

Returns
-------
Graphique camembert sexe
"""
plt.axis('equal')
plt.savefig('../data/graphsexe.png')
plt.show()


for loop in range (ma):
	mah.append(1)
	"""loop permettant de creer une liste mah utilisable par le module PYPLOT afin de creer un histogramme."""
for loop in range (m):
	mh.append(2)
	"""loop permettant de creer une liste mh utilisable par le module PYPLOT afin de creer un histogramme."""
plt.hist([mah, mh], bins = [1,2], color = ['pink', 'blue'],
            label = ['Madame', 'Monsieur'], histtype = 'bar') # Création du graphique: avec les valeurs des mah et mh, qui va de 1 à 2 en ordonnées, de coleur rose pour mah et bleu pour mh, avec une 									  légende (Madame=rose et Monsieur=bleu), sélection du type d'gistogramme classic (bar)
"""Fonction qui creer un graphique histogramme à partir de données.

Parameters
----------
data : liste [mah, mh]
   taille des différentes colonnes et leurs positions (exemple: [1, 1, 1] = colonne d'odonnées 3 et d'abcisses 1)
bins : liste [1,2]
   différentes parties de l'histogramme (exemple: de 1.1 à 1.5, de 1.5 à 1.9, ..., il y une colonne)
pre: mah et mh de type liste
     [mah, mh] >= 0

Returns
-------
Graphique histogramme

Exemple
-------
Voir test unitaire
"""
plt.ylabel('Nombre')
plt.xlabel('Valeurs à ignorer')
plt.title('Histogramme des effectifs par sexe')
plt.savefig('../data/histsexe')
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
"""Programme qui permet de creer les catégories
A faire !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
"""

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
plt.savefig('../data/graphage.png')
plt.show()








data = [age]
plt.hist(data, bins = [0,5,10,15,20,25,30,35,40,45,50,55,60,65,70])
plt.title('Histogramme du nombre de personnes par tranche d\'âge de 5 ans', fontsize=10)
plt.xlabel('Age')
plt.ylabel('Nombre')
plt.savefig("../data/hist_age1.png")
plt.show()
"""Fonction qui creer un graphique histogramme (du nombre de personnes par tranche d'âge) à partir de données.

Parameters
----------
data : liste [age]
   taille des différentes colonnes et leurs positions (exemple: [1, 1, 1] = colonne d'ordonnées 3 et d'abcisses 1)
bins : liste [0,5,10,15,20,25,30,35,40,45,50,55,60,65,70]
   différentes parties de l'histogramme (exemple: de 0 à 5, de 5 à 10, ..., il y une colonne)
pre: age de type liste
     [age] >= 0

Returns
-------
Graphique histogramme
Exemple
-------
Voir test unitaire
"""

data = [age]
plt.hist(data, bins = [0,5,10,15,20,25,30,35,40,45,50,55,60,65,70], cumulative = -1)
plt.title('Histogramme cumulatif par tranche d\'âge de 5 ans', fontsize=10)
plt.savefig("../data/hist_age2.png")
plt.show()
"""Fonction qui creer un graphique histogramme (cumulatif par tranche d'âge) à partir de données.

Parameters
----------
data : liste [age]
   taille des différentes colonnes et leurs positions (exemple: [1, 1, 1] = colonne d'ordonnées 3 et d'abcisses 1)
bins : liste [0,5,10,15,20,25,30,35,40,45,50,55,60,65,70]
   différentes parties de l'histogramme (exemple: de 0 à 5, de 5 à 10, ..., il y une colonne)
pre: age de type liste
     [age] >= 0

Returns
-------
Graphique histogramme
Exemple
-------
Voir test unitaire
"""






classement=sheet.col_values(27)

nc=classement.count("NC   ")
c40=classement.count("40   ")
c305=classement.count("30/5 ")
c304=classement.count("30/4 ")
c303=classement.count("30/3 ")
c302=classement.count("30/2 ")
c301=classement.count("30/1 ")
c30=classement.count("30   ")
c155=classement.count("15/5 ")
c154=classement.count("15/4 ")
c153=classement.count("15/3 ")
c152=classement.count("15/2 ")
c151=classement.count("15/1 ")
c15=classement.count("15 ")
c56=classement.count("5/6  ")
c46=classement.count("4/6  ")
c36=classement.count("3/6  ")
c26=classement.count("2/6  ")
c16=classement.count("1/6  ")
c0=classement.count("0  ")
print('nombre de personnes n\'ayant pas de grade : ',nc)
print('nombre de personnes ayant le grade de 40 : ',c40)
print('nombre de personnes ayant le grade de 30/5 : ',c305)
print('nombre de personnes ayant le grade de 30/4 : ',c304)
print('nombre de personnes ayant le grade de 30/3 : ',c303)
print('nombre de personnes ayant le grade de 30/2 : ',c302)
print('nombre de personnes ayant le grade de 30/1 : ',c301)
print('nombre de personnes ayant le grade de 30 : ',c30)
print('nombre de personnes ayant le grade de 15/5 : ',c155)
print('nombre de personnes ayant le grade de 15/4 : ',c154)
print('nombre de personnes ayant le grade de 15/3 : ',c153)
print('nombre de personnes ayant le grade de 15/2 : ',c152)
print('nombre de personnes ayant le grade de 15/1 : ',c151)
print('nombre de personnes ayant le grade de 15 : ',c15)
print('nombre de personnes ayant le grade de 5/6 : ',c56)
print('nombre de personnes ayant le grade de 4/6 : ',c46)
print('nombre de personnes ayant le grade de 3/6 : ',c36)
print('nombre de personnes ayant le grade de 2/6 : ',c26)
print('nombre de personnes ayant le grade de 1/6 : ',c16)
print('nombre de personnes ayant le grade de 0/6 : ',c0)
print('')

for loop in range (nc):
	nch.append(1)
for loop in range (c40):
	ch40.append(2)
for loop in range (c305):
	ch305.append(3)
for loop in range (c304):
	ch304.append(4)
for loop in range (c303):
	ch303.append(5)
for loop in range (c302):
	ch302.append(6)
for loop in range (c301):
	ch301.append(7)
for loop in range (c30):
	ch30.append(8)
for loop in range (c155):
	ch155.append(9)
for loop in range (c154):
	ch154.append(10)
for loop in range (c153):
	ch153.append(11)
for loop in range (c152):
	ch152.append(12)
for loop in range (c151):
	ch151.append(13)
for loop in range (c15):
	ch15.append(14)
for loop in range (c56):
	ch56.append(15)
for loop in range (c46):
	ch46.append(14)
for loop in range (c36):
	ch36.append(15)
for loop in range (c26):
	ch26.append(16)
for loop in range (c16):
	ch16.append(17)
for loop in range (c0):
	ch0.append(18)
data = [nch, ch40, ch305, ch304, ch303, ch302, ch301, ch30, ch155, ch154, ch153, ch152, ch151, ch15, ch56, ch46, ch36, ch26, ch16, ch0]
plt.hist(data, bins = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18])
plt.title('Histogramme du classement des memebres du club', fontsize=10)
plt.xlabel('NC-40-30/5-30/4-30/3-30/2-30/1-30-15/5-15/4-15/3-15/2-15/1-15-6/5-4/6-3/6-2/6-1/6-0')
plt.ylabel('Nombre')
plt.savefig("../data/classement.png")
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
plt.savefig('../data/graphlic.png')
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
plt.savefig('../data/graphville.png')
plt.show()



webbrowser.open(os.getcwd()+"/../html/html/tennis.html")
"""Fonction qui ouvre un site web.

Parameters
----------
os : getcwd()
   Permet d'utilisé le chemin vers la page web
webbrowser : open(os.getcwd()+"/../html/html/tennis.html")
   Permet l'ouverture du fichier sur le web

Returns
-------
Site HTML
Exemple
-------
Voir test unitaire
"""



