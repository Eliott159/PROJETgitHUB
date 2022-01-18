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



