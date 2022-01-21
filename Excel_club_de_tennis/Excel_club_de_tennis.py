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
plt.legend()
plt.savefig('../data/histsexe')
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
plt.title('Camembert des effectifs par catégories', fontsize=10)
plt.legend()
plt.savefig('../data/graphage.png')
plt.show()
efficace=mini+poussin+junior+adulte
"""La variable efficace permet d'afficher le nombre d'effectif dans l'html"""






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
plt.hist(data, bins = range(15))
plt.title('Histogramme du classement des membres du club', fontsize=10)
plt.xlabel('NC-40-30/5-30/4-30/3-30/2-30/1-30-15/5-15/4-15/3-15/2-15/1-15-6/5-4/6-3/6-2/6-1/6-0')
plt.ylabel('Nombre')
plt.savefig("../data/classement.png")
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
plt.title('Camembert de la répartition par commune', fontsize=10)
plt.legend()
plt.savefig('../data/graphville.png')
plt.show()



webbrowser.open(os.getcwd()+"/../html/html/tennis2.html")
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

#######################################
#                                     #
#             Programme HTML          #
#                                     #
#######################################

def write_html_head(fichier, titre):

    fichier.write(
"""
<head>
<title></title>
<meta name="generator" content="Bluefish 2.2.12" >
<meta name="author" content="Eliott" >
<meta name="date" content="2022-01-17T22:58:55+0100" >
<meta name="copyright" content="">
<meta name="keywords" content="">
<meta name="description" content="">
<meta name="ROBOTS" content="NOINDEX, NOFOLLOW">
<meta http-equiv="content-type" content="text/html; charset=UTF-8">
<meta http-equiv="content-type" content="application/xhtml+xml; charset=UTF-8">
<meta http-equiv="content-style-type" content="text/css">
<meta http-equiv="expires" content="0">
<link href="../css/tennis.css" rel="stylesheet" type="text/css">
</head>""")
   
def write_html_body(fichier):
    fichier.write(
"""<body>""")

def write_html_header(fichier):
    fichier.write(

"""<header>
<div id="logo">
<img src="../Images/logo.png" width="125" height="82" alt="">
</div>

 <div id="titre">
	<h1>Tennis-Club Val de Boivre </h1>
</div>
</header>



<nav> 
<div class="menu">
		<section class="categorie">
			<a href="#categorie"><h3>Catégorie</h3></a>
		</section>
		<section class="categorie">
			<a href="#ville"><h3>Ville</h3></a>
		</section>
		<section class="categorie">
			<a href="#Sexe"><h3>Sexe</h3></a>
		</section>
		<section class="categorie">
			<a href="#classement"><h3>Classement</h3></a>
		</section>
		<section class="categorie">
			<a href="#hist"><h3>Age</h3></a>
		</section>

</nav>
<section>
<article id="accueil">
	<div id="img">
		<img src="../Images/tenniscour.jpg" width="100%" height="100%" alt="">
	</div>

	<div id="infosG">
		<div id="soust">
			<h2>Infos Générales sur le club</h2>	
		</div>
			<p>
    <b>BIENVENUE au TCVdB</p></b><hr style="width: 400px">
    Le club propose la pratique encadrée à partir de 5 ans, en loisirs et/ou compétition, dispensée par un entraîneur diplômé d’Etat et un entraîneur qualifié.
    Il dispose pour cela de nombreuses infrastructures, 20 terrains dont 10 couverts répartis à Biard centre, Vouneuil-sous-Biard centre, au Creps de Boivre ainsi qu’à Poitiers Saint-Nicolas.
    L’équipe dirigeante a étoffé les outils pédagogiques du club en investissant dans un lanceur de balles en Avril 2021.
    Venez nous rejoindre, que vous soyez débutant ou confirmé, jeune ou moins jeune, nous vous accueillerons les bras ouverts, raquette à la main ! </p>
	</div>
</article>

<article id="eff_ann">
<h1>Effectif de l'année 2022: """)
def write_html_eff(fichier, efficace):
    fichier.write("\t<td>"+str(efficace)+"</td>")
def write_html_header2(fichier):
    fichier.write(
"""</h1>
</article>

<article id="categorie">
	<h2>Par catégorie</h2>
	<img src="../../data/graphage.png" width="960" height="720" alt="">
</article>

<article id="ville">
	<h2>Par commune</h2>
	<img src="../../data/graphville.png" width="960" height="720" alt="">
</article>

<article id="Sexe">
	<h2>Par Sexe/Catégorie</h2>
	<img src="../../data/histsexe.png" width="960" height="720" alt="">
	<img src="../../data/graphsexe.png" width="960" height="720" alt="">
</article>

<article id="classement">
	<h2>Par Classement</h2>
	<img src="../../data/classement.png" width="960" height="720" alt="">
</article>

<article id="hist">
	<h2>Par âge</h2>
	<div id="hist1">
	<img src="../../data/hist_age2.png" width="768" height="576" alt="">
	</div>
	<div id="hist2">
	<img src="../../data/hist_age1.png" width="768" height="576" alt="">
	</div>
</article>




</section>""")

def write_html_footer(fichier):
    fichier.write(

"""<footer>
<div id="foot_text">
<br>
Contacter le club<br>
CREPS de POITIERS<br>
156 Route de Parthenay 86000 Poitiers<br>
</div>


</footer>


</body>""")


    
    
def write_html_end(fichier):
    fichier.write("</html>")
    
fichier = open("../html/html/tennis2.html", "w")
write_html_head(fichier,"mon fichier")
write_html_body(fichier)
write_html_header(fichier)
write_html_eff(fichier, efficace)
write_html_header2(fichier)
write_html_footer(fichier)



fichier.close()










