import matplotlib.pyplot as plt #importer la fonction pyplot de la bibliothèque matplotlib et l'utiliser avec la chaine de caractère "plt"

"""
.. module:: pyplot
   :platform: Unix, windows, java, macOS
   :synopsis: Module graphique permet de créer des graphiques

.. moduleauthor:: Marceau Eliott <eliott.marceau@univ-poitiers.fr>

"""

#Programme pour creer un graphique "camembert"
labels = 'Hommes', 'Femmes' 

m=1
ma=2

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
Graphique camembert
"""


plt.axis('equal')
"""Permet d'avoir la même échelle sur l'axe des abscisses et l'axe des ordonnées afin de préserver la forme lors de l'affichage"""

plt.savefig('graphsexe.png')
"""sauveguarder une image du graphique sous le nom "graphsexe.png" """

plt.show()
"""afficher le graphique"""




#Programme pour creer un graphique "histogramme"
age = [50.0, 13.0, 25.0, 36.0, 9.0, 49.0, 60.0, 25.0, 16.0, 7.0, 6.0, 7.0, 11.0, 59.0, 22.0, 54.0, 49.0, 9.0, 45.0, 52.0, 44.0, 27.0, 25.0, 9.0, 13.0, 10.0, 12.0, 47.0, 41.0, 36.0, 46.0, 14.0, 14.0, 62.0, 11.0, 58.0, 46.0, 14.0, 59.0, 32.0, 33.0, 63.0, 21.0, 55.0, 77.0, 28.0, 16.0, 49.0, 67.0, 11.0, 16.0, 43.0, 33.0, 57.0, 17.0, 51.0, 10.0, 7.0, 11.0, 18.0, 13.0, 47.0, 29.0, 7.0, 41.0, 49.0, 9.0, 48.0, 17.0, 8.0, 73.0, 12.0, 57.0, 26.0, 45.0, 15.0, 41.0, 34.0, 11.0, 61.0, 7.0, 62.0, 12.0, 8.0, 33.0, 9.0, 14.0, 60.0, 13.0, 47.0, 67.0, 15.0, 48.0, 6.0, 69.0, 53.0, 64.0, 11.0, 60.0, 11.0, 12.0, 9.0, 66.0, 9.0, 57.0, 32.0, 21.0, 17.0, 42.0, 49.0, 24.0, 15.0, 33.0, 33.0, 11.0, 12.0, 10.0, 13.0, 16.0, 47.0, 49.0, 64.0, 9.0, 69.0, 51.0, 16.0, 48.0, 53.0, 39.0, 19.0, 49.0, 15.0, 52.0, 65.0, 13.0, 50.0, 14.0, 13.0, 63.0, 23.0, 49.0, 48.0, 10.0, 72.0, 12.0, 11.0, 22.0]

data = [age]

plt.hist(data, bins = [0,5,10,15,20,25,30,35,40,45,50,55,60,65,70])

"""Fonction qui creer un graphique histogramme à partir de données.

Parameters
----------
data : liste [age]
   taille des différentes colonnes et leurs positions (exemple: [1, 1, 1] = colonne d'odonnées 3 et d'abcisses 1)
bins : liste [0,5,10,15,20,25,30,35,40,45,50,55,60,65,70]
   différentes parties de l'histogramme (exemple: de 0 à 5, de 5 à 10, ..., il y une colonne)
pre: age de type liste
     [age] >= 0

Returns
-------
Graphique histogramme
"""


plt.title('Histogramme du nombre de personnes par tranche d\'âge de 5 ans', fontsize=10)

plt.xlabel('Age')

plt.ylabel('Nombre')

plt.savefig("hist_age1.png")

plt.show()



