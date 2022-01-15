import matplotlib.pyplot as plt #importer la fonction pyplot de la bibliothèque matplotlib et l'utiliser avec la chaine de caractère "plt"


#Programme pour creer un graphique "camembert"
labels = 'Hommes', 'Femmes' #nom des différentes partie du camembert

sizes = [m, ma] #configuration de la taille des parties du camembert --> dans l'exemple il y aura 2 partie de taille "m" et "ma" 

colors = ['lightskyblue', 'lightcoral'] #configuration des couleurs des différentes parties du camembert --> dans l"exemple "lightskyblue" et "lightcoral"

plt.pie(sizes, labels=labels, colors=colors, 
        autopct='%1.1f%%', shadow=True, startangle=90) #création du camembert à partir des tailles, noms et coueleurs déja configurées, ainsi que des extensions: permettant de calculer les pourcentage de chaque parties du camembert (autopct='%1.1f%%'); ajouter une ombre sous le camembert (shadow=True); ajouter l'angle dans lequel le graphique va être
        
plt.axis('equal') #Permet d'avoir la même échelle sur l'axe des abscisses et l'axe des ordonnées afin de préserver la forme lors de l'affichage

plt.savefig('graphsexe.png') #sauveguarder une image du graphique sous le nom "graphsexe.png"

plt.show() #afficher le graphique

