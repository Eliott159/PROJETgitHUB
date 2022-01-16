import os

"""
.. module:: os
   :platform: Tous les système d'exploitation
   :synopsis: Fournit une façon portable d'utiliser les fonctionnalités dépendantes du système d'exploitation.

.. moduleauthor:: Mazzolini Tim <tim.mazzolini@univ-poitiers.fr>

"""

import webbrowser

"""
.. module:: webbrowser
   :platform: Tous les système d'exploitation
   :synopsis: Fournit une interface de haut niveau pour permettre l'affichage de documents Web aux utilisateurs.

.. moduleauthor:: Mazzolini Tim <tim.mazzolini@univ-poitiers.fr>

"""


#Ouverture d'une page web situé dans l'ordinateur
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
"""

