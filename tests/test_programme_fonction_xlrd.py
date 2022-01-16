import xlrd

"""
.. module:: xlrd
   :platform: Unix, windows, java, macOS
   :synopsis: Module excel permet d'extraire des données excel et de les traités

.. moduleauthor:: Marceau Eliott <eliott.marceau@univ-poitiers.fr>

"""

#lit le fichier exportADOC_2021-2022.xls et renvoie un objet de type xlrd.Book
document = xlrd.open_workbook("exportADOC_2021-2022.xls") 

#renvoie la feuille de nom indiqué
sheet=document.sheet_by_name("Détaillé") 

#renvoie nombre de lignes du fichier
nbligne=sheet.nrows 

#renvoie une tranche des valeurs des cellules de la colonne donnée
civilité=sheet.col_values(0)
