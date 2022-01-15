import xlrd 

document = xlrd.open_workbook("exportADOC_2021-2022.xls") #lit le fichier exportADOC_2021-2022.xls et renvoie un objet de type xlrd.Book

sheet=document.sheet_by_name("Détaillé") #renvoie la feuille de nom indiqué

nbligne=sheet.nrows #le nombre de lignes du fichier

civilité=sheet.col_values(0) #Renvoie une tranche des valeurs des cellules de la colonne donnée
