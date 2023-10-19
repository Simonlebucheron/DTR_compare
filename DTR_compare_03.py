# -*- coding: utf-8 -*-
"""
Created on Thu Oct 19 13:18:06 2023

@author: e_sdelai
"""

import pandas as pd
import os
import datetime

# Trouver le chemin absolu du répertoire contenant le script
dir_path = os.path.abspath(os.path.dirname(__file__))
# Concaténer avec le dossier "data" pour spécifier le répertoire d'entrée/sortie par défaut
data_path = os.path.join(dir_path, 'data')
# Récupérer les noms de fichiers Excel dans le dossier
excel_files = [os.path.join(data_path, f) for f in os.listdir(data_path) if f.endswith('.xlsx')]
# Récupérer le nom de chaque fichier et le stocker dans une liste
excel_files_names = [os.path.basename(file) for file in excel_files]

# Importer les deux fichiers Excel comme ensembles
set1 = set(pd.read_excel(excel_files[0]).stack().tolist())
set2 = set(pd.read_excel(excel_files[1]).stack().tolist())

# Créer un ensemble de tous les mots commençant par "DTR"
mots_dtr = set([mot for mot in set1.union(set2) if isinstance(mot, str) and mot.startswith('DTR')])
nb_DTR_total = len(mots_dtr)

# Trouver le sous-ensemble des mots commençant par "DTR" dans set2
DTR_presents = mots_dtr.intersection(set2)
nb_DTR_presents = len(DTR_presents)
 
# Trouver le sous-ensemble des mots commençant par "DTR" qui ne sont pas dans set2
DTR_absents = mots_dtr.difference(set2)
nb_DTR_absents = len(DTR_absents)

# Afficher la liste des DTR présents et absents
print("Nombre total de référence DTR trouvées : ", nb_DTR_total)
print("Nombre de DTR présents : ", nb_DTR_presents)
#print("Liste DTR présents : ", list(DTR_presents))
print("Nombre de DTR absents : ", nb_DTR_absents)
#print("Liste DTR absents : ", list(DTR_absents))
DTR_absents_format1 = [str(element) for element in DTR_absents]
DTR_absents_format2 = ', '.join(DTR_absents_format1)
print("Liste DTR absents formaté : ", DTR_absents_format2)

# Générer le fichier texte contenant les informations
now = datetime.datetime.now()
date_str = now.strftime("%d/%m/%Y %H:%M:%S")

# Créer un timestamp au format "YYYYMMDD_HHMMSS_"
timestamp = now.strftime('%Y%m%d_%H%M%S')

# Concaténer le timestamp avec le reste du nom de fichier
filename_output = os.path.join(data_path, timestamp + '_resultats_traitement.txt')

with open(filename_output, 'w') as f:
    f.write('Date de traitement : {}\n'.format(date_str))
    f.write('Auteur : S. DELAIGUE\n')
    f.write('Fichiers traités : {}\n'.format(", ".join(excel_files_names)))
    f.write('\n')
    f.write('Nombre total de référence DTR trouvées : {}\n'.format(nb_DTR_total))
    f.write('\n')
    f.write('Nombre de DTR présents : {}\n'.format(nb_DTR_presents))
    # On commente l'affichage de la liste des composants présents
    # f.write('Liste DTR présents : {}\n'.format(list(DTR_presents)))
    f.write('\n')
    f.write('Nombre de DTR absents : {}\n'.format(nb_DTR_absents))
    #f.write('Liste DTR absents : {}\n'.format(list(DTR_absents)))
    f.write('Liste DTR absents formaté : {}\n'.format(DTR_absents_format2))
