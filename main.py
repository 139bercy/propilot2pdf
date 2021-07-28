import nbformat
import os
import shutil
import datetime
from nbconvert.preprocessors import ExecutePreprocessor

mois_fiche = "Juin"

#Création du script pour les fiches parlementaires
def lancement_auto_notebook(notebook_filename):
    with open(notebook_filename) as f:
        nb = nbformat.read(f, as_version=4) # Ouverture du ipynb
    ep = ExecutePreprocessor(timeout=1000, kernel_name='python3') # Configuration pour l'execution
    ep.preprocess(nb, {'metadata': {'path': os.getcwd()}}) # Execution du notebook présent dans 'path
# Création de pp_dep
print("Création de pp_dep")
notebook_filename = 'chargement_propilot.ipynb'
lancement_auto_notebook(notebook_filename)

# Génération des fiches
print("Génération des fiches")
notebook_filename = 'build_reports.ipynb'
lancement_auto_notebook(notebook_filename)

# Transposition des commentaires
print("Transposition des commentaires et remplissages des fiches")
notebook_filename = 'transpose_comments.ipynb'
lancement_auto_notebook(notebook_filename)

## Warning: Les tables des matières sont à générer manuellement, et avant la conversion en pdf


# Export en pdf des fiches avec commentaires
print("Export pdf des fiches")
notebook_filename = 'docx2pdf.ipynb'
lancement_auto_notebook(notebook_filename)

# Conversion en zip
print("Conversion des fiches en zip")
date_today = datetime.date.today()
if date_today.day < 23:
    shutil.make_archive('Fiche_Parlementaire_{}_Avant_Osmose'.format(mois_fiche), 'zip', root_dir='Fiche_Avant_Osmose_pdf')
else:
    shutil.make_archive('Fiche_Parlementaire_{}_Apres_Osmose'.format(mois_fiche), 'zip', root_dir='reports_pdf')

