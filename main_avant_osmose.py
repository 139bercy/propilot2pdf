import os
import transpose_comments
import build_reports
import docx2pdf
import datetime
import zipfile
#Partie chargement propilot encore en ipynb
from nbconvert.preprocessors import ExecutePreprocessor
import nbformat


def main():
    # Création de pp_dep
    print("Création de pp_dep")
    notebook_filename = 'chargement_propilot.ipynb'
    lancement_auto_notebook(notebook_filename)
    print("Génération des fiches")
    build_reports.main_build_reports()
    print("Récupération des commentaires")
    transpose_comments.main_transpose_comments()
    print("Conversion en pdf")
    docx2pdf.main_docx2pdf_avant_osmose()
    print("Création des archives zip")
    # Obtention du mois de génération des fiches
    today = datetime.datetime.today()
    months = ('Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 
                'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre')
    today_str = f"{months[today.month-1]}_{today.year}"
    name_zip = 'Fiche_Parlementaire_Avant_Commentaires_{}.zip'.format(today_str)
    f=zipfile.ZipFile(name_zip,'w',zipfile.ZIP_DEFLATED)
    f.write("Fiche_Avant_Osmose_pdf")
    f.write("Fiche_Avant_Osmose")
    f.close()  
    

def lancement_auto_notebook(notebook_filename):
    with open(notebook_filename) as f:
        nb = nbformat.read(f, as_version=4) # Ouverture du ipynb
    ep = ExecutePreprocessor(timeout=1000, kernel_name='python3') # Configuration pour l'execution
    ep.preprocess(nb, {'metadata': {'path': os.getcwd()}}) # Execution du notebook présent dans 'path


if __name__ == "__main__":
    main()





