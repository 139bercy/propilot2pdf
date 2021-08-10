import os
import transpose_comments
import build_reports
import docx2pdf
import datetime
import zipfile

#Partie chargement propilot encore en ipynb
from nbconvert.preprocessors import ExecutePreprocessor
import nbformat

# Logger
import logging
import logging.config

#Définition du logger
logger = logging.getLogger("main")
logger.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
fh = logging.handlers.RotatingFileHandler("Parlementary_files.log", maxBytes=100000000, backupCount=5)
fh.setLevel(logging.DEBUG)
formatter = logging.Formatter("%(asctime)s - %(name)-20s - %(levelname)-8s - %(message)s")
ch.setFormatter(formatter)
fh.setFormatter(formatter)
logger.addHandler(ch)
logger.addHandler(fh)

def main():
    # Création de pp_dep
    logger.info("Création de pp_dep")
    notebook_filename = 'chargement_propilot.ipynb'
    auto_notebook_launch(notebook_filename)
    logger.info("Génération des fiches")
    build_reports.main_build_reports()
    logger.info("Récupération des commentaires")
    modified_docx_dir = "modified_reports"
    mkdir_ifnotexist(modified_docx_dir)
    if len(os.listdir(modified_docx_dir)) > 0:
        transpose_comments.main_transpose_comments()
        logger.info("Conversion en pdf")
        docx2pdf.main_docx2pdf_avant_osmose()
        logger.info("Création des archives zip")
        # Obtention du mois de génération des fiches
        today = datetime.datetime.today()
        months = ('Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 
                    'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre')
        today_str = f"{months[today.month-1]}_{today.year}"
        mkdir_ifnotexist("archive")
        path = os.path.join("archive", "{}".format(today_str))
        mkdir_ifnotexist(path)
        name_zip = os.path.join(path, 'reports_before_new_comment_{}.zip'.format(today_str))
        folder_pdf = "reports_before_new_comment_pdf"
        folder_docx = "reports_before_new_comment"
        create_zip_for_archive(name_zip, folder_pdf, folder_docx)
    else:
        logger.info("Le dossier modified_reports est vide. Arrêt du traitement")
        raise ValueError("Le dossier modified_reports est vide. Arrêt du traitement")
    

def auto_notebook_launch(notebook_filename: str):
    """
    Launch a notebook in a .py file
    """
    with open(notebook_filename) as f:
        nb = nbformat.read(f, as_version=4) # Ouverture du ipynb
    ep = ExecutePreprocessor(timeout=1000, kernel_name='python3') # Configuration pour l'execution
    ep.preprocess(nb, {'metadata': {'path': os.getcwd()}}) # Execution du notebook présent dans 'path


def mkdir_ifnotexist(path: str):
    """
    Creates a folder if it's doesn't exist
    """
    if not os.path.isdir(path):
        os.mkdir(path)

def create_zip_for_archive(name_zip: str, folder_pdf: str, folder_docx: str):
    """
    Creates a zip in archive/Month_Year with 2 folders: folder_pdf and forlder_docx
    """
    with zipfile.ZipFile(name_zip, "w", zipfile.ZIP_DEFLATED) as zfile:
            for root, _, files in os.walk(folder_pdf):
                for file in files:
                    zfile.write(os.path.join(root, file))
            for root, _, files in os.walk(folder_docx):
                for file in files:
                    zfile.write(os.path.join(root, file))

if __name__ == "__main__":
    main()
