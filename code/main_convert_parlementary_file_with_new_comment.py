import os
import sys
import zipfile
import datetime
# import filecmp
import shutil

# Logger
import logging
import logging.config

# Import des scripts dans le dossier code
sys.path.append(os.path.join(os.getcwd(), "code"))
import docx2pdf
from diff_pdf_visually import pdfdiff


# Définition du logger
logger = logging.getLogger("main")
logger.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
fh = logging.handlers.RotatingFileHandler("Parlementary_files_Conversion.log", maxBytes=100000000, backupCount=5)
fh.setLevel(logging.DEBUG)
formatter = logging.Formatter("%(asctime)s - %(name)-20s - %(levelname)-8s - %(message)s")
ch.setFormatter(formatter)
fh.setFormatter(formatter)
logger.addHandler(ch)
logger.addHandler(fh)

# Chemin des fiches pdf
path_to_folder1 = os.path.join("reports", "reports_before_new_comment_pdf")
path_to_folder2 = os.path.join("reports", "reports_pdf")


def main():
    logger.info("Conversion des fiches présentes dans modified_reports")
    docx2pdf.main_docx2pdf_apres_osmose()
    logger.info("Création des archives zip")
    # Obtention du mois de génération des fiches
    today = datetime.datetime.today()
    months = ('Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet',
              'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre')
    today_str = f"{months[today.month-1]}_{today.year}"

    mkdir_ifnotexist("archive")
    path = os.path.join("archive", "{}".format(today_str))
    mkdir_ifnotexist(path)

    name_zip = os.path.join(path, 'reports_with_new_comment_{}.zip'.format(today_str))
    folder_pdf = os.path.join("reports", "reports_pdf")
    folder_docx = os.path.join("reports", "modified_reports")
    create_zip_for_archive(name_zip, folder_pdf, folder_docx)
    # Création de l'archive à envoyer
    modified_or_not()


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



def modified_or_not(path_to_folder1: str = path_to_folder1, path_to_folder2: str = path_to_folder2):
    """
    Creates a zip that contains two folders:
        1) Parlementary files without modifications since the last month
        2) Parlementary files with modifications
    """
    old_files = os.listdir(path_to_folder1)
    new_files = os.listdir(path_to_folder2)
    path_modif = os.path.join(path_to_folder2, "Fiche_Modifiee")
    path_no_modif = os.path.join(path_to_folder2, "Fiche_Non_Modifiee")
    mkdir_ifnotexist(path_no_modif)
    mkdir_ifnotexist(path_modif)
    # Récupération du mois en cours
    today = datetime.datetime.today()
    months = ('Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet',
              'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre')
    today_str = f"{months[today.month-1]}_{today.year}"
    file = "01 - Suivi Territorial plan France relance Ain.pdf"
    for file in new_files:
        if file.endswith("pdf"):
            if pdfdiff(os.path.join(path_to_folder1, file), os.path.join(path_to_folder2, file), ):  # doc identique
                shutil.move(os.path.join(path_to_folder2, file), path_no_modif)
            else:
                shutil.move(os.path.join(path_to_folder2, file), path_modif)
    # Export en zip
    with zipfile.ZipFile("Fiches_Parlementaires_{}.zip".format(today_str), "w", zipfile.ZIP_DEFLATED) as zfile:
        for root, _, files in os.walk(path_to_folder2):
            for file in files:
                zfile.write(os.path.join(root, file))


if __name__ == "__main__":
    main()
