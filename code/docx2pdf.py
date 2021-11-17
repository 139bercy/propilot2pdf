import os
import re
import pandas as pd
import shutil
from docx2python import docx2python
from unidecode import unidecode
# Barre de progression
from tqdm import tqdm
import time

# Logger
import logging
import logging.handlers
# Définition du logger
logger = logging.getLogger("main.docx2pdf")
logger.setLevel(logging.DEBUG)

# Variable globale

# Import référentiel départements
taxo_dep_df = pd.read_csv(os.path.join('refs', 'taxo_deps.csv'), dtype={'dep': str, 'reg': str})

# Définition et création des dossiers
DIR_TO_CONVERT = os.path.join(os.getcwd(), "reports", "modified_reports")
OUTPUT_DIR = os.path.join(os.getcwd(), "reports", 'reports_pdf')
avant_osmose_pdf = os.path.join("reports", "reports_before_new_comment_pdf")
DIR_COPY_DOCX = os.path.join(os.getcwd(), "reports", "temp_docx")

# main


def main_docx2pdf_avant_osmose():
    mkdir_ifnotexist(OUTPUT_DIR)
    mkdir_ifnotexist(DIR_COPY_DOCX)
    mkdir_ifnotexist(avant_osmose_pdf)
    depname2num = create_dico_dep2num(taxo_dep_df)
    # Mapping docx -> nom pdf
    docx2pdf_filename, doc_odt = docxnames_to_pdfnames(DIR_TO_CONVERT, depname2num)
    check_duplicated_docx(docx2pdf_filename)
    # Archivage a faire ici des docx
    export_to_pdf_avant_osmose(depname2num)
    shutil.rmtree(os.path.join("reports", "temp_docx"))


def main_docx2pdf_apres_osmose():
    mkdir_ifnotexist(OUTPUT_DIR)
    mkdir_ifnotexist(DIR_COPY_DOCX)
    mkdir_ifnotexist(avant_osmose_pdf)
    depname2num = create_dico_dep2num(taxo_dep_df)
    # Mapping docx -> nom pdf
    docx2pdf_filename, doc_odt = docxnames_to_pdfnames(DIR_TO_CONVERT, depname2num)
    check_duplicated_docx(docx2pdf_filename)
    # Archivage a faire ici des docx
    export_to_pdf_apres_osmose(docx2pdf_filename, OUTPUT_DIR, doc_odt, depname2num, taxo_dep_df)


def mkdir_ifnotexist(path: str):
    """
    Creates a folder if it's doesn't exist
    """
    if not os.path.isdir(path):
        os.mkdir(path)


def create_dico_dep2num(taxo_dep_df_: pd.DataFrame) -> dict:
    """
    Creates crossing dictionnary between department's name and department's number
    """
    depname2num = {}
    for i, row in taxo_dep_df_.iterrows():
        if row['dep'] != '0':
            depname2num[row['libelle']] = row['dep']
    # depnum2name = {v: k for k, v in depname2num.items()}
    return depname2num


def normalize_name(name: str) -> str:
    """
    Normalize a str: delete whitespace, special characters, put in lowercase and keep only letters
    """
    name = name.lower()
    name = unidecode(name)
    name = re.sub('[^a-z]', ' ', name)
    name = re.sub(' +', '', name)
    return name


def get_dep_name_from_docx(docx_filename: str, taxo_dep_df_: pd.DataFrame = taxo_dep_df) -> str:
    """
    Extract department's name from docx's front page
    """
    try:
        content = docx2python(docx_filename)
        # Chercher la ligne "Données pour le département :..."
        for line in content.body[0][0][0]:
            if line.startswith("Données pour le département"):
                expr_with_dep_name = line
                dep_name = expr_with_dep_name.split(':')[-1].strip()
                return dep_name
    except BaseException as e:
        logger.info(f"Pas de nom de département trouvé pour {docx_filename} dans la page de garde.")
        logger.error(repr(e))
        return detect_dep_in_filename(taxo_dep_df_, docx_filename)


def detect_dep_in_filename(taxo_dep_df_: pd.DataFrame, docx_filename: str) -> str:
    """
    Keep departement in filename

    Returns:
        str: departement's name
    """
    for dep in taxo_dep_df_.libelle.unique():
        if dep in docx_filename:
            return dep

def docxnames_to_pdfnames(base_dir: str, depname2num: dict) -> list:
    """
    Creates crossing dictionnary between a file in base_dir and a required output file
    Returns:
        list[0]: Crossing dictionnary between docx name and pdf name
        list[1]: Crossing dictionnary between odt name and pdf name
    """
    docx_filenames = [os.path.join(base_dir, basename) for basename in os.listdir(base_dir) if not basename.endswith('#')]
    docx2pdf_filename = {}
    doc_odt = []
    # Faire correspondre chaque nom de docx vers un nom de pdf - ex : "75 - Suivi Territorial plan France relance Paris.pdf"
    for docx_filename in docx_filenames:
        # Extraire le nom du département
        try: # A tester
            if docx_filename.endswith("docx"):  # Condition pour ne traiter que les docx (ignore les #docx)
                dep_name = get_dep_name_from_docx(docx_filename)
                pdf_filename = f"{depname2num[dep_name]} - Suivi Territorial plan France relance {dep_name}.pdf"
                # Ajout du nom de fichier original dans le dictionnaire pour vérifier les doublons
                docx2pdf_filename[docx_filename] = pdf_filename
            elif docx_filename.endswith("odt"):
                doc_odt += [docx_filename]
            else:
                raise ValueError("L'extension du document {} n'est pas pris en charge".format(docx_filename))
        except: 
            pass
    return docx2pdf_filename, doc_odt


def check_duplicated_docx(docx2pdf_filename: dict, taxo_dep_df_: pd.DataFrame = taxo_dep_df):
    """
    Check if there are duplicate department in docx2pdf_filename and remove them
        - If docx files, keep the last modified and remove others
        - If odt files, remove docx file with the same department's name

    Input (for our main):
        docx2pdf_filename: {key: docx_name into modified_reports folder, value: target pdf name}
    """
    # Les doublons auront le même nom de fichier pdf.
    # Mapping pdf->[docx]
    pdf2docx_filenames = {}
    for docx_filename, pdf_filename in docx2pdf_filename.items():
        if pdf_filename not in pdf2docx_filenames:
            pdf2docx_filenames[pdf_filename] = []
        pdf2docx_filenames[pdf_filename].append(docx_filename)
    # Afficher les doublons
    for pdf_filename, docx_filenames in pdf2docx_filenames.items():
        dep_name = pdf_filename.split(os.sep)[-1].split('.')[0].split('relance ')[-1]
        if len(docx_filenames) > 1:
            # Lister les fichiers dupliqués
            logger.info(f"Dupliqués {dep_name} :")
            _ = [logger.info("\t", docx_filename) for docx_filename in docx_filenames]
            # Récupération des dates de dernières modifications
            lastmodified1 =  os.stat(docx_filenames[0]).st_mtime  # cf https://docs.python.org/3/library/stat.html
            lastmodified2 = os.stat(docx_filenames[1]).st_mtime
            # On garde la fiche la + récente
            if lastmodified1 > lastmodified2:
                os.remove(docx_filenames[1])
            else:
                os.remove(docx_filenames[0])
    # Ajouter le cas suppression des docx associés aux odt. 
    for filename in os.listdir(os.path.join('reports', 'modified_reports')):
        if filename.endswith('.odt'):
            # On récupère le nom du departement dans le odt
            dep = detect_dep_in_filename(taxo_dep_df_, filename)
            # Si la fiche initial au format docx est encore dans le dossier, alors on doit la supprimer
            if 'Suivi Territorial plan relance {}.docx'.format(dep) in os.listdir(os.path.join('reports', 'modified_reports')):
                # On part du principe (au vu de l'usage) que la fiche docx n'a pas été modifiée
                os.remove(os.path.join('reports', 'modified_reports','Suivi Territorial plan relance {}.docx'.format(dep)))


def export_to_pdf_apres_osmose(docx2pdf_filename: dict, out_dir: str, doc_odt: dict, depname2num: dict, taxo_dep_df_: pd.DataFrame):
    """
    Convert into pdf all keys in docx2pdf_filename and doc_odt.
    For our main: Convert all files to pdf from modified_report folder into reports_pdf
    """
    files_to_convert = list(docx2pdf_filename.keys())
    for filename in tqdm(files_to_convert, desc="Conversion PDF des fiches docx"):
        # Effectuer la copie : les noms de fichiers comportant un espace ou une apostrophe rencontrent
        # font planter la conversion
        clean_path = rename_docx_without_buggy_chars(filename)
        # Renommage des clés docx2pdf_filename pour faire correspondre les nouveaux noms de docx
        # vers les bons pdf
        replace_key(docx2pdf_filename, filename, clean_path)
        # Conversion en pdf
        # > test.log 2> warning.log permet de rediriger la sortie vers test.log et les warning vers warning.log
        os.system(f'libreoffice --headless -convert-to pdf --outdir "{out_dir}" "{clean_path}" > test.log 2> warning.log')
    time.sleep(5)
    for filename in docx2pdf_filename:
        clean_pdf_filename = docx2pdf_filename[filename]
        pdf_basename = re.sub('.' + filename.split('.')[-1], '.pdf', os.path.basename(filename))
        pdf_filename = os.path.join(out_dir, pdf_basename)
        os.rename(pdf_filename, os.path.join(out_dir, clean_pdf_filename))

    # Traitement des odt
    # create du dictionnaire de renommage
    renommage_odt = {}
    for filename in doc_odt:
        if "plan relance" in filename.lower():
            dep_name = detect_dep_in_filename(taxo_dep_df_, filename)
            dep = depname2num[dep_name]
            renommage_odt[filename] = str(dep) + " - Suivi Territorial plan France relance " + str(dep_name) + ".pdf"

    for filename in tqdm(doc_odt, desc="Conversion PDF des fiches odt"):
        # Effectuer la copie : les noms de fichiers comportant un espace ou une apostrophe rencontrent
        # font planter la conversion
        clean_path = rename_docx_without_buggy_chars(filename)
        # Renommage des clés docx2pdf_filename pour faire correspondre les nouveaux noms de docx
        # vers les bons pdf
        renommage_odt[clean_path] = renommage_odt[filename]
        del renommage_odt[filename]
        # Conversion en pdf
        os.system(f'libreoffice --headless -convert-to pdf --outdir "{out_dir}" "{clean_path}" > test.log 2> warning.log')
    time.sleep(5)
    for filename in renommage_odt:
        clean_pdf_filename = renommage_odt[filename]
        pdf_basename = re.sub('.' + filename.split('.')[-1], '.pdf', os.path.basename(filename))
        pdf_filename = os.path.join(out_dir, pdf_basename)
        os.rename(pdf_filename, os.path.join(out_dir, clean_pdf_filename))


def export_to_pdf_avant_osmose(depname2num: dict):
    """
    Convert docx into pdf.
    For our main: Convert all docx to pdf from reports_before_new_comment folder to reports_before_new_comment_pdf
    """
    # Pour les fiches avant le passage osmose
    docx2pdf_filename, doc_odt = docxnames_to_pdfnames(os.path.join(os.getcwd(), "reports", "reports_before_new_comment"), depname2num)
    output = os.path.join("reports", "reports_before_new_comment_pdf")
    # Conversion docx -> pdf - Peut prendre quelques minutes
    # CAVEAT : Fermer les applications Libreoffice ouverte avant de lancer cette cellule
    files_to_convert = list(docx2pdf_filename.keys())
    for filename in tqdm(files_to_convert, desc="Conversion des fiches docx"):
        # Effectuer la copie : les noms de fichiers comportant un espace ou une apostrophe rencontrent
        # font planter la conversion
        clean_path = rename_docx_without_buggy_chars(filename)
        # Renommage des clés docx2pdf_filename pour faire correspondre les nouveaux noms de docx
        # vers les bons pdf
        replace_key(docx2pdf_filename, filename, clean_path)
        os.system(f'libreoffice --headless -convert-to pdf --outdir "{output}" {clean_path} > test.log 2> warning.log')
    time.sleep(5)
    for filename in tqdm(docx2pdf_filename):
        clean_pdf_filename = docx2pdf_filename[filename]
        pdf_basename = re.sub('.' + filename.split('.')[-1], '.pdf', os.path.basename(filename))
        pdf_filename = os.path.join(output, pdf_basename)
        os.rename(pdf_filename, os.path.join(output, clean_pdf_filename))


def replace_key(dictionary: dict, old_key: str, new_key: str):
    """
    Replace a key (old key) by another key (new_key) into a dictionnary
    """
    dictionary[new_key] = dictionary[old_key]
    del dictionary[old_key]


def rename_docx_without_buggy_chars(src_file: str) -> str:
    """
    Replace white space and apostrophe by #
    With white space and apostrophe the command line to convert into pdf doesn't work
    """
    clean_filename = re.sub("[ ']", "#", src_file)
    clean_path = os.path.join(DIR_COPY_DOCX, clean_filename.split(os.sep)[-1])  # Mettre la copie dans un autre dossier
    if os.path.exists(clean_path):
        os.remove(clean_path)
    shutil.copyfile(src_file, clean_path)
    return clean_path


if __name__ == "__main__":
    main_docx2pdf_avant_osmose()
