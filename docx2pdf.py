import os
import re
import pandas as pd
from docx2python import docx2python
from unidecode import unidecode

# Variable globale

# Import référentiel départements
taxo_dep_df = pd.read_csv(os.path.join('refs', 'taxo_deps.csv'), dtype={'dep':str, 'reg':str})

# Définition et création des dossiers
DIR_TO_CONVERT = os.path.join(os.getcwd(), "modified_reports")
OUTPUT_DIR = os.path.join(os.getcwd(), 'reports_pdf')

# main

def main_docx2pdf_avant_osmose():
    mkdir_ifnotexist(OUTPUT_DIR)
    depname2num = creation_dico_dep2name(taxo_dep_df)
    # Mapping docx -> nom pdf
    docx2pdf_filename, doc_odt = docxnames_to_pdfnames(DIR_TO_CONVERT, depname2num)
    check_duclicated_docx(docx2pdf_filename)
    #Archivage a faire ici des docx
    export_to_pdf_avant_osmose(depname2num)


def main_docx2pdf_apres_osmose():
    mkdir_ifnotexist(OUTPUT_DIR)
    depname2num = creation_dico_dep2name(taxo_dep_df)
    # Mapping docx -> nom pdf
    docx2pdf_filename, doc_odt = docxnames_to_pdfnames(DIR_TO_CONVERT, depname2num)
    check_duclicated_docx(docx2pdf_filename)
    #Archivage a faire ici des docx
    export_to_pdf_apres_osmose(docx2pdf_filename, OUTPUT_DIR, doc_odt, depname2num)


def mkdir_ifnotexist(path) :
    if not os.path.isdir(path) :
        os.mkdir(path)
        

def creation_dico_dep2name(taxo_dep_df):
    depname2num = {}
    for i, row in taxo_dep_df.iterrows():
        if row['dep'] != '0':
            depname2num[row['libelle']] = row['dep']
    depnum2name = {v: k for k, v in depname2num.items()}
    return depname2num


def normalisation_name(name):
    # Normalise le nom de la mesure ou volet, notamment pour l'utiliser comme nom de code dans les commentaires
    name = name.lower()
    name = unidecode(name)
    name = re.sub('[^a-z]', ' ',  name)
    name = re.sub(' +', '', name)
    return name


def get_dep_name_from_docx(docx_filename):
    # Extraire le nom du département depuis la page de garde du docx
    content = docx2python(docx_filename)
    # Chercher la ligne "Données pour le département :..."
    for line in content.body[0][0][0]:
        if line.startswith("Données pour le département"):
            expr_with_dep_name = line
            dep_name = expr_with_dep_name.split(':')[-1].strip()
            return dep_name
    raise Exception(f"Pas de nom de département trouvé pour {docx_filename}")


def docxnames_to_pdfnames(base_dir, depname2num):
    # Lister les fichiers à convertir - ignorer les fichiers lock (.docx#)
    docx_filenames = [os.path.join(base_dir, basename) for basename in os.listdir(base_dir) if not basename.endswith('#')]
    docx2pdf_filename = {}
    doc_odt = []
    # Faire correspondre chaque nom de docx vers un nom de pdf - ex : "75 - Suivi Territorial plan France relance Paris.pdf"
    for docx_filename in docx_filenames:
        # Extraire le nom du département
        if docx_filename.endswith("docx"): #Condition pour ne traiter que les docx
            dep_name = get_dep_name_from_docx(docx_filename)
            clean_dep_name = normalisation_name(dep_name)
            pdf_filename = f"{depname2num[dep_name]} - Suivi Territorial plan France relance {dep_name}.pdf"
            # Ajout du nom de fichier original dans le dictionnaire pour vérifier les doublons
            docx2pdf_filename[docx_filename] = pdf_filename
        elif docx_filename.endswith("odt"):
            doc_odt += [docx_filename]
        else: 
            raise ValueError("L'extension du document {} n'est pas pris en charge".format(docx_filename))
    return docx2pdf_filename, doc_odt


def check_duclicated_docx(docx2pdf_filename):
    # Les doublons auront le même nom de fichier pdf.
    # Mapping pdf->[docx]
    pdf2docx_filenames = {}
    for docx_filename, pdf_filename in docx2pdf_filename.items():
        if pdf_filename not in pdf2docx_filenames:
            pdf2docx_filenames[pdf_filename] = []
        pdf2docx_filenames[pdf_filename].append(docx_filename)
    # Afficher les doublons
    flag_duplication = False
    for pdf_filename, docx_filenames in pdf2docx_filenames.items():
        dep_name = pdf_filename.split(os.sep)[-1].split('.')[0].split('relance ')[-1]
        if len(docx_filenames) > 1:
            # Lister les fichiers dupliqués
            print(f"Dupliqués {dep_name} :")
            _ = [print("\t", docx_filename) for docx_filename in docx_filenames]
            flag_duplication = True
    assert not flag_duplication, "Fichiers dupliqués : supprimez les fichiers en trop."
    

def export_to_pdf_apres_osmose(docx2pdf_filename, OUTPUT_DIR, doc_odt, depname2num):
    files_to_convert = docx2pdf_filename.keys()
    for filename in files_to_convert:
        # Conversion en pdf
        os.system(f'libreoffice --headless -convert-to pdf --outdir "{OUTPUT_DIR}" "{filename}"')
        
    for filename in files_to_convert:    
        clean_pdf_filename = docx2pdf_filename[filename]
        pdf_basename = re.sub('.'+filename.split('.')[-1], '.pdf', os.path.basename(filename))
        pdf_filename = os.path.join(OUTPUT_DIR, pdf_basename)
        os.rename(pdf_filename, os.path.join(OUTPUT_DIR, clean_pdf_filename))

    # Traitement des odt
    #Creation du dictionnaire de renommage
    renommage_odt = {}
    for filename in doc_odt:
        if "plan relance" in filename.lower():
            dep_name = filename.split(".odt")[0]
            dep_name = dep_name.split(" ")[-1]
            dep = depname2num[dep_name]
            renommage_odt[filename] = str(dep) + " - Suivi Territorial plan France relance " + str(dep_name) + ".pdf"
        

    for filename in doc_odt:
        # Conversion en pdf
        os.system(f'libreoffice --headless -convert-to pdf --outdir "{OUTPUT_DIR}" "{filename}"')
        
    for filename in doc_odt:
        clean_pdf_filename = renommage_odt[filename]
        pdf_basename = re.sub('.'+filename.split('.')[-1], '.pdf', os.path.basename(filename))
        pdf_filename = os.path.join(OUTPUT_DIR, pdf_basename)
        os.rename(pdf_filename, os.path.join(OUTPUT_DIR, clean_pdf_filename))


def export_to_pdf_avant_osmose(depname2num):
    # Pour les fiches avant le passage osmose
    docx2pdf_filename, doc_odt = docxnames_to_pdfnames(os.path.join(os.getcwd(), "Fiche_Avant_Osmose"), depname2num)
    output = "Fiche_Avant_Osmose_pdf"
    # Conversion docx -> pdf - Peut prendre quelques minutes
    # CAVEAT : Fermer les applications Libreoffice ouverte avant de lancer cette cellule
    files_to_convert = docx2pdf_filename.keys()
    for filename in files_to_convert:
        # Conversion en pdf
        print(filename)
        os.system(f'libreoffice --headless -convert-to pdf --outdir "{output}" {filename}')
        
    for filename in files_to_convert:    
        clean_pdf_filename = docx2pdf_filename[filename]
        pdf_basename = re.sub('.'+filename.split('.')[-1], '.pdf', os.path.basename(filename))
        pdf_filename = os.path.join(output, pdf_basename)
        os.rename(pdf_filename, os.path.join(output, clean_pdf_filename))


if __name__ == "__main__":
    main_docx2pdf_avant_osmose()
