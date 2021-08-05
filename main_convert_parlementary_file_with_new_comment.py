import docx2pdf
import zipfile
import datetime
import os 
import shutil

def main():
    print("Conversion des fiches présentes dans modified_reports")
    docx2pdf.main_docx2pdf_apres_osmose()
    print("Création des archives zip")
    # Obtention du mois de génération des fiches
    today = datetime.datetime.today()
    months = ('Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 
                'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre')
    today_str = f"{months[today.month-1]}_{today.year}"

    mkdir_ifnotexist("archive")
    path = os.path.join("archive", "{}".format(today_str))
    mkdir_ifnotexist(path)
    
    name_zip = os.path.join(path, 'reports_with_new_comment_{}.zip'.format(today_str))
    folder_pdf = "reports_pdf"
    folder_docx = "modified_reports"
    create_zip_for_archive(name_zip, folder_pdf, folder_docx)


def mkdir_ifnotexist(path: str):
    """
    Create a folder if it's doesn't exist
    """
    if not os.path.isdir(path):
        os.mkdir(path)

def create_zip_for_archive(name_zip: str, folder_pdf: str, folder_docx: str):
    """
    Create a zip in archive/Month_Year with 2 folders: folder_pdf and forlder_docx
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
