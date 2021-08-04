import docx2pdf
import zipfile
import datetime

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
    
    name_zip = os.path.join(path, 'parlementary_file_with_new_comment{}.zip'.format(today_str))
    f=zipfile.ZipFile(name_zip,'w',zipfile.ZIP_DEFLATED)
    f.write("reports_pdf")
    f.write("modified_reports")
    f.close() 


def mkdir_ifnotexist(path) :
    if not os.path.isdir(path) :
        os.mkdir(path)


if __name__ == "__main__":
    main()
