from unidecode import unidecode
import os
import re

# Permet la génération de word
import docx2python
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage

# Parsing des commentaires
from docx2python import docx2python

# Nettoyage du texte
import html

# Ouvrir les docx en tant que zip pour remplacer manipuler son code XML
import zipfile
# Suppression de dossier temporaire
import shutil

# Logger
import logging
import logging.handlers
# Définition du logger
logger = logging.getLogger("main.transpose_comments")
logger.setLevel(logging.DEBUG)

# Variable globale
BR_TOKEN = '###BR###'  # Les retours à la ligne encodés dans les docx seront remplacés par ce token
template_dir = os.path.join("reports", "reports_word")
modified_docx_dir = os.path.join("reports", "modified_reports")
transposed_docx_dir = os.path.join("reports", "reports_word", "transposed_reports")
image_folder = os.path.join("reports", "reports_word", "reports_images")
avant_osmose = os.path.join("reports", "reports_before_new_comment")
temp_transpo = os.path.join("reports",'reports_word', 'temp_transposition')
# Mesures des fiches V2 contenant des commentaires
volet2mesures = {
    'Ecologie': [#'Bonus écologique',
                  #"MaPrimeRénov'",
                  #'Modernisation des filières automobiles et aéronautiques',
                  #'Prime à la conversion des agroéquipements',
                  #'Prime à la conversion des véhicules légers',
                  #'Réhabilitation Friches (urbaines et sites pollués)',
                  'Rénovation bâtiments Etat'],
 'Compétitivité': ['AAP Industrie : Soutien aux projets industriels territoires',
                  'AAP Industrie : Sécurisation approvisionnements critiques',
                  'France Num : aide à la numérisation des TPE,PME,ETI',
                  #'Industrie du futur',
                  'Renforcement subventions Business France',
                  #'Soutien aux filières culturelles (cinéma, audiovisuel, musique, numérique, livre)'
                  ],
 'Cohésion': [
             #    'Apprentissage',
             #     'Contrats Initiatives Emploi (CIE) Jeunes',
             #     'Contrats de professionnalisation',
             #     'Garantie jeunes',
             #     'Parcours emploi compétences (PEC) Jeunes',
             #     "Prime à l'embauche des jeunes",
             #     "Prime à l'embauche pour les travailleurs handicapés",
             #     'Service civique'
             ]
}

comment_prefixes = set(['Espace Commentaires\xa0:', 'Espace Commentaires :',
                        'Exemples de lauréats :', 'Exemples de lauréats\xa0:',
                        "Commentaires généraux\xa0:", "Commentaires généraux :",
                        "Volet\xa0: Ecologie", "Volet\xa0: Compétitivité", "Volet\xa0: Cohésion",
                        "Volet : Ecologie", "Volet : Compétitivité", "Volet : Cohésion"])

# main


def main_transpose_comments():
    # Vérification de l'existence des dossiers
    mkdir_ifnotexist(avant_osmose)
    mkdir_ifnotexist(transposed_docx_dir)
    mkdir_ifnotexist(image_folder)
    mkdir_ifnotexist(modified_docx_dir)
    mkdir_ifnotexist(temp_transpo)
    print(len(os.listdir(modified_docx_dir)))
    assert len(os.listdir(modified_docx_dir)) > 0, f"Le dossier {modified_docx_dir} est vide. Vous devez y placer les fichiers docx contenant les commentaires à déplacer."

    templates = [os.path.join(template_dir, filename) for filename in os.listdir(template_dir) if filename.endswith('docx')]
    modified_docx = [os.path.join(modified_docx_dir, filename) for filename in os.listdir(modified_docx_dir) if filename.endswith('docx')]
    template2modified_docx = map_templates_to_modified_reports(templates, modified_docx)
    errors = transpose_modification_to_new_reports(template2modified_docx)
    if errors:
        logger.info("Erreurs rencontrées :", errors)
    assert len(os.listdir(transposed_docx_dir)) == 109
    shutil.rmtree(os.path.join("reports", "reports_word", "temp_transposition"))


# Fonction nécessaire


def mkdir_ifnotexist(path: str):
    """
    Create a folder if it's doesn't exist
    """
    if not os.path.isdir(path):
        os.mkdir(path)


def normalize_name(name: str) -> str:
    """
    Normalize a str: delete whitespace, special characters, put in lowercase and keep only letters
    """
    name = name.lower()
    name = unidecode(name)
    name = re.sub('[^a-z]', ' ', name)
    name = re.sub(' +', '', name)
    return name


def flatten(L: list):
    """
    Transform list like: [[., ., [[.]]]] into a classical structure list: [., ., .]
    """
    if type(L) is list:
        for item in L:
            yield from flatten(item)
    else:
        yield L


def gen_unit_list(L: list):
    """
    Creates a generator that generates the smallest list for nested lists:
    [ [[.], [., .], [.]], [ [a], [.] ] ] --> [.], [., .], [.], [a], [.]
    """
    if (type(L) is list) and (len(L) > 0) and (type(L[0]) is not list):
        yield L
    else:
        for item in L:
            yield from gen_unit_list(item)


def reformat_bullet_point(text: str) -> str:
    """
    When docx2python get comments form parlementary file and where, in these, there are enumerations,
    docx2python add special character. Format them into a correct shape
    """
    text = re.sub('--\t\t', '- ', text)
    text = re.sub('^--\t-*', '- ', text)
    return text


def reformat_url(text: str) -> str:
    """
    Format text into a correct URL shape
    """
    regex_clean = re.compile('<a href.*?>')
    text = re.sub(regex_clean, '', text)
    text = re.sub('</a>', "", text)
    return text


def fix_vanishing_break_lines(text: str) -> str:
    """
    In modify_docx_break_line, we replace break line (create by Enter or Shift+ Enter) by a special token.
    Replace this token by "\n"
    """
    text = re.sub(BR_TOKEN, '\n', text)
    return text


def extract_comment(textbox_content: list) -> list:
    """
    Extract comment from textbox_content. Comments are always preceded by one of these prefixes:
            - 'Espace Commentaires\xa0:'
            - 'Espace Commentaires :'
            - 'Exemples de lauréats :'
            - 'Exemples de lauréats\xa0:'
            - "Commentaires généraux\xa0:"
            - "Commentaires généraux :"
            - "Volet\xa0: Ecologie"
            - "Volet\xa0: Compétitivité"
            - "Volet\xa0: Cohésion"
            - "Volet : Ecologie"
            - "Volet : Compétitivité"
            - "Volet : Cohésion"
    """
    # Cleaning section
    texts = []
    for text in textbox_content:
        text = html.unescape(text)
        text = reformat_bullet_point(text)
        text = reformat_url(text)
        text = fix_vanishing_break_lines(text)
        text = text.strip()
        texts.append(text)
    textbox_content = texts

    # Concatenation
    textbox_content = [text.strip() for text in textbox_content]
    textbox_content = '\n'.join(textbox_content)
    textbox_content = textbox_content.strip()

    # Retirer un potentiel préfix (Espace Commentaires ...)
    textbox_content = re.sub('^[0-9]+[\r\n]+[0-9]+', '', textbox_content).strip()
    prefix_clean = False
    while not prefix_clean:
        # Les préfixes étant déclarés dans un set, dès lors qu'on retrouve un préfixe à retirer,
        # on re-parcourt le set de préfixes
        prefix_clean = True
        for prefix in comment_prefixes:
            if textbox_content.startswith(prefix):
                textbox_content = textbox_content.replace(prefix, "", 1)
                textbox_content = textbox_content.strip()
                prefix_clean = False
            if textbox_content.endswith(prefix):
                textbox_content = textbox_content[:-len(prefix)]
                textbox_content = textbox_content.strip()
                prefix_clean = False
    textbox_content = textbox_content.strip()
    # Carriage pour conserver les retours à la ligne
    textbox_content = re.sub("\n", "\r\n", textbox_content)
    return textbox_content


def alternate_texts_and_images(doc: DocxTemplate, textbox_content: str) -> list:
    """
    Comments may contain pictures. To keep them, we create a liste to store text and image path
    The image size is fixed since it cannot be extracted from the initial document

    Return Structure:
        list: [{text:..., image:...}, {text:..., image:...}]
    """
    # Renvoie une liste [{text:..., image:...}, {text:..., image:...}...]
    r = re.compile("----media/(.*?)----")  # Pattern pour les images
    image_names = r.findall(textbox_content) + [None]
    texts = r.split(textbox_content)
    frameworks = []
    for text, image_basename in zip(texts[0::2], image_names):
        if image_basename is not None:
            image_path = os.path.join(image_folder, image_basename)
            frameworks.append({'text': text, 'image': InlineImage(doc, image_path, height=Mm(40))})
        else:
            frameworks.append({'text': text, 'image': ''})
    return frameworks


def get_mesure_to_comment(doc: DocxTemplate, content: docx2python, volet2mesures: dict) -> dict:
    """
    Get measures (not volet) that have comments space and comment space's content
    Returns:
        dict: {key: measure name, value: comment space's content}
    """
    mesure2comment = {}

    # Pattern regex pour attraper le nom des mesures
    list_mesures = [mesure for mesures in volet2mesures.values() for mesure in mesures]
    re_group_mesures = "(" + '|'.join(list_mesures) + ")"
    re_title_mesure_pattern = f'(\d - <a href=.*>{re_group_mesures}</a>)'

    current_mesure = None
    num_blocks_to_pass = 0
    for text_list_block in content.body:
        text_list = list(flatten(text_list_block))
        if current_mesure is None:
            # On veut le nom de la mesure
            text_unit = " ".join(text_list)
            title_mesures = re.findall(re_title_mesure_pattern, text_unit)
            if len(title_mesures) > 0:
                current_mesure = title_mesures[0][1]
                num_blocks_to_pass = 6
        else:
            # On veut récupérer le commentaire
            # il faudra passer 3 tableaux + 3 retours à la ligne
            if num_blocks_to_pass == 0:

                # On extrait le commentaire
                text_list = list(flatten(text_list_block))
                textbox_content = extract_comment(text_list)
                frameworks = alternate_texts_and_images(doc, textbox_content)

                # On associe volet et commentaire
                encoded_mesure = normalize_name(current_mesure)
                mesure2comment[encoded_mesure] = frameworks
                # Reinit la mesure courante pour récupérer la suivante
                current_mesure = None
            num_blocks_to_pass -= 1
    assert len(mesure2comment) == len(list_mesures), f"{len(mesure2comment)} != {len(list_mesures)} attendues"
    return mesure2comment


def get_volet_to_comment(doc: DocxTemplate, content, volet2mesures: dict) -> dict:
    """
    Get volet that have comments space and comment space's content

    Returns:
        dict: {key: measure name, value: comment space's content}
    """
    volet2comment = {}
    volet = None
    text_unit_generator = gen_unit_list(content.body)
    volet_names_regex = "(" + '|'.join([volet for volet in volet2mesures]) + ")"  # -> (Ecologie|Compétitivité|Cohésion)

    for text_list in text_unit_generator:
        # On cherche à trouver le titre contenant le nom du volet (partie else)
        # puis on saura que le texte suivant contiendra le commentaire
        if volet is not None:
            # On extrait le commentaire
            textbox_content = extract_comment(text_list)
            frameworks = alternate_texts_and_images(doc, textbox_content)

            # On associe volet et commentaire
            encoded_volet = normalize_name(volet)
            volet2comment[encoded_volet] = frameworks

            # Reinitialise volet
            volet = None
        else:
            text_list = ' '.join(text_list)
            patterns = re.findall(f'(Volet [1-3] : {volet_names_regex})', text_list)
            if len(patterns) > 0:
                # On attrape le titre du volet et récupère le nom du volet
                volet = patterns[-1][-1]
    assert len(volet2comment) == 3
    return volet2comment


def modify_docx_break_line(input_docx: str, output_docx: str):
    """
    Replace <w:br/> XML character (Shift+Enter) by a special token ###BR###
    in order to translate into a classic break line when we are (re)writting parlementary file
    """
    # ###BR### replace by "\r\n"
    with open(input_docx, 'rb') as f:
        # Ouvrir le docx entrant en tant que zip et ouvrir le docx de sortie en tant que zip
        with zipfile.ZipFile(f) as inzip, zipfile.ZipFile(output_docx, "w") as outzip:
            # Itérer sur tous les fichiers du zip entrant
            for inzipinfo in inzip.infolist():
                # Ouvrir le fichier courant
                with inzip.open(inzipinfo) as infile:
                    # Si le fichier est celui qui contient le texte du document
                    if inzipinfo.filename == "word/document.xml":  # word/document.xml will always contains the textual content
                        xml_content = infile.read().decode()
                        # Modify the content of the file by replacing a string
                        xml_content = re.sub('<w:br/>', f'<w:t>{BR_TOKEN}</w:t>', xml_content)
                        # Write content
                        outzip.writestr(inzipinfo.filename, xml_content)
                    else:  # Si pas de texte, simple recopiage du contenu du fichier
                        outzip.writestr(inzipinfo.filename, infile.read())


def transpose_comments(src_filename: str, template_filename: str, output_filename: str, volet2mesures: dict) -> str:
    """
    Complete comments spaces in src_filename with template_filename comments. Save the new document in output_filename
    """
    # Remplacer les retour à la ligne par un token spécial
    tmp_docx = os.path.join(temp_transpo, 'temp.docx')
    modify_docx_break_line(src_filename, tmp_docx)

    # Lecture du document
    content = docx2python(src_filename, image_folder=image_folder)
    doc_template = DocxTemplate(template_filename)

    # Parse les commentaires sous les volets et mesures
    mesure2comment = get_mesure_to_comment(doc_template, content, volet2mesures)
    volet2comment = get_volet_to_comment(doc_template, content, volet2mesures)
    context = {**mesure2comment, **volet2comment}

    # On génère un nouveau document avec les commentaires recopiés
    doc_template.render(context, autoescape=True)
    doc_template.save(output_filename)


def fill_template(template_filename: str, output_filename: str, volet2mesures: dict) -> str:
    """
    Complete space comment into template_filename by void.
    In fact, template_filename's space comment contain a tag allowing text insertion. This function delete this tag.
    """
    ordered_volets = list(volet2mesures.keys())

    context = {normalize_name(volet): [{'text': '', 'image': ''}] for volet in ordered_volets}
    doc = DocxTemplate(template_filename)
    doc.render(context, autoescape=True)
    doc.save(output_filename)


def map_templates_to_modified_reports(templates: list, modified_docx: list) -> dict:
    """
    After the comments phase parlementary file's name could be modified.
    Create a dictionnary to match original name and name after comment phase

    Return:
        dict: {key: new name, value: old name}
    """
    mapping = {filename: None for filename in templates}

    # Faire correspondre le nom des départements encodés vers le bon template
    encoded_dep_name2template = {}
    for filename in mapping:
        raw_dep_name = filename.split('_')[-1].split('.')[0]
        encoded_dep_name = normalize_name(raw_dep_name)
        encoded_dep_name2template[encoded_dep_name] = filename
    # assert len(encoded_dep_name2template) == 109

    # Faire correspondre le nom du département
    duplicated_dep = []
    for modified in modified_docx:
        content = docx2python(modified)
        expr_with_dep_name = content.body[0][0][0][7]
        logger.info(f"Extrait de {modified} : ", expr_with_dep_name)
        expr_with_dep_name.split(':')
        dep_name = expr_with_dep_name.split(':')[-1].strip()
        clean_dep_name = normalize_name(dep_name)
        target_template = encoded_dep_name2template[clean_dep_name]
        if mapping[target_template] is None:
            mapping[target_template] = modified
        else:
            duplicated_dep.append(dep_name)
            logger.info(f"!!! {target_template} is not None -> probably duplicated \n----See {modified}")
    if duplicated_dep != []:
        logger.info("Fiches dupliquées (à retirer manuellement puis relancer le script) :\n", str(duplicated_dep))
    logger.info(f"{len(mapping)} hits")
    return mapping


# Les nouveaux documents contiennent le texte transposé et sous le même nom que leur template mais
# dans un dossier différent


def transpose_modification_to_new_reports(template2modified_docx: dict):
    """
    Transpose every comment from old parlementary_file to new parlementary file.
    template2modified_docx is a crossing dictionnary between original name and name after comment phase
    """
    hit, unhit = 0, 0
    errors = []
    for template_path, modified_docx_path in template2modified_docx.items():
        output_basename = template_path.split(os.sep)[-1]
        output_path = os.path.join(transposed_docx_dir, output_basename)
        fill_template(template_path, os.path.join(avant_osmose, output_basename), volet2mesures)

        if modified_docx_path is None:
            unhit += 1
            logger.info(f'Pas de transposition pour {template_path}')
            fill_template(template_path, output_path, volet2mesures)
        else:
            logger.info(f'Transpose {template_path} vers {output_path}')
            try:
                transpose_comments(template_path, modified_docx_path, output_path, volet2mesures)
                hit += 1
            except BaseException as e:
                logger.info(f"** Transposition impossible. Génération d'une fiche vide dans {template_path}**")
                fill_template(template_path, output_path, volet2mesures)
                unhit += 1
                logger.info("erreur rencontrée")
                errors.append(e)
    if errors:
        return errors
    logger.info(f"Hit : {hit} | Unhit : {unhit}")


if __name__ == "__main__":
    main_transpose_comments()
