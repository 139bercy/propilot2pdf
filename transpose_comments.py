from unidecode import unidecode
import os
import urllib.request
import json
import re
import datetime

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pprint import pprint

# Permet la génération de word
import docx
from docx import Document
from docxcompose.composer import Composer
from docxtpl import DocxTemplate, R, Listing, InlineImage
from docx.shared import Mm

# Parsing des commentaires
from docx2python import docx2python

# Nettoyage du texte
import html
from collections import Counter


# Variable globale

template_dir = "reports_word"
modified_docx_dir = "modified_reports"
transposed_docx_dir = os.path.join("reports_word", "transposed_reports")
image_folder = os.path.join("reports_word", "reports_images")
avant_osmose = "Fiche_Avant_Osmose"
# Les mesures des fiches v2
volet2mesures = {'Ecologie': ['Bonus écologique',
  "MaPrimeRénov'",
  'Modernisation des filières automobiles et aéronautiques',
  'Prime à la conversion des agroéquipements',
  'Prime à la conversion des véhicules légers',
  'Réhabilitation Friches (urbaines et sites pollués)',
  'Rénovation bâtiments Etat'],
  'Compétitivité': ['AAP Industrie : Soutien aux projets industriels territoires',
  'AAP Industrie : Sécurisation approvisionnements critiques',
  'France Num : aide à la numérisation des TPE,PME,ETI',
  'Industrie du futur',
  'Renforcement subventions Business France',
  'Soutien aux filières culturelles (cinéma, audiovisuel, musique, numérique, livre)'],
  'Cohésion': ['Apprentissage',
  'Contrats Initiatives Emploi (CIE) Jeunes',
  'Contrats de professionnalisation',
  'Garantie jeunes',
  'Parcours emploi compétences (PEC) Jeunes',
  "Prime à l'embauche des jeunes",
  "Prime à l'embauche pour les travailleurs handicapés",
  'Service civique']}

comment_prefixes = set(['Espace Commentaires\xa0:', 'Espace Commentaires :', 
                        'Exemples de lauréats :', 'Exemples de lauréats\xa0:', 
                       "Commentaires généraux\xa0:", "Commentaires généraux :", 
                       "Volet\xa0: Ecologie", "Volet\xa0: Compétitivité", "Volet\xa0: Cohésion",
                       "Volet : Ecologie", "Volet : Compétitivité", "Volet : Cohésion",])

# main

def main_transpose_comments():
    #Vérification de l'existence des dossiers
    mkdir_ifnotexist(avant_osmose)    
    mkdir_ifnotexist(transposed_docx_dir)
    mkdir_ifnotexist(image_folder)
    mkdir_ifnotexist(modified_docx_dir)
    assert len(os.listdir(modified_docx_dir)) > 0, f"Le dossier {modified_docx_dir} est vide. Vous devez y placer les fichiers docx contenant les commentaires à déplacer."

    templates = [os.path.join(template_dir, filename) for filename in os.listdir(template_dir) if filename.endswith('docx')]
    modified_docx = [os.path.join(modified_docx_dir, filename) for filename in os.listdir(modified_docx_dir) if filename.endswith('docx')]
    template2modified_docx = map_templates_to_modified_reports(templates, modified_docx)
    transpose_modification_to_new_reports(template2modified_docx)
    assert len(os.listdir(transposed_docx_dir)) == 110

# Fonction nécessaire

def mkdir_ifnotexist(path) :
    if not os.path.isdir(path) :
        os.mkdir(path)
        

def encode_name(name):
    # Normalise le nom de la mesure ou volet, notamment pour l'utiliser comme nom de code dans les commentaires
    name = name.lower()
    name = unidecode(name)
    name = re.sub('[^a-z]', ' ',  name)
    name = re.sub(' +', '', name)
    return name

def flatten(L):
    # Aplatir une liste imbriquée [[., ., [[.]]]] -> [., ., .]
    if type(L) is list:
        for item in L:
            yield from flatten(item)
    else:
        yield L
        

def gen_unit_list(L):
    if (type(L) is list) and (len(L) > 0) and (type(L[0]) is not list):
        yield L
    else:
        for item in L:
            yield from gen_unit_list(item)


def count_occurence(texts):
    counter = Counter()
    for text in texts:
        counter[text] += 1
    return counter


def reformat_bullet_point(text):
    return re.sub('^--\t', '- ', text)


def reformat_url(text):
    regex_clean = re.compile('<a href.*?>')
    text = re.sub(regex_clean, '', text)
    text = re.sub('</a>', "", text)
    return text


def extract_comment(textbox_content):
    occurences = count_occurence(textbox_content)
    text_to_keep = []
    flag_text = None
    for text in textbox_content:
        if occurences[text] == 2 and flag_text == text:
            # On rencontre le début du commentaire pour la deuxième fois.
            break
        elif occurences[text] == 2 and flag_text == None and not text.isdigit():
            # On renconre le début du commentaire pour la première fois
            flag_text = text
        text_to_keep.append(text)

    textbox_content = text_to_keep
    
    # Cleaning section
    texts = []
    for text in textbox_content:
        text = html.unescape(text)
        text = reformat_bullet_point(text)
        text = reformat_url(text)
        
        texts.append(text)
    textbox_content = texts
    ###

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
    # Changer tous les "plan de relance" en "plan France Relance" ------------------------------------ !!!!!!!!!!!!!!!!!!!!
    textbox_content = re.sub("plan de relance", "plan France Relance", textbox_content)
    return textbox_content
    

def alternate_texts_and_images(doc, textbox_content):
    r = re.compile("----media/(.*?)----")
    image_names = r.findall(textbox_content) + [None]
    texts = r.split(textbox_content)
    
    frameworks = []
    for text, image_basename in zip(texts[0::2], image_names):
        if image_basename is not None:
            image_path = os.path.join(image_folder, image_basename)
            frameworks.append({'text': text, 'image': InlineImage(doc, image_path, height=Mm(40))})
        else:
            frameworks.append({'text': text, 'image': ''})
        
        if text.strip().isdigit():
            print('---- Récupération ratée ! Probablement un numéro de page attrapé au lieu du texte')
    return frameworks


def get_volet_to_comment(doc, content, volets2mesures):
    volet2comment = {}
    body = content.body
    count_mesures = 0
    for text_list in gen_unit_list(content.body):                
        volet = None
        for text in text_list:
            clean_text = text.lower().strip()
            clean_text = re.sub('\xa0', ' ', clean_text)
            if clean_text.startswith('volet :'):
                volet = clean_text.split(':')[-1].strip()
                
        if volet is not None:
            # Filtrer des retours à la lignes
            textbox_content = text_list
            textbox_content = textbox_content[:-7]

            # On extrait le commentaire
            textbox_content = extract_comment(textbox_content)
            frameworks = alternate_texts_and_images(doc, textbox_content)

            # On associe volet et commentaire
            encoded_volet = encode_name(volet)
            volet2comment[encoded_volet] = frameworks
    return volet2comment


def get_mesure_to_comment(doc, content, volet2mesures):
    mesure2comment = {}
    body = content.body
    # Les mesures doivent apparaitre dans le même ordre que le document
    ordered_mesures = [mesure for mesures in volet2mesures.values() for mesure in mesures]

    for i in range(1, len(ordered_mesures)+1):
        # On extrait la partie textuelle à copier
        textbox_content = body[2 + 6 * i][0][0]

        # Filtrer des retours à la lignes et potentiels num page
        while len(textbox_content) > 0 and (textbox_content[0] == '' or textbox_content[0].strip().isdigit()):
            textbox_content = textbox_content[1:]
                
        # On extrait le commentaire
        textbox_content = extract_comment(textbox_content)
        frameworks = alternate_texts_and_images(doc, textbox_content)
        
        # On associe la mesure au commentaire
        encoded_mesure = encode_name(ordered_mesures[i-1])
        mesure2comment[encoded_mesure] = frameworks  #textbox_content
    assert len(mesure2comment) == len(ordered_mesures)
    
    return mesure2comment


def transpose_comments(src_filename, template_filename, output_filename, volet2mesures):
    # Lecture du document
    content = docx2python(src_filename, image_folder=image_folder)
    doc_template = DocxTemplate(template_filename)
    
    # Parse les commentaires sous les volets et mesures
    mesure2comment = get_mesure_to_comment(doc_template, content, volet2mesures)
    volet2comment = get_volet_to_comment(doc_template, content, volet2mesures)
    context = {**mesure2comment, **volet2comment}
    dep_name = output_filename.split('_')[-1].split('.docx')[0]

    # On génère un nouveau document avec les commentaires recopiés
    doc_template.render(context, autoescape=True)
    doc_template.save(output_filename)
    return output_filename


def fill_template(template_filename, output_filename, volet2mesures):
    ordered_mesures = [mesure for mesures in volet2mesures.values() for mesure in mesures]
    ordered_volets = list(volet2mesures.keys())
    
    context = {encode_name(volet): [{'text': '', 'image': ''}] for volet in ordered_volets}
    dep_name = output_filename.split('_')[-1].split('.docx')[0]
    doc = DocxTemplate(template_filename)
    doc.render(context, autoescape=True)
    doc.save(output_filename)
    return output_filename


def map_templates_to_modified_reports(templates, modified_docx):
    mapping = {filename:None for filename in templates}
    
    # Faire correspondre le nom des départements encodés vers le bon template
    encoded_dep_name2template = {}
    for filename in mapping:
        raw_dep_name = filename.split('_')[-1].split('.')[0]
        encoded_dep_name = encode_name(raw_dep_name)
        encoded_dep_name2template[encoded_dep_name] = filename
    #assert len(encoded_dep_name2template) == 109
    
    # Faire correspondre le nom du département
    duplicated_dep = []
    for modified in modified_docx:
        content = docx2python(modified)
        expr_with_dep_name = content.body[0][0][0][7]
        print(f"Extrait de {modified} : ", expr_with_dep_name)
        expr_with_dep_name.split(':')
        dep_name = expr_with_dep_name.split(':')[-1].strip()
        clean_dep_name = encode_name(dep_name)
        target_template = encoded_dep_name2template[clean_dep_name]
        if mapping[target_template] is None:
            mapping[target_template] = modified
        else:
            duplicated_dep.append(dep_name)
            print(f"!!! {target_template} is not None -> probably duplicated \n----See {modified}")
            
    print("Fiches dupliquées (à retirer manuellement puis relancer le script) :\n", duplicated_dep)
    print(f"{len(mapping)} hits")
    return mapping


# Les nouveaux documents contiennent le texte transposé et sous le même nom que leur template mais
# dans un dossier différent


def transpose_modification_to_new_reports(template2modified_docx):
    # Transpose le texte ajouté aux documents sur le template associé. 
    # La correspondance se fait à partir du mapping (dictionnaire template -> doc modifié)
    hit, unhit = 0, 0
    for template_path, modified_docx_path in template2modified_docx.items():
        output_basename = template_path.split(os.sep)[-1]
        output_path = os.path.join(transposed_docx_dir, output_basename)

        output_name = fill_template(template_path, os.path.join(os.getcwd(), 'Fiche_Avant_Osmose', output_basename), volet2mesures)
        
        if modified_docx_path is None:
            unhit += 1
            print(f'Pas de transposition pour {template_path}')
            output_name = fill_template(template_path, output_path, volet2mesures)
        
        else:
            print(f'Transpose {template_path} vers {output_path}')
            try:
                output_name = transpose_comments(template_path, modified_docx_path, output_path, volet2mesures)
                hit += 1
            except:
                print(f"** Transposition impossible. Génération d'une fiche vide dans {template_path}**")
                output_name = fill_template(template_path, output_path, volet2mesures)
                unhit += 1
                
    print(f"Hit : {hit} | Unhit : {unhit}")


if __name__ == "__main__":
    main_transpose_comments()