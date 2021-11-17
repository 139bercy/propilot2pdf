[![CircleCI](https://circleci.com/gh/139bercy/propilot2pdf.svg?style=svg)](https://circleci.com/gh/139bercy/propilot2pdf)

# Bienvenue sur le projet SGPR: Suivi des mesures du Plan de Relance


# Le [plan de relance](https://www.economie.gouv.fr/plan-de-relance): Qu'est-ce que c'est ?

Pour faire face à l’épidémie du Coronavirus Covid-19, le Gouvernement a mis en place dès le début de la crise, des mesures inédites de soutien aux entreprises et aux salariés, qui continuent aujourd'hui d'être mobilisables.

Afin de redresser rapidement et durablement l’économie française, un plan de relance exceptionnel de 100 milliards d’euros est déployé par le Gouvernement autour de 3 volets principaux : l'écologie, la compétitivité et la cohésion. Ce plan de relance, qui représente la feuille de route pour la refondation économique, sociale et écologique du pays, propose des mesures concrètes et à destination de tous. Que vous soyez un particulier, une entreprise, une collectivité ou bien une administration, retrouvez l’ensemble des mesures dont vous pouvez bénéficier dans le cadre du plan de relance !

Ce projet fait partie des outils de suivis mis en place par le Secrétariat Général au Plan de Relance (SGPR).  


# Dépendances

Libreoffice (si pas installé par défaut)

```
    sudo add-apt-repository ppa:libreoffice/ppa
    sudo apt update
    sudo apt install libreoffice
```


# Comment générer des fiches ?

## Obtenir les données

1. Demander au BercyHub la clé et l'URL du dépot de données.
2. Créer un dossier data/ puis s'y rendre ```mkdir data; cd data```
3. Lancer la commande ```sftp -P 2022 -i ../key url_sftp.com```
4. Obtenir les fichiers csv ```get *.csv```

## Générer des fiches reprenant le commentaires des précédentes versions. 

1. Si vous disposez d'anciennes fiches avec des commentaires, déposez le format modifiable (.docx / .odt) dans un dossier ```modified_reports```
2. Lancer le script ```main_create_parlementary_file.py```
3. Les fiches seront générées dans le dossier Fiche_Avant_Osmose (pour le format .docx) et dans le dossier Fiche_Avant_Osmose_pdf (pour le format pdf)
4. Il est possible de déposer les fiches contenues de ```reports_word/transposed_reports``` sur Osmose.

## Après le passage Osmose

Dans le cas où vous disposez juste de fiches avec des commentaires, et que vous souhaitez les convertir en .pdf

1. Déposer le format modifiable (.docx / .odt) dans un dossier ```modified_reports```
2. Lancer le script ```main_convert_parlementary_file_with_new_comment.py```
3. Le format final est contenu dans le dossier ```reports_pdf```

## Opérations effetuées sur les données

### Création d'un CSV intermédiaire contenant les données départementales: pp_dep

1. Création d'un dataframe unique avec les 6 fichiers suivants:
    - fact_financials.csv
    - dim_tree_nodes.csv
    - dim_effects.csv
    - dim_states.csv
    - dim_period.csv
    - dim_structures.csv
2. Renommage de la colonne ```period_month_year``` en ```Date``` et ```financials_cumulated_amount``` en ```valeur```
3. Split de la colonne ```effect_id``` qui contient le nom de l'indicateur en deux colonnes:
    - ```short_indic```: Contient le nom de l'indicateur 
    - ```indic_id```: Contient la clef de l'indicateur sous la forme d'un quadrigramme.
4. Renommage de ```effect_id``` en ```indicateur``` et formatage de la colonne ```short_indic```
5. Récupération des données utiles pour les fiches 
    - structure_name: ```Département```
    - period_month_tri: Toutes les valeurs sauf ```Total``` et ```Y```
    - state_id: ```Valeur Actuelle```
    - valeur: Toutes les valeurs non nulles
6. Suppression des lignes avec une date plus ancienne que celle du jour
7. Création des colonnes ```departement``` et ```mesure``` à partir de la colonne ```tree_node_name``` et formatage de ces dernières
8. Traduction en français des mois dans la colonne ```Date```
9. Renommage de certains indicateurs dans la colonne ```short_indic```
10. Renommage de certaines mesures dans la colonne ```short_mesure```
11. Ajout des colonnes:
    - ```libelle```: Libellé du département
    - ```reg```: Code de la Région
    - ```region```: Libellé de la région
12. Export de pp_dep

### Création des fiches

1. Récupération des mesures et indicateurs devant figurer sur une fiche
2. Calcul des valeurs Régionales et Nationales pour chaque indicateurs
3. Création des dataframes pp_reg et pp_nat qui contiennent les valeurs des indicateurs par région et le total
4. Calcul des poids valeurs Départementales/Régionales et Régionales/Nationales
5. Formatage de la colonne valeur: 
    - Passage en str
    - pour pp_dep et pp_reg, ajout d'un caractère espace et de la valeur du poids entre parenthèse 
6. Création d'un dicionnaire qui contiendra les données des tableaux à mettre dans les fiches
7. Création itérative des fiches, à partir des templates présents dans le dossier ```template```:
    - Création de la page de garde
    - Création des pages volets 
    - Création d'un document par mesure à insérer dans les fiches
    - Fusion itérative des pages dans l'ordre souhaité par le SGPR
    - Export dans le dossier ```reports_word```
8. Récupération des commentaires des anciennes fiches, contenues dans le dossier ```modified_reports``` si il existe.
    - Export dans le dossier ```reports_word/transposed_reports```
9. Conversion des fiches:
    - Copie des fiches docx dans le dossier ```reports_before_new_comment```
    - Conversion en PDF de ces fiches dans le dossier ```reports_before_new_comment_pdf```
10. Création d'un zip archive stocké dans ```archive/Mois_Annee```

### Conversion une fois les fiches commentées

Pour cette partie, il suffit d'utiliser le script ```main_convert_parlementary_file_with_new_comment.py```

1. Déposer les nouvelles fiches dans le dossier ```modified_reports``` (le créer si nécessaire)
2. Si plusieurs fiches pour le meme document, alors le script va supprimer les fiches les plus anciennes (date de modification)
2. Renommage des fiches, conversion en pdf et stockage dans le dossier reports_pdf
3. Création d'un zip archive stocké dans ```archive/Mois_Annee```

#### Dans le cas d'une création de fiche, dans le script build_reports, ligne 527 il y a une valeur à changer en fonction du mois souhaité