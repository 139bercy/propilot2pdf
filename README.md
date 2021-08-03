[![CircleCI](https://circleci.com/gh/139bercy/propilot2pdf.svg?style=svg)](https://circleci.com/gh/139bercy/propilot2pdf)

# Bienvenue sur ProPilot2PDF

<a href="reports/archive.zip">Télécharger toutes les fiches</a></br>
<a href="reports/Suivi_territorial_plan_relance_Ain.pdf">Télécharger fiche Ain</a>


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
2. Lancer le fichier ```main_create_parlementary_file.py```
3. Les fiches seront générées dans le dossier Fiche_Avant_Osmose (pour le format .docx) et dans le dossier Fiche_Avant_Osmose_pdf (pour le format pdf)
4. Il est possible de déposer les fiches contenues dans ```reports_word/transposed_reports``` sur Osmose.

## Après le passage Osmose

Dans le cas où vous disposez juste de fiches avec des commentaires, et que vous souhaitez les convertir en .pdf

1. Déposer le format modifiable (.docx / .odt) dans un dossier ```modified_reports```
2. Lancer le fichier ```main_convert_parlementary_file_with_new_comment.py```
3. Le format final est contenu dans le dossier ```reports_pdf```

