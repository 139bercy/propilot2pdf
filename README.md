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



## Générer des fiches reprenant le commentaires des précédentes versions

1. Lancer le notebook ```chargement_propilot.ipynb``` pour obtenir le fichier ```pp_dep.csv```.
2. Lancer le notebook ```build_reports.ipynb``` pour générer les documents Word (.docx) dans un dossier ```reports_words```. Ces fiches contiennent des inscriptions ```Jinja``` dans les zones de commentaires qui seront remplacées par les commentaires récupérées depuis la précédente version des fiches. 
3. Télécharger toutes les fiches d'Osmose dans un dossier ```modified_reports``` (créé après avoir lancé ```transpose_comments.ipynb``` s'il n'existe pas) à la racine du projet.
4. Lancer le notebook ```transpose_comments.ipynb``` pour une transposition des commentaires de ```modified_reports``` -> ```reports_word/transposed_reports``` qui réutilise les nouvelles fiches-templates générées dans ```reports_words```.
5. Il est possible de déposer les fiches contenues dans ```reports_word/transposed_reports``` sur Osmose.


## Obtenir le format final et immuable des fiches (conversion en PDF)

La conversion n'accepte que le format ```.docx```.

6. Télécharger les fiches (modifiées ou non) dans le dossier ```modified_reports```, ce qui remplace les anciennes versions. Les fiches n'auront pas de commentaire ou contiendront les commentaires d'une version autérieure si elles n'ont pas été modifiées le mois courant.
7. Lancer le notebook ```docx2pdf.ipynb``` pour convertir les fiches de ```modified_reports``` en pdf dans un dossier ```reports_pdf``` automatiquement créé (si ce n'était pas le cas). Il est possible que certaines fiches soient dupliquées, un test du notebook l'indiquera. Il faudra alors retirer manuellement les fiches en trop. Enfin, relancer la cellule et vérifier que ce même test passe.
8. Le format final est contenu dans le dossier ```reports_pdf```.
