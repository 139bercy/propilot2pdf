[![CircleCI](https://circleci.com/gh/139bercy/propilot2pdf.svg?style=svg)](https://circleci.com/gh/139bercy/propilot2pdf)

# Bienvenue sur ProPilot2PDF

<a href="reports/archive.zip">Télécharger toutes les fiches</a></br>
<a href="reports/Suivi_territorial_plan_relance_Ain.pdf">Télécharger fiche Ain</a>


# Comment générer des fiches ?

## Générer des fiches reprenant le commentaires des précédentes versions

1. Lancer le notebook ```chargement_propilot.ipynb``` pour obtenir le fichier ```pp_dep.csv```.
2. Lancer le notebook ```build_reports.ipynb``` pour générer les documents Word (.docx) dans un dossier ```reports_words```. Ces fiches contiennent des inscriptions ```Jinja``` dans les zones de commentaires qui seront remplacées par les commentaires récupérées depuis la précédente version des fiches. 
3. Télécharger toutes les fiches d'Osmose dans un dossier ```modified_reports``` (créer après avoir lancé ```transpose_comments.ipynb``` s'il n'existe pas) à la racine du projet.
4. Lancer le notebook ```transpose_comments.ipynb``` pour une transposition des commentaires de ```modified_reports``` -> ```reports_word/transposed_reports``` qui réutilise les nouvelles fiches-templates générées dans ```reports_words```.
5. Il est possible de déposer les fiches contenues dans ```reports_word/transposed_reports``` sur Osmose.


## Obtenir le format final et immuable des fiches

6. Télécharger les fiches modifiées dans le dossier ```modified_reports```, ce qui remplace les anciennes versions.
7. Lancer le notebook ```docx2pdf.ipynb``` pour convertir les fiches de ```modified_reports``` en pdf dans un dossier ```reports_pdf``` automatiquement créer (si ce n'était pas le cas).
8. Le format final est contenu dans le dossier ```reports_pdf```.
