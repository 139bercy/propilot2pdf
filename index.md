# Bienvenue sur ProPilot2PDF

## Archive complète
<a href="reports/archive.zip">Télécharger toutes les fiches</a>

## Fiches Départementales
<a href="reports/Suivi_territorial_plan_relance_Ain.pdf">Télécharger fiche Ain</a>

{% for image in site.static_files %}
    {% if image.path contains 'reports/' %}
        <img src="{{ site.baseurl }}{{ image.path }}" alt="{{ image }}" />
    {% endif %}
{% endfor %}
