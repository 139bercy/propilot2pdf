# Bienvenue sur ProPilot2PDF

## Archive complète
<a href="reports/archive.zip">Télécharger toutes les fiches</a>

## Fiches Départementales
<a href="reports/Suivi_territorial_plan_relance_Ain.pdf">Télécharger fiche Ain</a>

<nav>
{% for image in site.static_files %}
    {% if image.path contains 'reports/' %}
        <a href="{{ image.path }}">Télécharger {{ image.name }}<a/>
    {% endif %}
{% endfor %}
</nav>
