# Bienvenue sur ProPilot2PDF

## Archive complète
<a href="reports/archive.zip">Télécharger toutes les fiches en format zip</a>

## Fiches Départementales

<nav>
    <ul>
        {% for image in site.static_files %}
            {% if image.path contains 'reports/' %}
                <li><a href="reports/{{ image.name }}">Télécharger {{ image.name }}<a/>
            {% endif %}
        {% endfor %}
