# Prospecteur Foncier V3.7.3 — filtre zonage PTZ

Cette version part de la V3.7.2 stable et ajoute un filtre **Zonage PTZ**
dans la barre latérale, entre Département et Commune.

## Filtres disponibles

- Tous
- A
- B1
- B2
- C

Le choix **A** inclut également **A bis**, conformément à l'article D304-1
du Code de la construction et de l'habitation.

## Source

Le zonage est chargé depuis le dataset officiel data.gouv.fr :
« Liste des communes selon le zonage ABC », produit par le ministère chargé du Logement.

Le code tente de découvrir automatiquement la ressource CSV la plus récente.
En secours, il utilise la liste nationale en vigueur au 26 juin 2026,
issue de la révision par arrêté du 23 juin 2026.

## Fonctionnement

Ordre des sélections :

1. Périmètre
2. Département
3. Zonage PTZ
4. Commune

La liste des communes est filtrée immédiatement après le choix de la zone PTZ.
La zone PTZ de la commune sélectionnée est aussi affichée dans l'interface.

## Déploiement

Remplacer uniquement `app.py` sur GitHub.

Le `requirements.txt` de la V3.7.2 reste inchangé.
