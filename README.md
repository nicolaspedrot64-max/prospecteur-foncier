# Prospecteur Foncier V3.3.1 — correctif API GPU / PLUi

Cette version corrige le plantage HTTP 500 observé sur Capbreton lors de l'appel :

`/api/gpu/document?geom=...`

## Pourquoi Capbreton posait problème

Capbreton est couvert par le PLUi de la Communauté de communes Maremne Adour Côte-Sud (MACS).
Un PLUi est stocké dans le Géoportail de l'Urbanisme sous une partition de type :

`DU_<SIREN EPCI>`

La V3.3 dépendait trop du endpoint `/document?geom`. Si celui-ci renvoyait une erreur 500,
l'analyse s'arrêtait.

## Correctif

La V3.3.1 :

- récupère aussi le `codeEpci` via geo.api.gouv.fr ;
- réessaie automatiquement les erreurs temporaires 500/502/503/504 ;
- si `/gpu/document?geom` échoue, teste directement :
  - `DU_<codeEpci>` pour les PLUi ;
  - `DU_<codeINSEE>` pour les PLU communaux ;
- vérifie la partition en interrogeant directement `zone-urba` ;
- filtre un PLUi sur le code INSEE de la commune pour éviter de charger toutes les communes de l'EPCI.

Pour Capbreton, le fallback attendu est le PLUi MACS.

## Déploiement

Remplacer uniquement `app.py` sur GitHub.
Le `requirements.txt` de la V3.3 peut être conservé.
