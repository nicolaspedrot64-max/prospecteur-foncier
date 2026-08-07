# Prospecteur Foncier V3.3 — propriétaires personnes morales

## Changements

1. Le filtre visible « Conserver uniquement les parcelles destinées à l'habitat » est supprimé.
2. L'exclusion des secteurs à vocation d'activités économiques reste active.
3. Le logiciel tente maintenant d'identifier le propriétaire lorsqu'il s'agit d'une
   **personne morale** : société, SCI, collectivité, commune, etc.
4. Le tableau affiche :
   - nom / dénomination ;
   - type de propriétaire ;
   - forme juridique ;
   - SIREN lorsqu'il est fourni par la source.

## Source propriétaires

Le rapprochement utilise le jeu open data DGFiP « Fichiers des locaux et des parcelles
des personnes morales », via la version unifiée au format Parquet publiée sur data.gouv.fr.

Les propriétaires personnes physiques ne sont volontairement pas identifiés : si aucune
personne morale n'est trouvée, le champ reste vide.

## Important pour le déploiement

Cette version ajoute la dépendance `duckdb`.

Sur GitHub, remplacez donc :
- `app.py`
- `requirements.txt`

Puis faites Commit changes. Streamlit redéploiera automatiquement l'application.
