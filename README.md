# Prospecteur Foncier V3.2.1 — correctif KeyError

Correctif de la V3.2.

## Problème corrigé

Lorsqu'une commune ne renvoyait aucune parcelle après le filtre Habitat strict,
`analyse_commune()` retournait un DataFrame vide sans colonnes. L'interface essayait ensuite
d'accéder à `results["logements_estimes"]`, ce qui provoquait un `KeyError`.

## Nouveau comportement

- aucune parcelle habitat trouvée -> message explicatif, sans plantage ;
- résultat vide après les autres critères -> message explicatif, sans plantage ;
- le DataFrame conserve désormais toujours le schéma attendu, même s'il est vide.

## Déploiement

Remplacer uniquement `app.py` sur GitHub. Le reste du projet ne change pas.
