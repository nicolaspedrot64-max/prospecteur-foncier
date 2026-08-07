# Prospecteur Foncier V3.4.1 — correctif gabarit PLU

La V3.4 était trop stricte : toute parcelle dont l'emprise ou le nombre de niveaux
n'était pas extrait automatiquement disparaissait de l'affichage.

La V3.4.1 sépare désormais :

- **Parcelles dans la cible logements** : gabarit PLU extrait + nombre de logements calculé ;
- **Parcelles à vérifier** : parcelles conservées, mais dont l'emprise et/ou les niveaux
  n'ont pas encore été lus automatiquement.

Autre correction importante : le logiciel exploite désormais le fragment `#page=N`
que le standard CNIG peut placer dans `URLFIC` pour pointer directement vers la page
du règlement correspondant à la zone.

Aucune valeur de capacité n'est inventée lorsque le règlement n'est pas interprétable.

## Déploiement

Remplacer uniquement `app.py` sur GitHub.
Le `requirements.txt` de la V3.4 ne change pas.
