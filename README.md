# Prospecteur Foncier V3.1 — exclusion du collectif existant

Cette version repart de la V3 (ratios SDP / SHAB) et annule les logiques des V4 et V5.

## Règle métier

Le logiciel ne filtre plus les maisons individuelles, villas, grands bâtiments ou parcelles selon leur emprise.

Il élimine uniquement les parcelles pour lesquelles la BDNB identifie déjà un usage
**« Résidentiel collectif »**.

Donc :
- terrain nu : conservé ;
- maison individuelle : conservée ;
- grande villa occupant presque toute la parcelle : conservée ;
- bâtiment non résidentiel : conservé par ce filtre (les autres filtres urbanistiques continuent de s'appliquer) ;
- résidence / immeuble de logements collectifs existant : exclu.

## Déploiement

Remplacer uniquement `app.py` sur GitHub.
Le `requirements.txt` actuel de la V3 est compatible.
