# Prospecteur Foncier V4 — Maisons avec grand terrain

Nouvelle règle métier :

- les terrains nus restent éligibles ;
- dès qu'une parcelle est bâtie, elle doit être qualifiée **Résidentiel individuel** par la BDNB ;
- toute parcelle comportant du **Résidentiel collectif** est exclue ;
- les bâtiments tertiaires / indifférenciés sont également exclus des cibles bâties ;
- seuil réglable de **terrain libre minimum** autour de la maison (70 % par défaut) ;
- seuil réglable de **nombre maximum de logements existants** (2 par défaut).

## Données utilisées

- Cadastre Etalab : parcelles + empreinte des bâtiments ;
- Géoportail de l'Urbanisme / API Carto IGN : PLU / PLUi ;
- BDNB : usage principal du bâtiment, nombre de logements, niveaux et adresse principale ;
- Géoplateforme : géocodage inverse si l'adresse BDNB n'est pas disponible.

## Déploiement

Sur GitHub, remplacez `app.py` par celui de cette V4.
`requirements.txt` est identique à la V3 et peut être conservé.
