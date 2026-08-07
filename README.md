# Prospecteur Foncier V5 — Maison individuelle sans contrainte d'emprise

Correction métier par rapport à la V4 :

- un terrain nu reste éligible ;
- une parcelle déjà bâtie reste éligible si la BDNB la qualifie **Résidentiel individuel** ;
- une grande maison ou villa peut donc être sélectionnée même si elle occupe presque toute la parcelle ;
- une résidence / un immeuble en **Résidentiel collectif** reste exclu ;
- le nombre maximum de logements existants est réglable, avec **1 logement par défaut** ;
- la part de terrain libre est toujours calculée et affichée, mais **n'est plus un filtre d'exclusion** ;
- une option permet seulement de **prioriser** les maisons ayant plus de terrain libre dans le score.

## Déploiement

Sur GitHub, remplacez uniquement `app.py` par celui de cette V5.
Le `requirements.txt` de la V4/V3 peut être conservé.
