# Prospecteur Foncier V3.2 — Filtre Habitat strict

Cette version repart de la V3.1 et renforce fortement l'analyse du PLU/PLUi.

## Nouveau filtre Habitat

Par défaut, l'application conserve uniquement les parcelles pour lesquelles la destination
**Habitation / Logement** est suffisamment établie.

Ordre de preuve :

1. Standards CNIG récents :
   - `DESTOUI` contenant 20 ou 21 → habitat autorisé ;
   - `DESTCDT` contenant 20 ou 21 → habitat autorisé sous conditions ;
   - `DESTNON` contenant 20 ou 21 → habitat exclu.
2. Anciens PLU :
   - `DESTDOMI=01` → habitat ;
   - `DESTDOMI=03` → mixte habitat / activité.
3. Libellé / nom long de zone :
   - habitat, habitation, résidentiel, logement, etc.
4. Zone ambiguë sans preuve explicite :
   - exclue en mode strict afin de réduire les faux positifs.

## Zonage

- U : analysée ;
- AUc : analysée si l'option U + AUc est activée ;
- AUs : toujours exclue (zone à urbaniser bloquée) ;
- A : exclue ;
- N : exclue.

## Autres règles conservées

- calcul SDP / SHAB / logements ;
- exclusion des parcelles avec logement collectif existant via la BDNB ;
- terrain nu ou bâti ;
- carte cadastrale réelle ;
- adresses et lettres de prospection.

## Déploiement

Remplacer uniquement `app.py` sur GitHub. Le `requirements.txt` actuel reste compatible.
