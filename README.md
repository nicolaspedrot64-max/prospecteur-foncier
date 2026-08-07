# Prospecteur Foncier V3.7 — Préfaisabilité PLU renforcée

Cette version ajoute une vraie **base de règles PLU pré-interprétées par zone**.

## 1. Archive CNIG complète

Le logiciel utilise désormais le service officiel :

`/api/document/download-by-partition/<partition>`

pour récupérer l'archive CNIG du PLU/PLUi.

L'intérêt est important : l'API Carto fournit les objets graphiques, mais l'archive
peut aussi contenir les attributs supplémentaires `LIB_ATTR / LIB_VAL` publiés par
la collectivité.

Ces attributs sont notamment utilisés par le standard CNIG pour :
- l'emprise maximale (38-02 / COEF_EMPRISE_SOL_MAX) ;
- la hauteur (39-02 / HAUTEUR_METRES_MAX ou HAUTEUR_RPLUS_ETAGES) ;
- les reculs d'implantation (15 / VALEUR DE RECUL) ;
- le coefficient de biotope (42).

## 2. Base de règles par zone

Une seule lecture du règlement est effectuée par zone pour pré-interpréter :
- emprise maximale ;
- hauteur / niveaux ;
- recul par rapport aux voies ;
- recul aux limites latérales ;
- recul en fond de parcelle ;
- pleine terre / espaces verts ;
- stationnement par logement.

La base obtenue est affichée dans l'application et exportable en CSV.

## 3. Prescriptions graphiques supplémentaires

Le moteur analyse :
- 02 : limitations / interdictions de constructibilité ;
- 05 : emplacements réservés ;
- 15 : règles d'implantation / reculs ;
- 18 : OAP ;
- 38 : emprise ;
- 39 : hauteur ;
- 42 : biotope.

## 4. Capacité corrigée

Le calcul initial emprise × niveaux est ensuite corrigé :
- par une enveloppe prudente liée aux reculs ;
- par la pleine terre / le biotope ;
- par une éventuelle interdiction graphique.

Le tableau distingue :
- logements gabarit ;
- logements corrigés ;
- logements retenus.

## Prudence

Le mode « reculs prudents » applique le plus grand recul trouvé à tout le contour
de la parcelle faute d'identifier encore parfaitement la façade sur rue.
Il est donc volontairement conservateur.

## Déploiement

Remplacer :
- `app.py`
- `requirements.txt`

Nouvelle dépendance : `pyshp`.
