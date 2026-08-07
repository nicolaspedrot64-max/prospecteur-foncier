# Prospecteur Foncier V3.2.2 — Habitat hors secteurs économiques

Cette version conserve le moteur Habitat de la V3.2.1 et ajoute une exclusion prioritaire
des secteurs à vocation d'activités économiques.

## Exclusion économique

Une zone est rejetée avant toute autre analyse lorsque l'une des preuves suivantes est présente :

- ancien standard CNIG : `DESTDOMI=02` (activité) ;
- standard CNIG récent :
  - `FORMDOMI=0200` activité ;
  - `FORMDOMI=0201` activité industrielle / logistique / commerciale ;
  - `FORMDOMI=0202` activité commerces ;
  - `FORMDOMI=0203` activité bureaux ;
- libellé explicite indiquant notamment zone/secteur/parc d'activités, vocation économique,
  industrie, artisanat, commerce, logistique, tertiaire ou pôle économique.

Les secteurs mixtes habitat/activité (`DESTDOMI=03` / `FORMDOMI=0300`) ne sont pas exclus
par ce filtre économique : ils restent soumis au filtre Habitat.

## Autres règles conservées

- filtre Habitat strict ;
- calcul SDP / SHAB / logements ;
- exclusion du résidentiel collectif existant via la BDNB ;
- cadastre réel, carte, adresses et lettres.

## Déploiement

Remplacer uniquement `app.py` sur GitHub.
