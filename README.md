# Prospecteur Foncier V3.7.2 — Mode stable

Cette version est une version de secours volontairement basée sur le dernier socle
qui fonctionnait correctement dans Streamlit : le moteur V3.5.

Elle conserve :
- cadastre réel ;
- PLU/PLUi ;
- exclusion des secteurs à vocation économique ;
- exclusion du résidentiel collectif existant ;
- propriétaires personnes morales / communes ;
- prescriptions graphiques CNIG 38-02 (emprise maximale) ;
- prescriptions graphiques CNIG 39-02 (hauteur maximale) ;
- calcul surface brute / SDP / SHAB / logements ;
- carte, sélection et courriers.

Elle retire temporairement :
- téléchargement de l'archive CNIG complète ;
- lecture directe des shapefiles de l'archive ;
- dépendances pypdf / pyshp introduites dans les versions avancées.

## Déploiement

IMPORTANT : remplacer les DEUX fichiers sur GitHub :
1. app.py
2. requirements.txt

Cela force Streamlit à reconstruire un environnement plus léger.
