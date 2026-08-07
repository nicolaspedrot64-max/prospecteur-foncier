# Prospecteur Foncier V2

Cette version remplace les données de démonstration Bayonne/Anglet par des données publiques réelles.

## Sources utilisées

- Communes : API Découpage administratif (`geo.api.gouv.fr`)
- Parcelles et bâtiments : Cadastre Etalab, dernier millésime
- PLU / PLUi : API Carto IGN, module Géoportail de l'Urbanisme
- Adresses : Service de géocodage de la Géoplateforme
- Courriers : modèle SAGEC `lettre_sagec_modele.docx`

## Mise à jour sur GitHub

Remplacer votre ancien `app.py` et `requirements.txt` par ceux de cette V2.
Ajouter `lettre_sagec_modele.docx` à la racine du dépôt si ce n'est pas déjà fait.

Streamlit redéploiera automatiquement l'application après le commit.

## Ce que la V2 fait

- choix dynamique d'une commune de Nouvelle-Aquitaine ou d'Occitanie ;
- téléchargement du vrai cadastre de la commune ;
- téléchargement des bâtiments ;
- récupération du document PLU/PLUi et de son zonage ;
- exclusion des zones A/N ;
- analyse des zones U et, au choix, AU ;
- détection terrain nu/bâti ;
- estimation de capacité logement ;
- carte réelle des parcelles ;
- sélection, recherche d'adresse et génération de lettres.

## Limite à connaître

Le nombre de logements est encore une estimation de présélection. Pour une faisabilité réglementaire
plus précise, il faut développer un moteur de lecture automatique des règlements écrits du PLU
(emprise, hauteur, retraits, stationnement, pleine terre, destinations, OAP, prescriptions, etc.).
