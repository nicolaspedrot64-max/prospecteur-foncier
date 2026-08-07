# Prospecteur Foncier V3.9 — sélection simple Habitat

Cette version revient volontairement à un moteur simple et stable.

## Sélection
Une parcelle est conservée uniquement si :
- elle est en zone U ;
- ou en zone AUc/AU si cette option est choisie ;
- le PLU/PLUi fournit une indication Habitat/Logement suffisamment claire ;
- la zone n'est pas à vocation économique ;
- la parcelle n'est pas déjà identifiée comme logement collectif existant.

## Identification Habitat
Le moteur utilise :
1. DESTOUI / DESTCDT / DESTNON ;
2. FORMDOMI : 01xx = habitat, 0300 = mixte habitat/activité ;
3. DESTDOMI : 01 = habitat, 03 = mixte habitat/activité ;
4. libellés habitat / résidentiel / logement ;
5. fallback limité aux codes urbains courants UA, UB, UC, UD, UH, UR, UP.

Les secteurs d'activités, industrie, commerce, logistique, artisanat, équipements,
tourisme/loisirs, agriculture et nature sont exclus.

## Estimation logements
- SDP = surface cadastrale × ratio SDP (80 % par défaut)
- SHAB = SDP × ratio SHAB (80 % par défaut)
- logements = SHAB / SHAB moyenne par logement (55 m² par défaut)

## Propriétaires
Affichage du propriétaire uniquement lorsqu'il s'agit d'une société,
d'une commune ou d'une autre personne morale disponible dans l'open data DGFiP.

## Courriers et PDF
Le modèle de lettre SAGEC et le dossier PDF de sélection sont conservés.
Le PDF affiche l'estimation simple plutôt qu'un gabarit PLU automatique.

## Déploiement
Remplacer uniquement `app.py`.
Le `requirements.txt` et la lettre de la V3.7.4 restent compatibles.
