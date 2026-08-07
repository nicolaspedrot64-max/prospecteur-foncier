# Prospecteur Foncier V3.7.4 — nouvelle lettre + dossier PDF

## Nouvelle lettre SAGEC

Le modèle `lettre_sagec_modele.docx` a été remplacé par le fichier
`Lettre_SAGEC_amelioree(1).docx` fourni.

Le logiciel remplit automatiquement :
- l'adresse de la propriété ;
- la commune ;
- la date ;
- l'objet ;
- la section et le numéro cadastral ;
- le signataire ;
- la fonction ;
- l'e-mail.

Le texte et la présentation du nouveau modèle sont conservés.

## Dossier PDF des parcelles sélectionnées

Un nouveau bouton permet de télécharger un PDF professionnel contenant :
- une page de synthèse de la sélection ;
- le nombre de parcelles ;
- la surface cadastrale cumulée ;
- la SDP estimée cumulée ;
- le nombre de logements estimés ;
- un tableau récapitulatif ;
- une fiche par parcelle ;
- un schéma vectoriel de la parcelle cadastrale ;
- adresse et référence cadastrale ;
- propriétaire société / commune et SIREN lorsqu'ils sont disponibles ;
- zone PLU et zone PTZ ;
- emprise, hauteur/niveaux, surface brute, SDP, SHAB et logements ;
- prescriptions utilisées et niveau de confiance.

## Déploiement

Remplacer sur GitHub :
1. `app.py`
2. `requirements.txt`
3. `lettre_sagec_modele.docx`

La nouvelle dépendance est `reportlab`.
