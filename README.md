# Prospecteur Foncier — MVP

Prototype Streamlit pour :
- saisir des critères (ville, nombre de logements, PC, terrain nu/bâti, surface, score) ;
- éliminer automatiquement les parcelles marquées non constructibles ;
- afficher la liste des parcelles répondant aux critères ;
- sélectionner les parcelles à prospecter ;
- générer des courriers DOCX personnalisés à partir du modèle SAGEC fourni.

## Installation

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

## Données

Le fichier `data/parcelles_exemple.csv` sert de démonstration.

Colonnes attendues :
- ville
- adresse
- section
- numero
- surface_m2
- zone_plu
- constructible
- terrain_bati
- pc_obtenu
- logements_estimes
- score

## Important

Cette version est un MVP fonctionnel de l'interface et du moteur de filtrage.
Elle ne se connecte pas encore automatiquement au cadastre, au PLU/PLUi ni aux autorisations d'urbanisme.
Ces connexions seront la prochaine étape pour transformer le prototype en outil opérationnel.
