# Prospecteur Foncier V3.7.1 — correctif stabilité

Cette version corrige le comportement de la V3.7 qui pouvait donner l'impression
que l'application ne fonctionnait plus.

## Changements

- L'analyse de l'archive CNIG complète est désormais **optionnelle et décochée par défaut**.
- Le téléchargement d'un gros PLUi ne bloque plus le parcours normal.
- Le filtre nombre de logements repose d'abord sur le gabarit emprise/hauteur.
- Les reculs / pleine terre / contraintes avancées restent des informations de préfaisabilité
  et ne suppriment plus silencieusement les parcelles.
- Le mode de recul prudent est désactivé par défaut.
- Les parcelles sans gabarit calculé restent **visibles et sélectionnables**.
- Si des parcelles calculées existent, un second éditeur permet aussi de sélectionner
  les parcelles « à vérifier ».
- Une erreur de l'archive CNIG devient un message d'information ; l'analyse continue.

## Pourquoi

Depuis 2026, certains gros documents PLUi du GPU sont lourds et leurs archives ont été
réorganisées / scindées. Une application Streamlit ne doit pas dépendre du téléchargement
de l'archive complète pour simplement afficher les parcelles.

## Déploiement

Remplacer uniquement `app.py`.

Le `requirements.txt` de la V3.7 peut être conservé.
