# Prospecteur Foncier V3.4 — Gabarit PLU

Le logiciel extrait maintenant automatiquement dans le règlement écrit :
- l'emprise au sol maximale ;
- le nombre de niveaux (R+N / niveaux explicites) ;
- à défaut, la hauteur maximale convertie en niveaux.

Calcul :
- Emprise constructible = Surface terrain × Emprise PLU
- Surface brute = Emprise constructible × Nombre de niveaux
- SDP = Surface brute × ratio SDP
- SHAB = SDP × ratio SHAB
- Logements = SHAB / SHAB moyenne par logement

Si l'emprise ou les niveaux ne sont pas trouvés, le logiciel n'invente pas de valeur :
la parcelle n'entre pas dans le calcul automatique.

Un audit affiche la règle trouvée, son niveau de confiance, l'extrait du règlement et le lien PDF.

Limites restantes : retraits, bandes de constructibilité, pleine terre, stationnement,
OAP, prescriptions graphiques et servitudes peuvent réduire la capacité réelle.

Déploiement GitHub :
- remplacer `app.py`
- remplacer `requirements.txt` (ajout de `pypdf`)
