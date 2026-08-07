# Prospecteur Foncier V3.6 — moteur PLU hybride

Cette version vise à se rapprocher d'une préfaisabilité de type outil métier.

## Pourquoi la V3.5 était limitée

API Carto restitue les géométries et les attributs principaux des prescriptions,
mais sa documentation utilisateur ne liste pas les attributs complémentaires CNIG
`LIB_ATTRn / LIB_VALn` où peuvent se trouver `COEF_EMPRISE_SOL_MAX`,
`HAUTEUR_METRES_MAX` ou `HAUTEUR_RPLUS_ETAGES`.

Le moteur V3.5 pouvait donc connaître le périmètre d'une prescription 38/39
sans disposer de sa valeur chiffrée.

## Stratégie V3.6

1. Identifier la zone PLU/PLUi de la parcelle.
2. Lire une seule fois le règlement écrit lié à cette zone via `URLFIC`.
3. Extraire une règle de base :
   - emprise au sol maximale ;
   - R+N / nombre de niveaux / hauteur maximale.
4. Charger les prescriptions graphiques :
   - 38-02 : emprise maximale ;
   - 39-02 : hauteur maximale.
5. Si une prescription graphique n'a pas de valeur dans l'API, lire uniquement
   le document pointé par son `URLFIC`.
6. Découper géométriquement la parcelle en sous-surfaces :
   la règle graphique locale remplace la règle de zone uniquement là où elle s'applique.
7. Additionner la surface brute de chaque sous-surface.
8. Ne calculer les logements que si emprise + hauteur couvrent au moins 98 % de la parcelle.

## Formule

Surface brute = somme(surface de chaque sous-zone × emprise applicable × niveaux applicables)

Puis :
- SDP = surface brute × ratio SDP
- SHAB = SDP × ratio SHAB
- logements = SHAB / ratio SHAB par logement

## Limites

Même ce moteur ne remplace pas une étude de faisabilité complète.
Les retraits, bandes de constructibilité, pleine terre, stationnement,
OAP, risques et servitudes peuvent encore réduire le potentiel.

## Déploiement

Remplacer :
- `app.py`
- `requirements.txt`

Nouvelle dépendance : `pypdf`.
