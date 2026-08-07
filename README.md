# Prospecteur Foncier V3.8 — analyse PLU/PLUi précise par commune

## Objectif

La recherche est désormais recalculée spécifiquement pour la commune sélectionnée.

## Méthode

1. Charge le dernier cadastre disponible via le flux Etalab `latest`.
2. Résout le PLU/PLUi en vigueur sur la commune.
3. Charge uniquement le `ZONE_URBA` et les prescriptions graphiques de cette commune.
4. Intersecte **géométriquement chaque parcelle avec toutes les zones PLU/PLUi**.
5. Élimine avant tout calcul :
   - zones A ;
   - zones N ;
   - zones AUs ;
   - autres zones non constructibles ;
   - zones à vocation économique / industrielle / artisanale / logistique ;
   - secteurs graphiques CNIG `02-01` avec interdiction explicite de constructibilité.
6. Conserve la géométrie exacte de la partie de parcelle située en zone constructible.
7. Pour chaque sous-zone constructible :
   - cherche une prescription CNIG 38-02 d'emprise maximale ;
   - cherche une prescription CNIG 39-02 de hauteur maximale ;
   - si nécessaire, utilise le règlement PDF **lié à cette zone graphique** pour récupérer uniquement ces deux valeurs.
8. Une prescription graphique locale ne s'applique que sur son périmètre réel.
9. La surface brute est la somme des sous-surfaces :
   `surface × emprise applicable × niveaux applicables`.
10. La capacité automatique n'est calculée que si emprise + hauteur couvrent au moins 95 % de la partie constructible de la parcelle.

## Ce que cette version ne calcule pas

Volontairement :
- retraits ;
- stationnement ;
- pleine terre ;
- OAP ;
- servitudes ;
- risques.

La V3.8 se concentre uniquement sur **emprise au sol + hauteur**, conformément au besoin.

## Déploiement

Remplacer :
- `app.py`
- `requirements.txt`

La lettre SAGEC de la V3.7.4 peut être conservée telle quelle.
