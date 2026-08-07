# Prospecteur Foncier V3.5 — corrélation cadastre / règlement graphique

Cette version repart de la V3.3.1 stable et abandonne l'idée de lire d'abord
le règlement PDF pour calculer le gabarit.

## Source principale : prescriptions graphiques GPU

Le moteur croise chaque géométrie cadastrale avec :
- `prescription-surf`
- `prescription-lin`
- `prescription-pct`

Il recherche en priorité :
- CNIG `38-02` : emprise au sol maximale ;
- CNIG `39-02` : hauteur maximale ;
- variantes de hauteur localisées lorsqu'elles existent.

## Valeurs

Le moteur lit :
1. les couples `LIB_ATTRn / LIB_VALn` s'ils sont exposés ;
2. à défaut, les champs `TXT`, `LIBELLE`, `NATURE`.

Il reconnaît notamment :
- pourcentages d'emprise ;
- coefficients CES décimaux ;
- R+1, R+2, etc. ;
- nombre de niveaux ;
- hauteurs en mètres.

## Calcul

Si emprise + hauteur/niveaux sont suffisamment couverts graphiquement :
- Emprise constructible = surface cadastrale × emprise PLU
- Surface brute = emprise constructible × niveaux
- SDP = surface brute × ratio SDP
- SHAB = SDP × ratio SHAB
- Logements = SHAB / ratio SHAB par logement

Si le graphique n'est pas assez complet, la parcelle est conservée dans
« Gabarit PLU à vérifier » : elle n'est pas supprimée et aucune valeur n'est inventée.

## Limite nationale importante

Tous les PLU/PLUi ne spatialisent pas les règles d'emprise et de hauteur sous forme
de prescriptions graphiques 38/39. La couche de zonage est obligatoire, mais la
présence des autres couches dépend du contenu du document d'urbanisme.

Pour couvrir 100 % des cas, la prochaine brique devra utiliser en complément :
- le règlement structuré CNIG/SRU lorsqu'il est publié ;
- sinon le règlement écrit de la zone comme fallback ciblé.

## Déploiement

Remplacer uniquement `app.py` sur GitHub.
Le `requirements.txt` de la V3.3.1 reste compatible.
