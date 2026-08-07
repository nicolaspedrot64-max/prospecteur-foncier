# Prospecteur Foncier V3 — ratios SDP / SHAB

Évolution de la V2 demandée :

- suppression du critère de surface minimale de terrain ;
- conservation du filtre par nombre de logements ;
- ajout de 3 ratios modifiables :
  - SDP / surface brute : 80 % par défaut ;
  - SHAB / SDP : 80 % par défaut ;
  - SHAB moyenne par logement : 55 m² par défaut ;
- calcul automatique :
  - `SDP = surface brute × ratio SDP`
  - `SHAB = SDP × ratio SHAB`
  - `Nombre de logements = partie entière(SHAB / SHAB par logement)`
- affichage de la surface brute, SDP estimée, SHAB estimée et nombre de logements pour chaque parcelle.

## Hypothèse provisoire importante

Dans cette V3, la surface cadastrale de la parcelle sert encore de base de **surface brute de présélection**.
La version suivante devra remplacer cette approximation par une vraie surface brute constructible issue du
gabarit PLU : emprise au sol, nombre de niveaux/hauteur, retraits, pleine terre, prescriptions, OAP, etc.
