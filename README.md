# Application de calcul NUR

Application web locale pour calculer et visualiser le NUR des cellules, des
technologies, des sites, des régions et du réseau à partir des exports Excel
ORDC et NUR.

Le générateur PowerPoint trafic se trouve dans un dossier séparé :
`C:\Users\winne\Documents\Orange_Automatisation\TRAFFIC_PPT`.

## Installation

Dans PowerShell, depuis ce dossier :

```powershell
python -m pip install -r requirements.txt
```

## Lancement

```powershell
python -m streamlit run app.py
```

Le navigateur s'ouvre normalement sur `http://localhost:8501`.

## Tests

```powershell
python -m unittest discover -s tests -v
```

## Utilisation

1. Importer la base déjà filtrée. L'application utilise directement sa
   **deuxième feuille**, sans appliquer de filtre supplémentaire.
2. Importer le fichier NUR contenant les onglets `2G`, `3G` et `4G`.
3. Choisir la date de début et une période de 1, 7 ou 30 jours.
4. Consulter les tableaux et graphiques régionaux, technologiques et par site.
5. Télécharger le rapport Excel complet ou la présentation PowerPoint.

Les dates texte des exports NUR sont interprétées au format `MM/JJ/AAAA`.
Ainsi, `09/03/2026` correspond au 3 septembre 2026.

## Règles appliquées

La formule commune est :

```text
NUR = somme des indisponibilités
      / (nombre de cellules × 86 400 × nombre de jours)
      × 100 000
```

- 2G : `Site Name`, `Cell Name` et
  `R373:Cell Out-of-Service Duration(s)`.
- 3G : `NODEBNAME`, `Cell Name` et
  `VS.Cell.UnavailTime.Sys(s)`.
- 4G : `eNodeB Name`, `Cell Name` et
  `L.Cell.Unavail.Dur.Sys(s)`.
- Les suffixes `_GSM` et `_UMTS` sont retirés pour effectuer la
  correspondance avec `Site_Code`.
- La deuxième feuille de la base constitue directement le périmètre des sites
  admissibles. La première feuille n'est pas utilisée et aucun filtre
  `Type_transmission` n'est réappliqué.
- Les sites absents de la base et les lignes invalides apparaissent dans
  l'onglet `Exclusions` du rapport.

Le rapport téléchargé contient les onglets `Parametres`, `Cellules`,
`Sites_Technologies`, `Sites`, `Technologies`, `Regions`,
`Evolution_journaliere`, `Reseau` et `Exclusions`.

Les tableaux affichent l'indisponibilité en secondes, minutes et heures. Les
totaux représentent des **heures-cellules** : une heure d'indisponibilité sur
deux cellules correspond donc à deux heures-cellules.

L'interface propose une vue exécutive en grille, le NUR total de tout le
réseau, le NUR total 2G/3G/4G, des graphiques animés, une matrice
régions-technologies et des classements détaillés.

La présentation PowerPoint comprend :

- une synthèse exécutive avec les indicateurs clés ;
- l'indisponibilité par région commerciale (`Region_commerciale`, colonne K) ;
- la contribution de chaque technologie ;
- les sites prioritaires et leurs durées en heures ;
- la méthode de calcul et les contrôles de qualité des données.
