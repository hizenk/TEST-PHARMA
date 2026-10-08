# Rush 2 : ventes de la pharmacie

Un classeur Excel propre (`Donnees_propres.xlsx`) avec un **tableau de bord
qui fonctionne entièrement dans Excel**, et deux scripts Python : l'un fabrique
le tableau de bord, l'autre va chercher les informations dans l'Excel au moment
où on les demande.

```
TEST-PHARMA/
├── Donnees_propres.xlsx       # données propres + onglet « Tableau de bord »
├── creer_tableau_de_bord.py   # (re)crée l'onglet Tableau de bord dans le classeur
├── interroger_excel.py        # interroge le classeur en live (fichier local ou dépôt git)
├── exports/                   # les 4 exports CSV bruts du client
└── requirements.txt
```

## Le tableau de bord (dans Excel, sans macro)

Ouvrez `Donnees_propres.xlsx` : le classeur s'ouvre sur l'onglet
**Tableau de bord**. Seules les cellules jaunes se modifient :

1. **Médicaments** : `Oui` ou `Non` devant chacun des 8 groupes ATC (liste déroulante).
2. **Période** : date de début et date de fin (limitées aux données disponibles).
3. **Affichage** : regrouper par `Jour`, `Semaine`, `Mois`, `Année`,
   `Jour de la semaine` ou `Heure`, et afficher la `Somme` ou la `Moyenne par jour`.

Tout se recalcule immédiatement :

- **4 indicateurs** : quantité vendue, moyenne par jour, meilleur jour (et sa
  quantité), nombre de jours de données ;
- **le détail par médicament** : total, moyenne par jour, part de la sélection ;
- **le tableau des résultats** (jusqu'à 400 lignes) avec une colonne par
  médicament, le total de la sélection en barres et le nombre de jours de
  chaque ligne (une période incomplète se voit tout de suite) ;
- **4 graphiques** : évolution de la sélection, profil selon le jour de la
  semaine, profil heure par heure, total par médicament.

Un message rouge prévient en cas de problème : date de fin avant la date de
début, aucun médicament inclus, ou plus de 400 lignes (choisir alors un
regroupement plus large).

Comment c'est construit :

- uniquement des formules Excel classiques (`SUMIFS`, `COUNTIFS`,
  `SUMPRODUCT`, `INDEX`/`MATCH`), sans macro ni fonction récente : cela marche
  dans Excel 365 comme dans les versions plus anciennes ;
- les formules lisent directement les tableaux `Pharma_Ventes_Daily` et
  `Pharma_Ventes_Hourly`. Si on ajoute des lignes à ces tableaux, le tableau de
  bord les prend en compte sans rien recopier ;
- l'onglet est protégé, sans mot de passe, pour éviter d'écraser une formule
  par erreur. Pour le modifier : *Révision > Ôter la protection de la feuille* ;
- les calculs intermédiaires sont dans l'onglet masqué `Calculs`.

## Installation

```bash
git clone https://github.com/hizenk/TEST-PHARMA.git
cd TEST-PHARMA
pip install -r requirements.txt
```

## Recréer le tableau de bord

```bash
python creer_tableau_de_bord.py
```

Le script remplace l'onglet Tableau de bord (et l'onglet `Calculs`) dans
`Donnees_propres.xlsx`. Il ne touche pas aux feuilles de données. Il faut le
relancer si les données changent de structure, par exemple avec un nouveau code
ATC ou une colonne renommée. Avec `--excel` et `--sortie`, il peut travailler
sur un autre classeur.

## Structure de l'Excel lue par le script

| Feuille                 | Utilisée pour |
|-------------------------|---------------|
| `Pharma_Ventes_Daily`   | total, jour, semaine, mois, année, jour de la semaine |
| `Pharma_Ventes_Hourly`  | profils heure par heure |
| `Pharma_Ventes_Weekly`, `Pharma_Ventes_Monthly` | non lues par le script (voir plus bas) |

Dans chaque feuille, le script lit la colonne `datum` et toutes les colonnes
dont le nom est un code ATC (`M01AB`, `N05B`, `R03`…). Un nouveau code ATC
ajouté en colonne est donc pris en compte automatiquement. Si une feuille ou la
colonne de date est renommée, il suffit de changer les constantes en haut du
script (`FEUILLE_JOUR`, `FEUILLE_HEURE`, `COLONNE_DATE`).

Les semaines, les mois et les années sont recalculés à partir de la feuille
journalière. Cela donne des résultats cohérents entre eux, et cela marche pour
n'importe quelle période (par exemple du 15 mars au 10 avril). Les feuilles
hebdomadaire et mensuelle ne concordent pas exactement avec le journalier.
L'export mensuel brut diverge déjà sur 31 mois, et les arrondis ont été faits
séparément dans chaque feuille.

## Interroger l'Excel

### Mode interactif

```bash
python interroger_excel.py
```

```
question> N02BE 2019-03                 # total de paracétamol en mars 2019
question> N05B N05C 2018 mois           # anxiolytiques + hypnotiques, mois par mois
question> R03 2019 jour top 5           # les 5 plus gros jours de R03 en 2019
question> jour_semaine moyenne          # vente moyenne par jour de la semaine
question> 2019-10-07 heure              # le 7 octobre 2019 heure par heure
question> liste                         # médicaments disponibles
question> aide
question> q
```

### En une commande

```bash
python interroger_excel.py --atc N02BE --periode 2019-03
python interroger_excel.py --atc N05B N05C --periode 2018 --par mois
python interroger_excel.py --atc R03 --periode 2019 --par jour --top 5
python interroger_excel.py --par jour_semaine --stat moyenne
python interroger_excel.py --periode 2019-10-07 --par heure
python interroger_excel.py --atc R06 --periode 2018-03:2018-06 --par semaine --export r06.csv
python interroger_excel.py --help
```

| Paramètre   | Valeurs |
|-------------|---------|
| `--atc`     | un ou plusieurs codes ATC (par défaut : tous) |
| `--periode` | `2019`, `2019-03`, `2019-03-15`, `15/03/2019`, ou une plage `2019-01:2019-06`, `2019-01:` (par défaut : tout l'historique) |
| `--par`     | `total` (défaut), `jour`, `semaine`, `mois`, `annee`, `jour_semaine`, `heure` |
| `--stat`    | `somme` (défaut) ou `moyenne` (moyenne par jour) |
| `--top N`   | ne garder que les N plus fortes valeurs |
| `--export`  | enregistrer le résultat en `.csv` (format Excel français) ou `.xlsx` |
| `--excel`   | interroger un autre classeur (par défaut : `Donnees_propres.xlsx`) |
| `--git`     | lire l'Excel dans le dépôt git (dernière version poussée), voir plus bas |
| `--branche`, `--remote` | branche et dépôt distant lus avec `--git` (défaut : branche courante, `origin`) |

La colonne « Nb jours » montre combien de jours de données couvre chaque ligne,
ce qui rend visibles les périodes incomplètes (octobre 2019 : 8 jours).

### Comment fonctionne le « live »

Le script ne garde aucune copie des ventes. À chaque question, il lit la
feuille utile dans `Donnees_propres.xlsx`. En mode interactif, il vérifie
d'abord si le fichier a été enregistré depuis la question précédente. Si oui,
la feuille est relue avant de répondre. Sinon, il réutilise la lecture
précédente pour aller plus vite. Une correction faite dans Excel est donc prise
en compte dès qu'on enregistre (Ctrl+S). Chaque réponse indique la feuille lue
et l'heure du dernier enregistrement du fichier.

### Connecteur git : lire l'Excel directement dans le dépôt

```bash
python interroger_excel.py --git                         # branche courante, remote origin
python interroger_excel.py --git --branche main --atc N05B --periode 2019 --par mois
```

Avec `--git`, le script ne lit plus le fichier du disque. À chaque question :

1. il fait un `git fetch` de la branche sur le dépôt distant ;
2. il lit l'Excel tel qu'il est dans le dernier commit poussé (`git cat-file`),
   sans toucher à votre copie locale ni à vos fichiers en cours de modification ;
3. il ne relit les feuilles que si l'Excel a changé depuis la question
   précédente.

Quand un membre de l'équipe pousse une nouvelle version de l'Excel, la question
suivante y répond déjà, sans relancer le script. La réponse indique le commit
lu (numéro, date, auteur). Si le dépôt est injoignable (pas de réseau), le
script répond avec la dernière version récupérée et le signale.

Le script utilise les identifiants git déjà configurés sur la machine (clé SSH
ou HTTPS). Il faut donc le lancer depuis un clone du dépôt.

**Pas besoin de cron.** Un cron lance une tâche à heure fixe, que quelqu'un
pose une question ou non. Ici, c'est la question qui déclenche la lecture de
l'Excel (et le `git fetch` avec `--git`). La réponse est donc toujours à jour,
sans tâche planifiée.

Les fonctions sont aussi utilisables depuis un autre script ou un notebook :

```python
from interroger_excel import ClasseurLive, SourceGit, interroger

classeur = ClasseurLive("Donnees_propres.xlsx")                  # fichier local
# classeur = ClasseurLive(SourceGit("Donnees_propres.xlsx"))     # ou dernière version poussée
reponse = interroger(classeur, atc=["N05B"], periode="2019", par="mois")
print(reponse.table)
```
