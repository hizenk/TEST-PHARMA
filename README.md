# Rush 2 : ventes de la pharmacie

Un classeur Excel propre (`Donnees_propres.xlsx`) avec un **tableau de bord
qui fonctionne entièrement dans Excel**, et deux scripts Python : l'un fabrique
le tableau de bord, l'autre va chercher les informations dans l'Excel au moment
où on les demande.

```
TEST-PHARMA/
├── Donnees_propres.xlsx       # données propres + onglet « Tableau de bord »
├── creer_tableau_de_bord.py   # (re)construit le classeur et son tableau de bord
├── interroger_excel.py        # interroge le classeur en live (fichier local ou dépôt git)
├── exports/                   # les 4 exports CSV bruts du client
└── requirements.txt
```

## Le tableau de bord (dans Excel, sans macro)

Ouvrez `Donnees_propres.xlsx` (sur Mac : clic droit > *Ouvrir avec* >
*Microsoft Excel*). Le classeur s'ouvre sur l'onglet **Tableau de bord**, qui
se lit de haut en bas. Seules les **cellules vert clair** se modifient.

| Bloc | Contenu |
|------|---------|
| ① Médicaments | `Oui` / `Non` devant chacun des 8 groupes ATC, total de la période et part de chacun |
| ② Période et affichage | dates Du / Au, regroupement (`Jour`, `Semaine`, `Mois`, `Année`, `Jour de la semaine`, `Heure`), `Somme` ou `Moyenne par jour` |
| ③ Chiffres clés | quantité vendue, moyenne par jour, meilleur jour, nombre de jours |
| ④ Graphiques | évolution de la sélection, ventes moyennes par jour de la semaine et par heure |
| ⑤ Tableau des achats | quantité à commander par médicament (voir ci-dessous), jour et mois les plus forts |
| ⑥ Résultats détaillés | une ligne par jour, semaine, mois… (jusqu'à 400 lignes), une colonne par médicament, total et nombre de jours |

**Tableau des achats.** On indique le nombre de jours à couvrir, une marge de
sécurité et le stock actuel de chaque médicament. Pour chaque médicament :

    à commander = moyenne par jour sur la période choisie × jours à couvrir × (1 + marge) − stock actuel

Le résultat est arrondi à l'unité supérieure, et jamais négatif. Le jour de la
semaine et le mois les plus forts sont calculés sur tout l'historique : ils
indiquent quand prévoir plus de stock.

Un message rouge prévient en cas de problème : date de fin avant la date de
début, aucun médicament inclus, ou plus de 400 lignes (choisir alors un
regroupement plus large).

Comment c'est construit :

- **des formules Excel simples** (`SUMIF`, `SUMIFS`, `COUNTIFS`, `AVERAGEIF`,
  `INDEX`/`MATCH`), sans macro, sans nom de tableau ni nom défini. Le classeur
  marche dans Excel sur Windows, sur Mac et en ligne, comme dans LibreOffice ;
- **des valeurs déjà calculées** : les chiffres s'affichent dès l'ouverture,
  même avant le recalcul d'Excel ;
- **chaque graphique est ancré dans sa propre zone de cellules**, pour
  qu'aucun bloc ne se chevauche, quel que soit l'écran ;
- **l'onglet est protégé, sans mot de passe**, pour éviter d'écraser une
  formule par erreur. Pour le modifier : *Révision > Ôter la protection de la feuille* ;
- **des onglets de calcul masqués** : `Calculs` pour les paramètres, les
  périodes et les profils, `Jours` pour une ligne de calcul par jour de données.

## Installation

```bash
git clone https://github.com/hizenk/TEST-PHARMA.git
cd TEST-PHARMA
pip install -r requirements.txt
```

## Recréer le classeur et son tableau de bord

```bash
python creer_tableau_de_bord.py
python creer_tableau_de_bord.py --excel classeur_source.xlsx --sortie Donnees_propres.xlsx
```

Le script lit les feuilles de données du classeur source (par défaut
`Donnees_propres.xlsx`) et réécrit un classeur propre : les données recopiées
en tableaux Excel ordinaires, le tableau de bord et les onglets de calcul. Les
données sont recopiées sous forme de valeurs. Les requêtes Power Query d'un
classeur source ne sont pas reprises, car elles pointent vers des fichiers qui
n'existent pas chez le client. Relancez-le après chaque mise à jour des
données. Les formules sont dimensionnées sur le nombre de lignes présentes.

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
