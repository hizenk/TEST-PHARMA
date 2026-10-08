# Rush 2 : ventes de la pharmacie

Analyse des ventes de huit groupes de médicaments d'une pharmacie indépendante,
de janvier 2014 à octobre 2019. Le dépôt contient :
- les quatre exports du client ;
- le code Python qui produit tout le classeur à partir de ces exports ;
- le classeur lui-même (`Donnees_propres.xlsx`), avec l'outil de sélection du pharmacien.

```
TEST-PHARMA/
├── Donnees_propres.xlsx            # le classeur livré au client (généré par le code)
├── creer_classeur.py               # point d'entrée : exports → classeur complet
├── analyse/
│   ├── donnees.py                  # lecture des exports, nettoyage, contrôles de cohérence
│   ├── statistiques.py             # statistiques descriptives
│   ├── prevision.py                # test de prévision du mois suivant
│   ├── externe.py                  # source publique : grippe (réseau Sentinelles)
│   ├── tableau_de_bord.py          # outil de sélection (formules Excel, sans macro)
│   ├── feuilles.py                 # écriture des onglets d'analyse
│   └── mise_en_forme.py            # thème vert et mise en page communs
├── interroger_excel.py             # interroge le classeur en live (fichier local ou dépôt git)
├── exports/                        # les 4 exports CSV du client (horaire, jour, semaine, mois)
├── donnees_externes/               # copie des données Sentinelles utilisées
└── requirements.txt
```

## Installation

```bash
git clone https://github.com/hizenk/TEST-PHARMA.git
cd TEST-PHARMA
pip install -r requirements.txt
```

## Régénérer le classeur

```bash
python creer_classeur.py                  # exports/ → Donnees_propres.xlsx
python creer_classeur.py --maj-externe    # retélécharge aussi les données Sentinelles
```

Chaque mois, le client envoie les quatre fichiers avec tout l'historique. Il
suffit de remplacer ceux du dossier `exports/` et de relancer la commande. Les
contrôles, les statistiques, le test de prévision, l'analyse de la grippe et
l'outil de sélection sont recalculés, sans rien refaire à la main. Le code
tourne depuis une copie neuve du dépôt.

## Les onglets du classeur

| Onglet | Contenu |
|--------|---------|
| **Synthèse** | Chiffres clés, ce qu'il faut retenir, recommandations pour le pharmacien achats, le propriétaire et le manager. Le texte est généré à partir des chiffres calculés. |
| **Tableau de bord** | L'outil de sélection : médicaments `Oui`/`Non`, une période, un regroupement. Chiffres clés, graphiques, tableau des achats et résultats détaillés. Formules Excel, sans macro. |
| **Statistiques** | Poids de chaque groupe, irrégularité, tendance par année, saisonnalité mensuelle, jours de la semaine, tranches horaires, corrélations. |
| **Prévision** | Réponse à la question du manager : test sur 12 mois non utilisés, face à la prévision naïve, et prévision du mois suivant. |
| **Source externe** | Incidence de la grippe (réseau Sentinelles) et ventes hebdomadaires : ce qui relève de l'environnement. |
| **État des données** | Quel export sert à quoi, contrôles de cohérence, et ce qui a été corrigé ou écarté. |
| `Pharma_Ventes_*`, `Grippe_Sentinelles` | Les données (exports et grippe) sous forme de tableaux Excel. |

Sur Mac : clic droit sur le fichier, puis *Ouvrir avec* > *Microsoft Excel*.

### Choix sur les données (détail dans l'onglet État des données)

- **L'export journalier sert de référence.** Il concorde exactement avec
  l'horaire et l'hebdomadaire.
- **L'export mensuel est écarté**, car il diverge sur 31 mois. Les ventes
  mensuelles sont recalculées à partir du journalier.
- **Jours et mois incomplets.** Les statistiques par jour portent sur les jours
  complets (24 heures présentes). La prévision porte sur les mois complets
  (février 2014 → septembre 2019).
- **Les quantités sont gardées telles qu'enregistrées**, décimales comprises.
  Un premier import Power Query arrondissait cinq groupes à l'entier, parce
  qu'il avait déduit le type « nombre entier » des premières lignes.

### Prévision du mois suivant

1. Six méthodes simples sont comparées : prévision naïve, même mois l'an dernier,
   moyennes mobiles, saisonnalité × niveau, régression tendance + mois.
2. Chaque méthode ne voit que les mois qui précèdent le mois prévu.
3. La méthode de chaque groupe est choisie sur 12 mois de validation, puis
   jugée sur les 12 mois suivants, qu'elle n'a jamais vus.
4. Verdict « Oui » si elle fait au moins 15 % d'erreur en moins que la
   prévision naïve, c'est-à-dire refaire les ventes du mois précédent.

### Source externe

Les données viennent du [réseau Sentinelles](https://www.sentiweb.fr) (Inserm) :
l'incidence hebdomadaire des syndromes grippaux en France, sous forme de
données ouvertes. Le fichier utilisé est copié dans `donnees_externes/`.

Les ventes des semaines d'épidémie (au moins 150 cas pour 100 000 habitants)
sont comparées à celles des autres semaines, sur toute l'année et à saison
égale (novembre à mars). La comparaison à saison égale sépare l'effet de la
grippe de celui de l'hiver.

### L'outil de sélection (onglet Tableau de bord)

Seules les cellules vert clair se modifient :
1. `Oui`/`Non` devant chaque médicament ;
2. les dates Du / Au ;
3. le regroupement (jour, semaine, mois, année, jour de la semaine, heure) ;
4. somme ou moyenne par jour.

Les chiffres clés, les graphiques et les résultats détaillés se recalculent
aussitôt. Le **tableau des achats** propose une quantité à commander par
médicament :

    à commander = moyenne par jour × jours à couvrir × (1 + marge) − stock actuel

Comment c'est construit :
- des formules Excel simples (`SUMIF`, `SUMIFS`, `COUNTIFS`, `AVERAGEIF`,
  `INDEX`/`MATCH`), sans macro, sans nom de tableau ni nom défini ;
- des valeurs déjà calculées, pour que les chiffres s'affichent dès l'ouverture ;
- un onglet protégé, sans mot de passe (*Révision > Ôter la protection*).

## Qui a fait quoi

| Membre | Responsabilités |
|--------|-----------------|
| *à compléter* | |
| *à compléter* | |
| *à compléter* | |

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
