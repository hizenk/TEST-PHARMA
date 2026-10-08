"""Interroge en live le classeur Donnees_propres.xlsx.

Le script ne garde aucune copie des ventes : à chaque question, il va lire
l'Excel (et le relit dès que le fichier a été enregistré). Si quelqu'un met à
jour le classeur, la réponse suivante en tient compte immédiatement.

Avec --git, l'Excel n'est plus lu sur le disque mais dans le dépôt git : à
chaque question, le script récupère (git fetch) la dernière version poussée
sur la branche et la lit directement, sans toucher à la copie locale.

Exemples :
  python interroger_excel.py                                   # mode interactif
  python interroger_excel.py --git                             # mode interactif, Excel lu dans le dépôt
  python interroger_excel.py --git --branche main --atc N05B --periode 2019 --par mois
  python interroger_excel.py --atc N02BE --periode 2019-03
  python interroger_excel.py --atc N05B N05C --periode 2018 --par mois
  python interroger_excel.py --periode 2019-10-07 --par heure
  python interroger_excel.py --atc R03 --periode 2019 --par jour --top 5
  python interroger_excel.py --par jour_semaine --stat moyenne
  python interroger_excel.py --atc R06 --periode 2018-03:2018-06 --par semaine --export r06.csv
  python interroger_excel.py --liste
"""
import argparse
import io
import re
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

ICI = Path(__file__).resolve().parent
EXCEL_PAR_DEFAUT = ICI / "Donnees_propres.xlsx"

# Structure du classeur : à adapter ici si une feuille ou une colonne est renommée.
FEUILLE_JOUR = "Pharma_Ventes_Daily"
FEUILLE_HEURE = "Pharma_Ventes_Hourly"
COLONNE_DATE = "datum"
CODE_ATC = re.compile(r"[A-Z]\d{2}[A-Z]{0,2}")  # M01AB, N05B, R03...

LIBELLES = {
    "M01AB": "Anti-inflammatoires (AINS), dérivés de l'acide acétique : diclofénac",
    "M01AE": "Anti-inflammatoires (AINS), dérivés de l'acide propionique : ibuprofène",
    "N02BA": "Analgésiques, acide salicylique et dérivés : aspirine",
    "N02BE": "Analgésiques, anilides : paracétamol",
    "N05B": "Anxiolytiques",
    "N05C": "Hypnotiques et sédatifs",
    "R03": "Médicaments de l'asthme et de la BPCO",
    "R06": "Antihistaminiques (allergies)",
}
JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

# --par -> (feuille lue, colonne de regroupement calculée à partir de la date)
REGROUPEMENTS = {
    "total": (FEUILLE_JOUR, None),
    "jour": (FEUILLE_JOUR, "Date"),
    "semaine": (FEUILLE_JOUR, "Semaine"),
    "mois": (FEUILLE_JOUR, "Mois"),
    "annee": (FEUILLE_JOUR, "Année"),
    "jour_semaine": (FEUILLE_JOUR, "Jour"),
    "heure": (FEUILLE_HEURE, "Heure"),
}
STATS = ("somme", "moyenne")
ALIAS = {"année": "annee", "an": "annee", "jour-semaine": "jour_semaine", "heures": "heure", "moy": "moyenne"}
MOTS_IGNORES = {"par", "en", "de", "du", "des", "pour", "ventes", "vente", "le", "la", "les"}

AIDE = """
Tapez une question avec, dans n'importe quel ordre :
  - un ou plusieurs codes ATC .......... N02BE N05B      (rien = tous les médicaments)
  - une date ou une période ............ 2019 | 2019-03 | 2019-03-15 | 15/03/2019
                                          2019-01:2019-06 | 2019-01-01:2019-03-31 | 2019-01:  (rien = tout)
  - un regroupement .................... total | jour | semaine | mois | annee | jour_semaine | heure
  - une statistique .................... somme (défaut) | moyenne  (moyenne par jour)
  - un classement ...................... top 5  (les 5 plus fortes valeurs)

Exemples :
  N02BE 2019-03                 total de paracétamol vendu en mars 2019
  N05B N05C 2018 mois           anxiolytiques et hypnotiques, mois par mois en 2018
  R03 2019 jour top 5           les 5 jours où il s'est vendu le plus de R03 en 2019
  jour_semaine moyenne          ventes moyennes par jour de la semaine, tout l'historique
  2019-10-07 heure              ventes heure par heure le 7 octobre 2019

Autres commandes : liste (médicaments disponibles), aide, q (quitter)
"""


@dataclass
class Reponse:
    titre: str
    details: list
    table: pd.DataFrame
    debut: pd.Timestamp = None  # premier et dernier jour réellement couverts par les données
    fin: pd.Timestamp = None
    nb_jours: int = 0


def ajouter_calendrier(df):
    """Ajoute à une feuille de ventes les colonnes servant à filtrer et regrouper."""
    df = df.dropna(subset=[COLONNE_DATE]).copy()
    dates = pd.to_datetime(df[COLONNE_DATE])
    iso = dates.dt.isocalendar()
    df["Date"] = dates.dt.normalize()
    df["Heure"] = dates.dt.hour
    df["Année"] = dates.dt.year
    df["Mois"] = dates.dt.strftime("%Y-%m")
    df["Semaine"] = iso["year"].astype(str) + "-S" + iso["week"].astype(str).str.zfill(2)
    df["Jour"] = dates.dt.dayofweek.map(dict(enumerate(JOURS)))
    return df


class SourceFichier:
    """L'Excel tel qu'il est enregistré sur le disque."""

    def __init__(self, chemin):
        self.chemin = Path(chemin)
        self.nom = self.chemin.name
        self.verifie_le = None

    def actualiser(self, forcer=False):
        """Renvoie une signature qui change dès que le fichier est réenregistré."""
        if not self.chemin.exists():
            raise FileNotFoundError(f"Classeur introuvable : {self.chemin}")
        stat = self.chemin.stat()
        self.verifie_le = datetime.now()
        return stat.st_mtime_ns, stat.st_size

    def contenu(self):
        return self.chemin

    def decrire(self):
        modifie = datetime.fromtimestamp(self.chemin.stat().st_mtime)
        return f"{self.nom} (fichier enregistré le {modifie:%d/%m/%Y à %H:%M:%S})"


class SourceGit:
    """Connecteur git : l'Excel est lu dans la dernière version poussée sur une branche du dépôt."""

    def __init__(self, chemin, branche=None, remote="origin", intervalle_fetch=0):
        dossier = Path(chemin).resolve().parent
        try:
            self.depot = Path(self._git_dans(dossier, "rev-parse", "--show-toplevel").decode().strip()).resolve()
        except RuntimeError as erreur:
            raise RuntimeError(f"{dossier} n'est pas dans un dépôt git : clonez le dépôt "
                               f"ou lancez le script sans --git ({erreur})") from erreur
        self.fichier = Path(chemin).resolve().relative_to(self.depot).as_posix()
        self.nom = Path(chemin).name
        self.remote = remote
        self.branche = branche or self._git("rev-parse", "--abbrev-ref", "HEAD").decode().strip()
        if self.branche == "HEAD":
            raise RuntimeError("Aucune branche courante : précisez la branche à lire avec --branche")
        self.ref = f"{remote}/{self.branche}"
        self.blob, self.octets, self.commit, self.avertissement = None, None, "", ""
        self.intervalle_fetch = intervalle_fetch  # secondes minimum entre deux git fetch (0 = à chaque question)
        self._dernier_fetch = None
        self.verifie_le = None

    @staticmethod
    def _git_dans(dossier, *args):
        try:
            return subprocess.run(["git", "-C", str(dossier), *args], capture_output=True, check=True).stdout
        except FileNotFoundError as erreur:
            raise RuntimeError("git n'est pas installé ou introuvable dans le PATH") from erreur
        except subprocess.CalledProcessError as erreur:
            lignes = [l.strip() for l in erreur.stderr.decode(errors="replace").splitlines() if l.strip()]
            message = next((l for l in lignes if l.startswith(("fatal:", "error:"))), lignes[0] if lignes else "échec")
            raise RuntimeError(f"git {' '.join(args)} : {message}") from erreur

    def _git(self, *args):
        return self._git_dans(self.depot, *args)

    def actualiser(self, forcer=False):
        """Récupère la branche distante ; renvoie l'identifiant (blob) de la version poussée de l'Excel."""
        recent = self._dernier_fetch is not None and time.monotonic() - self._dernier_fetch < self.intervalle_fetch
        if self.blob is not None and recent and not forcer:
            return self.blob
        erreur_fetch = None
        try:
            self._git("fetch", "--quiet", self.remote, self.branche)
            self._dernier_fetch = time.monotonic()
            self.verifie_le = datetime.now()
            self.avertissement = ""
        except RuntimeError as erreur:
            erreur_fetch = erreur
            self.avertissement = f" | dépôt injoignable, dernière version récupérée utilisée ({erreur})"
        try:
            blob = self._git("rev-parse", f"{self.ref}:{self.fichier}").decode().strip()
        except RuntimeError as erreur:
            cause = f" | {erreur_fetch}" if erreur_fetch else ""
            raise RuntimeError(f"{self.fichier} introuvable sur {self.ref} (pas encore poussé ?){cause}") from erreur
        if blob != self.blob:
            self.octets = self._git("cat-file", "blob", blob)
            self.commit = self._git("log", "-1", "--date=format:%d/%m/%Y %H:%M",
                                    "--format=commit %h du %cd par %an", self.ref, "--", self.fichier
                                    ).decode().strip()
            self.blob = blob
        return blob

    def contenu(self):
        return io.BytesIO(self.octets)

    def decrire(self):
        return f"dépôt git {self.ref}:{self.fichier} ({self.commit}){self.avertissement}"


class ClasseurLive:
    """Accès au classeur : à chaque question, on regarde si l'Excel a changé et on relit les feuilles utiles."""

    def __init__(self, source):
        self.source = source if hasattr(source, "actualiser") else SourceFichier(source)
        self.signature = None
        self._cache = {}  # nom de feuille -> (signature de la version lue, DataFrame)

    def actualiser(self, forcer=False):
        """À appeler au début de chaque question (fichier réenregistré ? nouveau push sur le dépôt ?)."""
        self.signature = self.source.actualiser(forcer)

    def feuille(self, nom):
        if self.signature is None:
            self.actualiser()
        en_cache = self._cache.get(nom)
        if en_cache is None or en_cache[0] != self.signature:
            try:
                df = pd.read_excel(self.source.contenu(), sheet_name=nom)
            except ValueError as erreur:
                raise ValueError(f"Feuille « {nom} » absente du classeur {self.source.nom}") from erreur
            if COLONNE_DATE not in df.columns:
                raise ValueError(f"Colonne « {COLONNE_DATE} » absente de la feuille {nom}")
            self._cache[nom] = (self.signature, ajouter_calendrier(df))
        return self._cache[nom][1]

    def codes_atc(self):
        return [c for c in self.feuille(FEUILLE_JOUR).columns if CODE_ATC.fullmatch(str(c))]

    def medicaments(self):
        codes = self.codes_atc()
        return pd.DataFrame({"Libellé": [LIBELLES.get(c, "") for c in codes]}, index=pd.Index(codes, name="Code ATC"))


def _borne(texte, fin):
    """Transforme '2019', '2019-03', '03/2019', '2019-03-15' ou '15/03/2019' en date de début ou de fin."""
    t = texte.strip()
    try:
        if m := re.fullmatch(r"(\d{4})", t):
            debut = pd.Timestamp(int(m[1]), 1, 1)
            return debut + pd.offsets.YearEnd(0) if fin else debut
        if m := re.fullmatch(r"(\d{4})-(\d{1,2})", t) or re.fullmatch(r"(?P<m>\d{1,2})/(?P<a>\d{4})", t):
            annee, mois = (m["a"], m["m"]) if "/" in t else (m[1], m[2])
            debut = pd.Timestamp(int(annee), int(mois), 1)
            return debut + pd.offsets.MonthEnd(0) if fin else debut
        if m := re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", t):
            return pd.Timestamp(int(m[1]), int(m[2]), int(m[3]))
        if m := re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", t):
            return pd.Timestamp(int(m[3]), int(m[2]), int(m[1]))
    except ValueError as erreur:
        raise ValueError(f"Date invalide : {texte!r} ({erreur})") from erreur
    raise ValueError(f"Période non reconnue : {texte!r} (exemples : 2019, 2019-03, 2019-03-15, 2019-01:2019-06)")


def lire_periode(texte):
    """Renvoie (début, fin) inclusifs ; None signifie « pas de limite »."""
    if not texte:
        return None, None
    if ":" in texte:
        gauche, droite = texte.split(":", 1)
        debut = _borne(gauche, fin=False) if gauche.strip() else None
        fin = _borne(droite, fin=True) if droite.strip() else None
    else:
        debut, fin = _borne(texte, fin=False), _borne(texte, fin=True)
    if debut is not None and fin is not None and debut > fin:
        raise ValueError(f"Période à l'envers : {texte!r}")
    return debut, fin


def normaliser_atc(atc, codes_connus):
    if not atc:
        return list(codes_connus)
    choisis = []
    for code in (c.strip().upper() for bloc in atc for c in bloc.split(",") if c.strip()):
        if code not in codes_connus:
            raise ValueError(f"Code ATC inconnu : {code} (disponibles : {', '.join(codes_connus)})")
        if code not in choisis:
            choisis.append(code)
    return choisis


def interroger(classeur, atc=None, periode=None, par="total", stat="somme", top=None):
    """Va chercher dans l'Excel les ventes demandées et renvoie une Reponse."""
    par, stat = par or "total", stat or "somme"
    if par not in REGROUPEMENTS:
        raise ValueError(f"Regroupement inconnu : {par} (choix : {', '.join(REGROUPEMENTS)})")
    if stat not in STATS:
        raise ValueError(f"Statistique inconnue : {stat} (choix : {', '.join(STATS)})")

    classeur.actualiser()
    codes = normaliser_atc(atc, classeur.codes_atc())
    debut, fin = lire_periode(periode)
    nom_feuille, colonne = REGROUPEMENTS[par]
    ventes = classeur.feuille(nom_feuille)

    masque = pd.Series(True, index=ventes.index)
    if debut is not None:
        masque &= ventes["Date"] >= debut
    if fin is not None:
        masque &= ventes["Date"] <= fin
    ventes = ventes[masque]
    if ventes.empty:
        jour = classeur.feuille(FEUILLE_JOUR)["Date"]
        raise ValueError(f"Aucune vente sur cette période (données du {jour.min():%d/%m/%Y} "
                         f"au {jour.max():%d/%m/%Y})")

    quantites = ventes[codes].astype(float)
    if len(codes) > 1:
        quantites["Total"] = quantites.sum(axis=1)

    if colonne is None:
        table = pd.DataFrame({
            "Total": quantites.sum(),
            "Moyenne / jour": quantites.mean(),
            "Max / jour": quantites.max(),
            "Jour du max": [f"{ventes.at[i, 'Date']:%d/%m/%Y}" for i in quantites.idxmax()],
        })
        table.insert(0, "Libellé", [LIBELLES.get(c, "Tous les médicaments choisis" if c == "Total" else "")
                                    for c in table.index])
        table.index.name = "Code ATC"
    else:
        groupes = quantites.groupby(ventes[colonne], sort=True)
        table = groupes.sum() if stat == "somme" else groupes.mean()
        if par != "jour":
            table.insert(0, "Nb jours", groupes.size())
        if par == "jour_semaine":
            table = table.reindex([j for j in JOURS if j in table.index])
        if par == "jour":
            table.index = table.index.strftime("%d/%m/%Y")
        if par == "heure":
            table.index = [f"{h:02d}h" for h in table.index]
        if top:
            table = table.sort_values(table.columns[-1], ascending=False).head(top)

    jours = ventes["Date"]
    stat_txt = "" if colonne is None else f" ({'somme' if stat == 'somme' else 'moyenne par jour'})"
    titre = (f"Ventes de {', '.join(codes)} | "
             + ("total sur la période" if colonne is None else f"par {par}{stat_txt}")
             + (f" | top {top}" if top and colonne else ""))
    details = [
        f"Période : {jours.min():%d/%m/%Y} → {jours.max():%d/%m/%Y} "
        f"({jours.nunique()} jour{'s' if jours.nunique() > 1 else ''} de données)",
        f"Source : feuille {nom_feuille} de {classeur.source.decrire()}",
    ]
    return Reponse(titre, details, table, jours.min(), jours.max(), jours.nunique())


def format_fr(x):
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")


def afficher(reponse):
    print(f"\n=== {reponse.titre}")
    for ligne in reponse.details:
        print(f"    {ligne}")
    print()
    with pd.option_context("display.max_rows", 100, "display.min_rows", 30, "display.width", 250,
                           "display.max_columns", None, "display.max_colwidth", 70,
                           "display.float_format", format_fr):
        print(reponse.table)


def afficher_medicaments(classeur):
    classeur.actualiser()
    jour = classeur.feuille(FEUILLE_JOUR)["Date"]
    print(f"\nMédicaments disponibles dans {classeur.source.decrire()}\n"
          f"Données du {jour.min():%d/%m/%Y} au {jour.max():%d/%m/%Y} :\n")
    with pd.option_context("display.max_colwidth", 80, "display.width", 250):
        print(classeur.medicaments())


def exporter(reponse, chemin):
    chemin = Path(chemin)
    if chemin.suffix.lower() == ".xlsx":
        reponse.table.to_excel(chemin)
    else:  # CSV lisible directement par un Excel français
        reponse.table.to_csv(chemin, sep=";", decimal=",", encoding="utf-8-sig")
    print(f"\nRésultat exporté dans {chemin}")


def lire_question(ligne, codes_connus):
    """Traduit une question tapée en mode interactif (« N05B 2019 mois top 3 ») en paramètres."""
    params = {"atc": [], "periode": None, "par": None, "stat": None, "top": None}
    mots = ligne.replace(",", " ").split()
    i = 0
    while i < len(mots):
        mot = mots[i]
        m = ALIAS.get(mot.lower(), mot.lower())
        if mot.upper() in codes_connus:
            params["atc"].append(mot.upper())
        elif m in REGROUPEMENTS:
            params["par"] = m
        elif m in STATS:
            params["stat"] = m
        elif m == "top" and i + 1 < len(mots) and mots[i + 1].isdigit():
            params["top"] = int(mots[i + 1])
            i += 1
        elif re.fullmatch(r"top\d+", m):
            params["top"] = int(m[3:])
        elif re.fullmatch(r"[\d/:-]+", mot):
            params["periode"] = mot
        elif m not in MOTS_IGNORES:
            raise ValueError(f"Mot non compris : {mot!r} (tapez « aide »)")
        i += 1
    return params


def mode_interactif(classeur):
    source = classeur.source
    if isinstance(source, SourceGit):
        print(f"Interrogation live de {source.fichier} sur la branche {source.ref} (connecteur git)")
        print("Chaque question récupère la dernière version de l'Excel poussée sur le dépôt.")
    else:
        print(f"Interrogation live de {source.chemin}")
        print("Chaque question relit l'Excel s'il a été enregistré entre-temps.")
    print("Tapez « aide » pour les exemples, « q » pour quitter.")
    while True:
        try:
            ligne = input("\nquestion> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return
        commande = ligne.lower()
        if not ligne:
            continue
        if commande in ("q", "quit", "quitter", "exit"):
            return
        try:
            if commande in ("aide", "help", "?"):
                print(AIDE)
            elif commande in ("liste", "medicaments", "médicaments"):
                afficher_medicaments(classeur)
            else:
                afficher(interroger(classeur, **lire_question(ligne, classeur.codes_atc())))
        except (ValueError, FileNotFoundError, RuntimeError) as erreur:
            print(f"Erreur : {erreur}")


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Interroge en live le classeur des ventes de la pharmacie.",
        epilog="Exemples :" + __doc__.split("Exemples :", 1)[1],
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--excel", type=Path, default=EXCEL_PAR_DEFAUT, help="classeur à interroger")
    parser.add_argument("--git", action="store_true",
                        help="lire l'Excel dans le dépôt git (dernière version poussée) au lieu du disque")
    parser.add_argument("--branche", help="branche à lire avec --git (défaut : la branche courante)")
    parser.add_argument("--remote", default="origin", help="dépôt distant à lire avec --git (défaut : origin)")
    parser.add_argument("--atc", nargs="+", metavar="CODE", help="un ou plusieurs codes ATC (défaut : tous)")
    parser.add_argument("--periode", help="2019 | 2019-03 | 2019-03-15 | 2019-01:2019-06 (défaut : tout)")
    parser.add_argument("--par", choices=REGROUPEMENTS, help="regroupement (défaut : total)")
    parser.add_argument("--stat", choices=STATS, help="somme (défaut) ou moyenne par jour")
    parser.add_argument("--top", type=int, metavar="N", help="ne garder que les N plus fortes valeurs")
    parser.add_argument("--export", type=Path, metavar="FICHIER", help="enregistrer le résultat (.csv ou .xlsx)")
    parser.add_argument("--liste", action="store_true", help="afficher les médicaments disponibles")
    args = parser.parse_args(argv)

    try:
        source = SourceGit(args.excel, args.branche, args.remote) if args.git else SourceFichier(args.excel)
        classeur = ClasseurLive(source)
        if args.liste:
            afficher_medicaments(classeur)
            return 0
        question = (args.atc, args.periode, args.par, args.stat, args.top, args.export)
        if all(v is None for v in question):
            mode_interactif(classeur)
            return 0
        reponse = interroger(classeur, args.atc, args.periode, args.par, args.stat, args.top)
        afficher(reponse)
        if args.export:
            exporter(reponse, args.export)
    except (ValueError, FileNotFoundError, RuntimeError) as erreur:
        print(f"Erreur : {erreur}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
