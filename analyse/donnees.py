"""Lecture des quatre exports CSV, nettoyage et contrôles de cohérence.

L'export journalier sert de référence ; les autres exports sont confrontés à lui.
Tout ce qui est corrigé ou écarté est consigné pour l'onglet « État des données ».
"""
import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

CODE_ATC = re.compile(r"[A-Z]\d{2}[A-Z]{0,2}")
NOMS = {
    "M01AB": "Diclofénac", "M01AE": "Ibuprofène", "N02BA": "Aspirine", "N02BE": "Paracétamol",
    "N05B": "Anxiolytiques", "N05C": "Hypnotiques, sédatifs", "R03": "Asthme, BPCO", "R06": "Antihistaminiques",
}
LIBELLES = {
    "M01AB": "Anti-inflammatoires, dérivés de l'acide acétique (diclofénac)",
    "M01AE": "Anti-inflammatoires, dérivés de l'acide propionique (ibuprofène)",
    "N02BA": "Analgésiques, acide salicylique et dérivés (aspirine)",
    "N02BE": "Analgésiques, anilides (paracétamol)",
    "N05B": "Anxiolytiques",
    "N05C": "Hypnotiques et sédatifs",
    "R03": "Médicaments des syndromes obstructifs des voies aériennes (asthme, BPCO)",
    "R06": "Antihistaminiques à usage systémique (allergies)",
}
# clé -> (fichier, format de la colonne datum, nom de la feuille dans le classeur, rôle)
EXPORTS = {
    "horaire": ("Pharma_Ventes_Hourly.csv", "%m/%d/%Y %H:%M", "Pharma_Ventes_Hourly"),
    "journalier": ("Pharma_Ventes_Daily.csv", "%m/%d/%Y", "Pharma_Ventes_Daily"),
    "hebdomadaire": ("Pharma_Ventes_Weekly.csv", "%m/%d/%Y", "Pharma_Ventes_Weekly"),
    "mensuel": ("Pharma_Ventes_Monthly.csv", "%Y-%m-%d", "Pharma_Ventes_Monthly"),
}
TOLERANCE = 0.01   # écart en dessous duquel deux exports sont considérés égaux
SEUIL_FAIBLE = 0.15  # jour « très faible » : moins de 15 % de la vente médiane d'un jour


def nom(code):
    return f"{code} · {NOMS.get(code, code)}"


@dataclass
class Donnees:
    codes: list
    exports: dict             # clé -> DataFrame nettoyé (colonnes d'origine, datum en datetime)
    etat_exports: pd.DataFrame
    controles: pd.DataFrame
    ecarts_mensuel: pd.DataFrame
    decisions: list
    jours: pd.DataFrame       # index = date, une colonne par code (tous les jours du journalier)
    jours_complets: pd.DataFrame  # idem, limité aux jours dont les 24 heures sont présentes
    heures: pd.DataFrame      # export horaire avec colonnes Date et Heure, jours complets seulement
    mois: pd.DataFrame        # mois complets recalculés depuis le journalier (index = 1er du mois)
    jours_faibles: pd.Series
    pics: pd.Series

    @property
    def premier_jour(self):
        return self.jours.index.min()

    @property
    def dernier_jour(self):
        return self.jours.index.max()


def lire_export(chemin, format_date, codes=None):
    """Lit un export ; renvoie (DataFrame nettoyé, ligne d'état, codes ATC)."""
    brut = pd.read_csv(chemin)
    if "datum" not in brut.columns:
        raise SystemExit(f"Colonne « datum » absente de {chemin}")
    codes = codes or [c for c in brut.columns if CODE_ATC.fullmatch(str(c))]
    dates = pd.to_datetime(brut["datum"], format=format_date, errors="coerce")
    a_relire = dates.isna() & brut["datum"].notna()
    if a_relire.any():  # format différent dans un nouvel export : seconde lecture plus tolérante
        dates[a_relire] = pd.to_datetime(brut.loc[a_relire, "datum"], errors="coerce")
    illisibles = int(dates.isna().sum())
    df = brut.assign(datum=dates).dropna(subset=["datum"])
    for c in codes:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    manquants = int(df[codes].isna().sum().sum())
    negatifs = int((df[codes] < 0).sum().sum())
    df[codes] = df[codes].mask(df[codes] < 0)  # une quantité négative est écartée (case vide)
    doublons = int(df["datum"].duplicated().sum())
    df = df.sort_values("datum").drop_duplicates("datum", keep="last").reset_index(drop=True)
    etat = {"Lignes": len(brut), "Première date": df["datum"].min(), "Dernière date": df["datum"].max(),
            "Dates illisibles": illisibles, "Doublons supprimés": doublons, "Valeurs manquantes": manquants,
            "Valeurs négatives écartées": negatifs}
    return df, etat, codes


def comparer(reference, export, libelle):
    """Écarts entre un export agrégé et la même agrégation recalculée depuis le journalier."""
    joint = export.join(reference, rsuffix="_ref", how="inner")
    lignes = []
    for code in reference.columns:
        diff = joint[code] - joint[f"{code}_ref"]
        for periode in diff[diff.abs() > TOLERANCE].index:
            lignes.append({"Période": periode, "Code": code, f"Export {libelle}": joint.at[periode, code],
                           "Recalcul journalier": joint.at[periode, f"{code}_ref"], "Écart": diff[periode]})
    return pd.DataFrame(lignes), len(joint)


def charger(dossier):
    dossier = Path(dossier)
    lus, etats = {}, []
    codes = None
    for cle in ("journalier", "horaire", "hebdomadaire", "mensuel"):
        fichier, fmt, feuille = EXPORTS[cle]
        if not (dossier / fichier).exists():
            raise SystemExit(f"Export introuvable : {dossier / fichier}")
        df, etat, codes_lus = lire_export(dossier / fichier, fmt, codes)
        codes = codes or codes_lus
        lus[cle] = df
        etats.append({"Export": cle.capitalize(), "Fichier": fichier, **etat})

    jour = lus["journalier"].set_index("datum")[codes]
    jour.index = jour.index.normalize()
    horaire = lus["horaire"].copy()
    horaire["Date"] = horaire["datum"].dt.normalize()
    horaire["Heure"] = horaire["datum"].dt.hour

    # Contrôles de cohérence
    par_jour_h = horaire.groupby("Date")[codes].sum()
    commun = jour.index.intersection(par_jour_h.index)
    ecarts_h = int(((jour.loc[commun] - par_jour_h.loc[commun]).abs() > TOLERANCE).sum().sum())
    seulement = len(jour.index.symmetric_difference(par_jour_h.index))
    ecarts_hebdo, n_hebdo = comparer(jour.resample("W-SUN").sum(), lus["hebdomadaire"].set_index("datum")[codes],
                                     "hebdomadaire")
    mensuel = lus["mensuel"].set_index("datum")[codes]
    mensuel.index = mensuel.index.to_period("M").to_timestamp()
    ecarts_mensuel, n_mensuel = comparer(jour.resample("MS").sum(), mensuel, "mensuel")
    heures_par_jour = horaire.groupby("Date").size()
    incomplets = heures_par_jour[heures_par_jour != 24]
    calendrier = pd.date_range(jour.index.min(), jour.index.max())
    manquants = calendrier.difference(jour.index)

    complets = jour.index.intersection(heures_par_jour[heures_par_jour == 24].index)
    jc = jour.loc[complets]
    total = jc.sum(axis=1)
    faibles = total[total < SEUIL_FAIBLE * total.median()]
    q1, q3 = jc.quantile(0.25), jc.quantile(0.75)
    pics = (jc > q3 + 3 * (q3 - q1)).sum()

    # Mois complets (tous les jours présents et complets) recalculés depuis le journalier
    nb = pd.Series(1, index=complets).resample("MS").sum()
    mois = jc.resample("MS").sum()
    mois = mois[nb.reindex(mois.index).fillna(0).to_numpy() == mois.index.days_in_month]

    mensuel_ecarte = not ecarts_mensuel.empty
    utilisation = {
        "Journalier": "Référence : statistiques, prévision, outil de sélection",
        "Horaire": "Profils par heure ; repère des jours complets (24 heures)",
        "Hebdomadaire": "Contrôle de cohérence uniquement",
        "Mensuel": "Écarté : diverge du journalier" if mensuel_ecarte else "Contrôle de cohérence uniquement",
    }
    etat_exports = pd.DataFrame(etats).set_index("Export")
    etat_exports["Utilisation"] = [utilisation[e] for e in etat_exports.index]

    def conclusion(n_ecarts, ok, ko):
        return ok if n_ecarts == 0 else ko

    controles = pd.DataFrame([
        ("Horaire = journalier (somme par jour)", len(commun), ecarts_h + seulement,
         conclusion(ecarts_h + seulement, "Concordants", "Divergents")),
        ("Hebdomadaire = journalier (semaines lundi → dimanche)", n_hebdo, len(ecarts_hebdo),
         conclusion(len(ecarts_hebdo), "Concordants", "Divergents")),
        ("Mensuel = journalier (somme par mois)", n_mensuel, len(ecarts_mensuel),
         conclusion(len(ecarts_mensuel), "Concordants",
                    f"Divergents sur {ecarts_mensuel['Période'].nunique() if mensuel_ecarte else 0} mois : export écarté")),
        ("Jours manquants dans le calendrier journalier", len(calendrier), len(manquants),
         conclusion(len(manquants), "Calendrier continu", ", ".join(f"{d:%d/%m/%Y}" for d in manquants[:10]))),
        ("Jours incomplets dans l'horaire (moins de 24 heures)", len(heures_par_jour), len(incomplets),
         ", ".join(f"{d:%d/%m/%Y} ({n} h)" for d, n in incomplets.items()) or "Aucun"),
        ("Jours de très faible activité (< 15 % d'un jour médian)", len(complets), len(faibles),
         ", ".join(f"{d:%d/%m/%Y}" for d in faibles.index) or "Aucun"),
        ("Valeurs atypiques (> 3e quartile + 3 écarts interquartiles)", len(complets) * len(codes), int(pics.sum()),
         ", ".join(f"{c} {n}" for c, n in pics.items() if n) or "Aucune"),
    ], columns=["Contrôle", "Éléments comparés", "Écarts trouvés", "Conclusion"]).set_index("Contrôle")

    exclus_mois = [p for p in jour.resample("MS").sum().index if p not in mois.index]
    decisions = [
        "Source de référence : l'export journalier, qui concorde exactement avec l'horaire et l'hebdomadaire.",
        (f"Export mensuel écarté : il diverge du journalier sur {ecarts_mensuel['Période'].nunique()} mois "
         "(détail plus bas) ; les ventes mensuelles sont recalculées à partir du journalier.")
        if mensuel_ecarte else "Export mensuel concordant : conservé pour contrôle.",
        (f"Statistiques par jour calculées sur les {len(complets)} jours complets ; écartés : "
         + (", ".join(f"{d:%d/%m/%Y}" for d in incomplets.index) or "aucun") + " (journées incomplètes)."),
        (f"Analyses mensuelles et prévision sur les {len(mois)} mois complets "
         f"({mois.index.min():%m/%Y} → {mois.index.max():%m/%Y}) ; mois incomplets écartés : "
         + ", ".join(f"{p:%m/%Y}" for p in exclus_mois) + "."),
        "Colonnes Year, Month, Hour et Weekday Name du journalier non utilisées (recalculées ; « Hour » n'y a pas de sens).",
        ("Quantités conservées telles qu'enregistrées, décimales comprises (unité non documentée). Un premier import "
         "Power Query les arrondissait pour M01AB, N05B, N05C, R03 et R06 : le type « nombre entier » avait été "
         "déduit des premières lignes. Le code relit les exports bruts."),
        "Jours de très faible activité et valeurs atypiques conservés (rien n'indique une erreur de saisie), mais signalés.",
        "Valeurs négatives ou illisibles : écartées si un nouvel export en contient (comptées dans le tableau ci-dessus).",
    ]

    heures_completes = horaire[horaire["Date"].isin(complets)]
    return Donnees(codes=codes, exports=lus, etat_exports=etat_exports, controles=controles,
                   ecarts_mensuel=ecarts_mensuel, decisions=decisions, jours=jour, jours_complets=jc,
                   heures=heures_completes, mois=mois, jours_faibles=faibles, pics=pics)
