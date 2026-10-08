"""Statistiques descriptives calculées sur les jours complets de l'export journalier (et horaire)."""
import pandas as pd

from .donnees import nom

JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août",
        "Septembre", "Octobre", "Novembre", "Décembre"]
TRANCHES = [(0, 8, "0 h – 8 h"), (8, 10, "8 h – 10 h"), (10, 12, "10 h – 12 h"), (12, 14, "12 h – 14 h"),
            (14, 16, "14 h – 16 h"), (16, 18, "16 h – 18 h"), (18, 20, "18 h – 20 h"), (20, 24, "20 h – 24 h")]


def _index_noms(df):
    df = df.copy()
    df.index = [nom(c) for c in df.index]
    df.index.name = "Groupe"
    return df


def vue_ensemble(d):
    jc = d.jours_complets
    total = jc.sum()
    vue = pd.DataFrame({
        "Total": total,
        "Part du total": total / total.sum(),
        "Moyenne / jour": jc.mean(),
        "Médiane / jour": jc.median(),
        "Écart-type / jour": jc.std(),
        "Coef. de variation": jc.std() / jc.mean(),
        "Minimum / jour": jc.min(),
        "Maximum / jour": jc.max(),
        "Jours sans vente": (jc == 0).mean(),
        "Jour record": jc.idxmax(),
    })
    return _index_noms(vue)


def par_annee(d):
    """Vente moyenne par jour, par année civile (seules les années avec au moins 300 jours complets)."""
    jc = d.jours_complets
    nb = jc.groupby(jc.index.year).size()
    annees = nb[nb >= 300].index
    moy = jc.groupby(jc.index.year).mean().loc[annees].T
    moy.columns = [str(a) for a in moy.columns]
    premiere, derniere = moy.columns[0], moy.columns[-1]
    moy[f"Évolution {premiere} → {derniere}"] = moy[derniere] / moy[premiere] - 1
    return _index_noms(moy), nb


def saisonnalite(d):
    """Indice mensuel : vente moyenne par jour du mois / vente moyenne par jour sur toute la période."""
    jc = d.jours_complets
    indice = (jc.groupby(jc.index.month).mean() / jc.mean()).T
    indice.columns = MOIS
    indice["Mois le plus fort"] = indice[MOIS].idxmax(axis=1)
    indice["Mois le plus faible"] = indice[MOIS].idxmin(axis=1)
    return _index_noms(indice)


def jours_semaine(d):
    jc = d.jours_complets
    indice = (jc.groupby(jc.index.dayofweek).mean() / jc.mean()).T
    indice.columns = JOURS
    indice["Jour le plus fort"] = indice[JOURS].idxmax(axis=1)
    indice["Jour le plus faible"] = indice[JOURS].idxmin(axis=1)
    return _index_noms(indice)


def tranches_horaires(d):
    """Part des ventes de chaque groupe réalisée dans chaque tranche horaire (jours complets)."""
    h = d.heures
    parts = {}
    for debut, fin, libelle in TRANCHES:
        parts[libelle] = h.loc[(h["Heure"] >= debut) & (h["Heure"] < fin), d.codes].sum()
    parts = pd.DataFrame(parts)
    parts = parts.div(parts.sum(axis=1), axis=0)
    tous = h[d.codes].sum(axis=1)
    ensemble = pd.Series({libelle: tous[(h["Heure"] >= debut) & (h["Heure"] < fin)].sum()
                          for debut, fin, libelle in TRANCHES})
    parts = _index_noms(parts)
    parts.loc["Tous les groupes"] = ensemble / ensemble.sum()
    return parts


def correlations(d):
    c = d.jours_complets.corr()
    c = _index_noms(c)
    c.columns = d.codes
    return c


def calculer(d):
    annee, jours_par_annee = par_annee(d)
    return {"vue": vue_ensemble(d), "annee": annee, "jours_par_annee": jours_par_annee,
            "saison": saisonnalite(d), "semaine": jours_semaine(d), "heures": tranches_horaires(d),
            "correlations": correlations(d)}
