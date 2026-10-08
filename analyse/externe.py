"""Source publique : incidence hebdomadaire des syndromes grippaux en France (réseau Sentinelles).

Ce qui relève de l'environnement (épidémies de grippe) et ce qui relève de la pharmacie :
on compare les ventes hebdomadaires des semaines épidémiques à celles des autres semaines,
y compris « à saison égale » (semaines de novembre à mars seulement) pour ne pas confondre
l'effet de la grippe avec celui de l'hiver.
"""
import urllib.request
from pathlib import Path

import pandas as pd

URL = "https://www.sentiweb.fr/datasets/all/inc-3-PAY.csv"
SOURCE = "Réseau Sentinelles (Inserm, Sorbonne Université), indicateur 3 : syndromes grippaux, France entière"
SEUIL_EPIDEMIE = 150  # cas pour 100 000 habitants par semaine (seuil indicatif)
MOIS_HIVER = (11, 12, 1, 2, 3)


def telecharger(chemin):
    chemin = Path(chemin)
    chemin.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(URL, timeout=60) as reponse:
        chemin.write_bytes(reponse.read())
    return chemin


def charger(chemin):
    """Incidence hebdomadaire : index = lundi de la semaine ISO, colonne inc100 (cas / 100 000)."""
    brut = pd.read_csv(chemin, comment="#")
    brut["inc100"] = pd.to_numeric(brut["inc100"], errors="coerce")
    brut["Lundi"] = pd.to_datetime(brut["week"].astype(str) + "1", format="%G%V%u")
    return brut.dropna(subset=["inc100"]).set_index("Lundi").sort_index()[["week", "inc100"]]


def analyser(d, grippe):
    """Relie les ventes hebdomadaires (semaines de 7 jours complets) à l'incidence de la grippe."""
    jc = d.jours_complets
    lundis = jc.index - pd.to_timedelta(jc.index.dayofweek, unit="D")
    nb = jc.groupby(lundis).size()
    ventes = jc.groupby(lundis).sum()[nb == 7]
    semaines = ventes.join(grippe, how="inner")
    semaines.index.name = "Lundi"
    epidemie = semaines["inc100"] >= SEUIL_EPIDEMIE
    hiver = semaines.index.month.isin(MOIS_HIVER)

    lignes = []
    for code in d.codes:
        v = semaines[code]
        hausse = v[epidemie].mean() / v[~epidemie].mean() - 1
        hausse_hiver = v[epidemie & hiver].mean() / v[~epidemie & hiver].mean() - 1
        lignes.append({"Code": code,
                       "Corrélation (Pearson)": v.corr(semaines["inc100"]),
                       "Corrélation des rangs (Spearman)": v.rank().corr(semaines["inc100"].rank()),
                       "Ventes / semaine hors épidémie": v[~epidemie].mean(),
                       "Ventes / semaine en épidémie": v[epidemie].mean(),
                       "Écart en épidémie": hausse,
                       "Écart en épidémie, à saison égale (nov. → mars)": hausse_hiver})
    effets = pd.DataFrame(lignes).set_index("Code")
    table = pd.DataFrame({"Semaine Sentinelles": semaines["week"], "Grippe (cas / 100 000)": semaines["inc100"],
                          "Semaine épidémique": ["Oui" if e else "Non" for e in epidemie]},
                         index=semaines.index).join(semaines[d.codes])
    return {"effets": effets, "semaines": table, "nb_semaines": len(semaines),
            "nb_epidemie": int(epidemie.sum()), "nb_epidemie_hiver": int((epidemie & hiver).sum()),
            "nb_hiver_hors": int((~epidemie & hiver).sum())}
