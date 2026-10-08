"""Peut-on prévoir les ventes du mois qui suit le dernier mois complet ? Test sur des mois non utilisés.

Protocole (« origine glissante ») : pour chaque mois testé, chaque méthode ne voit que les mois
précédents. La méthode de chaque groupe est choisie sur 12 mois de validation, puis jugée sur les
12 mois suivants (le test), qu'elle n'a jamais vus, face à la prévision la plus simple : refaire
les ventes du mois précédent (prévision naïve).
"""
import numpy as np
import pandas as pd

METHODES = {
    "naif": "Naïve : ventes du mois précédent",
    "saison_naif": "Même mois l'année précédente",
    "mm3": "Moyenne des 3 derniers mois",
    "mm12": "Moyenne des 12 derniers mois",
    "saison_niveau": "Même mois l'an dernier × évolution sur 12 mois",
    "regression": "Régression : tendance + effet du mois",
}
N_TEST, N_VALIDATION = 12, 12
SEUIL_OUI = 0.15  # « oui » si l'erreur baisse d'au moins 15 % par rapport à la prévision naïve
Z_80 = 1.2816     # intervalle indicatif à 80 % (erreurs supposées normales)


def prevoir(y, t, methode):
    """Prévision du mois t à partir des seuls mois 0 … t-1 (NaN si l'historique est trop court)."""
    if methode == "naif":
        return y[t - 1] if t >= 1 else np.nan
    if methode == "saison_naif":
        return y[t - 12] if t >= 12 else np.nan
    if methode == "mm3":
        return y[t - 3:t].mean() if t >= 3 else np.nan
    if methode == "mm12":
        return y[t - 12:t].mean() if t >= 12 else np.nan
    if methode == "saison_niveau":
        if t < 24 or y[t - 24:t - 12].mean() == 0:
            return np.nan
        return y[t - 12] * y[t - 12:t].mean() / y[t - 24:t - 12].mean()
    if methode == "regression":
        if t < 24:
            return np.nan
        mois = np.arange(t) % 12
        X = np.column_stack([np.ones(t), np.arange(t)] + [(mois == k).astype(float) for k in range(1, 12)])
        beta, *_ = np.linalg.lstsq(X, y[:t], rcond=None)
        x = np.r_[1.0, t, [float(t % 12 == k) for k in range(1, 12)]]
        return float(x @ beta)
    raise ValueError(methode)


def evaluer(mois):
    """mois : DataFrame des mois complets (index = 1er du mois, une colonne par code)."""
    n = len(mois)
    if n < N_TEST + N_VALIDATION + 24:
        raise SystemExit(f"Seulement {n} mois complets : il en faut {N_TEST + N_VALIDATION + 24} pour le test.")
    validation = range(n - N_TEST - N_VALIDATION, n - N_TEST)
    test = range(n - N_TEST, n)
    prochain_mois = mois.index[-1] + pd.offsets.MonthBegin(1)

    resultats, erreurs_test, detail, prochain = [], {}, [], []
    for code in mois.columns:
        y = mois[code].to_numpy(dtype=float)
        prev = {m: {t: prevoir(y, t, m) for t in [*validation, *test]} for m in METHODES}
        mae = lambda m, ts: np.mean([abs(prev[m][t] - y[t]) for t in ts])
        mae_valid = {m: mae(m, validation) for m in METHODES if not np.isnan(mae(m, validation))}
        retenue = min(mae_valid, key=mae_valid.get)
        mae_test = {m: mae(m, test) for m in METHODES}
        erreurs = np.array([prev[retenue][t] - y[t] for t in test])
        rmse = float(np.sqrt(np.mean(erreurs ** 2)))
        gain = 1 - mae_test[retenue] / mae_test["naif"] if mae_test["naif"] else 0.0
        mape = float(np.mean([abs(prev[retenue][t] - y[t]) / y[t] for t in test if y[t]]))
        verdict = "Oui" if gain >= SEUIL_OUI else ("Gain faible" if gain > 0 else "Non")
        resultats.append({"Code": code, "Méthode retenue (validation)": METHODES[retenue],
                          "Erreur moyenne naïve (test)": mae_test["naif"],
                          "Erreur moyenne méthode (test)": mae_test[retenue],
                          "Gain sur la prévision naïve": gain, "Erreur relative moyenne (test)": mape,
                          "Prévisible ?": verdict, "_retenue": retenue, "_rmse": rmse})
        erreurs_test[code] = {METHODES[m]: mae_test[m] for m in METHODES}
        for t in test:
            detail.append({"Mois": mois.index[t], "Code": code, "Réel": y[t], "Prévision naïve": prev["naif"][t],
                           "Prévision méthode retenue": prev[retenue][t],
                           "Erreur méthode retenue": prev[retenue][t] - y[t]})
        p = prevoir(y, n, retenue)
        prochain.append({"Code": code, "Méthode": METHODES[retenue], "Prévision": p,
                         "Fourchette basse (80 %)": max(0.0, p - Z_80 * rmse), "Fourchette haute (80 %)": p + Z_80 * rmse,
                         "Prévision naïve": y[-1], "Même mois l'an dernier": y[n - 12],
                         "Fiabilité (test)": verdict})
    res = pd.DataFrame(resultats).set_index("Code")
    return {
        "resultats": res,
        "erreurs_test": pd.DataFrame(erreurs_test),
        "detail": pd.DataFrame(detail),
        "prochain": pd.DataFrame(prochain).set_index("Code"),
        "prochain_mois": prochain_mois,
        "validation": (mois.index[validation[0]], mois.index[validation[-1]]),
        "test": (mois.index[test[0]], mois.index[test[-1]]),
        "nb_mois": n,
    }
