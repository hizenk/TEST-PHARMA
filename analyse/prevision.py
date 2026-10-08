"""Peut-on prévoir les ventes du mois qui suit le dernier mois complet ? Test sur des mois non utilisés.

Neuf techniques, en trois familles (naïve, régression, lissage), sont comparées. Protocole
(« origine glissante ») : pour chaque mois testé, chaque technique ne voit que les mois précédents.
La technique de chaque groupe est choisie sur 12 mois de validation, puis jugée sur les 12 mois
suivants (le test), jamais vus pendant le choix, face à la prévision la plus simple :
Naïve · Dernier mois (refaire les ventes du mois précédent).
"""
import numpy as np
import pandas as pd

# clé -> (famille, technique, ce qu'elle capte, comment elle prévoit le mois suivant, adaptée quand)
METHODES = {
    "naive_saison": ("Naïve", "Même mois an dernier", "Saisonnalité",
                     "Reprend la valeur du même mois, un an avant", "Saison marquée, niveau stable"),
    "naive": ("Naïve", "Dernier mois", "Niveau récent",
              "Reprend la valeur du mois précédent", "Série lisse, sans saison"),
    "reg_tendance": ("Régression", "Tendance seule", "Tendance",
                     "Prolonge une droite ajustée sur tout l'historique", "Hausse ou baisse régulière, sans saison"),
    "reg_saison": ("Régression", "Saisonnalité seule", "Saisonnalité",
                   "Moyenne historique du mois visé", "Saison stable, pas de tendance"),
    "reg_tendance_saison": ("Régression", "Tendance + saisonnalité", "Les deux",
                            "Droite + écart habituel du mois visé", "Tendance régulière et saison stable"),
    "lissage_simple": ("Lissage", "Simple", "Niveau",
                       "Moyenne pondérée, les mois récents pèsent plus", "Ni tendance ni saison, niveau qui dérive"),
    "holt": ("Lissage", "Holt", "Niveau + tendance",
             "Niveau récent + pente récente", "Tendance qui change en cours de route"),
    "holt_winters_sans_tendance": ("Lissage", "Holt-Winters sans tendance", "Niveau + saisonnalité",
                                   "Niveau récent + écart saisonnier récent", "Saison marquée, niveau qui dérive"),
    "holt_winters": ("Lissage", "Holt-Winters", "Niveau + tendance + saisonnalité",
                     "Niveau + pente + écart saisonnier, tous récents", "Tendance et saison qui évoluent"),
}
REFERENCE = "naive"  # la prévision la plus simple, que n'importe qui ferait sans analyse
N_TEST, N_VALIDATION = 12, 12
SEUIL_OUI = 0.15  # « oui » si l'erreur baisse d'au moins 15 % par rapport à la prévision de référence
Z_80 = 1.2816     # intervalle indicatif à 80 % (erreurs supposées normales)
M = 12            # longueur de la saison (mois)

# Grilles des paramètres de lissage (choisis en minimisant l'erreur à un pas sur l'historique connu)
ALPHAS = np.round(np.arange(0.05, 1.0, 0.05), 2)
BETAS = np.array([0.01, 0.05, 0.1, 0.2, 0.3])
GAMMAS = np.array([0.01, 0.05, 0.1, 0.2, 0.3, 0.5])


def libelle(cle):
    famille, technique = METHODES[cle][:2]
    return f"{famille} · {technique}"


def tableau_methodes():
    return pd.DataFrame([{"Famille": f, "Technique": t, "Ce qu'elle capte": c,
                          "Comment elle prévoit le mois suivant": p, "Adaptée quand": a}
                         for f, t, c, p, a in METHODES.values()])


def _regression(y, t, tendance, saison):
    """Moindres carrés sur les mois 0 … t-1, prévision du mois t."""
    colonnes = [np.ones(t)]
    x = [1.0]
    if tendance:
        colonnes.append(np.arange(t, dtype=float))
        x.append(float(t))
    if saison:
        mois = np.arange(t) % M
        colonnes += [(mois == k).astype(float) for k in range(1, M)]
        x += [float(t % M == k) for k in range(1, M)]
    beta, *_ = np.linalg.lstsq(np.column_stack(colonnes), y[:t], rcond=None)
    return float(np.array(x) @ beta)


def _lissage(y, tendance, saison):
    """Lissage exponentiel additif (simple, Holt, Holt-Winters) ; prévision du point qui suit y.

    Toutes les combinaisons de paramètres de la grille sont calculées en même temps ; on garde
    celle dont l'erreur à un pas (somme des carrés) est la plus faible sur l'historique.
    """
    a = ALPHAS
    b = BETAS if tendance else np.array([0.0])
    g = GAMMAS if saison else np.array([0.0])
    A, B, G = (v.ravel() for v in np.meshgrid(a, b, g, indexing="ij"))
    n = len(y)
    if saison:  # initialisation sur les deux premières années, tendance retirée des écarts saisonniers
        base = y[:M].mean()
        pente0 = (y[M:2 * M].mean() - base) / M if tendance else 0.0
        milieu = (M - 1) / 2
        niveau = np.full(A.shape, base + pente0 * (M - 1 - milieu))  # niveau à la fin de la 1re année
        pente = np.full(A.shape, pente0)
        saisons = np.tile(y[:M] - (base + pente0 * (np.arange(M) - milieu)), (len(A), 1))
        debut = M
    else:
        niveau = np.full(A.shape, y[0])
        h = min(M, n - 1)
        pente = np.full(A.shape, (y[h] - y[0]) / h if tendance else 0.0)
        saisons = None
        debut = 1
    sse = np.zeros(A.shape)
    for t in range(debut, n):
        s = saisons[:, t % M] if saison else 0.0
        erreur = y[t] - (niveau + pente + s)
        sse += erreur ** 2
        nouveau = A * (y[t] - s) + (1 - A) * (niveau + pente)
        if tendance:
            pente = B * (nouveau - niveau) + (1 - B) * pente
        if saison:
            saisons[:, t % M] = G * (y[t] - nouveau) + (1 - G) * s
        niveau = nouveau
    k = int(np.argmin(sse))
    return float(niveau[k] + pente[k] + (saisons[k, n % M] if saison else 0.0))


def prevoir(y, t, methode):
    """Prévision du mois t à partir des seuls mois 0 … t-1 (NaN si l'historique est trop court)."""
    if methode == "naive":
        return y[t - 1] if t >= 1 else np.nan
    if methode == "naive_saison":
        return y[t - M] if t >= M else np.nan
    if methode == "reg_tendance":
        return _regression(y, t, True, False) if t >= M else np.nan
    if methode == "reg_saison":  # = moyenne historique du mois visé
        return float(y[t % M:t:M].mean()) if t >= M else np.nan
    if methode == "reg_tendance_saison":
        return _regression(y, t, True, True) if t >= 2 * M else np.nan
    if methode == "lissage_simple":
        return _lissage(y[:t], False, False) if t >= M else np.nan
    if methode == "holt":
        return _lissage(y[:t], True, False) if t >= M else np.nan
    if methode == "holt_winters_sans_tendance":
        return _lissage(y[:t], False, True) if t >= 2 * M else np.nan
    if methode == "holt_winters":
        return _lissage(y[:t], True, True) if t >= 2 * M else np.nan
    raise ValueError(methode)


def evaluer(mois):
    """mois : DataFrame des mois complets (index = 1er du mois, une colonne par code)."""
    n = len(mois)
    if n < N_TEST + N_VALIDATION + 2 * M:
        raise SystemExit(f"Seulement {n} mois complets : il en faut {N_TEST + N_VALIDATION + 2 * M} pour le test.")
    validation = range(n - N_TEST - N_VALIDATION, n - N_TEST)
    test = range(n - N_TEST, n)
    prochain_mois = mois.index[-1] + pd.offsets.MonthBegin(1)
    ref = libelle(REFERENCE)

    resultats, erreurs_test, detail, prochain = [], {}, [], []
    for code in mois.columns:
        y = mois[code].to_numpy(dtype=float)
        prev = {m: {t: prevoir(y, t, m) for t in [*validation, *test]} for m in METHODES}

        def mae(m, ts):
            return float(np.mean([abs(prev[m][t] - y[t]) for t in ts]))

        mae_valid = {m: mae(m, validation) for m in METHODES if not np.isnan(mae(m, validation))}
        retenue = min(mae_valid, key=mae_valid.get)
        mae_test = {m: mae(m, test) for m in METHODES}
        erreurs = np.array([prev[retenue][t] - y[t] for t in test])
        rmse = float(np.sqrt(np.mean(erreurs ** 2)))
        gain = 1 - mae_test[retenue] / mae_test[REFERENCE] if mae_test[REFERENCE] else 0.0
        mape = float(np.mean([abs(prev[retenue][t] - y[t]) / y[t] for t in test if y[t]]))
        verdict = "Oui" if gain >= SEUIL_OUI else ("Gain faible" if gain > 0 else "Non")
        resultats.append({"Code": code, "Famille retenue": METHODES[retenue][0],
                          "Technique retenue (validation)": METHODES[retenue][1],
                          f"Erreur moyenne {ref} (test)": mae_test[REFERENCE],
                          "Erreur moyenne technique retenue (test)": mae_test[retenue],
                          f"Gain sur {ref}": gain, "Erreur relative moyenne (test)": mape,
                          "Prévisible ?": verdict, "_retenue": retenue, "_rmse": rmse})
        erreurs_test[code] = {libelle(m): mae_test[m] for m in METHODES}
        for t in test:
            detail.append({"Mois": mois.index[t], "Code": code, "Réel": y[t], f"Prévision {ref}": prev[REFERENCE][t],
                           "Prévision technique retenue": prev[retenue][t],
                           "Erreur technique retenue": prev[retenue][t] - y[t]})
        p = prevoir(y, n, retenue)
        prochain.append({"Code": code, "Famille": METHODES[retenue][0], "Technique": METHODES[retenue][1],
                         "Prévision": p,
                         "Fourchette basse (80 %)": max(0.0, p - Z_80 * rmse), "Fourchette haute (80 %)": p + Z_80 * rmse,
                         ref: y[-1], "Naïve · Même mois an dernier": y[n - M], "Fiabilité (test)": verdict})
    erreurs = pd.DataFrame(erreurs_test)
    return {
        "resultats": pd.DataFrame(resultats).set_index("Code"),
        "erreurs_test": erreurs,
        "detail": pd.DataFrame(detail),
        "prochain": pd.DataFrame(prochain).set_index("Code"),
        "prochain_mois": prochain_mois,
        "validation": (mois.index[validation[0]], mois.index[validation[-1]]),
        "test": (mois.index[test[0]], mois.index[test[-1]]),
        "nb_mois": n,
    }
