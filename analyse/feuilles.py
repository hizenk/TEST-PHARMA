"""Écriture des onglets d'analyse : Synthèse, Statistiques, Prévision, Source externe, État des données."""
import math
from datetime import datetime

import pandas as pd

from .donnees import LIBELLES, NOMS, nom
from .externe import SEUIL_EPIDEMIE, SOURCE, URL
from .mise_en_forme import (F_DATE, F_DATE_HEURE, F_ENTIER, F_INDICE, F_PCT, F_PCT1, F_QTE, F_QTE1, GRIS, ORANGE_PALE,
                            VERT, VERT_BARRE, VERT_FONCE, VERT_PALE, barres, bloc, echelle_divergente, ecrire,
                            titre_feuille)
from .prevision import METHODES, N_TEST, N_VALIDATION, SEUIL_OUI
from .statistiques import MOIS

PALETTE = ["#2A78D6", "#EB6834", "#1BAF7A"]  # 3 premières couleurs catégorielles (lisibles par les daltoniens)


# ---------------------------------------------------------------------- petits utilitaires de texte
def fr(x, decimales=0):
    texte = f"{x:,.{decimales}f}".replace(",", " ").replace(".", ",")
    return texte


def pct(x, signe=True):
    return f"{'+' if signe and x > 0 else ''}{fr(100 * x)} %"


def liste(elements):
    elements = list(elements)
    if not elements:
        return "aucun"
    return elements[0] if len(elements) == 1 else ", ".join(elements[:-1]) + " et " + elements[-1]


def mois_txt(ts):
    return f"{MOIS[ts.month - 1].lower()} {ts.year}"


def lignes_texte(ws, st, ligne, textes, col=1, derniere_col=12, puce="•", fmt=None):
    """Paragraphes sur toute la largeur, hauteur de ligne adaptée à la longueur du texte."""
    largeur = 11 * (derniere_col - col + 1)
    for texte in textes:
        contenu = f"{puce} {texte}" if puce else texte
        ws.merge_range(ligne, col, ligne, derniere_col, contenu, fmt or st(text_wrap=True, valign="top"))
        ws.set_row(ligne, 15 * max(1, math.ceil(len(contenu) / (largeur * 1.15))) + 2)
        ligne += 1
    return ligne + 1


def entete(ws, st, ligne, texte, col=1, derniere_col=12):
    ws.merge_range(ligne, col, ligne, derniere_col, texte, st.entete)
    return ligne + 1


# ---------------------------------------------------------------------- feuilles de données
def ecrire_export(ws, st, df, nom_table):
    """Recopie un export (dates en vraies dates) sous forme de tableau Excel ; renvoie {colonne: n°}."""
    colonnes = list(df.columns)
    avec_heure = {c: pd.api.types.is_datetime64_any_dtype(df[c]) and (df[c].dt.hour != 0).any() for c in colonnes}
    fmt_date = {c: st(num_format=F_DATE_HEURE if avec_heure[c] else F_DATE) for c in colonnes}
    for j, c in enumerate(colonnes):
        est_date = pd.api.types.is_datetime64_any_dtype(df[c])
        for i, v in enumerate(df[c].tolist(), start=1):
            if est_date:
                if not pd.isna(v):
                    ws.write_datetime(i, j, v.to_pydatetime(), fmt_date[c])
            else:
                ecrire(ws, i, j, v)
        ws.set_column(j, j, 17 if avec_heure[c] else (12 if est_date else 11))
    ws.add_table(0, 0, len(df), len(colonnes) - 1, {"name": nom_table, "style": "Table Style Medium 7",
                                                    "columns": [{"header": str(c)} for c in colonnes]})
    ws.freeze_panes(1, 0)
    return {c: j for j, c in enumerate(colonnes)}


# ---------------------------------------------------------------------- Synthèse
def synthese(ws, st, d, stats, prev, ext, genere_le):
    ws.set_column(1, 12, 11)
    ligne = titre_feuille(ws, st, "Analyse des ventes de la pharmacie : synthèse",
                          f"Ventes du {d.premier_jour:%d/%m/%Y} au {d.dernier_jour:%d/%m/%Y} · {len(d.codes)} groupes "
                          f"de médicaments · classeur généré le {genere_le:%d/%m/%Y à %H:%M} par creer_classeur.py "
                          "à partir des exports", derniere_col=12)

    vue = stats["vue"]
    parts = vue["Part du total"].sort_values(ascending=False)
    res = prev["resultats"]
    oui = [c for c in res.index if res.at[c, "Prévisible ?"] == "Oui"]

    # Chiffres clés
    tuiles = [
        ("Jours analysés", fr(len(d.jours_complets)), f"jours complets sur {len(d.jours)}"),
        ("Quantité vendue", fr(vue["Total"].sum()), "unités enregistrées en caisse"),
        ("Premier groupe", parts.index[0].split(" · ")[1], f"{pct(parts.iloc[0], False)} des quantités"),
        ("Prévision du mois suivant", f"{len(oui)} groupes sur {len(res)}", "battent nettement la prévision naïve"),
    ]
    for k, (libelle, valeur, sous) in enumerate(tuiles):
        c0 = 1 + 3 * k
        tuile = dict(bg_color=VERT_PALE)
        ws.merge_range(ligne, c0, ligne, c0 + 2, libelle, st(bold=True, font_size=9, font_color=GRIS, indent=1, **tuile))
        ws.merge_range(ligne + 1, c0, ligne + 2, c0 + 2, valeur,
                       st(bold=True, font_size=18, font_color=VERT_FONCE, valign="vcenter", indent=1, **tuile))
        ws.merge_range(ligne + 3, c0, ligne + 3, c0 + 2, sous, st(font_size=9, font_color=GRIS, indent=1, **tuile))
    ligne += 5

    # Ce qu'il faut retenir
    annee = stats["annee"]
    evol = annee.columns[-1]
    hausse, baisse = annee[evol].idxmax(), annee[evol].idxmin()
    saison = stats["saison"]
    pics = [f"{g.split(' · ')[1]} en {saison.at[g, 'Mois le plus fort'].lower()} (×{fr(saison.loc[g, MOIS].max(), 2)})"
            for g in saison.index if saison.loc[g, MOIS].max() >= 1.3]
    semaine = stats["semaine"]
    jours_cols = [c for c in semaine.columns if c not in ("Jour le plus fort", "Jour le plus faible")]
    creux = [f"{g.split(' · ')[1]} le {semaine.at[g, 'Jour le plus faible'].lower()} (×{fr(semaine.loc[g, jours_cols].min(), 2)})"
             for g in semaine.index if semaine.loc[g, jours_cols].min() <= 0.8]
    week_end = semaine[["Samedi", "Dimanche"]].mean(axis=1)
    forts_we = [f"{g.split(' · ')[1]} (×{fr(week_end[g], 2)})" for g in week_end.index if week_end[g] >= 1.08]
    heures = stats["heures"].loc["Tous les groupes"]
    tranche = heures.idxmax()
    effets = ext["effets"]
    sensibles = effets[effets["Écart en épidémie, à saison égale (nov. → mars)"] >= 0.10]
    sensibles = sensibles.sort_values("Écart en épidémie, à saison égale (nov. → mars)", ascending=False)
    grippe_txt = liste(f"{NOMS.get(c, c)} ({c}) {pct(sensibles.at[c, 'Écart en épidémie, à saison égale (nov. → mars)'])}"
                       for c in sensibles.index)
    non = [c for c in res.index if res.at[c, "Prévisible ?"] != "Oui"]
    mensuel = d.controles.loc["Mensuel = journalier (somme par mois)"]

    ligne = entete(ws, st, ligne, "Ce qu'il faut retenir")
    retenir = [
        (f"Données fiables après contrôle : les exports horaire, journalier et hebdomadaire concordent ; l'export "
         + (f"mensuel diverge sur {d.ecarts_mensuel['Période'].nunique()} mois et a été écarté (ventes mensuelles "
            "recalculées depuis le journalier)." if not d.ecarts_mensuel.empty else "mensuel aussi.")
         + " Voir l'onglet État des données."),
        (f"Poids des groupes : {liste(f'{p.split(chr(183))[1].strip()} {pct(v, False)}' for p, v in parts.head(3).items())} "
         "des quantités vendues."),
        (f"Tendance {evol.replace('Évolution ', '')} (vente moyenne par jour) : la plus forte hausse pour "
         f"{hausse.split(' · ')[1]} ({pct(annee.at[hausse, evol])}), la plus forte baisse pour "
         f"{baisse.split(' · ')[1]} ({pct(annee.at[baisse, evol])})."),
        f"Saisonnalité marquée : pic des ventes pour {liste(pics)} par rapport à la moyenne de l'année.",
        ("Jours de la semaine : "
         + (f"ventes plus fortes le week-end pour {liste(forts_we)}" if forts_we else "pas d'effet week-end net")
         + (f" ; creux marqué pour {liste(creux)}." if creux else ".")),
        f"Heures : la tranche {tranche} concentre {pct(heures.max(), False)} des ventes de la journée.",
        (f"Environnement : pendant les semaines d'épidémie de grippe (au moins {SEUIL_EPIDEMIE} cas pour 100 000 "
         f"habitants, réseau Sentinelles), et à saison égale, les ventes augmentent pour {grippe_txt}."),
        (f"Prévision du mois suivant : oui pour {liste(f'{NOMS.get(c, c)} ({c})' for c in oui)} ; "
         f"pas mieux que la prévision naïve pour {liste(f'{NOMS.get(c, c)} ({c})' for c in non)}. "
         f"Réponse : « pour certains médicaments seulement »." if oui and non else
         f"Prévision du mois suivant : {'oui pour tous les groupes' if oui else 'non, pas mieux que la prévision naïve'}."),
    ]
    ligne = lignes_texte(ws, st, ligne, [f"{i}. {t}" for i, t in enumerate(retenir, 1)], puce="")

    # Recommandations
    def mois_forts(g, seuil=1.2):
        return [m.lower() for m in MOIS if saison.at[g, m] >= seuil]

    achats = [
        "Anticiper les pics saisonniers : " + liste(
            f"{g.split(' · ')[1]} ({', '.join(mois_forts(g))})" for g in saison.index if mois_forts(g)) + ".",
        (f"Suivre le bulletin hebdomadaire Sentinelles : dès le passage du seuil épidémique, renforcer les commandes de "
         f"{liste(f'{NOMS.get(c, c)} ({c})' for c in sensibles.index)}."),
        "Calculer les quantités à commander avec le bloc ⑤ « Tableau des achats » de l'onglet Tableau de bord.",
        (f"Pour {liste(oui)}, s'appuyer sur la prévision du mois suivant (onglet Prévision) ; "
         f"pour les autres groupes, la moyenne récente reste la meilleure référence."),
    ]
    proprietaire = [
        (f"Renforcer l'équipe sur la tranche {tranche} ({pct(heures.max(), False)} des ventes)"
         + (f" et le week-end ({liste(g.split(' (')[0] for g in forts_we)})." if forts_we else ".")),
        ("Une part des variations vient de l'environnement (épidémies de grippe, saisons) : elle se prévoit et se "
         "prépare ; les écarts entre jours et heures relèvent du fonctionnement de la pharmacie."),
        ("Ne pas transmettre l'export au laboratoire en l'état : les groupes N05B et N05C concernent l'anxiété et le "
         "sommeil, données de santé (voir la note écrite, partie 2)."),
    ]
    manager = [
        ("Analyse reproductible : la commande python creer_classeur.py régénère tout le classeur à partir des "
         "quatre exports, contrôles compris."),
        (f"Mission de suivi : mise à jour mensuelle des exports, alertes sur la qualité des données, prévision "
         f"pour {len(oui)} groupes sur {len(res)}."),
        "Point de vigilance : les quantités n'ont pas d'unité documentée et aucun prix ; à clarifier avec le client.",
    ]
    ligne = entete(ws, st, ligne, "Recommandations")
    for titre, textes in (("Pour le pharmacien chargé des achats", achats), ("Pour le propriétaire", proprietaire),
                          ("Pour le manager", manager)):
        ws.merge_range(ligne, 1, ligne, 12, titre, st(bold=True, font_color=VERT_FONCE, bottom=1, bottom_color=VERT_BARRE))
        ligne = lignes_texte(ws, st, ligne + 1, textes)

    # Contenu du classeur
    ligne = entete(ws, st, ligne, "Contenu du classeur")
    contenu = [
        ("Tableau de bord", "Outil de sélection : médicaments, période, regroupement, chiffres clés, achats (sans macro)"),
        ("Statistiques", "Statistiques descriptives : poids, tendance, saisonnalité, jours, heures, corrélations"),
        ("Prévision", f"Test de prévision du mois suivant sur {N_TEST} mois non utilisés, face à la prévision naïve"),
        ("Source externe", "Grippe (réseau Sentinelles) et ventes : ce qui relève de l'environnement"),
        ("État des données", "Quel export pour quoi, contrôles de cohérence, ce qui a été corrigé ou écarté"),
        ("Pharma_Ventes_Daily", "Données : les quatre exports (journalier, horaire, hebdomadaire, mensuel) et la grippe"),
    ]
    for feuille, description in contenu:
        ws.merge_range(ligne, 1, ligne, 3, "", st())
        ws.write_url(ligne, 1, f"internal:'{feuille}'!A1", st(font_color=VERT, underline=1, bold=True), feuille)
        ws.merge_range(ligne, 4, ligne, 12, description, st(font_color=GRIS))
        ligne += 1


# ---------------------------------------------------------------------- Statistiques
def statistiques(ws, st, d, stats):
    ws.set_column(1, 1, 26)
    ws.set_column(2, 17, 11)
    ligne = titre_feuille(ws, st, "Statistiques descriptives",
                          f"Calculées sur les {len(d.jours_complets)} jours complets de l'export journalier "
                          "(et horaire pour les heures) · quantités telles qu'enregistrées en caisse", derniere_col=15)
    vue = stats["vue"]
    ligne, p = bloc(ws, st, ligne, "Vue d'ensemble par groupe", vue, {
        "Total": F_ENTIER, "Part du total": F_PCT1, "Coef. de variation": F_INDICE, "Jours sans vente": F_PCT1,
        "Jour record": F_DATE, "*": F_QTE}, largeur_titre=15,
        note="Coefficient de variation = écart-type / moyenne : plus il est élevé, plus les ventes d'un jour à l'autre "
             "sont irrégulières (et difficiles à anticiper).")
    barres(ws, p, 3, len(vue))

    annee = stats["annee"]
    evol = annee.columns[-1]
    nb = stats["jours_par_annee"]
    ligne, p = bloc(ws, st, ligne, "Vente moyenne par jour, par année", annee, {evol: F_PCT1, "*": F_QTE},
                    largeur_titre=15,
                    note="Années retenues : au moins 300 jours complets ("
                         + ", ".join(f"{a} : {n} j" for a, n in nb.items()) + "). La moyenne par jour permet de "
                         "comparer des années de longueurs différentes.")
    echelle_divergente(ws, p, 1 + len(annee.columns), len(annee), 1, milieu=0)

    saison = stats["saison"]
    ligne, p_saison = bloc(ws, st, ligne, "Saisonnalité : indice mensuel (1 = moyenne de l'année)", saison,
                           {"*": F_INDICE}, largeur_titre=15,
                           note="Vente moyenne par jour du mois divisée par la vente moyenne par jour sur toute la "
                                "période. 1,50 = 50 % de plus qu'un jour moyen ; 0,50 = moitié moins.")
    echelle_divergente(ws, p_saison, 2, len(saison), 12)

    # place réservée au graphique de saisonnalité (inséré par graphe_saisonnalite)
    semaine = stats["semaine"]
    ligne_graphe = ligne
    ligne += 18
    ligne, p = bloc(ws, st, ligne, "Jour de la semaine : indice (1 = jour moyen)", semaine, {"*": F_INDICE},
                    largeur_titre=15)
    echelle_divergente(ws, p, 2, len(semaine), 7)
    heures = stats["heures"]
    ligne, p = bloc(ws, st, ligne, "Répartition des ventes de la journée par tranche horaire", heures, {"*": F_PCT1},
                    largeur_titre=15, note="Part des ventes de chaque groupe réalisée dans chaque tranche (jours complets).")
    ws.conditional_format(p, 2, p + len(heures) - 1, 1 + len(heures.columns),
                          {"type": "2_color_scale", "min_color": "#FFFFFF", "max_color": VERT_BARRE})
    corr = stats["correlations"]
    ligne, p = bloc(ws, st, ligne, "Corrélations entre groupes (ventes journalières)", corr, {"*": F_INDICE},
                    largeur_titre=15, note="1 = les deux groupes montent et descendent ensemble ; 0 = aucun lien. "
                                           "Une corrélation ne prouve pas une cause.")
    echelle_divergente(ws, p, 2, len(corr), len(corr.columns), milieu=0)
    return p_saison, ligne_graphe


def graphe_saisonnalite(wb, ws, d, p_saison, ligne_graphe):
    choix = [c for c in ("N02BE", "R03", "R06") if c in d.codes][:3]
    g = wb.add_chart({"type": "line"})
    for k, code in enumerate(choix):
        r = p_saison + d.codes.index(code)
        g.add_series({"name": ["Statistiques", r, 1], "categories": ["Statistiques", p_saison - 1, 2, p_saison - 1, 13],
                      "values": ["Statistiques", r, 2, r, 13], "line": {"color": PALETTE[k], "width": 2},
                      "marker": {"type": "circle", "size": 5, "fill": {"color": PALETTE[k]},
                                 "border": {"color": PALETTE[k]}}})
    g.set_title({"name": "Saisonnalité de trois groupes (indice mensuel)",
                 "name_font": {"size": 11, "bold": True, "color": VERT_FONCE}})
    g.set_y_axis({"min": 0, "num_format": "0.0", "major_gridlines": {"visible": True, "line": {"color": "#E0E0E0"}},
                  "num_font": {"size": 9, "color": GRIS}, "line": {"none": True}})
    g.set_x_axis({"num_font": {"size": 9, "color": GRIS}})
    g.set_legend({"position": "bottom"})
    g.set_chartarea({"border": {"color": "#C5E1A5"}})
    g.set_size({"width": 26 * 7 + 5 + 13 * 82, "height": 330})
    ws.insert_chart(ligne_graphe, 1, g, {"x_offset": 4, "y_offset": 4})


# ---------------------------------------------------------------------- Prévision
def prevision(ws, st, prev):
    ws.set_column(1, 1, 40)
    ws.set_column(2, 2, 36)
    ws.set_column(3, 9, 15)
    v0, v1 = prev["validation"]
    t0, t1 = prev["test"]
    ligne = titre_feuille(ws, st, "Peut-on prévoir les ventes du mois suivant ?",
                          f"Test sur {N_TEST} mois non utilisés pour construire les prévisions "
                          f"({mois_txt(t0)} → {mois_txt(t1)}), face à la prévision la plus simple", derniere_col=9)
    ligne = entete(ws, st, ligne, "Méthode", derniere_col=9)
    methode = [
        (f"Séries : ventes mensuelles de chaque groupe sur les {prev['nb_mois']} mois complets, recalculées depuis "
         "l'export journalier."),
        ("Origine glissante : pour prévoir un mois, chaque méthode n'utilise que les mois qui le précèdent, comme "
         "en situation réelle."),
        (f"Choix de la méthode sur {N_VALIDATION} mois de validation ({mois_txt(v0)} → {mois_txt(v1)}), puis "
         f"jugement sur les {N_TEST} mois de test suivants, jamais vus pendant le choix."),
        ("Référence : la prévision naïve (les ventes du mois précédent), celle que n'importe qui ferait sans analyse. "
         f"Verdict « Oui » si l'erreur moyenne baisse d'au moins {pct(SEUIL_OUI, False)} par rapport à elle."),
        "Méthodes comparées : " + "; ".join(METHODES.values()) + ".",
    ]
    ligne = lignes_texte(ws, st, ligne, methode, derniere_col=9)

    res = prev["resultats"].drop(columns=["_retenue", "_rmse"])
    res.index = [nom(c) for c in res.index]
    res.index.name = "Groupe"
    ligne, p = bloc(ws, st, ligne, "Résultats du test", res, {
        "Gain sur la prévision naïve": F_PCT, "Erreur relative moyenne (test)": F_PCT, "*": F_QTE1}, largeur_titre=9,
        note="Erreur moyenne = écart absolu moyen entre prévision et ventes réelles, en unités par mois.")
    col_verdict = 1 + list(res.columns).index("Prévisible ?") + 1
    for valeur, couleur in (("Oui", VERT_BARRE), ("Non", ORANGE_PALE), ("Gain faible", "#FFF2CC")):
        ws.conditional_format(p, col_verdict, p + len(res) - 1, col_verdict,
                              {"type": "cell", "criteria": "==", "value": f'"{valeur}"', "format": st(bg_color=couleur)})
    oui = [c for c in prev["resultats"].index if prev["resultats"].at[c, "Prévisible ?"] == "Oui"]
    reponse = ("pour certains médicaments seulement" if 0 < len(oui) < len(res) else ("oui" if oui else "non"))
    ligne = lignes_texte(ws, st, ligne - 1, [
        f"Réponse à la question du manager : {reponse} ({len(oui)} groupes sur {len(res)} : {liste(oui)})."],
        derniere_col=9, puce="", fmt=st(bold=True, font_color=VERT_FONCE, text_wrap=True))

    proch = prev["prochain"].copy()
    proch.index = [nom(c) for c in proch.index]
    proch.index.name = "Groupe"
    ligne, p = bloc(ws, st, ligne, f"Prévision pour {mois_txt(prev['prochain_mois'])} (mois qui suit le dernier mois complet)",
                    proch, {"*": F_ENTIER}, largeur_titre=9,
                    note="Fourchette indicative à 80 % tirée des erreurs observées pendant le test. Pour les groupes "
                         "non prévisibles, la prévision n'est pas plus fiable que la prévision naïve.")

    erreurs = prev["erreurs_test"].copy()
    erreurs.index.name = "Méthode (erreur moyenne sur le test)"
    ligne, p = bloc(ws, st, ligne, "Erreur moyenne de chaque méthode sur les mois de test", erreurs, {"*": F_QTE1},
                    largeur_titre=9)
    for j in range(len(erreurs.columns)):
        ws.conditional_format(p, 2 + j, p + len(erreurs) - 1, 2 + j,
                              {"type": "bottom", "value": 1, "format": st(bg_color=VERT_BARRE, bold=True)})

    detail = prev["detail"].copy()
    detail["Mois"] = detail["Mois"].dt.strftime("%m/%Y")
    detail = detail.set_index("Mois")
    bloc(ws, st, ligne, "Détail mois par mois sur la période de test", detail, {"*": F_QTE1}, largeur_titre=9)


# ---------------------------------------------------------------------- Source externe
def source_externe(wb, ws, st, ext, chemin, nb_lignes_donnees):
    ws.set_column(1, 1, 26)
    ws.set_column(2, 9, 15)
    ligne = titre_feuille(ws, st, "Source externe : la grippe et les ventes",
                          "Ce qui relève de l'environnement de la pharmacie et ce qui relève de son fonctionnement",
                          derniere_col=9)
    ligne = entete(ws, st, ligne, "Données utilisées", derniere_col=9)
    fichier = chemin.name
    infos = [
        f"Source : {SOURCE}. Données ouvertes, structurées, hebdomadaires.",
        f"Adresse : {URL} (copie dans donnees_externes/{fichier} ; « python creer_classeur.py --maj-externe » la met à jour).",
        (f"{ext['nb_semaines']} semaines reliées aux ventes (semaines de 7 jours complets). Semaine épidémique : au "
         f"moins {SEUIL_EPIDEMIE} cas pour 100 000 habitants ({ext['nb_epidemie']} semaines, dont "
         f"{ext['nb_epidemie_hiver']} entre novembre et mars)."),
        ("« À saison égale » : on ne compare que des semaines de novembre à mars, épidémiques ou non "
         f"({ext['nb_epidemie_hiver']} contre {ext['nb_hiver_hors']}), pour séparer l'effet de la grippe de celui de l'hiver."),
    ]
    ligne = lignes_texte(ws, st, ligne, infos, derniere_col=9)
    effets = ext["effets"].copy()
    effets.index = [nom(c) for c in effets.index]
    effets.index.name = "Groupe"
    ligne, p = bloc(ws, st, ligne, "Ventes hebdomadaires et épidémies de grippe", effets, {
        "Corrélation (Pearson)": F_INDICE, "Corrélation des rangs (Spearman)": F_INDICE,
        "Ventes / semaine hors épidémie": F_QTE1, "Ventes / semaine en épidémie": F_QTE1, "*": F_PCT}, largeur_titre=9)
    echelle_divergente(ws, p, 6, len(effets), 2, milieu=0)
    ligne = lignes_texte(ws, st, ligne - 1, [
        "Lecture : un écart de +40 % signifie 40 % de ventes en plus pendant les semaines d'épidémie. L'écart « à "
        "saison égale » est la part attribuable à la grippe ; le reste de l'écart brut vient de la saison."],
        derniere_col=9, puce="", fmt=st(font_size=9, font_color=GRIS, italic=True, text_wrap=True))

    # deux graphiques l'un sous l'autre, même axe du temps (pas de double axe)
    n = nb_lignes_donnees
    for k, (titre, col, couleur) in enumerate((("Grippe : cas pour 100 000 habitants par semaine", 2, PALETTE[1]),
                                               ("Paracétamol (N02BE) : ventes par semaine", None, PALETTE[0]))):
        g = wb.add_chart({"type": "line"})
        col_val = col if col is not None else ext["semaines"].columns.get_loc("N02BE") + 1
        g.add_series({"name": titre, "categories": ["Grippe_Sentinelles", 1, 0, n, 0],
                      "values": ["Grippe_Sentinelles", 1, col_val, n, col_val],
                      "line": {"color": couleur, "width": 1.5}})
        g.set_title({"name": titre, "name_font": {"size": 11, "bold": True, "color": VERT_FONCE}})
        g.set_legend({"none": True})
        g.set_x_axis({"date_axis": True, "num_format": "yyyy", "major_unit": 1, "major_unit_type": "years",
                      "num_font": {"size": 9, "color": GRIS}})
        g.set_y_axis({"min": 0, "num_font": {"size": 9, "color": GRIS}, "line": {"none": True},
                      "major_gridlines": {"visible": True, "line": {"color": "#E0E0E0"}}})
        g.set_chartarea({"border": {"color": "#C5E1A5"}})
        g.set_size({"width": 26 * 7 + 5 + 8 * 110, "height": 260})
        ws.insert_chart(ligne + 14 * k, 1, g, {"x_offset": 4, "y_offset": 4})


# ---------------------------------------------------------------------- État des données
def etat_donnees(ws, st, d):
    largeurs = [44, 24, 11, 17, 17, 11, 11, 11, 11, 46]  # colonnes B à K
    for j, l in enumerate(largeurs, start=1):
        ws.set_column(j, j, l)
    ligne = titre_feuille(ws, st, "État des données",
                          "Quel export pour quoi, ce qui a été contrôlé, corrigé ou écarté (recalculé à chaque exécution)",
                          derniere_col=10)
    texte = st(text_wrap=True, valign="top")
    nombre = st(num_format=F_ENTIER, valign="top")

    # Exports reçus : une colonne par information, l'utilisation dans la colonne large
    ligne = entete(ws, st, ligne, "Exports reçus et utilisation", derniere_col=10)
    etat = d.etat_exports
    titres = ["Export", *etat.columns]
    for j, t in enumerate(titres, start=1):
        ws.write(ligne, j, t, st.colonne if j in (1, 10) else st.colonne_droite)
    ws.set_row(ligne, 42)
    for i, (export, r) in enumerate(etat.iterrows(), start=1):
        ws.write(ligne + i, 1, export, st(bold=True, valign="top"))
        for j, col in enumerate(etat.columns, start=2):
            v = r[col]
            if isinstance(v, pd.Timestamp):
                ws.write_datetime(ligne + i, j, v.to_pydatetime(), st(num_format=F_DATE_HEURE, valign="top"))
            elif isinstance(v, str):
                ws.write(ligne + i, j, v, st(text_wrap=True, valign="top", indent=1))
            else:
                ecrire(ws, ligne + i, j, v, nombre)
    ligne += len(etat) + 2

    # Contrôles : conclusion fusionnée sur les colonnes E à K pour rester lisible
    ligne = entete(ws, st, ligne, "Contrôles de cohérence", derniere_col=10)
    for j, t in ((1, "Contrôle"), (2, "Éléments comparés"), (3, "Écarts trouvés")):
        ws.write(ligne, j, t, st.colonne if j == 1 else st.colonne_droite)
    ws.merge_range(ligne, 4, ligne, 10, "Conclusion", st.colonne)
    ws.set_row(ligne, 30)
    for i, (controle, r) in enumerate(d.controles.iterrows(), start=1):
        rr = ligne + i
        ws.write(rr, 1, controle, texte)
        ecrire(ws, rr, 2, r["Éléments comparés"], nombre)
        ecart = int(r["Écarts trouvés"])
        ws.write_number(rr, 3, ecart, st(num_format=F_ENTIER, valign="top", bold=ecart > 0,
                                         bg_color=ORANGE_PALE if ecart else "#FFFFFF"))
        conclusion = str(r["Conclusion"])
        ws.merge_range(rr, 4, rr, 10, conclusion, texte)
        ws.set_row(rr, 15 * max(1, math.ceil(max(len(conclusion) / 125, len(controle) / 44))) + 2)
    ligne += len(d.controles) + 2

    ligne = entete(ws, st, ligne, "Choix retenus : ce qui a été corrigé ou écarté", derniere_col=10)
    ligne = lignes_texte(ws, st, ligne, d.decisions, derniere_col=10)
    ligne = entete(ws, st, ligne, "Groupes de médicaments (codes ATC)", derniere_col=10)
    for code in d.codes:
        ws.write(ligne, 1, nom(code), st(bold=True))
        ws.merge_range(ligne, 2, ligne, 10, LIBELLES.get(code, ""), st(font_color=GRIS))
        ligne += 1
    ligne += 1
    if not d.ecarts_mensuel.empty:
        detail = d.ecarts_mensuel.copy()
        detail["Période"] = detail["Période"].dt.strftime("%m/%Y")
        detail = detail.set_index("Période")
        bloc(ws, st, ligne, "Détail des écarts : export mensuel / recalcul journalier", detail, {"*": F_QTE},
             largeur_titre=10)
