"""Onglet « Tableau de bord » : l'outil de sélection du pharmacien, sans macro.

Le pharmacien choisit les médicaments (Oui / Non), une période et un regroupement ;
les chiffres clés, les graphiques, le tableau des achats et les résultats détaillés
se recalculent par formules. Les formules n'utilisent que des références de cellules
simples (pas de nom de tableau ni de nom défini) et les valeurs d'ouverture sont
enregistrées déjà calculées (fonction calculer, qui reproduit les formules).
"""
import math
from datetime import date, datetime, timedelta

from xlsxwriter.utility import xl_col_to_name

from .donnees import NOMS
from .mise_en_forme import (F_DATE, F_PCT, F_QTE, GRIS, VERT, VERT_BARRE, VERT_BORD, VERT_FONCE,
                            VERT_MOYEN, VERT_PALE, VERT_SAISIE)

FEUILLE, CALCULS, JOURS_F = "Tableau de bord", "Calculs", "Jours"
JOUR, HEURE = "Pharma_Ventes_Daily", "Pharma_Ventes_Hourly"
MAX_LIGNES = 400

REGROUPEMENTS = ["Jour", "Semaine", "Mois", "Année", "Jour de la semaine", "Heure"]
VALEURS = ["Somme", "Moyenne par jour"]
JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
JOURS_COURTS = ["lun.", "mar.", "mer.", "jeu.", "ven.", "sam.", "dim."]
MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août",
        "Septembre", "Octobre", "Novembre", "Décembre"]

# Valeurs proposées à l'ouverture (la période est fixée sur les 12 derniers mois complets des données)
DEFAUT = {"regroupement": "Mois", "valeur": "Somme", "jours_couvrir": 30, "marge": 0.2}


def periode_par_defaut(dates):
    """Les 12 derniers mois complets présents dans les données (ou toute la période si moins d'un an)."""
    dernier = max(dates)
    fin = dernier if (dernier + timedelta(days=1)).day == 1 else date(dernier.year, dernier.month, 1) - timedelta(days=1)
    debut = date(fin.year - 1, fin.month, 1) + timedelta(days=32)
    debut = date(debut.year, debut.month, 1)
    return max(debut, min(dates)), fin

LARGEUR_B, LARGEUR_C = 30, 12.5


def serie(d):
    """Date au format numéro de série Excel (valeur pré-calculée d'une formule qui renvoie une date)."""
    if isinstance(d, datetime):
        return (d - datetime(1899, 12, 30)).total_seconds() / 86400
    return (d - date(1899, 12, 30)).days



# ---------------------------------------------------------------------- Calcul des valeurs affichées
def calculer(jours, heures, codes, entrees):
    """Reproduit en Python les formules du classeur pour les valeurs d'ouverture (même logique, mêmes cas)."""
    debut, fin = entrees["debut"], entrees["fin"]
    g = REGROUPEMENTS.index(entrees["regroupement"]) + 1
    par_jour = entrees["valeur"] == "Moyenne par jour"
    inclus = entrees["inclus"]
    lundi = debut - timedelta(days=debut.isoweekday() - 1)

    if fin < debut:
        nb_lignes = 0
    else:
        nb_lignes = [(fin - debut).days + 1, (fin - lundi).days // 7 + 1,
                     (fin.year - debut.year) * 12 + fin.month - debut.month + 1,
                     fin.year - debut.year + 1, 7, 24][g - 1]
    nb_aff = min(nb_lignes, MAX_LIGNES)

    j = []  # une entrée par jour, comme la feuille Jours
    for d, valeurs in jours:
        dans = debut <= d <= fin
        if not dans:
            cle = 0
        else:
            cle = [(d - debut).days + 1, (d - lundi).days // 7 + 1,
                   (d.year - debut.year) * 12 + d.month - debut.month + 1,
                   d.year - debut.year + 1, d.isoweekday(), 0][g - 1]
        sel = sum(valeurs[c] for c in codes if inclus[c]) if dans else ""
        j.append({"date": d, "jsem": d.isoweekday(), "mois": d.month, "dans": int(dans), "cle": cle, "sel": sel,
                  "valeurs": valeurs})

    nb_jours = sum(x["dans"] for x in j)
    total = sum(x["sel"] for x in j if x["dans"])
    max_jour = max([x["sel"] for x in j if x["dans"]], default=0)
    meilleur = next((x["date"] for x in j if x["dans"] and x["sel"] == max_jour), None) if max_jour > 0 else None

    meds = {}
    for c in codes:
        tot = sum(x["valeurs"][c] for x in j if x["dans"])
        moy_jsem = []
        for k in range(1, 8):
            v = [x["valeurs"][c] for x in j if x["jsem"] == k]
            moy_jsem.append(sum(v) / len(v) if v else 0)
        moy_mois = []
        for m in range(1, 13):
            v = [x["valeurs"][c] for x in j if x["mois"] == m]
            moy_mois.append(sum(v) / len(v) if v else 0)
        meds[c] = {"total": tot, "moy": tot / nb_jours if nb_jours else 0, "jsem": moy_jsem, "mois": moy_mois,
                   "meilleur_jsem": JOURS[moy_jsem.index(max(moy_jsem))],
                   "meilleur_mois": MOIS[moy_mois.index(max(moy_mois))]}

    def dans_heures(dt):
        return debut <= dt.date() < fin + timedelta(days=1)

    lignes = []
    for i in range(1, MAX_LIGNES + 1):
        actif = i <= nb_aff
        if not actif:
            lignes.append({"actif": 0})
            continue
        if g == 1:
            debut_l = debut + timedelta(days=i - 1)
        elif g == 2:
            debut_l = lundi + timedelta(days=7 * (i - 1))
        elif g == 3:
            m0 = debut.month - 1 + i - 1
            debut_l = date(debut.year + m0 // 12, m0 % 12 + 1, 1)
        elif g == 4:
            debut_l = date(debut.year + i - 1, 1, 1)
        else:
            debut_l = i if g == 5 else i - 1
        libelle = [lambda: f"{JOURS_COURTS[debut_l.isoweekday() - 1]} {debut_l:%d/%m/%Y}",
                   lambda: f"Semaine du {debut_l:%d/%m/%Y}", lambda: f"{MOIS[debut_l.month - 1]} {debut_l.year}",
                   lambda: f"{debut_l.year}", lambda: JOURS[i - 1], lambda: f"{i - 1:02d}h"][g - 1]()
        if g <= 5:
            selection = [x for x in j if x["cle"] == i]
            nb = len(selection)
            sommes = {c: sum(x["valeurs"][c] for x in selection) for c in codes}
        else:
            selection = [h for h in heures if h[1] == i - 1 and dans_heures(h[0])]
            nb = len(selection)
            sommes = {c: sum(h[2][c] for h in selection) for c in codes}
        valeurs = {c: (sommes[c] / max(1, nb) if par_jour else sommes[c]) if inclus[c] else "" for c in codes}
        tot_ligne = sum(v for v in valeurs.values() if v != "")
        lignes.append({"actif": 1, "debut": debut_l, "libelle": libelle, "valeurs": valeurs, "total": tot_ligne,
                       "nb": nb, "x": debut_l if g <= 4 else "#N/A", "y": tot_ligne if g <= 4 else "#N/A"})

    profil_jsem = []
    for k in range(1, 8):
        v = [x["sel"] for x in j if x["dans"] and x["jsem"] == k]
        profil_jsem.append(sum(v) / len(v) if v else 0)
    profil_heure = []
    for h in range(24):
        v = [sum(vals[c] for c in codes if inclus[c]) for dt, heure, vals in heures if heure == h and dans_heures(dt)]
        profil_heure.append(sum(v) / len(v) if v else 0)

    return {"g": g, "lundi": lundi, "nb_lignes": nb_lignes, "nb_aff": nb_aff, "par_jour": int(par_jour),
            "nb_jours": nb_jours, "total": total, "max_jour": max_jour, "meilleur": meilleur, "jours": j,
            "meds": meds, "lignes": lignes, "profil_jsem": profil_jsem, "profil_heure": profil_heure}


# ---------------------------------------------------------------------- Écriture
def ajouter_tableau_de_bord(wb, fmt, ws, jours, heures, codes, cj, ch, nd, nh):
    """Remplit la feuille ws (déjà créée) et ajoute les onglets masqués Calculs et Jours.

    jours : [(date, {code: quantité})] dans l'ordre de la feuille journalière ;
    heures : [(datetime, heure, {code: quantité})] ; cj / ch : n° de colonne (0 = A) de
    datum, Hour et de chaque code dans les feuilles journalière et horaire ;
    nd / nh : dernière ligne Excel de ces feuilles.
    """
    n = len(codes)
    debut, fin = periode_par_defaut([d for d, _ in jours])
    DEF = dict(DEFAUT, debut=debut, fin=fin)
    entrees = dict(DEF, inclus={c: True for c in codes})
    v = calculer(jours, heures, codes, entrees)

    # Plages des données (références simples, valables dans tous les tableurs)
    def plage(feuille, col, fin_ligne):
        lettre = xl_col_to_name(col)
        return f"{feuille}!${lettre}$2:${lettre}${fin_ligne}"
    D_DATE = plage(JOUR, cj["datum"], nd)
    D = {c: plage(JOUR, cj[c], nd) for c in codes}
    H_DATE, H_HEURE = plage(HEURE, ch["datum"], nh), plage(HEURE, ch["Hour"], nh)
    H = {c: plage(HEURE, ch[c], nh) for c in codes}
    J = {col: f"{JOURS_F}!${col}$2:${col}${len(jours) + 1}" for col in "ABCDEF"}
    TDB = f"'{FEUILLE}'"

    calc = wb.add_worksheet(CALCULS)
    wj = wb.add_worksheet(JOURS_F)

    # ---------------- plan de l'onglet Tableau de bord (lignes Excel, 1 = première ligne)
    L_MED = 8                       # premier médicament
    L_TOT = L_MED + n               # total de la sélection
    L_PER = L_TOT + 2               # en-tête « Période et affichage »
    L_DU, L_AU, L_REG, L_VAL, L_DISPO = L_PER + 1, L_PER + 2, L_PER + 3, L_PER + 4, L_PER + 5
    L_MSG = max(L_DISPO + 1, 24)
    L_GRAPH = L_MSG + 2             # bande des graphiques
    L_ACH = L_GRAPH + 34            # bande du tableau des achats
    L_RES = L_ACH + 7 + n           # bande des résultats détaillés
    C_DEB, C_FIN = f"$C${L_DU}", f"$C${L_AU}"

    # Calculs : paramètres (lignes 2 à 11), médicaments (17…), périodes (28…), profils
    P = {"debut": "$B$2", "fin": "$B$3", "g": "$B$4", "lundi": "$B$5", "nb_lignes": "$B$6", "nb_aff": "$B$7",
         "par_jour": "$B$8", "nb_jours": "$B$9", "total": "$B$10", "max": "$B$11"}
    CP = {k: f"{CALCULS}!{ref}" for k, ref in P.items()}
    L_CMED = 17
    FLAG = {c: f"{CALCULS}!$B${L_CMED + i}" for i, c in enumerate(codes)}
    L_CPER = 28
    crit_heure = f'{H_DATE},">="&{CP["debut"]},{H_DATE},"<"&({CP["fin"]}+1)'

    def chaine_date(ref):
        return f'RIGHT("0"&DAY({ref}),2)&"/"&RIGHT("0"&MONTH({ref}),2)&"/"&YEAR({ref})'

    def choix(index, valeurs):
        return f"CHOOSE({index}," + ",".join(f'"{x}"' for x in valeurs) + ")"

    # ================================================================== Feuille Jours (masquée)
    for c, titre in enumerate(["Date", "Jour semaine", "Mois", "Dans la période", "Ligne du résultat",
                               "Ventes de la sélection"]):
        wj.write(0, c, titre, fmt(bold=True))
    somme_sel = lambda r: "+".join(f"{JOUR}!{xl_col_to_name(cj[c])}{r}*{FLAG[c]}" for c in codes)
    for k, x in enumerate(v["jours"]):
        r = k + 2
        wj.write_formula(r - 1, 0, f"={JOUR}!{xl_col_to_name(cj['datum'])}{r}", fmt(num_format=F_DATE),
                         serie(x["date"]))
        wj.write_formula(r - 1, 1, f"=WEEKDAY(A{r},2)", None, x["jsem"])
        wj.write_formula(r - 1, 2, f"=MONTH(A{r})", None, x["mois"])
        wj.write_formula(r - 1, 3, f"=IF(AND(A{r}>={CP['debut']},A{r}<={CP['fin']}),1,0)", None, x["dans"])
        wj.write_formula(r - 1, 4, f"=IF(D{r}=0,0,CHOOSE({CP['g']},A{r}-{CP['debut']}+1,"
                                   f"INT((A{r}-{CP['lundi']})/7)+1,"
                                   f"(YEAR(A{r})-YEAR({CP['debut']}))*12+MONTH(A{r})-MONTH({CP['debut']})+1,"
                                   f"YEAR(A{r})-YEAR({CP['debut']})+1,B{r},0))", None, x["cle"])
        wj.write_formula(r - 1, 5, f'=IF(D{r}=0,"",{somme_sel(r)})', None, x["sel"])
    wj.set_column(0, 5, 14)
    wj.hide()

    # ================================================================== Feuille Calculs (masquée)
    calc.write(0, 0, "Paramètre", fmt(bold=True))
    calc.write(0, 1, "Valeur", fmt(bold=True))
    reg = f"{TDB}!$C${L_REG}"
    num_reg = "".join(f'IF({reg}="{r}",{i},' for i, r in enumerate(REGROUPEMENTS, 1)) + "3" + ")" * 6
    parametres = [
        ("Début", f"={TDB}!{C_DEB}", serie(DEF["debut"]), F_DATE),
        ("Fin", f"={TDB}!{C_FIN}", serie(DEF["fin"]), F_DATE),
        ("N° de regroupement", "=" + num_reg, v["g"], None),
        ("Lundi de la 1re semaine", f"={P['debut']}-WEEKDAY({P['debut']},2)+1",
         serie(v["lundi"]), F_DATE),
        ("Nombre de lignes", f"=IF({P['fin']}<{P['debut']},0,CHOOSE({P['g']},{P['fin']}-{P['debut']}+1,"
                             f"INT(({P['fin']}-{P['lundi']})/7)+1,"
                             f"(YEAR({P['fin']})-YEAR({P['debut']}))*12+MONTH({P['fin']})-MONTH({P['debut']})+1,"
                             f"YEAR({P['fin']})-YEAR({P['debut']})+1,7,24))", v["nb_lignes"], None),
        ("Lignes affichées", f"=MIN({P['nb_lignes']},{MAX_LIGNES})", v["nb_aff"], None),
        ("Moyenne par jour ?", f'=IF({TDB}!$C${L_VAL}="Moyenne par jour",1,0)', v["par_jour"], None),
        ("Jours dans la période", f"=SUM({J['D']})", v["nb_jours"], None),
        ("Ventes de la sélection", f"=SUM({J['F']})", v["total"], None),
        ("Meilleur jour (quantité)", f"=MAX({J['F']})", v["max_jour"], None),
    ]
    for i, (libelle, formule, valeur, nf) in enumerate(parametres, start=1):
        calc.write(i, 0, libelle)
        calc.write_formula(i, 1, formule, fmt(num_format=nf) if nf else None, valeur)

    entetes_med = (["Code", "Inclus", "Total période", "Moyenne / jour"] + [f"Moy. {j_[:3]}." for j_ in JOURS]
                   + [f"Moy. {m[:4]}." for m in MOIS] + ["Jour le plus fort", "Mois le plus fort"])
    for c, titre in enumerate(entetes_med):
        calc.write(L_CMED - 2, c, titre, fmt(bold=True))
    for i, code in enumerate(codes):
        r = L_CMED + i
        m = v["meds"][code]
        calc.write(r - 1, 0, code)
        calc.write_formula(r - 1, 1, f'=IF({TDB}!$C${L_MED + i}="Oui",1,0)', None, 1)
        calc.write_formula(r - 1, 2, f"=SUMIF({J['D']},1,{D[code]})", None, m["total"])
        calc.write_formula(r - 1, 3, f"=IF({P['nb_jours']}>0,C{r}/{P['nb_jours']},0)", None, m["moy"])
        for k in range(7):
            calc.write_formula(r - 1, 4 + k, f"=IFERROR(AVERAGEIF({J['B']},{k + 1},{D[code]}),0)", None, m["jsem"][k])
        for k in range(12):
            calc.write_formula(r - 1, 11 + k, f"=IFERROR(AVERAGEIF({J['C']},{k + 1},{D[code]}),0)", None, m["mois"][k])
        calc.write_formula(r - 1, 23, f"={choix(f'MATCH(MAX(E{r}:K{r}),E{r}:K{r},0)', JOURS)}", None,
                           m["meilleur_jsem"])
        calc.write_formula(r - 1, 24, f"={choix(f'MATCH(MAX(L{r}:W{r}),L{r}:W{r},0)', MOIS)}", None,
                           m["meilleur_mois"])

    for c, titre in enumerate(["N°", "Affichée", "Début", "Libellé", "X graphique", "Y graphique"]):
        calc.write(L_CPER - 2, c, titre, fmt(bold=True))
    L_RES_DATA = L_RES + 2  # première ligne de données du tableau de résultats (onglet Tableau de bord)
    col_tot, col_nb = 2 + n, 3 + n
    l_tot, l_nb = xl_col_to_name(col_tot), xl_col_to_name(col_nb)
    for i in range(1, MAX_LIGNES + 1):
        r = L_CPER + i - 1
        ligne = v["lignes"][i - 1]
        rd = L_RES_DATA + i - 1
        actif = ligne["actif"]
        calc.write_number(r - 1, 0, i)
        calc.write_formula(r - 1, 1, f"=IF(A{r}<={P['nb_aff']},1,0)", None, actif)
        calc.write_formula(r - 1, 2, f'=IF(B{r}=0,"",CHOOSE({P["g"]},{P["debut"]}+A{r}-1,{P["lundi"]}+7*(A{r}-1),'
                                     f'DATE(YEAR({P["debut"]}),MONTH({P["debut"]})+A{r}-1,1),'
                                     f'DATE(YEAR({P["debut"]})+A{r}-1,1,1),A{r},A{r}-1))',
                           fmt(num_format=F_DATE),
                           (serie(ligne["debut"])
                            if isinstance(ligne.get("debut"), date) else ligne.get("debut", "")) if actif else "")
        libelle = (f'IF(B{r}=0,"",CHOOSE({P["g"]},'
                   f'{choix(f"WEEKDAY(C{r},2)", JOURS_COURTS)}&" "&{chaine_date(f"C{r}")},'
                   f'"Semaine du "&{chaine_date(f"C{r}")},'
                   f'{choix(f"MONTH(C{r})", MOIS)}&" "&YEAR(C{r}),'
                   f'""&YEAR(C{r}),{choix(f"A{r}", JOURS)},RIGHT("0"&(A{r}-1),2)&"h"))')
        calc.write_formula(r - 1, 3, "=" + libelle, None, ligne["libelle"] if actif else "")
        x = ligne["x"] if actif else "#N/A"
        calc.write_formula(r - 1, 4, f"=IF(AND(B{r}=1,{P['g']}<=4),C{r},NA())", fmt(num_format=F_DATE),
                           serie(x) if isinstance(x, date) else x)
        calc.write_formula(r - 1, 5, f"=IF(AND(B{r}=1,{P['g']}<=4),{TDB}!${l_tot}${rd},NA())", None,
                           ligne["y"] if actif else "#N/A")

    calc.write(L_CPER - 2, 7, "Jour", fmt(bold=True))
    calc.write(L_CPER - 2, 8, "Moyenne / jour", fmt(bold=True))
    for k in range(7):
        r = L_CPER + k
        calc.write(r - 1, 7, JOURS_COURTS[k].capitalize())
        calc.write_formula(r - 1, 8, f"=IFERROR(SUMIFS({J['F']},{J['B']},{k + 1},{J['D']},1)"
                                     f"/COUNTIFS({J['B']},{k + 1},{J['D']},1),0)", None, v["profil_jsem"][k])
    calc.write(L_CPER - 2, 10, "Heure", fmt(bold=True))
    calc.write(L_CPER - 2, 11, "Moyenne / jour", fmt(bold=True))
    for h in range(24):
        r = L_CPER + h
        somme = "+".join(f"$B${L_CMED + i}*SUMIFS({H[c]},{H_HEURE},{h},"
                         f'{H_DATE},">="&{P["debut"]},{H_DATE},"<"&({P["fin"]}+1))'
                         for i, c in enumerate(codes))
        calc.write(r - 1, 10, f"{h:02d}h")
        calc.write_formula(r - 1, 11, f'=IFERROR(({somme})/COUNTIFS({H_HEURE},{h},{H_DATE},">="&{P["debut"]},'
                                      f'{H_DATE},"<"&({P["fin"]}+1)),0)', None, v["profil_heure"][h])
    calc.set_column(0, 0, 24)
    calc.set_column(3, 3, 24)
    calc.hide()

    # ================================================================== Tableau de bord
    ws.hide_gridlines(2)
    ws.set_zoom(90)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, LARGEUR_B)
    ws.set_column(2, 12, LARGEUR_C)
    ws.set_column(13, 13, 2)
    ws.set_row(0, 8)
    ws.set_row(1, 26)
    ws.set_row(2, 26)

    titre_f = fmt(bold=True, font_size=20, font_color="#FFFFFF", bg_color=VERT_FONCE, valign="vcenter", indent=1)
    ws.merge_range("B2:M3", "Ventes de la pharmacie : tableau de bord", titre_f)
    ws.merge_range("B4:M4", "Modifiez les cellules vert clair : tout se met à jour automatiquement, sans macro.",
                   fmt(italic=True, font_color=GRIS, valign="vcenter"))

    entete_f = fmt(bold=True, font_color="#FFFFFF", bg_color=VERT, valign="vcenter", indent=1)
    col_f = fmt(bold=True, font_size=9, font_color=GRIS, bottom=1, bottom_color=VERT_BORD, valign="bottom")
    col_f_d = fmt(bold=True, font_size=9, font_color=GRIS, bottom=1, bottom_color=VERT_BORD, align="right",
                  text_wrap=True, valign="bottom")
    saisie = dict(bg_color=VERT_SAISIE, border=1, border_color=VERT_BORD, bold=True, font_color=VERT_FONCE,
                  align="center", valign="vcenter", locked=False)
    petit = fmt(font_size=9, font_color=GRIS)
    qte = fmt(num_format=F_QTE)

    # ---------------- ① Médicaments
    ws.merge_range(f"B6:E6", "① Médicaments", entete_f)
    for c, titre in zip("BCDE", ["Médicament", "Inclure", "Total période", "Part"]):
        ws.write(f"{c}7", titre, col_f if c == "B" else col_f_d)
    for i, code in enumerate(codes):
        r = L_MED + i
        m = v["meds"][code]
        ws.write(f"B{r}", f"{code} · {NOMS.get(code, code)}")
        ws.write(f"C{r}", "Oui", fmt(**saisie))
        ws.write_formula(f"D{r}", f"={CALCULS}!$C${L_CMED + i}", qte, m["total"])
        ws.write_formula(f"E{r}", f'=IF(AND({FLAG[code]}=1,{CP["total"]}>0),D{r}/{CP["total"]},"")',
                         fmt(num_format=F_PCT), m["total"] / v["total"] if v["total"] else "")
    tot_f = fmt(bold=True, top=1, top_color=VERT_BORD, num_format=F_QTE)
    ws.write(f"B{L_TOT}", "Total de la sélection", fmt(bold=True, top=1, top_color=VERT_BORD))
    ws.write(f"C{L_TOT}", "", tot_f)
    ws.write_formula(f"D{L_TOT}", f"={CP['total']}", tot_f, v["total"])
    ws.write(f"E{L_TOT}", "", tot_f)
    ws.data_validation(f"C{L_MED}:C{L_TOT - 1}", {"validate": "list", "source": ["Oui", "Non"],
                                                  "error_title": "Valeur non valide",
                                                  "error_message": "Choisissez Oui ou Non."})
    ws.conditional_format(f"B{L_MED}:E{L_TOT - 1}", {"type": "formula", "criteria": f'=$C{L_MED}="Non"',
                                                     "format": fmt(font_color="#9E9E9E")})
    ws.conditional_format(f"E{L_MED}:E{L_TOT - 1}", {"type": "data_bar", "bar_color": VERT_BARRE, "bar_solid": True,
                                                     "min_type": "num", "min_value": 0,
                                                     "max_type": "num", "max_value": 1})

    # ---------------- ② Période et affichage
    ws.merge_range(f"B{L_PER}:E{L_PER}", "② Période et affichage", entete_f)
    ws.write(f"B{L_DU}", "Du")
    ws.write(f"B{L_AU}", "Au")
    ws.write(f"B{L_REG}", "Regrouper par")
    ws.write(f"B{L_VAL}", "Valeur")
    date_saisie = fmt(num_format=F_DATE, **saisie)
    ws.write_number(f"C{L_DU}", serie(DEF["debut"]), date_saisie)
    ws.write_number(f"C{L_AU}", serie(DEF["fin"]), date_saisie)
    ws.merge_range(f"C{L_REG}:D{L_REG}", DEF["regroupement"], fmt(**saisie))
    ws.merge_range(f"C{L_VAL}:D{L_VAL}", DEF["valeur"], fmt(**saisie))
    ws.write(f"B{L_DISPO}", "Données disponibles", petit)
    ws.write_formula(f"C{L_DISPO}", f"=MIN({D_DATE})", fmt(font_size=9, font_color=GRIS, num_format=F_DATE),
                     serie(min(d for d, _ in jours)))
    ws.write_formula(f"D{L_DISPO}", f"=MAX({D_DATE})", fmt(font_size=9, font_color=GRIS, num_format=F_DATE),
                     serie(max(d for d, _ in jours)))
    ws.data_validation(f"C{L_DU}:C{L_AU}", {"validate": "date", "criteria": "between",
                                            "minimum": f"=$C${L_DISPO}", "maximum": f"=$D${L_DISPO}",
                                            "error_title": "Date hors des données",
                                            "error_message": "Choisissez une date comprise dans les données disponibles."})
    ws.data_validation(f"C{L_REG}", {"validate": "list", "source": REGROUPEMENTS,
                                     "error_title": "Valeur non valide", "error_message": "Choisissez dans la liste."})
    ws.data_validation(f"C{L_VAL}", {"validate": "list", "source": VALEURS,
                                     "error_title": "Valeur non valide", "error_message": "Choisissez dans la liste."})
    message = (f'=IF({C_FIN}<{C_DEB},"⚠ La date de fin est avant la date de début.",'
               f'IF(COUNTIF($C${L_MED}:$C${L_TOT - 1},"Oui")=0,"⚠ Aucun médicament inclus : mettez Oui devant au moins un médicament.",'
               f'IF({CP["nb_lignes"]}>{MAX_LIGNES},"⚠ "&{CP["nb_lignes"]}&" lignes : seules les {MAX_LIGNES} premières '
               f'sont affichées. Choisissez un regroupement plus large.","")))')
    ws.merge_range(f"B{L_MSG}:M{L_MSG}", "", fmt(bold=True, font_color="#C62828"))
    ws.write_formula(f"B{L_MSG}", message, fmt(bold=True, font_color="#C62828"), "")

    # ---------------- ③ Chiffres clés (tuiles) et mode d'emploi
    ws.merge_range("G6:M6", "③ Chiffres clés de la sélection", entete_f)
    tuile = dict(bg_color=VERT_PALE)
    tuiles = [
        ("G", 7, "Quantité vendue", f"={CP['total']}", "#,##0", v["total"],
         f'="sur "&{CP["nb_jours"]}&" jours"', f"sur {v['nb_jours']} jours"),
        ("K", 7, "Moyenne par jour", f"=IF({CP['nb_jours']}>0,{CP['total']}/{CP['nb_jours']},0)", "#,##0.0",
         v["total"] / v["nb_jours"] if v["nb_jours"] else 0, '="pour la sélection"', "pour la sélection"),
        ("G", 12, "Meilleur jour",
         f'=IF({CP["max"]}<=0,"—",INDEX({J["A"]},MATCH({CP["max"]},{J["F"]},0)))', F_DATE,
         serie(v["meilleur"]) if v["meilleur"] else "—",
         f'=IF({CP["max"]}<=0,"","avec "&FIXED({CP["max"]},0)&" ventes")',
         f"avec {v['max_jour']:,.0f} ventes".replace(",", " ")),
        ("K", 12, "Jours de données", f"={CP['nb_jours']}", "0", v["nb_jours"],
         f'="du "&{chaine_date(C_DEB)}&" au "&{chaine_date(C_FIN)}',
         f"du {DEF['debut']:%d/%m/%Y} au {DEF['fin']:%d/%m/%Y}"),
    ]
    for col, r, libelle, formule, nf, valeur, sous_formule, sous_valeur in tuiles:
        c0 = ord(col) - 65
        ws.merge_range(r - 1, c0, r - 1, c0 + 2, libelle, fmt(bold=True, font_size=9, font_color=GRIS, indent=1, **tuile))
        ws.merge_range(r, c0, r + 1, c0 + 2, "", fmt(**tuile))
        ws.write_formula(r, c0, formule, fmt(bold=True, font_size=20, font_color=VERT_FONCE,
                                                          num_format=nf, valign="vcenter", indent=1, **tuile), valeur)
        ws.merge_range(r + 2, c0, r + 2, c0 + 2, "", fmt(font_size=9, font_color=GRIS, indent=1, **tuile))
        ws.write_formula(r + 2, c0, sous_formule, fmt(font_size=9, font_color=GRIS, indent=1, **tuile), sous_valeur)
    ws.merge_range("G17:M17", "Mode d'emploi", entete_f)
    consignes = [
        "1. Mettez Oui ou Non devant chaque médicament.",
        "2. Saisissez les dates Du et Au (jj/mm/aaaa).",
        "3. Choisissez le regroupement et la valeur (somme ou moyenne par jour).",
        "4. Lisez les chiffres clés, les graphiques et les résultats plus bas.",
        "5. Tableau des achats : indiquez les jours à couvrir, la marge et vos stocks.",
        "Seules les cellules vert clair sont modifiables.",
    ]
    for k, texte in enumerate(consignes):
        ws.merge_range(17 + k, 6, 17 + k, 12, texte, fmt(font_size=10, font_color=GRIS, indent=1))

    # ---------------- ④ Graphiques (chacun ancré dans sa zone de cellules)
    ws.merge_range(f"B{L_GRAPH}:M{L_GRAPH}", "④ Graphiques de la sélection", entete_f)
    px_b, px_c = int(LARGEUR_B * 7 + 0.5) + 5, int(LARGEUR_C * 7 + 0.5) + 5

    def style_graphe(graphe, titre):
        graphe.set_title({"name": titre, "name_font": {"size": 11, "bold": True, "color": VERT_FONCE}})
        graphe.set_legend({"none": True})
        graphe.set_chartarea({"border": {"color": "#C5E1A5"}})
        graphe.set_y_axis({"min": 0, "num_format": "#,##0", "num_font": {"size": 9, "color": GRIS},
                           "major_gridlines": {"visible": True, "line": {"color": "#E0E0E0"}},
                           "line": {"none": True}})

    r0, r1 = L_CPER - 1, L_CPER + MAX_LIGNES - 2
    evolution = wb.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
    evolution.add_series({"name": "Sélection", "categories": [CALCULS, r0, 4, r1, 4], "values": [CALCULS, r0, 5, r1, 5],
                          "line": {"color": VERT, "width": 2},
                          "marker": {"type": "circle", "size": 5, "fill": {"color": VERT}, "border": {"color": VERT}}})
    style_graphe(evolution, "Évolution des ventes (regroupement par jour, semaine, mois ou année)")
    evolution.set_x_axis({"num_format": "mm/yyyy", "num_font": {"size": 9, "color": GRIS},
                          "major_gridlines": {"visible": False}, "line": {"color": "#BDBDBD"}})
    evolution.show_na_as_empty_cell()
    evolution.set_size({"width": px_b + 11 * px_c - 12, "height": 16 * 20 - 8})
    ws.insert_chart(f"B{L_GRAPH + 1}", evolution, {"x_offset": 6, "y_offset": 4})

    def colonnes_graphe(titre, col_cat, col_val, nb, largeur_px):
        g = wb.add_chart({"type": "column"})
        g.add_series({"categories": [CALCULS, r0, col_cat, r0 + nb - 1, col_cat],
                      "values": [CALCULS, r0, col_val, r0 + nb - 1, col_val],
                      "fill": {"color": VERT_MOYEN}, "border": {"none": True}, "gap": 60})
        style_graphe(g, titre)
        g.set_y_axis({"min": 0, "num_format": "#,##0.0", "num_font": {"size": 9, "color": GRIS},
                      "major_gridlines": {"visible": True, "line": {"color": "#E0E0E0"}}, "line": {"none": True}})
        g.set_x_axis({"num_font": {"size": 9, "color": GRIS}, "line": {"color": "#BDBDBD"}})
        g.set_size({"width": largeur_px, "height": 15 * 20 - 8})
        return g

    ws.insert_chart(f"B{L_GRAPH + 18}", colonnes_graphe("Ventes moyennes selon le jour de la semaine", 7, 8, 7,
                                                        px_b + 5 * px_c - 12), {"x_offset": 6, "y_offset": 4})
    ws.insert_chart(f"H{L_GRAPH + 18}", colonnes_graphe("Ventes moyennes par jour selon l'heure", 10, 11, 24,
                                                        6 * px_c - 12), {"x_offset": 6, "y_offset": 4})

    # ---------------- ⑤ Tableau des achats
    ws.merge_range(f"B{L_ACH}:M{L_ACH}", "⑤ Tableau des achats", entete_f)
    L_JC, L_MARGE, L_NOTE, L_ACH_T = L_ACH + 1, L_ACH + 2, L_ACH + 3, L_ACH + 4
    ws.write(f"B{L_JC}", "Jours à couvrir")
    ws.write_number(f"C{L_JC}", DEF["jours_couvrir"], fmt(num_format="0", **saisie))
    ws.write(f"D{L_JC}", "jours", petit)
    ws.write(f"B{L_MARGE}", "Marge de sécurité")
    ws.write_number(f"C{L_MARGE}", DEF["marge"], fmt(num_format="0%", **saisie))
    ws.data_validation(f"C{L_JC}", {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": 365,
                                    "error_title": "Valeur non valide", "error_message": "Entre 1 et 365 jours."})
    ws.data_validation(f"C{L_MARGE}", {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 2,
                                       "error_title": "Valeur non valide", "error_message": "Entre 0 % et 200 %."})
    note = (f'="Calcul : moyenne par jour du "&{chaine_date(C_DEB)}&" au "&{chaine_date(C_FIN)}'
            f'&" × jours à couvrir × (1 + marge) − stock actuel, arrondi à l\'unité supérieure."')
    ws.merge_range(f"B{L_NOTE}:M{L_NOTE}", "", petit)
    ws.write_formula(f"B{L_NOTE}", note, petit,
                     f"Calcul : moyenne par jour du {DEF['debut']:%d/%m/%Y} au {DEF['fin']:%d/%m/%Y} × jours "
                     f"à couvrir × (1 + marge) − stock actuel, arrondi à l'unité supérieure.")
    titres_achats = ["Médicament", "Moyenne / jour", None, "Marge de sécurité", "Stock actuel", "À commander",
                     "Jour le plus fort", "Mois le plus fort"]
    ws.set_row(L_ACH_T - 1, 30)
    for k, titre in enumerate(titres_achats):
        cellule = f"{chr(66 + k)}{L_ACH_T}"
        if titre is None:
            ws.write_formula(cellule, f'="Besoin sur "&$C${L_JC}&" jours"', col_f_d,
                             f"Besoin sur {DEF['jours_couvrir']} jours")
        else:
            ws.write(cellule, titre, col_f if k == 0 else col_f_d)
    for i, code in enumerate(codes):
        r = L_ACH_T + 1 + i
        m = v["meds"][code]
        besoin = m["moy"] * DEF["jours_couvrir"]
        marge = besoin * DEF["marge"]
        ws.write(f"B{r}", f"{code} · {NOMS.get(code, code)}")
        ws.write_formula(f"C{r}", f"={CALCULS}!$D${L_CMED + i}", qte, m["moy"])
        ws.write_formula(f"D{r}", f"=C{r}*$C${L_JC}", qte, besoin)
        ws.write_formula(f"E{r}", f"=D{r}*$C${L_MARGE}", qte, marge)
        ws.write_number(f"F{r}", 0, fmt(num_format="#,##0", **saisie))
        ws.write_formula(f"G{r}", f"=ROUNDUP(MAX(0,D{r}+E{r}-F{r}),0)", fmt(bold=True, num_format="#,##0"),
                         math.ceil(round(max(0, besoin + marge), 9)))
        ws.write_formula(f"H{r}", f"={CALCULS}!$X${L_CMED + i}", fmt(align="right"), m["meilleur_jsem"])
        ws.write_formula(f"I{r}", f"={CALCULS}!$Y${L_CMED + i}", fmt(align="right"), m["meilleur_mois"])
        ws.conditional_format(f"B{r}:I{r}", {"type": "formula", "criteria": f'=$C${L_MED + i}="Non"',
                                             "format": fmt(font_color="#9E9E9E")})
    L_ACH_FIN = L_ACH_T + n
    ws.data_validation(f"F{L_ACH_T + 1}:F{L_ACH_FIN}", {"validate": "decimal", "criteria": ">=", "value": 0,
                                                       "error_title": "Valeur non valide",
                                                       "error_message": "Le stock doit être positif ou nul."})
    ws.conditional_format(f"G{L_ACH_T + 1}:G{L_ACH_FIN}", {"type": "data_bar", "bar_color": VERT_BARRE,
                                                          "bar_solid": True, "min_type": "num", "min_value": 0})
    ws.write(f"B{L_ACH_FIN + 1}", "Jour et mois les plus forts : moyenne par jour sur tout l'historique des ventes.",
             petit)

    # ---------------- ⑥ Résultats détaillés
    titre_res = f'="⑥ Résultats par "&LOWER($C${L_REG})&IF({CP["par_jour"]}=1," (moyenne par jour)"," (somme)")'
    ws.merge_range(f"B{L_RES}:L{L_RES}", "", entete_f)
    ws.write_formula(f"B{L_RES}", titre_res, entete_f,
                     f"⑥ Résultats par {DEF['regroupement'].lower()} "
                     f"({'moyenne par jour' if v['par_jour'] else 'somme'})")
    ws.set_row(L_RES, 30)
    for k, titre in enumerate(["Période", *codes, "Total sélection", "Nb jours"]):
        ws.write(L_RES, 1 + k, titre, col_f if k == 0 else col_f_d)
    for i in range(1, MAX_LIGNES + 1):
        rr = L_CPER + i - 1
        rd = L_RES_DATA + i - 1
        ligne = v["lignes"][i - 1]
        actif = f"{CALCULS}!$B${rr}"
        num = f"{CALCULS}!$A${rr}"
        ws.write_formula(rd - 1, 1, f'=IF({actif}=1,{CALCULS}!$D${rr},"")', None,
                         ligne["libelle"] if ligne["actif"] else "")
        ws.write_formula(rd - 1, col_nb, f'=IF({actif}=0,"",IF({CP["g"]}<=5,COUNTIF({J["E"]},{num}),'
                                         f'COUNTIFS({H_HEURE},{num}-1,{crit_heure})))', fmt(num_format="0"),
                         ligne["nb"] if ligne["actif"] else "")
        for k, code in enumerate(codes):
            formule = (f'=IF(OR({actif}=0,{FLAG[code]}=0),"",IF({CP["g"]}<=5,SUMIF({J["E"]},{num},{D[code]}),'
                       f'SUMIFS({H[code]},{H_HEURE},{num}-1,{crit_heure}))'
                       f'/IF({CP["par_jour"]}=1,MAX(1,${l_nb}{rd}),1))')
            ws.write_formula(rd - 1, 2 + k, formule, qte, ligne["valeurs"][code] if ligne["actif"] else "")
        ws.write_formula(rd - 1, col_tot, f'=IF({actif}=0,"",SUM(C{rd}:{xl_col_to_name(col_tot - 1)}{rd}))',
                         fmt(bold=True, num_format=F_QTE), ligne["total"] if ligne["actif"] else "")
    derniere = L_RES_DATA + MAX_LIGNES - 1
    ws.conditional_format(f"{l_tot}{L_RES_DATA}:{l_tot}{derniere}",
                          {"type": "data_bar", "bar_color": VERT_BARRE, "bar_solid": True,
                           "min_type": "num", "min_value": 0})
    ws.conditional_format(f"B{L_RES_DATA}:{l_nb}{derniere}",
                          {"type": "formula", "criteria": f'=AND($B{L_RES_DATA}<>"",MOD(ROW(),2)=0)',
                           "format": fmt(bg_color=VERT_PALE)})
    for k, code in enumerate(codes):
        cellule = f"{xl_col_to_name(2 + k)}{L_RES + 1}"
        ws.conditional_format(cellule, {"type": "formula", "criteria": f'=$C${L_MED + k}="Non"',
                                        "format": fmt(font_color="#BDBDBD")})

    # ---------------- impression, protection, ouverture
    ws.print_area(f"A1:M{L_RES - 2}")
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.protect("", {"select_locked_cells": True, "select_unlocked_cells": True})
    return v
