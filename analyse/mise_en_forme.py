"""Thème vert et fonctions d'écriture communes à tous les onglets du classeur."""
import math
from datetime import date, datetime

import pandas as pd

VERT_FONCE, VERT, VERT_MOYEN = "#1B5E20", "#2E7D32", "#43A047"
VERT_PALE, VERT_SAISIE, VERT_BORD, GRIS = "#F1F8E9", "#DCEDC8", "#689F38", "#616161"
VERT_BARRE = "#A5D6A7"  # barres dans les cellules : assez clair pour que le chiffre reste lisible
ORANGE_PALE, ROUGE = "#F8CBAD", "#C62828"
F_QTE, F_QTE1, F_ENTIER = "#,##0.00", "#,##0.0", "#,##0"
F_DATE, F_DATE_HEURE, F_PCT, F_PCT1, F_INDICE = "dd/mm/yyyy", "dd/mm/yyyy hh:mm", "0%", "0.0%", "0.00"


class Styles:
    """Formats XlsxWriter créés à la demande et réutilisés (une seule police pour tout le classeur)."""

    def __init__(self, wb):
        self.wb = wb
        self._cache = {}

    def __call__(self, **proprietes):
        cle = tuple(sorted(proprietes.items()))
        if cle not in self._cache:
            self._cache[cle] = self.wb.add_format(
                dict(proprietes, font_name="Calibri", font_size=proprietes.get("font_size", 11)))
        return self._cache[cle]

    @property
    def entete(self):
        return self(bold=True, font_color="#FFFFFF", bg_color=VERT, valign="vcenter", indent=1)

    @property
    def colonne(self):
        return self(bold=True, font_size=9, font_color=GRIS, bottom=1, bottom_color=VERT_BORD, valign="bottom",
                    text_wrap=True)

    @property
    def colonne_droite(self):
        return self(bold=True, font_size=9, font_color=GRIS, bottom=1, bottom_color=VERT_BORD, valign="bottom",
                    text_wrap=True, align="right")

    @property
    def petit(self):
        return self(font_size=9, font_color=GRIS)


def valeur_cellule(v):
    """Convertit une valeur pandas/numpy en valeur écrivable (None pour une case vide)."""
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.to_pydatetime()
    if hasattr(v, "item"):
        v = v.item()
        if isinstance(v, float) and math.isnan(v):
            return None
    return v


def ecrire(ws, ligne, col, v, fmt=None):
    v = valeur_cellule(v)
    if v is None:
        if fmt is not None:
            ws.write_blank(ligne, col, None, fmt)
    elif isinstance(v, datetime):
        ws.write_datetime(ligne, col, v, fmt)
    elif isinstance(v, date):
        ws.write_datetime(ligne, col, datetime.combine(v, datetime.min.time()), fmt)
    elif isinstance(v, bool):
        ws.write_boolean(ligne, col, v, fmt)
    elif isinstance(v, (int, float)):
        ws.write_number(ligne, col, v, fmt)
    else:
        ws.write_string(ligne, col, str(v), fmt)


def titre_feuille(ws, st, titre, sous_titre, derniere_col=10):
    """Bandeau vert en haut d'un onglet (lignes 2 à 4)."""
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_row(0, 8)
    ws.set_row(1, 26)
    ws.set_row(2, 26)
    ws.merge_range(1, 1, 2, derniere_col, titre,
                   st(bold=True, font_size=20, font_color="#FFFFFF", bg_color=VERT_FONCE, valign="vcenter", indent=1))
    ws.merge_range(3, 1, 3, derniere_col, sous_titre, st(italic=True, font_color=GRIS, valign="vcenter"))
    return 5  # première ligne libre (indice 0)


def bloc(ws, st, ligne, titre, df, formats=None, col=1, largeur_titre=None, note=None, index_nom=None):
    """Écrit un titre de section puis un tableau (index en première colonne). Renvoie la ligne libre suivante.

    formats : {nom de colonne: format de nombre} ; "*" s'applique aux colonnes non citées.
    """
    formats = formats or {}
    n_col = len(df.columns) + 1
    largeur_titre = largeur_titre or n_col
    ws.merge_range(ligne, col, ligne, col + largeur_titre - 1, titre, st.entete)
    ligne += 1
    if note:
        ws.merge_range(ligne, col, ligne, col + largeur_titre - 1, note, st(font_size=9, font_color=GRIS, italic=True,
                                                                            text_wrap=True, valign="top"))
        ws.set_row(ligne, 15 * max(1, math.ceil(len(note) / (11 * largeur_titre))))
        ligne += 1
    ws.write(ligne, col, index_nom or (df.index.name or ""), st.colonne)
    for j, nom in enumerate(df.columns, start=1):
        ws.write(ligne, col + j, str(nom), st.colonne_droite)
    plus_long = max([len(str(c)) for c in df.columns] + [1])
    ws.set_row(ligne, max(30, 12 * math.ceil(plus_long / 14) + 6))
    premiere = ligne + 1
    for i, (idx, rang) in enumerate(df.iterrows()):
        r = premiere + i
        ecrire(ws, r, col, idx, st(num_format=F_DATE) if isinstance(idx, (pd.Timestamp, date)) else None)
        for j, nom in enumerate(df.columns, start=1):
            nf = formats.get(nom, formats.get("*"))
            v = rang[nom]
            if isinstance(v, str):
                ecrire(ws, r, col + j, v, st(align="right"))
            else:
                ecrire(ws, r, col + j, v, st(num_format=nf) if nf else None)
    return premiere + len(df) + 1, premiere


def zone(premiere, col_debut, nb_lignes, nb_cols):
    """Plage Excel (indices 0) pour une mise en forme conditionnelle."""
    return premiere, col_debut, premiere + nb_lignes - 1, col_debut + nb_cols - 1


def echelle_divergente(ws, premiere, col, nb_lignes, nb_cols, milieu=1.0):
    """Indices autour de 1 : orange pâle sous la moyenne, blanc à 1, vert au-dessus."""
    ws.conditional_format(*zone(premiere, col, nb_lignes, nb_cols), {
        "type": "3_color_scale", "min_color": ORANGE_PALE, "mid_type": "num", "mid_value": milieu,
        "mid_color": "#FFFFFF", "max_color": VERT_BARRE})


def barres(ws, premiere, col, nb_lignes):
    ws.conditional_format(premiere, col, premiere + nb_lignes - 1, col, {
        "type": "data_bar", "bar_color": VERT_BARRE, "bar_solid": True, "min_type": "num", "min_value": 0})
