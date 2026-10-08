"""Construit tout le classeur d'analyse à partir des quatre exports du client.

    python creer_classeur.py                      # exports/ -> Donnees_propres.xlsx
    python creer_classeur.py --maj-externe        # retélécharge aussi les données Sentinelles
    python creer_classeur.py --exports autre_dossier --sortie autre_classeur.xlsx

Chaque mois, il suffit de remplacer les quatre fichiers du dossier exports/ et de relancer la
commande : contrôles, statistiques, prévision, source externe et outil de sélection sont recalculés.
"""
import argparse
import os
import tempfile
from datetime import datetime
from pathlib import Path

import xlsxwriter

from analyse import donnees, externe, feuilles, prevision, statistiques
from analyse.mise_en_forme import VERT, VERT_FONCE, Styles
from analyse.tableau_de_bord import ajouter_tableau_de_bord

ICI = Path(__file__).resolve().parent
EXPORTS_PAR_DEFAUT = ICI / "exports"
EXTERNE_PAR_DEFAUT = ICI / "donnees_externes" / "sentinelles_grippe_france.csv"
SORTIE_PAR_DEFAUT = ICI / "Donnees_propres.xlsx"


def construire(dossier_exports, fichier_externe, sortie, maj_externe=False):
    fichier_externe = Path(fichier_externe)
    if maj_externe or not fichier_externe.exists():
        print(f"Téléchargement : {externe.URL}")
        externe.telecharger(fichier_externe)

    d = donnees.charger(dossier_exports)
    stats = statistiques.calculer(d)
    prev = prevision.evaluer(d.mois)
    ext = externe.analyser(d, externe.charger(fichier_externe))
    genere_le = datetime.now()

    sortie = Path(sortie)
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", dir=sortie.resolve().parent, delete=False)
    tmp.close()
    wb = xlsxwriter.Workbook(tmp.name)
    st = Styles(wb)

    # Ordre des onglets : analyse d'abord, données ensuite, calculs masqués à la fin
    ws = {nom: wb.add_worksheet(nom) for nom in
          ("Synthèse", "Tableau de bord", "Statistiques", "Prévision", "Source externe", "État des données")}
    # Pages de données : seulement celles que lisent les formules et les graphiques, masquées.
    # (Les exports hebdomadaire et mensuel ne servent qu'aux contrôles, faits ici en Python :
    # ils restent dans le dossier exports/ mais ne sont pas recopiés dans le classeur.)
    colonnes = {}
    pages_donnees = []
    for cle in ("journalier", "horaire"):
        nom_feuille = donnees.EXPORTS[cle][2]
        page = wb.add_worksheet(nom_feuille)
        colonnes[cle] = feuilles.ecrire_export(page, st, d.exports[cle], nom_feuille)
        pages_donnees.append(page)
    page = wb.add_worksheet("Grippe_Sentinelles")
    feuilles.ecrire_export(page, st, ext["semaines"].reset_index(), "Grippe_Sentinelles")
    pages_donnees.append(page)

    # Outil de sélection (formules sur les feuilles journalière et horaire)
    jour, horaire = d.exports["journalier"], d.exports["horaire"]
    nombre = lambda v: 0.0 if v != v else float(v)  # case vide -> 0, comme dans les formules
    jours = [(ts.date(), {c: nombre(v) for c, v in zip(d.codes, valeurs)})
             for ts, *valeurs in jour[["datum", *d.codes]].itertuples(index=False)]
    heures = [(ts.to_pydatetime(), int(h), {c: nombre(v) for c, v in zip(d.codes, valeurs)})
              for ts, h, *valeurs in horaire[["datum", "Hour", *d.codes]].itertuples(index=False)]
    ajouter_tableau_de_bord(wb, st, ws["Tableau de bord"], jours, heures, d.codes, colonnes["journalier"],
                            colonnes["horaire"], len(jour) + 1, len(horaire) + 1)

    feuilles.synthese(ws["Synthèse"], st, d, stats, prev, ext, genere_le)
    p_saison, ligne_graphe = feuilles.statistiques(ws["Statistiques"], st, d, stats)
    feuilles.graphe_saisonnalite(wb, ws["Statistiques"], d, p_saison, ligne_graphe)
    feuilles.prevision(ws["Prévision"], st, prev)
    feuilles.source_externe(wb, ws["Source externe"], st, ext, fichier_externe, len(ext["semaines"]))
    feuilles.etat_donnees(ws["État des données"], st, d)

    for nom in ("Synthèse", "Statistiques", "Prévision", "Source externe", "État des données"):
        ws[nom].set_landscape()
        ws[nom].set_paper(9)
        ws[nom].fit_to_pages(1, 0)
    ws["Synthèse"].set_tab_color(VERT_FONCE)
    ws["Tableau de bord"].set_tab_color(VERT_FONCE)
    for nom in ("Statistiques", "Prévision", "Source externe", "État des données"):
        ws[nom].set_tab_color(VERT)
    ws["Synthèse"].activate()
    for page in pages_donnees:
        page.hide()
    wb.close()
    os.replace(tmp.name, sortie)

    res = prev["resultats"]
    print(f"Classeur écrit : {sortie}")
    print(f"  {len(d.jours_complets)} jours complets, {len(d.mois)} mois complets, {ext['nb_semaines']} semaines "
          f"reliées à la grippe")
    print("  Prévisible le mois suivant : " + ", ".join(f"{c} {res.at[c, 'Prévisible ?']}" for c in res.index))


def main():
    parser = argparse.ArgumentParser(description="Construit le classeur d'analyse à partir des quatre exports.")
    parser.add_argument("--exports", type=Path, default=EXPORTS_PAR_DEFAUT, help="dossier des quatre exports CSV")
    parser.add_argument("--externe", type=Path, default=EXTERNE_PAR_DEFAUT, help="fichier Sentinelles (grippe)")
    parser.add_argument("--maj-externe", action="store_true", help="retélécharger les données Sentinelles")
    parser.add_argument("--sortie", type=Path, default=SORTIE_PAR_DEFAUT, help="classeur à écrire")
    args = parser.parse_args()
    construire(args.exports, args.externe, args.sortie, args.maj_externe)


if __name__ == "__main__":
    main()
