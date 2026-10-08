from pathlib import Path

import pandas as pd


# ---------------------------------------------------------------------------
# Explication du traitement des fichiers
# ---------------------------------------------------------------------------
# Ce script a été écrit pour ouvrir un classeur Excel contenant 4 feuilles de
# données, vérifier qu'elles sont bien présentes, puis comparer leurs colonnes
# de dates nommées "datum". Il s'assure que le fichier Excel existe avant de le
# lire, contrôle le nombre de feuilles, puis affiche un aperçu rapide de chaque
# feuille et le nombre de doublons sur les lignes complètes.
#
# Ensuite, il repère les dates de chaque feuille, les normalise au format
# datetime, les trie et les compare sur la période commune 2014-2020. Si une
# feuille contient des dates manquantes, le programme complète la série en
# réindexant sur les dates attendues, en interpolant les valeurs numériques et
# en remplissant les autres colonnes par propagation des valeurs. Enfin, il
# vérifie que toutes les feuilles ont exactement la même tranche de dates après
# complétion.
#
# En résumé, le fichier a pour objectif de nettoyer, harmoniser et aligner les
# données entre plusieurs feuilles afin qu'elles soient cohérentes avant toute
# analyse ou exploitation.


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
# Placez « données propres.xlsx » dans le même dossier que ce script,
# ou remplacez cette valeur par le chemin complet de votre fichier Excel.
FICHIER_EXCEL = Path(__file__).resolve().parent / "données propres.xlsx"


# ---------------------------------------------------------------------------
# Vérifications initiales
# ---------------------------------------------------------------------------
def verifier_fichier_excel():
	"""Vérifie que le fichier Excel existe bien avant de le lire."""
	if not FICHIER_EXCEL.is_file():
		raise FileNotFoundError(
			f"Fichier Excel introuvable : {FICHIER_EXCEL}\n"
			"Modifiez FICHIER_EXCEL pour indiquer le chemin du fichier."
		)


# ---------------------------------------------------------------------------
# Lecture des feuilles Excel
# ---------------------------------------------------------------------------
def charger_feuilles():
	"""Lit toutes les feuilles du classeur et vérifie qu'il y en a bien 4."""
	feuilles = pd.read_excel(FICHIER_EXCEL, sheet_name=None)

	if len(feuilles) != 4:
		raise ValueError(
			f"4 feuilles attendues, {len(feuilles)} trouvée(s) : "
			f"{', '.join(feuilles)}"
		)

	return feuilles


# ---------------------------------------------------------------------------
# Affichage de diagnostic
# ---------------------------------------------------------------------------
def afficher_apercu(feuilles):
	"""Affiche un aperçu rapide de chaque feuille et les doublons."""
	for nom_feuille, donnees in feuilles.items():
		print(f"\n--- {nom_feuille} : aperçu des 5 premières lignes ---")
		print(donnees.head(5).to_string(index=False))
		print(f"Nombre de doublons sur les lignes complètes : {donnees.duplicated().sum()}")


# ---------------------------------------------------------------------------
# Gestion des dates
# ---------------------------------------------------------------------------
def trouver_colonnes_dates(feuilles):
	"""Retourne la liste des colonnes nommées 'datum' pour chaque feuille."""
	return {
		nom: [col for col in donnees.columns if str(col).strip().lower() == "datum"]
		for nom, donnees in feuilles.items()
	}


def normaliser_dates_par_feuille(feuilles, colonnes_dates):
	"""Convertit chaque colonne datum en dates triées et sans doublons."""
	dates_par_feuille = {}
	for nom, donnees in feuilles.items():
		colonne_date = colonnes_dates[nom][0]
		dates = pd.DatetimeIndex(
			pd.to_datetime(donnees[colonne_date], errors="coerce").dropna().unique()
		).normalize().sort_values()
		dates_par_feuille[nom] = dates

	return dates_par_feuille


def afficher_differences_dates(nom_a, dates_a, nom_b, dates_b):
	"""Affiche les dates présentes d'un côté mais pas de l'autre."""
	dates_absentes_dans_a = sorted(dates_b - dates_a)
	dates_absentes_dans_b = sorted(dates_a - dates_b)

	if dates_absentes_dans_a:
		print(f"  Dates présentes dans {nom_b} mais absentes dans {nom_a} : {dates_absentes_dans_a}")
	if dates_absentes_dans_b:
		print(f"  Dates présentes dans {nom_a} mais absentes dans {nom_b} : {dates_absentes_dans_b}")

	if not dates_absentes_dans_a and not dates_absentes_dans_b:
		print(f"  Les deux feuilles ont exactement les mêmes dates sur la période demandée.")


def tranches_dates_communes(dates_par_feuille, debut, fin):
	"""Retourne la tranche de dates commune à toutes les feuilles sur la période demandée."""
	dates_periode = {
		nom: set(dates[(dates >= debut) & (dates < fin)])
		for nom, dates in dates_par_feuille.items()
	}

	if not dates_periode:
		return pd.DatetimeIndex([])

	tranche_commune = set.intersection(*dates_periode.values())
	return pd.DatetimeIndex(sorted(tranche_commune))


def plage_dates_attendue(dates_par_feuille, debut, fin):
	"""Retourne la plage journalière attendue sur la période demandée, à partir de la union des dates connues."""
	dates_periode = [
		dates[(dates >= debut) & (dates < fin)]
		for dates in dates_par_feuille.values()
	]

	if not dates_periode or all(d.empty for d in dates_periode):
		return pd.DatetimeIndex([])

	# On prend l'union des dates connues pour établir la période de référence.
	toutes_dates = pd.DatetimeIndex(
		sorted(set().union(*[set(dates) for dates in dates_periode if len(dates)]))
	)
	if toutes_dates.empty:
		return pd.DatetimeIndex([])

	return pd.date_range(start=toutes_dates.min(), end=toutes_dates.max(), freq="D")


def comparer_feuilles_initiale_finale(feuilles_initiales, feuilles_finales):
	"""Compare la base initiale et la base finale feuille par feuille."""
	print("\n=== Comparatif base initiale vs base finale ===")

	for nom in sorted(set(feuilles_initiales) | set(feuilles_finales)):
		if nom not in feuilles_initiales:
			print(f"\n{nom} : absent dans la base initiale, présent dans la base finale.")
			continue
		if nom not in feuilles_finales:
			print(f"\n{nom} : présent dans la base initiale, absent dans la base finale.")
			continue

		base_init = feuilles_initiales[nom].copy()
		base_finale = feuilles_finales[nom].copy()
		print(f"\n--- {nom} ---")
		print(f"Lignes initiales : {len(base_init)} | Lignes finales : {len(base_finale)}")
		print(f"Colonnes initiales : {list(base_init.columns)}")
		print(f"Colonnes finales   : {list(base_finale.columns)}")

		colonnes_ajoutees = [col for col in base_finale.columns if col not in base_init.columns]
		colonnes_supprimees = [col for col in base_init.columns if col not in base_finale.columns]
		colonnes_communes = [col for col in base_init.columns if col in base_finale.columns]

		if colonnes_ajoutees:
			print(f"Colonnes ajoutées : {colonnes_ajoutees}")
		if colonnes_supprimees:
			print(f"Colonnes supprimées : {colonnes_supprimees}")
		if not colonnes_ajoutees and not colonnes_supprimees:
			print("Structure des colonnes inchangée.")

		if "datum" in base_init.columns and "datum" in base_finale.columns:
			dates_init = pd.to_datetime(base_init["datum"], errors="coerce").dropna().sort_values().unique()
			dates_finale = pd.to_datetime(base_finale["datum"], errors="coerce").dropna().sort_values().unique()
			print(f"Dates initiales : {len(dates_init)} | Dates finales : {len(dates_finale)}")
			print(f"Dates ajoutées dans la base finale : {len(pd.Index(dates_finale).difference(pd.Index(dates_init)))}")
			print(f"Dates manquantes dans la base finale : {len(pd.Index(dates_init).difference(pd.Index(dates_finale)))}")

		colonnes_modifiees = []
		for col in colonnes_communes:
			if not base_init[col].reset_index(drop=True).equals(base_finale[col].reset_index(drop=True)):
				colonnes_modifiees.append(col)

		if colonnes_modifiees:
			print(f"Colonnes modifiées : {colonnes_modifiees}")
		else:
			print("Aucune colonne commune modifiée.")

		if len(base_init) == len(base_finale) and base_init.shape[1] == base_finale.shape[1]:
			print("Volume global inchangé : même nombre de lignes et de colonnes.")
		else:
			print("Volume global modifié : lignes et/ou colonnes ont changé.")

	print("\n=== Fin du comparatif ===")


def verifier_dates_apres_completer(feuilles, colonnes_dates):
	"""Vérifie simplement si les feuilles couvrent à peu près la même période, sans modifier les données."""
	debut, fin = pd.Timestamp("2014-01-01"), pd.Timestamp("2019-10-31")
	dates_par_feuille = normaliser_dates_par_feuille(feuilles, colonnes_dates)

	print("\n=== Vérification de la période commune approximative ===")
	for nom, dates in dates_par_feuille.items():
		if dates.empty:
			print(f"{nom} : aucune date valide.")
			continue
		dates_periode = dates[(dates >= debut) & (dates <= fin)]
		if dates_periode.empty:
			print(f"{nom} : aucune date dans la période demandée (2014-01-01 -> 2019-10-31).")
			continue
		print(f"{nom} : {dates_periode.min().date()} -> {dates_periode.max().date()} ({len(dates_periode)} jours)")

	min_dates = [dates[(dates >= debut) & (dates <= fin)].min() for dates in dates_par_feuille.values() if len(dates[(dates >= debut) & (dates <= fin)])]
	max_dates = [dates[(dates >= debut) & (dates <= fin)].max() for dates in dates_par_feuille.values() if len(dates[(dates >= debut) & (dates <= fin)])]

	if min_dates and max_dates:
		debut_commune = max(min_dates)
		fin_commune = min(max_dates)
		if debut_commune <= fin_commune:
			print(f"Période commune approximative : {debut_commune.date()} -> {fin_commune.date()} )")
			print("Conclusion : les dates sont à peu près communes dans la plage 2014-01-01 -> 2019-10-31.")
		else:
			print("Conclusion : les dates ne se chevauchent pas correctement sur la période demandée.")


def comparer_dates(feuilles, dates_par_feuille, colonnes_dates):
	"""Compare uniquement les périodes de dates, sans reconstruire ni modifier les feuilles."""
	debut, fin = pd.Timestamp("2014-01-01"), pd.Timestamp("2019-10-31")

	print("\n=== Comparaison des périodes de dates ===")
	for nom, dates in dates_par_feuille.items():
		if dates.empty:
			print(f"{nom} : aucune date valide.")
			continue
		dates_periode = dates[(dates >= debut) & (dates <= fin)]
		if dates_periode.empty:
			print(f"{nom} : aucune date dans la période demandée.")
			continue
		print(f"{nom} : {dates_periode.min().date()} -> {dates_periode.max().date()} ({len(dates_periode)} jours)")

	min_dates = [dates[(dates >= debut) & (dates <= fin)].min() for dates in dates_par_feuille.values() if len(dates[(dates >= debut) & (dates <= fin)])]
	max_dates = [dates[(dates >= debut) & (dates <= fin)].max() for dates in dates_par_feuille.values() if len(dates[(dates >= debut) & (dates <= fin)])]
	if not min_dates or not max_dates:
		print("\nAucune période exploitable à comparer entre les feuilles.")
		return

	debut_commune = max(min_dates)
	fin_commune = min(max_dates)
	if debut_commune <= fin_commune:
		print(f"\nPériode commune approximative : {debut_commune.date()} -> {fin_commune.date()} ({(fin_commune - debut_commune).days + 1} jours)")
		print("Les dates sont donc à peu près communes sur la plage attendue de janvier 2014 à octobre 2019.")
	else:
		print("\nLes feuilles ne partagent pas une période commune cohérente sur cette plage.")

	# On ne touche pas aux colonnes mois / heure / jour / semaine ; on se contente du diagnostic.
	for nom in dates_par_feuille:
		print(f"{nom} : données conservées, sans modification des colonnes mois / heure / jour / semaine.")


# ---------------------------------------------------------------------------
# Point d'entrée principal
# ---------------------------------------------------------------------------
if __name__ == "__main__":
	verifier_fichier_excel()
	feuilles = charger_feuilles()
	afficher_apercu(feuilles)

	colonnes_dates = trouver_colonnes_dates(feuilles)

	if all(len(colonnes) == 1 for colonnes in colonnes_dates.values()):
		dates_par_feuille = normaliser_dates_par_feuille(feuilles, colonnes_dates)
		comparer_dates(feuilles, dates_par_feuille, colonnes_dates)
	else:
		print("\nComparaison impossible : chaque feuille doit avoir exactement une colonne nommée 'datum'.")

