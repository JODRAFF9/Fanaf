"""Lit les formulaires FANAF au format Excel (.xls/.xlsx), une feuille par societe
(dossiers "Stes vie 2020" et "Stes non vie 2020"), et renvoie des fiches au meme
format que extraire_annuaire.lire_fiche, pour la fusion dans la Base FANAF.

Regles :
- seules les feuilles au format fiche societe sont lues (A1 = "NOM DE LA SOCIETE"
  et blocs A, B, C presents) ; les feuilles de marche et les reassureurs sont exclus ;
- quand un classeur existe en version revue ("... REVUE.xls", "... REVU.xls"),
  seule cette version est lue ;
- comme pour les PDF, seules les donnees saisies sont reprises (pas de totaux,
  parts, evolutions, ratios, autres actifs ni taux de couverture).
"""
import glob
import os
import re
from datetime import datetime

import extraire_annuaire as ex

try:
    import xlrd
except ImportError:  # pragma: no cover
    xlrd = None


def texte(v):
    return " ".join(str(v).split()) if v not in (None, "") else ""


def nombre(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return ex.valeur(str(v))


def annee(v):
    n = nombre(v)
    if n is not None and 10000 < n < 30000 and n == int(n):
        n = int(str(int(n))[:2] + str(int(n))[-2:])  # faute de frappe : 20219 -> 2019
    return int(n) if n is not None and 1990 <= n <= 2100 and n == int(n) else None


def feuilles(chemin):
    """(nom de feuille, grille de valeurs, datemode) pour chaque feuille du classeur."""
    wb = xlrd.open_workbook(chemin)
    for s in wb.sheets():
        grille = [[s.cell_value(r, c) for c in range(s.ncols)] for r in range(s.nrows)]
        yield s.name, grille, wb.datemode


def cellule(g, r, c):
    return g[r][c] if 0 <= r < len(g) and 0 <= c < len(g[r]) else ""


def date_creation(v, datemode):
    if isinstance(v, float) and v > 1000:
        try:
            return xlrd.xldate_as_datetime(v, datemode).strftime("%d/%m/%Y")
        except Exception:
            return texte(v)
    return texte(v)


def lire_feuille(g, datemode, source):
    col_b = [texte(cellule(g, r, 1)).upper() for r in range(len(g))]

    def titre(motif, debut=0):
        return next((r for r in range(debut, len(g)) if re.search(motif, col_b[r])), None)

    if texte(cellule(g, 0, 0)).upper() != "NOM DE LA SOCIETE":
        return None
    iA = titre(r"EMISSIONS")
    iB = titre(r"PRESTATIONS|SINISTRALITE", (iA or 0) + 1)
    iC = next((r for r in range((iB or 0) + 1, len(g)) if texte(cellule(g, r, 0)).lower().startswith("rubriques")), None)
    if None in (iA, iB, iC):
        return None
    branche = "Non-vie" if "SINISTRALITE" in col_b[iB] else "Vie"
    an1, an2 = annee(cellule(g, iA + 1, 1)), annee(cellule(g, iA + 1, 3))
    if not an1 or not an2:
        return None
    annees = (an1, an2)
    brut_annees = (cellule(g, iA + 1, 1), cellule(g, iA + 1, 3))

    ident = {"Societe": texte(cellule(g, 0, 1)), "Pays": texte(cellule(g, 1, 1)),
             "DG": texte(cellule(g, 2, 1)), "Date": date_creation(cellule(g, 3, 1), datemode),
             "Capital": texte(cellule(g, 4, 1)) if not isinstance(cellule(g, 4, 1), float)
             else str(int(cellule(g, 4, 1))),
             "Cadres": nombre(cellule(g, 5, 3)), "Maitrise": nombre(cellule(g, 6, 3)),
             "Employes": nombre(cellule(g, 7, 3))}
    entrees, controles = [], []
    if any(isinstance(v, float) and v > 3000 for v in brut_annees):
        controles.append(f"annees saisies {int(brut_annees[0])} et {int(brut_annees[1])}, lues {an1} et {an2}")

    def bloc(t, fin, nom_bloc):
        sous = cellule(g, t + 2, 1)
        if not texte(cellule(g, t + 2, 0)) and isinstance(sous, str) and sous.strip():
            # non-vie : deux mesures par annee (N-1 en B et C, N en D et E) ; ratios ignores
            mesures = []
            for c_n1, c_n in ((1, 3), (2, 4)):
                lib = texte(cellule(g, t + 2, c_n1))
                if lib and "/" not in lib:
                    mesures.append((ex.canonique(lib, ["Primes émises", "Primes acquises (PA)"])
                                    if not lib.startswith("(CS)") else "Charges de sinistres (CS)", c_n1, c_n))
            debut, categorie = t + 3, "Affaires directes"
        else:
            m = "Emissions nettes" if nom_bloc == "Emissions" else "Prestations versees"
            mesures, debut, categorie = [(m, 1, 3)], t + 2, ""
        reference = ex.BRANCHES_NV if branche == "Non-vie" else ex.RUBRIQUES_VIE
        n_vie = 0
        sommes, totaux = {}, {}
        for r in range(debut, fin):
            lab = ex.libelle(texte(cellule(g, r, 0)))
            if not lab:
                continue
            low = lab.lower()
            vide = all(texte(cellule(g, r, c)) == "" for c in range(1, 6))
            if vide and low.startswith("assurances "):
                categorie = ex.libelle(lab.split(None, 1)[1])
                continue
            if low.startswith("branches") or low.startswith("(chiffre"):
                continue
            if low.startswith("total") or low.startswith("ensemble") or low.startswith("chiffre d"):
                k = "total" if low.startswith("total") else "ensemble"
                for mes, c1, c2 in mesures:
                    totaux[(k, mes, annees[0])] = nombre(cellule(g, r, c1))
                    totaux[(k, mes, annees[1])] = nombre(cellule(g, r, c2))
                continue
            lab = ex.canonique(lab, reference)
            if branche == "Vie" and ex.cle(lab).startswith("contrat") and ex.cle(lab).endswith("de vie"):
                n_vie += 1
                categorie = "Individuelles" if n_vie == 1 else "Collectives"
            if lab == "Contrats en cas de vie" and categorie == "Collectives":
                lab = "Contrat en cas de vie"
            cat = "Acceptations" if lab == "Acceptations" else categorie
            for mes, c1, c2 in mesures:
                for an, c in ((annees[0], c1), (annees[1], c2)):
                    v = nombre(cellule(g, r, c))
                    entrees.append(("EP", nom_bloc, cat, lab, mes, an, v))
                    k = ("acc", mes, an) if cat == "Acceptations" else (mes, an)
                    sommes[k] = sommes.get(k, 0) + (v or 0)
        for mes, _, _ in mesures:
            for an in annees:
                s, acc = sommes.get((mes, an), 0), sommes.get(("acc", mes, an), 0)
                t_ = totaux.get(("total", mes, an))
                e_ = totaux.get(("ensemble", mes, an))
                if t_ is not None and abs(s - t_) > 2:
                    controles.append(f"{nom_bloc} / {mes} / {an} : somme des branches {s:,.0f} "
                                     f"<> total affaires directes {t_:,.0f}".replace(",", " "))
                if e_ is not None and abs(s + acc - e_) > 2:
                    controles.append(f"{nom_bloc} / {mes} / {an} : branches + acceptations {s + acc:,.0f} "
                                     f"<> ensemble {e_:,.0f}".replace(",", " "))

    bloc(iA, iB, "Emissions")
    bloc(iB, iC, "Sinistralite" if branche == "Non-vie" else "Prestations")

    # Bloc C : annees N-1 et N en colonnes C et D
    vals_c = {}
    for r in range(iC + 1, len(g)):
        lab = ex.libelle(texte(cellule(g, r, 0)))
        if not lab:
            continue
        if lab.lower().startswith("taux"):
            break
        lab = ex.canonique(lab, ex.RUBRIQUES_C)
        v = [nombre(cellule(g, r, 2)), nombre(cellule(g, r, 3))]
        vals_c[lab] = v
        if lab == "Autres actifs":
            continue
        for k, an in enumerate(annees):
            entrees.append(("CC", "", "", lab, "", an, v[k]))
    aa, liq, au = vals_c.get("Actifs admis"), vals_c.get("Dont liquidités"), vals_c.get("Autres actifs")
    if aa and liq and au:
        for k, an in enumerate(annees):
            if None not in (aa[k], liq[k], au[k]) and abs(aa[k] - liq[k] - au[k]) > 2:
                controles.append(f"Chiffres cles / {an} : actifs admis - liquidites "
                                 f"{aa[k] - liq[k]:,.0f} <> autres actifs {au[k]:,.0f}".replace(",", " "))

    return {"page": source, "source": source, "page_imprimee": "", "branche": branche, "annees": annees,
            "ident": ident, "entrees": entrees, "controles": controles}


def lire_dossiers(dossiers):
    """Lit tous les classeurs des dossiers. Renvoie (fiches, feuilles ignorees)."""
    fiches, ignorees = [], []
    for d in dossiers:
        chemins = sorted(glob.glob(os.path.join(d, "*.xls")) + glob.glob(os.path.join(d, "*.xlsx")))
        revus = {re.sub(r"\s+REVUE?(?=\.xlsx?$)", "", c, flags=re.I) for c in chemins
                 if re.search(r"\sREVUE?\.xlsx?$", c, re.I)}
        for chemin in chemins:
            if chemin in revus:
                continue  # remplace par sa version revue
            rel = f"{os.path.basename(d)}/{os.path.basename(chemin)}"
            for nom, g, dm in feuilles(chemin):
                if re.search(r"march|societes", nom, re.I):
                    continue
                source = f"{rel}, feuille {nom.strip()}"
                try:
                    f = lire_feuille(g, dm, source)
                except Exception as e:  # feuille atypique
                    f = None
                    ignorees.append(f"{source} ({e})")
                    continue
                if f is None:
                    ignorees.append(source)
                    continue
                p = ex.pays_normalise(f["ident"]["Pays"])
                f["ident"]["Pays"] = p
                fiches.append(f)
    return fiches, ignorees
