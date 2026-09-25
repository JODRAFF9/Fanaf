"""Fusionne deux editions de l'annuaire FANAF dans un classeur unique.

Chaque edition couvre deux exercices (32e edition : 2022-2023 ; 33e edition :
2023-2024). Les tables Emission&Prestations et Chiffres cles ont une ligne
par societe, rubrique et annee, et deux colonnes de valeurs, une par edition :
l'annee commune (2023) a donc ses deux valeurs cote a cote.

Les societes sont rapprochees d'une edition a l'autre (meme pays et meme
branche) par le nom, la date de creation, le directeur general et le capital,
ce qui couvre les changements de nom (Atlantique -> AFG) et les noms
illisibles du PDF 2026.

Usage : python fusionner_annuaires.py ANCIENNE.pdf RECENTE.pdf sortie.xlsx
"""
import difflib
import re
import sys
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font

import extraire_annuaire as ex

EDITIONS = ["32e edition", "33e edition"]
PAYS_ALIAS = {"CONGO BRAZZAVILLE": "CONGO"}


# Noms de societes illisibles dans le PDF 2026 (police mal encodee), decodes a la main :
# chaque page remplace les lettres par d'autres caracteres ; la correspondance se
# retrouve a partir des mots connus ("SOCIETE", "D'ASSURANCES", "REASSURANCES"...).
NOMS_DECODES = {
    '!"#A%C%DEF*A#EA+%DI-E!!.*E+#%!D%CDI%D*%E!!.*E+#%':
        "SOCIETE AFRICAINE D'ASSURANCES ET DE REASSURANCE",
    '!"#A%CC%DE"F*%+%D*!+%I-.I*F-*!%DL0-EE#I-!F%EDDA*%':
        "NOUVELLE SOCIETE INTERAFRICAINE D'ASSURANCES VIE",
    '!"#A%C%DE)CA"E)*%D+I)!!-.)E#%!D%CD+%D.%)!!-.)E#%!DLA%':
        "SOCIETE NATIONALE D'ASSURANCES ET DE REASSURANCES VIE",
    '!"#A%C%DE)CA"E)*%D+I)!!-.)E#%!D%CD+%D.%)!!-.)E#%!':
        "SOCIETE NATIONALE D'ASSURANCES ET DE REASSURANCES",
    '!"#ABCDE)G+IB--./BD!)G)MG/)B--./BD!)G-BN.GDO.#BD':
        "COMPAGNIE D'ASSURANCE ET REASSURANCE SABU NYUMAN",
    '!"#$%&%G($)%*$%((%G+,-!!.*-(#%!G%&G+%G*%-!!.*-(#%!':
        "SOCIETE NIGERIENNE D'ASSURANCES ET DE REASSURANCES",
}


def nom_illisible(nom):
    return not nom or bool(re.search(r'[!"#$%&*+]{2}|[!"#][A-Z%]', nom))


MOTS_GENERIQUES = {"assurance", "assurances", "sa", "s", "a", "de", "des", "du", "d", "l", "la", "le", "les", "et",
                   "vie", "iard", "iardt", "cie", "compagnie", "ex", "insurance", "cote", "ivoire", "benin",
                   "burkina", "faso", "cameroun", "cameroon", "congo", "brazzaville", "gabon", "mali", "niger",
                   "senegal", "tchad", "togo", "centrafrique"}


def distinctif(nom):
    """Nom sans les mots generiques ni le pays ("SAAR ASSURANCE COTE D'IVOIRE" -> "saar")."""
    return " ".join(m for m in ex.cle(nom).split() if m not in MOTS_GENERIQUES) or ex.cle(nom)


def sim(a, b):
    return difflib.SequenceMatcher(None, distinctif(a), distinctif(b)).ratio()


def normaliser(fiches):
    for f in fiches:
        f["ident"]["Pays"] = PAYS_ALIAS.get(f["ident"]["Pays"], f["ident"]["Pays"])
        brut = f["ident"]["Societe"]
        f["decode"] = brut in NOMS_DECODES
        if f["decode"]:
            f["ident"]["Societe"] = NOMS_DECODES[brut]
        f["illisible"] = nom_illisible(f["ident"]["Societe"])
    return fiches


def rapprocher(anciennes, recentes):
    """Couples (ancienne, recente) : un a un, meilleur score d'abord."""
    candidats = []
    for i, a in enumerate(anciennes):
        for j, r in enumerate(recentes):
            if (a["ident"]["Pays"], a["branche"]) != (r["ident"]["Pays"], r["branche"]):
                continue
            nom = 0 if r["illisible"] else sim(a["ident"]["Societe"], r["ident"]["Societe"])
            date = bool(a["ident"]["Date"]) and a["ident"]["Date"] == r["ident"]["Date"]
            dg = bool(a["ident"]["DG"]) and sim(a["ident"]["DG"], r["ident"]["DG"]) > 0.8
            cap = bool(a["ident"]["Capital"]) and ex.capital(a["ident"]["Capital"]) == ex.capital(r["ident"]["Capital"])
            dates_diff = bool(a["ident"]["Date"]) and bool(r["ident"]["Date"]) and not date
            # nom quasi identique (identique si les dates de creation different),
            # ou meme date de creation confirmee par un autre indice
            if nom >= (0.99 if dates_diff else 0.85) or (date and (nom >= 0.35 or dg or cap)):
                candidats.append((nom + 0.5 * date + 0.3 * dg + 0.2 * cap, i, j))
    candidats.sort(reverse=True)
    pris_a, pris_r, couples = set(), set(), {}
    for s, i, j in candidats:
        if i not in pris_a and j not in pris_r:
            pris_a.add(i)
            pris_r.add(j)
            couples[j] = (i, s)
    return couples


def construire(f32, f33):
    """Liste des societes : [{'fiches': [fiche32|None, fiche33|None], 'nom': ...}]"""
    couples = rapprocher(f32, f33)
    societes, notes = [], []
    lies = set()
    for j, r in enumerate(f33):
        a = None
        if j in couples:
            i, s = couples[j]
            a = f32[i]
            lies.add(i)
        nom = r["ident"]["Societe"]
        if r["decode"]:
            notes.append(("Nom illisible", f"33e edition p. {r['page']}", f"nom decode : {nom}"))
        if r["illisible"]:
            if a:
                notes.append(("Nom illisible", f"33e edition p. {r['page']}",
                              f"nom repris de la 32e edition : {a['ident']['Societe']}"))
                nom = a["ident"]["Societe"]
            else:
                nom = f"(nom illisible, 33e edition p. {r['page']})"
                notes.append(("Nom illisible", f"33e edition p. {r['page']}", "aucune correspondance trouvee"))
        if a and ex.cle(a["ident"]["Societe"]) != ex.cle(nom):
            notes.append(("Changement de nom", f"32e p. {a['page']} / 33e p. {r['page']}",
                          f"{a['ident']['Societe']} -> {nom}"))
        societes.append({"fiches": [a, r], "nom": nom, "pays": r["ident"]["Pays"], "branche": r["branche"]})
    for i, a in enumerate(f32):
        if i not in lies:
            societes.append({"fiches": [a, None], "nom": a["ident"]["Societe"], "pays": a["ident"]["Pays"],
                             "branche": a["branche"]})
    ordre_pays = {}
    for s in societes:
        ordre_pays.setdefault(s["pays"], len(ordre_pays))
    societes.sort(key=lambda s: (s["pays"], s["branche"] != "Vie", ex.cle(s["nom"])))
    return societes, notes


def ecrire(societes, notes, sortie):
    wb = Workbook()
    wb.remove(wb.active)
    gras = Font(bold=True)

    def feuille(nom, entetes, cache=False):
        ws = wb.create_sheet(nom)
        ws.append(entetes)
        for c in ws[1]:
            c.font = gras
        ws.freeze_panes = "A2"
        if cache:
            ws.sheet_state = "hidden"
        return ws

    e_id = ["No enregistrement", "No societe", "Edition", "Societe", "Pays", "Branche", "Directeur general",
            "Date de creation", "Capital social (F CFA)", "Cadres", "Maitrise", "Employes", "Annee N", "Source"]
    idv = feuille("Identification", e_id)
    e_ep = ["No societe", "Societe", "Pays", "Branche", "Bloc", "Categorie", "Rubrique", "Mesure", "Annee"]
    e_cc = ["No societe", "Societe", "Pays", "Branche", "Rubrique", "Annee"]
    epv = feuille("Emission&Prestations", e_ep + [f"Valeur {e} (F CFA)" for e in EDITIONS])
    ccv = feuille("Chiffres clés", e_cc + [f"Valeur {e} (F CFA)" for e in EDITIONS])
    epb = feuille("Emission&Prestations (brut)", e_ep + [f"Valeur {e} (milliers F CFA)" for e in EDITIONS]
                  + [f"Source {e}" for e in EDITIONS], True)
    ccb = feuille("Chiffres clés (brut)", e_cc + [f"Valeur {e} (milliers F CFA)" for e in EDITIONS]
                  + [f"Source {e}" for e in EDITIONS], True)

    n_enr = 0
    for no, s in enumerate(societes, 1):
        base_soc = [no, s["nom"], s["pays"], s["branche"]]
        # Identification : une ligne par fiche (edition)
        for k, f in enumerate(s["fiches"]):
            if f is None:
                continue
            n_enr += 1
            i = f["ident"]
            idv.append([n_enr, no, EDITIONS[k], s["nom"] if f["illisible"] else i["Societe"], s["pays"],
                        s["branche"], i["DG"] or None, ex.date_ou_texte(i["Date"]), ex.capital(i["Capital"]),
                        i["Cadres"], i["Maitrise"], i["Employes"], f["annees"][1], f"PDF p. {f['page']}"])
        # Valeurs : cle -> [v32, v33], sources
        lignes_ep, lignes_cc = {}, {}
        for k, f in enumerate(s["fiches"]):
            if f is None:
                continue
            for typ, bloc, cat, rub, mes, an, v in f["entrees"]:
                if typ == "EP":
                    d, cle_l = lignes_ep, (bloc, cat or None, rub, mes, an)
                else:
                    d, cle_l = lignes_cc, (rub, an)
                e = d.setdefault(cle_l, [None, None, None, None])
                e[k] = v
                e[2 + k] = f"PDF p. {f['page']}"
        for cle_l, (v1, v2, s1, s2) in lignes_ep.items():
            epb.append(base_soc + list(cle_l) + [v1, v2, s1, s2])
            r = epb.max_row
            epv.append(base_soc + list(cle_l) + [
                f"=IF('Emission&Prestations (brut)'!{c}{r}=\"\",\"\",'Emission&Prestations (brut)'!{c}{r}*1000)"
                for c in ("J", "K")])
        for cle_l, (v1, v2, s1, s2) in lignes_cc.items():
            ccb.append(base_soc + list(cle_l) + [v1, v2, s1, s2])
            r = ccb.max_row
            ccv.append(base_soc + list(cle_l) + [
                f"=IF('Chiffres clés (brut)'!{c}{r}=\"\",\"\",'Chiffres clés (brut)'!{c}{r}*1000)"
                for c in ("G", "H")])

    nt = feuille("Notes", ["Type", "Reference", "Detail"])
    for n in notes:
        nt.append(list(n))

    for c in idv["H"][1:]:
        c.number_format = "dd/mm/yyyy"
    for c in idv["I"][1:]:
        c.number_format = "#,##0"
    for ws, cols in ((epv, "JK"), (ccv, "GH")):
        for col in cols:
            for c in ws[col][1:]:
                c.number_format = "#,##0"
    for ws, cols in ((epb, "JK"), (ccb, "GH")):
        for col in cols:
            for c in ws[col][1:]:
                c.number_format = "#,##0.000"
    for ws in wb.worksheets:
        for col in ws.columns:
            larg = max((len(str(c.value)) for c in col[:300]
                        if c.value is not None and not str(c.value).startswith("=")), default=12)
            ws.column_dimensions[col[0].column_letter].width = min(max(10, larg + 2), 60)
    wb.save(sortie)


def main(pdf_ancien, pdf_recent, sortie):
    f32, ign32 = ex.lire_pdf(pdf_ancien)
    f33, ign33 = ex.lire_pdf(pdf_recent)
    normaliser(f32)
    normaliser(f33)
    societes, notes = construire(f32, f33)
    for ed, fs, ign in (("32e edition", f32, ign32), ("33e edition", f33, ign33)):
        if ign:
            notes.append(("Pages non lues", ed, ", ".join(map(str, ign))))
        for f in fs:
            for c in f["controles"]:
                notes.append(("Controle", f"{ed} p. {f['page']} {f['ident']['Societe'][:60]}", c))
            if f["annees"] != (fs[0]["annees"] if ed == "32e edition" else (2023, 2024)):
                notes.append(("Exercices", f"{ed} p. {f['page']} {f['ident']['Societe']}",
                              f"fiche publiee avec les exercices {f['annees'][0]}-{f['annees'][1]}"))
    ecrire(societes, notes, sortie)
    deux = sum(1 for s in societes if all(s["fiches"]))
    print(f"32e edition : {len(f32)} fiches ; 33e edition : {len(f33)} fiches")
    print(f"{len(societes)} societes : {deux} dans les deux editions, "
          f"{sum(1 for s in societes if s['fiches'][1] is None)} seulement 32e, "
          f"{sum(1 for s in societes if s['fiches'][0] is None)} seulement 33e")
    return societes, notes


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
