"""Rapport de verification de l'extraction, source par source.

Pour chaque valeur non nulle extraite, trois niveaux :
- verifiee individuellement : recalculee a partir d'une colonne calculee de la source
  (evolution N/N-1, ratio CS/PA) et conforme a l'arrondi pres ;
- verifiee par un total : la somme des branches redonne le total imprime ;
- non verifiee : aucun controle ne la couvre (valeur seule sur sa ligne, par ex.).
Les echecs sont listes : erreur de lecture ou incoherence de la source elle-meme.

Usage : python verifier_extraction.py "libelle=chemin" ...   (meme syntaxe que la fusion)
"""
import os
import sys

import extraire_annuaire as ex
import lire_formulaires_xls as lx


def lire(chemin):
    if ";" in chemin or os.path.isdir(chemin):
        return lx.lire_dossiers([c for c in chemin.split(";") if c])
    return ex.lire_pdf(chemin)


def bilan(fiches):
    tot = ind = som = 0
    echecs = []
    tests = 0
    for f in fiches:
        v = f["verifs"]
        tests += v["tests"]
        for e in f["entrees"]:
            if not e[6]:
                continue
            tot += 1
            k = e[:6]
            if k in v["individuel"]:
                ind += 1
            elif k in v["somme"]:
                som += 1
        for x in v["echecs"]:
            echecs.append((f.get("source") or f"p. {f['page']}", f["ident"]["Societe"], x))
    return tot, ind, som, tests, echecs


def main(sources):
    resultats = []
    for lib, chemin in sources:
        fiches, _ = lire(chemin)
        tot, ind, som, tests, echecs = bilan(fiches)
        resultats.append((lib, len(fiches), tot, ind, som, tests, echecs))
        print(f"\n== {lib} : {len(fiches)} fiches, {tot} valeurs non nulles")
        print(f"   verifiees individuellement : {ind} ({100 * ind / tot:.1f} %)")
        print(f"   verifiees par un total     : {som} ({100 * som / tot:.1f} %)")
        print(f"   non verifiees              : {tot - ind - som} ({100 * (tot - ind - som) / tot:.1f} %)")
        print(f"   tests de recalcul : {tests}, echecs : {len(echecs)} ({100 * len(echecs) / max(tests, 1):.2f} %)")
        for src, soc, x in echecs:
            print(f"     - {src} {soc[:40]} : {x}")
    return resultats


if __name__ == "__main__":
    main([tuple(a.split("=", 1)) for a in sys.argv[1:]])
