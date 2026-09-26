"""Tests de non-regression de l'extraction des annuaires FANAF.

Les valeurs attendues ont ete relevees a la main sur les pages rendues en image
(controle visuel), puis figees ici. Les PDF sont cherches dans FANAF_PDF_DIR
(par defaut le sous-dossier documents) ; un test est ignore si son PDF est absent.

Lancement, depuis ce dossier : python -m pytest tests -q
"""
import os
import sys

import pytest

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(RACINE, "scripts"))
import extraire_annuaire as ex  # noqa: E402

DOSSIER = os.environ.get("FANAF_PDF_DIR", os.path.join(RACINE, "documents"))
_cache = {}


def fiche(pdf, page):
    chemin = os.path.join(DOSSIER, pdf)
    if not os.path.exists(chemin):
        pytest.skip(f"{pdf} absent de {DOSSIER}")
    if (pdf, page) not in _cache:
        import pdfplumber
        with pdfplumber.open(chemin) as doc:
            _cache[(pdf, page)] = ex.lire_fiche(doc.pages[page - 1], page)
    return _cache[(pdf, page)]


def valeurs(f):
    return {(e[1], e[2], e[3], e[4], e[5]): e[6] for e in f["entrees"]}


# ---------------------------------------------------------------- 27e edition, p. 50 (vie)
def test_27e_nsia_vie_cameroun():
    f = fiche("annuaire_fanaf_2020_27eme_edition.pdf", 50)
    assert f["branche"] == "Vie" and f["annees"] == (2017, 2018)
    i = f["ident"]
    assert (i["Societe"], i["Pays"], i["Date"]) == ("NSIA VIE ASSURANCES", "CAMEROUN", "21/10/2013")
    assert (i["Cadres"], i["Maitrise"], i["Employes"]) == (7, 5, 5)
    v = valeurs(f)
    E, P = "Emissions nettes", "Prestations versees"
    assert v[("Emissions", "Individuelles", "Contrat en cas de décès", E, 2017)] == 33642
    assert v[("Emissions", "Collectives", "Contrat en cas de décès", E, 2018)] == 1286366
    assert v[("Emissions", "Acceptations", "Acceptations", E, 2017)] == 28423
    assert v[("Prestations", "Individuelles", "Epargne", P, 2017)] is None  # case vide dans le PDF
    assert v[("Prestations", "Individuelles", "Epargne", P, 2018)] == 4361
    assert v[("Prestations", "Collectives", "Epargne", P, 2017)] == 16169
    assert v[("", "", "Produits financiers nets", "", 2018)] == -55390
    assert v[("", "", "Résultats d’exploitations nets", "", 2017)] == -150967
    # "Fonds propres nets" (ancien libelle) rattache aux capitaux propres
    assert v[("", "", "Total des capitaux propres et réserves", "", 2018)] == 1399367
    assert v[("", "", "Dont liquidités", "", 2018)] == 603014
    assert not f["verifs"]["echecs"]


# ---------------------------------------------------------------- 30e edition, p. 110 (non-vie)
def test_30e_nsia_cote_ivoire():
    f = fiche("annuaire_fanaf_2023_30eme_edition.pdf", 110)
    assert f["branche"] == "Non-vie" and f["annees"] == (2020, 2021)
    v = valeurs(f)
    A, PE, PA, CS = "Affaires directes", "Primes émises", "Primes acquises (PA)", "Charges de sinistres (CS)"
    assert v[("Emissions", A, "Accidents corporels et maladie", PE, 2020)] == 11177947
    assert v[("Emissions", A, "Accidents corporels et maladie", PA, 2020)] == 11187443
    assert v[("Emissions", A, "Accidents corporels et maladie", PE, 2021)] == 13227424
    assert v[("Emissions", A, "Accidents corporels et maladie", PA, 2021)] == 13263007
    assert v[("Emissions", "Acceptations", "Acceptations", PA, 2021)] == 151299
    assert v[("Sinistralite", A, "Transports Aériens", CS, 2020)] == -64
    assert v[("Sinistralite", A, "Autres Risques", CS, 2020)] == -564323
    assert v[("", "", "Primes acquises aux réassureurs", "", 2021)] == 474645
    assert not f["verifs"]["echecs"]


# ---------------------------------------------------------------- 33e edition, p. 180 (non-vie)
def test_33e_wafa_senegal():
    f = fiche("annuaire_fanaf_2026_33eme_edition.pdf", 180)
    assert f["annees"] == (2023, 2024)
    v = valeurs(f)
    A, CS = "Affaires directes", "Charges de sinistres (CS)"
    assert v[("Emissions", A, "RC Générale", "Primes acquises (PA)", 2024)] == 309015
    assert v[("Sinistralite", A, "Transports Maritimes", CS, 2023)] == 0
    assert v[("Sinistralite", A, "Transports Maritimes", CS, 2024)] == 28133
    assert v[("Sinistralite", A, "Incendie et autres dommages aux biens", CS, 2024)] == -121279
    assert v[("", "", "Résultats au Bilan", "", 2023)] == -1072084
    assert ("", "", "Autres actifs", "", 2024) not in v  # valeur calculee : non reprise


# ---------------------------------------------------------------- 32e edition, variantes de mise en page
def test_32e_atlantique_benin_vie():
    f = fiche("annuaire_fanaf_2025_32eme_edition.pdf", 15)
    v = valeurs(f)
    assert v[("Emissions", "Collectives", "Epargne", "Emissions nettes", 2023)] == 1217705
    assert v[("Prestations", "Collectives", "Epargne", "Prestations versees", 2023)] == 1910483
    assert len(f["verifs"]["individuel"]) == sum(1 for e in f["entrees"] if e[6])  # tout est recalcule


def test_32e_citoyenne_vie_tirets_et_nombres_coupes():
    # p. 165 : "-" pour zero, nombre coupe en deux ("9" + "5 037"), titre de bloc C different
    f = fiche("annuaire_fanaf_2025_32eme_edition.pdf", 165)
    v = valeurs(f)
    assert v[("Emissions", "Individuelles", "Complémentaires", "Emissions nettes", 2022)] == 0
    assert v[("", "", "Résultats d’exploitations nets", "", 2022)] == -34041
    assert v[("", "", "Autres charges de l'exercice", "", 2023)] == 451399  # "Frais Generaux"


def test_33e_titres_de_blocs_mal_numerotes():
    # p. 105 : "B - EMISSIONS NETTES" et "C - SINISTRALITE"
    f = fiche("annuaire_fanaf_2026_33eme_edition.pdf", 105)
    assert f is not None and f["branche"] == "Non-vie"
    assert valeurs(f)[("Emissions", "Affaires directes", "Accidents corporels et maladie",
                       "Primes émises", 2023)] == 13249131


def test_30e_entete_bloc_c_incomplet():
    # p. 65 : l'en-tete du bloc C n'indique que 2021
    f = fiche("annuaire_fanaf_2023_30eme_edition.pdf", 65)
    v = valeurs(f)
    assert v[("", "", "Produits financiers nets", "", 2020)] == 43371
    assert v[("", "", "Produits financiers nets", "", 2021)] == 12866


# ---------------------------------------------------------------- fonctions de base
def test_valeur():
    assert ex.valeur("1 177 646") == 1177646
    assert ex.valeur("-") == 0
    assert ex.valeur("- 34 041") == -34041
    assert ex.valeur("12,5") == 12.5
    assert ex.valeur("abc") is None


def test_canonique():
    assert ex.canonique("autes actifs", ex.RUBRIQUES_C) == "Autres actifs"
    assert ex.canonique("Marg disponible", ex.RUBRIQUES_C) == "Marge disponible"
    assert ex.canonique("Frais Généraux", ex.RUBRIQUES_C) == "Autres charges de l'exercice"
    assert ex.canonique("U°E", ex.BRANCHES_NV) == "Accidents corporels et maladie"
