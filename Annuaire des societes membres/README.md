# Annuaire des societes membres

Publication annuelle : une fiche par societe membre (vie et non vie), montants en milliers de F CFA.

## Contenu

| Element | Role |
|---|---|
| `Modele.xlsx` | Formulaire de collecte |
| `vba/` | Macros Excel de collecte (feuilles Collecte Vie et Collecte Non Vie) et de copie de feuilles |
| `scripts/` | Extraction des annuaires PDF et des formulaires Excel, fusion, verification |
| `tests/` | Tests de non regression de l'extraction |
| `donnees/` | Base FANAF 2017 2024 et donnees de la 32e edition |

## Editions

| Edition | Exercices | Etat |
|---|---|---|
| 19e (2012) | 2010 (a verifier) | Disponible, non integree |
| 20e (2013) | 2011 (a verifier) | Disponible, non integree |
| 25e (2018) | 2015, 2016 | A recuperer |
| 26e (2019) | 2016, 2017 | Disponible, non integree |
| 27e | 2017, 2018 | Integree |
| 28e (2021) | 2018, 2019 (a verifier) | A recuperer |
| 30e (2023) | 2020, 2021 | Integree |
| 32e (2025) | 2022, 2023 | Integree |
| 33e (2026) | 2023, 2024 | Integree |
| Formulaires 2020 | 2020 | Integres |

## Utilisation

Depuis ce dossier :

```
python scripts/fusionner_annuaires.py "donnees/Base FANAF 2017 2024.xlsx" "33e edition=chemin/annuaire.pdf" ...
FANAF_PDF_DIR=chemin/des/pdf python -m pytest tests -q
```
