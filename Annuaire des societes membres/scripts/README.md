# Scripts Python

Prerequis : `pip install pdfplumber openpyxl xlrd`

| Script | Role |
|---|---|
| `extraire_annuaire.py` | Lit les fiches societes d'un annuaire PDF (vie et non vie) |
| `lire_formulaires_xls.py` | Lit les formulaires Excel (.xls) des dossiers "Stes vie 2020" et "Stes non vie 2020" |
| `fusionner_annuaires.py` | Fusionne plusieurs sources dans la Base FANAF |
| `verifier_extraction.py` | Affiche le bilan des controles, source par source |

## Commandes

Depuis le dossier `Annuaire des societes membres`, sources de la plus ancienne a la plus recente :

```
python scripts/extraire_annuaire.py annuaire.pdf sortie.xlsx

python scripts/fusionner_annuaires.py "donnees/Base FANAF 2017 2024.xlsx" \
    "27e edition=ANNUAIRE_FANAF_2018_27e_Edition.pdf" \
    "Formulaires 2020=Stes vie 2020;Stes non vie 2020" \
    "30e edition=ANNUAIRE_FANAF_2021_30e_Edition.pdf" \
    "32e edition=ANNUAIRE_FANAF_2023_32e_Edition.pdf" \
    "33e edition=FANAF-ANNUAIRE-MARCHES-2026.pdf"

python scripts/verifier_extraction.py "33e edition=FANAF-ANNUAIRE-MARCHES-2026.pdf"
```

Une source est un PDF d'annuaire, ou une liste de dossiers de formulaires separes par ";".

## Principes

- Seules les donnees saisies sont reprises. Les valeurs deduites (totaux, parts, evolutions, ratios, autres actifs, taux de couverture) sont laissees de cote.
- Ces valeurs deduites servent a controler la lecture :
  - la somme des branches est comparee au total imprime ;
  - les evolutions N / N-1 sont recalculees ;
  - les ratios CS/PA sont recalcules.
- Les societes sont rapprochees d'une source a l'autre par pays, branche, nom, date de creation, directeur general et capital.
