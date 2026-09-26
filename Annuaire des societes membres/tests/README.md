# Tests

Tests de non regression de l'extraction. Les valeurs attendues ont ete relevees a la main sur les pages rendues en image, puis figees dans `test_extraction.py`.

| Test | Page controlee |
|---|---|
| 27e edition | page 50 |
| 30e edition | pages 65 et 110 |
| 32e edition | pages 15 et 165 |
| 33e edition | pages 105 et 180 |
| Lecture des nombres et des libelles | sans PDF |

## Lancement

Depuis le dossier `Annuaire des societes membres` :

```
python -m pytest tests -q
```

Les PDF sont lus dans `documents/`, ou dans le dossier indique par la variable `FANAF_PDF_DIR`. Un test est ignore si son PDF est absent.
