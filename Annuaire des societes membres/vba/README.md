# Macros Excel

Modules a importer dans un classeur .xlsm (Alt+F11, Fichier, Importer un fichier).

## ModConsolidation.bas : collecte des formulaires

| Macro | Role |
|---|---|
| `InitialiserClasseur` | Cree les feuilles Collecte Vie et Collecte Non Vie, avec les boutons Enregistrer et Supprimer le precedent |
| `EnregistrerVie`, `EnregistrerNonVie` | Controlent le formulaire colle en A:F puis l'ajoutent aux tables |
| `SupprimerPrecedentVie`, `SupprimerPrecedentNonVie` | Suppriment le dernier enregistrement de la branche, apres confirmation |
| `RafraichirCopies` | Recree les feuilles visibles a partir des feuilles brutes |

Tables produites :

| Feuille | Contenu |
|---|---|
| Identification | Une ligne par formulaire : societe, pays, branche, directeur general, capital, effectifs, annee |
| Emission&Prestations | Une ligne par bloc, categorie, rubrique, mesure et annee |
| Chiffres cles | Une ligne par rubrique et annee |

- Chaque table existe en version "(brut)", masquee, avec les valeurs en milliers de F CFA, la date d'import et la cellule source.
- La version visible reprend les valeurs par formule, multipliees par 1000.
- Le numero d'enregistrement est croissant et n'est jamais reutilise.

## ModCopieFeuilles.bas : copie des classeurs d'un dossier

`CopierFeuillesDuDossier` copie dans le classeur toutes les feuilles des classeurs Excel du meme dossier.

- Seules les valeurs, les formats et les largeurs de colonnes sont copies ; le code VBA, les boutons et les liaisons ne le sont pas.
- Chaque feuille est nommee "<fichier> - <feuille>", en 31 caracteres au plus.
- Sur OneDrive ou SharePoint, une fenetre demande de choisir le dossier local.

## Remarques

- Les fichiers sont en ASCII pur avec des fins de ligne CRLF ; les accents sont produits par `ChrW`, ce qui evite les caracteres illisibles a l'import.
- En cas d'erreur "Impossible d'executer la macro", lancer Debogage, Compiler VBAProject pour voir la ligne en cause.
