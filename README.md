# Construction du panel quantitatif (mémoire Master Finance)

## Fichiers livrables
- `build_quant_panel.py` : script principal pour extraire, harmoniser et assembler les données.
- `etude_quantitative_livrable2_panel_final.xlsx` : classeur Excel multi-feuilles (généré par le script).
- `panel_final.csv` : export CSV de la feuille `Panel_final` (généré par le script).

## Relancer le script
```bash
python3 build_quant_panel.py
```

## Dépendances Python
- pandas
- requests
- openpyxl

Installation typique :
```bash
python3 -m pip install pandas requests openpyxl
```

## Notes
- Le script utilise un cache local dans `./cache` pour éviter de re-télécharger les mêmes réponses API.
- En cas d'indisponibilité HTTP (API ou CSV de ratings), le script journalise un avertissement et continue.
