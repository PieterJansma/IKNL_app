# IKNL Interface – README

Deze repository bevat een Windows GUI-applicatie om IKNL-gegevens (CSV) te converteren naar REDCap-importbestanden met behulp van een sleutellijst (Excel) en een dictionary-Excel met alle mappings.

De GUI is gebouwd met Tkinter; de eigenlijke dataverwerking gebeurt in `final_iknl_exe_ready.py`.

## 1. Projectstructuur

Belangrijkste bestanden:

- `iknl_gui_bundle.py` – GUI (Tkinter)
- `final_iknl_exe_ready.py` – Dataverwerking
- `iknl_dictionary.xlsx` – Dictionary
- `data.csv.xlsx` – Data.csv
- `sleutellijst.xlsx` – Sleutellijst
- Outputbestanden:
  - `treat_iknl_import.csv`
  - `mal_iknl_import.csv`
  - `bl_iknl_import.csv`
  - `pb_iknl_import.csv`
  - `sx_iknl_import.csv`
  - `unknown_mappings.csv`

## 2. Vereisten

- Windows 10+
- Python 3.9+
- Modules:
  - pandas  
  - numpy  
  - openpyxl  
  - pyinstaller  

Installeren:

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install pandas numpy openpyxl pyinstaller
```

## 3. Applicatie lokaal draaien

```powershell
venv\Scripts\Activate.ps1
python iknl_gui_bundle.py
```

## 4. .exe bouwen met PyInstaller

### 4.1. Build-commando

```powershell
pyinstaller --clean --noconfirm `
  --onefile `
  --windowed `
  --name IKNL_Interface `
  --hidden-import pandas `
  --hidden-import numpy `
  --hidden-import openpyxl `
  --add-data "final_iknl_exe_ready.py;." `
  iknl_gui_bundle.py
```

Na de build staat de executable in:

```
dist/IKNL_Interface.exe
```

## 5. Gebruik van de .exe

1. Start **IKNL_Interface.exe**
2. Selecteer:
   - Sleutellijst (Excel)
   - Data (CSV)
   - Dictionary Excel
   - Outputmap
3. Klik **Run**

Output verschijnt in de gekozen map.

## 6. Opnieuw bouwen na wijzigingen

```powershell
venv\Scripts\Activate.ps1

Remove-Item -Recurse -Force build
Remove-Item -Recurse -Force dist
Remove-Item -Force IKNL_Interface.spec

pyinstaller --clean --noconfirm `
  --onefile `
  --windowed `
  --name IKNL_Interface `
  --hidden-import pandas `
  --hidden-import numpy `
  --hidden-import openpyxl `
  --add-data "final_iknl_exe_ready.py;." `
  iknl_gui_bundle.py
```

## 7. Troubleshooting

### GUI start niet
Start via commandline:

```powershell
python iknl_gui_bundle.py
```

### `final_iknl_exe_ready.py` niet gevonden
Controleer dat deze regel aanwezig is:

```
--add-data "final_iknl_exe_ready.py;."
```

### Onbekende mappingwaarden
Bekijk `unknown_mappings.csv` en vul ontbrekende keys aan in de dictionary-Excel.

---

