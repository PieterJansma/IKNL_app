# IKNL Interface

⚠️ **BELANGRIJK – EERST LEZEN**

Deze repository bevat de **broncode** van de IKNL Interface.  
**Voor normaal gebruik hoef je GEEN Python te installeren en GEEN code te draaien.**

➡️ **Wil je de tool gebruiken? Download uitsluitend de kant-en-klare `IKNL_Interface.exe`.**  
Deze repository is **alleen bedoeld voor developers** die wijzigingen willen aanbrengen en zelf een nieuwe `.exe` willen bouwen.

Als je hier bent om “gewoon de tool te gebruiken”: **je hebt alleen de .exe nodig.**

---

## Voor gebruikers (aanbevolen – 99% van de mensen)

1. Download `IKNL_Interface.exe`
2. Dubbelklik om te starten (geen installatie nodig)
3. Selecteer:
   - Sleutellijst (Excel)
   - Data (CSV)
   - Dictionary (Excel)
   - Outputmap
4. Klik op **Run**
5. De REDCap-importbestanden worden aangemaakt in de gekozen map

Klaar.

---

## Voor developers (alleen als je de tool wilt aanpassen)

Alleen relevant als je de code wilt wijzigen en zelf een nieuwe `.exe` wilt bouwen.

---

## Projectstructuur

- iknl_gui_bundle.py – GUI (Tkinter)
- final_iknl_exe_ready.py – Dataverwerking
- iknl_dictionary.xlsx – Dictionary
- data.csv.xlsx – Data
- sleutellijst.xlsx – Sleutellijst

Output:
- treat_iknl_import.csv
- mal_iknl_import.csv
- bl_iknl_import.csv
- pb_iknl_import.csv
- sx_iknl_import.csv
- unknown_mappings.csv

---

## Vereisten (development)

- Windows 10+
- Python 3.9+
- pandas
- numpy
- openpyxl
- pyinstaller

---

## Installatie (development)

python -m venv venv  
venv\Scripts\Activate.ps1  
pip install --upgrade pip  
pip install pandas numpy openpyxl pyinstaller  

---

## Lokaal draaien (development)

venv\Scripts\Activate.ps1  
python iknl_gui_bundle.py  

---

## .exe bouwen

pyinstaller --clean --noconfirm ^
  --onefile ^
  --windowed ^
  --name IKNL_Interface ^
  --hidden-import pandas ^
  --hidden-import numpy ^
  --hidden-import openpyxl ^
  --add-data "final_iknl_exe_ready.py;." ^
  iknl_gui_bundle.py

Na build:  
dist/IKNL_Interface.exe

---

## Opnieuw bouwen

venv\Scripts\Activate.ps1  
Remove-Item -Recurse -Force build  
Remove-Item -Recurse -Force dist  
Remove-Item -Force IKNL_Interface.spec  

pyinstaller --clean --noconfirm ^
  --onefile ^
  --windowed ^
  --name IKNL_Interface ^
  --hidden-import pandas ^
  --hidden-import numpy ^
  --hidden-import openpyxl ^
  --add-data "final_iknl_exe_ready.py;." ^
  iknl_gui_bundle.py

---

## Troubleshooting

GUI start niet:  
python iknl_gui_bundle.py  

final_iknl_exe_ready.py niet gevonden:  
Zorg dat deze regel in het PyInstaller-commando staat:  
--add-data "final_iknl_exe_ready.py;."

Onbekende mappingwaarden:  
Bekijk unknown_mappings.csv en vul ontbrekende keys aan in de dictionary-Excel.

---

## TL;DR

Gebruiken? → download en start `IKNL_Interface.exe`  
Aanpassen? → gebruik deze repo + bouw zelf een nieuwe .exe  
Twijfel? → je hebt waarschijnlijk alleen de .exe nodig
