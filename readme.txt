------------------------------------------------
Erklaerung der Funktion des Programms:
Programm erstellt aus einer Excel Datei einzelne .rac2 Dateien fuer die Concept2 ErgRace-App.

Programm starten: main.py mit python starten
(funktioniert mit python-version 3.11.9)

------------------------------------------------
benoetigte libraries:  
pip install pandas
pip install numpy
pip install openpyxl
pip install PySimpleGUI

-----------------------------------------------
ggf. Starting Virtual enviroment:
pip install virtualenv
python3 -m virtualenv venv
Set-ExecutionPolicy Unrestricted -Scope Process
venv\Scripts\activate

-----------------------------------------------
todo:
- besseres Errorhandling
- testing
- Uebersichtlichkeit und Einheitlichkeit






----------------------------------------------
Hinweise zur funktion des Programms:
Die im GUI ausgewählte Excel file wird geöffnet und zur Sheet "Rennlisten" navigiert. (Die Excel-Datei muss so wie die Musterdatei aufgebaut sein)
Es werden einzelne Rennen erstellt:
	- der Dateiname setzt sich aus der Rennnummer und dem Abschnitt(bspw A1, A2...) zusammen
	- wenn in der Bemerkungsspalte "I" etwas steht (bspw."A2 startet mit Rennen 4 A1") wird dies beruecksichtigt
Es werden außerdem die folgenden Daten zu den einzelnen Rennen extrahiert:
	- Name und Vorname
	- Rennbezeichnung (bspw "Masters A-H, weiblich, 30 Min.")
	- Rennart und Rennstrecke (aus Spalte "H")(bspw. "Zeit", "30min")
	- Klasse
	- Ergonummer
	- Verein     !hierbei werden nur die ersten 4 buchstaben in der ErgRace App angezeigt
-------------------------------------------
bekannte Probleme:
- Affiliation (Vereinsname) wird automatisch auf max. 4 Buchstaben gekuerzt
- Teams(bspw. Doppel) funktionieren nicht mit dem Programm

