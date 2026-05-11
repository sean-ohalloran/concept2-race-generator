----------Erklaerung der Bedienung des Programms:----------------------------------
Das Programm erstellt aus einer Excel Datei einzelne .rac2 Dateien fuer die Concept2 ErgRace App.

Programm starten: 
- ExcelZuErgRace.exe doppelklicken
- es oeffnet sich ein Fenster (das kann einige Sekunden dauern)
    - in der Zeile "Exceldatei" auf "Browse" klicken und die zu verarbeitende Datei anwaehlen
    - in der Zeile "Rennen-Ordner" auf "Browse" klicken und den Ordner auswählen, in den die fertigen .rac2 Dateien gespeichert werden sollen
    - mit "submit" bestaetigen
- jetzt laueft das Programm ab. Das kann einige Sekunden dauern
- es kommt die Meldung "das Programm ist durchlaufen!"
    - diese mit "ok" bestaetigen
- schließlich kommt die Meldung "List of Errors:..."
    - hier werden die Rennen angezeigt, die nicht verarbeitet werden konnten
    - fuer diese Rennen wurde keine .rac2 Datei erstellt. Das muss haendisch nachgeholt werden


--------bekannte Probleme:-----------------------------------------------------------
- Affiliation(Vereinsname) wird automatisch auf max. 4 Buchstaben gekuerzt
- Teams(bspw. Doppel) funktionieren nicht mit dem Programm


--------Hinweise dazu, wie das Programm funktioniert:--------------------------------
Die im GUI ausgewählte Exceldatei wird geoeffnet und es wird zur Sheet "Rennlisten" navigiert. 
(Die Excel-Datei muss genau so wie die Musterdatei aufgebaut sein!)
Es warden einzelne Rennen erstellt:
	- der Dateiname setzt sich aus der Rennnummer und dem Abschnitt(bspw A1, A2...) zusammen
	- wenn in der Bemerkungsspalte "I" etwas steht wird dies beruecksichtigt
	    - moegliche Bemerkungen:    (diese muessen auch genau so formuliert werden!)
	            - "Startet mit Rennen 20"
	            - "Startet mit Rennen 13 A3"
	            - "A2 startet mit Rennen 4 A1"
Es warden außerdem die Folgenden Daten zu den einzelnen Rennen extrahiert:
	- Name und Vorname          (Spalte "B", "C")
	- Rennbezeichnung           (bspw "Masters A-H, weiblich, 30 Min.") (Spalte "E")
	- Rennart und Rennstrecke   (Spalte "H")(bspw. "Zeit", "30min")
	- Klasse                    (Spalte "E")
	- Ergonummer                (Spalte "F")
	- Verein                    (Spalte "I") (!Vereinsnamen wird auf ersten 4 Buchstaben gekuerzt)