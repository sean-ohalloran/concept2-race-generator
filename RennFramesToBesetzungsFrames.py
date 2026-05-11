import pandas as pd
import numpy as np
import json
import re
import os

#macht aus den einzelnen Panadaframes der Rennen pandaframes der Besetzungen
#Eine Besetzung beseteht aus allen Teilnehmern, die gleichzeitig auf den Ergos sitzen

def strecke_auswerten(strecke):
    #wertet die Strecke aus: es wird entweder bei einem zeit-Rennen die Zeit in sekunden oder bei einem Distanz-Rennen die Distanz in Metern returned
    try:
        if zeit_oder_distanz(strecke) == "time":
            #es handelt sich um ein Zeit-Rennen
            durationMin = re.findall(r'[0-9]+', strecke)[0]
            #Zeit von Minuten in Sekunden umrechenen
            durationSec = int(durationMin) * 60
            return durationSec
        if zeit_oder_distanz(strecke) == "meters":
            #es handelt sich um ein Distanz-rennen
            distance = int(re.findall(r'[0-9]+', strecke)[0])
            return distance
        else:
            print("Fehler: Strecke (Spalte H in Excel) kann nicht ausgewertet werden!")
            raise Exception("Fehler: Strecke (Spalte H in Excel) kann nicht ausgewertet werden!")
            return "error"
    except:
        print("Fehler: Strecke (Spalte H in Excel) kann nicht ausgewertet werden!")
        raise Exception("Fehler: Strecke (Spalte H in Excel) kann nicht ausgewertet werden!")
        return "error"


def zeit_oder_distanz(strecke):
    #prueft, ob es sich um ein Distanz-oder Streckenrennen handelt
    try:
        if re.search(r'[0-9]+ ?min', strecke):
            return "time"
        elif(re.search(r'[0-9]+ ?m', strecke)):
            return "meters"
        else:
            print("Fehler: Strecke (Spalte H in Excel) kann nicht ausgewertet werden!")
            raise Exception("Fehler: Strecke (Spalte H in Excel) kann nicht ausgewertet werden!")
            return "error"
    except:
        print("Fehler: Strecke (Spalte H in Excel) kann nicht ausgewertet werden!")
        raise Exception("Fehler: Strecke (Spalte H in Excel) kann nicht ausgewertet werden!")
        return "error"
    
def delete_files_in_directory(directory_path):
   #deletes all files in the Directory
   try:
     files = os.listdir(directory_path)
     for file in files:
       file_path = os.path.join(directory_path, file)
       if os.path.isfile(file_path):
         os.remove(file_path)
     print("All files in"+directory_path+" deleted successfully.")
   except OSError:
     print("Error occurred while deleting files.")
     return ("Error occurred while deleting Besetzungsframe files from "+directory_path)


def RennFramesToBesetzungsFrames():
    print("\n\n ------------------------------- Function: RennFramesToBesetzungsFrames was called:\n")
    listDateiNamen = []
    listErrors= []
    toBeMergedBesetzungsframes = []

    #---------------------- neuen ornder erstellen falls noch nicht vorhanden
    directory_path = r'./BesetzungenFrames' 
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        print("new direcory was created:   "+ directory_path)
    # alle Dateien aus dem Ordner BesetzungenJson loeschen
    listErrors.append(delete_files_in_directory(directory_path))

    with open(r"./RennenFrames/FileNames.json", "r") as file:
        ListOfFileNames = json.load(file)
    for dateiNameRennFrame in ListOfFileNames:
        print(dateiNameRennFrame)
        try:
            RennNr = re.findall(r'\((.*?)\)', dateiNameRennFrame)[0]   #alles was zwischen den Klammern steht
            #------------------------excel-Dateo RennFrame einlesen
            # usecols gibt die Excel Spalten an, aus denen der Dataframe erstellt wird
            #dtype=str macht aus allen Einträgen strings
            #header = none --> Spaltenbezeichnung wird auf B=1, C=2, D=3...
            dfRennen = pd.read_excel(dateiNameRennFrame, index_col=None, header = 0, usecols="B:K", dtype=str, engine="openpyxl")
            #print(dfRennen.loc[:,:])  #Zeile, Spalte   --> printed das Dataframes


            #pruefen, ob mehrere "A"        (ob das Rennen in mehrerer Besetungen aufgeteilt wird)
            contains_A = dfRennen.loc[:, 10].str.contains('A').any()
            if contains_A == True:
                #print("es ist ein A enthalten: das Rennen muss in mehrere Besetzungen aufgespalten werden!")
                #rennliste an den A-Stellen Spalten:
                # Überprüfen, an welchen Stellen 'A' in der Spalte 10 vorhanden ist
                mask_A = dfRennen.loc[:, 10].str.contains('A', na=False)
                #print(mask_A)
                # DataFrame an den Stellen splitten, an denen 'A' vorkommt
                df_RennTeile = []
                start_idx = 1  # Start des ersten Teils

                # Iteriere durch alle Zeilen, an denen 'A' vorkommt
                for idx, is_A in enumerate(mask_A):
                    if is_A:  # Wenn an dieser Stelle ein 'A' ist
                        dfRennteil = dfRennen.loc[start_idx:idx-1]     # Teil-Datenframe bis zum Index mit 'A'
                        dfRennteil = dfRennteil.set_index(pd.Index(range(2, 2 + len(dfRennteil))))    #Index aendern, sodass er bei 2 anfaengt
                        df_RennTeile.append(dfRennteil)  
                        start_idx = idx  # Setze den Start für das nächste Teil nach dem 'A'
                dfRennteil = dfRennen.loc[start_idx:idx-1]     #Das letzte Segment nach dem letzten 'A'
                dfRennteil = dfRennteil.set_index(pd.Index(range(2, 2 + len(dfRennteil))))    #Index aendern, sodass er bei 2 anfaengt
                df_RennTeile.append(dfRennteil) 
                df_RennTeile.pop(0)
                #for part in df_RennTeile:
                    #print(part)
                    #print("\n")
            else:
                df_RennTeile = [dfRennen.loc[2:]]


            for dfTeil in df_RennTeile:

                #--------------non-Nested Key-Value-Pairs----------------
                nameLong = dfRennen.loc[0,4]
                strecke = dfRennen.loc[0,7]       #Excel Rennen-kopfzeile, Spalte "H"
                duration = strecke_auswerten(strecke)
                RennenArt = zeit_oder_distanz(strecke)

                A = dfTeil.loc[2, 10]
                #BesetzungsName
                if pd.isna(A):
                    BesetzungsName = "r"+str(RennNr)
                else:
                    BesetzungsName = "r"+str(RennNr)+str(A)
                # Dict mit allen key-Value-Pairs auf der untersten ebene
                if RennenArt == "time":
                    dictNonNested = {
                        'Besetzung': BesetzungsName,
                        'Rennen': RennNr,
                        'A': A,
                        'Rennen_2': np.nan,
                        'A_2': np.nan,
                        '~': 'ab hier beginnen die für die Rac-Datei relevanten Infos',
                        'c2_race_id': '\"\"',
                        'duration': duration,
                        'duration_type': RennenArt,
                        'event_name': 'Ergocup',
                        'handicap_enabled': "false",
                        'name_long': nameLong,
                        'name_short': '\"\"',
                        'race_id': '\"\"',
                        'race_type': 'individual',
                        'split_value': duration,
                        'team_size': 1,
                        'time_cap': 0
                    }
                else:
                    dictNonNested = {
                        'Besetzung': BesetzungsName,
                        'Rennen': RennNr,
                        'A': A,
                        'Rennen_2': np.nan,
                        'A_2': np.nan,
                        '~': 'ab hier beginnen die für die Rac-Datei relevanten Infos',
                        'c2_race_id': '\"\"',
                        'duration': duration,
                        'duration_type': RennenArt,
                        'event_name': 'Ergocup',
                        'handicap_enabled': "false",
                        'name_long': nameLong,
                        'name_short': '\"\"',
                        'race_id': '\"\"',
                        'race_type': 'individual',
                        'split_value': duration,
                        'team_size': 1,
                        'time_cap': 0
                        }

                # Convert dictionary to DataFrame
                dfBesetzung = pd.DataFrame(list(dictNonNested.items()), columns=[1,2], index=range(len(dictNonNested)))
                dfBesetzung[[3, 4, 5, 6]] = np.nan


                #--------------------------boats (array)
                # Zeile mit boats eifuegen:
                boatsRowNum = len(dfBesetzung)   #Row bei der das Array boats beginnt
                boatsRow = ["boats", "["]
                dfBesetzung.loc[boatsRowNum] = boatsRow + [np.nan] * (6 - len(boatsRow))   #mit nan auffuellen

                #boats einfuegen:
                listdfBoats=[]
                anzahlBoats =  dfTeil.loc[2:,1].count()     #einfach alle nicht-NaN-eintraege in der Namenspalte zaehlen
                boatNum = 0
                for BoatNum in range(anzahlBoats):
                    affiliation = dfTeil.loc[2+BoatNum,8]   #!!!!!!!!!!!!!!!!vielleicht sind die Vereinsnamen zu lang!!!
                    class_name_raw = dfTeil.loc[2+BoatNum,4]
                    if pd.isna(class_name_raw):
                        class_name = "-"
                    else:
                        class_name = class_name_raw
                    lane_number = int(dfTeil.loc[2+BoatNum,5])
                    name = dfTeil.loc[2+BoatNum,1] + ", "+ dfTeil.loc[2+BoatNum,2]
                    #dictionary boat erstellen
                    boat = {
                        1: [np.nan] * 7,  # Padding with NaN for column '1'
                        2: [BoatNum] * 7,  # specifies place in Array '2'
                        3: ['affiliation', 'class_name', 'lane_number', 'name', 'participants', np.nan, np.nan],
                        4: [affiliation, class_name, lane_number, name, '[', 0, ']'],
                        5: [np.nan] * 5 + ['name', np.nan],  # Padding with NaN for column '5'
                        6: [np.nan] * 5 + ['\"\"', np.nan]  # Padding with NaN for column '6'
                    }
                    dfBoat = pd.DataFrame(boat)
                    listdfBoats.append(dfBoat)


                #merge all the boats dataframes with the non-nested-Dataframe
                dfBesetzung = pd.concat([dfBesetzung]+listdfBoats, ignore_index=True, sort=False)

                # Zeile mit ']' als abschluss des Boats-Array eifuegen:
                dfBesetzung.loc[len(dfBesetzung)] = [np.nan, "]"] + [np.nan] * (4)   #mit nan auffuellen


                

                # Besetzungsframes als Excel-Datei speichern:
                A = "" if pd.isna(A) else str(A)       # A fuer Dateiname passend formatieren
                dfBesetzung.to_excel(r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsName+r").xlsx") 

                #Liste mit gespeicherten Excel dateien ergaenzen
                listDateiNamen.append("BesetzungFrame_("+BesetzungsName+r").xlsx")

                print("Meldung: Successfully saved File: "+r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsName+r").xlsx")
                
                #----mehrere besetzungen gleichzeitig---!!!!!!!!!!!!!!muss noch geaendert werden!!!!!!!!!!!!!!
                # checken, ob mehrere Rennen gleichzeitig starten --> ob Besetzungen zusammengefuegt werden muessen
                Bemerkung  = dfRennen.loc[0,8]      #in der Bemerkung steht welche bzw. ob mehrere Rennen gleichzeitig starten
                if pd.isna(Bemerkung) == False:
                    if re.search(r"\AStartet", Bemerkung):
                        # Rennen1 beinhaltet kein A 
                        rennenName1 = "r" + str(RennNr)
                    elif re.search(r"\AA", Bemerkung):
                        ABezeichnung = re.findall(r"\AA[0-9]", Bemerkung)[0]
                        rennenName1 = "r" + str(RennNr)+ABezeichnung
                    if re.search("Rennen [0-9]+ ?A?[0-9]?", Bemerkung):
                        rennenName2 = re.findall("Rennen [0-9]+ ?A?[0-9]?", Bemerkung)[0]
                        # Use regular expression to remove "Rennen" and all spaces
                        rennenName2 = "r" + re.sub(r"Rennen|\s+", "", rennenName2)
                    toBeMergedBesetzungsframes.append((rennenName1, rennenName2))
        except Exception as e:
            listErrors.append(dateiNameRennFrame+"          "+str(e))
            print(e)
            print(dfRennen)
            #raise

 
    #save list of Filenames of to be merged Besetzungsframes to a Json File
    with open(r"./BesetzungenFrames/ToBeMergedFileNames.json", "w") as file:
        json.dump(toBeMergedBesetzungsframes, file, indent=4)
    
    #save list of Besetzungsframe - Filenames to Json File
    with open(r"./BesetzungenFrames/FileNames.json", "w") as file:
        json.dump(listDateiNamen, file, indent=4)


    print("\n")
    if len(listErrors) == 0:
        print("Meldung: successfully saved ... BesezungsFrame_() files")
    else:
        print("There were Errors with thte follwoing file(s):")
        for Error in listErrors:
            print(Error)
    print("\n\n")

    
    return listErrors

#test = RennFramesToBesetzungsFrames()
