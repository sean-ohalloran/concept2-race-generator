import pandas as pd
import numpy as np
import re
import os
import json

#-----------Dieses Programm Merged Besetzungsframes, die Gleichzeitig beginnen

pd.options.mode.copy_on_write = True   #because of:  SettingWithCopyWarning:A value is trying to be set on a copy of a slice from a DataFrame


def MergeBesetungsFrames():
    print("\n\n-------------------------------Function:  MergeBesetungsFrames    was called:\n")
    MergedFileNames = []
    listErrors = []

    #---------------------------open list of Filesnames of to be merged Besetzungsframes-------
    with open(r"./BesetzungenFrames/ToBeMergedFileNames.json", "r") as file:
        toBeMergedBesetzungsframes = json.load(file)
    #remove duplictaes from list of to be merged Besetzungen:
    unique_tuples = set(tuple(sorted(tup)) for tup in toBeMergedBesetzungsframes)
    toBeMergedBesetzungsframes = list(unique_tuples)            # Convert back to a list


    for BesetzungsPair in toBeMergedBesetzungsframes:
        #open besetzungsframe
        BesetzungsFrame0 = pd.read_excel(r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsPair[0]+").xlsx", index_col=0, header = 0, usecols="A:G", engine="openpyxl")
        BesetzungsFrame1 = pd.read_excel(r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsPair[1]+").xlsx", index_col=0, header = 0, usecols="A:G", engine="openpyxl")
        BesetzungsFrameMerged = BesetzungsFrame0    #Merged Besetzungsframe baut auf dem BesetzungsFrame0 auf

        #non-nested
        BesetzungsName = BesetzungsFrame0.loc[0,2]+"_"+ BesetzungsFrame1.loc[0,2]     #neuen BesetzungsName
        BesetzungsFrameMerged.loc[0,2] = BesetzungsName                               #neuen BesetzungsNamen an der richtigen Stelle im Frame speichern
        BesetzungsFrameMerged.loc[3,2] = BesetzungsFrame1.loc[1,2]                    #Rennen_2
        BesetzungsFrameMerged.loc[4,2] = BesetzungsFrame1.loc[2,2]                    #A_2
        BesetzungsFrameMerged.loc[11,2] = BesetzungsFrame0.loc[11,2]+" & "+BesetzungsFrame1.loc[11,2] #name long

        #boats
        #letzte Zeile loeschen:
        BesetzungsFrameMerged = BesetzungsFrameMerged.drop(BesetzungsFrameMerged.index[-1])

        #extract the boats frames from Dataframe1
        Boatindex1 = BesetzungsFrame1.loc[BesetzungsFrame1.loc[:,1].str.contains("boats", na=False)].index[0]
        Boats1  = BesetzungsFrame1.loc[(Boatindex1 +1):,:]        #extrahiert dataframe ab "boats"
        #ArrayNr anpassen:
        LetzteArryNummer0 = BesetzungsFrameMerged.iloc[-1, 1]
        for line in Boats1.index:
            if re.search("[0-9]", str(Boats1.loc[line,2])):
                ArrayNr = Boats1.loc[line,2]
                ArrayNr = int(ArrayNr) + int(LetzteArryNummer0) + 1 
                Boats1.loc[line,2] = str(ArrayNr)         

        #merge the boats Dataframes from 0 and 1
        BesetzungsFrameMerged = pd.concat([BesetzungsFrameMerged]+[Boats1], ignore_index=True, sort=False)





        try:
            #Als Excel speichern:
            BesetzungsFrameMerged.to_excel(r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsName+r").xlsx")
            print("new Besetzungsframe saved to excel-file: " + r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsName+r").xlsx")
        except Exception as e:
            listErrors.append(r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsName+r").xlsx"+"          "+str(e))
            print("error while trying to save file: ", r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsName+r").xlsx")
            print("maby the excel-file is open?")
            print(e)

        MergedFileNames.append(r"BesetzungFrame_("+BesetzungsName+r").xlsx")


    #----------------die Liste mit den Datei-Namen der Besetungsframes anpassen
    # Datei welche die Liste speichert oeffnen:
    with open(r"./BesetzungenFrames/FileNames.json", "r") as file:
        ListOfFileNames = json.load(file)
    #Besetzungsframes, aus denen das Gemergede Besetzungsframe erstellt wurde loeschen:
    for BesetzungsPair in toBeMergedBesetzungsframes:
        #!!!!!!!!!!!!funktioniert iwie nicht!!!!!!!!!!!
        ListOfFileNames.remove(r"BesetzungFrame_("+BesetzungsPair[0]+").xlsx")
        ListOfFileNames.remove(r"BesetzungFrame_("+BesetzungsPair[1]+").xlsx")
    #Dateinamen der gemergedn Besetzungsframes hinzufuegen:
    for MergedFilename in MergedFileNames:
        ListOfFileNames.append(MergedFilename)
    #die geaenderte Liste als Datei speichern:
    with open(r"./BesetzungenFrames/FileNames.json", "w") as file:
        json.dump(ListOfFileNames, file, indent=4)
    print("\nDatei mit Liste von Dateinamen der Besetzungsframes wurde aktualisiert!")


    #-----------Die Dateien loeschen, aus denen die Merged Besetzungsframes erstellt wurden
    for BesetzungsPair in toBeMergedBesetzungsframes:
        try:
            #Alten Besetzungsframes-Dateien loeschen:
            os.remove(r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsPair[0]+").xlsx")
            os.remove(r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsPair[1]+").xlsx")
        except Exception as e:
            listErrors.append(r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsPair[0]+").xlsx"+"\nbzw\n"+r"./BesetzungenFrames/BesetzungFrame_("+BesetzungsPair[1]+").xlsx"+"          "+str(e)+"\n error while trying to remove files: maby File is open?")
    print("\nDie alten Besetzungsframe-Excel-Dateien wurden geloescht!")

    #Errors ausgeben:
    for Error in listErrors:
        print(Error)

    return listErrors