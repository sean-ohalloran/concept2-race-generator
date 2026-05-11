import pandas as pd
import os
import json



def test():
    print("test was called!")
    pass

def delete_files_in_directory(directory_path):
   try:
     files = os.listdir(directory_path)
     for file in files:
       file_path = os.path.join(directory_path, file)
       if os.path.isfile(file_path):
         os.remove(file_path)
     print("All files in"+directory_path+" deleted successfully.")
   except OSError:
     print("Error occurred while deleting files.")
     return ("Error occurred while deleting files from "+directory_path)


def exl_to_df(ExcelFilePath):
    listErrors = []
    print("\n\n-------------------------------function: exl_to_df was called:\n")
    #----------------- neuen ornder erstellen falls noch nicht vorhanden
    directory_path = r'.\RennenFrames' 
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        print("new direcory was created:   "+ directory_path)
    # alle Dateien aus dem Ordner BesetzungenJson loeschen
    listErrors.append(delete_files_in_directory(directory_path))


    #-------------Excelfile "Rennliste" in ein großes Dataframe "df" parsen-------------------------
    # usecols gibt die Excel Spalten an, aus denen der Dataframe erstellt wird
    #dtype=str macht aus allen Einträgen strings
    #header = none --> Spaltenbezeichnung wird auf B=1, C=2, D=3...
    df = pd.read_excel(ExcelFilePath, sheet_name="Rennlisten", index_col=None, header = None, usecols="B:K", dtype=str, engine="openpyxl")
    print(df.loc[0:15,1:9])  #Zeile, Spalte   --> printed die ersten paar Zeilen des Dataframes



    #-------------Dataframe "df" in einzelne kleine Dataframes(Rennen) zerteilen-------------------
    #Der nachfolgende Code sucht nach Zeilen(i) in der Excelliste, wo dieses Pattern erfüllt wird:
    #Pattern:
    #             B            C
    # i         Zahl        "Ergocup"
    # i+1       "Name"      "Vorname"
    #
    matchZahl = df.loc[:,1].str.match(r'.*[0-9][0-9]?', na=False)           #pruefen, ob in Spalte B Zahlen von 0 bis 99 enthalten sind
    matchName = df.loc[:,1].str.match('Name',na=False)                   #pruefen, ob in Spalte B "Name" enthalten ist
    matchVorname = df.loc[:,2].str.match('Vorname',na=False)                   #pruefen, ob in Spalte C "Vorname" enthalten ist


    # überall wo dieses Pattern erfüllt ist, wird das dataframe df zerteilt
    # Es ergeben sich einzelne Datasheets, welche jeweils ein Rennen beinhalten
    #Renn-Datasheets werden im Dictionary Rennen gespeichert. Key ist die Rennummer
    RennenFrames = {}
    ErstesFrame = True
    print(df)
    for i in range(len(df)-1):
        if matchZahl.loc[i] and matchName.loc[i+1] and matchVorname[i+1]:
            if ErstesFrame:
                ErstesFrame = False
                AnfangIndex = i
                RennNr = int(df.loc[i,1])
                continue
            EndIndex = i-1
            RennenFrames[RennNr] = df.loc[AnfangIndex:EndIndex, :]

            AnfangIndex = i

            #RennNr = int(df.loc[i,1])
            RennNr = df.loc[i,1]
    EndIndex = i-1
    RennenFrames[RennNr] = df.loc[AnfangIndex:EndIndex, :]


    #----------------Die einzelnen Rennframes als Excel-Dateien speichern--------------

    listOfFileNames = []
    for RennNr in RennenFrames:
        RennenFrames[RennNr].to_excel("./RennenFrames/RennenFrame_({0}).xlsx".format(RennNr)) 
        listOfFileNames.append("./RennenFrames/RennenFrame_({0}).xlsx".format(RennNr))
    with open(r"./RennenFrames/FileNames.json", "w") as file:
        json.dump(listOfFileNames, file, indent=4)
    print("Meldung:   Excel-file: \"Rennliste\" was sucessfully parsed to Panda Dataframes and saved as Files \"RennenFrame_().xlsx\"")
    print("\n\n")
    return listErrors

#test = exl_to_df(r"./ERGOCUP 2025.xlsx")