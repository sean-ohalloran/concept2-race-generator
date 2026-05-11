import re
import pandas as pd
import json


def neuer_eintrag_Abkuerzungsverzeichnis():
    print("neuen Eintrag erstellen:\n")
    Vereinsname_Lang = input("Wie lautet der volle Vereinsname?")
    abkuerzung = input("Wie lautet die Abkuerzung?")
    # dem Abkuerzungsdictionary hinzufuegen:
    dict_VereinsAbkuerzungen[Vereinsname_Lang] = abkuerzung
    return abkuerzung


ExcelFilePath_AbkVerz = "./Abkürzungen_Vereinsnamen_2025.xlsx"
SheetName_AbkVerz = "nach_Orten"

#-------------Excelfile "Abkürzungen_Vereinsnamen_2025.xlsx" in Dataframe parsen-------------------------
# usecols gibt die Excel Spalten an, aus denen der Dataframe erstellt wird
#dtype=str macht aus allen Einträgen strings
#header = 2 --> Spaltenbezeichnung wird gesetzt auf:     index	Abkuerzung	Ort	Name_Lang
df = pd.read_excel(ExcelFilePath_AbkVerz, sheet_name=SheetName_AbkVerz, index_col=0, header = 2, usecols="A:D", engine="openpyxl")
#print(df.loc[:,:])  #Zeile, Spalte   --> printed die ersten paar Zeilen des Dataframes

#--------------------- Dataframe in Dictionary umwandeln
#key ist voller vereinsname ("Name_Lang")
#Value ist abkuerzung ("Abkuerzung")
dict_VereinsAbkuerzungen = dict(zip(df.loc[:,"Name_Lang"], df.loc[:,"Abkuerzung"]))




#---------------Excel-Datei Meldeliste (Ergocup 2025.xlsx) oeffnene:
ExcelFilePath_MeldeListen = r"./ERGOCUP 2025.xlsx"
SheetName_MeldeListen = "Eingabe Meldungen"
df_MeldeListe = pd.read_excel(ExcelFilePath_MeldeListen, sheet_name=SheetName_MeldeListen, index_col=None, header = 0, usecols="A:J", engine="openpyxl")
#print(df_MeldeListe.loc[:20,"Name":"Vorname"])  #Zeile, Spalte   --> printed die ersten paar Zeilen des Dataframes

#----------------------.rac2(json) datei laden und Vereinsnamen aendern:
rac2FileName = input("\n\nName eingeben:\t")
rac2FilePath = r"./BesetzungenJson/" + rac2FileName + ".rac2"
with open(rac2FilePath, "r", encoding='utf-8') as file:
        dict_Besetzung = json.load(file)
        list_boats = dict_Besetzung["race_definition"]["boats"]
        for boat in list_boats:
                splittedName = boat["name"].split(',')
                if len(splittedName)==2:
                    Name = splittedName[0].strip()
                    Vorname = splittedName[1].strip()
                else:
                    #der Name kann nicht ordnungsgemaess gesplitted werden
                    #einfach ueberspringen
                    print("ERROR: Kann nicht gesplitted werden")
                    print(splittedName)
                    continue
                    
                print("Name, Vorname:\t", Name, ", ", Vorname)
                #--------------Anhand von Name und Vorname den vollen Vereinsnamen Raussuchen:
                df_passendeVereinsnamen = df_MeldeListe[df_MeldeListe['Name'].str.contains(Name) & df_MeldeListe['Vorname'].str.contains(Vorname)]['Verein']
                df_passendeVereinsnamen = df_passendeVereinsnamen.drop_duplicates()       # remove duplicates
                if len(df_passendeVereinsnamen) == 1:
                    print("Vereinsname gefunden!:\t", df_passendeVereinsnamen.iloc[0])
                    Vereinsname_Lang = df_passendeVereinsnamen.iloc[0]
                elif len(df_passendeVereinsnamen) == 0:
                    print("Problem: Name nicht in Meldeliste gefunden!")
                    Vereinsname_Lang = input("bitte gebe den Vereinsnamen ein:")
                else:
                    print("mehrere Namen mit unterschiedlichen Vereinen gefunden!")
                    print(df_passendeVereinsnamen)
                    Vereinsname_Lang = input("bitte gebe den Vereinsnamen ein:")


                #-------------------passende Abkuerzung raussuchen:
                if Vereinsname_Lang in dict_VereinsAbkuerzungen:
                    #Vereinsname ist im Dictionary vorhanden
                    abkuerzung = dict_VereinsAbkuerzungen[Vereinsname_Lang]
                    print("Vereinsname: ", Vereinsname_Lang, "eindeutig zugeordnet:", abkuerzung)
                    
                else:
                    #Vereinsname ist nicht im Dictionary vorhanden
                    print("\n-----------\nVereinsname ist nicht im Abkuerzungsverzeichnis enthalten! Es muss eine neue hinzugefuegt werden:")
                    print("\t\tRenn-Name:\t", dict_Besetzung["race_definition"]["name_long"])
                    print("\t\taktuelles boot\t", boat)
                    print("\t\tname:\t", boat["name"])
                    abkuerzung = neuer_eintrag_Abkuerzungsverzeichnis()
 

                #Vereinsnamen-Anpassung vornehmen:
                boat["affiliation"] = abkuerzung


#------------------Geaenderte .rac2(json)-Datei speichern:
with open(rac2FilePath,  'w', encoding='utf-8') as file:
    json_Besetzungsframe = json.dumps(dict_Besetzung, indent=4)
    file.write(json_Besetzungsframe)



#-------------------Excelfile mit dem Vereinsabkuerzungsverzeichnis aktualisiseren:
#dictionary to dataframe:
list_Vereinsnamen_Abkuerzungen = dict_VereinsAbkuerzungen.items()
df_NewVereinsAbkuerzungen = pd.DataFrame(list_Vereinsnamen_Abkuerzungen, columns=['Name_Lang', 'Abkuerzung'])

df_NewVereinsAbkuerzungen['Ort'] = df["Ort"]

# Reorder the columns to match the desired output
df_NewVereinsAbkuerzungen = df_NewVereinsAbkuerzungen[['Abkuerzung', 'Ort', 'Name_Lang']]


with pd.ExcelWriter(r"./test12345.xlsx") as writer:
    # Write the DataFrame to the Excel file
    df_NewVereinsAbkuerzungen.to_excel(writer, startrow=3, index=True, sheet_name="nach_Orten", header=True)
    # Write the title and metadata at the top of the Excel file
    writer.sheets['nach_Orten'] = pd.DataFrame([['Abkürzungen der Vereinsnamen 2020', '', '']]) #funktioniert nicht 
    writer.sheets['nach_Orten'] = pd.DataFrame([['Stand:', '16/01/2025', '']])                  #funktioniert nicht 













kommentar = """
                #---------------Liste mit nicht eindeutig zuordnenbaren Vereinsnamen:
                list_nichtEindeutig = ["Rude","Stut"]
                if Vereinsname_Lang in list_nichtEindeutig:
                    #Vereinsname ist nicht eindeutig einer Abkuerzung zuorndbar
                    print("\nvereinsname nicht eindeutig zuordnenbar. ")
                    print("\t\tRenn-Name:\t", dict_Besetzung["race_definition"]["name_long"])
                    print("\t\taktuelles boot\t", boat)
                    print("\t\tname:\t", boat["name"])

                    neuOderNicht = input("muss ein neuer Eintrag erstellt werden? (y/n)")
                    while True:
                        if neuOderNicht =="y":
                            abkuerzung = neuer_eintrag_Abkuerzungsverzeichnis()
                            break
                        elif neuOderNicht =="n":
                            print("es wird kein neuer eintrag in das Abkuerzungsverzeichnis angelegt.")
                            abkuerzung = input("Wie lautet die Abkuerzung?")
                            break
                        """


