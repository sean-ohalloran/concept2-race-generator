import pandas as pd
import numpy as np
import os
import re
import json

#-------------------Programmbeschreibung----------------------
#macht aus den einzelnen Panadaframes der Rennen pandaframes der Besetzungen
#Eine Besetzung beseteht aus allen Teilnehmern, die gleichzeitig auf den Ergos sitzen


def df_to_dict(df):
    #wandelt ein Dataframe in ein Dictionary um
    #das Dataframe muss aufgebaut sein, wie in der Muster-Excel-Datei
    dictGes = {}
    startColumn = 1     #gibt die Spalte an, auf der sich die Keys der untersten Ebene befinden

    #alle Key-Value-Pairs auf der untersten Ebene in das Dictionary einfuegen
    dictNonNested = kvp_nonnested(df, startColumn)
    dictGes.update(dictNonNested)     #append dictNonNested to dictGes

    #Array in Dictionary umwandeln und in das dictGes einfuegen
    dictArray = arrays(df, startColumn)
    dictGes.update(dictArray)

    return dictGes

def kvp_nonnested(df, startColumn):
    #Key-Value-Pairs non-nested
    #die Key-Value-Pairs auf der untersten Ebene als dictionary returned
    dfNonNested = df.loc[(df[startColumn+1] != "[") & (df[startColumn].notna())]                #Alle rows, die nicht zu einem Array gehoeren rausfiltern
    dictNonNested = dict(zip(dfNonNested.loc[:,startColumn], dfNonNested.loc[:,startColumn+1])) #Das Dataframe zum Dictionary umwandeln 
    return dictNonNested

def arrays(df, startColumn):
    #ein Dataframe mit array wird heruntergebrochen und als dictionary returned
    #das returned Dictionary hat nur ein key-Value-Pair: Key: name des Arrays    Value: Eintraege des Arrays
    listArrayDf=[]          #list of Dataframes in the Array
    listArrayDicts=[]       #list of Dictionaries made out of the Array-Elements
    dictArray = {}          #dictionary that will be returned

    # das Array wird aus dem gesamt-Dataframe herausgefiltert
    # es wird in einen eigenes Dataframe "dfArray" geschrieben
    arrayStartRow = df.loc[(df[startColumn+1] == "[")].index[0] +1      #die Row mit "[" wird nicht mit reingezählt
    arrayEndRow = df.loc[(df[startColumn+1] == "]")].index[0] -1        #die Row mit "]" wird nicht mit reingezählt
    arrayName = df.loc[arrayStartRow-1, startColumn]                    #sucht den Namen des Arrays raus
    dfArray = df.loc[arrayStartRow:arrayEndRow,startColumn+1:]

    #Array-Frame in die einzelnen Array-Items-Frames aufteilen (Items sind die durch Komma getrennten Eintraege)
    #Einzeneln Item-Frames werden in die Liste listArrayDf geschrieben
    groupedArray = dfArray.groupby(startColumn+1)
    for ArrayNum, DfArrayItem in groupedArray:
        listArrayDf.append(DfArrayItem)

    #jetzt wird jedes Item-Frame einzelnd behandelt:
    #   die Key-Value-Pairs auf der Untersten Ebene werden mit der Funktion kvp_nonnested() in das Dictionary dictArrayItem hinzugefuegt
    #   dann wird geprueft, ob ein verschateltes Array vorliegt. Wenn ja, wird die funktion arrays() von sich selbst aufgerufen
    for DfArrayItem in listArrayDf:
        dictArrayItem = {}      #dictionary of the current array-item
        dictArrayitemNonNested = kvp_nonnested(DfArrayItem, startColumn+2)
        dictArrayItem.update(dictArrayitemNonNested)

        #zunaechst pruefen, ob sich ein Array in dem ArrayItemDataframe befindet:
        #       (Check if the character '[' exists at all in the Column)
        if DfArrayItem[startColumn+3].str.contains(r'\[', regex=True).any()==False:
            listArrayDicts.append(dictArrayItem)
            continue
        dictArrayItemArray = arrays(DfArrayItem, startColumn+2)
        dictArrayItem.update(dictArrayItemArray)

        listArrayDicts.append(dictArrayItem)

        
    #es wird ein Dictionary mit nur einem Key-Value-Pair returned:
    #Key: Name des Arrays          Value: liste der Array-Elemente
    #Die einzelnen Array-Elemente sind dabei jeweils Dictionaries
    dictArray[arrayName] = listArrayDicts
    return dictArray

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
     return ("Error occurred while deleting JSON files from "+directory_path)

#-----------------------------------

listErrors = []
def BesetzungsFramesToJson():
    print("\n\n-------------------------------function: BesetzungsFramesToJson was called:\n\n")
    #-------------------- neuen ornder erstellen, falls noch nicht vorhanden
    directory_path = r"./BesetzungenJson" 
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)
        print("new direcory was created:   "+ directory_path)
    #alle Dateien aus dem Ordner BesetzungenJson loeschen
    listErrors.append(delete_files_in_directory(directory_path))

    #------------------Excelfile in Pandaframe laden--------------
    with open(r"./BesetzungenFrames/FileNames.json", "r") as file:
        ListOfFileNames = json.load(file)
    
    for BesetzungsNum in range(len(ListOfFileNames)):
        try:
            dateiNameBesFrame = ListOfFileNames[BesetzungsNum]
            # usecols gibt die Excel Spalten an, aus denen der Dataframe erstellt wird
            #dtype=str macht aus allen Einträgen strings
            #header = none --> Spaltenbezeichnung wird festgelegt auf: B=1, C=2, D=3...
            BesetzungsFrame_1 = pd.read_excel(r"./BesetzungenFrames/"+dateiNameBesFrame, index_col=0, header = None, usecols="A:G", engine="openpyxl")
            BesetzungsFrame_1 = BesetzungsFrame_1.replace('\"\"', '', regex=True)       #sodass cellen in Excel-file mit "" als empty string gelesen werden
            BesetzungsFrame_1 = BesetzungsFrame_1.replace('false', False, regex=True)   #sodass Excel Zellen mit "false" als boolean-Value False gelesen werden
            BesetzungsFrame_1 = BesetzungsFrame_1.replace('true', True, regex=True)     #sodass Excel Zellen mit "true" als boolean-Value True gelesen werden
            #print(BesetzungsFrame_1.loc[:,:])  #Zeile, Spalte   --> printed das Dataframe


            #--------------------- Key-Value-Pairs fuer intere Zwecke--------------
            #die Einträge bis zum Tilde-Zeichen sollen in ein Seperates dicitonary "intern_dict"
            intern_dict = dict(zip(BesetzungsFrame_1.loc[0:4,1], BesetzungsFrame_1.loc[0:4,2]))



            #--------------------- Dataframe in Dictionary umwandeln
            dataframe = BesetzungsFrame_1.loc[6:,:]
            dict_BesetungsFrame = {"race_definition": df_to_dict(dataframe)}


            #---------------------Dictionary in Json umwandeln
            json_Besetzungsframe = json.dumps(dict_BesetungsFrame, indent=4)
            #print(json_Besetzungsframe)


            #--------------------Json-String als File speichern:
            
            pattern = r"r[0-9]+A?[0-9]?_?r?[0-9]?[0-9]?A?[0-9]?"  # Looks for 'r' followed by one or more alphanumeric characters
            # Apply the regex pattern to each filename and extract the match
            bez = re.findall(pattern, dateiNameBesFrame)[0]

            filePath = os.path.join(r".\BesetzungenJson", "BesetzungJson_("+bez+").rac2")
            with open(filePath, 'w', encoding='utf-8') as file:
                file.write(json_Besetzungsframe)

            print("Meldung:   Excel-file: \"",dateiNameBesFrame,"\" was sucessfully parsed to Json and saved as File \"BesetzungJson_("+bez+").rac2\"")
        except Exception as e:
            listErrors.append(dateiNameBesFrame+"          "+str(e))
            print(e)

    print("\n")
    if len(listErrors) == 0:
        print("Meldung: sucessfully parsed to Json and saved as Files")
    else:
        print("There were Errors with thte follwoing file(s):")
        for Error in listErrors:
            print(Error)
    print("\n\n")
    return listErrors