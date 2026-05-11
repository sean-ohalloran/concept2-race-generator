import pandas as pd


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
print(dict_VereinsAbkuerzungen)

#-------------------Excelfile mit dem Vereinsabkuerzungsverzeichnis aktualisiseren:
#dictionary to dataframe:
list_Vereinsnamen_Abkuerzungen = dict_VereinsAbkuerzungen.items()
print(list_Vereinsnamen_Abkuerzungen)
df_NewVereinsAbkuerzungen = pd.DataFrame(list_Vereinsnamen_Abkuerzungen, columns=['Abkuerzung', 'Name_Lang'])

df_NewVereinsAbkuerzungen['Ort'] = df["Ort"]

# Reorder the columns to match the desired output
df_NewVereinsAbkuerzungen = df_NewVereinsAbkuerzungen[['Abkuerzung', 'Ort', 'Name_Lang']]


with pd.ExcelWriter(r"./test12345.xlsx") as writer:
    # Write the DataFrame to the Excel file
    df_NewVereinsAbkuerzungen.to_excel(writer, startrow=3, index=True, sheet_name="nach_Orten", header=True)
    # Write the title and metadata at the top of the Excel file
    writer.sheets['nach_Orten'] = pd.DataFrame([['Abkürzungen der Vereinsnamen 2020', '', '']])
    writer.sheets['nach_Orten'] = pd.DataFrame([['Stand:', '16/01/2025', '']])
