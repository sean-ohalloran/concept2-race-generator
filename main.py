import ExlToDf
import RennFramesToBesetzungsFrames
import BesetzungsFramesToJson
import BesungsFrameMergen
import CopyDirectory
import PySimpleGUI as sg

#-----GUI
sg.theme('Light Blue 2')

layout = [[sg.Text('Waehle die Exceldatei mit den Rennen, und den Ordner, wo die Rennen gespeichert werden, aus')],
          [sg.Text('Exceldatei', size=(12, 1)), sg.Input(), sg.FileBrowse()],
          [sg.Text('Rennen-Ordner', size=(12, 1)), sg.Input(), sg.FolderBrowse()],
          [sg.Submit(), sg.Cancel()]]

window = sg.Window('Excel zu Json.rac2 Dateien', layout)

event, values = window.read()
window.close()

ExcelFilePath = values[0]
RennFolderPath = values[1]

print(f'Du hast geclickt: {event}')
print(f'Du hast die Exceldatei ausgewaehlt: {ExcelFilePath}')
print(f'Du hast den Ordner ausgewaehlt: {RennFolderPath}')
sg.popup_auto_close(f'Du hast die Exceldatei ausgewaehlt: {ExcelFilePath}\nDas Programm läuft, bitte warten...')


listOfErrors = []
#--------------------Unterprogramme aufrufen:
#listOfErrors.append("\nexl_to_df:")
listOfErrors.extend(ExlToDf.exl_to_df(ExcelFilePath))
#listOfErrors.append("\nRennFramesToBesetzungsFrames:")
listOfErrors.extend(RennFramesToBesetzungsFrames.RennFramesToBesetzungsFrames())
#listOfErrors.append("\n\BesungsFrameMergen:")
listOfErrors.extend(BesungsFrameMergen.MergeBesetungsFrames())
#listOfErrors.append("\n\BesetzungsFramesToJson:")
listOfErrors.extend(BesetzungsFramesToJson.BesetzungsFramesToJson())
listOfErrors.extend(CopyDirectory.copy_directory(RennFolderPath))

sg.popup_no_titlebar(f'Das Programm ist durchgelaufen!\n Die fertigen JSON-Dateien finden sie in dem Ordner: \"{RennFolderPath}\"')

#---------Errors ausgeben:
print("\n\n\n\n-----------------------------------------------------------")
print("There were Errors with the follwoing file(s):")
for Error in listOfErrors:
    print(Error)
print("\n\n")

#-------------Errors im GUI Ausgeben:
#Error-List zu einem String machen
ErrorString = "List of Errors:\n(Zu den aufgelisteten Rennen wurden keine Dateien erstellt)\n\n"
for Error in listOfErrors:
        if not Error is None:
            ErrorString = ErrorString + Error + "\n"
sg.popup_no_titlebar(ErrorString)
    
