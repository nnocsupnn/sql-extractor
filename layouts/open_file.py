import PySimpleGUI as sg
import os

def openFile(filePath):
    openFileLayout = [
        [sg.Titlebar("Report Extracted")],
        [sg.Text("File is extracted [" + filePath + "]")],
        [sg.Button('Open file'), sg.Button('Close')]
    ]

    openFileWindow = sg.Window('Report Extracted', layout=openFileLayout)

    while True:
        event, values = openFileWindow.read()
        if event in (sg.WIN_CLOSED, "Exit"):
            openFileWindow.close()
            break
        elif event == 'Close':
            openFileWindow.close()
            break
        elif event == "Open file":
            os.startfile(filePath)
            break

    openFileWindow.close()