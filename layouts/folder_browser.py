import PySimpleGUI as sg

def chooseFolder():
    fileChooser = [
            [sg.Text("Select folder path"),
            sg.Input(key="folderPath", disabled=True),
            sg.FolderBrowse(target="folderPath")],
            [sg.Push(), sg.Button("OK")],
        ]

    fileChooserWindow = sg.Window('Choose folder to SAVE', fileChooser)

    filename = ""
    while True:
        event, values = fileChooserWindow.read()
        if event in (sg.WIN_CLOSED, "Exit"):
            break
        elif event == "OK":
            filename = values['folderPath']
            break

    fileChooserWindow.close()
    
    return filename