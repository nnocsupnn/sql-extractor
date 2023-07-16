import PySimpleGUI as sg
import pandas as pd
import pyodbc as db
import os
import datetime as dt
import random


def getResults(query):
    cnxn = db.connect('Driver={SQL Server};'
                                'Server=localhost;'
                                'Database=PYCONVERTER;'
                                'UID=SA;'
                                'PWD=yourStrong(!)Password;', autocommit=False)

    df = pd.read_sql(query, cnxn)
    return df

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


def save(format, df, filePath):
    if format == "xlsx":
        df.to_excel(filePath, index=False)
    elif format == "csv":
        df.to_csv(filePath, index=False)
    elif format == "json":
        df.to_json(filePath)
    elif format == "xml":
        df.to_xml(filePath)
    else:
        df.to_excel(filePath, index=False)

theme_name_list = sg.theme_list()
randomTheme = theme_name_list[random.randint(0, len(theme_name_list) - 1)]

sg.theme('SandyBeach')   # Add a touch of color
# All the stuff inside your window.
layout = [  
            [sg.Text('SQL Script'), sg.Push(), sg.Text(key="resultCount", text_color='red', text="0"), sg.Text(text='row(s)', text_color='red')],
            [sg.Multiline(size=(60, 5), key="script", default_text='select * from persons', enter_submits=True, focus=True, font=('Courier New', 12),)],
            [sg.Push(), sg.Button("Connection Settings", tooltip="Coming soon", button_color="black")],
            [sg.Text("Extraction Option")],
            [sg.FolderBrowse(target="folderPath", tooltip="Select the output folder.", button_text="Select Folder", initial_folder="./", enable_events=True, key="chooseFolder"), sg.Input('./', key="folderPath")],
            [sg.T("")],
            [
                sg.Submit('Load', button_color="green"), 
                sg.Exit('Close'), 
                sg.Push(), 
                sg.Combo(disabled=True, default_value='xlsx', values=['xlsx', 'json', 'xml', 'csv'], key="format", tooltip="Select format", size=(10, 5)), 
                sg.Submit('Generate', tooltip="Choose a folder to save then submit.", disabled=True)] 
        ]

window = sg.Window('Report Generator', layout=layout, finalize=True)

def handleFolder(event):
    print(event)


df = None
while True:
    try:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "Exit"):
            window.close()
            break

        if event == 'Close':
            window.close()
            break

        if event == 'Load':
            query = values['script']
            df = getResults(query)
            resultCount = df.shape[0]
            window['resultCount'].update(resultCount)
            if resultCount > 0:
                window['Generate'].update(disabled=False)
                window['format'].update(disabled=False)

        if event == 'Generate':
            folderPath = values['folderPath']
            if folderPath == None:
                folderPath = chooseFolder()

            filePath = os.path.join(folderPath, 'report_' + dt.datetime.now().strftime("%Y_%m_%d_%H_%M_%S") + '.' + values['format'])
            save(values['format'], df, filePath)
            openFile(filePath)
    except Exception as e:
        sg.popup_error_with_traceback("Something went wrong.", e)

# try:
#     df = getResults(query)

#     if df.empty:
#         sg.popup_auto_close(title="Empty", message="Result is empty.")
#     else:
#         event, values = window.read()
#         print(df.shape[0])
#         window['resultCount'].update(df.shape[0])
#         fileChooser = [
#             [sg.Text("Select folder path"),
#             sg.Input(key="folderPath", disabled=True),
#             sg.FolderBrowse(target="folderPath")],
#             [sg.Button("Submit")],
#         ]

#         fileChooserWindow = sg.Window('Choose folder to SAVE', fileChooser)

#         filename = ""
#         while True:
#             event, values = fileChooserWindow.read()
#             if event in (sg.WIN_CLOSED, "Exit"):
#                 break
#             elif event == "Submit":
#                 filename = values['folderPath']
#                 break

#         fileChooserWindow.close()
#         filePath = os.path.join(filename, 'report_' + dt.datetime.now().strftime("%Y_%m_%d_%H_%M_%S") + '.xlsx')
#         df.to_excel(filePath, index=False)
#         os.startfile('report.xlsx')
# except Exception as e:
#     sg.popup_error_with_traceback("Something went wrong.", e)