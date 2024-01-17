import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
import xlsxwriter
import PySimpleGUI as sg
from def_PreFoglioPy import *
import time
import subprocess
import sys
import io


## INTERFACCIA 
sg.theme('DarkBrown1')   # Add a little color to your windows
# All the stuff inside your window. This is the PSG magic code compactor...
layout = [[sg.Text('PreFoglioPy', size = (40, 1), justification = 'center', font='Courier 15', text_color='White')],
          [sg.Text('importazione delle sollecitazioni max e min nel SuperFoglio', size = (75, 1), justification = 'center', font='Courier 8', text_color='White')],
          [sg.T(' '*2)],
          [sg.T('Source Folder Input '), sg.In(key='PathIn'), sg.FolderBrowse(target='PathIn')],
          [sg.T('Source Folder Output'), sg.In(key='PathOut'), sg.FolderBrowse(target='PathOut')],
          [sg.T(' '*1)],
          [sg.Text("NOTA 1: percorso della cartella contenente i dati e risultati esportati da MIDAS in {Folder Input}", size = (75, 2), justification='left', font='Courier 8', text_color='Grey')],
          [sg.Text("NOTA 2: in {Folder Output} indicare la cartella dove si sta lavorando cn il SuperFoglio", size = (75, 2), justification='left', font='Courier 8', text_color='Grey')],
          [sg.T(' '*1)],
          [sg.OK(size = (30, 1)), sg.Cancel(size = (30, 1))],
          [sg.T(' '*1)],
          [sg.Multiline("", size=(70, 5),autoscroll=True, background_color='black', text_color='white', key="output")] , #   per visualizzare gli output cmd
          [sg.T(' '*1)],
          [sg.Text('Autore:   Ing. Domenico Gaudioso', text_color='White'), sg.Text(' '*76)],
          [sg.Text('Contatti: d.gaudioso@matildi.com', text_color='White'), sg.Text(' '*76)],
          [sg.Text("per qualsiasi segnalazione rivolgersi all'autore"), sg.Text(' '*60)]]

# Create the Window
window = sg.Window('PreFoglioPy', layout, element_justification = 'c')

# Event Loop to process "events"
while True:             
    event, values = window.read()
    namePC = os.environ['COMPUTERNAME']
    pathSave = r'Z:\tools\script_ToPreFoglio\activity'
    orario = time.strftime("%H:%M:%S")
    data = time.strftime("%d/%m/%Y")
    testo = "{};{};{}\n".format(namePC, data, orario)

    if event in (sg.WIN_CLOSED, 'OK'):
        PathIn = values['PathIn']
        PathOut = values['PathOut']

        with open(os.path.join (pathSave, 'activity.txt'), 'a') as f:
            f.write(testo)

        #scriptPath = r'Z:\tools\script_ToPreFoglio\def_PreFoglioPy'
        #fName = 'Run_ExportOut_SuperFoglio'
        #pythonPath = r'Z:\tools\Portable Python-3.9.13 x64\App\Python\python.exe'
        #run_function(pythonPath, scriptPath, fName, PathIn, PathOut)

        captured_output = io.StringIO()
        captured_error = io.StringIO()
        sys.stdout = captured_output
        sys.stderr = captured_error
        
        try:
            Run_Export1Out_SuperFoglio(PathIn, PathOut)
        except Exception as e:
            print( 'No write variecose NTC')
            #captured_error.write(str(e))

        try:
            Run_Export2Out_SuperFoglio(PathIn, PathOut, metodo = 2)
        except Exception as e:
            print( 'No write variecoseFatica ')

        #Run_Export1Out_SuperFoglio(PathIn, PathOut)
        #Run_Export2Out_SuperFoglio(PathIn, PathOut)
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        output = captured_output.getvalue() + captured_error.getvalue()
        window["output"].update(output)

        #break

    elif event in (sg.WIN_CLOSED, 'Cancel'):
        break

window.close()

