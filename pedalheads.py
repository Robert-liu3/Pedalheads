import openpyxl
import sys
import pandas
import PySimpleGUI as sg
from functionF import *

#FILE = (r"c:\Users\rober\Desktop\list.xlsx")

#C:\Users\Robert Liu\Desktop\list.xlsx
#C:\Users\Robert Liu\Desktop\test.xlsx


#variables with names and buttons
input_file_path1 = [sg.Text("INSERT FILE LOCATION"), sg.In(size=(55,1), enable_events=True, key = "-FILE1-")]
#button1 = [sg.Button("Submit file")]
input_file_path2 = [sg.Text("INSERT NEW FILE LOCATION"), sg.In(size=(50,1), enable_events=True, key = "-FILE2-")]
button = [sg.Button("Generate file")]

layout = [
    [
        input_file_path1,
        input_file_path2, button,
    ]
]
#variables that are windows
window = sg.Window("AUTOMATIC REGISTRATION", layout, margins=(300,100))
while True:
    event, values = window.read()
    if event == "Generate file":
        file = values["-FILE1-"]
        filepath = values["-FILE2-"]
        break
    if event == sg.WIN_CLOSED:
        exit()
window.close()


#FILE = input("ENTER PATH OF FILE (WITH FILE):")

#filepath = input("ENTER NEW FILE (WITH PATH):")

try:
    numberOfReg(str(file))
except:
    raise Exception("Not a valid input, REMINDER you need a file AND the path to the file")

try:
    half(str(file)) 
except:
    raise Exception("Not a valid input, REMINDER you need a file AND the path to the file")

try:
    allday(str(file))
except:
    raise Exception("Not a valid input, REMINDER you need a file AND the path to the file")


try:
    copyRow(str(file), str(filepath))
except:
    raise Exception("Not a valid input, REMINDER you need a file AND the path to the file")

exit()

