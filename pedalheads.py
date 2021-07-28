import openpyxl
import sys
import PySimpleGUI as sg
from functionF import *



#FILE = (r"c:\Users\rober\Desktop\list.xlsx")
#C:\Users\rober\OneDrive\Desktop

#C:\Users\Robert Liu\Desktop\list.xlsx
#C:\Users\Robert Liu\Desktop\test.xlsx

#variables for instructions
instruction1 = [sg.Text("Thank you for trying my (Robert's) automatic registration system \nWhen typing in a file path, the format should be similar to C:\\Users\\Name\\Etc\\file_name")]

#instruction2 = [sg.Text("When typing in a file path, the format should be similar to", r"C:\Users\Name\Etc\file_name")]

nextButton = [sg.Button("NEXT")]

layoutInstruction = [
    [
        instruction1,
        #instruction2,
        nextButton,
    ]
]
#variables with names and buttons
input_file_path1 = [sg.Text("INSERT FILE LOCATION"), sg.In(size=(55,1), enable_events=True, key = "-FILE1-")]
#button1 = [sg.Button("Submit file")]
input_file_path2 = [sg.Text("INSERT NEW FILE LOCATION"), sg.In(size=(50,1), enable_events=True, key = "-FILE2-")]

button = [sg.Button("Generate file")]

layout = [
    [
        input_file_path1,
        input_file_path2, 
        button,
    ]
]

#variables that are windows
window = sg.Window("AUTOMATIC REGISTRATION", layout, margins=(300,100))
instructionWindow = sg.Window("BEGINNING",layoutInstruction, margins=(300,100))

while True:
    eventIn, valueIn = instructionWindow.read()
    if eventIn == "NEXT":
        break
    if eventIn == sg.WIN_CLOSED:
        exit()
instructionWindow.close()


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

try:
    orderSorting(str(filepath))
except:
    raise Exception("Not a valid input, REMINDER you need a file AND the path to the file")
    
exit()

