import openpyxl
import sys
import pandas
from functionF import *

#FILE = (r"c:\Users\rober\Desktop\list.xlsx")

#C:\Users\Robert Liu\Desktop\list.xlsx
#C:\Users\Robert Liu\Desktop\test.xlsx


FILE = input("ENTER PATH OF FILE (WITH FILE):")

filepath = input("ENTER NEW FILE (WITH PATH):")

try:
    numberOfReg(str(FILE))
except:
    raise Exception("Not a valid input, REMINDER you need a file AND the path to the file")

try:
    half(str(FILE)) 
except:
    raise Exception("Not a valid input, REMINDER you need a file AND the path to the file")

try:
    allday(str(FILE))
except:
    raise Exception("Not a valid input, REMINDER you need a file AND the path to the file")


try:
    copyRow(str(FILE),filepath)
except:
    raise Exception("Not a valid input, REMINDER you need a file AND the path to the file")

exit()

