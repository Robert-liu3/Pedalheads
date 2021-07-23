import openpyxl
import sys
from functionF import *

#FILE = (r"c:\Users\rober\Desktop\list.xlsx")
#C:\Users\Robert Liu\Desktop\list.xlsx
#C:\Users\Robert Liu\Desktop\test.xlsx

FILE = input("ENTER PATH OF FILE (WITH FILE):")
filepath = input("ENTER NEW FILE (WITH PATH):")

numberOfReg(str(FILE))

half(str(FILE)) 

allday(str(FILE))


copyRow(str(FILE),filepath)

exit()

