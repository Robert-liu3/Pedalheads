import openpyxl
import sys
from functionF import *

#FILE = (r"c:\Users\rober\Desktop\list.xlsx")
#C:\Users\Robert Liu\Desktop
#PULL REQUEST TEST


FILE = input("ENTER PATH OF FILE (WITH FILE):")
filepath = input("ENTER NEW FILE (WITH PATH):")

numberOfReg(str(FILE))
#hello
half(str(FILE)) 

allday(str(FILE))

copyRow(str(FILE),filepath)

exit()

