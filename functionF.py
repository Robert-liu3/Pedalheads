#reading excel files
import openpyxl
import sys



#hihihihihi
#hello

def numberOfReg(fileloc):
    #open workbook
    wb = openpyxl.load_workbook(fileloc)

    sheet = wb.get_sheet_by_name('Class List')

    sheet['A1'].value

    sh = wb.active
    numReg = 0

    for i in range(7, sh.max_row+7):
        cell = 'A' + str(i)
        if isinstance(sheet[cell].value, int) == True :
            numReg += 1
            #print(sheet[cell].value)  
    print("TOTAL NUMBER OF REGISTRATIONS: ", numReg)


def half(fileloc):
    wb = openpyxl.load_workbook(fileloc)

    sheet = wb.get_sheet_by_name('Class List')


    sh = wb.active
    halfam1 = 0
    halfam2 = 0
    halfam3 = 0
    halfam4 = 0
    halfam5 = 0
    halfam6 = 0

    halfpm1 = 0
    halfpm2 = 0
    halfpm3 = 0
    halfpm4 = 0
    halfpm5 = 0
    halfpm6 = 0

    for i in range(7, sh.max_row+7):
        cell1 = 'I' + str(i)
        cell2 = 'H' + str(i)
        
        if sheet[cell1].value == "Bike Half (main price): 09:00 AM - 12:00 PM":
            if sheet[cell2].value == "Level 1 - Newbees":
                halfam1+= 1
            if sheet[cell2].value == "Level 2 - Advanced Newbees":
                halfam2 += 1
            if sheet[cell2].value == "Level 3 - Pedalheads":
                halfam3 += 1
            if sheet[cell2].value == "Level 4 - Advanced Pedalheads":
                halfam4 += 1
            if sheet[cell2].value == "Level 5 - Gearheads":
                halfam5 += 1
            if sheet[cell2].value == "Level 6 - Treadheads":
                halfam6 += 1
        if sheet[cell1].value == "Bike Half (main price): 01:00 PM - 04:00 PM":
            if sheet[cell2].value == "Level 1 - Newbees":
                halfpm1+= 1
            if sheet[cell2].value == "Level 2 - Advanced Newbees":
                halfpm2 += 1
            if sheet[cell2].value == "Level 3 - Pedalheads":
                halfpm3 += 1
            if sheet[cell2].value == "Level 4 - Advanced Pedalheads":
                halfpm4 += 1
            if sheet[cell2].value == "Level 5 - Gearheads":
                halfpm5 += 1
            if sheet[cell2].value == "Level 6 - Treadheads":
                halfpm6 += 1
    print("========================================\n========================================")
    print("NUMBER OF HALF DAY AM LEVEL 1:", halfam1, "registrations")
    print("NUMBER OF HALF DAY AM LEVEL 2:", halfam2, "registrations")
    print("NUMBER OF HALF DAY AM LEVEL 3:", halfam3, "registrations")
    print("NUMBER OF HALF DAY AM LEVEL 4:", halfam4, "registrations")
    print("NUMBER OF HALF DAY AM LEVEL 5:", halfam5, "registrations")
    print("NUMBER OF HALF DAY AM LEVEL 6:", halfam6, "registrations")
    print("========================================\n========================================")
    print("NUMBER OF HALF DAY PM LEVEL 1:", halfpm1, "registrations")
    print("NUMBER OF HALF DAY PM LEVEL 2:", halfpm2, "registrations")
    print("NUMBER OF HALF DAY PM LEVEL 3:", halfpm3, "registrations")
    print("NUMBER OF HALF DAY PM LEVEL 4:", halfpm4, "registrations")
    print("NUMBER OF HALF DAY PM LEVEL 5:", halfpm5, "registrations")
    print("NUMBER OF HALF DAY PM LEVEL 6:", halfpm6, "registrations")
def allday(fileloc):
    wb = openpyxl.load_workbook(fileloc)

    sheet = wb.get_sheet_by_name('Class List')

    sh = wb.active

    all1 = 0
    all2 = 0
    all3 = 0
    all4 = 0
    all5 = 0
    all6 = 0

    for i in range(7, sh.max_row+7):
        cell1 = 'I' + str(i)
        cell2 = 'H' + str(i)

        if sheet[cell1].value == "Bike All (main price): 09:00 AM - 04:00 PM":
            if sheet[cell2].value == "Level 1 - Newbees":
                all1+= 1
            if sheet[cell2].value == "Level 2 - Advanced Newbees":
                all2 += 1
            if sheet[cell2].value == "Level 3 - Pedalheads":
                all3 += 1
            if sheet[cell2].value == "Level 4 - Advanced Pedalheads":
                all4 += 1
            if sheet[cell2].value == "Level 5 - Gearheads":
                all5 += 1
            if sheet[cell2].value == "Level 6 - Treadheads":
                all6 += 1
    print("========================================\n========================================")
    print("NUMBER OF ALL DAY LEVEL 1:", all1, "registrations")
    print("NUMBER OF ALL DAY LEVEL 2:", all2, "registrations")
    print("NUMBER OF ALL DAY LEVEL 3:", all3, "registrations")
    print("NUMBER OF ALL DAY LEVEL 4:", all4, "registrations")
    print("NUMBER OF ALL DAY LEVEL 5:", all5, "registrations")
    print("NUMBER OF ALL DAY LEVEL 6:", all6, "registrations")

def copyRow(fileloc,filepath):
    wb = openpyxl.load_workbook(fileloc)

    sheet = wb.get_sheet_by_name('Class List')

    sh = wb.active

    wb2 = openpyxl.Workbook()

    #filepath = (r"c:\Users\rober\Desktop\test2.xlsx")

    ws2 = wb2.active

    ws3 = wb2.create_sheet(0)

    ws3.title = "Class List"

    #max row and max column
    mr = sh.max_row
    mc = sh.max_column 
    for j in range (1, mc + 1):
        c = sh.cell(row = 5, column = j)
        ws2.cell(row = 5, column = j).value = c.value
        #for i in range(7, sh.max_row+7):
            #cell1 = 'I' + str(i)
            #cell = 'A' + str(i)
            #if isinstance(sheet[cell].value, int) == True :
                #if sheet[cell1].value == "Bike All (main price): 09:00 AM - 04:00 PM" or sheet[cell1].value == "Bike Half (main price): 09:00 AM - 12:00 PM":
                    #print(sheet[cell].value)
    
    for i in range(7, mr+7):
        orderNum = 'A' + str(i)
        regTime = 'I' + str(i)
        if isinstance(sheet[orderNum].value, int) == True :
            if sheet[regTime].value == "Bike All (main price): 09:00 AM - 04:00 PM" or sheet[regTime].value == "Bike Half (main price): 09:00 AM - 12:00 PM": 
                for j in range (1, mc + 1):
                    for x in range(7, mr+7):
                        c = sh.cell(row = i, column = j)
                        ws2.cell(row = i, column = j).value = c.value
    
    
    
    
    wb2.save(filepath)

