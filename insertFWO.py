from openpyxl import load_workbook as load
import csv

workbook = "../2019/Gift Aid Analysis 2019.xlsx"
sheet_name = "directDebit"
csv_file = "../2019/AlexForExport13_09_19"

try:
    wb = load(workbook, data_only = True)
except PermissionError:
    print("You currently have the excel workbook open")
    raise

sheet = wb[sheet_name]
cashNames = "glen", "wood"

def insertDD(sheet, csv_file, cashNames):
    with open(csv_file+".csv") as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        #goes down list of names in csv file
        for row in csv_reader:
            j = 0
            while j <= 82:
                name = str(sheet.cell(column = 3, row = (j + 5)).value)
                #attempts to match name in excel file
                #with name in csv (all lower case) 
                if name.lower() in row[8].lower().split()\
                   and name.lower() not in cashNames:
                    month = int(row[9][3:5])
                    amountPaid = float(row[14])
                    #existing amount in the distination cell
                    currCellValue = 0.0
                    valStr = (str
                              (sheet.cell
                               (column = (5 + month), row = (j + 5)).value))
                    if valStr != "None":
                        currCellValue = float(valStr)
                    #adds amount to existing amount in appropriate cell
                    sheet.cell(column =(5 + month),
                              row = (j + 5),
                              value = (amountPaid + currCellValue))
                    break
                j += 1
    wb.save("Gift Aid Analysis 2019 FWO copy.xlsx")
    
insertDD(sheet, csv_file, cashNames)
