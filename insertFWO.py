from openpyxl import load_workbook as load
import csv

workbook = "Gift Aid Analysis 2019.xlsx"
sheet_name = "directDebit"

try:
    wb = load(workbook, data_only = True)
except PermissionError:
    print("You currently have the excel workbook open")
    raise

sheet = wb[sheet_name]

#test
#sheet.cell(row = 5, column = 6, value = "hello")
#wb.save("Gift Aid Analysis 2019 copy.xlsx")
csv_file = "AlexForExport13_09_19"

with open(csv_file+".csv") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    #goes down list of names in csv file
    for row in csv_reader:
        j = 0
        while j <= 82:
            name = str(sheet.cell(column = 3, row = (j + 5)).value)
            if j == 82:
                print(row[8]," not found")
            #attempts to match name in excel file
            #with name in csv (all lower case) 
            elif name.lower() in row[8].lower().split():            
#                print(row[8]+" found in row "+str(j+5))
                break
            j += 1
            
        

    
##def getTemps(csv_reader):
##    """Extract temp values from csv reader
##    return as list
##    """
##    line_count = 0
##    temps = []
##    for row in csv_reader:
##        if line_count == 0:
##            line_count += 1
##        else:
##            temps.append(float(row[1]))
##            line_count += 1
##    return temps[::-1] #reverse order 
##
##def plotTimeVsTemp(file):
##    with open(file+".csv") as csv_file:
##        csv_reader = csv.reader(csv_file, delimiter=',')
##        line_count = 0
##        temps = getTemps(csv_reader)

##    for i in range(10):
##        print(temps[i])
##
##    xVals = [] #mins
##    for i in range(2295):
##        xVals.append(i*0.5)
##
##    pylab.plot(xVals,temps[:2295], label = file)
##    pylab.title("Water bath temp during sonication")
##    pylab.xlabel("time (mins)")
##    pylab.ylabel("temp, Celcius")
##    pylab.legend()
