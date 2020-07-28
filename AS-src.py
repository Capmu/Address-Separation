#--------------------------------------------------------------------------------------------------------------------
#   Define a Python library.
#--------------------------------------------------------------------------------------------------------------------
import xlrd #for read excel workbook [READ ONLY!]

#--------------------------------------------------------------------------------------------------------------------
#   Define a constant variables.
#--------------------------------------------------------------------------------------------------------------------
xlsName = "Pickuplist 2020-06-21.xls"
xlsPath = "Excel files/" + xlsName

#--------------------------------------------------------------------------------------------------------------------
#   Open files.
#--------------------------------------------------------------------------------------------------------------------
workbookVar = xlrd.open_workbook(xlsPath)   #for [xlrd] libraly
readerVar = workbookVar.sheet_by_index(0)

#--------------------------------------------------------------------------------------------------------------------
#   Read contents.
#--------------------------------------------------------------------------------------------------------------------
testingRow = 15
rowLength = len(str(readerVar.cell(testingRow,5)))
strTemp = ""
indexPicker = rowLength - 11
temp = str(readerVar.cell(testingRow,5))[indexPicker]
print("-------------------------------------------------------------------------")
for i in range(3):
    while(temp != ' '):
        strTemp = strTemp + temp
        indexPicker -= 1
        temp = str(readerVar.cell(testingRow,5))[indexPicker]
    print(strTemp[::-1]) #[::-1] will reverse text in a string variable.
    strTemp = ""
    indexPicker -= 1
    temp = str(readerVar.cell(testingRow,5))[indexPicker]
print(str(readerVar.cell(testingRow,4).value))
