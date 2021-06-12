import xlrd
from re import search 
 
def findCell(sh, searchedValue):
    for row in range(sh.nrows):
        for col in range(sh.ncols):
            myCell = sh.cell(row, col)
            if searchedValue in str(myCell.value):
                return (row)
    return -1
def import_drawinglist():
    global drawings_tosearch
    drawings = open("files"+".txt", "r")
    lines = drawings.readlines()
    drawings_tosearch = []
    for element in lines:
        drawings_tosearch.append(element.strip())
    drawings.close()
def convertTuple(tup): 
    str =  ''.join(tup) 
    return str
import_drawinglist()
ssheet = input("Enter Spreadsheet Number")
for element in drawings_tosearch:
    searchedValue=element

    for sh in xlrd.open_workbook(ssheet +".xls").sheets():  
            result = findCell(sh, searchedValue)
            if result != -1:
                print(searchedValue, (result+1))
                f=open("Locations.txt", "a")
                locn=open("Line.txt", "a+")
                result=result+1
                result=str(result)
                locn.write(result)
                locn.write("\n")
                d="Drawing:,"
                l="Location:,"
                text=(d+searchedValue,",", l , result)
                txt=convertTuple(text)
                f.write(txt)
                f.write("\n")

input('Press ENTER to exit')
