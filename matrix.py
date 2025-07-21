import openpyxl
from openpyxl import load_workbook

def processSheet(wb, dict):
    # parse the student ID, city and their address in dictionary
    
    sheet = wb.active
    

    for row_num, row in enumerate(sheet.iter_rows(values_only = True), start = 1):
        if row_num == 1:
            pass

    
def main():
    studentWb = load_workbook("C:\\Users\steve\SAIT\Sem2\\Capstone\\Inputs\Students\\ModifiedClassList.xlsx")
    hostListWb = load_workbook("C:\\Users\steve\SAIT\Sem2\\Capstone\\Inputs\Locations\\ModifiedHostList.xlsx")
    studentDict = {}
    hostDict = {}
    # parse the origins (student address) 
    print("Student Location \n")
    processSheet(studentWb, studentDict)
    print("Host Location \n")
    processSheet(hostListWb, hostDict)




main()