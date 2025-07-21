import openpyxl
from openpyxl import load_workbook

def processSheet(wb):
    # parse the student ID, city and their address in dictionary
    
    sheet = wb.active
    for row in sheet.iter_rows(values_only = True):
        print(row)

    
def main():
    studentWb = load_workbook("C:\\Users\steve\SAIT\Sem2\\Capstone\\Inputs\Students\\ModifiedClassList.xlsx")
    hostListWb = load_workbook("C:\\Users\steve\SAIT\Sem2\\Capstone\\Inputs\Locations\\ModifiedHostList.xlsx")

    # parse the origins (student address) 
    print("Student Location \n")
    processSheet(studentWb)
    print("Host Location \n")

    processSheet(hostListWb)




main()