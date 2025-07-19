import openpyxl

def processStudents(wb):
    # parse the student ID, city and their address in dictionary
    ws = wb.active
    studentDict = {}
    # testing
    print(ws)

def main():
    studentWb = openpyxl.Workbook(r"/Users/stevesu/Downloads/Classlist.xlsx")
    hostListWb = openpyxl.Workbook(r"/Users/stevesu/Downloads/ECEDHostList.xlsx")
    # parse the origins (student address) 
    processStudents(studentWb)


main()