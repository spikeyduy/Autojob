# Will copy serial nbr from excel and alt+tab to
# chrome and paste. Enter. Tab twice, if there is a "inactive"
# on the webpage. Else, Alert me.
# Move through with ? tabs and down arrow twice.
# Move through with ? tabs and up/down arrow depending on what is selected.
# Tab ? times and input the custom date. Tab to save and complete, loop.
import openpyxl
import os
import urllib

class autoJob():
    def __init__(self):
        self.getfile()

    # function to copy a variable to the clipboard
    def addToClip(text):
        if(type(text) == str):
            command = 'echo ' + text + '| clip'
            os.system(command)

    # This does all the leg work
    def humanJob(serial):
        # print(serial)

        autoJob.addToClip(serial)
        # print(type(serial), serial)

    def getfile(self):
        filePath = input("Please enter the full path of the spreadsheet: ")
        # print(filePath)
        # returnDate = input("Please enter the date in mm/dd/yyyy format for day of trade in: ")

        wb = openpyxl.load_workbook(filePath)
        # print(type(wb))

        # fetch all the sheets in the workbook
        for x in wb.get_sheet_names():
            currSheet = wb.get_sheet_by_name(x)
            # print(wb.get_sheet_names()[x])
            for i in range(1, currSheet.max_row+1):
                # print(currSheet.cell(row=i, column=1).value, currSheet)
                autoJob.humanJob(currSheet.cell(row=i, column=1).value)

autoJob()
