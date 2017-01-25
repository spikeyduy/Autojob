# Will copy serial nbr from excel and alt+tab to
# chrome and paste. Enter. Tab twice, if there is a "inactive"
# on the webpage. Else, Alert me.
# Move through with ? tabs and down arrow twice.
# Move through with ? tabs and up/down arrow depending on what is selected.
# Tab ? times and input the custom date. Tab to save and complete, loop.
import openpyxl
import os
from bs4 import BeautifulSoup
import urllib
import requests
import ctypes, time

class autoJob():
    def __init__(self):
        ctypes.windll.user32.keybd_event(0x12, 0, 0, 0) # alt
        ctypes.windll.user32.keybd_event(0x09, 0, 0, 0) # tab

        time.sleep(2) # wait 2

        ctypes.windll.user32.keybd_event(0x09, 0, 2, 0) # !alt
        ctypes.windll.user32.keybd_event(0x012, 0, 2, 0) # !tab

        self.getfile()

    # function to copy a variable to the clipboard
    def addToClip(self, text):
        if(type(text) == str):
            command = 'echo ' + text + '| clip'
            os.system(command)

    # This does all the leg work
    def humanJob(self, serial):
        # print(serial)
        # MAY NOT NEED THIS
        # could do with just letting it sit in the variable form
        # and just input it into the text box
        self.addToClip(serial)
        ctypes.windll.user32.keybd_event(0xA2, 0, 0, 0) # ctrl
        ctypes.windll.user32.keybd_event(0x56, 0, 0, 0) # v
        ctypes.windll.user32.keybd_event(0xA2, 0, 2, 0) # !ctrl
        ctypes.windll.user32.keybd_event(0x56, 0, 2, 0) # !v
        ctypes.windll.user32.keybd_event(0xA2, 0, 2, 0)
        ctypes.windll.user32.keybd_event(0x56, 0, 2, 0)
        time.sleep(1)
        ctypes.windll.user32.keybd_event(0x0D, 0, 0, 0) # enter
        ctypes.windll.user32.keybd_event(0x0D, 0, 2, 0)
        time.sleep(1)
        ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)  # tab
        ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)
        ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)  # tab
        ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)
        ctypes.windll.user32.keybd_event(0x0D, 0, 0, 0)  # enter
        ctypes.windll.user32.keybd_event(0x0D, 0, 2, 0)
        time.sleep(1)
        for i in range(29):
            print('tabbing through')
            ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)  # tab
            ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)
            time.sleep(0.2)
        ctypes.windll.user32.keybd_event(0x0D, 0, 0, 0)  # enter
        ctypes.windll.user32.keybd_event(0x0D, 0, 2, 0)
        time.sleep(3)
        # print(type(serial), serial)
        # alt tab over

    def getfile(self):
        # filePath = raw_input("Please enter the full path of the spreadsheet:")
        # print(filePath)
        # returnDate = input("Please enter the date in mm/dd/yyyy format for day"
        #                    " of trade in: ")
        # wb = openpyxl.load_workbook(filePath)
        wb = openpyxl.load_workbook('test.xlsx')
        # print(type(wb))

        # fetch the sheets of the workbook
        for x in wb.get_sheet_names():
            currSheet = wb.get_sheet_by_name(x)
            for i in range(1, currSheet.max_row+1):
                if type(currSheet.cell(row=i, column=1).value) == unicode:
                    regString = currSheet.cell(row=i, column=1).value.encode('ascii', 'ignore')
                    self.humanJob(regString)
                    # print(currSheet.cell(row=i, column=1).value)

autoJob()