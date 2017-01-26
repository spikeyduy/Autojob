#  Will copy serial nbr from excel and alt+tab to
#  chrome and paste. Enter. Tab twice, if there is a "inactive"
#  on the webpage. Else, Alert me.
import openpyxl
import os
import ctypes
import time


class AutoJob:
    def __init__(self):
        ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)  # alt
        ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)  # tab
        time.sleep(1)  # wait 2
        ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)  # !alt
        ctypes.windll.user32.keybd_event(0x012, 0, 2, 0)  # !tab

        self.get_file()

    def alt_tab(self):
        ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)  # alt
        ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)  # tab
        time.sleep(1)  # wait 2
        ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)  # !alt
        ctypes.windll.user32.keybd_event(0x012, 0, 2, 0)  # !tab

    def tab(self):
        ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)  # tab
        ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)

    def paste(self):
        ctypes.windll.user32.keybd_event(0xA2, 0, 0, 0)  # ctrl
        ctypes.windll.user32.keybd_event(0x56, 0, 0, 0)  # v
        ctypes.windll.user32.keybd_event(0xA2, 0, 2, 0)  # !ctrl
        ctypes.windll.user32.keybd_event(0x56, 0, 2, 0)  # !v

    def enter(self):
        ctypes.windll.user32.keybd_event(0x0D, 0, 0, 0)  # enter
        ctypes.windll.user32.keybd_event(0x0D, 0, 2, 0)

    def first_click(self):
        for i in range(2):
            self.tab()
        self.enter()

    def change_status(self):
        for i in range(3):
            time.sleep(0.5)
            print("Status")
            self.tab()
            time.sleep(0.2)
        ctypes.windll.user32.keybd_event(0x55, 0, 0, 0)  # u
        ctypes.windll.user32.keybd_event(0x55, 0, 2, 0)

    def change_storage(self):
        for i in range(6):
            print("Storage Location")
            self.tab()
            time.sleep(0.2)
        ctypes.windll.user32.keybd_event(0x53, 0, 0, 0)  # s
        ctypes.windll.user32.keybd_event(0x53, 0, 2, 0)
        ctypes.windll.user32.keybd_event(0x45, 0, 0, 0)  # e
        ctypes.windll.user32.keybd_event(0x45, 0, 2, 0)

    def change_date(self):
        for i in range(12):
            print("End Cust Maint")
            self.tab()
            time.sleep(0.2)
        ctypes.windll.user32.keybd_event(0x31, 0, 0, 0)  # 1
        ctypes.windll.user32.keybd_event(0x31, 0, 2, 0)
        time.sleep(0.2)
        ctypes.windll.user32.keybd_event(0x28, 0, 0, 0)  # Down arrow
        ctypes.windll.user32.keybd_event(0x28, 0, 2, 0)
        time.sleep(0.2)
        self.enter()

    def cancel_butt(self):
        for i in range(8):
            print('Cancel')
            self.tab()
            time.sleep(0.2)
        self.enter()
        time.sleep(0.5)
        self.enter()

    # alt tabs through and checks the excel
    def check_tab(self):
        ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)  # alt
        ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)  # tab
        ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)  # tab
        time.sleep(2)
        ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)  # !alt
        ctypes.windll.user32.keybd_event(0x012, 0, 2, 0)  # !tab
        time.sleep(1)
        ctypes.windll.user32.keybd_event(0xA2, 0, 0, 0)  # ctrl
        ctypes.windll.user32.keybd_event(0x22, 0, 0, 0)  # pgup
        ctypes.windll.user32.keybd_event(0xA2, 0, 2, 0)  # ctrl
        ctypes.windll.user32.keybd_event(0x22, 0, 2, 0)  # pgup
        time.sleep(0.5)
        ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)  # alt
        ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)  # tab
        time.sleep(1)  # wait 2
        ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)  # !alt
        ctypes.windll.user32.keybd_event(0x012, 0, 2, 0)  # !tab

    # function to copy a variable to the clipboard
    def add_to_clip(self, text):
        if type(text) == str:
            command = 'echo ' + text + '| clip'
            os.system(command)

    # This does all the leg work
    def human_job(self, serial):
        self.add_to_clip(serial)
        self.paste()
        time.sleep(1)
        self.enter()
        time.sleep(1)
        self.first_click()
        time.sleep(1)
        self.change_status()
        time.sleep(1)
        self.change_storage()
        time.sleep(1)
        self.change_date()
        time.sleep(1)
        self.cancel_butt()
        time.sleep(1)

    def get_file(self):
        # filePath = raw_input("Please enter the full path of the spreadsheet:")
        # print(filePath)
        # returnDate = input("Please enter the date in mm/dd/yyyy format for day"
        #                    " of trade in: ")
        # wb = openpyxl.load_workbook(filePath)
        wb = openpyxl.load_workbook('test.xlsx')
        # print(type(wb))

        # fetch the sheets of the workbook
        for x in wb.get_sheet_names():
            curr_sheet = wb.get_sheet_by_name(x)
            for i in range(1, curr_sheet.max_row+1):
                if type(curr_sheet.cell(row=i, column=1).value) == unicode:
                    reg_string = curr_sheet.cell(row=i, column=1).value.encode('ascii', 'ignore')
                    self.human_job(reg_string)
                    # print(currSheet.cell(row=i, column=1).value)
            time.sleep(1)
            self.check_tab()
AutoJob()
