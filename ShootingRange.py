from selenium import webdriver
from bs4 import BeautifulSoup
import requests
import ctypes, time

# browser = webdriver.Firefox()
# browser.get('http://seleniumhq.org/')

# res = requests.get('http://mysoft.healthcare.uiowa.edu/')
# # res.raise_for_status()
# testSoup = BeautifulSoup(res.text, 'html.parser')
# # print(testSoup)
# inField = testSoup.select('input')
# for i in range(0,len(inField)):
#     print(str(inField[i]))

ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)

time.sleep(2)

ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)
ctypes.windll.user32.keybd_event(0x012, 0, 2, 0)

time.sleep(2)

ctypes.windll.user32.keybd_event(0xA2, 0, 0, 0)
ctypes.windll.user32.keybd_event(0x56, 0, 0, 0)
ctypes.windll.user32.keybd_event(0xA2, 0, 2, 0)
ctypes.windll.user32.keybd_event(0x56, 0, 2, 0)
time.sleep(1)
ctypes.windll.user32.keybd_event(0x0D, 0, 0, 0)
ctypes.windll.user32.keybd_event(0x0D, 0, 2, 0)
time.sleep(2)
ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)
ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)
# ctypes.windll.user32.keybd_event(0x09, 0, 0, 0)
# ctypes.windll.user32.keybd_event(0x09, 0, 2, 0)
time.sleep(0.2)
ctypes.windll.user32.keybd_event(0x0D, 0, 0, 0)
ctypes.windll.user32.keybd_event(0x0D, 0, 2, 0)

