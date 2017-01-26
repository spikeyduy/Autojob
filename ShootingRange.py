import ctypes, time

# browser = webdriver.Firefox()
# browser.get('http://seleniumhq.org/')
ctypes.windll.user32.keybd_event(0x12, 0, 0, 0) # alt
ctypes.windll.user32.keybd_event(0x09, 0, 0, 0) # tab
time.sleep(1) # wait 2
ctypes.windll.user32.keybd_event(0x09, 0, 2, 0) # !alt
ctypes.windll.user32.keybd_event(0x012, 0, 2, 0) # !ta
time.sleep(1)
ctypes.windll.user32.keybd_event(0x12, 0, 0, 0) # alt
ctypes.windll.user32.keybd_event(0x09, 0, 0, 0) # tab
ctypes.windll.user32.keybd_event(0x09, 0, 0, 0) # tab
time.sleep(2)
ctypes.windll.user32.keybd_event(0x09, 0, 2, 0) # !alt
ctypes.windll.user32.keybd_event(0x012, 0, 2, 0) # !tab
time.sleep(1)
ctypes.windll.user32.keybd_event(0xA2, 0, 0, 0) # ctrl
ctypes.windll.user32.keybd_event(0x22, 0, 0, 0) # pgup
ctypes.windll.user32.keybd_event(0xA2, 0, 2, 0) # ctrl
ctypes.windll.user32.keybd_event(0x22, 0, 2, 0) # pgup
time.sleep(0.5)
ctypes.windll.user32.keybd_event(0x12, 0, 0, 0) # alt
ctypes.windll.user32.keybd_event(0x09, 0, 0, 0) # tab
time.sleep(1) # wait 2
ctypes.windll.user32.keybd_event(0x09, 0, 2, 0) # !alt
ctypes.windll.user32.keybd_event(0x012, 0, 2, 0) # !tab