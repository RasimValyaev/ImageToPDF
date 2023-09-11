# https://stackoverflow.com/questions/63197559/how-to-convert-a-word-docx-to-an-image-using-python

# pip install PyAutoGUI

import win32com.client as win32
import pyautogui
import win32gui
import time

docfile = rfilename = r"C:\Rasim\Python\Prestige\TelegramBot\001694011666.docx"
shotfile = r'C:\Rasim\Python\Prestige\TelegramBot\111.png'


def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))


word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
word.WindowState = 1  # maximize

top_windows = []
win32gui.EnumWindows(windowEnumerationHandler, top_windows)

for i in top_windows:  # all open apps
    if "word" in i[1].lower():  # find word (assume only one)
        try:
            win32gui.ShowWindow(i[0], 5)
            win32gui.SetForegroundWindow(i[0])  # bring to front
            break
        except:
            pass

doc = word.Documents.Add(docfile)  # open file

time.sleep(2)  # wait for doc to load

myScreenshot = pyautogui.screenshot()  # take screenshot
myScreenshot.save(shotfile)  # save screenshot

# close doc and word app
doc.Close()
word.Application.Quit()