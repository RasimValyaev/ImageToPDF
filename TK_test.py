#!/usr/bin/env python3
import time
import tkinter
from datetime import datetime
from queue import Empty, Queue
from threading import Thread
import requests


def display_result(label, q):
    try:
        label['text'] = q.get(block=False)
    except Empty:  # update time at 100ms boundary
        label.after(round(100 - (1000 * time.time()) % 100),
                    display_result, label, q)
        label['text'] = str(datetime.now().strftime("%H:%M:%S.%f")[:-3])


def get_result(q):
    # blocking function
    q.put('ip=' + requests.get('https://httpbin.org/delay/3').json()['origin'])


# get result in a background thread
result_queue = Queue()
Thread(target=get_result, args=[result_queue], daemon=True).start()

# display result
root = tkinter.Tk()
label = tkinter.Label(font=(None, 100))
label.pack()
display_result(label, result_queue)

# center window
# root.eval('tk::PlaceWindow %s center' % root.winfo_pathname(root.winfo_id()))
root.mainloop()
