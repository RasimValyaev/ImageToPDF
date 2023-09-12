# file_path = filedialog.askopenfilename(filetypes=[("Excel files", ["*.xls","*.xlsx"])])

from tkinter import *
from tkinter import ttk
from tkinter import filedialog

from Main import merge_files_to_one

root = Tk()
root.title("Prestige")
root.geometry("250x200")

root.grid_rowconfigure(index=0, weight=1)
root.grid_columnconfigure(index=0, weight=1)
# root.grid_columnconfigure(index=1, weight=1)

text_editor = Text()
text_editor.grid(column=0, columnspan=1, row=0)


# открываем файл в текстовое поле
def open_file():
    excel_file_source = filedialog.askopenfilename(filetypes=[("Excel files", ["*.xls", "*.xlsx"])])
    if excel_file_source != '':
        merge_files_to_one(excel_file_source)


open_button = ttk.Button(text="Открыть файл", command=open_file)
open_button.grid(column=0, row=1, sticky=NSEW, padx=10)

root.mainloop()
