import tkinter as tk
from tkinter import filedialog

from MergeExcleWord import merge_excle_word_main


def select_file():
    excel_file_source = filedialog.askopenfilename(filetypes=[("Excel files", ["*.xls", "*.xlsx"])])
    if excel_file_source != '':
        print("Вы выбрали файл", excel_file_source)
        merge_excle_word_main(excel_file_source)

root = tk.Tk()
root.title("PrestigeProduct")
root.geometry("400x100")
root.grid_rowconfigure(index=0, weight=1)
root.grid_columnconfigure(index=0, weight=1)

label = tk.Label(root, text="Выбираемый Excel файл должен быть в папке со сканами.\n"
                            "Наименование файлов должно начинаться на ВН или ТТН.\n"                            
                            "В Worde не забудьте обновить список платежей\n"
                 )
label.pack()

button = tk.Button(root, text="Выберите Excel файл", command=select_file)
button.pack()

root.mainloop()
