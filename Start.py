# pip install pyinstaller
# pyinstaller --onefile start.py

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from MergeExcleWord import merge_excle_word_main


def select_file():
    excel_file_source = filedialog.askopenfilename(filetypes=[("Excel files", ["*.xls", "*.xlsx"])])
    if excel_file_source != '':
        print("Вы выбрали файл", excel_file_source)
        merge_excle_word_main(excel_file_source)
        messagebox.showinfo("PrestigeProduct", "Завершено!")
        sys.exit(0)


if __name__ == '__main__':

    root = tk.Tk()
    root.title("PrestigeProduct")
    root.geometry("600x200")
    root.grid_rowconfigure(index=0, weight=1)
    root.grid_columnconfigure(index=0, weight=1)

    label = tk.Label(root, text="Выбираемый Excel файл должен быть в папке со сканами."
                                "\n\nНаименование файлов должно начинаться на ВН и/или ТТН и иметь расширение pdf."
                                "\n(иначе в письме будут пустоты)"
                                "\n\nБанковские выписки: БВ дата.pdf"
                                "\n(иначе письма НЕ формируются)\n"
                     )
    label.pack()

    template = r"\\PRESTIGEPRODUCT\Scan\Maket.docx"

    msg_no_template = f"Не найден файл с макетом {template}!\nБудет сформирован Excel с доп. колонками\nПродолжить?"
    msg_header = "Отсутствует шаблон письма"

    if not os.path.exists(template):
        result = messagebox.askquestion(msg_header, msg_no_template)
        if result != 'yes':
            sys.exit(0)

    print("Нашел Шаблон письма")
    button = tk.Button(root, text="Выберите Excel файл", command=select_file)
    button.pack()

    root.mainloop()
