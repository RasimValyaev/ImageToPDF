import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from MergeExcleWord import merge_excle_word_main


def select_file():
    excel_file_source = filedialog.askopenfilename(filetypes=[("Excel files", ["*.xls", "*.xlsx"])])
    if excel_file_source != '':
        print("Вы выбрали файл", excel_file_source)
        merge_excle_word_main(excel_file_source)
        messagebox.showinfo("PrestigeProduct", "Завершено!")
        sys.exit(0)

root = tk.Tk()
root.title("PrestigeProduct")
root.geometry("400x100")
root.grid_rowconfigure(index=0, weight=1)
root.grid_columnconfigure(index=0, weight=1)

label = tk.Label(root, text="Выбираемый Excel файл должен быть в папке со сканам.\n"
                            "Наименование файлов должно начинаться на ВН или ТТН и иметь расширение pdf.\n"
                            "Банковские выписки: БВ дата.pdf\n"
                 )
label.pack()

template = r"\\PRESTIGEPRODUCT\Scan\Maket.docx"

msg_no_template = f"Не найден файл с макетом {template}!\nБудет сформирован Excel с доп. колонками\nПродолжить?"
msg_header = "Отсутствует шаблон письма"

if not os.path.exists(template):
    result = messagebox.askquestion(msg_header, msg_no_template)
    if result != 'yes':
        sys.exit(0)
#
# result = messagebox.askquestion("Инфо банка", "Если данные о платежах НЕ обновлены, тогда выйти?")
#
# if result == 'yes':
#     sys.exit(0)

print("Нашел Шаблон письма")
button = tk.Button(root, text="Выберите Excel файл", command=select_file)
button.pack()

root.mainloop()
