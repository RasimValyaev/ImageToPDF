import win32com.client as win32


def word_2_pdf(in_file_word, out_file_pdf):
    word = win32.DispatchEx("Word.Application")
    doc = word.Documents.Open(in_file_word)
    doc.SaveAs(out_file_pdf, FileFormat=17)
    doc.Close()
    word.Quit()


if __name__ == '__main__':
    in_file_word = r'c:\Users\Rasim\Desktop\Scan\9.docx'
    out_file_pdf = r'c:\Users\Rasim\Desktop\Scan\9_13.pdf'
    word_2_pdf(in_file_word, out_file_pdf)
