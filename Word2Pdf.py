# pip install pywin32
import win32com.client as win32


def word_2_pdf(in_file_word, out_file_pdf):
    try:
        wdFormatPDF = 17
        word = win32.DispatchEx("Word.Application")
        doc = word.Documents.Open(in_file_word)
        doc.SaveAs(out_file_pdf, FileFormat=wdFormatPDF) # or FileFormat=17
        doc.Close()
        word.Quit()

    except Exception as e:
        print(str(e))
        pass


if __name__ == '__main__':
    in_word_file = rfilename = r"C:\Rasim\Python\Prestige\TelegramBot\001694011666.docx"
    out_pdf_file = r'c:\Users\Rasim\Desktop\Scan\9_13.pdf'
    word_2_pdf(in_word_file, out_pdf_file)
