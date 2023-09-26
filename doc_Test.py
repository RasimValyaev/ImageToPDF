from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt

# Открываем документ
doc = Document(r'C:\Rasim\Python\ImageToPDF\gfg.docx')

# Находим параграф с нумерованным списком
for i, paragraph in enumerate(doc.paragraphs):
    if paragraph.style.name.startswith('List'):
        # Получаем текущий номер пункта
        # current_number = int(paragraph.runs[0].text.split('.')[0])

        # Создаем новый параграф с новым пунктом
        new_paragraph = paragraph.insert_paragraph_before(f'{i + 1}. Ваш новый пункт ', style='List Bullet')
        new_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        for run in new_paragraph.runs:
            run.font.size = Pt(12)  # Установите желаемый размер шрифта

# Сохраняем документ
doc.save('ваш_документ_с_новым_пунктом.docx')
