# pip install pdf2image
# download folder from https://github.com/oschwartz10612/poppler-windows/releases/ and paste to current dir

from pdf2image import convert_from_path

# images = convert_from_path('1.pdf', poppler_path=r'c:\Rasim\Python\ImageToPDF\poppler-23.08.0\Library\bin', \
#                            dpi=200, thread_count=4, output_folder=r'C:\Users\Rasim\Desktop\Scan\7', transparent=True, \
#                             fmt='PNG',last_page=True
#                            )
images = convert_from_path('1.pdf', poppler_path=r'c:\Rasim\Python\ImageToPDF\poppler-23.08.0\Library\bin', dpi=200)
x = 1
for image in images:
    image.save(r'C:\Users\Rasim\Desktop\Scan\7\deneme_' + str(x) + '.jpeg', 'JPEG')
    break
