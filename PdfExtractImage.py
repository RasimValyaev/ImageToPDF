# pip3 install PyMuPDF
# pip3 install Pillow
# https://github.com/x4nth055/pdf-tools-python
# https://github.com/x4nth055/pythoncode-tutorials/blob/master/web-scraping/pdf-image-extractor/pdf_image_extractor.py

# извлекает изображения из pdf

import os
import fitz  # PyMuPDF
import io
from PIL import Image
from pathlib import Path

Image.MAX_IMAGE_PIXELS = None


def extract_image(file, output_dir, output_format="jpeg"):
    # Minimum width and height for extracted images
    min_width = 100
    min_height = 100
    # Create the output directory if it does not exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # open the file
    pdf_file = fitz.open(file)
    # Iterate over PDF pages

    for page_index in range(len(pdf_file)):
        # Get the page itself
        page = pdf_file[page_index]
        # Get image list
        image_list = page.get_images(full=True)
        # Print the number of images found on this page
        if image_list:
            print(f"[+] Found a total of {len(image_list)} images in page {page_index}")
        else:
            print(f"[!] No images found on page {page_index}")

        # Iterate over the images on the page
        for image_index, img in enumerate(image_list, start=1):
            # Get the XREF of the image
            xref = img[0]
            # Extract the image bytes
            base_image = pdf_file.extract_image(xref)
            image_bytes = base_image["image"]
            # Get the image extension
            image_ext = base_image["ext"]
            # Load it to PIL
            image = Image.open(io.BytesIO(image_bytes))
            # Check if the image meets the minimum dimensions and save it
            if image.width >= min_width and image.height >= min_height:
                path = Path(file)
                print(str(path.absolute()))
                print(path.name)
                print(path.absolute().as_uri())
                print(path.stem)
                image.save(
                    open(os.path.join(output_dir, f"{path.stem}_{page_index + 1}_{image_index}.{output_format}"), "wb"),
                    format=output_format.upper())
            else:
                print(f"[-] Skipping image {image_index} on page {page_index} due to its small size.")


if __name__ == '__main__':
    file = r"\\PRESTIGEPRODUCT\Scan\ЕСП - Copy\Resize\ТТН 13731 10.05.2023.pdf"
    # Output directory for the extracted images
    output_dir = r"\\PRESTIGEPRODUCT\Scan\ЕСП - Copy\copy"
    # Desired output image format
    output_format = "jpeg"
    extract_image(file, output_dir, output_format)
