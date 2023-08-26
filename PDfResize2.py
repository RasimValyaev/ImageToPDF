import os
import pypdftk  # pdftk main prog has to be installed and added to path too...
import pdf2image
from PIL import Image


def crunchPdfImages(file_to_crunch):
    working_folder = os.path.dirname(file_to_crunch)
    working_dir = os.path.join(working_folder, "temp_working_dir")
    if not (os.path.exists(working_dir)): os.mkdir(working_dir)

    # Get all the image...
    extractPDFImages(file_to_crunch, working_dir)
    # Compress all the images... (no resize, just optimise)
    all_image_list = [entry.path for entry in os.scandir(working_dir) if isImage(entry.path)]
    if (len(all_image_list) > 0):
        for an_image in all_image_list:
            img_picture = Image.open(an_image).convert("RGB")
            img_picture.save(an_image, "JPEG", optimize=True)
    else:
        print("No images found in PDF...")

    # Uncompress the PDF
    pdf_folder = os.path.join(working_dir, "pdf_uncompressed")
    if not (os.path.exists(pdf_folder)): os.mkdir(pdf_folder)
    pdf_datain_file = os.path.join(pdf_folder, "uncompressed_pdf.pdf")
    pdf_dataout_file = os.path.join(pdf_folder, "new_images_pdf.pdf")
    print("Uncompressing PDF...")
    pypdftk.uncompress('"' + file_to_crunch + '"', '"' + pdf_datain_file + '"')

    # Now get to work...
    #   The PDF is comprised of objects, some of which are lablled as images.
    #   Each image has the line "/Subtype /Image" before the "stream" which is then ended by "endstream" then "endobj".
    #   In between the stream and endstream is the encoded image data... hopefully I can replace this in the same order that
    #   the images were taken out.
    picture_replace_count = 0
    pdf_openfile_in = open(pdf_datain_file, "rb")
    pdf_openfile_out = open(pdf_dataout_file, "wb")
    pdf_file_lines = pdf_openfile_in.readlines()

    looking_for_next_stream = False
    found_stream_and_removing = False
    skip_a_line = False

    for line in pdf_file_lines:
        new_line_addition = ""  # For adding to byte count, resetting to null here just in case
        current_line_val = line.decode("Latin-1").strip()

        if (looking_for_next_stream):
            # Last image tag has been found but not dealt with, so find the stream then
            if (current_line_val[:8] == "/Length "):
                # Update the length
                skip_a_line = True
                new_img_size = str(os.path.getsize(all_image_list[picture_replace_count]))
                new_line = r"/Length " + new_img_size + "\n"
                pdf_openfile_out.write(new_line.encode("latin-1"))  # add new line
            if (current_line_val == "stream"):
                print("Stream start found... skipping stream information")
                looking_for_next_stream = False  # it's been found
                found_stream_and_removing = True  # time to delete

                new_line_addition = "stream\n".encode("latin-1")
                pdf_openfile_out.write(new_line_addition)  # add the line in or it will be skipped

        elif (found_stream_and_removing):
            if (current_line_val == "endstream"):
                print("Stream end found")
                found_stream_and_removing = False  # Passed through all image data line
                # Now, add in the new image data and continue on.
                print("Adding new image data...")

                image = open(all_image_list[picture_replace_count], 'rb')
                pdf_openfile_out.write(image.read())
                image.close()

                picture_replace_count += 1
                pdf_openfile_out.write("\n".encode("latin-1"))  # add new line

        elif (current_line_val == r"/Subtype /Image"):
            print("Found an image place, number " + str(picture_replace_count))
            print("Looking for stream start...")
            looking_for_next_stream = True
            # Find next

        if not (found_stream_and_removing) and not (skip_a_line):
            pdf_openfile_out.write(line)

        skip_a_line = False

    pdf_openfile_in.close()
    pdf_openfile_out.close()

    print("Rebuilding xref table (post newfile creation)")
    rebuildXrefTable(pdf_dataout_file)


def rebuildXrefTable(pdf_file_in, pdf_file_out=None):
    # Updating the xref table:
    #   * Assumes uncompressed PDF file
    #   To do this I need the number of bytes that precede and object (this is used as a reference).
    #   So, each line I will need to count the byte number and tally up
    #   When an object is found, the byte_count will be added to the reference list and then used to create the xref table
    #   Also need to update the "startxref" at the bottom (similar principle).

    if (pdf_file_out == None): pdf_file_out = os.path.join(os.path.dirname(pdf_file_in), "rebuilt_xref_pdf.pdf")
    print("Updating xref table of: " + os.path.basename(pdf_file_in))

    byte_count = 0
    xref_start = 0
    object_location_reference = []
    updating_xref_stage = 1
    pdf_openfile_in = open(pdf_file_in, "rb")
    pdf_openfile_out = open(pdf_file_out, "wb")
    pdf_file_lines = pdf_openfile_in.readlines()

    for line in pdf_file_lines:
        current_line_val = line.decode("Latin-1").strip()
        if (" obj" in current_line_val):
            # Check if the place is an object loc, store byte reference and object index
            obj_ref_index = current_line_val.split(" ")[0]
            print("Found new object (index, location): (" + str(obj_ref_index) + ", " + str(byte_count) + ")")
            object_location_reference.append((int(obj_ref_index), byte_count))
        elif ("startxref" in current_line_val):
            # This is the last thing to edit (right at the bottom, update the xref start location and then add the file end.
            print("Updating the xref start value with new data...")
            new_line = "startxref\n" + str(xref_start) + "\n" + r"%%EOF"
            pdf_openfile_out.write(new_line.encode("latin-1"))
            break
        elif ("xref" in current_line_val):
            print("Recording the new xref byte location")
            preceeding_str = current_line_val.split("xref")[0]
            preceeding_count = len(preceeding_str.encode("latin-1"))
            xref_start = byte_count + preceeding_count  # used at the end
            updating_xref_stage = 2

        elif (updating_xref_stage == 2 or updating_xref_stage == 3):
            # This stage simply skips the first 2 xref data lines (and prints it o the new file as is)
            updating_xref_stage += 1
        elif (updating_xref_stage == 4):
            print("Creating new xref object byte location table...")
            object_location_reference.sort()  # Sort the collected xref locations by their object index.
            # Now add the new xref data information
            for xref_loc in object_location_reference:
                new_val = str(xref_loc[1]).zfill(10)  # Pad the number out
                new_val = new_val + " 00000 n \n"
                pdf_openfile_out.write(new_val.encode("latin-1"))
            updating_xref_stage = 5
        elif (updating_xref_stage == 5):
            # Stage 5 doesn't record the read in lines into new file, step 6 will.
            if ("trailer" in current_line_val): updating_xref_stage = 6

        # Write to file
        if not (updating_xref_stage == 5):
            pdf_openfile_out.write(line)
            byte_count += len(line)

    pdf_openfile_in.close()
    pdf_openfile_out.close()


# To use the PDF compression:
crunchPdfImages(r"C:\Users\Person\Desktop\Test Folder\Pdf File.pdf")