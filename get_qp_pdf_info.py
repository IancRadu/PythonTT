import pathlib
from tkinter import filedialog

import PyPDF2
import fitz

# Function to get the path to the QP
def get_qp_path(current_data):
    path_to_qp_name = pathlib.Path(current_data["PathtoTO"]).parent.parent
    qp_folder = str(path_to_qp_name.name)[:9]

    path_to_qp = pathlib.Path(current_data["PathtoTO"]).parent.parent.parent
    full_path_to_qp = pathlib.Path(f'{path_to_qp}/01_SPECIFICATION/01_QP/{qp_folder}')
    try:
        for child in full_path_to_qp.iterdir():
            if 'pdf' in str(child):
                current_data['PathToQP'] = str(child)
                return child
        # print(path_to_qp)
    except FileNotFoundError:
        try:
            for child in full_path_to_qp.parent.iterdir():
                if qp_folder[3:] in str(child):
                    full_path = pathlib.Path(f'{child}')
                    # print("Got trough here --------------------1")
                    # print(full_path)
                    for child2 in full_path.iterdir():
                        # print("Got trough here --------------------2")
                        if 'pdf' in str(child2):
                            # print("Got trough here --------------------3")
                            current_data['PathToQP'] = str(child2)
                            return child2
        except FileNotFoundError:
            print(f"QP PDF file was not found at {full_path_to_qp}")
            print("Select correct path to the PDF file")
            path = filedialog.askopenfile().name
            return path


def get_page_number(current_data):
    project_data = current_data
    # Open the pdf file
    pdf_source = PyPDF2.PdfFileReader(get_qp_path(current_data), strict=False)
    QP_text = {}
    QP_text_clean = {}
    # Iterate trough each page from the pdf and return the text for each one
    for i in range(0, len(pdf_source.pages)):
        page = pdf_source.pages[i]
        text = page.extract_text()
        QP_text[i] = text.lower().replace(' ', '')
        QP_text_clean[i] = text
    for key in project_data['TestFlow']:
        pages_to_get = []
        word_clean = project_data['TestFlow'][key]['Test name']
        word = project_data['TestFlow'][key]['Test name'].lower().replace("test", '').replace(" ", '').strip()
        data = project_data['TestFlow'][key]
        # print(f'Search pdf for {word}')
        # Get raw string of standards used in test
        for key in QP_text:
            if word in QP_text[key].lower():
                if 'standard' in QP_text[key].lower():
                    if 'parameters' in QP_text[key].lower():
                        # print(QP_text_clean[key])
                        if '........' not in QP_text[key].lower(): # To bypass the table of contents
                            # print("Got trough here --------------------1")
                            if 'track' not in QP_text[key].lower(): # To bypass the planning(flow)
                                # print("Got trough here --------------------2")
                                # print(QP_text_clean[key])
                                try:
                                    data['QP_read_standards_page'] = \
                                    QP_text_clean[key].split('Standard')[1].split('Test')[
                                        0].split('Deviatio')
                                    # print(data['QP_read_standards_page'])
                                    break
                                except IndexError:
                                    data[
                                        'QP_read_standards_page'] = f'Failed to read standards.'
                                    break
                                    # print(data['QP_read_standards_page'])
            else:
                data[
                    'QP_read_standards_page'] = f'Failed to read standards. Check if {word_clean} is the same in TO ' \
                                                f'as it is in QP '

                # print(QP_text_clean[key])
        # For each page, look for the key_word given by the user and find the first occurrence, which usually is the
        # table of contents
        for key in QP_text:
            if word in QP_text[key]:
                page = QP_text[key].replace(".", '').strip().splitlines()
                # print(page)
                for i in range(0, len(page)):
                    if word in page[i]:
                        # Checks if the first occurrence is the table of content, if not go to the next occurrence of
                        # word
                        try:
                            test_first_page_number = int(page[i].strip()[-2:])
                            # print(test_first_page_number)
                            next_test_page_number = int(page[i + 1].strip()[-2:])
                            page_number = test_first_page_number
                            pages_to_get.append(page_number)
                            for i in range(page_number, next_test_page_number - 1):
                                page_number += 1
                                pages_to_get.append(page_number)
                            data['QP_read_pages'] = pages_to_get
                            break
                        except ValueError:
                            try:
                                if data['QP_read_pages'] is None:
                                    print("Trow KeyError")
                            except KeyError:
                                print(
                                    f"Failed to read page number: Check if {word_clean} is the same in TO as it is in QP ")
                                pages_to_get.append(1)
                                data['QP_read_pages'] = pages_to_get
                            continue


# Retrieve image of page from pdf file using pymupdf module, and export it to project file
def output_pdf_page(current_data):

    pdffile = f"{get_qp_path(current_data)}"
    zoom = 2  # to increase the resolution
    mat = fitz.Matrix(zoom, zoom)
    doc = fitz.open(pdffile)
    for i in range(1,len(current_data["TestFlow"])+1):
        for m in range(0, len(current_data["TestFlow"][i]['QP_read_pages'])):
            page = doc.load_page(current_data["TestFlow"][i]['QP_read_pages'][m]-1)  # number of page
            pix = page.get_pixmap(matrix=mat)
            output = f"{current_data['TestFlow'][i]['Pathto04_Snipping']}/{current_data['TestFlow'][i]['QP_read_pages'][m]}.png"
            pix.save(output)

