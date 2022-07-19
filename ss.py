import PyPDF2
import re

# Open the pdf file
# object = PyPDF2.PdfFileReader(r"C:/Users/iancr/OneDrive - Continental AG/JavaScript Projects/Python_TT/QP.pdf", strict=False)
# object = PyPDF2.PdfFileReader(r"C:/Users/iancr/OneDrive - Continental AG/JavaScript Projects/Python_TT/QP1.pdf", strict = False)
object = PyPDF2.PdfFileReader(r"C:/Users/iancr/OneDrive - Continental AG/JavaScript Projects/Python_TT/QP3.pdf", strict = False)
# object = PyPDF2.PdfFileReader(r"C:/Users/iancr/OneDrive - Continental AG/JavaScript Projects/Python_TT/QP3-saveas.pdf", strict = False)

word = 'HIGH TEMPERATURE STORAGE TEST'.lower().replace("test", '').strip()
print(word)
QP_text = {}
def get_content_page():
    # Iterate trough each page from the pdf and return the text for each one
    for i in range(0, len(object.pages)):
        page = object.pages[i]
        text = page.extract_text().lower()
        QP_text[i] = text
        pages_to_get = []
    # For each page, look for the key_word given by the user and find the first page on which
    for key in QP_text:
        # print(f'{key} is {QP_text[key]}')
        if word in QP_text[key]:
            page = QP_text[key].replace(".", '').strip().splitlines()
            print(page)
            for i in range(0, len(page)):
                if word in page[i]:
                    try:
                        test_first_page_number = int(page[i].strip()[-2:])
                        # print(test_first_page_number)
                        next_test_page_number = int(page[i + 1].strip()[-2:])
                        page_number = test_first_page_number
                        pages_to_get.append(page_number)
                        for i in range(page_number, next_test_page_number - 1):
                            page_number += 1
                            pages_to_get.append(page_number)
                        return pages_to_get
                    except ValueError:
                        continue


# Based on Table of Content from the pdf, search first page of the test and then look for the page of the next test
# def find_page(content_page):
#     page = content_page.replace(".",'').strip().splitlines()
#     print(page)
#     pages_to_get = []
#     for i in range(0,len(page)):
#         if word in page[i]:
#             test_first_page_number = int(page[i].strip()[-2:])
#             print(test_first_page_number)
#             next_test_page_number = int(page[i+1].strip()[-2:])
#             page_number = test_first_page_number
#             pages_to_get.append(page_number)
#             for i in range(page_number,next_test_page_number-1):
#                 page_number +=1
#                 pages_to_get.append(page_number)
#     print(pages_to_get)

print(get_content_page())
