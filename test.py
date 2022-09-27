from pdfminer.high_level import extract_pages

for page_layout in extract_pages("SO：NBYE155339.pdf"):
    for element in page_layout:
        print(element)