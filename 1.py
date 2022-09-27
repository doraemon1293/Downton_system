from PyPDF2 import PdfReader

reader = PdfReader("SO：NBYE155339.pdf")
page = reader.pages[0]
print(page.extract_text())

print(reader.metadata)