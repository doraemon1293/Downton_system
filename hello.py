import pdfplumber

with pdfplumber.open("E7120F0F371103A2E0530A008761F9EB.pdf") as pdf:
   first_page = pdf.pages[0]
   print(first_page.extract_text())