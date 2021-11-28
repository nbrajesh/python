from PyPDF2 import PdfFileMerger

pdfs = ['file1.pdf', 'file2.pdf', 'file3.pdf']

merger = PdfFileMerger()

for pdf in pdfs:
    merger.append(pdf)

merger.write("result.pdf")
print("Successfully Merged All PDF files")
merger.close()