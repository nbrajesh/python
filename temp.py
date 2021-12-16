from pathlib import Path # Standard Python Library
import xlwings as xw # pip install xlwings
from PyPDF2 import PdfFileMerger, PdfFileReader # pip install PyPDF2
def main():
wb = xw.Book.caller()
sheet = wb.sheets[0]
merger = PdfFileMerger()
sheet.range(&quot;status&quot;).clear_contents()
source_dir = sheet.range(&quot;source_dir&quot;).value
output_name = sheet.range(&quot;output_name&quot;).value + &quot;.pdf&quot;
pdf_files = list(Path(source_dir).glob(&quot;*.pdf&quot;))
for pdf_file in pdf_files:
merger.append(PdfFileReader(str(pdf_file), &quot;rb&quot;))

output_path = str(Path(__file__).parent / output_name)
merger.write(output_path)
sheet.range(&quot;status&quot;).value = f&quot;The file have been saved here: {output_path}&quot;
if __name__ == &quot;__main__&quot;:
xw.Book(&quot;pdfmerger.xlsm&quot;).set_mock_caller()
main()
