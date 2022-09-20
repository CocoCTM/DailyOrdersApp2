from PyPDF2 import PdfFileReader, PdfFileWriter
import tabula

desktop = 'C:/Users/corina/Downloads/'
file_name = '17166533880.pdf'

pdf_data = tabula.read_pdf('C:/Users/corina/Downloads/17166533880.pdf',area=(146.129,48.557,248.74,205.222), pages=1, stream=True)
print(pdf_data)
#area=(146.129,48.557,248.74,205.222)