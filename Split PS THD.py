import pandas as pd
from PyPDF2 import PdfFileReader, PdfFileWriter
import tabula
import PyPDF2
import pandas as pd
from datetime import date

desktop = 'C:/Users/corina/Downloads/'
file_name = '17363975311.pdf'
ord_data = pd.read_excel("Y:/User/Order Entry User/CTM Household Appliances Orders Processed/The Home Depot Inc/FILES/Daily Order/THD_Order_DTLS_2022-11-04.xlsx", sheet_name='Sheet1')
ord_data_sub = ord_data[['PO Number', 'Location']].drop_duplicates()
calc_ship_from = ord_data.groupby('PO Number')['Location'].nunique().to_frame().reset_index()
thd_ord_ship = calc_ship_from.merge(ord_data_sub, on='PO Number', how='left')
thd_ord_ship.loc[thd_ord_ship["Location_x"] != 1, "Location_y"] = "Multi-Location - Edit PS Manually"
thd_ord_ship = thd_ord_ship.drop_duplicates()
thd_ord_ship = thd_ord_ship[['PO Number', 'Location_y']]
thd_ord_ship['PO Number'] = thd_ord_ship['PO Number'].astype(str)
thd_ord_ship['PO Number'] = thd_ord_ship['PO Number'].str.strip()  # PO# # IF NOT 8 DIGITS ADD 0 IN FRONT
thd_ord_ship['PO Number'] = thd_ord_ship['PO Number'].str.zfill(8)  # PO# # IF NOT 8 DIGITS ADD 0 IN FRONT
thd_ord_ship = thd_ord_ship.rename(columns={"Location_y": "Location", "PO Number": "PO#"})  # RENAME ORDER DATE COLUMN
thd_ord_ship['PO#'].astype(str)
#pdf_data = tabula.read_pdf('C:/Users/corina/Downloads/17239562369.pdf',area=(146.129,48.557,248.74,205.222), pages=1, stream=True)

# LOWES SOLD TO area=(146.129,48.557,248.74,205.222)
# LOWES PO FIELD pdf_data = tabula.read_pdf('C:/Users/corina/Downloads/17239562369.pdf',area=(170,500,190,556), pages=1, stream=True)[0]
#pdf_data.iloc[0,0]
#tabula.read_pdf('C:/Users/corina/Downloads/17239562369.pdf',area=(170,500,190,556), pages=1, stream=True)[0].iloc[0,0]
file = open((desktop+file_name), 'rb')
readpdf = PyPDF2.PdfFileReader(file)
totalpages = readpdf.numPages

pdf_po = pd.DataFrame(columns=['PO#', 'P1'], index=range(totalpages))

#pdf_po.at[0,'PageNumber'] = tabula.read_pdf('C:/Users/corina/Downloads/17252079327.pdf', area=(170, 500, 190, 556), pages=(1),stream=True)[0].iloc[0, 0]
for i in range(totalpages):
    pdf_po.at[i, 'PO#'] = tabula.read_pdf((desktop + file_name), area=(90, 410, 100, 500), pages=(i+1), stream=True)[0].columns.tolist()[0]
    pdf_po.at[i, 'P1'] = (i+1)
pdf_po = pdf_po.sort_values(by=['P1'])
pdf_po = pdf_po.reset_index()
del pdf_po['index']

pdf_po['PO#'] = pdf_po['PO#'].astype(str)
pdf_po = pdf_po.merge(thd_ord_ship, on="PO#", how='left')

for i in range(len(pdf_po)):
    pdfW = PdfFileWriter()
    pdfW.addPage(readpdf.getPage((pdf_po.iloc[i, 1] - 1)))
    output_filename = 'C:/Users/corina/Downloads/THD PO#' + pdf_po.iloc[i, 0] + " PS " + pdf_po.iloc[i, 2]+'.pdf'
    with open(output_filename, 'wb') as output:
        pdfW.write(output)

pdfW = PdfFileWriter()


