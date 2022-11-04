import pandas as pd
from PyPDF2 import PdfFileReader, PdfFileWriter
import tabula
import PyPDF2
import pandas as pd
from datetime import date

desktop = 'C:/Users/corina/Downloads/'
file_name = '17363991188.pdf'
ord_data = pd.read_excel("Y:/User/Order Entry User/CTM Household Appliances Orders Processed/Lowe's/FILES/Daily Order/Lowe's_Order_DTLS_2022-11-04.xlsx", sheet_name='Sheet1')
ord_data_sub = ord_data[['PO Number', 'Location']].drop_duplicates()
calc_ship_from = ord_data.groupby('PO Number')['Location'].nunique().to_frame().reset_index()
lowes_ord_ship = calc_ship_from.merge(ord_data_sub, on='PO Number', how='left')
lowes_ord_ship.loc[lowes_ord_ship["Location_x"] != 1, "Location_y"] = "Multi-Location - Edit PS Manually"
lowes_ord_ship = lowes_ord_ship.drop_duplicates()
lowes_ord_ship = lowes_ord_ship[['PO Number', 'Location_y']]
lowes_ord_ship = lowes_ord_ship.rename(columns={"Location_y": "Location", "PO Number": "PO#"})  # RENAME ORDER DATE COLUMN
#pdf_data = tabula.read_pdf('C:/Users/corina/Downloads/17239562369.pdf',area=(146.129,48.557,248.74,205.222), pages=1, stream=True)

# LOWES SOLD TO area=(146.129,48.557,248.74,205.222)
# LOWES PO FIELD pdf_data = tabula.read_pdf('C:/Users/corina/Downloads/17239562369.pdf',area=(170,500,190,556), pages=1, stream=True)[0]
#pdf_data.iloc[0,0]
#tabula.read_pdf('C:/Users/corina/Downloads/17239562369.pdf',area=(170,500,190,556), pages=1, stream=True)[0].iloc[0,0]
file = open((desktop+file_name), 'rb')
readpdf = PyPDF2.PdfFileReader(file)
totalpages = readpdf.numPages

pdf_po = pd.DataFrame(columns=['PO#', 'P1', 'P2'], index=range(totalpages))

#pdf_po.at[0,'PageNumber'] = tabula.read_pdf('C:/Users/corina/Downloads/17252079327.pdf', area=(170, 500, 190, 556), pages=(1),stream=True)[0].iloc[0, 0]
for i in range(totalpages):
    try:
        pdf_po.at[i, 'P1'] = (i + 1)
        pdf_po.at[i, 'P2'] = (i + 1)
        pdf_po.at[i, 'PO#'] = tabula.read_pdf((desktop+file_name), area=(170, 500, 190, 556), pages=(i + 1),
                        stream=True)[0].iloc[0, 0]
        #pdf_po.at[i, 'PO#'] = tabula.read_pdf('C:/Users/corina/Downloads/17248138421.pdf', area=(170, 500, 190, 556), pages=(i+1),stream=True)[0].iloc[0, 0]
    except IndexError: #Error
        pdf_po.at[i, 'P1'] = (i + 1)
        pdf_po.at[i, 'P2'] = (i + 1)
        pdf_po.at[i, 'PO#'] = tabula.read_pdf((desktop+file_name), area=(170, 500, 190, 556), pages=(i),stream=True)[0].iloc[0, 0]
df = pdf_po.groupby('PO#').agg({'P1':'min', 'P2':'max'})[['P1','P2']].reset_index()
df = df.sort_values(by=['P1'])
df = df.reset_index()
del df['index']

df = df.merge(lowes_ord_ship, on="PO#", how='left')

for i in range(len(df)):
    pdfW = PdfFileWriter()
    if df.iloc[i,1]==df.iloc[i,2]: #PACKING SLIP HAS 1 PAGE
        pdfW.addPage(readpdf.getPage((df.iloc[i, 1] - 1)))
    else:
        pdfW.addPage(readpdf.getPage((df.iloc[i, 1] - 1)))
        pdfW.addPage(readpdf.getPage((df.iloc[i, 2] - 1)))
    output_filename = 'C:/Users/corina/Downloads/Lowe\'s PO#' + df.iloc[i, 0].astype(str) + " PS " + str(df.iloc[i, 3]) +'.pdf'
    with open(output_filename, 'wb') as output:
        pdfW.write(output)

pdfW = PdfFileWriter()


