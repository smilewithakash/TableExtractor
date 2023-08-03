import os
import camelot
import pandas as pd

#try to extract tables without performing OCR
tables = camelot.read_pdf('Presentation2.pdf', flavor='stream', pages="1-end")

#If no tables are found, perform OCR and try again
if len(tables) == 0:
    os.system("ocrmypdf Presentation2.pdf output.pdf")
    tables = camelot.read_pdf('output.pdf', flavor='stream', pages="1-end")

#If tables are still not found, print a message and exit
if len(tables) == 0:
    print('No tables found in document.')
    exit()

writer = pd.ExcelWriter('output.xlsx', engine='openpyxl')

# Loop over the tables and write each one to a separate sheet in the Excel file
for i, table in enumerate(tables):
    # Convert the table to a pandas DataFrame
    df = table.df
    
    df.to_excel(writer, sheet_name=f'Table {i}', index=False)

writer.save()
