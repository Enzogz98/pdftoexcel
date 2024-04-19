import os
import PyPDF2
import pandas as pd
import re


def pdf_to_excel(pdf_path, excel_path):
    with open(pdf_path, 'rb') as file:
        pdf_reader=PyPDF2.PdfReader(file)
        text= ''
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]

            text += page.extract_text()

    text_clean = ''.join(char for char in text if ord(char) < 128)

    text_clean = ''.join(char if char.isprintable() else ' ' for char in text_clean)

    lines = text_clean.split('\n')
    df=pd.DataFrame(lines)

    df.to_excel(excel_path, index=False, header=False)

def convertir_pdfs_a_excel(directorio_pdf,directorio_excel):
    if not os.path.exists(directorio_excel):
        os.makedirs(directorio_excel)

    for filename in os.listdir(directorio_pdf):
        if filename.endswith('.pdf'):
            pdf_path=os.path.join(directorio_pdf,filename)
            excel_path=os.path.join(directorio_excel,filename.replace('.pdf','.xlsx'))
            pdf_to_excel(pdf_path,excel_path)
            print(f'{pdf_path} convertido a {excel_path}')

directorio_pdf= 'C:\Enzo\pdfs'
directorio_excel='C:\Enzo\excel'

convertir_pdfs_a_excel(directorio_pdf,directorio_excel)
