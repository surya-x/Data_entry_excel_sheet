from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

import time
import io
import openpyxl
import sys
from config import *


def convert_pdf_to_txt(path, data):
    logging.info("method 'convert_pdf_to_txt' called ")
    try:
        rsrcmgr = PDFResourceManager()
        retstr = io.StringIO()
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, codec='utf-8', laparams=laparams)
        fp = open(path, 'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = ""
        maxpages = 0
        caching = True
        pagenos = set()

        for pageN, page in enumerate(PDFPage.get_pages(fp, pagenos, maxpages=maxpages,
                                      password=password,
                                      caching=caching,
                                      check_extractable=True)):
            if pageN % 2 == 0 :
                interpreter.process_page(page)
                text = retstr.getvalue()
                data.append(text)

        fp.close()
        device.close()
        retstr.close()
        return data
    except Exception as e:
        print("Error : Reading pages from pdf")
        print(e)
        logging.error(e)
        sys.exit()

def write_data(excel_path, excel_format, data):
    logging.info("method write_data called ")
    try:
        workbook = openpyxl.load_workbook(excel_format)
        sheet = workbook.worksheets[0]

        for index, each_row in enumerate(data):
            for i, each_cell in zip(range(1, 13), each_row):
                sheet.cell(row=index+2, column=i).value = each_cell

        workbook.save(excel_path)
    except PermissionError:
        print("Error Occured : Close the Excel Window before executing code")
        sys.exit()
    except Exception as e:
        print("Error : Writing in full_pay.xlsx")
        print(e)
        logging.error(e)
        sys.exit()