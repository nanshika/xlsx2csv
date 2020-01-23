# python Export_all.py abc.xlsx
import sys
from pathlib import Path
from subprocess import call

# pdf2txt.py のパス
py_path = Path(sys.exec_prefix) / "Scripts" / "pdf2txt.py"

# pdf2txt.py の呼び出し
call(["py", str(py_path), "-o extract-sample.txt", "-p 1", "extract-sample.pdf"])

import PyPDF2

FILE_PATH = './files/executive_order.pdf'

with open(FILE_PATH, mode='rb') as f:
    reader = PyPDF2.PdfFileReader(f)
    page = reader.getPage(0)
    print(page.extractText())

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfparser import PDFDocument
from pdfminer.pdfparser import PDFPage
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams
from pdfminer.layout import LTTextBoxHorizontal

import docx2txt
text = docx2txt.process('sample.docx')
text

fp = open('sample.pdf', 'rb')
parser = PDFParser(fp)
document = PDFDocument()
parser.set_document(document)
password=''
document.set_parser(parser)
document.initialize(password)
if not document.is_extractable:
    raise #PDFTextExtractionNotAllowed
rsrcmgr = PDFResourceManager()
laparams = LAParams()
device = PDFPageAggregator(rsrcmgr, laparams=laparams)

interpreter = PDFPageInterpreter(rsrcmgr, device)
pages = list(document.get_pages())
texts = []
for page in pages:
    interpreter.process_page(page)
    layout = device.get_result()
    for l in layout:
        if isinstance(l, LTTextBoxHorizontal):
            text = l.get_text()
            texts.append(text)

print('\n'.join(texts))

