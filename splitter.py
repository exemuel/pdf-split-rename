#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Event Participant Certificates (PDF) Splitter
   
   This module splits a PDF file of event participant 
   certificates.
   
"""

__author__ = "Samuel I. G. Situmeang"
__copyright__ = "Copyright 2022, Institut Teknologi Del"
__credits__ = ["Samuel I. G. Situmeang"]
__license__ = "GPL"
__version__ = "1.0.0"
__maintainer__ = "Samuel I. G. Situmeang"
__email__ = "samsitumeang@gmail.com"
__status__ = "Development"


import os
import sys
import pandas as pd

from PyPDF2 import PdfFileWriter, PdfFileReader
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def main():
    Tk().withdraw()
    print("Please select the certificates file (PDF).")
    certificate_filename = askopenfilename()
    ext = os.path.splitext(certificate_filename)[-1].lower()
    if ext == ".pdf":
        print("You select", certificate_filename, "\n")
    else:
        print("Please select a pdf file.")
        sys.exit()


    print("Please select the participants file (XLSX).")
    info_filename = askopenfilename()
    ext = os.path.splitext(info_filename)[-1].lower()
    if ext == ".xlsx":
        print("You select", info_filename, "\n")
    else:
        print("Please select an xlsx file.")
        sys.exit()


    df = pd.read_excel(info_filename,
                       header = 0,
                       converters = {'No':int, 'Nama Lengkap':str,
                                     'Mata Pelajaran':str, 'Asal Sekolah':str,
                                     'No. Urut':str})
    inputpdf = PdfFileReader(open(certificate_filename, "rb"), strict=False)


    for i in range(inputpdf.numPages):
        output = PdfFileWriter()
        output.addPage(inputpdf.getPage(i))
        new_filename = "output/" + str(df.at[i, 'No. Urut'].split('/')[0]) + \
                       '-ITDel-PANDAI-WorkshopCT-Sertifikat-20221112-' + \
                       str(df.iloc[i].loc['Nama Lengkap'].split(',')[0]) + \
                       '.pdf'
        with open(new_filename, 'wb') as outputStream:
            output.write(outputStream)
            print(new_filename, "done!")

if __name__ == "__main__":
    main()
