# -*- coding: utf-8 -*-
"""
Created on Wed Jan 10 13:46:20 2018

@author: reema.malhotra
"""

#code that reads all doc,docx,pdf files from a location, converts to textfiles and save to a destination location
import os.path
import win32com.client

dir_path="C:\\Users\\reema.malhotra\\source"

dest="C:\\Users\\reema.malhotra\\dest"

# changing working directory to source directory
os.chdir(dir_path)

#filename='test1.doc'

error_dict = {}
wordapp = win32com.client.gencache.EnsureDispatch("Word.Application")
for filename in os.listdir(dir_path):
    if filename.split(".")[-1] in ['doc', 'DOC', 'docx', 'DOCX']:
        try:
            cv_key = filename.split(".")[0]
            #txtfile = cv_key + '.txt'
            print(cv_key + '.txt')
            wordapp.Visible = False
            wordapp.Documents.Open(os.path.abspath(filename))
            wdFormatTextLineBreaks = 1
            wdFormatUnicodeText = 7
            doc = wordapp.ActiveDocument
            tables = doc.Tables
            text1 = doc.Range().Text

            # print(text)
            text1 = text1.replace('\t', '')
            text1 = text1.replace('\x07', ' ')
            sv = os.path.join(dest, cv_key + ".txt")
            file1 = open(sv, "w", encoding='utf-8')
            file1.write(text1)
            file1.close()
            wordapp.ActiveDocument.Close()

            
        except BaseException as e:
            error_dict[cv_key] = str(e)

    elif (filename.split(".")[-1] == "pdf"):
        print(filename)
        
        try:
            cv_key = filename.split(".")[0]
            import PyPDF2

            sv = os.path.join(dest, cv_key + ".txt")
            file1 = open(sv, "w", encoding='utf-8')
            pdfFileObj = open(filename, 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            pages = pdfReader.getNumPages()
            cv_read = ""
            for p in range(pages):
                pageObj = pdfReader.getPage(p)
                page = pdfReader.getPage(p)
                text = page.extractText()
                file1.write(text)
            file1.close()

        except BaseException as e:
            error_dict[cv_key] = str(e)

