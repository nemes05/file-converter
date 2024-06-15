#Legacy code without UI
from typing import Type
import win32com.client
import os
import glob

#This program converts those Powerpoint and Word documents to pdf which are in the folder

#stores the names of the PowerPoint documents
powPoints = []
#stores the names of the Word documents
words = []
#stores the names of the Excel documents
excels = []

#put the names of the Powerpoint, Word and Excel documents to the correct array
for x in glob.glob("*.pptx"):
    powPoints.append(x)
for x in glob.glob("*.docx"):
    words.append(x)
for x in glob.glob("*.xls"):
    excels.append(x)

#converts the Powerpoints to pdf
for x in powPoints:
    in_file = os.path.realpath(x)
    x = x.replace('.pptx','')
    out_file = os.getcwd() + "\\" + x + ".pdf"
    powerpoint = win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
    powerpoint.Visible = True
    pdf = powerpoint.Presentations.Open(in_file)
    pdf.SaveAs(out_file, FileFormat = 32)
    pdf.Close()
    powerpoint.Quit()

#converts the Word documents to pdf 
for x in words:
    in_file = os.path.realpath(x)
    x = x.replace('.docx','')
    out_file = os.getcwd() + "\\" + x + ".pdf"
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = True
    pdf = word.Documents.Open(in_file)
    pdf.SaveAs(out_file, FileFormat = 17)
    pdf.Close()
    word.Quit()

#converts the Excel documents to pdf 
for x in excels:
    in_file = os.path.realpath(x)
    x = x.replace('.xls','')
    out_file = os.getcwd() + "\\" + x + ".pdf"
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    pdf = excel.Workbooks.Open(in_file)
    #pdf.SaveAs(out_file, FileFormat = 17)
    pdf.ExportAsFixedFormat(0,out_file)
    pdf.Close()
    excel.Quit()
