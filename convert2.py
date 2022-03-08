import win32com.client
import os
import glob

#This program converts those Powerpoint and Word documents to pdf which are in the folder

#stores the names of the PowerPoint documents
powPoints = []
#stores the names of the Word documents
words = []

#put the names of the Powerpoint and Word documents to the correct array
for x in glob.glob("*.pptx"):
    powPoints.append(x)
for x in glob.glob("*.docx"):
    words.append(x)

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
    print(out_file)
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = True
    pdf = word.Documents.Open(in_file)
    pdf.SaveAs(out_file, FileFormat = 17)
    pdf.Close()
    word.Quit()