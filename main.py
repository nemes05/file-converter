from tkinter import *
from tkinter import filedialog
import customtkinter
from os import path
import win32com.client
import os
import glob 

class App:
    def __init__(self):
        self.words = []
        self.excels = []
        self.powerPoints = []
        self.folder = ""

        customtkinter.set_appearance_mode('dark')
        root = customtkinter.CTk()
        root.geometry("300x400")

        self.word = customtkinter.StringVar(value=0)
        self.powerPoint = customtkinter.StringVar(value=0)
        self.excel = customtkinter.StringVar(value=0)

        openFolder = customtkinter.CTkButton(master=root, text="Choose folder", command=self.select_folder)
        convert = customtkinter.CTkButton(master=root, text="Convert", command=self.select_files)
        button = customtkinter.CTkButton(master=root, text="Quit", command=root.destroy)
        
        wordCheckBox = customtkinter.CTkCheckBox(root, text="Word", variable=self.word)
        powerPointCheckBox = customtkinter.CTkCheckBox(root, text="PowerPoint", variable=self.powerPoint)
        excelCheckBox = customtkinter.CTkCheckBox(root, text="Excel",  variable=self.excel)

        self.error_label = customtkinter.CTkLabel(root, text="", text_color="red")
        self.folder_label = customtkinter.CTkLabel(root, text="")
        
        openFolder.place(relx=0.5, rely=0.6, anchor='center')
        convert.place(relx=0.5, rely=0.7, anchor='center')
        button.place(relx=0.5, rely=0.8, anchor='center')
        
        wordCheckBox.place(relx=0.5, rely=0.3, anchor='center')
        powerPointCheckBox.place(relx=0.5, rely=0.2, anchor='center')
        excelCheckBox.place(relx=0.5, rely=0.1, anchor='center')

        self.error_label.place(relx=0.5,rely=0.9, anchor='center')
        self.folder_label.place(relx=0.5,rely=0.4, anchor='center')
        
        root.mainloop()
    def select_folder(self):
        self.folder = filedialog.askdirectory(initialdir='/')
        self.error_label.configure(text='')
        self.folder_label.configure(text='Path: ' + self.folder)
    def select_files(self):
        if(self.folder == ''):
            self.error_label.configure(text='First select a folder')
        if(self.word.get() == '1'):
            for x in glob.glob(path.join(self.folder,'*.{}'.format('docx'))):
                self.words.append(x)
            self.convert(self.words,'.docx')
        if(self.powerPoint.get() == '1'):
            for x in glob.glob(path.join(self.folder,'*.{}'.format('pptx'))):
                self.powerPoints.append(x)
            self.convert(self.powerPoints,'.pptx')
        if(self.excel.get() == '1'):
            for x in glob.glob(path.join(self.folder,'*.{}'.format('xlsx'))):
                self.excels.append(x)
            self.convert(self.excels,'.xlsx')
    def convert(self, array, extension):
       for x in array:
            in_file = os.path.realpath(x)
            x = x.replace(extension,'')
            out_file = x + ".pdf"
            if(extension == '.docx'):
                file = win32com.client.gencache.EnsureDispatch('Word.Application')
                file.Visible = False
                pdf = file.Documents.Open(in_file)
                pdf.SaveAs(out_file, FileFormat = 17)
            elif(extension == '.pptx'):
                file = win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
                pdf = file.Presentations.Open(in_file, WithWindow=False)
                print(out_file)
                pdf.SaveAs(out_file, 32)
            elif(extension == '.xlsx'):
                file = win32com.client.gencache.EnsureDispatch('Excel.Application')
                file.Visible=False
                pdf = file.Workbooks.Open(in_file)
                pdf.ExportAsFixedFormat(win32com.client.constants.xlTypePDF, out_file)
            pdf.Close()
            file.Quit()

App()