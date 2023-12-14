# Python Fitz MuPDF
import fitz  
from tkinter import filedialog, GROOVE, messagebox, ttk
from tkinter.ttk import *
from tkinter import *
import os, sys

class splitPDF:

    def __init__(self, root):
        self.root = root
        self.root.title('Split PDF Document')
        self.root.geometry('620x400')
        #Filename with absolute path for source PDF
        self.sourcepdfFileName = StringVar()    
        #Absolute path for target splitted PDF path
        self.targetpdfFolderPath = StringVar()    
        #Total number of pages of source pdf
        self.srcPDFtotPages: int = 0
        #Variable binding from_page entry 
        self.from_pageno = IntVar()
        #Variable binding to_page entry 
        self.to_pageno = IntVar()

        #Container Frame for textbox entry for path of PDF file to split
        srcPDFFileTxtFrame = LabelFrame(self.root, text="Source PDf File", font=("Arial", 9, "bold"), 
                                          bg="#4d65a3", fg="white", bd=1, relief=GROOVE)
        srcPDFFileTxtFrame.place(x=0, y=0, width=420, height=50)

        #Textbox entry for path of PDF file to split
        Entry(srcPDFFileTxtFrame, textvariable=self.sourcepdfFileName, state="readonly", width=40).grid(row=0, column=0, padx=5, pady=5)
        
        #Button to browse and open the PDF file path 
        self.sourcePDFBtn = Button(srcPDFFileTxtFrame, text="Open PDF", command=lambda: self.open_source_PDF_File(), 
                                   width=5, height=0, font=("arial", 9, "bold"), fg="#9ed6e6", bg="#4c4f55")
        #Position of Button inside frame 
        self.sourcePDFBtn.grid(row=0, column=1, padx=5, pady=5)


        #Container Frame for setting the target path in entrybox to splitted PDF file 
        trgDirTxtFrame = LabelFrame(self.root, text="Browse PDF Target Folder", font=("Arial", 9, "bold"), bg="#4d65a3",
                                    fg="white", bd=5, relief=GROOVE)
        trgDirTxtFrame.place(x=0, y=50, width=420, height=50)

        #Setting the target directory path in entrybox to splitted PDF file 
        Entry(trgDirTxtFrame, textvariable=self.targetpdfFolderPath, state="readonly",
                           width=40).grid(row=0, column=0, padx=5, pady=5)


        #Button to browse and open the target path where splitted PDF file will be saved.
        self.browseFolderBtn = Button(trgDirTxtFrame, text="Browse", command=lambda: self.open_directory(),
                                      state="disabled", width=5, height=0, font=("arial", 9, "bold"), fg="#9ed6e6",
                                      bg="#4c4f55")
        self.browseFolderBtn.grid(row=0, column=1, padx=5, pady=5)

        #Container Frame for label and textbox entry for from page number filter
        targetPagenoFrame = LabelFrame(self.root, text="Page No to Split", font=("Arial", 9, "bold"), 
                                          bg="#4d65a3", fg="white", bd=1, relief=GROOVE)
        targetPagenoFrame.place(x=0, y=101, width=210, height=50)

        #Label for textbox entry for from page number filter
        Label(targetPagenoFrame, text="From Page", font=("Arial", 9, "bold"),bg="#4d65a3", fg="#ecd609").grid(row=0, column=0, padx=5, pady=5)

        #Textbox entry for from page number filter
        txtFromPageNo = Entry(targetPagenoFrame, textvariable=self.from_pageno, state="normal", width=3)
        txtFromPageNo.grid(row=0, column=1, padx=0, pady=5)
        txtFromPageNo.bind('<FocusOut>',self.validatePageNo)

        #Label for textbox entry for to page number filter
        Label(targetPagenoFrame, text="To Page", font=("Arial", 9, "bold"), bg="#4d65a3", fg="#ecd609").grid(row=0, column=2, padx=5, pady=5)

        #Textbox entry for to page number filter
        txtToPageNo = Entry(targetPagenoFrame, textvariable=self.to_pageno, state="normal", width=3)
        txtToPageNo.grid(row=0, column=3, padx=0, pady=5)
        txtToPageNo.bind('<FocusOut>', self.validatePageNo)
        

        #Container Frame for button for splitting source pdf pagewise with filtered pageno
        splitPDFBtnFrame = LabelFrame(self.root, text="Split PDF", font=("Arial", 9, "bold"), 
                                          bg="#4d65a3", fg="white", bd=1, relief=GROOVE)
        splitPDFBtnFrame.place(x=0, y=151, width=80, height=50)

        #Button for splitting source pdf pagewise with filtered pageno
        self.splitPDFBtn = Button(splitPDFBtnFrame, text="Split PDF", command=lambda: self.splitPDFFiles(),
                                      state="normal", width=5, height=0, font=("arial", 9, "bold"), fg="#9ed6e6",
                                      bg="#4c4f55")
        self.splitPDFBtn.grid(row=0, column=0, padx=5, pady=5)        
        self.splitPDFBtn.pack(pady=5)

        # Frame which contain ProgressBar showing splitted PDF export progress of a 
        # selected source PDF with multiple pages
        pbframe = LabelFrame(self.root, text="", font=("arial", 12, "bold"), bg="#8F00FF", fg="white",
                             bd=5, relief=GROOVE)
        pbframe.place(x=0, y=201, width=400, height=40)

        #ProgressBar showing splitted PDF export progress of a 
        #selected source PDF with multiple pages
        self.pb = Progressbar(pbframe, orient='horizontal', mode='determinate', length=280)
        self.pb.pack()
        self.pb['value'] = 0        

    #Validating pageno. Page no should not be less than zero, from pageno should not be 
    #greater than to pageno if so from and to pageno will be reset to default
    def validatePageNo(self, event):
        if self.from_pageno.get() >self.to_pageno.get() or self.from_pageno.get()<0 or self.to_pageno.get()<=0:
            self.from_pageno.set(1)
            self.to_pageno.set(self.srcPDFtotPages)

            

    #Function called from button self.sourcePDFBtn to browse and open the target path where splitted PDF file will be saved.
    def open_directory(self):
        directory_path = filedialog.askdirectory()
        if directory_path:
            self.targetpdfFolderPath.set(directory_path)
        self.splitPDFBtn['state'] = NORMAL


    #Function called on Button self.sourcePDFBtn to browse and open the PDF file path 
    def open_source_PDF_File(self):
        pdffile = filedialog.askopenfilename(defaultextension=".pdf",
                                                    filetypes=[("PDF files", "*.pdf")])
        if pdffile:
            self.sourcepdfFileName.set(pdffile)
            input_pdf = fitz.open(self.sourcepdfFileName.get())
            self.srcPDFtotPages = input_pdf.page_count
            input_pdf.close()
            self.from_pageno.set(1)
            self.to_pageno.set(self.srcPDFtotPages)
            self.browseFolderBtn['state'] = NORMAL

    #Splitting filename , filename extension and file directory path to extract filename
    #which will be used as prefix in splitted pdf files
    def getFileNameFromPAth(self, fileNameWithPAthAndExtension: str):
            if '/' in fileNameWithPAthAndExtension and '.' in fileNameWithPAthAndExtension:
                return fileNameWithPAthAndExtension.split('/')[-1:][0].split('.')[0]

    #Function called on Button self.splitPDFBtn to split the source multipages PDF files to 
    #single page multiple pdf files
    def splitPDFFiles(self):
        input_pdf = fitz.open(self.sourcepdfFileName.get())
        
        #Setting ProgressBar maimum value by getting the difference between from 
        #page and to page filter
        self.pb['maximum'] = self.to_pageno.get() - self.from_pageno.get() + 1
        
        self.pb['value'] = 0
        for page_index in range(self.from_pageno.get()-1,self.to_pageno.get()):
            # if page_index>= self.from_pageno.get() and page_index<=self.to_pageno.get():
            target = fitz.open()
            target.insert_pdf(input_pdf, from_page=page_index, to_page=page_index)
            #Filename of target splitted pdf is taken from source pdf indexed with page number
            output_file = f"{self.targetpdfFolderPath.get()}/{self.getFileNameFromPAth(self.sourcepdfFileName.get())}_{page_index + 1}.pdf"
            target.save(output_file)
            target.close()
            self.pb['value'] +=1
            self.root.update_idletasks()
        messagebox.showinfo(title="Succes", message="Split of PDF completed successfully")
        self.sourcepdfFileName.set("")
        self.targetpdfFolderPath.set("")
        self.srcPDFtotPages = 0
        self.from_pageno.set(0)
        self.to_pageno.set(0)
        self.sourcePDFBtn['state'] = NORMAL
        self.browseFolderBtn['state'] = DISABLED
        self.splitPDFBtn['state'] = DISABLED

root = Tk()
splitPDF(root)
root.mainloop()