# PyMuPDF
import fitz  
from tkinter import filedialog, GROOVE, messagebox, ttk
from tkinter.ttk import *
from tkinter import *
import os, sys
from PDFNetPython3.PDFNetPython import PDFDoc, Optimizer, SDFDoc, PDFNet
class mergePDF:

    def __init__(self, root):
        self.root = root
        self.root.title('CSV to Excel Converter')
        self.root.geometry('620x400')
        self.dir_path_var = StringVar()
        self.pdfFileName = StringVar()
        self.Checkbutton = IntVar()
        self.fileList: list = [str]
        # self.conn: Connection = None
        # self.tableList: list

        #Container Frame for textbox entry for new excel file path
        srcPDFFileTxtFrame = LabelFrame(self.root, text="Merge to PDF", font=("Arial", 9, "bold"), 
                                          bg="#4d65a3", fg="white", bd=5, relief=GROOVE)
        srcPDFFileTxtFrame.place(x=0, y=0, width=400, height=50)

        #New textbox entry for new excel file path
        textbox = Entry(srcPDFFileTxtFrame, textvariable=self.pdfFileName, state="readonly", width=40)
        textbox.pack(pady=5)

        #Container Frame for Save As button to save the excel file on the path mentioned in textbox above
        srcTargetPDFBtnFrame = LabelFrame(self.root, text="", font=("Arial", 12, "bold"), bg="#4d65a3",
                                          fg="white", bd=5, relief=GROOVE)
        srcTargetPDFBtnFrame.place(x=401, y=0, width=100, height=50)

        #Save As button to save the excel file on the path mentioned in textbox above
        self.TargetPDFBtn = Button(srcTargetPDFBtnFrame, text="Save As", command=lambda: self.save_target_PDF_File(), 
                                   width=5, height=0, font=("arial", 9, "bold"), fg="#9ed6e6", bg="#4c4f55")
        self.TargetPDFBtn.grid(row=0, column=0, padx=0, pady=5)
        self.TargetPDFBtn.pack()

        #Container Frame for setting the path in entrybox to csv file source folder for bulk conversion from csv to excel
        srcDirTxtFrame = LabelFrame(self.root, text="Browse PDF Source Folder", font=("Arial", 9, "bold"), bg="#4d65a3",
                                    fg="white", bd=5, relief=GROOVE)
        srcDirTxtFrame.place(x=0, y=50, width=400, height=50)

        #Setting the path in entrybox to csv file source folder for bulk conversion from csv to excel
        dirPathTxt = Entry(srcDirTxtFrame, textvariable=self.dir_path_var, state="readonly",
                           width=40)
        dirPathTxt.pack(pady=5)

        #Container Frame for containing button for browsing the csv source folder for setting the 
        #path in entrybox for bulk conversion from csv to excel
        srcDirBtnFrame = LabelFrame(self.root, text="", font=("Arial", 12, "bold"),
                                    bg="#4d65a3", fg="white", bd=5, relief=GROOVE)
        srcDirBtnFrame.place(x=401, y=50, width=100, height=50)

        #Button for browsing the csv source folder for setting the path in entrybox 
        #for bulk conversion from csv to excel
        self.browseFolderBtn = Button(srcDirBtnFrame, text="Browse", command=lambda: self.open_directory(),
                                      state="disabled", width=5, height=0, font=("arial", 9, "bold"), fg="#9ed6e6",
                                      bg="#4c4f55")
        self.browseFolderBtn.grid(row=0, column=0, padx=0, pady=5)
        self.browseFolderBtn.pack()

        # Frame which contain ProgressBar showing tables export progress of a selected schema of database
        tblListframe = LabelFrame(self.root, text="List of source PDFs", font=("arial", 12, "bold"), bg="#8F00FF",
                                  fg="white",
                                  bd=5, relief=GROOVE)
        tblListframe.place(x=0, y=102, width=600, height=200)
        #Vertical Scrollbar for schema tables listbox
        scroll_y = Scrollbar(tblListframe, orient=VERTICAL)
        scroll_x = Scrollbar(tblListframe, orient=HORIZONTAL)

        #Database Schema tables listbox
        self.LB = Listbox(tblListframe, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set, selectbackground="#8d8df6", selectmode=EXTENDED,
                          font=("arial", 12, "bold"), bg="#c9f56f", fg="navyblue", bd=5, relief=GROOVE)
        self.LB.bind('<<ListboxSelect>>', self.items_selected)
        # self.LB.pack(fill=BOTH)

        #Vertical Scrollbar configuration for schema tables listbox
        scroll_y.config(command=self.LB.yview)
        scroll_x.config(command=self.LB.xview)
        scroll_y.pack(side=RIGHT, fill=Y)
        scroll_x.pack(side=BOTTOM, fill=X)
        self.LB.pack(expand=True, fill=BOTH, side=LEFT)
        self.LB['state'] = DISABLED

        tblBrowseFileframe = LabelFrame(self.root, text="", font=("arial", 12, "bold"), bg="#4d65a3",
                                  fg="white", bd=5, relief=GROOVE)
        tblBrowseFileframe.place(x=601, y=102, width=20, height=20)

        #Button for browsing the csv source folder for setting the path in entrybox 
        #for bulk conversion from csv to excel
        self.browseFileBtn = Button(tblBrowseFileframe, text="dasdsadasd", command=lambda: self.choose_file(),
                                      state="disabled", width=5, height=5, font=("arial", 9, "bold"), 
                                      fg="#9ed6e6",bg="#4c4f55")
        self.browseFileBtn.grid(row=0, column=0, padx=0, pady=5)
        self.browseFileBtn.pack()
        
        tblChkFileframe = LabelFrame(self.root, text="", font=("arial", 12, "bold"), bg="#4d65a3",
                                  fg="white", bd=5, relief=GROOVE)
        tblChkFileframe.place(x=601, y=125, width=20, height=20)


        self.chkButton = Checkbutton(tblChkFileframe,command=self.onCheck,text = "",variable = self.Checkbutton, 
                      state="disabled", onvalue = 1, offvalue = 0, height = 2, width = 10) 
        self.chkButton.grid(row=1, column=0, padx=0, pady=5)
        self.chkButton.pack()


        
        #Container Frame for containing button for conversion from csv to excel
        srcMergeBtnFrame = LabelFrame(self.root, text="", font=("Arial", 9, "bold"),
                                        bg="#4d65a3", fg="white", bd=5, relief=GROOVE)
        srcMergeBtnFrame.place(x=0, y=303, width=100, height=50)

        #Button for conversion from csv to excel
        self.mergePDFBtn = Button(srcMergeBtnFrame, text="Merge", command=lambda: self.mergeToSinglePDF(), width=5,
                                    state="disabled", height=0, font=("arial", 9, "bold"), fg="#9ed6e6", bg="#4c4f55")
        self.mergePDFBtn.grid(row=0, column=0, padx=0, pady=5)
        self.mergePDFBtn.pack()

        # Frame which contain ProgressBar showing tables export progress of a selected schema of database
        pbframe = LabelFrame(self.root, text="", font=("arial", 12, "bold"), bg="#8F00FF", fg="white",
                             bd=5, relief=GROOVE)
        pbframe.place(x=0, y=354, width=400, height=40)

        #ProgressBar showing tables export progress of a selected schema of database
        self.pb = Progressbar(pbframe, orient='horizontal', mode='determinate', length=280)
        # self.pb.grid(column=0, row=0, columnspan=2, padx=20, pady=40)
        self.pb.pack()
        self.pb['value'] = 0

   #Open directory function to browse directory path for pdf source folder
    def open_directory(self):
        directory_path = filedialog.askdirectory()
        if directory_path:
            self.dir_path_var.set(directory_path)
            # self.text_var.set(directory_path)
            self.fileList = sorted(os.listdir(self.get_directory_path()), key=self.sort_criteria, reverse=True)
            self.LB['state'] = NORMAL
            self.LB.delete(0,END)
            for item in self.fileList:
                if self.isPDFExtensionName(item):
                    self.LB.insert(END, item)
            self.mergePDFBtn['state'] = NORMAL
            self.chkButton['state'] = NORMAL
            # self.LB.pack(expand=True, fill=BOTH, side=LEFT)
        else:
            self.mergePDFBtn['state'] = DISABLED
            self.chkButton['state'] = DISABLED

    #Short creteria is set on datetime in ascending order used for os.listdir
    def sort_criteria(self,item):
        return os.path.getctime(os.path.join(self.get_directory_path(), item))
    
    def onCheck(self):
        if self.Checkbutton.get() == 1:
            self.LB.selection_set(0,END)
        else:
            self.LB.selection_clear(0, END)

        selected_indices = self.LB.curselection()
        # get selected items
        selected_langs = ",".join([self.LB.get(i) for i in selected_indices])
        self.selectedfileleList = selected_langs.split(',')
        if len(selected_langs) > 0:
            self.mergePDFBtn['state'] = NORMAL
        else:
            self.mergePDFBtn['state'] = DISABLED


    def choose_file(self):
        self.pdffiles = filedialog.askopenfilenames(defaultextension=".pdf",
                                                    filetypes=[("PDF files", "*.pdf")])
        if self.pdffiles:
            self.LB['state'] = NORMAL
            for item in self.pdffiles:
                if self.isPDFExtensionName(item):
                    self.LB.insert(END, item)
                    self.fileList.append(item)
            self.chkButton['state'] = NORMAL

    def isPDFExtensionName(self, filename: str):
        return filename.lower().endswith('pdf')
            
    #Get directory path returns path of chosen directory
    def get_directory_path(self):
        return self.dir_path_var.get()

    #Get the excel filename that is set while saving the new excel file
    def get_pdfFileName(self):
        return self.pdfFileName.get()

    #Function to save the new excel file which will be used to append csv file content
    def save_target_PDF_File(self):
        pdfFileName = filedialog.asksaveasfilename(initialfile='Output.pdf',
                                                     defaultextension=".pdf",
                                                     filetypes=[("PDF Documents", "*.pdf")])
        self.pdfFileName.set(pdfFileName)
        if pdfFileName:
            self.browseFolderBtn['state'] = NORMAL
            self.browseFileBtn['state'] = NORMAL
        else:
            self.browseFolderBtn['state'] = DISABLED
            self.mergePDFBtn['state'] = DISABLED

    def items_selected(self, event):
        # get selected indices
        selected_indices = self.LB.curselection()
        # get selected items
        selected_langs = ",".join([self.LB.get(i) for i in selected_indices])
        self.selectedfileleList = selected_langs.split(',')
        if len(self.selectedfileleList) > 0:
            self.mergePDFBtn['state'] = NORMAL
        else:
            self.mergePDFBtn['state'] = DISABLED


    def getFormattedFileSize(self, size: int):
        if size/(1024*1024) == 0:
             if size/1024 == 0:
                 return str(size) + ' Bytes'
             else:
                return str(round(size*1.0/1024,2)) + ' KB'
        else:
            return str(round(size*1.0/(1024*1024),2)) + ' MB'


        
    def compress_pdf(self, pdf_file):
        # Load PDF using pymupdf
        doc = fitz.open(pdf_file)
        compressed_pdf_bytes = doc.tobytes(
            deflate=True,
            garbage=4,
        )
        doc.save(pdf_file)
        print(len(compressed_pdf_bytes)) # Check output of compressed pdf
        return compressed_pdf_bytes


    #Function to convert all csv files into the excel file in individual 
    #sheets with sheetname as csv filename
    def mergeToSinglePDF(self):
        if len(self.selectedfileleList)==0:
            return
        self.pb['maximum'] = len(self.selectedfileleList)
        self.pb['value'] = 0
        target_pdf_file_path = self.get_pdfFileName()
        self.browseFolderBtn['state'] = DISABLED
        self.mergePDFBtn['state'] = DISABLED
        target = fitz.open()
        for x in self.selectedfileleList:
            src_pdf_file_path = self.get_directory_path() + '/' + x
            try:
                input_pdf = fitz.open(src_pdf_file_path)
                target.insert_pdf(input_pdf)
                self.pb['value'] +=1
                self.root.update_idletasks()
            except Exception as e:
                print(f"An exception occurred while processing {x}: {e}")
                messagebox.showerror(title="Error", message="An exception occurred while processing {x}: {e}")
                return
        target.save(target_pdf_file_path, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, pretty=True)
        target.close()
        initial_filesize = os.path.getsize(target_pdf_file_path)
        # compressed_size = len(self.compress_pdf(target_pdf_file_path))
        # messagebox.showinfo(title="Succes", message="Conversion from CSV to Excel completed successfully with size = " + self.getFormattedFileSize(initial_filesize) + " compressed to " + self.getFormattedFileSize(compressed_size))
        messagebox.showinfo(title="Succes", message="Conversion from CSV to Excel completed successfully with size = " + self.getFormattedFileSize(initial_filesize))
        self.dir_path_var.set("")
        self.pdfFileName.set("")
        self.TargetPDFBtn['state'] = NORMAL
        # self.browseFolderBtn['state'] = DISABLED
        self.mergePDFBtn['state'] = DISABLED
        self.LB.delete(0,END)
        self.LB['state'] = DISABLED
    

# Press the green button in the gutter to run the script.
# if __name__ == '__main__':
root = Tk()
mergePDF(root)
root.mainloop()
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
