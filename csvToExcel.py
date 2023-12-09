from tkinter import filedialog, GROOVE, messagebox, ttk
from tkinter.ttk import *
from tkinter import *

import pandas as pd
import os


class csvToExcel:

    def __init__(self, root):
        self.root = root
        self.root.title('CSV to Excel Converter')
        self.root.geometry('500x200')
        self.dir_path_var = StringVar()
        self.xlFileName = StringVar()
        # self.conn: Connection = None
        # self.tableList: list

        #Container Frame for textbox entry for new excel file path
        srcExcelFileTxtFrame = LabelFrame(self.root, text="Save as Excel", font=("Arial", 9, "bold"), 
                                          bg="#4d65a3", fg="white", bd=5, relief=GROOVE)
        srcExcelFileTxtFrame.place(x=0, y=0, width=400, height=50)

        #New textbox entry for new excel file path
        textbox = Entry(srcExcelFileTxtFrame, textvariable=self.xlFileName, state="readonly", width=40)
        textbox.pack(pady=5)

        #Container Frame for Save As button to save the excel file on the path mentioned in textbox above
        srcSaveExcelBtnFrame = LabelFrame(self.root, text="", font=("Arial", 12, "bold"), bg="#4d65a3",
                                          fg="white", bd=5, relief=GROOVE)
        srcSaveExcelBtnFrame.place(x=401, y=0, width=100, height=50)

        #Save As button to save the excel file on the path mentioned in textbox above
        self.saveExcelBtn = Button(srcSaveExcelBtnFrame, text="SaveAs", command=lambda: self.save_Excel_File(), 
                                   width=5, height=0, font=("arial", 9, "bold"), fg="#9ed6e6", bg="#4c4f55")
        self.saveExcelBtn.grid(row=0, column=0, padx=0, pady=5)
        self.saveExcelBtn.pack()

        #Container Frame for setting the path in entrybox to csv file source folder for bulk conversion from csv to excel
        srcDirTxtFrame = LabelFrame(self.root, text="Browse CSV Folder", font=("Arial", 9, "bold"), bg="#4d65a3",
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

        #Container Frame for containing button for conversion from csv to excel
        srcConvertBtnFrame = LabelFrame(self.root, text="", font=("Arial", 9, "bold"),
                                        bg="#4d65a3", fg="white", bd=5, relief=GROOVE)
        srcConvertBtnFrame.place(x=0, y=102, width=100, height=50)

        #Button for conversion from csv to excel
        self.convertCSVBtn = Button(srcConvertBtnFrame, text="Convert", command=lambda: self.convertToExcel(), width=5,
                                    state="disabled", height=0, font=("arial", 9, "bold"), fg="#9ed6e6", bg="#4c4f55")
        self.convertCSVBtn.grid(row=0, column=0, padx=0, pady=5)
        self.convertCSVBtn.pack()

        # Frame which contain ProgressBar showing tables export progress of a selected schema of database
        pbframe = LabelFrame(self.root, text="", font=("arial", 12, "bold"), bg="#8F00FF", fg="white",
                             bd=5, relief=GROOVE)
        pbframe.place(x=0, y=153, width=400, height=40)

        #ProgressBar showing tables export progress of a selected schema of database
        self.pb = Progressbar(pbframe, orient='horizontal', mode='determinate', length=280)
        # self.pb.grid(column=0, row=0, columnspan=2, padx=20, pady=40)
        self.pb.pack()
        self.pb['value'] = 0

    #Open directory function to browse directory path for csv source folder
    def open_directory(self):
        directory_path = filedialog.askdirectory()
        if directory_path:
            self.dir_path_var.set(directory_path)
            # self.text_var.set(directory_path)
        if directory_path and self.get_directory_path():
            self.convertCSVBtn['state'] = NORMAL
        else:
            self.convertCSVBtn['state'] = DISABLED

    #Get directory path returns path of chosen directory
    def get_directory_path(self):
        return self.dir_path_var.get()

    #Get the excel filename that is set while saving the new excel file
    def get_ExcelFileName(self):
        return self.xlFileName.get()

    #Function to save the new excel file which will be used to append csv file content
    def save_Excel_File(self):
        excelFileName = filedialog.asksaveasfilename(initialfile='Untitled.xlsx',
                                                     defaultextension=".xlsx",
                                                     filetypes=[("Excel (2003) Documents", "*.xls"),
                                                                ("Excel Documents", "*.xlsx")])
        self.xlFileName.set(excelFileName)
        if excelFileName:
            self.browseFolderBtn['state'] = NORMAL
        else:
            self.browseFolderBtn['state'] = DISABLED
            self.convertCSVBtn['state'] = DISABLED

    #Function to convert all csv files into the excel file in individual 
    #sheets with sheetname as csv filename
    def convertToExcel(self):
        fileList = os.listdir(self.get_directory_path())
        self.pb['maximum'] = len(fileList)
        self.pb['value'] = 0
        excel_file_path = self.get_ExcelFileName()
        self.saveExcelBtn['state'] = DISABLED
        self.browseFolderBtn['state'] = DISABLED
        self.convertCSVBtn['state'] = DISABLED
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter', mode='w+') as writer:
            for x in fileList:
                csv_file_path = self.get_directory_path() + '/' + x
                if x[-4:].lower() == '.csv':
                    try:
                        df = pd.read_csv(csv_file_path, sep='~')
                        df.to_excel(writer, index=False, sheet_name=x[:-4])
                        self.pb['value'] +=1
                        self.root.update_idletasks()
                    except pd.errors.EmptyDataError:
                        print(f"Warning: CSV file {x} is empty. Skipping.")
                        messagebox.showinfo(title="Warning", message="CSV file {x} is empty. Skipping.")
                    except Exception as e:
                        print(f"An exception occurred while processing {x}: {e}")
                        messagebox.showinfo(title="Error", message="An exception occurred while processing {x}: {e}")
                        return
        
        messagebox.showinfo(title="Succes", message="Conversion from CSV to Excel completed successfully")
        self.dir_path_var.set("")
        self.xlFileName.set("")
        self.saveExcelBtn['state'] = NORMAL
        self.browseFolderBtn['state'] = DISABLED
        self.convertCSVBtn['state'] = DISABLED


# Press the green button in the gutter to run the script.
# if __name__ == '__main__':
root = Tk()
csvToExcel(root)
root.mainloop()
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
