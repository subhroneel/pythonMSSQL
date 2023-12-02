import tkinter
from tkinter import filedialog, GROOVE, ttk
from tkinter.ttk import *
from tkinter import *

import pandas as pd
import os
import sys


class csvToExcel:

    def __init__(self, root):
        self.root = root
        self.root.title('Database Exporter')
        self.root.geometry('850x300')
        self.dir_path_var = StringVar()
        self.xlFileName = StringVar()
        # self.conn: Connection = None
        # self.tableList: list

        srcExcelFileTxtFrame = LabelFrame(self.root, text="Save as Excel", font=("Arial", 9, "bold"), bg="#4d65a3",
                                          fg="white", bd=5, relief=GROOVE)
        srcExcelFileTxtFrame.place(x=0, y=0, width=400, height=50)

        # Create a Textbox to display the selected directory path
        textbox = Entry(srcExcelFileTxtFrame, textvariable=self.xlFileName, state="readonly", width=40)
        textbox.pack(pady=5)

        srcSaveExcelBtnFrame = LabelFrame(self.root, text="", font=("Arial", 12, "bold"), bg="#4d65a3",
                                          fg="white", bd=5, relief=GROOVE)
        srcSaveExcelBtnFrame.place(x=401, y=0, width=100, height=50)
        self.saveExcelBtn = Button(srcSaveExcelBtnFrame, text="SaveAs", command=lambda: self.save_Excel_File(), width=5,
                                   height=0, font=("arial", 9, "bold"), fg="#9ed6e6", bg="#4c4f55")
        self.saveExcelBtn.grid(row=0, column=0, padx=0, pady=5)
        self.saveExcelBtn.pack()

        srcDirTxtFrame = LabelFrame(self.root, text="Browse CSV Folder", font=("Arial", 9, "bold"), bg="#4d65a3",
                                    fg="white", bd=5, relief=GROOVE)
        srcDirTxtFrame.place(x=0, y=50, width=400, height=50)

        # Create a Textbox to display the selected directory path
        dirPathTxt = Entry(srcDirTxtFrame, textvariable=self.dir_path_var, state="readonly",
                           width=40)
        dirPathTxt.pack(pady=5)

        srcDirBtnFrame = LabelFrame(self.root, text="", font=("Arial", 12, "bold"),
                                    bg="#4d65a3", fg="white", bd=5, relief=GROOVE)
        srcDirBtnFrame.place(x=401, y=50, width=100, height=50)
        self.browseFolderBtn = Button(srcDirBtnFrame, text="Browse", command=lambda: self.open_directory(),
                                      state="disabled", width=5, height=0, font=("arial", 9, "bold"), fg="#9ed6e6",
                                      bg="#4c4f55")
        self.browseFolderBtn.grid(row=0, column=0, padx=0, pady=5)
        self.browseFolderBtn.pack()

        srcConvertBtnFrame = LabelFrame(self.root, text="", font=("Arial", 9, "bold"),
                                        bg="#4d65a3", fg="white", bd=5, relief=GROOVE)
        srcConvertBtnFrame.place(x=0, y=102, width=100, height=50)
        self.convertCSVBtn = Button(srcConvertBtnFrame, text="Convert", command=lambda: self.convertToExcel(), width=5,
                                    state="disabled", height=0, font=("arial", 9, "bold"), fg="#9ed6e6", bg="#4c4f55")
        self.convertCSVBtn.grid(row=0, column=10, padx=0, pady=5)
        self.convertCSVBtn.pack()

    def open_directory(self):
        directory_path = filedialog.askdirectory()
        if directory_path:
            self.dir_path_var.set(directory_path)
            # self.text_var.set(directory_path)
        if directory_path and self.get_directory_path():
            self.convertCSVBtn['state'] = NORMAL
        else:
            self.convertCSVBtn['state'] = DISABLED

    def get_directory_path(self):
        return self.dir_path_var.get()

    def get_ExcelFileName(self):
        return self.xlFileName.get()

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

    def convertToExcel(self):
        fileList = os.listdir(self.get_directory_path())
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
                    except pd.errors.EmptyDataError:
                        print(f"Warning: CSV file {x} is empty. Skipping.")
                    except Exception as e:
                        print(f"An exception occurred while processing {x}: {e}")
        self.saveExcelBtn['state'] = NORMAL
        self.browseFolderBtn['state'] = NORMAL
        self.convertCSVBtn['state'] = NORMAL


# Press the green button in the gutter to run the script.
# if __name__ == '__main__':
root = Tk()
csvToExcel(root)
root.mainloop()
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
