#Copyright 2023 TukangM
#
#Licensed under the Creative Commons Zero v1.0 Universal License;
#you may not use this file except in compliance with the License.
#You may obtain a copy of the License at
#
#   https://github.com/TukangM/file_to_csv/blob/main/LICENSE
#
#Unless required by applicable law or agreed to in writing, software
#distributed under the License is distributed on an "AS IS" BASIS,
#WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#See the License for the specific language governing permissions and
#limitations under the License.

import os
import glob
import pandas as pd
from tkinter import *
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def ConvertToCSV():
    inputdir = textboxInputDir.get()
    outputfile = textboxOutputFile.get()
    filetype = textboxFileType.get()

    # Jika checkbox "Include Subdirectories" diaktifkan, tambahkan '**/' ke awal file type
    if checkboxSubdirectory.get():
        filetype = '**/' + filetype

    if not os.path.isdir(inputdir):
        consoleLog.insert(END, "Input directory not found!\n")
        return

    consoleLog.insert(END, "Please wait...\n")

    files = glob.glob(os.path.join(inputdir, filetype), recursive=checkboxSubdirectory.get())

    if not files:
        consoleLog.insert(END, "No files found!\n")
        return

    df = pd.DataFrame(files, columns=['File Path'])

    # Jika checkbox "Add Columns" diaktifkan, tambahkan kolom Name, Type/Extension, dan Size
    df['Name'] = df['File Path'].apply(os.path.basename)
    df['Type/Extension'] = df['File Path'].apply(lambda path: os.path.splitext(path)[1] if os.path.isfile(path) else 'Folders')
    df['Size'] = df['File Path'].apply(lambda path: os.path.getsize(path) if os.path.isfile(path) else 0)
    df['Size'] = df['Size'].apply(lambda size: sizeof_fmt(size))

    # Menyimpan versi CSV
    df.to_csv(outputfile, index=False, encoding='utf-8')
    consoleLog.insert(END, "Done!\n")

    # Menyimpan versi XLSX
    outputfileXLSX = outputfile.replace('.csv', '.xlsx')
    df.to_excel(outputfileXLSX, index=False)
    consoleLog.insert(END, "Done! Making .xlsx version\n")

def BrowseFolder():
    folderBrowser = filedialog.askdirectory(title="Select Input Directory")
    textboxInputDir.delete(0, END)
    textboxInputDir.insert(0, folderBrowser)

def BrowseSaveFile():
    saveFileDialog = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    textboxOutputFile.delete(0, END)
    textboxOutputFile.insert(0, saveFileDialog)

def sizeof_fmt(num, suffix='B'):
    for unit in ['','Ki','Mi','Gi','Ti','Pi','Ei','Zi']:
        if abs(num) < 1024.0:
            return "%3.1f %s%s" % (num, unit, suffix)
        num /= 1024.0
    return "%.1f %s%s" % (num, 'Yi', suffix)

# Main window
root = Tk()
root.title("File to CSV and XLSX Converter")

# Labels and Entry widgets
labelInputDir = Label(root, text="Input Directory:")
labelInputDir.grid(row=0, column=0, padx=5, pady=5)

textboxInputDir = Entry(root, width=50)
textboxInputDir.grid(row=0, column=1, columnspan=2, padx=5, pady=5)

buttonBrowseInput = Button(root, text="Browse", command=BrowseFolder)
buttonBrowseInput.grid(row=0, column=3, padx=5, pady=5)

labelOutputFile = Label(root, text="Output CSV File:")
labelOutputFile.grid(row=1, column=0, padx=5, pady=5)

textboxOutputFile = Entry(root, width=50)
textboxOutputFile.grid(row=1, column=1, columnspan=2, padx=5, pady=5)

buttonBrowseOutput = Button(root, text="Browse", command=BrowseSaveFile)
buttonBrowseOutput.grid(row=1, column=3, padx=5, pady=5)

labelFileType = Label(root, text="File Type:")
labelFileType.grid(row=2, column=0, padx=5, pady=5)

textboxFileType = Entry(root, width=20)
textboxFileType.grid(row=2, column=1, padx=5, pady=5)

checkboxSubdirectory = IntVar()
checkboxSubDir = Checkbutton(root, text="Include Subdirectories", variable=checkboxSubdirectory)
checkboxSubDir.grid(row=2, column=2, sticky=W)

buttonExecute = Button(root, text="Execute!", command=ConvertToCSV)
buttonExecute.grid(row=5, column=0, pady=10)

consoleLog = Text(root, height=10, width=60)
consoleLog.grid(row=6, column=0, columnspan=4, padx=10, pady=10)

root.mainloop()
