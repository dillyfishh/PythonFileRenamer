import sys, os
import tkinter as tk
from tkinter import filedialog, messagebox
import numpy as np
import pandas as pd
from natsort import natsorted

root = tk.Tk()
root.withdraw()

tk.messagebox.showinfo("Alert","Please Select Excel File")
dataImport = filedialog.askopenfilename()
filetype = dataImport.split('.', 1)[1].lower()

if not dataImport:
    tk.messagebox.showerror("File not Found.","You didn't select a file. Please Try Again.")
    sys.exit()
elif not filetype == "xlsx":
    tk.messagebox.showerror("Incorrect File Type","The file you have selected is the wrong filetype. Please Try Again with an xlsx file.")
    sys.exit()

messagebox.showinfo("Alert","Please Select Directory you would like to rename files in.")
dir = filedialog.askdirectory()
files = os.listdir(dir)

print(str.format("There are {} files in {}.",len(files),dir))

dfFilenames = pd.read_excel(dataImport)
fileList = []


if [*dfFilenames] == ['FileName']:
    for(_,filename) in dfFilenames.itertuples():
        fileList.append(filename)
print("UnSorted:",files)
files = natsorted(files)
print("Sorted:",files)

print(str.format("There are {} files in {} and {} names in the Excel Sheet",len(files),dir,len(fileList)))

if len(files) == len(fileList):
    for index, file in enumerate(files):
        newName = fileList[index] + ".docx"
        os.rename(os.path.join(dir,files[index]),os.path.join(dir,newName))
else:
    messagebox.showinfo("Alert",str.format("There are {} files in {} and {} names in the Excel Sheet. The file amount in the directory must be the same as the excel file. NOT INCLUDING THE HEADER",len(files),dir,len(fileList)))
    sys.exit()