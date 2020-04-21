import pandas as pd
# import pywintypes
# import win32api
import tkinter
import csv
import re
from tkinter import *
from tkinter import filedialog
import os

global file_name
global df


def import_csv():
    global df
    global file_name

    files = [('All Files', '*.*'),('csv Files', '*.csv'), ('Excel Files', '*.xlsx')]
    csv_file_path = filedialog.askopenfile(filetypes=files)
    if csv_file_path is None:
        return
    file_name.set(csv_file_path.name)
    name = os.path.basename(csv_file_path.name)
    if name.endswith('.csv'):
        df = pd.read_csv(csv_file_path, encoding="utf-8", header=None)
    if name.endswith('.xlsx'):
        xl = pd.ExcelFile(name)
        df = xl.parse("Sheet1", index_col=None, header=None)
        # df = pd.read_excel(csv_file_path, 'Sheet1', encoding="cp1252", index_col=None)
    print(df.head())
    df.columns = ["timestamp", "name", "text"]
    for index, col in df.iterrows():
        if re.search(r"ï»¿", str(col["timestamp"])):  # not sure how to fix encoding utf-8 here
            text = re.sub(r"ï»¿", "", str(col["timestamp"]))
            df.loc[index]["timestamp"] = text
            break
    # print(df.head())


def write_excel():
    files = [('All Files', '*.*'),
             ('Excel Files', '*.xlsx')]
    excel_save = filedialog.asksaveasfile(mode='w', filetypes=files, defaultextension=".xlsx")
    if excel_save is None:  # asksaveasfile return `None` if dialog closed with "cancel".
        return
    print(excel_save.name)
    df.to_excel(excel_save.name, header=False, index=False)


def clean_text(command):
    global df
    regex = ""

    if command == "square_bracket":
        regex = "(\[.*?\])"
    if command == "um":
        regex = "\s(um)\s*?|\s*?(um)\s|\s(uh)\s*?|\s*?(uh)\s|\s(ah)\s*?|\s*?(ah)\s"
    texts = []
    for index, col in df.iterrows():
        text = re.sub(regex, "", col["text"])
        texts.append(text)
    df["cleaned text"] = texts
    print(df["cleaned text"])


root = tkinter.Tk()
root.title("NGSX Clean Transcript")
w = 600  # width for the tkinter root
h = 300  # height for the tkinter root

ws = root.winfo_screenwidth()  # width of the screen
hs = root.winfo_screenheight()  # height of the screen

# calculate x and y coordinates for the tkinter root window
x = (ws / 4) - (w / 4)
y = (hs / 4) - (h / 4)

# set the dimensions of the screen
# and where it is placed
root.geometry('%dx%d+%d+%d' % (w, h, x, y))
tkinter.Button(root, text='Browse .csv or Excel File', command=import_csv).grid(sticky="w",row=1, column=0)
file_name = tkinter.StringVar()
tkinter.Label(root, textvariable=file_name,wraplength=400, justify=LEFT, anchor="w").grid(row=0, column=3)
tkinter.Button(root, text='Clean all square brackets', command=lambda:clean_text("square_bracket"), anchor="w").grid(sticky="w",row=2, column=0)
tkinter.Button(root, text='Clean all standalone um ah uh', command=lambda:clean_text("um"), anchor="w").grid(sticky="w",row=2, column=1)
tkinter.Button(root, text='Save as', command=write_excel,anchor="w").grid(sticky="w",row=3, column=0)
tkinter.Button(root, text='Close', command=root.destroy,anchor="w").grid(sticky="w",row=4, column=0)

bottom = tkinter.Frame(root)
bottom_label = tkinter.Label(bottom, text="Feature requests to CMai@clarku.edu")
bottom_label.pack(side="bottom", fill="x")
bottom.grid(pady=y-30)

root.mainloop()
