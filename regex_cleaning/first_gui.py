import tkinter as tk
import pandas as pd
import csv
from tkinter import *
from tkinter import filedialog as fd


def import_csv():
    global file_name
    global df
    csv_file_path = fd.askopenfile(filetypes=[("csv files", "*.csv")])
    file_name.set(csv_file_path)
    df = pd.read_csv(csv_file_path)
    print(df.head())


def write_csv():
    files = [('All Files', '*.*'),
             ('csv Files', '*.csv'),
             ('Text Document', '*.txt')]
    mylist = [1,2,3]
    fp = fd.asksaveasfile(mode='w', filetypes=files, defaultextension=".csv")
    if fp is None:  # asksaveasfile return `None` if dialog closed with "cancel".
        return
    csv_writer = csv.writer(fp, delimiter=",",
                            quotechar="'",
                            quoting=csv.QUOTE_MINIMAL,
                            lineterminator="\n")
    csv_writer.writerow(mylist)


root = tk.Tk()
root.title("NGSX Clean Transcript")
w = 800  # width for the Tk root
h = 650  # height for the Tk root

ws = root.winfo_screenwidth()  # width of the screen
hs = root.winfo_screenheight()  # height of the screen

# calculate x and y coordinates for the Tk root window
x = (ws / 2) - (w / 2)
y = (hs / 2) - (h / 2)

# set the dimensions of the screen
# and where it is placed
root.geometry('%dx%d+%d+%d' % (w, h, x, y))
tk.Button(root, text='Browse .csv File', command=import_csv).grid(row=1, column=0)
file_name = tk.StringVar()
tk.Label(root, textvariable=file_name).grid(row=0, column=1)
tk.Button(root, text='Close', command=root.destroy).grid(row=1, column=2)
tk.Button(root, text='Save as', command=write_csv).grid(row=1, column=3)

root.mainloop()
