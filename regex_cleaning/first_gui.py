import pandas as pd
# import pywintypes
# import win32api
import tkinter
import csv
import re
from tkinter import *
from tkinter import filedialog
import os
import xlrd

global FILE_NAME
global df
global root
global header_var, replace_text_var


def import_csv():
    global df
    global FILE_NAME
    global root

    files = [('All Files', '*.*'), ('csv Files', '*.csv'), ('Excel Files', '*.xlsx')]
    allowed_ext = ['.csv','.xlsx']

    while True:
        csv_file_path = filedialog.askopenfile(filetypes=files)
        if csv_file_path is None:
            return
        extension = os.path.splitext(csv_file_path.name)[1]
        if extension in allowed_ext:
            break
        else:
            popupdialog("Invalid file types. Only .csv and .xlsx. Please try again")
            root.protocol("WM_DELETE_WINDOW", root.destroy)
    name = os.path.basename(csv_file_path.name)
    FILE_NAME.set(name)
    if name.endswith('.csv'):
        df = pd.read_csv(csv_file_path, encoding="utf-8", header=None)
    if name.endswith('.xlsx'):
        xl = xlrd.open_workbook(csv_file_path.name, encoding_override='utf-8')
        df = pd.read_excel(xl, header=None)
        # xl = pd.ExcelFile(csv_file_path)
        print(csv_file_path)
        # df = xl.parse("Sheet1", index_col=None, header=None)
        # df = pd.read_excel(csv_file_path, 'Sheet1', encoding="cp1252", index_col=None)
    df.columns = ["timestamp", "name", "text"]
    for index, col in df.iterrows():
        if re.search(r"ï»¿", str(col["timestamp"])):  # not sure how to fix encoding utf-8 here
            text = re.sub(r"ï»¿", "", str(col["timestamp"]))
            df.loc[index]["timestamp"] = text
            break
    # print(df.head())


def write_excel():
    global df
    global header_var, replace_text_var

    header = bool(header_var.get())
    replace = bool(replace_text_var.get())
    if replace:
        df.drop(columns=["text"], inplace=True)
    files = [('All Files', '*.*'),
             ('Excel Files', '*.xlsx')]
    excel_save = filedialog.asksaveasfile(mode='w', filetypes=files, defaultextension=".xlsx")
    if excel_save is None:  # asksaveasfile return `None` if dialog closed with "cancel".
        return
    print(excel_save.name)
    try:
        df.to_excel(excel_save.name, index=False, header=header)
        saved_message = "File saved as " + os.path.basename(excel_save.name)
        done_message(5, 1, saved_message)
    except OSError as e:
        print(e)
        popupdialog("OS error")
        root.protocol("WM_DELETE_WINDOW", root.destroy)


def clean_text(command):
    global df
    regex = ""
    row = 0

    if command == "clear_square_bracket":
        regex = "(\[.*?\])"
        row = 2
    if command == "clear_um":
        # regex = "\s(um)\s*?|\s*?(um)\s|\s(uh)\s*?|\s*?(uh)\s|\s(ah)\s*?|\s*?(ah)\s"
        regex = "\s(um|uh|ah)\s*?|\s*?(um|uh|ah)\s"
        row = 3
    texts = []
    for index, col in df.iterrows():
        text = re.sub(regex, "", col["text"])
        texts.append(text)
    df[command] = texts
    done_message(row, 1, "Cleaned!")


def space_square_bracket():
    global df

    texts = []
    for index, col in df.iterrows():
        # find a sequence of space inside bracket not follow by ']'
        pat1 = re.compile(r"\s+(?=[^\[]*\])")
        # find a sequence of space not follow by a char and delete it out eg spaces between brackets
        pat2 = re.compile(r"\s+(?![a-z])", re.I)
        # case of beginning of line
        pat3 = re.compile(r"^(\[.*?\])\s")

        text = pat2.sub('', pat1.sub('', col['text']))
        text = pat3.sub(r'\1', text)
        texts.append(text)
    df["space_square_brackets"] = texts
    done_message(4,1,"Cleaned!")


def done_message(row, col, text="Done!"):
    t = tkinter.Label(text=text)
    t.grid(row=row, column=col)


def init_gui():
    global FILE_NAME
    global df
    global root
    global header_var, replace_text_var
    root = tkinter.Tk()
    root.title("NGSX Clean Transcript")
    w = 600  # width for the tkinter root
    h = 300  # height for the tkinter root

    ws = root.winfo_screenwidth()  # width of the screen
    hs = root.winfo_screenheight()  # height of the screen
    x = (ws / 4) - (w / 4)
    y = (hs / 4) - (h / 4)
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))

    browse_button = tkinter.Button(root, text='Browse .csv or Excel File', command=import_csv)
    browse_button.grid(sticky="w", row=1, column=1)
    FILE_NAME = tkinter.StringVar()
    file_name_entry = tkinter.Entry(root, textvariable=FILE_NAME, justify=LEFT, width=50)
    file_name_entry.grid(row=1, column=0)
    clear_brac_button = tkinter.Button(root, text='Clean square brackets and content inside', command=lambda: clean_text("clear_square_bracket"), anchor="w")
    clear_brac_button.grid(sticky="e", row=2, column=0)
    um_button = tkinter.Button(root, text='Clean all standalone um ah uh', command=lambda: clean_text("clear_um"), anchor="w")
    um_button.grid(sticky="e", row=3, column=0)
    space_square_button = tkinter.Button(root, text='Clean spaces inside brackets and surrounding space', command=space_square_bracket,anchor="w")
    space_square_button.grid(sticky="e", row=4, column=0)
    save_button = tkinter.Button(root, text='Save as', command=write_excel, anchor="w")
    save_button.grid(sticky="e", row=5, column=0)
    header_var = tkinter.IntVar()
    replace_text_var = tkinter.IntVar()
    header_check = tkinter.Checkbutton(root,text="Header in Excel File", variable=header_var)
    header_check.grid(sticky="w", row=5, column=0, padx=20)
    replace_text_check = tkinter.Checkbutton(root,text="Replace text", variable=replace_text_var)
    replace_text_check.grid(row=5, column=0, padx=(100, 10))
    restart_button = tkinter.Button(root, text='Restart', command=restart)
    restart_button.grid(sticky="e", row=6, column=0)
    close_button = tkinter.Button(root, text='Close', command=root.destroy, anchor="w")
    close_button.grid(sticky="e", row=7, column=0)

    bottom = tkinter.Frame(root)
    bottom_label = tkinter.Label(bottom, text="Feature requests to CMai@clarku.edu")
    bottom_label.pack(side="bottom")
    bottom.grid(pady=y - 30)

    root.mainloop()


def restart():
    root.destroy()
    init_gui()


def popupdialog(message):
    popup = tkinter.Tk()
    popup.wm_title("WARNING")
    label = tkinter.Label(popup, text=message)
    label.pack(side="top", fill="x", pady=10)
    accept_button = tkinter.Button(popup, text="Okay", command=popup.destroy)
    accept_button.pack()
    popup.mainloop()


if __name__ == '__main__':
    init_gui()