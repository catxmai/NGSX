import pandas as pd
# import pywintypes
# import win32api
import re
import tkinter as tk
from tkinter import filedialog
from tkinter.ttk import *
import win32com.client as wincl
import os
import xlrd
import sys

# global FILE_NAME, FILE_PATH, filename_str
# global df
# global root, cda
# global TreeFrame
# global header_var, replace_text_var
# global treeIsClicked, CDAIsClicked


class initRoot(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.treeIsClicked = False
        self.CDAIsClicked = False
        self.title("NGSX Clean Transcript")
        self.root_width = 600  # width for the tkinter root
        self.root_height = 600  # height for the tkinter root

        ws = master.winfo_screenwidth()  # width of the screen
        hs = root.winfo_screenheight()  # height of the screen
        x = (ws / 4) - (self.root_width / 4)  # centering
        y = (hs / 4) - (self.root_height / 4)  # centering
        master.geometry('%dx%d+%d+%d' % (self.root_width, self.root_height, x, y))

    # buttons
    browse_button = tk.Button(root, text='Browse .csv or Excel File', command=import_csv)
    browse_button.grid(sticky="w", row=1, column=1)
    FILE_NAME = tk.StringVar()
    file_name_entry = tk.Entry(root, textvariable=FILE_NAME, justify=tk.LEFT, width=50)
    file_name_entry.grid(row=1, column=0)
    load_file_button = tk.Button(root, text="Load file view", command=lambda: treeview(root), anchor="w")
    load_file_button.grid(sticky="e", row=2, column=0)
    clear_brac_button = tk.Button(root, text='Clean square brackets ',
                                  command=lambda: clean_text("clear_square_bracket"), anchor="w")
    clear_brac_button.grid(sticky="e", row=3, column=0)
    clear_brac_hover = CreateToolTip(clear_brac_button,
                                     text="Eg juice [gesture] apple -> juice apple")
    um_button = tk.Button(root, text='Clean ums', command=lambda: clean_text("clear_um"), anchor="w")
    um_button.grid(sticky="e", row=4, column=0)
    um_hover = CreateToolTip(um_button, text="Clean all um ah uh Eg wood, um, why -> wood,  why")
    space_square_button = tk.Button(root, text='Square bracket space', command=space_square_bracket, anchor="w")
    space_square_button.grid(sticky="e", row=5, column=0)
    space_square_hover = CreateToolTip(space_square_button,
                                       text="Eg penguin [ 2 second pause] -> penguin[2secondpause]")
    save_button = tk.Button(root, text='Save as', command=write_excel, anchor="w")
    save_button.grid(sticky="e", row=6, column=0)
    header_var = tk.IntVar()
    replace_text_var = tk.IntVar()
    header_check = tk.Checkbutton(root, text="Header in Excel File", variable=header_var,
                                  command=lambda: treeview(root, save_mode=True))
    header_check.grid(sticky="w", row=6, column=0, padx=20)
    header_hover = CreateToolTip(header_check, text="Check if you want header in your saved Excel file")
    replace_text_check = tk.Checkbutton(root, text="Replace text", variable=replace_text_var)
    replace_text_check.grid(row=6, column=0, padx=(100, 10))
    cda_button = tk.Button(root, text="Run this in CDA", command=init_CDA, anchor="w")
    cda_button.grid(sticky="e", row=7, column=0)
    restart_button = tk.Button(root, text='Restart', command=restart)
    restart_button.grid(sticky="e", row=8, column=0)
    close_button = tk.Button(root, text='Close', command=root.destroy, anchor="w")
    close_button.grid(sticky="e", row=9, column=0)

    bottom_label = tk.Label(root, text="Feature requests to CMai@clarku.edu")
    bottom_label.place(relx=.5, rely=1, anchor="s")


    def import_csv():
        global df
        global FILE_NAME, filename_str
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
        filename_str = name
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
        for index, row in df.iterrows():
            if re.search(r"ï»¿", str(row["timestamp"])):  # not sure how to fix encoding utf-8 here
                text = re.sub(r"ï»¿", "", str(row["timestamp"]))
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
        row_id = 0

        if command == "clear_square_bracket":
            regex = "(\[.*?\])"
            row_id = 3
        if command == "clear_um":
            # regex = "\s(um)\s*?|\s*?(um)\s|\s(uh)\s*?|\s*?(uh)\s|\s(ah)\s*?|\s*?(ah)\s"
            regex = "\s(,?um|,?uh|,?ah)(?!\w)\s*?|\s*?(?<!\w)(um,?|uh,?|ah,?)\s"
            row_id = 4
        texts = []
        for index,row in df.iterrows():
            text = re.sub(regex, "", row["text"])
            texts.append(text)
        df[command] = texts
        done_message(row_id, 1, "Cleaned!")
        treeview(root)


    def space_square_bracket():
        global df

        texts = []
        for index,row in df.iterrows():
            # find a sequence of space inside bracket not follow by ']'
            pat1 = re.compile(r"\s+(?=[^\[]*\])")
            # find a sequence of space not follow by a char and delete it out eg spaces between brackets
            pat2 = re.compile(r"\s+(?![a-z])", re.I)
            # case of beginning of line
            pat3 = re.compile(r"^(\[.*?\])\s")

            text = pat2.sub('', pat1.sub('', row['text']))
            text = pat3.sub(r'\1', text)
            texts.append(text)
        df["space_square_brackets"] = texts
        done_message(5,1,"Cleaned!")
        treeview(root)


    def done_message(row, col, text="Done!"):
        t = tk.Label(text=text)
        t.grid(row=row, column=col)


    def treeview(root, save_mode=False):
        global df
        global TreeFrame
        global treeIsClicked

        if treeIsClicked:
            TreeFrame.destroy()
        treeIsClicked = True

        df_columns = list(df.columns)
        tree_columns = df_columns[1:len(df.columns)]
        TreeFrame = tk.Frame(root)
        header = bool(header_var.get())
        if save_mode and not header:
            TreeFrame.destroy()
            TreeFrame = tk.Frame(root)
            tree = tk.ttk.Treeview(TreeFrame, columns=tree_columns, height=16, show="tree")
        elif save_mode and header:
            tree = tk.ttk.Treeview(TreeFrame, columns=tree_columns, height=16)
        elif not save_mode:
            tree = tk.ttk.Treeview(TreeFrame, columns=tree_columns, height=16)
            header_var.set(value=1)
        # scrollbar
        yscroll = tk.ttk.Scrollbar(TreeFrame, orient="vertical", command=tree.yview)
        yscroll.pack(side='right',fill="y")
        xscroll = tk.ttk.Scrollbar(TreeFrame, orient="horizontal", command=tree.xview)
        xscroll.pack(side='bottom',fill="x")
        tree.configure(xscrollcommand=xscroll.set,yscrollcommand=yscroll.set)

        # Treeview oddity with not first value ie timestamp
        #columns
        tree.column("#0", width=200,minwidth=100,stretch=True)
        tree.column("#1", width=100,minwidth=100,stretch=True)
        for i in range(2,len(tree_columns)+1): # timestamp and name is incl
            tree.column("#"+str(i), width=200,minwidth=300,stretch=True)

        #headers
        tree.heading("#0", text="timestamp", anchor="w")
        tree.heading("#1", text="name", anchor="w")
        for i in range(2, len(tree_columns) + 1):  # timestamp and name is incl
            tree.column("#" + str(i), width=200, minwidth=300, stretch=True)
            tree.heading("#" + str(i), text=df_columns[i], anchor="w")

        for index,row in df.iterrows():
            value_list=[]
            for heading in tree_columns:
                value_list.append(row[heading])
            tree.insert("",'end', text=row["timestamp"],values=value_list)

        tree.pack(fill=tk.BOTH, side="left")
        TreeFrame.grid(columnspan=10,rowspan=10,sticky="nsew")
        # TreeFrame.rowconfigure(index=1, minsize=200, weight=1)
        # update root size
        # root.update()
        width = root.winfo_screenwidth()-300
        height = root.winfo_screenheight()-200
        root.geometry(f'{width}x{height}')


    def init_gui():
        global FILE_NAME, FILE_PATH
        global df
        global root
        global header_var, replace_text_var
        global treeIsClicked, CDAIsClicked


        root.mainloop()


def init_CDA():
    global df, cda
    global filename_str, CDAIsClicked
    CDAIsClicked = True
    cda = tk.Tk()
    cda.wm_title("NGSX Clean Transcript")
    cda_width = 600  # width for the tkinter root
    cda_height = 600  # height for the tkinter root

    ws = root.winfo_screenwidth()  # width of the screen
    hs = root.winfo_screenheight()  # height of the screen
    x = (ws / 4) - (cda_width / 4)  # centering
    y = (hs / 4) - (cda_height / 4)  # centering

    try:
        a = filename_str
    except NameError:
        cda.geometry('+400+300')
        label1 = Label(cda, text="CDA: No file selected")
        label1.pack(side="top",pady=10)
        accept_button = tk.Button(cda, text="Okay", command=cda.destroy)
        accept_button.pack()
        return
    cda.geometry('+%d+%d' % (x, y))
    cda.geometry('%dx%d'%(cda_width,cda_height))
    file_label = tk.Label(cda, text=filename_str)
    file_label.pack(side="top", fill="x", pady=10)
    cda.focus_force()
    popupdialog("Before you do this, make sure you've enabled macro in your Excel so you can run CDA. To do this,"
                "\n open any Excel file, go to File->Options->Trust Center->Settings->Macros->Enable all macros.\n"
                "You can disable this later if you want.")


    cda.mainloop()


def run_CDA():
    excel_macro = wincl.DispatchEx("Excel.application")
    excel_path = os.path.expanduser("C:\\Users\\Cat Mai\\Documents\\Work\\NGSX\\Coding\\NGSX\\regex_cleaning\\CDA_Sample.xlsm")
    workbook = excel_macro.Workbooks.Open(Filename = excel_path)
    excel_macro.Application.Run("GenerateData0")
    workbook.SaveAs("C:\\Users\\Cat Mai\\Documents\\Work\\NGSX\\Coding\\NGSX\\regex_cleaning\\CDA_Sampleresult_1.xlsm")
    excel_macro.Application.Quit()
    del excel_macro


def restart():
    if CDAIsClicked:
        cda.destroy()
    root.destroy()
    init_gui()


def popupdialog(message):
    popup = tk.Tk()
    popup.wm_title("WARNING")
    popup.geometry("+200+200")
    label = tk.Label(popup, text=message)
    label.pack(side="top", fill="x", pady=10)
    accept_button = tk.Button(popup, text="Okay", command=popup.destroy)
    accept_button.pack()
    popup.focus_force()
    popup.mainloop()
    


class CreateToolTip(object):
    # create a tooltip for a given widget
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.close)
    def enter(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = tk.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw, text=self.text, justify='left',
                       background='#75b7e0', relief='solid', borderwidth=1,
                       font=("times", "8", "normal"))
        label.pack(ipadx=1)
    def close(self, event=None):
        if self.tw:
            self.tw.destroy()


if __name__ == '__main__':
    init_gui()