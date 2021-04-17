import math
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import scipy as scipy
import  numpy as np
from scipy import stats
import xlrd as df
import self as self
import xlrd as df
import matplotlib.pyplot as plt

import pandas as pd

# initalise the tkinter GUI
root = tk.Tk()
text = tk.StringVar()

root.geometry("1000x500")  # set the root dimensions
root.pack_propagate(False)  # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0)  # makes the root window fixed in size.

# Frame for TreeView
frame1 = tk.LabelFrame(root, text="Excel Data")
frame1.place(height=250, width=500)

required_data = []
required_data1 = []


# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File")
file_frame.place(height=100, width=400, rely=0.65, relx=0)

# Buttons
button1 = tk.Button(file_frame, text="Browse A File", command=lambda: File_dialog())
button1.place(x=300, y=55)

button2 = tk.Button(file_frame, text="Load File", command=lambda: Load_excel_data())
button2.place(x=220, y=55)

button3 = tk.Button(file_frame, text="Karl Method", command=lambda: Load_excel_data1())
button3.place(x=60, y=55)

button4 = tk.Button(file_frame, text="Rank", command=lambda: Load_excel_data2())
button4.place(x=15, y=55)

button5 = tk.Button(file_frame,text = "Quit", command=lambda: root.destroy())
button5.place(x=175, y=55)

v = tk.StringVar(root, "1")
graph1 = tk.Radiobutton(root, text="Scatter Plot", variable=v, value=1, command=lambda: show_graph1(required_data,required_data1))
graph1.place(x=530, y=10)
graph2 = tk.Radiobutton(root, text="Pie of X", variable=v, value=2, command=lambda: show_graph2(required_data,required_data1))
graph2.place(x=610, y=10)
graph2 = tk.Radiobutton(root, text="Pie of Y", variable=v, value=3, command=lambda: show_graph3(required_data,required_data1))
graph2.place(x=690, y=10)
graph2 = tk.Radiobutton(root, text="Line Graph", variable=v, value=4, command=lambda: show_graph4(required_data,required_data1))
graph2.place(x=760, y=10)
graph2 = tk.Radiobutton(root, text="Bar Graph", variable=v, value=5, command=lambda: show_graph5(required_data,required_data1))
graph2.place(x=860, y=10)




# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)

## Treeview Widget
tv1 = ttk.Treeview(frame1)


tv1.place(relheight=1, relwidth=1)  # set the height and width of the widget to 100% of its container (frame1).

treescrolly = tk.Scrollbar(frame1, orient="vertical",
                           command=tv1.yview)  # command means update the yaxis view of the widget
treescrollx = tk.Scrollbar(frame1, orient="horizontal",
                           command=tv1.xview)  # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set,
              yscrollcommand=treescrolly.set)  # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x")  # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y")  # make the scrollbar fill the y axis of the Treeview widget

tv2 = ttk.Treeview()

tv2.place(x=30, y=255)
tv2.place(relheight=0.15, relwidth=0.35)
treescrolly1 = tk.Scrollbar(frame1, orient="vertical",
                            command=tv2.yview)  # command means update the yaxis view of the widget
treescrollx1 = tk.Scrollbar(frame1, orient="horizontal", command=tv2.xview)


def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file["text"] = filename
    return None


def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview"""
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None


    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column)  # let the column heading = column name

    df_rows = df.to_numpy().tolist()  # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row)
        # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
    return None

def Load_excel_data1():
    file_path = label_file["text"]
    excel_filename = r"{}".format(file_path)
    workbook = df.open_workbook(excel_filename)
    sheets = workbook.sheet_names()
    required_data = []
    required_data1 = []
    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data.append((row_valaues[0]))

    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data1.append((row_valaues[1]))
    required_data.pop(0)
    required_data1.pop(0)
    npx = np.array(required_data)
    npy = np.array(required_data1)
    c = str(stats.pearsonr(npx,npy)[0])
    tv2["column"] = c
    tv2["show"] = "headings"
    for column in tv2["columns"]:
        tv2.heading(column, text="Karl Pearson: "+column)


def Load_excel_data2():
    file_path = label_file["text"]
    excel_filename = r"{}".format(file_path)
    workbook = df.open_workbook(excel_filename)
    sheets = workbook.sheet_names()
    required_data = []
    required_data1 = []
    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data.append((row_valaues[0]))

    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data1.append((row_valaues[1]))

    required_data.pop(0)
    required_data1.pop(0)
    npx =np.array(required_data)
    npy = np.array(required_data1)
    print(npx)
    c=stats.spearmanr(npx,npy)[0]
    print(c)
    tv2["column"] = c
    tv2["show"] = "headings"
    for column in tv2["columns"]:
        tv2.heading(column, text="Rank Co-Relation: "+column)






def show_graph1(x,y):
    file_path = label_file["text"]
    excel_filename = r"{}".format(file_path)
    workbook = df.open_workbook(excel_filename)
    sheets = workbook.sheet_names()
    required_data = []
    required_data1 = []
    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data.append((row_valaues[0]))

    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data1.append((row_valaues[1]))

    required_data.pop(0)
    required_data1.pop(0)
    plt.scatter(required_data,required_data1)
    plt.show()


def show_graph2(x,y):
    file_path = label_file["text"]
    excel_filename = r"{}".format(file_path)
    workbook = df.open_workbook(excel_filename)
    sheets = workbook.sheet_names()
    required_data = []
    required_data1 = []
    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data.append((row_valaues[0]))

    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data1.append((row_valaues[1]))

    required_data.pop(0)
    plt.pie(required_data)
    plt.show()

def show_graph3(x,y):
    file_path = label_file["text"]
    excel_filename = r"{}".format(file_path)
    workbook = df.open_workbook(excel_filename)
    sheets = workbook.sheet_names()
    required_data = []
    required_data1 = []
    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data.append((row_valaues[0]))

    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data1.append((row_valaues[1]))

    required_data1.pop(0)
    plt.pie(required_data1)
    plt.show()

def show_graph4(x,y):
    file_path = label_file["text"]
    excel_filename = r"{}".format(file_path)
    workbook = df.open_workbook(excel_filename)
    sheets = workbook.sheet_names()
    required_data = []
    required_data1 = []
    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data.append((row_valaues[0]))

    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data1.append((row_valaues[1]))

    required_data.pop(0)
    required_data1.pop(0)
    plt.plot(required_data,required_data1)
    plt.show()

def show_graph5(x,y):
    file_path = label_file["text"]
    excel_filename = r"{}".format(file_path)
    workbook = df.open_workbook(excel_filename)
    sheets = workbook.sheet_names()
    required_data = []
    required_data1 = []
    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data.append((row_valaues[0]))

    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)

        for rownum in range(sh.nrows):
            row_valaues = sh.row_values(rownum)

            required_data1.append((row_valaues[1]))

    required_data.pop(0)
    required_data1.pop(0)
    plt.bar(required_data,required_data1,color = 'green')
    plt.show()

root.mainloop()
