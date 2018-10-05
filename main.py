import os
import sys
import tkinter
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import xlrd as xlrd
import xlwt as xlwt
from PIL import ImageTk, Image


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def parseExcel(filename):
    book = xlrd.open_workbook(filename)
    first_sheet = book.sheet_by_index(0)
    count = 0

    # delete previous file

    try:
        script_dir = os.path.dirname(__file__)  # <-- absolute dir the script is in
        rel_path = "output.xls"
        abs_file_path = script_dir + "/" + rel_path
        os.remove(abs_file_path)
    except:
        print("exception in opening file")

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Released')

    row = 0
    for i in range(0, first_sheet.nrows):

        if str.lower(first_sheet.cell(i, 9).value) == "released":
            count = count + 1
            data = [first_sheet.cell_value(i, col) for col in range(first_sheet.ncols)]
            for index, value in enumerate(data):
                sheet.write(row, index, value)
            row = row + 1

    release(count)

    openCount = 0
    open_row_list = list()
    for i in range(0, first_sheet.nrows):
        dumpStr = str.lower(first_sheet.cell(i, 9).value)
        if dumpStr == "open" or not dumpStr:
            open_row_list.append(i)

    row = 0
    openedrow = 0
    sheet = workbook.add_sheet("Opened-ITALK")
    openedSheet = workbook.add_sheet("Opened-NON ITALK")
    for j in open_row_list:
        dumpStr = str.lower(first_sheet.cell(j, 21).value)
        if not dumpStr:
            openCount = openCount + 1
            data = [first_sheet.cell_value(j, col) for col in range(first_sheet.ncols)]
            inserted = 0
            inserted2 = 0
            for index, value in enumerate(data):
                if str("ITALK").lower() in str(data[3]).lower():
                    sheet.write(row, index, value)
                    inserted = inserted + 1
                else:
                    openedSheet.write(openedrow,index,value)
                    inserted2 = inserted2 + 1
            if inserted > 0:
                row = row + 1
            if inserted2 > 0:
                openedrow = openedrow + 1

    open(openCount)

    retestCount = 0
    sheet = workbook.add_sheet("Retest-ITALK")
    RetestSheet = workbook.add_sheet("Retest-NON ITALK")
    row = 0
    retestRow = 0
    open_row_list = list()
    for i in range(0, first_sheet.nrows):
        dumpStr = str.lower(first_sheet.cell(i, 9).value)
        if dumpStr == "open" and dumpStr:
            open_row_list.append(i)

    for j in open_row_list:
        dumpStr = str.lower(first_sheet.cell(j, 21).value)
        if dumpStr:
            retestCount = retestCount + 1
            data = [first_sheet.cell_value(j, col) for col in range(first_sheet.ncols)]
            inserted = 0
            inserted2 = 0
            for index, value in enumerate(data):
                if str("ITALK").lower() in str(data[3]).lower():
                    sheet.write(row, index, value)
                    inserted = inserted + 1
                else:
                    RetestSheet.write(retestRow, index, value)
                    inserted2 = inserted2 + 1

            if inserted > 0:
                row = row + 1
            if inserted2 > 0:
                retestRow = retestRow + 1

    retest(retestCount)
    workbook.save('./output.xls')
    # open_file_button()


def uploadCallBack():
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename()
    print(filename)
    parseExcel(filename)


def center_window(width=300, height=200):
    # get screen width and height
    screen_width = top.winfo_screenwidth()
    screen_height = top.winfo_screenheight()

    # calculate position x and y coordinates
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    # top.configure(background='white')
    top.geometry('%dx%d+%d+%d' % (width, height, x, y))


def make_button():
    b = tkinter.Button(top, text="Select File", highlightthickness=0, bd=0, command=uploadCallBack)
    rel_path = "excel.png"
    path = resource_path(rel_path)
    image = ImageTk.PhotoImage(Image.open(path))
    b.config(image=image)
    b.image = image
    b.pack()
    b.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)


def logo():
    script_dir = os.path.dirname(__file__)  # <-- absolute dir the script is in
    rel_path = "vf_logo.png"
    abs_file_path = script_dir + "/" + rel_path
    path = resource_path(rel_path)
    image_logo = ImageTk.PhotoImage(Image.open(path))
    imglabel = tkinter.Label(top, image=image_logo)
    imglabel.image = image_logo
    imglabel.pack()
    imglabel.place(relx=0.482, rely=0.168, anchor=tkinter.CENTER)


def title():
    lable = tkinter.Label(top, text="Queue Status Tool")
    lable.config(font=("Courier", 18))
    lable.pack()
    lable.place(relx=0.5, rely=0.265, anchor=tkinter.CENTER)


def release(count=0):
    lable = tkinter.Label(top, text="Release = " + str(count))
    lable.config(font=("Arial", 10))
    lable.pack()
    lable.place(relx=0.1, rely=0.7, anchor=tkinter.CENTER)


def open(count=0):
    lable = tkinter.Label(top, text="Open = " + str(count))
    lable.config(font=("Arial", 10))
    lable.pack()
    lable.place(relx=0.5, rely=0.7, anchor=tkinter.CENTER)


def retest(count=0):
    lable = tkinter.Label(top, text="Retest = " + str(count))
    lable.config(font=("Arial", 10))
    lable.pack()
    lable.place(relx=0.9, rely=0.7, anchor=tkinter.CENTER)


def open_excel():
    os.system('start excel.exe ./output.xls')


# def open_file_button():
#     b = tkinter.Button(top, text="Open File", highlightthickness=0, bd=0, command=open_excel)
#     rel_path = "output_image.png"
#     path = resource_path(rel_path)
#     image = ImageTk.PhotoImage(Image.open(path))
#     b.config(image=image)
#     b.image = image
#     b.pack()
#     b.place(relx=0.5, rely=0.85, anchor=tkinter.CENTER)


top = tkinter.Tk()
top.title("Queue Status Tool")
top.resizable(0, 0)
rel_path = "vf_logo.png"
path = resource_path(rel_path)
image_logo = ImageTk.PhotoImage(Image.open(path))
top.call('wm', 'iconphoto', top._w, image_logo)

center_window(500, 500)

logo()
title()
make_button()
release()
open()
retest()

top.mainloop()
