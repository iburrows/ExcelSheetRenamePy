from openpyxl import load_workbook
import os
import tkinter
from tkinter import filedialog


def getWorkBook():
    tkinter.Tk().withdraw()
    filepath = filedialog.askopenfilename()
    filename = os.path.basename(filepath)

    wb = load_workbook(filename = filepath)

    setNewName(wb, filename)

def setNewName(wb, filename):
    new_name = input("What should the worksheet be named? ")
    j = input("Starting number: ")
    i = int(j)

    for worksheets in wb:
        sheet = worksheets.title
        if sheet != "Testing Overview" and sheet != "Master Data":
            worksheets.title = new_name + "{0:02}".format(i)
            print(worksheets.title)
            i += 1

    wb.save("new_" + filename)

if __name__ == "__main__":
    getWorkBook()