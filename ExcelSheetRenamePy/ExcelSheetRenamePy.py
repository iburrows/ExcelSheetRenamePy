from openpyxl import load_workbook
import os
import tkinter
from tkinter import filedialog

#gets the workbook using a filechooser
def getWorkBook():
    tkinter.Tk().withdraw()
    filepath = filedialog.askopenfilename()
    #save the filename of the file for later
    filename = os.path.basename(filepath)

    wb = load_workbook(filename = filepath)

    setNewName(wb, filename)

def setNewName(wb, filename):
    new_name = input("What should the worksheet be named? ")
    j = input("Starting number: ")
    i = int(j)

    #loop through each sheet to  be able to change it
    for worksheets in wb:
        sheet = worksheets.title
        #we don't want to rename testing overview and master data worksheets
        if sheet != "Testing Overview" and sheet != "Master Data":
            #format will output 0 in front of single digit numbers e.g. 01,02 etc.
            worksheets.title = new_name + "{0:02}".format(i)
            print(worksheets.title)
            i += 1

    wb.save("new_" + filename)

if __name__ == "__main__":
    getWorkBook()