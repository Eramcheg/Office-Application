import tkinter as tk
import openpyxl
import os
from tkinter import filedialog
class MainClass:
    def __init__(self):
        self.filepath=""
        self.imagePath=""
    def startApp(self):
        root = tk.Tk()
        root.geometry("600x600")
        root.title("Main")
        root.configure(background="black")

        buttonFrame = tk.Frame(root)
        buttonFrame.columnconfigure(0, weight=1)
        buttonFrame.columnconfigure(1, weight=1)
        buttonFrame.pack(pady=200)
        labelExcel = tk.Label(buttonFrame, text=self.filepath)
        labelExcel.grid(row=0, column=0)
        labelImage = tk.Label(buttonFrame, text=self.imagePath)
        labelImage.grid(row=1, column=0)
        but=tk.Button(buttonFrame,text="Excel file", command=lambda:self.openFile(labelExcel))
        but.grid(row=0, column=1)
        but1 = tk.Button(buttonFrame,text="Image file", command=lambda:self.openImage(labelImage))
        but1.grid(row=1, column=1)
        quit=tk.Button(buttonFrame,text="quit", command=self.quitcode)
        quit.grid(row=2,column=1)
        root.mainloop()

    def search_for_file_path(self):
        currdir = os.getcwd()
        tempdir = filedialog.askdirectory( initialdir=currdir, title='Please select a path to excel file')
        if len(tempdir) > 0:
            print("You chose: %s" % tempdir)
        print(tempdir)
    def openFile(self, label):
        tempdir = filedialog.askopenfilename(initialdir="/", title="Select An Excel File", filetypes=(
        ("excel files", "*.xlsx"), ("All files", "*.*")))

        if len(tempdir) > 0:
            self.filepath=tempdir
            label.configure(text=tempdir)
        label.update()
        print(tempdir)
        #fileName=filedialog.askopenfilename(initialdir="/",title="select a file", filetype=(("jpeg", "*.jpg")))

    def openImage(self,label):
        tempdir = filedialog.askopenfilename(initialdir="/", title="Select An Excel File", filetypes=(
            (("jpeg files", "*.jpg"), ("png files", "*.png"))))

        if len(tempdir) > 0:
            self.imagePath = tempdir
            label.configure(text=tempdir)
        label.update()
        print(tempdir)

    def quitcode(self):
        quit()
m=MainClass()
m.startApp()







