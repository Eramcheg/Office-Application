import tkinter as tk
import openpyxl
from PIL import Image
import os
from tkinter import filedialog



class MainClass:
    def __init__(self):
        self.filepath = 'Test.xlsx'
        self.dirImages = []
        self.imagePath = ""

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

        buttonExcel = tk.Button(buttonFrame, text="Excel file", command=lambda: self.openFile(labelExcel))
        buttonExcel.grid(row=0, column=1)

        buttonImage = tk.Button(buttonFrame, text="Image folder", command=lambda: self.openImage(labelImage))
        buttonImage.grid(row=1, column=1)

        buttonSave = tk.Button(buttonFrame, text="Insert", command=self.saveImage)
        buttonSave.grid(row=2, column=1)
        quit = tk.Button(buttonFrame, text="quit", command=self.quitcode)
        quit.grid(row=3, column=1)

        root.mainloop()

    def openFile(self, label):
        tempdir = filedialog.askopenfilename(initialdir="/", title="Select An Excel File", filetypes=(
            ("excel files", "*.xlsx"), ("All files", "*.*")))
        if len(tempdir) > 0:
            self.filepath = tempdir
            label.configure(text=tempdir)
        label.update()
        print(tempdir)

    def openImage(self, label):
        tempdir = filedialog.askdirectory(initialdir="/", title="Select An Image Folder")

        if len(tempdir) > 0:
            self.imagePath = tempdir
            label.configure(text=tempdir)
        label.update()
        self.dirImages = os.listdir(tempdir)
        self.imagePath = tempdir
        print(tempdir)

    def saveImage(self):

        start = 1
        finish = len(self.dirImages)

        for i in range(start, finish + 1):

            workbook = openpyxl.load_workbook('Test.xlsx')  # opening workbook
            worksheet = workbook.get_sheet_by_name("List1") # opening sheet 1
            artical_number = str(worksheet['B' + str(i)].value) #unique number of product

            for j in range(finish):
                if artical_number in self.dirImages[j]:
                    img = Image.open(str(self.imagePath + "/" + self.dirImages[j]))

                    # img1=Image(img)

                    # img=img.resize((400,400))
                    #img.thumbnail((200, 200))

                    img1 = openpyxl.drawing.image.Image(img)

                    img1.width = 200*img1.width/img1.height
                    img1.height = 200
                    worksheet.add_image(img1, "A" + str(i))
                    worksheet.row_dimensions[i].height = int(150)
                    worksheet.column_dimensions["A"].width = 44.6
                    workbook.save(self.filepath)

                    break

        print("Program is successful")

    def quitcode(self):
        quit()


m = MainClass()
m.startApp()
