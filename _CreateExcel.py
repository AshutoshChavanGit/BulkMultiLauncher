import os
from tkinter import messagebox as mb
from openpyxl import Workbook

class BookmarkManager:
    def __init__(self):
        self.bookmarks_dir = "Bookmarks"
        self.excel_file = "Links.xlsx"
        self.excel_path = os.path.join(self.bookmarks_dir, self.excel_file)
        self.check_create_bookmarks_dir()
        self.create_excel_file()

    def check_create_bookmarks_dir(self):
        if not os.path.exists(self.bookmarks_dir):
            os.mkdir(self.bookmarks_dir)
            mb.showinfo("created",f"Created directory: {self.bookmarks_dir}")
        else:
            pass
            #mb.showerror("Already exists",f"Directory already exists: {self.bookmarks_dir}")

    def create_excel_file(self):
        if not os.path.exists(self.excel_path):
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Example"
            sheet['A1'] = "www.example.com"
            workbook.save(self.excel_path)
            mb.showinfo("operation successful",f"Created Excel file: {self.excel_path}")
        else:
            selection = mb.askyesno("Already Exists",f"Excel file already exists: {self.excel_path} would you like to create another Excel sheet")
            if selection == True:
                    list = os.listdir(r"Bookmarks/")
                    num = len(list)

                    # Create a new Excel workbook
                    workbook = Workbook()
                    sheet = workbook.active
                    sheet.title = "Example"
                    sheet['A1'] = "www.example.com"
                    # Save the workbook


                    text = self.excel_path
                    new_name = text[:-6]+str(num)+text[-5:]
                    workbook.save(new_name)
                    mb.showinfo("operation successful", f"Created Excel file: {self.excel_path}")



            else:
                mb.showwarning("Clicked No","Sheet not created as you entered no")

