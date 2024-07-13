import tkinter as tk

class FAQ:
    def __init__(self,root):
        root = tk.Tk()
        root.title("FAQ Page")
        #root.iconbitmap(default="appicon.ico")
        tk.Label(root, text="(Q)How to use this Tool?").pack(anchor='w')
        tk.Label(root,text="-"*100).pack(anchor='w')
        tk.Label(root, text='Along with this tool you get an Excel file named "Links.xlsx",').pack(anchor='w')
        tk.Label(root,
                 text='In case if you havent got it create one using "Create New File" Button on the initialization page.',).pack(anchor='w')
        tk.Label(root, text='After you have created the excel file : open excel file in Microsoft excel,').pack(anchor='w')
        tk.Label(root, text='Enter all the links you desire to open in bulk in web browser.').pack(anchor='w')
        tk.Label(root, text='caution : Enter them in "A" column of sheet only and one link in each cell.').pack(anchor='w')
        tk.Label(root,
                 text='For Creating profiles, you need to rename the workbook name generally at the bottom of the excel sheet in microsoft Excel.').pack(anchor='w')
        tk.Label(root,
                 text='For launching all the websites in the browser, click the next button below in the initialization page').pack(anchor='w')
        tk.Label(root, text='As you click next select profile in drop down box.').pack(anchor='w')
        tk.Label(root,
                 text='If you want to see links you are about to open click the "refresh" button to see links in your desired profile.').pack(anchor='w')
        tk.Label(root,
                 text='you want to load click refresh and to open it in web-browser click "Open In Browser" button".').pack(anchor='w')
        tk.Button(root,text="Exit",font=("times new Roman",20),command=root.destroy).pack()
        root.mainloop()