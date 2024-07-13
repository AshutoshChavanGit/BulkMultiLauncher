from PIL import Image,ImageTk
import tkinter as tk
import _MainGui as mg
import _CreateExcel as CE
import _FAQ as faq


class Initialize:
    def __init__(self):
        self.root = tk.Tk()
        self.root.iconbitmap(default="images/url.ico")
        self.root.title("BulkOpener")
        self.InfoString = tk.StringVar()
        self.TitleFrame = tk.Frame(self.root)
        self.TitleFrame.pack()
        #image
        self.image = Image.open(fp="images/url-icon.png")
        tk_image = ImageTk.PhotoImage(self.image,size=(20,20))
        tk.Label(self.TitleFrame, image=tk_image).pack(side=tk.LEFT)
        tk.Label(self.TitleFrame, text="  BulkOpener\n"+"-"*50, font=("Times New Roman", 80)).pack(side=tk.LEFT)

        def OnClick():
            obj = CE.BookmarkManager()

        def onClick():
            self.root.destroy()
            obj = mg.MainGui()


        def center_window(window, width, height):
            # Get screen width and height
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()

            # Calculate position x and y
            x = (screen_width // 2) - (width // 2)
            y = (screen_height // 2) - (height// 2)

            # Set the window position
            window.geometry(f'{width}x{height}+{x}+{y}')



        # Set window size
        window_width = 1500
        window_height = 800
        center_window(self.root, window_width, window_height)


        string = "This application is used to automatically run all of your desired links in your default browser.\n This tool is used to automatically launch your websites automatically.\n"+"-"*50
        self.InfoString.set(string)
        self.frame_content = tk.Frame(self.root).pack()#.grid_configure(row=1,column=0)
        tk.Label(self.frame_content,textvariable=self.InfoString,font=("times new Roman",29)).pack()#.grid_configure(row=0,column=0)
        tk.Label(self.frame_content,text="Used for initialization purposes only: skip if you already have excel sheet: Links.xlsx",font=("times new Roman",29)).pack(anchor="w")
        self.Frame1 = tk.Frame(self.root)
        self.Frame1.pack()
        self.Button_frame = tk.Frame(self.frame_content,width=50,height=50).pack()
        tk.Button(self.Frame1,text="Create New File",command=OnClick,font=("times new Roman",20)).pack(anchor="w",side=tk.LEFT)
        tk.Button(self.Frame1,text="(Q)How to use this tool",command= lambda : faq.FAQ(self.root),font=("times new Roman",20)).pack(anchor='e',side=tk.LEFT)

        self.Frame2 = tk.Frame(self.root)
        self.Frame2.pack()

        tk.Button(self.root,text="Next >>",command=onClick,font=("times new Roman",20)).pack()

        self.root.mainloop()

