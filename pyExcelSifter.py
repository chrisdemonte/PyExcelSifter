import pandas as pd
from tkinter import *
from tkinter import filedialog as fd
from tkinter.ttk import *
import math 

dynamicControls = []
dynamicVars = []
fileName = 0
df = 0
headerList = []

#This code for the Verticle Scroll Frame was taken from this stack overflow post:
# https://stackoverflow.com/questions/40780634/tkinter-canvas-window-size
class VerticalScrolledFrame(Frame):
    """A pure Tkinter scrollable frame that actually works!
    * Use the 'interior' attribute to place widgets inside the scrollable frame.
    * Construct and pack/place/grid normally.
    * This frame only allows vertical scrolling.
    """
    def __init__(self, parent, *args, **kw):
        Frame.__init__(self, parent, *args, **kw)
        
        # Create a canvas object and a vertical scrollbar for scrolling it.
        vscrollbar = Scrollbar(self, orient=VERTICAL)
        vscrollbar.pack(fill=Y, side=RIGHT, expand=FALSE)
        canvas = Canvas(self, bd=0, highlightthickness=0, yscrollcommand=vscrollbar.set,height=400, width=580)
        
        canvas.pack(side=LEFT, fill=BOTH, expand=TRUE)
        
        vscrollbar.config(command=canvas.yview)

        # Reset the view
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)

        # Create a frame inside the canvas which will be scrolled with it.
        self.interior = interior = Frame(canvas, width=580)
        interior_id = canvas.create_window(0, 0, window=interior, anchor=NW)

        # Track changes to the canvas and frame width and sync them,
        # also updating the scrollbar.
        def _configure_interior(event):
            # Update the scrollbars to match the size of the inner frame.
            size = (interior.winfo_reqwidth(), interior.winfo_reqheight())
            canvas.config(scrollregion="0 0 %s %s" % size)
            if interior.winfo_reqwidth() != canvas.winfo_width():
                # Update the canvas's width to fit the inner frame.
                canvas.config(width=interior.winfo_reqwidth())
        interior.bind('<Configure>', _configure_interior)

        def _configure_canvas(event):
            if interior.winfo_reqwidth() != canvas.winfo_width():
                # Update the inner frame's width to fill the canvas.
                canvas.itemconfigure(interior_id, width=canvas.winfo_width())
        canvas.bind('<Configure>', _configure_canvas)

class SampleApp(Tk):
    def __init__(self, *args, **kwargs):
        root = Tk.__init__(self, *args, **kwargs)

        self.geometry("600x400")
        self.title("Excel Sifter")

        img = PhotoImage(width=1, height=1)

        def openFile():
            global fileName
            fileName = fd.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("Comma Separated Values", "*.csv")])

        def createDataFrame():
            global fileName

            global df
            if fileName.endswith(".xlsx"):
                df = pd.read_excel(fileName, )
                return 1
            elif fileName.endswith(".csv"):
                df = pd.read_csv(fileName)
                return 1
            return 0

        def createHeaderList():
            global headerList
            headerList = list(df)

        def populateWindow():
            global headerList
            global dynamicControls
            global dynamicVars

            dynamicControls = []
            dynamicVars = []

            r = 0
            c = 0
            wid = 90
            if len(headerList) > 1325:
                wid = int(75 / math.ceil((len(headerList) / 1325)))
    
            for i in range(len(headerList)):
                
                dynamicVars.append(IntVar(value=0))
                checkbox = Checkbutton(self.frame.interior, text= headerList[i], variable=dynamicVars[i], width=wid)
                checkbox.grid(padx= 10, column= c, row = r, sticky=W)
                dynamicControls.append(checkbox)

                r += 1
                if r > 1325:
                    r = 0
                    c += 1

        def createFilter():
            global headerList
            global dynamicVars
            dfFilter = []
            for i in range(len(headerList)):
                if dynamicVars[i].get():
                    dfFilter.append(headerList[i])
            return dfFilter

        def createFilteredDataFrame(dfFilter):
            global df
            return df.filter(dfFilter, axis = 1)

        def exportFile(newDf, newFile):
            global fileName
            if fileName.endswith(".xlsx"):
                newDf.to_excel(newFile + ".xlsx")
            elif fileName.endswith(".csv"):
                newDf.to_csv(newFile + ".csv")

        def openFileAction():
            openFile()
            if createDataFrame() == 1:
                createHeaderList()
                populateWindow()
        
        def exportFileAction():
            newFile = fd.asksaveasfilename(filetypes=[("Excel Files", "*.xlsx"), ("Comma Separated Values", "*.csv")])
            dfFilter = createFilter()
            newDf = createFilteredDataFrame(dfFilter)
            exportFile(newDf, newFile)

        menubar = Menu(self)
        self.config(menu=menubar)
        fileMenu = Menu(menubar, tearoff= 0)

        menubar.add_cascade(label="File", menu=fileMenu)
        fileMenu.add_command(label="Open", command=openFileAction)
        fileMenu.add_command(label="Export", command=exportFileAction )
        fileMenu.add_separator()
        fileMenu.add_command(label="Exit", command=quit)

        self.frame = VerticalScrolledFrame(root)
        self.frame.pack()



app = SampleApp()
app.mainloop()

