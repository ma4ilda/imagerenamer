from Tkinter import *
class AutoScrollbar(Scrollbar):
    # a scrollbar that hides itself if it's not needed.  only
    # works if you use the grid geometry manager.
    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            # grid_remove is currently missing from Tkinter!
            self.tk.call("grid", "remove", self)
        else:
            self.grid()
        Scrollbar.set(self, lo, hi)
    def pack(self, **kw):
        raise TclError, "cannot use pack with this widget"
    def place(self, **kw):
        raise TclError, "cannot use place with this widget"

class ScrollbarLabelFrame(LabelFrame):

    def __init__(self, parent, cnf={}, **kw):
        
        LabelFrame.__init__(self, parent, cnf={}, **kw)
        hscrollbar = AutoScrollbar(self, orient=HORIZONTAL)
        hscrollbar.grid(row=1, column=0, sticky="EW")
        
        self.canvas = Canvas(self,
                        xscrollcommand=hscrollbar.set, width = 580)
        self.canvas.grid(row=0, column=0, sticky='NSEW')
        
        hscrollbar.config(command=self.canvas.xview)
        
        # make the canvas expandable
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        # create canvas contents
        
        self.frame = Frame(self.canvas)
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(1, weight=1)

        self.canvas.create_window(0, 0, anchor=NW, window=self.frame)

    def set_background(self, bg):
        self.frame.config(background = bg)

    def create_interior(self):
        #create interior then update tasks
        self.frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def clear(self):
        for child in self.frame.winfo_children():
            child.destroy()

class TasksLabelFrame(ScrollbarLabelFrame):
    def display_task(self, task, task_num):
        Label(self.frame, text = str(task_num) + '. ' + task).grid(sticky = "W")
        ScrollbarLabelFrame.create_interior(self)
        
class ExcelTableFrame(ScrollbarLabelFrame):
    
    def __init__(self, parent, text, nrows = 4, ncols = 6):
        # use black background so it "peeks through" to 
        # form grid lines
    
        ScrollbarLabelFrame.__init__(self, parent, background="grey", text = text)
        
        self.nrows = nrows
        self.ncols = ncols
        self._sheet = None
        self.set_background("grey")
        self.create_interior(nrows, ncols)
        
    def set_excel_sheet(self, sheet):
        self._sheet = sheet
        self.clear()
        self.create_interior(self.nrows, sheet.ncols if self.ncols < sheet.ncols else self.ncols)
    
    def create_interior(self, nrows, columns):
        for row in range(nrows):
            for column in range(columns):
                text = self.gettext(row, column)
                label = Label(self.frame, text = text,
                                 borderwidth=1)
                if not text:
                    label.config(width =  10)
                label.grid(row=row, column=column, sticky="nsew", padx=1, pady=1)
        
        for column in range(columns):
            self.grid_columnconfigure(column, weight=1)
        ScrollbarLabelFrame.create_interior(self)
        
    def gettext(self, row, column):
        if self._sheet:
            try:
                cell = self._sheet.cell(row,column)
                return cell.value
            except IndexError as er:
                #ignore exception to print empty value in such case
                pass
        return ""