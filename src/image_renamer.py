#! /usr/bin/python
from xlrd import open_workbook, empty_cell, XLRDError
import os, traceback
from Tkinter import *
import tkFont
from tkFileDialog import askopenfilename, askdirectory
import re
import shutil

class Renamer(Frame):
    
    def __init__(self, parent):
        Frame.__init__(self, parent)
        #settings global variables
        self._log_file = "Renamer_errorlog.txt"
        self.default_configuration = '<inPath>/thumbnail/<Filename>.jpg -> <outPath>/Brickyard/thumb/<CatalogItem>_thumb.jpg\n'\
        '<inPath>/ 800 zoom/<Filename>.jpg -> <outPath>/Brickyard/zoom/<CatalogItem>_zoom.jpg\n'\
        '<inPath>/main image/<Filename>.jpg -> <outPath>/Brickyard/main/<CatalogItem>_main.jpg'
        self.config_file = "renamer_configfile.txt"
        self.config_separator = "->"
        default_font = tkFont.nametofont("TkDefaultFont")
        self.customFont = default_font
        self.vars = {}
        
        self.pack()
        self.create_widgets()
        try:
            self.tasks = self.read_config(self.config_file)
        except ValueError as e:
            self.display_message(e.message + " Please check " + self.config_file + " and reload application.", error = True)
    
    def read_config(self, config):
        path = os.path.join(os.path.expanduser("~"), "Renamer/" + config)
    
        if not os.path.isfile(path):
            self.make_dirs(path)
            f = open(path, "w")
            f.write(self.default_configuration)
            f.close()
        
        tasks = []
        f = open(path, "r")
        for line in f:       
            task = line.rstrip("\r\n")
            if len(task) > 0:
                self.add_task(tasks, task)
                self.tasksFrame.display_task( task, len(tasks))
        return tasks
    
    def show_in_row(self, panel, row, height):
        panel.grid( row = row, column = 0, sticky='W', \
                 padx=10, pady=5, ipady=5)
        panel.config(font=self.customFont, width = 580, height = height)
        if self.isWindows(): panel.config(width = 500)
        panel.grid_propagate(0)
        
    def create_widgets(self):    
        inputFrame = LabelFrame(self, text="Enter details: ")
        self.show_in_row(inputFrame, 0, height = 140)
        
        self.tasksFrame = TasksLabelFrame(self, text = "Tasks:")
        self.show_in_row(self.tasksFrame, 1, height = 100)
        
        self.filePreviewFrame = ExcelTableFrame(self, text = "Data file preview:")
        self.show_in_row(self.filePreviewFrame, 2, height = 120)  
        
        outputFrame = LabelFrame(self, text="Process results: ")
        self.show_in_row(outputFrame, 3, height = 120)
    
        
        self.result = Text(outputFrame, wrap=WORD, height = 13, font=tkFont.nametofont("TkDefaultFont"))
        if self.isWindows(): self.result.config(width = 80)
        else: self.result.config(width = 69) 
        self.result.pack(side=LEFT)
        self.result.tag_configure('error', foreground='red')
        self.scrollbar = Scrollbar(outputFrame)
        self.scrollbar.pack(side = RIGHT, fill=Y)
        self.result.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.result.yview)
        
        self.add_block("inPath:", inputFrame, 0, self.select_dir_event, self.customFont, 'inPath', required = True)
        self.add_block("outPath:", inputFrame, 1, self.select_dir_event, self.customFont, 'outPath', required = True)
        fileName = self.add_block("Data file*:", inputFrame, 3, self.select_file_event, self.customFont, required = True)

        self.start_button = Button(inputFrame, command = lambda: self.start_button_event(fileName.get()), text = "Start", font=self.customFont)
        if self.isWindows(): self.start_button.config(width = 5)
        self.start_button.grid(row = 4, column = 4, sticky='W', padx=10, pady=2)

    def add_block(self, label, parent, row, buttonevent, font, key = None, required = False):
        Label(parent, text = label, font = font).grid(row = row, column = 0, sticky=E, padx=5, pady=2)
        var = StringVar()
        entry = Entry(parent, textvariable = var, width = 45, font=font)
        entry.grid(row = row, column = 2, pady=2)
        if self.isWindows(): entry.config(width = 50)
        button = Button(parent, command = lambda: buttonevent(var, key), text = "...", font=font)
        button.grid(row = row, column = 3, sticky='W', padx=5, pady=2)
        if self.isWindows(): button.config(width = 4)
        return var
    
    def select_file_event(self, stringvar, key):
        stringvar.set(askopenfilename())
    
    def start_button_event(self, fileName): 
        self.read_spreadsheet(fileName)
        
    def display_message(self, text, delete = False, error = False):
        if delete:
            self.result.delete(0.0, END)
       
        self.result.insert(END, text + "\n", ( 'error' if error else '')) 
        
    def select_dir_event(self, stringvar, key):
        path = askdirectory()
        stringvar.set(path)
        self.vars[key] =  path

    def sanitize(self, string):
        return re.sub('[^A-Za-z0-9]+', "_", string)

    def read_spreadsheet(self, filename):
        
        if not filename:
            self.display_message( "Data file location was not specified.", True, True)
            return

        try:
            sp = open_workbook(filename)
        except XLRDError as e:
            self.display_message("Unsupported format or corrupted file.", error = True);
            return
        self.display_message("Reading... " + filename + "\n", True)        
        try:      
            s = sp.sheets()[0]
 
            self.filePreviewFrame.set_excel_sheet(s)
            
            if s.nrows > 0:             
                columns = s.row(0)
                for row in range(1, s.nrows):
                    
                    self.build_excel_var(self.vars, columns, s.row(row))
                    try:
                        for i in range(len(self.tasks)):
                            
                            #self.display_message("Processing task #" + str(i + 1))
                            template = self.tasks[i]
                            self.execute(template[0], template[1], self.vars)
                    except ValueError as er:
                        self.display_message("Skipping row number " + str(row + 1) + ": " + er.message, error = True)
            self.display_message("\nCompleted.\n")
        except KeyError as er:
            self.display_message("Value for tag <" + er.message + "> was not found." , error = True)
        except IOError as er:
             self.display_message(er.filename + " was not found ", error = True)
        except OSError as er:
            self.display_message(er.strerror, error = True)
        
        except Exception:
            self.display_message("Exception occured. Please look at " + self._log_file + " for details.", error = True)
            sys.stderr = open(self._log_file, 'w')
            sys.stderr.write(traceback.format_exc())
            sys.stderr.close()

    def build_excel_var(self, vars, keys, values):       
        for i in range(len(keys)):
            vars[str(keys[i].value)] = self.sanitize(str(values[i].value))
    
    def add_task(self, tasks, task):
        if not self.config_separator in task:
            raise ValueError( "Separator '" + self.config_separator + "' is missing in task '" + task + "'.")
        task = re.split(r'\s*->\s*', task)
        tasks.append(tuple(task))       
        return tasks

    def make_dirs(self, path):
        dirs = os.path.dirname(path)   
        if not os.path.exists(dirs):
            os.makedirs(dirs)        
    def rename(self, oldpath, newpath):
        self.make_dirs(newpath)
        return shutil.copy(oldpath, newpath)
    
    def execute(self, old, new, vars):
        return self.rename(self.compile(old, vars), self.compile(new, vars))

    def compile(self, template, vars):
        return re.sub(r'<([a-zA-Z0-9_]+)>', lambda matchobj: self.match(matchobj.group(1), vars), template)

    def match(self, key, vars):
        result = vars[key]
        if result:
            return result
        else: raise ValueError( "Value for <" + key + "> is empty.")
    def isWindows(self):
        return sys.platform == 'win32'
    
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
        hscrollbar.grid(row=1, column=0, sticky=E+W)
        
        self.canvas = Canvas(self,
                        xscrollcommand=hscrollbar.set, width = 580)
        self.canvas.grid(row=0, column=0, sticky=N+S+E+W)
        
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

if __name__ == '__main__':
    import sys
    
    root = Tk()
    root.title("ImageRenamer Script")
    root.resizable(0,0)
    #hide window until in will be filled in with widgets
    root.attributes('-alpha', 0.0)
    root.iconify()
    try:
        r = Renamer(root)
        #center window
        root.deiconify()
        # make window always be on top of others
        root.attributes('-topmost', 1)
        root.update()
        root.attributes('-topmost', 0)
        position_x = (root.winfo_screenwidth() - root.winfo_width())/ 2
        position_y = (root.winfo_screenheight() - root.winfo_height())/ 2
        root.geometry("+" + str(position_x) + "+" + str(position_y))       
        root.attributes('-alpha', 1.0) 
        root.mainloop()
    except:
        sys.stderr = open('Reanmer_errorlog_global.txt', 'w')
        sys.stderr.write(traceback.format_exc())
        sys.stderr.close()