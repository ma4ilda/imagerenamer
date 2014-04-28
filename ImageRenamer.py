#! /usr/bin/python
from xlrd import open_workbook, empty_cell, XLRDError
import os, traceback
from Tkinter import *
import tkFont
from tkFileDialog import askopenfilename, askdirectory
import re
import shutil
from scrollbarframes import TasksLabelFrame, ExcelTableFrame

PROJECT = "ImageRenamer"

CONFIG = """<inPath>/100 thumbnail/<Filename>.jpg -> <outPath>/Brickyard/thumb/<CatalogItem>_thumb.jpg
<inPath>/800 zoom/<Filename>.jpg -> <outPath>/Brickyard/zoom/<CatalogItem>_zoom.jpg
<inPath>/350 main image/<Filename>.jpg -> <outPath>/Brickyard/main/<CatalogItem>_main.jpg"""

class Worker(object):
    tasks = []

    def __init__(self, config_file="config.txt", log_file="errorlog.txt"):
        self.work_dir = PROJECT
        self.config_file = config_file
        self.log_file = log_file

    def read_config(self):
        path = self.get_config()
        if not os.path.isfile(path):
            with open(path, "w") as f:
                f.write(CONFIG)
            config = CONFIG.split("\n")
        else:
            with open(path, "r") as f:
                config = f.readlines()

        config = [line.rstrip("\r\n") for line in config]
        for task in config:
            if len(task) > 0:
                self.add_task(task)
        return config

    def get_work_dir(self):
        path = os.path.join(os.path.expanduser("~"), self.work_dir)
        if not os.path.exists(path):
            self.make_dirs(path)
        return path

    def get_config(self):
        return os.path.join(self.get_work_dir(), self.config_file)

    def get_log(self):
        return os.path.join(self.get_work_dir(), self.log_file)
    
    def make_dirs(self, path):
        dirs = os.path.dirname(path)
        if not os.path.exists(dirs):
            os.makedirs(dirs)

    def rename(self, oldpath, newpath):
        self.make_dirs(newpath)
        return shutil.copy(oldpath, newpath)
    
    def execute(self, variables):
        for task in self.tasks:
            self.rename(self.compile(task[0], variables), self.compile(task[1], variables))

    def compile(self, template, variables):
        return re.sub(r'<([a-zA-Z0-9_]+)>', lambda matchobj:  self.match(matchobj.group(1), variables), template)

    def match(self, key, vars):
        result = vars[key]
        #check empty cells only for keys existed in task configuration
        if result:
            return result
        else:
            raise ValueError( "Value for <" + key + "> is empty.")
        
    def add_task(self, task):
        try:
            s = re.split(r'\s*->\s*', task)
            self.tasks.append((s[0], s[1]))
        except:
            raise ValueError("Cannot parse task '" + task + "'.")

class GUI(Worker, Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        Worker.__init__(self)
        self.customFont = tkFont.nametofont("TkDefaultFont")
        self.variables = {}
        self.pack()
        self.create_widgets()
    
    def show_in_row(self, panel, row, height):
        panel.grid( row = row, column = 0, sticky='W', padx=10, pady=5, ipady=5)
        panel.config(font=self.customFont, width = 580, height = height)
        if self.is_windows():
            panel.config(width = 500)
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
     
        self.result = Text(outputFrame, wrap=WORD, height = 13, font=self.customFont)
        if self.is_windows():
            self.result.config(width = 80)
        else:
            self.result.config(width = 69) 
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
        if self.is_windows():
            self.start_button.config(width = 5)
        self.start_button.grid(row = 4, column = 4, sticky='W', padx=10, pady=2)

        try:
            tasks = self.read_config()
            self.tasksFrame.display_tasks(tasks)
        except ValueError as e:
            self.message(e.message + " Please check " + self.get_config() + " and reload application.")

    def add_block(self, label, parent, row, buttonevent, font, key = None, required = False):
        Label(parent, text = label, font = font).grid(row = row, column = 0, sticky=E, padx=5, pady=2)
        var = StringVar()
        entry = Entry(parent, textvariable = var, width = 45, font=font)
        entry.grid(row = row, column = 2, pady=2)
        button = Button(parent, command = lambda: buttonevent(var, key), text = "...", font=font)
        button.grid(row = row, column = 3, sticky='W', padx=5, pady=2)
        if self.is_windows():
            button.config(width = 4)
            entry.config(width = 50)
        return var
    
    def is_windows(self):
        return sys.platform == 'win32'
    
    def select_file_event(self, stringvar, key):
        stringvar.set(askopenfilename())
        
    def message(self, text, clear = False, severity = 'error'):
        if clear:
            self.result.delete(0.0, END)
        self.result.insert(END, text + "\n", ('error' if severity == 'error' else '')) 
        
    def select_dir_event(self, stringvar, key):
        path = askdirectory()
        stringvar.set(path)
        self.variables[key] = path

    def start_button_event(self, filename):
        if not filename:
            self.message("Data file location was not specified.", True)
            return
        try:
            reader = ExcelReader(filename, self.message)
            s = reader.get_sheets()[0]
            self.filePreviewFrame.set_excel_sheet(s)
            reader.run(0, self.execute, self.variables.copy())
        except XLRDError as e:
            self.message("Unsupported format or corrupted file.")
        except Exception:
            self.message("Exception occured. Please look at " + self.get_log() + " for details.")
            sys.stderr = open(self.get_log(), "a")
            sys.stderr.write(traceback.format_exc())
            sys.stderr.close()

class ExcelReader(object):
    def __init__(self, path, message=None):
        if message:
            self.message = message
        self.message("Reading... " + path + "\n", clear = True, severity = 'info')  
        self.workbook = open_workbook(path)

    def message(self, text, clear=False, severity="error"):
        if clear:
            print "\n"
        print severity, ":", text
        
    def get_sheets(self):     
        return self.workbook.sheets()
    
    def run(self, num, execute, variables={}):
        print variables
        try:
            s = self.get_sheets()[num]
            if s.nrows > 0:             
                columns = s.row(0)
                for row in range(1, s.nrows):
                    try:
                        variables = self.build_excel_var(variables, columns, s.row(row))
                        execute(variables)
                    except IOError as e:
                        self.message("Skipping row number " +  str(row + 1) + ": " + e.strerror + " " + e.filename)
                    except ValueError as e:
                        self.message("Skipping row number " +  str(row + 1) + ": " + e.message)
            self.message("Completed.\n", severity="info")
        except KeyError as er:
            self.message("Value for tag <" + er.message + "> was not found.")
        except OSError as er:
            self.message(er.strerror)
        
    def build_excel_var(self, variables, keys, values):       
        for i in range(len(keys)):
            variables[str(keys[i].value)] = self.sanitize(str(values[i].value))
        return variables

    def sanitize(self, string):
        return re.sub('[^A-Za-z0-9]+', "_", string)

if __name__ == '__main__':
    root = Tk()
    root.title(PROJECT)
    root.resizable(0,0)
    #hide window until in will be filled in with widgets
    root.attributes('-alpha', 0.0)
    root.iconify()
    try:
        worker = GUI(root)
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
        sys.stderr = open(Worker().get_log(), 'a')
        sys.stderr.write(traceback.format_exc())
        sys.stderr.close()