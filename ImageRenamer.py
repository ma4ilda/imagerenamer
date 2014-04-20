#! /usr/bin/python
from xlrd import open_workbook, empty_cell, XLRDError
import os, traceback
from Tkinter import *
import tkFont
from tkFileDialog import askopenfilename, askdirectory
import re
import shutil
from scrollbarframes import TasksLabelFrame, ExcelTableFrame

class ImageRenamer(Frame):
    
    def __init__(self, parent, config_helper, log_filename):
        Frame.__init__(self, parent)
        #settings global variables
        self._log_file = config_helper.get_app_dir() + log_filename       
        default_font = tkFont.nametofont("TkDefaultFont")
        self.customFont = default_font
        self.vars = {}
        self.config_helper = config_helper
        self.pack()
        self.create_widgets()
    
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
     
        self.result = Text(outputFrame, wrap=WORD, height = 13, font=self.customFont)
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

        try:
            tasks = self.config_helper.read_config()
            self.tasksFrame.display_tasks(tasks)
        except ValueError as e:
            self.display_message(e.message + " Please check " + self._log_file + " and reload application.", severity = 'error')
 

    def add_block(self, label, parent, row, buttonevent, font, key = None, required = False):
        Label(parent, text = label, font = font).grid(row = row, column = 0, sticky=E, padx=5, pady=2)
        var = StringVar()
        entry = Entry(parent, textvariable = var, width = 45, font=font)
        entry.grid(row = row, column = 2, pady=2)
        button = Button(parent, command = lambda: buttonevent(var, key), text = "...", font=font)
        button.grid(row = row, column = 3, sticky='W', padx=5, pady=2)
        if self.isWindows():
            button.config(width = 4)
            entry.config(width = 50)
        return var
    
    def isWindows(self):
        return sys.platform == 'win32'
    
    def select_file_event(self, stringvar, key):
        stringvar.set(askopenfilename())
        
    def display_message(self, text, clear = False, severity = 'info'):
        if clear:
            self.result.delete(0.0, END)
       
        self.result.insert(END, text + "\n", ( 'error' if severity == 'error' else '')) 
        
    def select_dir_event(self, stringvar, key):
        path = askdirectory()
        stringvar.set(path)
        self.vars[key] =  path

    def start_button_event(self, filename):
        
        if not filename:
            self.display_message( "Data file location was not specified.", True, severity = 'error')
            return
        try:
            reader = ExcelReader(filename, self.config_helper, self.display_message, self.vars)
        except XLRDError as e:
            self.display_message("Unsupported format or corrupted file.", severity = 'error');
            return
           
        try:      
            s = reader.get_sheets()[0]
 
            self.filePreviewFrame.set_excel_sheet(s)
            reader.read_sheet(0)
 
        except Exception:
            self.display_message("Exception occured. Please look at " + self._log_file + " for details.", severity = 'error')
            sys.stderr = open(self._log_file, "a")
            sys.stderr.write(traceback.format_exc())
            sys.stderr.close()

class ExcelReader():
    def __init__(self, path, data_processor, show_message_func = None, var_dict = {}):
        
        if show_message_func:
            self.show_message = show_message_func
        self.show_message("Reading... " + path + "\n", clear = True)  
        self.workbook = open_workbook(path)
        self.vars = var_dict
        self.data_processor = data_processor
    def show_message(self, text, clear = False, severity = 'info'):
        if clear:
            print "\n"
        print severity , ":", text
        
    def get_sheets(self):     
        return self.workbook.sheets()
    
    def read_sheet(self, num):
        try:
            s = self.workbook.sheets()[num]
            if s.nrows > 0:             
                columns = s.row(0)
                for row in range(1, s.nrows):
                    try:
                        self.build_excel_var(self.vars, columns, s.row(row))
                        self.data_processor.execute(self.vars)
                    except ValueError as er:
                        self.show_message("Skipping row number " +  str(row + 1) + ": " + er.message, severity = 'error')
                    except IOError as er:
                        self.show_message("Skipping row number " +  str(row + 1) + ": " + er.filename + " was not found ", severity = 'error')
            self.show_message("Completed.\n")
        except KeyError as er:
            self.show_message("Value for tag <" + er.message + "> was not found." , severity = 'error')
        except OSError as er:
            self.show_message(er.strerror, severity = 'error')

        
    def build_excel_var(self, vars, keys, values):       
        for i in range(len(keys)):
            vars[str(keys[i].value)] = self.sanitize(str(values[i].value))
        return vars

    def sanitize(self, string):
        return re.sub('[^A-Za-z0-9]+', "_", string)
    
class ConfigHelper():
    def __init__(self, dirname, config_separator = "->", filename = "renamer_configfile.txt"):
        self.config_separator = config_separator
        self.dirname = dirname
        self.config_file = filename
        self.default_configuration = '<inPath>/100 thumbnail/<Filename>.jpg -> <outPath>/Brickyard/thumb/<CatalogItem>_thumb.jpg\n'\
        '<inPath>/800 zoom/<Filename>.jpg -> <outPath>/Brickyard/zoom/<CatalogItem>_zoom.jpg\n'\
        '<inPath>/350 main image/<Filename>.jpg -> <outPath>/Brickyard/main/<CatalogItem>_main.jpg'
        self.tasks = []

    def read_config(self):
        path = self.get_app_dir() + self.config_file
    
        if not os.path.isfile(path):
            f = open(path, "w")
            f.write(self.default_configuration)
            f.close()

        f = open(path, "r")
        tmp_list = []
        for line in f:       
            task = line.rstrip("\r\n")
            if len(task) > 0:
                tmp_list.append(task)
                self.add_task(task)
        return tmp_list

    
    def get_app_dir(self):
        path = os.path.join(os.path.expanduser("~"), self.dirname + "/")
        if not os.path.exists(path):
            self.make_dirs(path)
        return path
    
    def make_dirs(self, path):
        dirs = os.path.dirname(path)
        if not os.path.exists(dirs):
            os.makedirs(dirs)
            
    def rename(self, oldpath, newpath):
        self.make_dirs(newpath)
        return shutil.copy(oldpath, newpath)
    
    def execute(self, vars):
        for i in range(len(self.tasks)):

            template = self.tasks[i]
            self.rename(self.compile(template[0], vars), self.compile(template[1], vars))

    def compile(self, template, vars):
        return re.sub(r'<([a-zA-Z0-9_]+)>', lambda matchobj: self.match(matchobj.group(1), vars), template)

    def match(self, key, vars):
        result = vars[key]
        if result:
            return result
        else: raise ValueError( "Value for <" + key + "> is empty.")

    def add_task(self, task):
        if not self.config_separator in task:
            raise ValueError( "Separator '" + self.config_separator + "' is missing in task '" + task + "'.")
        task = re.split(r'\s*->\s*', task)
        self.tasks.append(tuple(task))


if __name__ == '__main__':
    title = "ImageRenamer"
    root = Tk()
    root.title(title)
    root.resizable(0,0)
    #hide window until in will be filled in with widgets
    root.attributes('-alpha', 0.0)
    root.iconify()
    config = ConfigHelper(title)
    log_file = "Renamer_errorlog.txt" 
    try:
        r = ImageRenamer(root, config, log_file)
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
        sys.stderr = open(config.get_app_dir() + log_file, 'a')
        sys.stderr.write(traceback.format_exc())
        sys.stderr.close()