imagerenamer
============

Image Renamer Script

File renamer script based on configuration tasks and exel spreadheet data.

ImageRenamer.py - main functionality

integrationtest.py - integration tests. Recieves variables and spreadsheet location urls through the command line. 
Example:
python integrationtest.py 
inPath=/Users/ma4ilda/imagerenamer/tests 
outPath=/Users/ma4ilda/imagerenamer/tests 
/Users/ma4ilda/imagerenamer/tests/BRICKYARD_FILE_LOAD.xls

scroolbarframes.py - contains customized tkinter LabelFrame with attached horizontal Scrollbar

build.bat - bat file uses setup.py for building Win application installation with the cx_Freeze lib.

MacOS folder - contains ImageRenamer's app package for Mac OS X 10.6+
Windows folder - contains .exe file, msi and zip installation folder of ImageRenamer application built.
