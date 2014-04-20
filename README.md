imagerenamer
============

Image Renamer Script

Purpose: The goal of the script is to use the data provided in an Excel spreadsheet to rename product images into readable and SEO-friendly image names, and to separate the contents of image folders into site-specific images.


Short description:

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

3. Configuration.

Script configuration is performed by tasks defined in renamer_configfile.txt. This configuration file should be located in ImageRenamer folder of current user:
C:/User/<current user folder>/ImageRenamer/ for Windows OS
/Users/<current user>/ImageRenamer/ for MacOS
If ImageRenamer folder and configuration file doesn’t exist script automatically generates folder and dumb config file.
Every task defines rules by which paths for copying and renaming files will be generated.
Example1 of config file content. 
Created on Mac OS:
<inPath>/thumbnail/<Filename>.jpg -> <outPath>/Brickyard/thumb/<CatalogItem>_thumb.jpg

<inPath>/ 800 zoom/<Filename>.jpg -> <outPath>/Brickyard/zoom/<CatalogItem>_zoom.jpg

<inPath>/main image/<Filename>.jpg -> <outPath>/Brickyard/main/<CatalogItem>_main.jpg

This example contains definition of three tasks (in general there could be any number of tasks defined). Each tag corresponds to columns of the data file or variables of script’s “Enter Details:” panel. Tags are case sensitive. 
<inPath> and <outPath>- correspond to InPath entered in the “Enter Details:” panel
<Filename> - corresponds to column in the data file
<CatalogItem> - corresponds to column in the data file
-> - required separator for In and Out paths of the task definition
