imagerenamer
============

Image Renamer Script

Purpose: The goal of the script is to use the data provided in an Excel spreadsheet to rename product images into readable and SEO-friendly image names, and to separate the contents of image folders into site-specific images.


Short description:

ImageRenamer.py - main functionality

integrationtest.py - integration tests. Recieves variables and spreadsheet location urls through the command line. 
Example:
python integrationtest.py inPath=/Users/ma4ilda/imagerenamer/tests outPath=/Users/ma4ilda/imagerenamer/tests/ /Users/ma4ilda/imagerenamer/tests/BRICKYARD_FILE_LOAD.xls /Users/ma4ilda/imagerenamer/tests/test_wrong_column.xlsx

scroolbarframes.py - contains customized tkinter LabelFrame with attached horizontal Scrollbar

build.bat - bat file uses setup.py for building Win application installation with the cx_Freeze lib.

MacOS folder - contains ImageRenamer's app package for Mac OS X 10.6+
Windows folder - contains .exe file, msi and zip installation folder of ImageRenamer application built.

3. Configuration.

Script configuration is performed by tasks defined in config.txt. This configuration file should be located in ImageRenamer folder of current user:
C:/User/<current user folder>/ImageRenamer/ for Windows OS
/Users/<current user>/ImageRenamer/ for MacOS

If ImageRenamer folder and configuration file doesn’t exist script automatically generates folder and default config file.
Every task defines rules by which paths for copying and renaming files will be generated.

Example1 of config file content. Created on Mac OS:

<inPath>/thumbnail/<Filename>.jpg -> <outPath>/Brickyard/thumb/<CatalogItem>_thumb.jpg

<inPath>/ 800 zoom/<Filename>.jpg -> <outPath>/Brickyard/zoom/<CatalogItem>_zoom.jpg

<inPath>/main image/<Filename>.jpg -> <outPath>/Brickyard/main/<CatalogItem>_main.jpg

This example contains definition of three tasks (in general there could be any number of tasks defined). Each tag corresponds to columns of the data file or variables of script’s “Enter Details:” panel. Tags are case sensitive. 
<inPath> and <outPath>- correspond to InPath entered in the “Enter Details:” panel
<Filename> - corresponds to column in the data file
<CatalogItem> - corresponds to column in the data file
-> - required separator for In and Out paths of the task definition

Variables after user imnput:
InPath = /Users/ma4ilda/Documents/Work/test/master/
outPath = /Users/ma4ilda/Documents/Work/test/master/
Filename = topic_2
CatalogItem = 1/18 Plastic Race Car Dale Jr. 2 (after spaces and special characters “sanitizing”)
Result:
File /Users/ma4ilda/Documents/Work/test/master/thumbnail/topic_2.jpg renamed and copied to /Users/ma4ilda/Documents/Work/test/master/Brickyard/thumb/1_18_Plastic_Race_Car_Dale_Jr_2_thumb.jpg

File /Users/ma4ilda/Documents/Work/test/master/ 800 zoom/topic_2.jpg renamed and copied to /Users/ma4ilda/Documents/Work/test/master/Brickyard/zoom/1_18_Plastic_Race_Car_Dale_Jr_2_zoom.jpg

File /Users/ma4ilda/Documents/Work/test/master/main image/topic_2.jpg renamed and copied to /Users/ma4ilda/Documents/Work/test/master/Brickyard/main/1_18_Plastic_Race_Car_Dale_Jr_2_main.jpg

If folders /Brickyard/thumb/, /Brickyard/main/, /Brickyard/zoom/ don’t exist they will be created.

Example2 of config file content. 

Created on Windows 7.
<inPath>/<type>/<Filename>.jpg -> <outPath>/<Project>/<type>/<CatalogItem>_<type>.jpg

This example contains single task definition.
<inPath> and <outPath>- correspond to InPath entered in the “Enter Details:” panel
<Filename> - corresponds to column in the data file
<CatalogItem> - corresponds to column in the data file
<type> - corresponds to column in the data file
<Project> - corresponds to column in the data file

-> - required separator for In and Out paths of the task definition

Variables:
InPath = C:/Users/IEUser/Downloads/master/
outPath = C:/Users/IEUser/Downloads/static/
Filename = old_image_1
CatalogItem = New_Product_Image_Name1 (after spaces and special characters “sanitizing”)
type = zoom
Project = Indy


Result:
C:/Users/IEUser/Downloads/master/
/zoom/old_image_1.jpg -> C:/Users/IEUser/Downloads/static/Indy/zoom/New_Product_Image_Name1_zoom.jpg

2) Second raw after tasks compilation results in the following paths:
InPath = C:/Users/IEUser/Downloads/master/
outPath = C:/Users/IEUser/Downloads/static/
Filename = old_image_2
CatalogItem = New_Product_Image_Name2 (after spaces and special characters “sanitizing”)
type = zoom
Project = Indy

Result:
C:/Users/IEUser/Downloads/master/
/zoom/old_image_2.jpg -> C:/Users/IEUser/Downloads/static/Indy/zoom/New_Product_Image_Name2_zoom.jpg

3) Third raw after tasks compilation results in the following paths:
InPath = C:/Users/IEUser/Downloads/master/
outPath = C:/Users/IEUser/Downloads/static/
Filename = old_image_3
CatalogItem = New_Product_Image_Name3 (after spaces and special characters “sanitizing”)
type = zoom
Project = Brickyard

Result:
C:/Users/IEUser/Downloads/master/
/zoom/old_image_3.jpg -> C:/Users/IEUser/Downloads/static/Brickyard/zoom/New_Product_Image_Name3_zoom.jpg
