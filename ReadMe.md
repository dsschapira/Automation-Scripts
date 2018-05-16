These are a number of miscellaneous scripts that I've made to automate tasks.

toggleDevMode
=============
From the current directory will run a VBScript to change a setting in all JS files here and in subdirectories. Usage is:
```
toggleDevMode <var name> <delimiter plus any spaces between delim and value ex: ": "> <desired value>

Example:
toggleDevMode devMode ": " false
```

zipBanners
===========
Requirements:
7-zip is needed and the path to the .exe file needs to be set in zipBanners.vbs 

From the current directory, this will read a manifest.txt file to determine how banner files should be zipped up using 7-zip. Usage is:
```
zipbanners <true or false>

True will include a campaign name in zipped files, False will not
```
The manifest.txt file is absolutely mandatory for this and has a specific format.  An example is below:
```
FinalZipFileName
-SpecificSize-z
-SpecificSize2-z
-Campaign1
--Size1-z
--Size2-z
-Campaign2
--Size1-z
--Size2-z
```
All files marked with a "-z" at the end will be zipped up. The number of dashes at the start of a line determines how many folders deep the zip-script will look and how it is packaged up at the end.  There should be no spaces between these dashes or between the dashes and name of the folders.

Additionally, the folder names there should correspond to the folders used when developing.  This both determines the search-through-folders structure and the output structure.

learndir, workdir, staff_transfer
==========
 Move quickly between some commonly used directories
