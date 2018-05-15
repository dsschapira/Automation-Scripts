@echo off

set var=%1
set delim=%2
set val=%3
set SCRIPT_LOCATION="C:\Users\Dan Schapira\Documents\batch_files\toggleDevMode.vbs"
cscript %SCRIPT_LOCATION% %var% %delim% %val%