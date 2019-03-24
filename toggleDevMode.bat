@echo off

set var=%1
set delim=%2
set val=%3
set SCRIPT_LOCATION="C:\Users\Dan\Documents\batch_scripts\toggleDevMode.vbs"
cscript %SCRIPT_LOCATION% %var% %delim% %val%