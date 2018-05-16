@echo off
::This can be called with truthy or falsy argument
:: truthy = we will add the campaign name to zipped-up files that are inside of campaign folders
:: falsey = we will not add the campaign name
:: default is falsey
:: must have 7zip to use this and update the 7zip path variable in zipBanners.vbs
set includeCampaignName=%1
set SCRIPT_LOCATION="C:\Users\Dan Schapira\Documents\batch_files\zipBanners.vbs"
cscript %SCRIPT_LOCATION% %includeCampaignName%