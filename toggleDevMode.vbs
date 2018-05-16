Const ForReading = 1
Const ForWriting = 2
fileArray = array()

cwd = CreateObject("WScript.Shell").CurrentDirectory
Dim FSO, startFolder, currFolder
counter=0

Set FSO = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments

if args.count>0 then
    if args(2)="true" then
        searchVal = "false"
    else
        searchVal= "true"
    end if
    delim = args(1)
    searchVar = args(0)
    searchString = searchVar & delim & searchVal
    destString = searchVar & delim & args(2)
else
    searchString = "devMode: true"
    destString = "devMode: false"
end if

Set startFolder = FSO.GetFolder(cwd)

getSubFolders startFolder

' All of the files ending in JS are now inside of the fileArray variable
For Each oFile in fileArray
    set objFile = FSO.OpenTextFile(oFile,ForReading)
    strText = objfile.ReadAll
    objFile.Close

    strNewText = Replace(strText, searchString, destString)
    set objFile = FSO.OpenTextFile(oFile,ForWriting)
    objFile.Write strNewText
    objFile.Close
    WScript.Echo "Updated: "&oFile
Next
if(UBound(fileArray)<=0) then
    WScript.Echo "No .js files found in this or any subdirectories."
end if
WScript.Echo "Finished"


Sub getSubFolders(Folder)
    For Each Subfolder In Folder.SubFolders
        Set currFolder=FSO.GetFolder(Subfolder.Path)

        For Each objFile In currFolder.Files 
            if FSO.getextensionname(objFile.Path) = "js" then
                'ReDim'ing the array here is a HEAVY operation it's O n^2 so this will get SLOW for a large number of files
                ReDim Preserve fileArray(UBound(fileArray)+1)
                fileArray(UBound(fileArray)) = objFile.Path
            end if
        Next
        getSubFolders Subfolder
    Next

    For Each objFile In Folder.Files 
        if FSO.getextensionname(objFile.Path) = "js" then
            'ReDim'ing the array here is a HEAVY operation it's O n^2 so this will get SLOW for a large number of files
            ReDim Preserve fileArray(UBound(fileArray)+1)
            fileArray(UBound(fileArray)) = objFile.Path
        end if
    Next
End Sub