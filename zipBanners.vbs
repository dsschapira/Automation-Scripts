Const zipExe = "C:\Program Files\7-Zip\7z.exe" 'UPDATE 7z.exe PATH HERE IF NEEDED
Const ForReading = 1

Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments

If args.count > 0 Then
    if NOT(vartype(args(0)) = vbBoolean) Then
        useCampaignName = false
    Else
        useCampaignName = args(0)
    End If
End If

cwd = WShell.CurrentDirectory

' 7-zip command line command format is:
' 7z <command> [<switch>...]
' Here, command a = add, switch = -t is type of archive, type of archive is zip
' For the rest of the command we will add the Destination Name and what gets zipped

' This reads from a manifest.txt file in the 
' Working directory which defines the zip format

strCommand = """"&"C:\Program Files\7-Zip\7z.exe"""&" a -tzip"
manifest = cwd&"\manifest.txt"

set manifestFile = FSO.OpenTextFile(manifest,ForReading)
manifestStr = manifestFile.ReadAll
manifestFile.Close

folders = Split(manifestStr,vbCrLf) 'splits the manifest string into an array by splitting on newline character (vbCrLf)
ReDim folderPaths(-1)
ReDim preserve folderPaths(UBound(folders))  'make an array to fill with the paths
ReDim parentDirs(-1)
ReDim preserve parentDirs(UBound(folders))  'Holds the parent path for each level
ReDim foldersToZip(-1)
ReDim preserve foldersToZip(UBound(folders))

'Make the first folder before looping through
FSO.CreateFolder cwd&"\"&folders(0)&"_Zipped"
folderPaths(0) = cwd
zippedFolderParent = cwd&"\"&folders(0)&"_Zipped"
parentDirs(0) = ""
parentDirs(1) = ""
foldersToZip(0) = true
first = true
index = 1 'Start at 1 since we are skipping the first entry
parentDir = ""
prevLevel = 999 'start really high so that it's like we're starting at the first level again
'Loop through to fill our arrays with information about how to structure the zipped up banner files
For Each folder in folders 
    continueFor = true
    If first Then 
        'We want to skip the first folder since this was already done before looping
        continueFor=false
        first = false
    End If
    ' Determine how far down (the level) the folder tree we need to be
    level = 0
    sliceStrIndex = 0
    continueCounting = true
    For i=1 to Len(folder)
        If Mid(folder,i,1) = "-" Then
            If continueCounting Then
                level = level + 1
            End If
        Else
            If continueCounting Then
                sliceStrIndex = i
                continueCounting = false
            End If
        End If
    Next
    ' We now have our folder level
    If continueFor Then
        If Mid(folder,len(folder)-1, len(folder)) = "-z" Then
            'If the folder ends in -z, we will zip it
            foldersToZip(index) = true
            folderStr = Mid(folder,sliceStrIndex,len(folder)-sliceStrIndex-1)
        Else
            foldersToZip(index) = false
            folderStr = Mid(folder,sliceStrIndex,len(folder)-sliceStrIndex+1)
        End If
        parentDirs(level+1) = parentDirs(level)&"\"&folderStr
        folderPaths(index) = parentDirs(level)&"\"&folderStr
        index = index + 1
        prevLevel = level
    End If
Next

'Relative folder paths are now in arrays

For index = 0 to UBound(folderPaths) Step 1
    If index > 0 Then
        If NOT(foldersToZip(index)) Then
            FSO.CreateFolder zippedFolderParent&folderPaths(index)
            If useCampaignName Then
                'If we have this set to true, then we want to include _<campaign name> before the _Zipped
                campaignName = "_"&Mid(folderPaths(index),2,len(folderPaths(index)))
            Else
                campaignName=""
            End If
        Else
            'Zip these files up!
            fileToZip = """"&folderPaths(0)&folderPaths(index)&""""
            zipDest = """"&zippedFolderParent&folderPaths(index)&campaignName&".zip"&""""
            WScript.Echo "Zipping:"&VBTab&fileToZip&vbCrLf&"To:"&VBTab&VBTab&zipDest
            WShell.run strCommand&" "&zipDest&" "&fileToZip
        End If
    End If
Next

'Finally, zip up the main Zipp folder
fileToZip = """"&zippedFolderParent&""""
zipDest =  """"&zippedFolderParent&".zip"&""""
WScript.Echo "Zipping:"&VBTab&fileToZip&vbCrLf&"To:"&VBTab&VBTab&zipDest
WShell.run strCommand&" "&zipDest&" "&fileToZip