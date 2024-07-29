Option Explicit

' Constants
Const ZIP_DIR = "C:\path\to\zip\directory"
Const DEST_DIR = "C:\path\to\destination\directory"
Const LIST_FILE = "C:\path\to\list\file.txt"
Const LOG_FILE = "C:\path\to\log\file.txt"
Const SEARCH_PATTERN = "tem" ' Pattern to search in file names

Dim fso, zipFileList, zipFile, zipName, destPath, zipFileName, logFile
Dim countDict

Set fso = CreateObject("Scripting.FileSystemObject")
Set zipFileList = fso.OpenTextFile(LIST_FILE, 1)
Set countDict = CreateObject("Scripting.Dictionary")

' Process each ZIP file in the list
Do Until zipFileList.AtEndOfStream
    zipFile = zipFileList.ReadLine
    zipName = fso.GetFileName(zipFile)
    
    ' Initialize count for this ZIP file
    countDict(zipName) = 0
    
    ' Process each ZIP file
    ProcessZipFile ZIP_DIR & "\" & zipFile, zipName
Loop

' Close the list file
zipFileList.Close

' Log results
Set logFile = fso.OpenTextFile(LOG_FILE, 2, True)
For Each zipFileName In countDict.Keys
    logFile.WriteLine "File: " & zipFileName & " - Count: " & countDict(zipFileName)
Next
logFile.Close

' Function to process a ZIP file
Sub ProcessZipFile(zipPath, zipName)
    Dim shellApp, zipFolder, item
    Set shellApp = CreateObject("Shell.Application")
    Set zipFolder = shellApp.NameSpace(zipPath)
    
    If Not zipFolder Is Nothing Then
        Dim subFolder
        For Each item In zipFolder.Items
            If item.IsFolder Then
                For Each subFolder In item.SubFolders
                    ProcessSubFolder subFolder, zipFolder, zipName
                Next
            Else
                ProcessFile item, zipFolder, zipName
            End If
        Next
    End If
End Sub

' Function to process a file
Sub ProcessFile(file, zipFolder, zipName)
    Dim fileName, destFileName
    fileName = file.Name
    
    ' Check if the file name contains the search pattern
    If InStr(fileName, SEARCH_PATTERN) > 0 Then
        ' Construct destination path and file name
        destFileName = DEST_DIR & "\" & zipName & "_" & fileName
        
        ' Extract file
        zipFolder.CopyHere file, 4 ' 4 = Do not display progress
        countDict(zipName) = countDict(zipName) + 1
    End If
End Sub

' Function to process a subfolder
Sub ProcessSubFolder(subFolder, zipFolder, zipName)
    Dim file
    For Each file In subFolder.Items
        ProcessFile file, zipFolder, zipName
    Next
End Sub
