Option Explicit

' Path to the text file containing the ZIP file names
Const ZIP_FILE_LIST_PATH = "C:\path\to\zipfilelist.txt"

' Pattern to match the files inside the ZIP archives (e.g., "*.txt")
Const FILE_PATTERN = "*.txt"

' Directory to extract the matching files
Const EXTRACT_TO_DIR = "C:\path\to\extracted\files\"

Dim fso, shell, zipFileList, zipFilePath, zipFile, extractedFolder
Dim fileList, zipFileStream, fileName, fileInZip
Dim fileSystem, folder, file

' Create FileSystemObject and Shell.Application
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("Shell.Application")

' Read the ZIP file names from the text file
Set zipFileList = fso.OpenTextFile(ZIP_FILE_LIST_PATH, 1) ' 1 = ForReading

Do Until zipFileList.AtEndOfStream
    zipFilePath = zipFileList.ReadLine
    If fso.FileExists(zipFilePath) Then
        ' Extract files matching the pattern
        ProcessZipFile zipFilePath
    Else
        WScript.Echo "ZIP file not found: " & zipFilePath
    End If
Loop

zipFileList.Close
WScript.Echo "Processing completed."

' Subroutine to process each ZIP file
Sub ProcessZipFile(zipFilePath)
    Dim zipFolder, file, files, extractedFolderPath
    
    ' Open the ZIP file as a folder
    Set zipFolder = shell.NameSpace(zipFilePath)
    
    If Not zipFolder Is Nothing Then
        ' Create the extraction directory if it doesn't exist
        If Not fso.FolderExists(EXTRACT_TO_DIR) Then
            fso.CreateFolder(EXTRACT_TO_DIR)
        End If
        
        ' Iterate over the files in the ZIP folder
        For Each file In zipFolder.Items
            If file.IsFolder Then
                ' Skip folders
                Continue For
            End If
            
            ' Check if the file matches the pattern
            If LCase(file.Name) Like LCase(FILE_PATTERN) Then
                ' Extract the file
                WScript.Echo "Extracting: " & file.Name
                ' Use Shell to copy file to destination folder
                shell.NameSpace(EXTRACT_TO_DIR).CopyHere file
            End If
        Next
    Else
        WScript.Echo "Unable to open ZIP file: " & zipFilePath
    End If
End Sub
