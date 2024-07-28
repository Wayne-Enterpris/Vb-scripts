Option Explicit

' Define the file name to search for, the path to the ZIP file, and the extraction folder
Dim targetFileName, zipFilePath, extractFolder
targetFileName = "example.txt" ' Change this to the file you're looking for
zipFilePath = "C:\Path\To\Your\Archive.zip" ' Change this to the path of your ZIP file
extractFolder = "C:\Path\To\Extract\Folder" ' Change this to the path of the folder where you want to extract the file

' Call the function to start the search and extraction
ExtractFileFromZip zipFilePath, targetFileName, extractFolder

' Function to search for and extract the file from the ZIP archive
Sub ExtractFileFromZip(zipPath, fileName, outputFolder)
    Dim shellApp, zipFolder, fileFound
    fileFound = False
    
    ' Create the Shell.Application object
    Set shellApp = CreateObject("Shell.Application")
    
    ' Open the ZIP file
    Set zipFolder = shellApp.NameSpace(zipPath)
    
    ' Check if the ZIP file is accessible
    If zipFolder Is Nothing Then
        WScript.Echo "ZIP file not found or inaccessible: " & zipPath
        Exit Sub
    End If
    
    ' Search and extract the file
    fileFound = SearchAndExtractFile(zipFolder, fileName, outputFolder)
    
    If fileFound Then
        WScript.Echo "File extracted successfully: " & fileName
    Else
        WScript.Echo "File not found in ZIP archive: " & fileName
    End If
End Sub

' Recursive function to search for a file in all folders within the ZIP archive and extract it
Function SearchAndExtractFile(folder, fileName, outputFolder)
    Dim file, subFolder
    Dim found

    found = False

    ' Search in the current folder
    For Each file In folder.Items
        If file.IsFolder Then
            ' Recursively search in subfolders
            found = SearchAndExtractFile(file, fileName, outputFolder)
        ElseIf LCase(file.Name) = LCase(fileName) Then
            ' Extract the file if found
            file.CopyHere outputFolder & "\" & file.Name
            found = True
            Exit For
        End If
    Next

    SearchAndExtractFile = found
End Function