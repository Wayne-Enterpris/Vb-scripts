Option Explicit

Dim objFSO, objShell, objFolder, objFile

Dim sourceFolder, extractDir, matchFileNamePattern
sourceFolder = "C:\path\to\your\source\folder"   ' Replace with the path to the folder containing zip files
extractDir = "C:\path\to\extract\dir"             ' Replace with the directory where you want to extract files
matchFileNamePattern = "file*.txt"                 ' Replace with the pattern to match (e.g., "file*.txt")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' Create the target extraction directory if it does not exist
If Not objFSO.FolderExists(extractDir) Then
    objFSO.CreateFolder extractDir
End If

' Get the folder object for the source folder
Set objFolder = objFSO.GetFolder(sourceFolder)

' Iterate through each file in the folder
For Each objFile In objFolder.Files
    ' Check if the file is a zip file
    If LCase(objFSO.GetExtensionName(objFile.Path)) = "zip" Then
        ExtractMatchingFiles objFile.Path, extractDir, matchFileNamePattern
    End If
Next

Sub ExtractMatchingFiles(zipFilePath, extractDir, matchFileNamePattern)
    Dim objArchive, objSource, objTarget, objExtractedFile
    Dim zipFileName, baseFileName, extractFileName
    
    ' Open the zip file
    Set objArchive = objShell.NameSpace(zipFilePath)
    
    ' Extract matching files
    For Each objFile In objArchive.Items
        ' Get the base name of the file (without extension)
        baseFileName = objFSO.GetBaseName(objFile.Path)
        
        ' Check if the current file matches the pattern
        If LCase(Left(baseFileName, Len(matchFileNamePattern) - 5)) = Left(matchFileNamePattern, Len(matchFileNamePattern) - 5) Then
            ' Determine the zip file name without extension
            zipFileName = objFSO.GetBaseName(zipFilePath)
            
            ' Determine the extract file name including the zip file name
            extractFileName = extractDir & "\" & zipFileName & "\" & objFile.Name
            
            ' Check if the target extraction directory exists, create it if it doesn't
            If Not objFSO.FolderExists(extractDir & "\" & zipFileName) Then
                objFSO.CreateFolder extractDir & "\" & zipFileName
            End If
            
            ' Extract the contents of the file to the specified extract file name
            Set objSource = objArchive.Items.Item(objFile.Path)
            Set objTarget = objFSO.CreateTextFile(extractFileName, True)
            
            For Each objExtractedFile in objSource.Items
                objTarget.Write objExtractedFile.Path
            Next
            
            objTarget.Close
        End If
    Next
End Sub
