Option Explicit

' Initialize variables
Dim fso, objShell, fileNamesFile, zipFolder, extractFolder, logFile, pattern, logFilePath
Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' Set your parameters here
fileNamesFile = "C:\Users\nimba\OneDrive\Desktop\Learning\VBScirpt\Practice\zipFileList.txt" ' Path to the file containing zip filenames
zipFolder = "C:\Users\nimba\OneDrive\Desktop\Learning\VBScirpt\Practice\logsFolder"        ' Folder containing zip files
extractFolder = "C:\Users\nimba\OneDrive\Desktop\Learning\VBScirpt\Practice\extractFolder" ' Folder to extract files to
logFilePath = "C:\Users\nimba\OneDrive\Desktop\Learning\VBScirpt\Practice\logFile.txt"     ' Path for the log file
pattern = "stat"                               ' Pattern to search within zip files

' Ensure the log directory exists
If Not fso.FolderExists(fso.GetParentFolderName(logFilePath)) Then
    fso.CreateFolder fso.GetParentFolderName(logFilePath)
End If

' Read zip filenames from file
Dim zipFilesToExtract
zipFilesToExtract = ReadFileLines(fileNamesFile)

' Process each zip file
Dim zipFile, extractedCount
For Each zipFile In zipFilesToExtract
    If fso.FileExists(fso.BuildPath(zipFolder, zipFile)) Then
        extractedCount = ExtractAndLog(zipFolder, zipFile, extractFolder, pattern, logFilePath)
    Else
        LogMessage logFilePath, "Zip file not found: " & zipFile
    End If
Next

WScript.Echo "Process completed. Check the log file for details."

' Function to read lines from a file and return as an array
Function ReadFileLines(filePath)
    Dim file, lines, line

    If fso.FileExists(filePath) Then
        Set file = fso.OpenTextFile(filePath, 1)

        Do Until file.AtEndOfStream
            line = Trim(file.ReadLine)

            If line <> "" Then
                If IsEmpty(lines) Then
                    lines = Array(line)
                Else
                    ReDim Preserve lines(UBound(lines) + 1)
                    lines(UBound(lines)) = line
                End If
            End If
        Loop

        file.Close

    Else
        WScript.Echo "File not found: " & filePath
        WScript.Quit
    End If

    If IsEmpty(lines) Then
        lines = Array() ' Initialize an empty array if no lines were found
    End If

    ReadFileLines = lines
End Function

' Function to extract files from a zip and log details
Function ExtractAndLog(zipFolderPath, zipFileName, destFolderPath, filePattern, logFilePath)
    Dim zipFilePath, fileCount, zipNamespace, item
    zipFilePath = fso.BuildPath(zipFolderPath, zipFileName)

    On Error Resume Next
    Set zipNamespace = objShell.Namespace(zipFilePath)
    On Error GoTo 0

    ' Debug: Check if zipNamespace is set correctly
    If zipNamespace Is Nothing Then
        WScript.Echo "Failed to open zip file: " & zipFileName
        LogMessage logFilePath, "Error: Failed to open zip file " & zipFileName
        ExtractAndLog = 0
        Exit Function
    End If

    fileCount = 0
    For Each item In zipNamespace.Items
        If item.IsFolder Then
            fileCount = fileCount + ProcessSubFolder(zipNamespace, item, destFolderPath, filePattern, zipFileName)
        ElseIf InStr(1, item.Name, filePattern, vbTextCompare) > 0 Then
            ' Debug: Check if the file item can be parsed
            On Error Resume Next
            Set fileItem = zipNamespace.ParseName(item.Name)
            If fileItem Is Nothing Then
                LogMessage logFilePath, "Error parsing file: " & item.Name
                WScript.Echo "Error parsing file: " & item.Name
            Else
                ExtractFileFromZip zipFilePath, item.Path, destFolderPath, zipFileName
                fileCount = fileCount + 1
            End If
            On Error GoTo 0
        End If
    Next

    ' Log the result
    LogMessage logFilePath, "Extracted " & fileCount & " files from " & zipFileName
    ExtractAndLog = fileCount
End Function

' Function to extract a single file from a zip
Sub ExtractFileFromZip(zipFilePath, fileName, destFolderPath, zipFileName)
    Dim destFilePath, fileItem
    On Error Resume Next
    Set fileItem = objShell.Namespace(zipFilePath).ParseName(fileName)
    On Error GoTo 0

    If Not fileItem Is Nothing Then
        destFilePath = fso.BuildPath(destFolderPath, zipFileName & "_" & fileItem.Name)
        On Error Resume Next
        fileItem.CopyHere destFilePath
        On Error GoTo 0
    Else
        WScript.Echo "Error extracting file: " & fileName
    End If
End Sub

' Function to process subfolders in a zip file
Function ProcessSubFolder(parentNamespace, zipSubFolder, destFolderPath, filePattern, zipFileName)
    Dim subItem, count, subFolderNamespace
    count = 0

    ' Use the ParentNamespace to navigate through the subfolders
    Set subFolderNamespace = objShell.Namespace(zipSubFolder.Path)
    
    If Not subFolderNamespace Is Nothing Then
        For Each subItem In subFolderNamespace.Items
            If subItem.IsFolder Then
                count = count + ProcessSubFolder(subFolderNamespace, subItem, destFolderPath, filePattern, zipFileName)
            ElseIf InStr(1, subItem.Name, filePattern, vbTextCompare) > 0 Then
                ExtractFileFromZip subFolderNamespace.Self.Path, subItem.Path, destFolderPath, zipFileName
                count = count + 1
            End If
        Next
    End If

    ProcessSubFolder = count
End Function

' Function to log messages to a log file
Sub LogMessage(logFilePath, message)
    Dim logFile

    ' Ensure the log directory exists
    If Not fso.FolderExists(fso.GetParentFolderName(logFilePath)) Then
        fso.CreateFolder fso.GetParentFolderName(logFilePath)
    End If

    ' Attempt to open the log file
    On Error Resume Next
    Set logFile = fso.OpenTextFile(logFilePath, 8, True)

    ' Check if the log file was opened successfully
    If Err.Number <> 0 Then
        WScript.Echo "Error opening log file: " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If

    On Error GoTo 0

    ' Write the message to the log file
    logFile.WriteLine Now & " - " & message
    logFile.Close
End Sub
