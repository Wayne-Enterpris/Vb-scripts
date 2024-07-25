Option Explicit

Dim objFSO, objFolder, objFile, outFile
Dim folderPath, outputFilePath

folderPath = "C:\path\to\your\folder"   ' Replace with the path to your folder
outputFilePath = "C:\path\to\output\file.txt"   ' Replace with the path where you want to save the output file

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check if the folder exists
If objFSO.FolderExists(folderPath) Then
    ' Open the folder
    Set objFolder = objFSO.GetFolder(folderPath)
    
    ' Create or overwrite the output file
    Set outFile = objFSO.CreateTextFile(outputFilePath, True)
    
    ' Iterate through each file in the folder and write its name to the output file
    For Each objFile In objFolder.Files
        outFile.WriteLine objFile.Name
    Next
    
    ' Close the output file
    outFile.Close
    
    MsgBox "File names have been saved to: " & outputFilePath, vbInformation, "File Names Saved"
Else
    MsgBox "Folder does not exist: " & folderPath, vbExclamation, "Folder Not Found"
End If
