
Sub ListFilesHyperlinks()
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim i As Integer
Dim MyPath As Variant
Dim FinalRow As Integer

Range("A1:XFD1048576").Clear


'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
MyPath = InputBox("Enter a file path in following format:  C:\Users\Documents")
Set objFolder = objFSO.GetFolder(MyPath)
Cells(1, 1).Value = "The files found in folder " & objFolder.Name & " are:"
Cells(1, 2).Value = "The file paths"
Cells(1, 3).Value = "Hyperlinks"
Range("A1:C1").Font.Bold = True
i = 1
'loops through each file in the directory and prints their names and path
For Each objFile In objFolder.Files
    'print file name
    Cells(i + 1, 1) = objFile.Name
    'print file path
    Cells(i + 1, 2) = objFile.Path
    'Cells(i + 1, 3) = objFile.Path
    i = i + 1
Next objFile
FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
Range("C2").Formula = "=hyperlink(B2)"
Range("C2:C" & FinalRow).FillDown
Columns("A:C").AutoFit


End Sub
