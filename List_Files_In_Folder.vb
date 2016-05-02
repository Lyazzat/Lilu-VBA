Sub ListAllFile()
     
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim ws As Worksheet
     
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set ws = Worksheets.Add
     
     'Get the folder object associated with the directory
    mypath = InputBox("Enter a file path", "Title Here")
    Set objFolder = objFSO.GetFolder(mypath)
    ws.Cells(1, 1).Value = "The files found in folder " & objFolder.Name & " are:"
     
     'Loop through the Files collection
    For Each objFile In objFolder.Files
        ws.Cells(ws.UsedRange.Rows.Count + 1, 1).Value = objFile.Name
        
    Next
     
     'Clean up!
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
     
End Sub
