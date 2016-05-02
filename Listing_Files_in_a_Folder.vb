'Listing files in a folder

Dim fso As Object
Dim strName As String
Dim strArr(1 to 1048576, 1 to 1) As String, i As Long

'Enter the folder name here
Const strDir As String ="C:\Users\Lilu\Documents\VBA\Lilu-VBA"

Let strName=dir$(strDir & "*.xlsx")
Do While strName<> vbNullString
	Let i = i+1 
	Let strArr(i,1) = strDir & strName
	Let strName = Dir$ ()
Loop
Set fso = CreateObject("Scripting.FileSystemObject")
Call recurseSubFolders (fso.GetFolder(strDir), strArr(), 1)
Set fso = Nothing
If i > 0 Then
	Range ("A1").Resize(1).Value = strArr
End If

'Next loop through all found files
'and break into path and filename

FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 1 to FinalRow
	This Entry = Cells (i,1)
	For j = Len(ThisEntry) To 1 Step -1
		If Mid(ThisEntry, j, 1) = Application.PathSeparator Then 
			Cells (i,2) = Left(ThisEntry, j)
			Cells (i, 3) = Mid(ThisEntry, j + 1)
			Exit For
		End If
	Next j
Next i

