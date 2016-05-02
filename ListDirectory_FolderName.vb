'The code below retrieves the name and path of the folders and files
'in that directory using the DIR() function and lists them in column A and B:

Sub Example4()

Dim varDirectory As Variant
Dim flag As Boolean
Dim i As Integer
Dim strDirectory As String


strDirectory = InputBox("Enter a file path", "Title Here")   'here change your directory
i = 1
flag = True
varDirectory = Dir(strDirectory, vbDirectory)

While flag = True
    If varDirectory = "" Then
        flag = False
    Else
        Cells(i + 1, 1) = varDirectory
        Cells(i + 1, 2) = strDirectory + varDirectory
        'returns the next file or directory in the path
        varDirectory = Dir
        i = i + 1
    End If
Wend
End Sub
