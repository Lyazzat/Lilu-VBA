Private Sub UppercaseLetters()
    Dim x As Variant
    Dim FinalRow As Integer
    FinalRow = Worksheets("Letters").Cells(Rows.Count, 1).End(xlUp).Row
   ' Loop to cycle through each cell in the specified range.
   For Each x In Worksheets("Letters").Range("A4:A" & FinalRow)
      ' Change the text in the range to uppercase letters.
      x.Value = UCase(x.Value)
   Next
End Sub