Option Explicit
Public FinalRow As Integer

Private Sub ManufacturerCount()
    
    'This function counts number of items per manufacturer.
    ' Entire column E is named 'manu'
    ' Since it is a lop we do not need to use
    ' Range("A3" & FinalRow).FillDown
    
    FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
         
          Range("A3").Formula = "=COUNTIF(manu,E3)"
          Range("A3" & FinalRow).FillDown
          Range("A3:A" & FinalRow).Value = Range("A3:A" & FinalRow).Value
    
End Sub

Private Sub DownloadedDate()

    FinalRow = Cells(Rows.Count, 12).End(xlUp).Row  ' Find the last row

    'Loop to fill function to extract the date and convert into values
         Range("L3").Formula = "=CONCATENATE(MID(J3,5,6),"","" , RIGHT(J3,4))"
         Range("L3" & FinalRow).FillDown
         Range("L3:L" & FinalRow).Value = Range("L3:L" & FinalRow).Value

End Sub
    
Private Sub CountPerDate()
    
    'This function counts number of items per date.
    'Entire column L is named 'dldate'
    
    FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
           
         Range("M3").Formula = "=COUNTIF(dldate,L3)"
         Range("M3" & FinalRow).FillDown
         Range("M3:M" & FinalRow).Value = Range("M3:M" & FinalRow).Value
      
End Sub
Private Sub GetLetter()
    
    FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
         
         Range("N3").Formula = "=LEFT(E3, 1)"
         Range("N3" & FinalRow).FillDown
         Range("N3:N" & FinalRow).Value = Range("N3:N" & FinalRow).Value
         
   
End Sub

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


Private Sub LetterGroupAssign()
Dim lookupValue As Variant
Dim LetterTech As String
Dim FinalRowJMSDSData As Integer
Dim FinalRowLetters As Integer
Dim I As Integer
Dim LookupArray As Variant

FinalRowJMSDSData = Worksheets("JMSDSData").Cells(Rows.Count, 14).End(xlUp).Row
FinalRowLetters = Worksheets("Letters").Cells(Rows.Count, 1).End(xlUp).Row
Set LookupArray = Worksheets("Letters").Range("A4:B" & FinalRowLetters)
    
     For I = 3 To FinalRowJMSDSData
     lookupValue = Worksheets("JMSDSData").Range("N" & I).Value
     Worksheets("JMSDSData").Cells(I, 15).Value = Application.VLookup(lookupValue, LookupArray, 2, False)
    
Next

End Sub


Public Sub Assign_Letter_Groups()
    MsgBox "WARNING: Do not change name of worksheets: 'JMSDSData' and 'Letters'.               To change letter assignment, please update table in the 'Letters' tabs."
    ManufacturerCount
    DownloadedDate
    CountPerDate
    GetLetter
    UppercaseLetters
    LetterGroupAssign
    MsgBox "SUCCESS! Please check table borders in the JMSDSData tab prior to updating Pivot Tables sheet."
End Sub

