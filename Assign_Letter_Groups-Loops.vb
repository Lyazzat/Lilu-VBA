Private Sub ManufacturerCount()
    
    'This function counts number of items per manufacturer.
    ' Entire column E is named 'manu'
    ' Since it is a lop we do not need to use
    ' Range("A3" & FinalRow).FillDown
    
    FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
                        
    For I = 3 To FinalRow
    
          Range("A" & I).Formula = "=COUNTIF(manu,E" & I & ")"
          Cells(I, "A").Value = Cells(I, "A").Value
         
    Next I
    
End Sub

Private Sub DownloadedDate()

    FinalRow = Cells(Rows.Count, 12).End(xlUp).Row  ' Find the last row

    'Loop to fill function to extract the date and convert into values
    
    For I = 3 To FinalRow
        Cells(I, 12).Formula = "=CONCATENATE(MID(J" & I & ",5,6),"","" , RIGHT(J" & I & ",4))"
        Cells(I, 12).Value = Cells(I, 12).Value
    

    Next I
End Sub
    
Private Sub CountPerDate()
    
    'This function counts number of items per date.
    'Entire column L is named 'dldate'
    
    FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
                        
     For I = 3 To FinalRow
    
         Range("M" & I).Formula = "=COUNTIF(dldate,L" & I & ")"
         Cells(I, "M").Value = Cells(I, "M").Value
         
     Next I
    
End Sub
Private Sub GetLetter()
    
    FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
                        
     For I = 3 To FinalRow
    
         Range("N" & I).Formula = "=LEFT(E" & I & ", 1)"
         Cells(I, "N").Value = Cells(I, "N").Value
         
     Next I
End Sub

Private Sub LetterGroupAssign()
    'Figure out what the last row is
    FinalRow = Cells(Rows.Count, 14).End(xlUp).Row
    'Below is the loop function to assign letter group to technician:
    For I = 3 To FinalRow
        Select Case Cells(I, 14)
        Case 1, 3, "H", "O", "M", "Y", "C"
            Cells(I, 15).Value = "Cindy McComb"
        Case "N", "U", "K"
            Cells(I, 15).Value = "Gwen"
        Case "X", "G", "F", "W", "D", "L", "Z"
            Cells(I, 15).Value = "Matt"
        Case "Q", "V", "S", "I", "R", "E"
            Cells(I, 15).Value = "Leo"
        Case "P", "J", "T", "A", "B"
            Cells(I, 15).Value = "Lina"
        Case Else
            
    End Select
        
    Next I
    
End Sub

Public Sub AssignLetterGroups()

    ManufacturerCount
    DownloadedDate
    CountPerDate
    GetLetter
    LetterGroupAssign

End Sub


Private Sub ConvertFormulaToValue()
FinalRow = Cells(Rows.Count, 14).End(xlUp).Row
    For I = 3 To FinalRow
    Cells(I, 14).Value = Cells(I, 14).Value
   Next I
End Sub


Private Sub Test()
'Dim strFormulas(1 To 2) As Variant
FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
With ThisWorksheet
        strFormulas(1) = "=COUNTIF(12,L)"
        'strFormulas(2) = "=Date"
        'strFormulas(3) = "=A2/B2"

        .Range("A3").Formula = strFormulas
        .Range("A3" & FinalLRow).FillDown
        
    End With
End Sub
    
Private Sub DoesThisWork()
    Dim sht As Worksheet
    Dim LastRow As Long

    Set sht = ThisWorkbook.Worksheets("FromJMSDS")
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    LastRow = sht.ListObjects("Table5").Range.Rows.Count
    For I = 3 To FinalRow
    Cells(I, 1).Formula = "=COUNTIF(manu,'E' & i)" '"=COUNTIF(manu,Table5[[#This Row],[Manufacturer]])"
    
    Next I
    
End Sub

Private Sub Lilu()
    
    FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
            
            For I = 3 To FinalRow
                ' Range("A3").Formula = "=COUNTIF($E$3:$E$" & FinalRow & ", E" & I & ")"
                Range("A3").Formula = "=COUNTIF(manu, E3)"
                 'Range("A3").Formula = "=COUNTIF(manu, E" & I & ")"
                 Range("A3" & FinalRow).FillDown
                ' Cells(I, 1).Value = Cells(I, 1).Value
                'Range ("A3:A182,L3:L182,M3:M182,N3:N182")     if you want to reference non-consequitive ranges)
                'Range("A3" & FinalRow).FillDown
       
           
                ' For i = 3 To FinalRow
       
                   ' Range("L3").Formula = "=CONCATENATE(MID(J" & I & ",5,6),"" , "",RIGHT(J" & I & ",4))"
                   ' Range("L3" & FinalRow &).FillDown
            Next I
End Sub



Private Sub Macro6()
        For Each C In Range("A3:A182")
        C.Value = "=COUNTIF(manu,E"
        '"=COUNTIF(RC[4]:R[179]C[4],RC[4])"
        I = I + 1

    Next
    
End Sub
Private Sub ResizeTable()
'
' Resize table 5
    FinalRow = Cells(Rows.Count, "E").End(xlUp).Row
     'LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row ' For all used range. It won't work for my template
    'because I have formulas already pasted there
   
    ActiveSheet.ListObjects("Table5").Resize Range("$A$2:O" & FinalRow)
    
End Sub

Option Explicit

 'From Overstack
Private Sub FormatDate()
Dim FinalRow As Long
Dim V As Variant
Dim I As Long

    FinalRow = Cells(Rows.Count, 12).End(xlUp).Row
    For I = 3 To FinalRow
        V = Split(Cells(I, "B"))
        V(1) = Month(DateValue(V(1) & " 01")) 'Change month string into number.

        Cells(I, 12) = DateSerial(V(5), V(1), V(2))
        Cells(I, 12).NumberFormat = "mmm dd yyyy"
    Next I

End Sub
