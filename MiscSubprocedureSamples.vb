
Private Sub ConvertFormulaToValue()
FinalRow = Cells(Rows.Count, 14).End(xlUp).Row
Dim I as Integer 
    For I = 3 To FinalRow
    Cells(I, 14).Value = Cells(I, 14).Value
   Next I
End Sub


Private Sub Test()
'Dim strFormulas(1 To 2) As Variant
FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
Dim strFormulas as Variant
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
    Dim I as Integer         
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
