Option Explicit
Public FinalRow As Integer


Private Sub ManufacturerCount()
    
    'This function counts number of items per manufacturer.
    ' Entire column E is named 'manu'
    ' Since it is a lop we do not need to use
    ' Range("A3" & FinalRow).FillDown
    
    FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
	Dim I as Integer   'Since we declared "Option Explicit" at the top we have to declare all our variables in each
	'sub-procedure
                        
    For I = 3 To FinalRow
    
          Range("A" & I).Formula = "=COUNTIF(manu,E" & I & ")"
          Cells(I, "A").Value = Cells(I, "A").Value
         
    Next I
    
End Sub

Private Sub DownloadedDate()

    FinalRow = Cells(Rows.Count, 12).End(xlUp).Row  ' Find the last row
	Dim I as Integer 
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
    Dim I as Integer                     
     For I = 3 To FinalRow
    
         Range("M" & I).Formula = "=COUNTIF(dldate,L" & I & ")"
         Cells(I, "M").Value = Cells(I, "M").Value
         
     Next I
    
End Sub
Private Sub GetLetter()
    
    FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
    Dim I as Integer                    
     For I = 3 To FinalRow
    
         Range("N" & I).Formula = "=LEFT(E" & I & ", 1)"
         Cells(I, "N").Value = Cells(I, "N").Value
         
     Next I
End Sub

Private Sub LetterGroupAssign()
    'Figure out what the last row is
    FinalRow = Cells(Rows.Count, 14).End(xlUp).Row
	Dim I as Integer 
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

