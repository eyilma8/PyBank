# PyBank
Sub PyBank()

    Dim I As Integer
    Dim j As Integer
    Dim p As Integer
    Dim TotalNumberofMonth As Integer
    Dim NetProfLoss As Long
    Dim GreatestIncrease As Long
    Dim GreatestDecrease As Long
    Dim AverageChange As Long

TotalNumberofMonth = Application.WorksheetFunction.Count(Range("B2:B87"))
NetProfLoss = Application.WorksheetFunction.Sum(Range("b2:b87"))

R = 3
For p = 2 To 87
    Cells(R, 3) = Cells(p, 2) - Cells(p + 1, 2)
    R = R + 1
    Next p
 
 AverageChange = Application.WorksheetFunction.Average(Cells(2, 3), Cells(87, 3))
 
    
    GreatestIncrease = 0
    For I = 2 To 87
    If (Cells(I + 1, 2) - Cells(I, 2)) > GreatestIncrease Then
    GreatestIncrease = (Cells(I + 1, 2) - Cells(I, 2))
    Else
    End If
Next I

    GreatestDecrease = 0
    For j = 2 To 87
    If (Cells(j + 1, 2) - Cells(j, 2)) < GreatestDecrease Then
    GreatestDecrease = (Cells(j + 1, 2) - Cells(j, 2))
    Else
End If
Next j
   
Range("W10") = TotalNumberofMonth
Range("W11") = NetProfLoss
Range("W12") = AverageChange
Range("W13") = GreatestIncrease
Range("W14") = GreatestDecrease

End Sub

End Sub
