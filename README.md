# PyBank
Sub PyBank()

    Dim I As Integer
    Dim j As Integer
    Dim TotalNumberofMonth As Integer
    Dim NetProfLoss As Long
    Dim GreatestIncrease As Long
    Dim GreatestDecrease As Long

TotalNumberofMonth = Application.WorksheetFunction.Count(Range("B2:B87"))
NetProfLoss = Application.WorksheetFunction.Sum(Range("b2:b87"))
    
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
Range("W12") = GreatestIncrease
Range("W13") = GreatestDecrease

End Sub
