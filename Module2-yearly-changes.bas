Attribute VB_Name = "Module2"
Sub yearly():

'This script tabulates the yearly change for each ticker, and the yearly change %
'This script also colors cells green if yearly change > 0, and red if < 0

'set row 2 values to help loop from row 3 onwards
Cells(2, 12).Value = 262
Cells(2, 13).Value = Cells(262, 6).Value - Cells(2, 3).Value
Cells(2, 14).Value = Cells(2, 13).Value / Cells(2, 3).Value

For i = 3 To 2836

    'yearly change for each ticker
    Cells(i, 13) = Cells(Cells(i, 12).Value, 6).Value - Cells(Cells(i - 1, 15).Value, 3).Value

    'yearly change % for each ticker
    Cells(j, 14) = Cells(j, 13).Value / Cells(Cells(j - 1, 15).Value, 3).Value

    'if yearly change <0, color cells red, otherwise color cells green
    If Cells(j, 13) < 0 Then
    Cells(j, 13).Interior.ColorIndex = 3
    Else
    Cells(j, 13).Interior.ColorIndex = 4
    End If

Next i

End Sub
