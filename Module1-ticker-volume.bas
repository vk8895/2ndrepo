Attribute VB_Name = "Module1"
Sub VBA_wall_st():

'Note: There are two scripts for completing the HW assignment b/c this script was taking a very long time to run
'This script, VBA_wall_st identifies all the unique stock tickers traded each and totals their volume traded
'The other script in Module 2 computes all the required yearly changes,%, and cell coloring

'Declare all variables
For Each ws In Worksheets

Dim worksheetame As String
worksheetname = ws.Name

Dim stockname As String
Dim ticker_total_volume As Double

Dim tablerow As Double
tablerow = 2

'in order to determine yearly changes, count row# for easier reference in the other script
Dim counter As Double
counter = 1

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'main loop to identify tickers and total share volume traded

For i = 2 To lastrow

    'if cells rows are not equal in column A, then
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    'make array/list of all unique stocknames/tickers
    stockname = Cells(i, 1).Value
    
    'display array in Column I
    ws.Range("I" & tablerow).Value = stockname
    tablerow = tablerow + 1
    
    'if cell rows are equal in column A, then
    Else
    
    'sumup total of all shares traded for each stock/ticker in the year
    ticker_total_volume = ticker_total_volume + ws.Cells(i + 1, 7).Value
    
    'display array in Column J
    ws.Range("J" & tablerow).Value = ticker_total_volume
    
    '"Counter" counts the #trading days in a year for a particular stock/ticker
    counter = counter + 1
    ws.Range("K" & tablerow).Value = counter
    ws.Range("L" & tablerow).Value = counter + 1
                        
    End If

Next i

Next ws

End Sub
