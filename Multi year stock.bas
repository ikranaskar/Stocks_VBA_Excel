Attribute VB_Name = "Module1"
Sub Multiyearstock():


Dim Ticker As String

Dim Yearly_Change As Double

Dim Percent_Change As Double

Dim Total_Stock_Volume As Long

Dim Volume As Double

Dim Stock_Open As Double

Dim Stock_Close As Double

Dim Stock_Change As Double

Dim Lastrow As Long





'Loop through worksheet
Dim Worksheet As String
For Each ws In Worksheets
ws.Activate

'Create columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Total_Stock_Volume = 0
Percent_Change = 0
Yearly_Change = 0



Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all rows

For i = 2 To Lastrow

'Ticker name
    TickerName = ws.Cells(i, 1).Value

    Dim Summary_table_Row As Double
    Summary_table_Row = 2




    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    If ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Total = Total + Cells(i, 7).Value
    
    End If
    
    




    If Stock_Open = 0 Then
        Stock_Change = 0
        Percent_Change = 0
    Else
        Stock_Change = Stock_Close - Stock_Open
        Percent_Change = (Stock_Close - Stock_Open) / Stock_Open
    End If


'Condioning formating Green



    If ws.Range("J" & i).Value > 0 Then
        ws.Range("J" & r).Interior.ColorIndex = 4

    ElseIf ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3
        
    End If

Next i
Next ws

End Sub
