VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_aggregator()

Dim ws As Worksheet

Dim ticker As String
Dim stock_open As Double
Dim stock_close As Double
Dim stock_difference As Double
Dim stock_total As Double
Dim year_first As Double
Dim year_last As Double
Dim percent_change As Double
Dim stock_volume As LongLong




For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    stock_total = 0
    stock_volume = 0
    
    year_first = ws.Cells(2, 3).Value
        
    For i = 2 To ws.UsedRange.Rows.Count
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            
            
        Else

            year_last = ws.Cells(i, 6).Value
            ticker = ws.Cells(i, 1).Value
            
            ws.Cells(ws.Cells(ws.Rows.Count, 9).End(xlUp).Row + 1, 9).Value = ticker
    
    stock_difference = year_last - year_first
            ws.Cells(ws.Cells(ws.Rows.Count, 10).End(xlUp).Row + 1, 10).Value = stock_difference
            
            If stock_difference > 0 Then
                ws.Cells(ws.Cells(ws.Rows.Count, 10).End(xlUp).Row, 10).Interior.ColorIndex = 4
            ElseIf stock_difference < 0 Then
                ws.Cells(ws.Cells(ws.Rows.Count, 10).End(xlUp).Row, 10).Interior.ColorIndex = 3
            End If
                        
            
            If year_first <> 0 Then
                percent_change = ((year_last - year_first) / year_first)
            Else
                percent_change = 0
            End If
            
            ws.Cells(ws.Cells(ws.Rows.Count, 11).End(xlUp).Row + 1, 11).Value = percent_change
            
            
            ws.Cells(ws.Cells(ws.Rows.Count, 12).End(xlUp).Row + 1, 12).Value = stock_volume
            
            stock_volume = 0
            year_first = ws.Cells(i + 1, 3).Value
            
        End If
    Next i
Next ws


End Sub

