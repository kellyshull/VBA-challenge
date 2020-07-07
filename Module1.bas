Attribute VB_Name = "Module1"
Sub stockdata()
'evaluating everything
Dim ticker As String
Dim I As Variant
Dim last_row As Variant
Dim summary_table_row As Integer
Dim vol As Variant
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim ws As Worksheet

For Each ws In Worksheets
summary_table_row = 2
'overflow avoidance
On Error Resume Next


'evaluate last row
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'headings
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Pecent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
        For I = 2 To last_row
            'conditional
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
            vol = ws.Cells(I, 7).Value
            ticker = ws.Cells(I, 1).Value
            year_open = ws.Cells(I, 3).Value
            year_close = ws.Cells(I, 6).Value
            
            'functions
            yearly_change = year_close - year_open
            percent_change = year_change / year_close
            
            'print
            ws.Cells(summary_table_row, 9).Value = ticker
            ws.Cells(summary_table_row, 10).Value = yearly_change
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 12).Value = vol
            summary_table_row = summary_table_row + 1
            
            vol = 0
            
            End If
            
            Next I
            
    ws.Columns("K").NumberFormat = "0.00%"
    
            
            
            
Next ws

End Sub
