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
year_open = ws.Cells(2, 3).Value
'overflow avoidance
On Error Resume Next


'evaluate last row
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'headings
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Pecent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    ws.Cells(4, 16).Value = 0
    ws.Cells(2, 16).Value = -1000
    ws.Cells(3, 16).Value = 1000
    
        For I = 2 To last_row
            'conditional
            vol = ws.Cells(I, 7).Value + vol
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
            
            ticker = ws.Cells(I, 1).Value
            
            year_close = ws.Cells(I, 6).Value
            
            'functions
            yearly_change = year_close - year_open
            percent_change = yearly_change / year_open
            year_open = ws.Cells(I + 1, 3).Value
            'print
            ws.Cells(summary_table_row, 9).Value = ticker
            ws.Cells(summary_table_row, 10).Value = yearly_change
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 12).Value = vol
            
            If ws.Cells(4, 16).Value < vol Then
            
            ws.Cells(4, 16).Value = vol
            ws.Cells(4, 15).Value = ticker
            End If
    
            
            If ws.Cells(2, 16).Value < percent_change Then
            
    
            ws.Cells(2, 16).Value = percent_change
            ws.Cells(2, 15).Value = ticker
            
            End If
            
            If ws.Cells(3, 16).Value > percent_change Then
            
            ws.Cells(3, 16).Value = percent_change
            ws.Cells(3, 15).Value = ticker
            End If
            
            
            If yearly_change >= 0 Then
            ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
            End If
            
            summary_table_row = summary_table_row + 1
            vol = 0
            percent_change = 0
            
            End If
            
            Next I
            
            
       
            
 
       
    
       
            
            
Next ws

End Sub
