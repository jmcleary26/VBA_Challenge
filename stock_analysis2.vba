Sub stock()
For Each ws In Worksheets

    Dim worksheetname As String
    worksheetname = ws.Name
    
    Dim ticker As String
    Dim open_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim volume As LongLong
    
    yearly_change = 0
    percent_change = 0
    volume = 0
    
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To lastrow
    
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ticker = ws.Cells(i, 1).Value
    open_price = ws.Cells(i, 3).Value
    yearly_change = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
    percent_change = Round(((ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 3).Value * 100), 2)
    volume = volume + ws.Cells(i, 7).Value
    
    ws.Range("I" & summary_table_row).Value = ticker
    ws.Range("J" & summary_table_row).Value = Round(yearly_change, 2)
    ws.Range("K" & summary_table_row).Value = "%" & percent_change
    ws.Range("L" & summary_table_row).Value = volume
    
    summary_table_row = summary_table_row + 1
    
    yearly_change = 0
    percent_change = 0
    volume = 0
    
    Else
    volume = volume + ws.Cells(i, 7).Value
    
    End If

Next i

For i = 2 To lastrow

     If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If

Next i

Next ws

MsgBox ("Done!")

End Sub
