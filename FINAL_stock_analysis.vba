Sub stock()
For Each ws In Worksheets

    'define variables
    Dim worksheetname As String
    worksheetname = ws.Name
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim volume As LongLong
    Dim previous_close_price As Double
    Dim open_price As Double
    
    'define starting opening price
    open_price = ws.Cells(2, 3).Value
    volume = 0
    
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    'establish last rows
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    last_formatting_row = ws.Cells(Rows.Count, 10).End(xlUp).Row

'start the loop
For i = 2 To lastrow
    
    previous_close_price = ws.Cells(i, 6).Value
        
    'if next row ticker does not equal current row ticker then...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'set values
        ticker = ws.Cells(i, 1).Value
        close_price = ws.Cells(i, 6).Value
        yearly_change = close_price - open_price
        volume = volume + ws.Cells(i, 7).Value
        
        'print values
        ws.Range("I" & summary_table_row).Value = ticker
        ws.Range("J" & summary_table_row).Value = Round(yearly_change, 2)
        ws.Range("L" & summary_table_row).Value = volume
        
        'calculate percent change to account for values = 0
            If open_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / open_price) * 100
            End If
            
        'print value
        ws.Range("K" & summary_table_row).Value = "%" & percent_change
        
        'go to next row
        summary_table_row = summary_table_row + 1
        
        'reset values
        yearly_change = 0
        percent_change = 0
        volume = 0
        open_price = ws.Cells(i + 1, 3).Value
        
    'if next row ticker = current row ticker
    Else
    volume = volume + ws.Cells(i, 7).Value
    ws.Range("J" & summary_table_row).Value = Round(yearly_change, 2)
    
    End If

Next i

For i = 2 To lastrow
    'set conditional formatting
     If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If

Next i

Next ws

MsgBox ("Done!")

End Sub



