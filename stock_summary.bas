Attribute VB_Name = "Module1"
Sub stock_loop():

    'declaring looping variables
    Dim i As Integer
    'Dim j As Integer
    
    
    'declare intermediary variables
    Dim counter As Integer
    'Dim dates As Integer
    Dim tick_open As Integer
    Dim tick_close As Integer
    
    'variables to create
    Dim ticker As String
    Dim year_change As Double
    Dim pct_change As Double
    Dim vol_tot As Double
    
    'Loop through each worksheet
For Each ws In Worksheets
    
    'set starting values for variables
    counter = 1
    vol_tot = 0
    tick_open = ws.Cells(2, 3).Value
    
    'name new columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'loop to save the ticker symbol in same row new column
    'For i = 2 To 22771 'change to syntax for auto
    
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To last_row
    'For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
       ' If tick_open = -99 Then
       vol_tot = vol_tot + ws.Cells(i, 7)
            
        
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            ticker = ws.Cells(i, 1)
            counter = counter + 1
            tick_close = ws.Cells(i, 6)
            
            'insert statistics
            ws.Cells(counter, 9).Value = ticker
            ws.Cells(counter, 10).Value = (tick_open - tick_close)
            ws.Cells(counter, 11).Value = (tick_open / tick_close) * 100
            ws.Cells(counter, 12).Value = vol_tot
            
            'conditional formatting
           ' If (Cells(counter, 9) >= 0) Then
                'Cells(counter,9).
                
            
            'reset values
            tick_open = ws.Cells(i + 1, 3)
            vol_tot = 0
        
        End If
        
        
        
    Next i
        
Next ws

End Sub

Sub stock_summary():
    
    
    Dim i As Integer
    'Dim j As Integer
    '
    Dim inc_ticker As String
    Dim dec_ticker As String
    Dim vol_ticker As String
    
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    
For Each ws In Worksheets
    
    'specify defaults for max values
    max_increase = -99
    max_decrease = 200
    max_volume = 0
    
    Dim last_row
    last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To 51 'replace with fancy code
        If ws.Cells(i, 10).Value >= 0 Then
            ws.Range("J" & i).Interior.ColorIndex = 43
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Range("J" & i).Interior.ColorIndex = 53
        Else
            Range("J" & i).Interior.ColorIndex = 1 'should only happen if bug
        End If
        
        If ws.Range("K" & i).Value > max_increase Then
        
            max_increase = ws.Range("K" & i).Value
            inc_ticker = ws.Range("I" & i).Value
        End If
        
        'If Range("K" & i).Value < max_decrease Then
         '   max_decrease = Range("K" & i).Value
         '   dec_ticker = Range("I" & i).Value
        'End If
        If max_decrease > ws.Range("K" & i).Value And ws.Range("K" & i).Value Then
            max_decrease = ws.Range("K" & i).Value
           dec_ticker = ws.Range("I" & i).Value
        End If
        
        If ws.Range("L" & i).Value > max_volume Then
            max_volume = ws.Range("L" & i).Value
            vol_ticker = ws.Range("I" & i).Value
        End If
        
    Next i
            
  ws.Range("P2").Value = inc_ticker
  ws.Range("Q2").Value = max_increase
  
  ws.Range("P3").Value = dec_ticker
  ws.Range("Q3").Value = max_decrease
  
  ws.Range("P4").Value = vol_ticker
  ws.Range("Q4").Value = max_volume
    
Next ws

End Sub
