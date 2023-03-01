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
    
    'set starting values for variables
    counter = 1
    vol_tot = 0
    tick_open = Cells(2, 3).Value
    
    'name new columns
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'loop to save the ticker symbol in same row new column
    'For i = 2 To 22771 'change to syntax for auto
    
    Dim last_row As Long
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To last_row
    'For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
       ' If tick_open = -99 Then
       vol_tot = vol_tot + Cells(i, 7)
            
        
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            ticker = Cells(i, 1)
            counter = counter + 1
            tick_close = Cells(i, 6)
            
            'insert statistics
            Cells(counter, 9).Value = ticker
            Cells(counter, 10).Value = (tick_open - tick_close)
            Cells(counter, 11).Value = (tick_open / tick_close) * 100
            Cells(counter, 12).Value = vol_tot
            
            'conditional formatting
           ' If (Cells(counter, 9) >= 0) Then
                'Cells(counter,9).
                
            
            'reset values
            tick_open = Cells(i + 1, 3)
            vol_tot = 0
        
        End If
        
        
        
    Next i
        
    

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
    
    
    'specify defaults for max values
    max_increase = -99
    max_decrease = 200
    max_volume = 0
    
    Dim last_row
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To 51 'replace with fancy code
        If Cells(i, 10).Value >= 0 Then
            Range("J" & i).Interior.ColorIndex = 43
        ElseIf Cells(i, 10).Value < 0 Then
            Range("J" & i).Interior.ColorIndex = 53
        Else
            Range("J" & i).Interior.ColorIndex = 1 'should only happen if bug
        End If
        
        If Range("K" & i).Value > max_increase Then
        
            max_increase = Range("K" & i).Value
            inc_ticker = Range("I" & i).Value
        End If
        
        'If Range("K" & i).Value < max_decrease Then
         '   max_decrease = Range("K" & i).Value
         '   dec_ticker = Range("I" & i).Value
        'End If
        If max_decrease > Range("K" & i).Value Then
            max_decrease = Range("K" & i).Value
           dec_ticker = Range("I" & i).Value
        End If
        
        If Range("L" & i).Value > max_volume Then
            max_volume = Range("L" & i).Value
            vol_ticker = Range("I" & i).Value
        End If
        
    Next i
            
  Range("P2").Value = inc_ticker
  Range("Q2").Value = max_increase
  
  Range("P3").Value = dec_ticker
  Range("Q3").Value = max_decrease
  
  Range("P4").Value = vol_ticker
  Range("Q4").Value = max_volume
    
End Sub
