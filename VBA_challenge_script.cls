VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub loop_Stocks()

    'Declare "ws" as worksheet
    Dim ws As Worksheet
    
    'loop through each worksheet
    For Each ws In Worksheets
    
    'Column Headers
    ws.Range("I1").value = "Ticker"
    ws.Range("J1").value = "Yearly Change"
    ws.Range("K1").value = "Percent Change"
    ws.Range("L1").value = "Total Stock Volume"
    ws.Range("O2").value = "Greatest % Increase"
    ws.Range("O3").value = "Greatest % Decrease"
    ws.Range("O4").value = "Greatest Total Volume"
    ws.Range("P1").value = "Ticker"
    ws.Range("Q1").value = "Value"
    
    
    'Variables
    Dim Ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock As Double
    
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_total As Double
    
    Dim open_price As Double
    Dim close_price As Double
    
    Dim previous_amount As Long
    
    Dim last_row As Long
    Dim last_row_k As Long
    
    
    
    total_stock = 0
    summary_table = 2
    previous_amount = 2
    greatest_increase = 0
    greatest_decrease = 0
    greatest_total = 0
    
    
    
    'Finds the last row with data in column A (ticker Column)
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'Loop through each row in column A starting from row 2'
    For i = 2 To last_row
    
        total_stock = total_stock + ws.Cells(i, 7).value
        
        
    
        ' Compare the next ticker with the previous ticker
        If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
            Ticker = ws.Cells(i, 1).value
            ws.Range("I" & summary_table).value = Ticker
            ws.Range("L" & summary_table).value = total_stock
            
            total_stock = 0
            
            open_price = ws.Range("C" & previous_amount)
            
            close_price = ws.Range("F" & i)
            
            yearly_change = close_price - open_price
            ws.Range("J" & summary_table).value = yearly_change
            
            
            'Format change of column J with $
            ws.Range("J" & summary_table).NumberFormat = "$0.00"
            
            
            'if yearly open price is 0 then percent change is 0
            If open_price = 0 Then
                percent_change = 0
                
                
                Else
                yearly_open = ws.Range("C" & previous_amount)
                percent_change = yearly_change / open_price
                
                
                
            End If
            
            ws.Range("K" & summary_table).value = percent_change
            
            ws.Range("K" & summary_table).NumberFormat = "0.00%"
            
            
            'Highlight cell in green if yearly change is positive
            If ws.Range("J" & summary_table).value >= 0 Then
            ws.Range("J" & summary_table).Interior.ColorIndex = 4
            
            
                Else
                'Highligh cell in red if yearly change is negative
                
                ws.Range("J" & summary_table).Interior.ColorIndex = 3
                
                
            End If
            
            
            summary_table = summary_table + 1
            previous_amount = i + 1
            
            
        End If
        
        
        
        Next i
        
        'last row in K column
        last_row_k = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        
        For i = 2 To last_row_k
        
        
            'Greatest % increase
            If ws.Range("K" & i).value > ws.Range("Q2").value Then
                ws.Range("Q2").value = ws.Range("K" & i).value
                ws.Range("P2").value = ws.Range("I" & i).value
                
                
            End If
            
            
            'Greatest % decrease
            If ws.Range("K" & i).value < ws.Range("Q3").value Then
                ws.Range("Q3").value = ws.Range("K" & i).value
                ws.Range("P3").value = ws.Range("I" & i).value
                
                
            End If
            
            
            'Greatest total
            If ws.Range("L" & i).value > ws.Range("Q4").value Then
                ws.Range("Q4").value = ws.Range("L" & i).value
                ws.Range("P4").value = ws.Range("I" & i).value
                
                
            End If
            
            Next i
            
        'Format change of Q2 & Q3 to %
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        
    Next ws
    
End Sub
    
