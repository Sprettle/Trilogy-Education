Attribute VB_Name = "Module1"
Sub totals():
    
       
    Dim Ticker      As String
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim total_stock As Double
    Dim Ticker_data As Integer
    Dim Row         As Boolean
    
    total_stock = 0
    open_value = 0
    Percentage_Change = 0
    Ticker_data = 2
    Yearly_Change = 0
    
    Row = True
    
    For i = 2 To 9999999:
        
        If Row = True Then
            open_value = Cells(i, 3)
            
            Row = False
            
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1) Then
            
            Ticker = Cells(i, 1).Value
            closing_value = Cells(i, 6)
            Yearly_Change = (closing_value - open_value)
            Percentage_Change = Yearly_Change / open_value
            total_stock = total_stock + Cells(i, 7).Value
            
            Range("I" & Ticker_data).Value = Ticker
            Range("J" & Ticker_data).Value = Yearly_Change
            Range("K" & Ticker_data).Value = Percentage_Change
            Range("K" & Ticker_data).NumberFormat = "0.00%"
            Range("L" & Ticker_data).Value = total_stock
            
            Ticker_data = Ticker_data + 1
            
            Yearly_Change = 0
            Row = True
            
                
        ElseIf Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        Else
            Cells(i, 10).Interior.ColorIndex = 3
            
        End If
        
    Next i
    
   
    
End Sub

