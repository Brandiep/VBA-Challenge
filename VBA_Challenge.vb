Sub Stocks()

    'loop through the tabs
    Dim ws As Worksheet
    For Each ws In Worksheets

    'set the variables
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim i As Long
    
    'set initial stock total
    Total_Stock_Volume = 0
    
    'set up the summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'dictate where to start the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Open_Price = ws.Cells(2, 3).Value
    
    'set up the loop
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'create a check for same ticker, if its not then....
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'set the ticker and print it
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        'add to the total stock volume and print it
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
        'set the closing price
            Close_Price = ws.Cells(i, 6).Value
            
        'calculate yearly change and print it
            Yearly_Change = Open_Price - Close_Price
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
         'calculate percent change and print it
            Percent_Change = Yearly_Change / Open_Price * 100
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        'add a row to the summary table
            Summary_Table_Row = Summary_Table_Row + 1
            
        'reset the total stock volume
            Total_Stock_Volume = 0
            
        'set open price
        'cannot divide by 0 - take next ticker open price
            If ws.Cells(i + 1, 3).Value = 0 Then
                ' For loop to loop through open price to find new non zero
                Open_Price = Cells(i + 2, 3).Value
            Else
                Open_Price = Cells(i + 1, 3).Value
            End If
            
       'if the ticker is the same as the previous one then...
        Else
         'calculate the total volume for the first ticker
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
            
         
        End If
    Next i
    
    'add conditional formatting
    For x = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        If ws.Cells(x, 10).Value > 0 Then
            ws.Cells(x, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(x, 10).Interior.ColorIndex = 3
        End If
    Next x
        
   Next ws

End Sub

