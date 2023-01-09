Sub assignment2()

  Dim WS As Worksheet
  For Each WS In Worksheets

    
    'setting up the variables for the script
    Dim Ticker_Name As String
    Dim Ticker_Volume As Double
    Ticker_Volume = 0
    Dim Ticker_Summary As Integer
    Ticker_Summary = 2
    Dim open_price As Double
    open_price = WS.Cells(2, 3).Value
    Dim close_price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    'creating lables for table
    WS.Cells(1, 9).Value = "Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Volume"
    
    'creating the for loop statement that looks for change in cell value
    For i = 2 To lastrow
    
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
        Ticker_Name = WS.Cells(i, 1).Value
        close_price = WS.Cells(i, 6).Value
        Yearly_Change = (close_price - open_price)
        Ticker_Volume = Ticker_Volume + WS.Cells(i, 7).Value
        WS.Range("I" & Ticker_Summary).Value = Ticker_Name
        WS.Range("L" & Ticker_Summary).Value = Ticker_Volume
        WS.Range("J" & Ticker_Summary).Value = Yearly_Change
        
      
     'creating the loop to calculate percent change, including 0
         If (open_price = 0) Then
         Percent_Change = 0
            Else
            Percent_Change = (Yearly_Change / open_price)
        End If
        
     'Printing the yearly change in the table
        WS.Range("K" & Ticker_Summary).Value = Percent_Change
        WS.Range("K" & Ticker_Summary).NumberFormat = "0.00%"
        
        Ticker_Volume = 0
        Ticker_Summary = Ticker_Summary + 1
    
        open_price = WS.Cells(i + 1, 3).Value
        
        Else
        
        Ticker_Volume = Ticker_Volume + WS.Cells(i, 7).Value
        
        End If
        
    Next i
        
        lastrow_summary_table = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow_summary_table
            If WS.Cells(i, 11).Value > 0 Then
            WS.Cells(i, 10).Interior.ColorIndex = 4
            Else: WS.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
            Next i
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        
    For i = 2 To lastrow_summary_table
            If WS.Cells(i, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & lastrow_summary_table)) Then
            WS.Cells(2, 16).Value = WS.Cells(i, 9).Value
            WS.Cells(2, 17).Value = WS.Cells(i, 11).Value
            WS.Cells(2, 17).NumberFormat = "0.00%"
            End If
            
            If WS.Cells(i, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & lastrow_summary_table)) Then
            WS.Cells(3, 16).Value = WS.Cells(i, 9).Value
            WS.Cells(3, 17).Value = WS.Cells(i, 11).Value
            WS.Cells(3, 17).NumberFormat = "0.00%"
            End If
            
            If WS.Cells(i, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & lastrow_summary_table)) Then
            WS.Cells(4, 16).Value = WS.Cells(i, 9).Value
            WS.Cells(4, 17).Value = WS.Cells(i, 12).Value
            
            End If
        Next i
    Next WS
    End Sub
