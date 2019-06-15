Sub StockData()

    For Each WS In Worksheets
    
' Determine the Last Row

LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

' Print the headings for summary table

        WS.Cells(1, "I").Value = "Ticker"
        WS.Cells(1, "J").Value = "Yearly Change"
        WS.Cells(1, "K").Value = "Percent Change"
        WS.Cells(1, "L").Value = "Total Stock Volume"
        
'Set up the variables

        Dim Open_Value As Double
        Dim Close_Value As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
'Set First Open Value

        Open_Value = WS.Cells(2, Column + 2).Value
         
         ' Loop through all tickers
        
        For i = 2 To LastRow
        
         ' Check if we are still within the same ticker symbol, if it is not...
         
            If WS.Cells(i + 1, Column).Value <> WS.Cells(i, Column).Value Then
            
          'Print the Ticker name
          
                Ticker_Name = WS.Cells(i, Column).Value
                WS.Cells(Row, Column + 8).Value = Ticker_Name
                
           'Print the Close Value
           
                Close_Value = WS.Cells(i, Column + 5).Value
                
            'Calculate the Yearly Change
            
                Yearly_Change = Close_Value - Open_Value
                WS.Cells(Row, Column + 9).Value = Yearly_Change
                
            ' Add Percent Change
            
                If (Open_Value = 0 And Close_Value = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Value = 0 And Close_Value <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Value
                    WS.Cells(Row, Column + 10).Value = Percent_Change
                    WS.Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                
            ' Print the Total of the Volume of the ticker
            
                Volume = Volume + WS.Cells(i, Column + 6).Value
                WS.Cells(Row, Column + 11).Value = Volume
                
            ' Add 1 to the summary table row
            
                Row = Row + 1
                
            ' reset the Open Value
            
                Open_Value = WS.Cells(i + 1, Column + 2)
                
            ' reset the Volume Total
            
                Volume = 0
                
            'if cells contain the same ticker
            
            Else
                Volume = Volume + WS.Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Find the Last Row of Yearly Change per each Worksheet
        
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        ' Set the Cell Colors
        
        For j = 2 To YCLastRow
            If (WS.Cells(j, Column + 9).Value > 0 Or WS.Cells(j, Column + 9).Value = 0) Then
                WS.Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf WS.Cells(j, Column + 9).Value < 0 Then
                WS.Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        ' Print the values of Greatest % Increase, % Decrease, and Total Volume
        
        WS.Cells(2, Column + 14).Value = "Greatest % Increase"
        WS.Cells(3, Column + 14).Value = "Greatest % Decrease"
        WS.Cells(4, Column + 14).Value = "Greatest Total Volume"
        WS.Cells(1, Column + 15).Value = "Ticker"
        WS.Cells(1, Column + 16).Value = "Value"
        
        ' Loop through each row and find the greatest value and the corresponding ticker
        
        For Z = 2 To YCLastRow
            If WS.Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                WS.Cells(2, Column + 15).Value = WS.Cells(Z, Column + 8).Value
                WS.Cells(2, Column + 16).Value = WS.Cells(Z, Column + 10).Value
                WS.Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf WS.Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                WS.Cells(3, Column + 15).Value = WS.Cells(Z, Column + 8).Value
                WS.Cells(3, Column + 16).Value = WS.Cells(Z, Column + 10).Value
                WS.Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf WS.Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                WS.Cells(4, Column + 15).Value = WS.Cells(Z, Column + 8).Value
                WS.Cells(4, Column + 16).Value = WS.Cells(Z, Column + 11).Value
            End If
        Next Z
        
   'Autofit Column
   
WS.Range("J1").EntireColumn.AutoFit
WS.Range("K1").EntireColumn.AutoFit
WS.Range("L1").EntireColumn.AutoFit
WS.Range("O1").EntireColumn.AutoFit
WS.Range("Q1").EntireColumn.AutoFit

    Next WS
        
End Sub

