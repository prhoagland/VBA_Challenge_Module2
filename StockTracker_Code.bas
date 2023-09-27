Attribute VB_Name = "Module1"
Sub StockTracker()

    ' Code written by Peter Hoagland
    ' Produces a Summary Table and Max Data Table for each worksheet in workbook

    Dim ws As Worksheet
    Dim LastRow As Long
    
    For Each ws In ThisWorkbook.Sheets
       
        ' Produce a Summary Table and a Max Data Table for the stock data for each worksheet
    
        ' Label the columns in the Summary Table
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        
        ' Label the rows and columns for the Max Data Table
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        
        ' Find how many rows of data in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Declare RowCounter to track the rows in the Summary Table
        Dim RowCounter As Integer
        RowCounter = 2
        
        ' Declare the variables used for the Summary Table
        Dim BeginOpening As Double
        Dim EndClosing As Double
        Dim YearlyChange As Double
        Dim VolumeCounter As Integer
        Dim PercentChange As Double
        Dim TickerSymbol As String
        
        ' Declare Variables for Max Data Table
        Dim MaxIncTicker As String
        Dim MaxDecTicker As String
        Dim MaxVolTicker As String
        Dim MaxIncValue As Double
        Dim MaxDecValue As Double
        
        ' Move the first Ticker Symbol over into the first row of the Summary Table
        TickerSymbol = ws.Cells(2, 1).Value
        ws.Cells(2, 9) = TickerSymbol
        
        ' Start tracking the Beginning-Year Opening Price and the Volume Counter
        BeginOpening = ws.Cells(2, 3).Value
        StockVolCounter = ws.Cells(2, 7).Value
           
        For row_ = 3 To LastRow
        
            If (ws.Cells(row_, 1).Value <> ws.Cells(row_ - 1, 1)) Then
                
                ' Record the End-of-Year Closing from previous row
                EndClosing = ws.Cells(row_ - 1, 6).Value
                
                ' Calculate the Yearly Change
                YearlyChange = EndClosing - BeginOpening
                ws.Cells(RowCounter, 10) = YearlyChange
                
                ' Highlight the yearly change green for positive and red for negative
                If (YearlyChange > 0) Then
                
                    ws.Cells(RowCounter, 10).Interior.ColorIndex = 4
                    
                ElseIf (YearlyChange < 0) Then
                
                    ws.Cells(RowCounter, 10).Interior.ColorIndex = 3
                    
                End If
                           
                ' Calculate the percent change
                PercentChange = (YearlyChange / BeginOpening)
                ws.Cells(RowCounter, 11) = FormatPercent(PercentChange, 2)
                          
                ' Move the Stock Volume counter over to the Summary Table
                ws.Cells(RowCounter, 12) = StockVolCounter
                
                ' Update Values in Max Table if Necessary
                If (PercentChange > MaxIncValue) Then
                
                    MaxIncValue = PercentChange
                    MaxIncTicker = TickerSymbol
                    
                ElseIf (PercentChange < MaxDecValue) Then
                    
                    MaxDecValue = PercentChange
                    MaxDecTicker = TickerSymbol
                        
                End If
                
                If (StockVolCounter > MaxVolValue) Then
                
                    MaxVolValue = StockVolCounter
                    MaxVolTicker = TickerSymbol
                    
                End If
                
                          
                ' Move down to the next row in the Summary Table and increase the RowCounter
                RowCounter = RowCounter + 1
                           
                ' Move the new ticker symbol to the next row
                TickerSymbol = ws.Cells(row_, 1).Value
                ws.Cells(RowCounter, 9) = TickerSymbol
                
                ' Start tracking Beginning-Year opening and Stock Volume
                BeginOpening = ws.Cells(row_, 3).Value
                StockVolCounter = ws.Cells(row_, 7).Value
                
            Else
                
                ' If same ticker symbol, just update Stock Volume Counter
                StockVolCounter = StockVolCounter + ws.Cells(row_, 7).Value
                   
            End If
            
        Next row_
        
        ' Fill in values for Max Data Table
        ws.Cells(2, 16) = MaxIncTicker
        ws.Cells(3, 16) = MaxDecTicker
        ws.Cells(2, 17) = FormatPercent(MaxIncValue, 2)
        ws.Cells(3, 17) = FormatPercent(MaxDecValue, 2)
        ws.Cells(4, 16) = MaxVolTicker
        ws.Cells(4, 17) = MaxVolValue
        
    Next ws
       
End Sub
