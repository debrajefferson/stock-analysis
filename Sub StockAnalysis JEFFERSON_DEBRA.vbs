Sub StockAnalysis()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestIncreasePercent As Double
    Dim GreatestDecreaseTicker As String
    Dim GreatestDecreasePercent As Double
    Dim GreatestTotalVolumeTicker As String
    Dim GreatestTotalVolume As Long
    Dim SummaryTableRow As Integer
    

    
    ' Loop through each sheet
    For Each ws In Worksheets
        ' Find the last row with data
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize variables
        Ticker = ws.Cells(2, 1).Value
        OpenPrice = ws.Cells(2, 3).Value
        TotalVolume = 0
        SummaryTableRow = 2
        
        'Place new labels for Columns and "Greatest" table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest Percent Increase"
        ws.Cells(3, 15).Value = "Greatest Percent Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        
        
        ' Loop through the data
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> Ticker Then
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If
                
                ' Output the information
                ws.Cells(i, 9).Value = Ticker
                ws.Cells(i, 10).Value = YearlyChange
                ws.Cells(i, 11).Value = PercentChange
                ws.Cells(i, 12).Value = TotalVolume
                
                'Add Stock Information to Summary Table
                
                'Print the Ticker symbol in the Summary Table
                ws.Range("I" & SummaryTableRow).Value = Ticker
                
                'Print the Yearly Change in the Summary Table
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
                'Print the Percent Change in the Summary Table
                ws.Range("K" & SummaryTableRow).Value = PercentChange & "%"
                
                'Print the Total Stock Volume in the Summary Table
                ws.Range("L" & SummaryTableRow).Value = TotalVolume
                
                'Add one to the Summary Table Row interger
                SummaryTableRow = SummaryTableRow + 1
                
                ' Reset variables for the next ticker
                Ticker = ws.Cells(i + 1, 1).Value
                OpenPrice = ws.Cells(i + 1, 3).Value
                TotalVolume = 0
                
                Else
                
                TotalVolume = TotalVolume + ws.Cells(i + 1, 7).Value
            
            End If
        
        Next i
        
        ' Calculate and store Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        
           ' Find the last row of Summary Table Data
            LastRowTable = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
         
            'Loop through Summary Table Data
            
                GreatestIncreaseTicker = ws.Cells(2, 16).Value
                GreatestIncreasePercent = ws.Cells(2, 11).Value
                GreatestDecreaseTicker = ws.Cells(3, 16).Value
                GreatestDecreasePercent = ws.Cells(2, 11).Value
                GreatestTotalVolumeTicker = ws.Cells(4, 16).Value
                GreatestTotalVolume = ws.Cells(2, 12).Value
           
           For j = 2 To LastRowTable
        
                If ws.Cells(j, 11).Value > GreatestIncreasePercent Then
                    GreatestIncreaseTicker = ws.Cells(j, 9).Value
                    GreatestIncreasePercent = ws.Cells(j, 11).Value
                End If
            
                If ws.Cells(j, 11).Value < GreatestDecreasePercent Then
                    GreatestDecreaseTicker = ws.Cells(j, 9).Value
                    GreatestDecreasePercent = ws.Cells(j, 11).Value
                End If
            
                If ws.Cells(j, 12).Value > GreatestTotalVolume Then
                    GreatestTotalVolumeTicker = ws.Cells(j, 9).Value
                    GreatestTotalVolume = ws.Cells(j, 11).Value
                End If
        
            
            ' Output Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
    
                ws.Cells(2, 16).Value = GreatestIncreaseTicker
                ws.Cells(2, 17).Value = GreatestIncreasePercent & "%"
    
                ws.Cells(3, 16).Value = GreatestDecreaseTicker
                ws.Cells(3, 17).Value = GreatestDecreasePercent & "%"
    
                ws.Cells(4, 16).Value = GreatestTotalVolumeTicker
                ws.Cells(4, 17).Value = GreatestTotalVolume
          
            
            Next j
        
        'Color Yearly Change on Summary Table Red if Negative or Green if Positive
            For k = 2 To LastRowTable
                If ws.Cells(k, 10).Value > 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 4
            
                ElseIf ws.Cells(k, 10).Value < 0 Then
                    ws.Cells(k, 10).Interior.ColorIndex = 3
        
                End If
            
            Next k
    
    Next ws
    
End Sub

