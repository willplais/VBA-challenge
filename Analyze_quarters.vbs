Attribute VB_Name = "Module1"
Sub analyzedata()

    ' Loop through all sheets in worksheet
    For Each ws In Worksheets
        
        ' Create headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
        
        ' Set variables for each ticker
        Dim ActiveTicker As String
        Dim SummaryTableRow As Integer
        SummaryTableRow = 1
        Dim QuarterlyChange As Double
        QuarterlyChange = 0
        Dim PercentChange As Double
        PercentChange = 0
        Dim QuarterlyOpen As Double
        QuarterlyOpen = 0
        Dim QuarterlyClose As Double
        QuarterlyClose = 0
        Dim QuarterlyVolume As Double
        QuarterlyVolume = 0
        
        ' Set variables for summary table
        Dim HighestTickerIncreaseName As String
        Dim HighestTickerIncreaseVal As Double
        HighestTickerIncreaseVal = 0
        
        Dim HighestTickerDecreaseName As String
        Dim HighestTickerDecreaseVal As Double
        HighestTickerDecreaseVal = 0
        
        Dim HighestTickerVolumeName As String
        Dim HighestTickerVolumeVal As Double
        HighestTickerVolumeVal = 0
        
        ' Loop through each row
        For Row = 2 To LastRow
            
            ' If its a new ticker
            If ws.Cells(Row, 1).Value <> ActiveTicker Then
                
              ' Set the Ticker name
              ActiveTicker = ws.Cells(Row, 1).Value
        
              ' Add to the Ticker Total
              SummaryTableRow = SummaryTableRow + 1
              
              QuarterlyOpen = ws.Cells(Row, 3).Value
              QuarterlyVolume = ws.Cells(Row, 7).Value
            
            ' If its the last ticker in a section
            ElseIf ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
              
              ws.Cells(SummaryTableRow, 9).Value = ActiveTicker
        
              QuarterlyClose = ws.Cells(Row, 6).Value
              QuarterlyChange = QuarterlyClose - QuarterlyOpen
              ws.Cells(SummaryTableRow, 10).Value = QuarterlyChange
              
              If QuarterlyChange > 0 Then
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
              Else
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
              End If
              
              PercentChange = QuarterlyChange / QuarterlyOpen
              ws.Cells(SummaryTableRow, 11).Value = PercentChange
              ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
              
              QuarterlyVolume = QuarterlyVolume + ws.Cells(Row, 7).Value
              ws.Cells(SummaryTableRow, 12).Value = QuarterlyVolume
            
                ' Final summary block
                If PercentChange > HighestTickerIncreaseVal Then
                    HighestTickerIncreaseVal = PercentChange
                    HighestTickerIncreaseName = ActiveTicker
                ElseIf PercentChange < HighestTickerDecreaseVal Then
                    HighestTickerDecreaseVal = PercentChange
                    HighestTickerDecreaseName = ActiveTicker
                End If
            
                If QuarterlyVolume > HighestTickerVolumeVal Then
                    HighestTickerVolumeVal = QuarterlyVolume
                    HighestTickerVolumeName = ActiveTicker
                End If
                    
            
            ' If the cell immediately following a row is the same ticker
            Else
        
              QuarterlyVolume = QuarterlyVolume + ws.Cells(Row, 7).Value
        
            End If
        
        Next Row
        
        ' Set the quarterly summary headers
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Set the cell values for the overall summary
        ws.Cells(2, 16).Value = HighestTickerIncreaseName
        ws.Cells(2, 17).Value = HighestTickerIncreaseVal
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = HighestTickerDecreaseName
        ws.Cells(3, 17).Value = HighestTickerDecreaseVal
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = HighestTickerVolumeName
        ws.Cells(4, 17).Value = HighestTickerVolumeVal
        
        ws.Columns("I:Q").AutoFit

    Next ws
End Sub
