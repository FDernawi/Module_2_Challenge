Attribute VB_Name = "Module1"
'Defining all the variables to be used in the analysis
Sub AnalyzeStockDataMultipleSheets()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim LastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim startRow As Long
    Dim endRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    
    ' Variables to track the greatest values
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    
    ' Variables to store the tickers associated with the greatest values
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
    
    ' Array of sheet names
    sheetNames = Array("B", "C", "D", "E", "F")
    
    ' Loop through each sheet in the array
    For Each sheetName In sheetNames
        ' Set the worksheet to the current sheet
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' Initialize greatest values for this sheet
        greatestIncrease = -999999
        greatestDecrease = 999999
        greatestVolume = 0
        
        ' Adding headers for the output columns in the current sheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Adding headers for the greatest values in columns P and Q
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Finding the last row of data in the ticker column
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize the output row
        outputRow = 2
        
        ' Looping through all rows of data
        i = 2
        Do While i <= LastRow
            ticker = ws.Cells(i, 1).Value
            startRow = i
            
            ' Loop to find the end of the quarter for the same ticker
            Do While ws.Cells(i, 1).Value = ticker And i <= LastRow
                i = i + 1
            Loop
            
            endRow = i - 1
            
            ' Get opening and closing prices
            openPrice = ws.Cells(startRow, 3).Value
            closePrice = ws.Cells(endRow, 6).Value
            
            ' Calculate the quarterly change and percentage change
            quarterlyChange = closePrice - openPrice
            
            If openPrice <> 0 Then
                percentChange = quarterlyChange / openPrice
            Else
                percentChange = 0
            End If
            
            ' Calculate the total volume (ChatGPT)
            totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))
            
            ' Output the ticker data
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = percentChange
            ws.Cells(outputRow, 12).Value = totalVolume
            
            ' Check for greatest percentage increase
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                tickerGreatestIncrease = ticker
            End If
            
            ' Check for greatest percentage decrease
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                tickerGreatestDecrease = ticker
            End If
            
            ' Check for greatest volume
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                tickerGreatestVolume = ticker
            End If
            
            ' Increment the output row
            outputRow = outputRow + 1
        Loop
        
        ' Output the greatest values for this sheet
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = tickerGreatestIncrease
        ws.Cells(2, 17).Value = greatestIncrease
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = tickerGreatestDecrease
        ws.Cells(3, 17).Value = greatestDecrease
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = tickerGreatestVolume
        ws.Cells(4, 17).Value = greatestVolume
        
        ' Apply percent format to the Percent Change column (column K)
        ws.Range("K2:K" & outputRow - 1).NumberFormat = "0.00%"
        
        ' Apply conditional formatting to highlight positive (green) and negative (red) percent changes (ChatGPT)
        With ws.Range("K2:K" & outputRow - 1)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative
        End With
        
        ' Format percent values in the output for greatest increase/decrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
    Next sheetName
    
End Sub
