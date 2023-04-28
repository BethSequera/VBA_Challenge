Sub StockAnalysis()
    ' Set variables
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim lastRow As Long
    Dim i As Long
     Dim myRange As Range
    Dim totalCount As Double
    Dim nonBlankCount As Double
    Dim percentage As Double
    Dim start As Long
    Dim outputindex As Integer
For Each ws In Worksheets

    start = 2
    outputindex = 2
    ' Set initial values
    yearlyChange = 0
    percentChange = 0
    totalVolume = 0
    ' Set column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ' Loop through each stock
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Set myRange = ws.Range("A1:A" & lastRow)
    totalCount = myRange.Cells.Count
    nonBlankCount = WorksheetFunction.CountA(myRange)
    percentage = nonBlankCount / totalCount * 100
       ' Get last row of data
       For i = 2 To lastRow ' Start from row 2 since row 1 is headers
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    
        ' Check if we've moved on to a new stock
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
          ' Update variables for the current stock
          closingPrice = ws.Cells(i, 6).Value
          ticker = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(start, 3).Value
          ' Calculate yearly and percent change
          yearlyChange = closingPrice - openingPrice
          If openingPrice <> 0 Then
              percentChange = yearlyChange / openingPrice
          Else
              percentChange = 0
          End If
        start = i + 1
        ' Output information for the previous stock
            ws.Range("I" & outputindex).Value = ticker
            ws.Range("J" & outputindex).Value = yearlyChange
            ws.Range("K" & outputindex).Value = percentChange
            ws.Range("K" & outputindex).NumberFormat = "0.00%" ' Format percent as a percentage
            ws.Range("L" & outputindex).Value = totalVolume
                
            If yearlyChange < 0 Then
                 ws.Range("J" & outputindex).Interior.ColorIndex = 4
            ElseIf yearlyChange > 0 Then
                ws.Range("J" & outputindex).Interior.ColorIndex = 3
            End If
            outputindex = outputindex + 1
            ' Reset variables for the new stock
            yearlyChange = 0
            percentChange = 0
            totalVolume = 0
            
        End If
  Next i
Max = WorksheetFunction.Max(ws.Range("K2:K" & outputindex))
MaxTickerIndex = WorksheetFunction.Match(Max, ws.Range("K2:K" & outputindex), 0)
ws.Range("Q2") = "%" & Max * 100
ws.Range("P2") = ws.Cells(MaxTickerIndex + 1, 9)
Min = WorksheetFunction.Min(ws.Range("K2:K" & outputindex))
MinTickerIndex = WorksheetFunction.Match(Min, ws.Range("K2:K" & outputindex), 0)
ws.Range("Q3") = "%" & Min * 100
ws.Range("P3") = ws.Cells(MinTickerIndex + 1, 9)
Total = WorksheetFunction.Max(ws.Range("L2:L" & outputindex))
TotalTickerIndex = WorksheetFunction.Match(Total, ws.Range("L2:L" & outputindex), 0)
ws.Range("Q4") = Total
ws.Range("P4") = Cells(TotalTickerIndex + 1, 9)

Next ws

End Sub