Attribute VB_Name = "Module1"
Sub CalculatorMain()
    Dim ws As Worksheet
    
    ' Iterate through each worksheet
    For Each ws In Worksheets
        ' Call subroutine to write headers
        Call WriteHeaders(ws)
        
        ' Call subroutine to calculate and fill symbol summary
        Call FillSymbolSummary(ws)
        ' Formatting the new columns
        Call FormatColumns(ws)
        ' Calculate and display greatest % increase, % decrease, and total volume
        Call CalculateAndDisplayGreatestValues(ws)
    Next ws
End Sub

Sub WriteHeaders(ws As Worksheet)
    ' Write headers
    ws.Cells(1, 9).Value = "Tickers"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
End Sub



Sub FillSymbolSummary(ws As Worksheet)
    Dim rowsCount As Long
    Dim symbolCount As Long
    Dim startOfYearOpen As Double, EndOfYearClose As Double
    Dim totalStockVolume As Double
    Dim i As Long ' Declare i as Long
    
    ' Calculate the last row in column A
    rowsCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Reset variables
    symbolCount = 0
    startOfYearOpen = 0
    EndOfYearClose = 0
    totalStockVolume = 0
    
    ' Iterate through each row in the worksheet
    For i = 2 To rowsCount
        ' Check if we have a new symbol
        If IsNewSymbol(ws, i) Then
            ' Calculate and fill data for previous symbol if it exists
            Call CalculateSymbol(ws, i, symbolCount, startOfYearOpen, totalStockVolume)
            
            ' Initialize for new symbol
            Call InitializeNewSymbol(ws, i, symbolCount, startOfYearOpen, totalStockVolume)
        End If
        
        ' Calculate total stock volume for the symbol
        totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
    Next i
    
    ' Calculate and fill data for the last symbol in the worksheet
    Call CalculateSymbol(ws, i, symbolCount, startOfYearOpen, totalStockVolume)
End Sub


Function IsNewSymbol(ws As Worksheet, rowIndex As Long) As Boolean
    ' Check if the current symbol is different from the previous row
    If ws.Cells(rowIndex, 1).Value <> ws.Cells(rowIndex - 1, 1).Value Then
        IsNewSymbol = True
    Else
        IsNewSymbol = False
    End If
End Function

Sub CalculateSymbol(ws As Worksheet, i As Long, symbolCount As Long, _
                     startOfYearOpen As Double, totalStockVolume As Double)
    Dim EndOfYearClose As Double
    
    ' Check if symbolCount > 0, indicating that it's not the first symbol
    If symbolCount > 0 Then
        ' Calculate the end of year close for the previous symbol if it wasn't the first symbol
        EndOfYearClose = ws.Cells(i - 1, 6).Value
        
        ' Write the yearly change and total stock volume for the symbol
        ws.Cells(symbolCount + 1, 10).Value = EndOfYearClose - startOfYearOpen
        ws.Cells(symbolCount + 1, 11).Value = (EndOfYearClose - startOfYearOpen) / startOfYearOpen
        ws.Cells(symbolCount + 1, 12).Value = totalStockVolume
    End If
End Sub

Sub InitializeNewSymbol(ws As Worksheet, i As Long, ByRef symbolCount As Long, _
                        ByRef startOfYearOpen As Double, ByRef totalStockVolume As Double)
    ' Initialize for new symbol
    startOfYearOpen = ws.Cells(i, 3).Value
    symbolCount = symbolCount + 1
    ws.Cells(symbolCount + 1, 9).Value = ws.Cells(i, 1).Value
    totalStockVolume = 0
End Sub

Sub FormatColumns(ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    
    ' Find the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    ' Format column 10 and 11(Yearly Change)
    Call ApplyConditionalFormatting(ws)
    
    ' Format column 11 (Total Stock Volume) as percentage
    Set rng = ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11))
    rng.NumberFormat = "0.00%"
    
    ' Format columns 9, 10, 11, and 12 to autofit
    ws.Columns("I:L").AutoFit
End Sub

Sub ApplyConditionalFormatting(ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    
    ' Find the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    ' Define the range for column 10 and 11 (Yearly Changes)
    Set rng = ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 11))
    
    ' Clear existing conditional formatting rules
    rng.FormatConditions.Delete
    
    ' Add new conditional formatting rule for negative numbers
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = vbRed
    End With
    
    ' Add new conditional formatting rule for non-negative numbers
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
        .Interior.Color = vbGreen
    End With
End Sub

Sub CalculateAndDisplayGreatestValues(ws As Worksheet)
    Dim lastRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    
    ' Find the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    ' Initialize variables
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    tickerIncrease = ""
    tickerDecrease = ""
    tickerVolume = ""
    
    ' Loop through the data to find greatest % increase, % decrease, and total volume
    For i = 2 To lastRow
        ' Check for greatest % increase
        If ws.Cells(i, 11).Value > greatestIncrease Then
            greatestIncrease = ws.Cells(i, 11).Value
            tickerIncrease = ws.Cells(i, 9).Value
        End If
        
        ' Check for greatest % decrease
        If ws.Cells(i, 11).Value < greatestDecrease Then
            greatestDecrease = ws.Cells(i, 11).Value
            tickerDecrease = ws.Cells(i, 9).Value
        End If
        
        ' Check for greatest total volume
        If ws.Cells(i, 12).Value > greatestVolume Then
            greatestVolume = ws.Cells(i, 12).Value
            tickerVolume = ws.Cells(i, 9).Value
        End If
    Next i
    
    ' Display the results in columns O, P, and Q
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ws.Cells(2, 16).Value = tickerIncrease
    ws.Cells(3, 16).Value = tickerDecrease
    ws.Cells(4, 16).Value = tickerVolume
    
    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(4, 17).Value = greatestVolume
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ' Autofit columns P, Q, and R
    ws.Columns("P:Q").AutoFit
End Sub

