Sub VbaStocks():

    Dim workSheetName As String
    Dim ticker As String
    Dim openValue As Double
    Dim closeValue As Double
    Dim stockVolume As Double
    Dim lastRow As Long
    Dim lastColumn As Integer
    Dim reportRow As Integer
    Dim rowCounter As Integer
    Dim yearlyChange As Double
    Dim yearlyPCTChange As Double
    Dim lastRowHeader As Long
    Dim Max As Double
    Dim Min As Double
    Dim MaxV As Double
    
     
'Iterate through all the worksheets in the file.
    For Each ws In Worksheets
        workSheetName = ws.Name
        lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        reportRow = 2
        stockVolume = 0
        rowCounter = 0
        yearlyChange = 0
        yearlyPCTChange = 0
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        'Iterate through all the rows in the worksheet.
        For i = 2 To lastRow
            Dim sV As Double
            sV = ws.Cells(i, 7).Value

            
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                ticker = ws.Cells(i, 1).Value
                stockVolume = stockVolume + sV
                ws.Cells(reportRow, 9).Value = ticker
                ws.Cells(reportRow, 12).Value = stockVolume
                closeValue = ws.Cells(i, 6)
                yearlyChange = closeValue - openValue
                ws.Cells(reportRow, 10).Value = yearlyChange
                              
                

                If openValue <> 0 Then
                    yearlyPCTChange = (yearlyChange / openValue) * 100
                    If yearlyPCTChange > 0 Then
                    ws.Cells(reportRow, 11).Value = yearlyPCTChange
                    ws.Cells(reportRow, 10).Interior.ColorIndex = 4
                    ElseIf yearlyPCTChange < 0 Then
                    ws.Cells(reportRow, 11).Value = yearlyPCTChange
                    ws.Cells(reportRow, 10).Interior.ColorIndex = 3
                    End If
                 End If
                
                reportRow = reportRow + 1
                stockVolume = 0
                'Variable to check the first row of the ticket to capture Open Value.
                rowCounter = 0
            Else
                stockVolume = sV + stockVolume
                If (rowCounter = 0) Then
                    openValue = ws.Cells(i, 3).Value
                    ws.Cells(reportRow, 11).Value = openValue

                    rowCounter = rowCounter + 1
                End If
                
            End If
            
        Next i
    lastColumnHeasder = ws.Cells(11, Columns.Count).End(xlToLeft).Column
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % increase"
    ws.Range("O3") = "Greatest % decrease"
    ws.Range("O4") = "Greatest total volume"
    lastRowHeader = ws.Cells(Rows.Count, 9).End(xlUp).Row
    'code to capture and display min and max values.
    ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K2", ws.Cells(lastRowHeader, 11)))
    ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K2", ws.Cells(lastRowHeader, 11)))
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2", ws.Cells(lastRowHeader, 12)))
    Max = WorksheetFunction.Max(ws.Range("K2", ws.Cells(lastRowHeader, 11)))
    Min = WorksheetFunction.Min(ws.Range("K2", ws.Cells(lastRowHeader, 11)))
    MaxV = WorksheetFunction.Max(ws.Range("L2", ws.Cells(lastRowHeader, 12)))
    ws.Columns("A:Z").AutoFit
    For e = 2 To lastRowHeader
        If (Max = ws.Cells(e, 11)) Then
            ws.Range("P2") = ws.Cells(e, 9)
        ElseIf (Min = ws.Cells(e, 11)) Then
                ws.Range("P3") = ws.Cells(e, 9)
        ElseIf (MaxV = ws.Cells(e, 12)) Then
                ws.Range("P4") = ws.Cells(e, 9)
         
        End If
    Next e
    Next ws


End Sub





