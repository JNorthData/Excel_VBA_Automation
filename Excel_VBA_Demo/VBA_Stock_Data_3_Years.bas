Attribute VB_Name = "Module1"
Sub Main()
    Call ExtractDistinctNames
    Call CalculateYearlyChange
    Call CalculateTotalVolume
    Call FinalComparisons
End Sub

'   Find all DISTINCT stock tickers
Sub ExtractDistinctNames()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim namesList As New Collection
    
    '   Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        ' Add column headers for Ticker and Yearly Change
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Voume"
    
        '   Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        '   Loop through each row of data
        For i = 2 To lastRow
            Dim name As Variant
            
            '   Get the name from column A
            name = ws.Cells(i, "A").Value
            
            '   Check if the name already exists in the collection
            On Error Resume Next
            namesList.Add name, CStr(name)
            On Error GoTo 0
        Next i
    Next ws
    
    '   Output the distinct names in column I
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        '   Output the distinct names in column I
        Dim rowIndex As Long
        rowIndex = 2
        For Each name In namesList
            ws.Cells(rowIndex, "I").Value = name
            rowIndex = rowIndex + 1
        Next name
    Next

End Sub


' Calculate yearly change for each stock ticker
Sub CalculateYearlyChange()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim startDate As Date
    Dim stopDate As Date
    Dim yrOpen As Double
    Dim yrClose As Double
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through each row of data
        For i = 2 To lastRow
            ' Get the ticker from column A
            ticker = ws.Cells(i, "A").Value
            
            ' Check if the ticker value is different from the cell value above it
            If ticker <> ws.Cells(i - 1, "A").Value Then
                ' Start date is the current row's date
                startDate = CDate(ws.Cells(i, "B").Value)
                
                ' Find the last row with the same ticker
                Dim j As Long
                j = i + 1
                Do Until j > lastRow Or ticker <> ws.Cells(j, "A").Value
                    j = j + 1
                Loop
                
                ' Stop date is the date of the last row with the same ticker
                stopDate = CDate(ws.Cells(j - 1, "B").Value)
                
                ' Set yrOpen and yrClose based on Start and Stop dates
                yrOpen = CDate(ws.Cells(i, "C").Value)
                yrClose = CDate(ws.Cells(j - 1, "F").Value)
                
                ' Calculate yearly change
                Dim yearlyChange As Double
                yearlyChange = yrClose - yrOpen
                
                ' Calculate percent change
                Dim percentChange As Double
                If yrOpen <> 0 Then
                    percentChange = round((yearlyChange / yrOpen), 2)
                Else
                    percentChange = 0
                End If
                
                ' Insert yearly change and percent change in columns J and K for the corresponding ticker in column I
                Dim tickerRange As Range
                Set tickerRange = ws.Range("I:I").Find(What:=ticker, LookIn:=xlValues, LookAt:=xlWhole)
                
                If Not tickerRange Is Nothing Then
                    Dim tickerRow As Long
                    tickerRow = tickerRange.Row
                    
                    ws.Cells(tickerRow, "J").Value = yearlyChange
                    ws.Cells(tickerRow, "K").Value = percentChange
                End If
                
                ' Skip the rows for the current ticker
                i = j - 1
            End If
        Next i
    Next ws

End Sub

' Calculate total volume for each stock ticker
Sub CalculateTotalVolume()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim totalVolume As Double
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in column A
        lastRow = CDate(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1)
        
        ' Loop through each row of data
        For i = 2 To lastRow
            ' Get the ticker from column A
            ticker = CDate(ws.Cells(i, "A").Value)
            
            ' Check if the ticker value is different from the cell value above it
            If ticker <> CDate(ws.Cells(i - 1, "A").Value Then)
                ' Calculate total volume
                totalVolume = WorksheetFunction.SumIf(ws.Range("A:A"), ticker, ws.Range("G:G"))
                
                ' Insert total volume in column L for the corresponding ticker in column I
                Dim tickerRange As Range
                Set tickerRange = CDate(ws.Range("I:I").Find(What:=ticker, LookIn:=xlValues, LookAt:=xlWhole))
                
                If Not tickerRange Is Nothing Then
                    Dim tickerRow As Long
                    tickerRow = tickerRange.Row
                    
                    CDate(ws.Cells(tickerRow, "L").Value = totalVolume)
                End If
            End If
        Next i
    Next ws

End Sub

' Perform final comparisons
Sub FinalComparisons()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxPercentChange As Double
    Dim minPercentChange As Double
    Dim maxTotalVolume As Double
    Dim maxPercentChangeTicker As String
    Dim minPercentChangeTicker As String
    Dim maxTotalVolumeTicker As String
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in column I
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row - 1
        
        ' Find the maximum percent change and corresponding ticker
        maxPercentChange = WorksheetFunction.Max(ws.Range("K:K"))
        maxPercentChangeTicker = ws.Cells(Application.Match(maxPercentChange, ws.Range("K:K"), 0), "I").Value
        
        ' Find the minimum percent change and corresponding ticker
        minPercentChange = WorksheetFunction.Min(ws.Range("K:K"))
        minPercentChangeTicker = ws.Cells(Application.Match(minPercentChange, ws.Range("K:K"), 0), "I").Value
        
        ' Find the maximum total volume and corresponding ticker
        maxTotalVolume = WorksheetFunction.Max(ws.Range("L:L"))
        maxTotalVolumeTicker = ws.Cells(Application.Match(maxTotalVolume, ws.Range("L:L"), 0), "I").Value
        
        ' Insert headers and values
        ws.Range("O2").Value = "Greatest % Increase:"
        ws.Range("O3").Value = "Greatest % Decrease:"
        ws.Range("O4").Value = "Greatest Total Volume:"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("P2").Value = maxPercentChangeTicker
        ws.Range("Q2").Value = maxPercentChange
        
        ws.Range("P3").Value = minPercentChangeTicker
        ws.Range("Q3").Value = minPercentChange
        
        ws.Range("P4").Value = maxTotalVolumeTicker
        ws.Range("Q4").Value = maxTotalVolume
    Next ws

End Sub

