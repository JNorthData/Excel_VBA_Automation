Attribute VB_Name = "Module1"
Sub Main()
    Call ExtractDistinctNames
    Call CalculateYearlyChange
    Call CalculateTotalVolume
    Call FinalComparisons
    Call FormatColumnsJK
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
                startDate = ws.Cells(i, "B").Value
                
                ' Find the last row with the same ticker
                Dim j As Long
                j = i + 1
                Do Until j > lastRow Or ticker <> ws.Cells(j, "A").Value
                    j = j + 1
                Loop
                
                ' Stop date is the date of the last row with the same ticker
                stopDate = ws.Cells(j - 1, "B").Value
                
                ' Set yrOpen and yrClose based on Start and Stop dates
                yrOpen = ws.Cells(i, "C").Value
                yrClose = ws.Cells(j - 1, "F").Value
                
                ' Calculate yearly change
                Dim yearlyChange As Double
                yearlyChange = yrClose - yrOpen
                
                ' Calculate percent change
                Dim percentChange As Double
                If yrOpen <> 0 Then
                    percentChange = Round((yearlyChange / yrOpen) * 100, 0)
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
                    ws.Cells(tickerRow, "K").Value = percentChange & "%"

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
    Dim totalVolume As Double
    Dim uniqueTickers As New Collection
    Dim cell As Range
    Dim item As Variant ' Explicitly declare the loop variable as Variant

    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Clear the collection for each worksheet
        Set uniqueTickers = New Collection

        ' Collect unique tickers from column A
        On Error Resume Next ' In case of attempting to add a duplicate key
        For Each cell In ws.Range("A2:A" & lastRow)
            uniqueTickers.Add cell.Value, CStr(cell.Value)
        Next cell
        On Error GoTo 0 ' Turn back on regular error handling

        ' Loop through each unique ticker
        For Each item In uniqueTickers
            ' Calculate total volume for the ticker
            totalVolume = WorksheetFunction.SumIf(ws.Range("A2:A" & lastRow), item, ws.Range("G2:G" & lastRow))
            
            ' Find first occurrence of the ticker in column I and write total volume in column L
            Set cell = ws.Range("I:I").Find(What:=item, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not cell Is Nothing Then
                ws.Cells(cell.Row, "L").Value = totalVolume
                ' Set the number format for the cell to avoid scientific notation
                ws.Cells(cell.Row, "L").NumberFormat = "#,##0"
            End If
        Next item
    Next ws
End Sub


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
        ws.Range("Q2").Value = Format(maxPercentChange * 100, "0") & "%"
        
        ws.Range("P3").Value = minPercentChangeTicker
        ws.Range("Q3").Value = Format(minPercentChange * 100, "0") & "%"
        
        ws.Range("P4").Value = maxTotalVolumeTicker
        ws.Cells(ws.Range("Q4").Row, "Q").Value = maxTotalVolume

        ' Set the number format for the "Greatest Total Volume" to avoid scientific notation
        ws.Cells(ws.Range("Q4").Row, "Q").NumberFormat = "#,##0"
    Next ws

End Sub


Sub FormatColumnsJK()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long

    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "2018" Or ws.Name = "2019" Or ws.Name = "2020" Then
            lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
            Set rng = ws.Range("J1:J" & lastRow)
            
            For Each cell In rng
                If IsNumeric(cell.Value) Then
                    If cell.Value > 0 Then
                        ' Positive Value Formatting: Dark Green Text on Light Green Fill
                        cell.Font.Color = RGB(0, 100, 0) ' Dark Green Text
                        cell.Interior.Color = RGB(198, 239, 206) ' Light Green Fill
                        ws.Range("K" & cell.Row).Font.Color = RGB(0, 100, 0)
                        ws.Range("K" & cell.Row).Interior.Color = RGB(198, 239, 206)
                    ElseIf cell.Value < 0 Then
                        ' Negative Value Formatting: Dark Red Text on Light Red Fill
                        cell.Font.Color = RGB(156, 0, 0) ' Dark Red Text
                        cell.Interior.Color = RGB(255, 199, 206) ' Light Red Fill
                        ws.Range("K" & cell.Row).Font.Color = RGB(156, 0, 0)
                        ws.Range("K" & cell.Row).Interior.Color = RGB(255, 199, 206)
                    End If
                End If
            Next cell
        End If
    Next ws
End Sub