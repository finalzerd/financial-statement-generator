Sub CreateStatementOfChangesInEquity(targetWorkbook As Workbook)
    Dim ws As Worksheet
    Dim infoSheet As Worksheet
    Dim trialBalance1 As Worksheet
    Dim trialBalance2 As Worksheet
    Dim trialPL1 As Worksheet
    Dim trialPL2 As Worksheet
    Dim row As Long
    Dim year As String
    Dim PreviousYear As String
    Dim isFirstYear As Boolean
    
    ' Check if it's a "บริษัทจำกัด"
    Set infoSheet = targetWorkbook.Sheets("Info")
    If infoSheet.Range("B2").Value <> "บริษัทจำกัด" Then
        Exit Sub
    End If
    
    ' Create a new sheet named SCE
    Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    ws.Name = "SCE"
    
    ' Create the header
    CreateHeader ws, "Statement of Changes in Equity"
    
    ' Get the year from the Info sheet
    year = infoSheet.Range("B3").Value
    PreviousYear = CStr(CLng(year) - 1)
    
    ' Check if it's the first year
    On Error Resume Next
    Set trialBalance2 = targetWorkbook.Sheets("Trial Balance 2")
    isFirstYear = (trialBalance2 Is Nothing)
    On Error GoTo 0
    
    ' Set references to Trial Balance and Trial PL sheets
    Set trialBalance1 = targetWorkbook.Sheets("Trial Balance 1")
    Set trialPL1 = targetWorkbook.Sheets("Trial PL 1")
    If Not isFirstYear Then
        Set trialPL2 = targetWorkbook.Sheets("Trial PL 2")
    End If
    
    ' Start adding details
    row = 5 ' Start below the header
    
    ' Add "หน่วย:บาท" with top and bottom borders
    With ws.Range(ws.Cells(row, 3), ws.Cells(row, 9))
        .Merge
        .Value = "หน่วย:บาท"
        .HorizontalAlignment = xlCenter
    End With
    
    row = row + 1
    
    ' Add column headers with bottom borders, wrap text, and center alignment
    With ws.Cells(row, 3)
        .Value = "ทุนเรือนหุ้นที่ออกและชำระแล้ว"
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Cells(row, 6)
        .Value = "กำไร(ขาดทุน)สะสม"
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Cells(row, 9)
        .Value = "รวม"
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Adjust row height to accommodate wrapped text
    ws.Rows(row).AutoFit

    row = row + 1
    
    If Not isFirstYear Then
        ' Previous Year
        ' Add opening balance for previous year
        ws.Cells(row, 1).Value = "ยอดคงเหลือ ณ วันที่ 1 มกราคม " & PreviousYear
        Dim paidUpCapitalPreviousYear As Double
        paidUpCapitalPreviousYear = GetAmountFromSheet(trialBalance2, "3010", "F")
        ws.Cells(row, 3).Value = paidUpCapitalPreviousYear
        row = row + 1
        
        ' Add net profit for previous year
        ws.Cells(row, 1).Value = "กำไร (ขาดทุน) สุทธิ สำหรับปี " & PreviousYear
        Dim netProfitPreviousYear As Double
        netProfitPreviousYear = GetLastAmountInColumn(trialPL2, "G")
        ws.Cells(row, 6).Value = netProfitPreviousYear
        ws.Cells(row, 9).Value = netProfitPreviousYear
        row = row + 1
        
        ' Add closing balance for previous year
        ws.Cells(row, 1).Value = "ยอดคงเหลือ ณ วันที่ 31 ธันวาคม " & PreviousYear
        Dim retainedEarningsPreviousYear As Double
        retainedEarningsPreviousYear = GetAmountFromSheet(trialBalance2, "3020", "F")
        ws.Cells(row, 3).Value = paidUpCapitalPreviousYear
        ws.Cells(row, 6).Value = retainedEarningsPreviousYear
        ws.Cells(row, 9).Value = paidUpCapitalPreviousYear + retainedEarningsPreviousYear
        
        ' Calculate and update the opening balance for previous year
        ws.Cells(row - 2, 6).Value = retainedEarningsPreviousYear - netProfitPreviousYear
        ws.Cells(row - 2, 9).Value = paidUpCapitalPreviousYear + retainedEarningsPreviousYear - netProfitPreviousYear
        
        ' Add two blank rows
        row = row + 3
    End If
    
    ' Current Year
    ' Add opening balance for current year
    ws.Cells(row, 1).Value = "ยอดคงเหลือ ณ วันที่ 1 มกราคม " & year
    Dim paidUpCapitalCurrentYear As Double
    paidUpCapitalCurrentYear = GetAmountFromSheet(trialBalance1, "3010", "F")
    ws.Cells(row, 3).Value = paidUpCapitalCurrentYear
    If isFirstYear Then
        ws.Cells(row, 6).Value = 0  ' For the first year, opening retained earnings is 0
    Else
        ws.Cells(row, 6).Value = retainedEarningsPreviousYear  ' Use previous year's closing balance
    End If
    ws.Cells(row, 9).Value = ws.Cells(row, 3).Value + ws.Cells(row, 6).Value
    row = row + 1
    
    ' Add net profit for current year
    ws.Cells(row, 1).Value = "กำไร (ขาดทุน) สุทธิ สำหรับปี " & year
    Dim netProfitCurrentYear As Double
    netProfitCurrentYear = GetLastAmountInColumn(trialPL1, "G")
    ws.Cells(row, 6).Value = netProfitCurrentYear
    ws.Cells(row, 9).Value = netProfitCurrentYear
    row = row + 1
    
    ' Add closing balance for current year
    ws.Cells(row, 1).Value = "ยอดคงเหลือ ณ วันที่ 31 ธันวาคม " & year
    Dim retainedEarningsCurrentYear As Double
    retainedEarningsCurrentYear = GetAmountFromSheet(trialBalance1, "3020", "F")
    ws.Cells(row, 3).Value = paidUpCapitalCurrentYear
    ws.Cells(row, 6).Value = retainedEarningsCurrentYear
    ws.Cells(row, 9).Value = paidUpCapitalCurrentYear + retainedEarningsCurrentYear
    
    ' Format the worksheet
    FormatSCEWorksheet ws
End Sub

Function GetAmountFromSheet(ws As Worksheet, accountCode As String, column As String) As Double
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value = accountCode Then
            GetAmountFromSheet = ws.Cells(i, ws.Columns(column).column).Value
            Exit Function
        End If
    Next i
    
    GetAmountFromSheet = 0
End Function

Function GetLastAmountInColumn(ws As Worksheet, column As String) As Double
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, ws.Columns(column).column).End(xlUp).row
    GetLastAmountInColumn = ws.Cells(lastRow, ws.Columns(column).column).Value
End Function

Sub FormatSCEWorksheet(ws As Worksheet)
    ' Apply Thai Sarabun font and font size 14 to the worksheet
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Set number format to use comma style for columns C, E, and G
    ws.Columns("C:C").NumberFormat = "#,##0.00"
    ws.Columns("F:F").NumberFormat = "#,##0.00"
    ws.Columns("I:I").NumberFormat = "#,##0.00"
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 2
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D:E").ColumnWidth = 2
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G:H").ColumnWidth = 2
    ws.Columns("I").ColumnWidth = 14
    
    ' Center align headers
    ws.Range("A1:G4").HorizontalAlignment = xlCenter
    
    ' Right align amount columns
    ws.Columns("C:G").HorizontalAlignment = xlRight
    
    ' Add borders
    ' ws.Range(ws.Cells(6, 3), ws.Cells(ws.Cells(ws.Rows.Count, 1).End(xlUp).row, 7)).Borders.LineStyle = xlContinuous
End Sub
