Public Type LongTermLoanCurrentPortion
    CurrentYear As Double
    PreviousYear As Double
End Type

Public gLoanCurrentPortion As LongTermLoanCurrentPortion

Sub GenerateBalanceSheet()
    Dim ws As Worksheet
    Dim trialBalanceSheets As Collection
    Dim trialPLSheets As Collection
    
    ' Get the TrialBalance and TrialPL sheets
    Set trialBalanceSheets = GetWorksheetsWithPrefix("Trial Balance")
    Set trialPLSheets = GetWorksheetsWithPrefix("Trial PL")
    
    ' Check the number of TrialBalance and TrialPL sheets
    If trialBalanceSheets.Count = 1 And trialPLSheets.Count = 1 Then
        CreateFirstYearBalanceSheet trialBalanceSheets(1)
    ElseIf trialBalanceSheets.Count = 2 And trialPLSheets.Count = 2 Then
        CreateMultiPeriodBalanceSheet trialBalanceSheets
    Else
        MsgBox "Invalid number of TrialBalance or TrialPL sheets. Please ensure there are either one or two sheets of each type.", vbExclamation
    End If
End Sub

Sub CreateFirstYearBalanceSheet(trialBalanceSheet As Worksheet)
    Dim wsAsset As Worksheet
    Dim wsLiability As Worksheet
    Dim targetWorkbook As Workbook
    
    Set targetWorkbook = trialBalanceSheet.Parent
    
    ' Create new sheets for the balance sheet
    Set wsAsset = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    wsAsset.Name = "ABS"
    Set wsLiability = targetWorkbook.Sheets.Add(After:=wsAsset)
    wsLiability.Name = "LBS"
    
    ' Asset Side
    CreateAssetBalanceSheet wsAsset, trialBalanceSheet
    
    ' Liability Side
    CreateLiabilityBalanceSheet wsLiability, trialBalanceSheet
    
    ' Format both worksheets
    FormatWorksheet wsAsset
    FormatWorksheet wsLiability
End Sub

Sub CreateMultiPeriodBalanceSheet(trialBalanceSheets As Collection)
    Dim wsAsset As Worksheet
    Dim wsLiability As Worksheet
    Dim targetWorkbook As Workbook
    
    Set targetWorkbook = trialBalanceSheets(1).Parent
    
    ' Create new sheets for the balance sheet with corrected names
    Set wsAsset = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    wsAsset.Name = "MPA"
    Set wsLiability = targetWorkbook.Sheets.Add(After:=wsAsset)
    wsLiability.Name = "MPL"
    
    ' Asset Side
    CreateMultiPeriodAssetBalanceSheet wsAsset, trialBalanceSheets
    
    ' Liability Side
    CreateMultiPeriodLiabilityBalanceSheet wsLiability, trialBalanceSheets
    
    ' Format both worksheets
    FormatWorksheet wsAsset
    FormatWorksheet wsLiability
End Sub

Sub CreateAssetBalanceSheet(ws As Worksheet, trialBalanceSheet As Worksheet)
    Dim row As Long
    Dim currentAssetsStartRow As Long
    Dim nonCurrentAssetsStartRow As Long
    Dim year As Variant
    Dim tempResults As Variant
    
    ' Get the financial year
    year = GetFinancialYears(ws)
    If IsError(year) Then
        MsgBox "Failed to get financial year: " & year, vbExclamation
        Exit Sub
    End If

    ' Create the header
    CreateHeader ws, "Balance Sheet"

    ' Add details
    row = 5 ' Start below the header
    AddBalanceSheetHeaderDetails ws, row
    row = row + 1

    ' Add "สินทรัพย์"
    With ws.Cells(row, 2)
        .Value = "สินทรัพย์"
        .Font.Bold = True
    End With
    row = row + 1

    ' Add "สินทรัพย์หมุนเวียน" and year
    With ws.Cells(row, 2)
        .Value = "สินทรัพย์หมุนเวียน"
        .Font.Bold = True
    End With
    With ws.Cells(row, 9)
        .Value = year
        .Font.Underline = xlUnderlineStyleSingle
    End With
    row = row + 1

    ' Store current assets start row
    currentAssetsStartRow = row
    
    ' Create collection of trial balance sheets
    Dim trialBalanceSheets As Collection
    Set trialBalanceSheets = CreateTrialBalanceCollection(trialBalanceSheet)

    ' Add current assets
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินสดและรายการเทียบเท่าเงินสด", "1010", "1099", "", True, False, False, False)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น", "1140", "1299", "1141", True, False, False, False)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินให้กู้ยืมระยะสั้น", "1141", "1141", "", True, True, False, True)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "สินค้าคงเหลือ", "1510", "1530", "", True, False, False, False)

    ' Add total current assets with formula
    With ws.Cells(row, 2)
        .Value = "รวมสินทรัพย์หมุนเวียน"
        .Font.Bold = True
    End With
    Dim currentAssetsRow As Long
    currentAssetsRow = row ' Store this row for later reference
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(currentAssetsStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add "สินทรัพย์ไม่หมุนเวียน"
    With ws.Cells(row, 2)
        .Value = "สินทรัพย์ไม่หมุนเวียน"
        .Font.Bold = True
    End With
    row = row + 1

    ' Store non-current assets start row
    nonCurrentAssetsStartRow = row

    ' Add non-current assets
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "ที่ดิน อาคารและอุปกรณ์", "1600", "1659", "", True, False, False, False)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "สินทรัพย์ไม่หมุนเวียนอื่น", "1660", "1700", "", True, False, False, False)

    ' Add total non-current assets with formula
    With ws.Cells(row, 2)
        .Value = "รวมสินทรัพย์ไม่หมุนเวียน"
        .Font.Bold = True
    End With
    Dim nonCurrentAssetsRow As Long
    nonCurrentAssetsRow = row ' Store this row for later reference
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(nonCurrentAssetsStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add total assets with formula
    With ws.Cells(row, 2)
        .Value = "รวมสินทรัพย์"
        .Font.Bold = True
    End With
    With ws.Cells(row, 9)
        .Formula = "=" & ws.Cells(currentAssetsRow, 9).Address & "+" & ws.Cells(nonCurrentAssetsRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
End Sub


Sub CreateLiabilityBalanceSheet(ws As Worksheet, trialBalanceSheet As Worksheet)
    Dim row As Long
    Dim currentLiabilitiesStartRow As Long
    Dim nonCurrentLiabilitiesStartRow As Long
    Dim equityStartRow As Long
    Dim partnersEquity As Collection
    Dim equityItem As Variant
    Dim tempResults As Variant
    Dim isLimitedPartnership As Boolean
    Dim liabilityAndEquityTerm As String
    Dim equityTerm As String
    Dim targetWorkbook As Workbook
    Dim paidUpCapitalStartRow As Long
    
    Set targetWorkbook = ws.Parent
    
    ' Check if it's a limited partnership
    isLimitedPartnership = (targetWorkbook.Sheets("Info").Range("B2").Value = "ห้างหุ้นส่วนจำกัด")
    
    ' Set terms based on company type
    If isLimitedPartnership Then
        liabilityAndEquityTerm = "หนี้สินและส่วนของผู้เป็นหุ้นส่วน"
        equityTerm = "ส่วนของผู้เป็นหุ้นส่วน"
    Else
        liabilityAndEquityTerm = "หนี้สินและส่วนของผู้ถือหุ้น"
        equityTerm = "ส่วนของผู้ถือหุ้น"
    End If
    
    ' Initial setup
    CreateHeader ws, "Balance Sheet"
    row = 5
    AddBalanceSheetHeaderDetails ws, row
    row = row + 1

    ' Add main headers
    With ws.Cells(row, 2)
        .Value = liabilityAndEquityTerm
        .Font.Bold = True
    End With
    row = row + 1

    With ws.Cells(row, 2)
        .Value = "หนี้สินหมุนเวียน"
        .Font.Bold = True
    End With
    row = row + 1

    ' Store current liabilities start row
    currentLiabilitiesStartRow = row

    ' Create collection of trial balance sheets
    Dim trialBalanceSheets As Collection
    Set trialBalanceSheets = CreateTrialBalanceCollection(trialBalanceSheet)

    ' Add current liabilities
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน", "2001", "2009", "", True, False, False, False)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น", "2010", "2999", "2030,2045,2050,2051,2052,2100,2120,2121,2122,2123", True, False, False, False)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "ส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี", "0", "0", "", True, True, False, True)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินกู้ยืมระยะสั้น", "2030", "2030", "", True, True, False, True)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "ภาษีเงินได้นิติบุคคลค้างจ่าย", "2045", "2045", "", True, True, False, True)

    ' Add total current liabilities with formula
    With ws.Cells(row, 2)
        .Value = "รวมหนี้สินหมุนเวียน"
        .Font.Bold = True
    End With
    Dim currentLiabilitiesRow As Long
    currentLiabilitiesRow = row ' Store this row for later reference
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(currentLiabilitiesStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add non-current liabilities header
    With ws.Cells(row, 2)
        .Value = "หนี้สินไม่หมุนเวียน"
        .Font.Bold = True
    End With
    row = row + 1

    ' Store non-current liabilities start row
    nonCurrentLiabilitiesStartRow = row

    ' Add non-current liabilities
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินกู้ยืมระยะยาวจากสถาบันการเงิน", "2120", "2123", "2121", True, False, False, False)
    ' Subtract current portion
    With ws.Cells(row - 1, 9)
        .Value = .Value - gLoanCurrentPortion.CurrentYear
    End With
    
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินกู้ยืมระยะยาว", "2050", "2052", "", True, False, True, False)

    ' Add total non-current liabilities with formula
    With ws.Cells(row, 2)
        .Value = "รวมหนี้สินไม่หมุนเวียน"
        .Font.Bold = True
    End With
    Dim nonCurrentLiabilitiesRow As Long
    nonCurrentLiabilitiesRow = row ' Store this row for later reference
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(nonCurrentLiabilitiesStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add total liabilities with formula
    With ws.Cells(row, 2)
        .Value = "รวมหนี้สิน"
        .Font.Bold = True
    End With
    With ws.Cells(row, 9)
        .Formula = "=" & ws.Cells(currentLiabilitiesRow, 9).Address & "+" & ws.Cells(nonCurrentLiabilitiesRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    row = row + 1

    ' Add equity section header
    With ws.Cells(row, 2)
        .Value = equityTerm
        .Font.Bold = True
    End With
    row = row + 1

    ' Store equity start row
    equityStartRow = row

    ' Add equity accounts
    If isLimitedPartnership Then
        Set partnersEquity = GetPartnersEquity(targetWorkbook)
        paidUpCapitalStartRow = row
        For Each equityItem In partnersEquity
            ws.Cells(row, 3).Value = equityItem(0) ' Account name
            ws.Cells(row, 9).Value = equityItem(1) ' Amount
            row = row + 1
        Next equityItem
    Else
        ' Add registered capital section
        ws.Cells(row, 3).Value = "ทุนจดทะเบียน"
        row = row + 1
        
        Dim shares As Long
        Dim shareValue As Double
        shares = CLng(targetWorkbook.Sheets("Info").Range("B6").Value)
        shareValue = CDbl(targetWorkbook.Sheets("Info").Range("B7").Value)
        
        ws.Cells(row, 4).Value = "หุ้นสามัญ " & Format(shares, "#,##0") & " หุ้น มูลค่าหุ้นละ " & Format(shareValue, "#,##0.00") & " บาท"
        ws.Cells(row, 9).Value = shares * shareValue
        With ws.Cells(row, 9)
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        row = row + 1
        
        ' Add paid-up capital section
        ws.Cells(row, 3).Value = "ทุนที่ออกและชำระแล้ว"
        row = row + 1
        ws.Cells(row, 4).Value = "หุ้นสามัญ " & Format(shares, "#,##0") & " หุ้น มูลค่าหุ้นละ " & Format(shareValue, "#,##0.00") & " บาท"
        ' Store the row where paid-up capital starts for equity calculation
        
        paidUpCapitalStartRow = row
        tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "", "3010", "3019", "", False, True, False, False)
    End If

    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, GetRetainedEarningsText(targetWorkbook), "3020", "3020", "", False, True, False, True)

    ' Add total equity with formula
    With ws.Cells(row, 2)
        .Value = "รวม" & equityTerm
        .Font.Bold = True
    End With
    Dim equityTotalRow As Long
    equityTotalRow = row
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(paidUpCapitalStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add total liabilities and equity with formula
    With ws.Cells(row, 2)
        .Value = "รวม" & liabilityAndEquityTerm
        .Font.Bold = True
    End With
    With ws.Cells(row, 9)
        .Formula = "=" & ws.Cells(currentLiabilitiesRow + 2, 9).Address & "+" & ws.Cells(equityTotalRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
End Sub


Sub CreateMultiPeriodAssetBalanceSheet(ws As Worksheet, trialBalanceSheets As Collection)
    Dim row As Long
    Dim currentAssetsStartRow As Long
    Dim nonCurrentAssetsStartRow As Long
    Dim years As Variant
    Dim targetWorkbook As Workbook
    Dim tempResults As Variant

    ' Get the target workbook
    Set targetWorkbook = ws.Parent

    ' Get financial years
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
        Exit Sub
    End If
    
    ' Create header and details
    CreateHeader ws, "Balance Sheet"
    row = 5
    AddBalanceSheetHeaderDetails ws, row
    row = row + 1

    ' Add "สินทรัพย์"
    With ws.Cells(row, 2)
        .Value = "สินทรัพย์"
        .Font.Bold = True
    End With
    row = row + 1

    ' Add headers and years
    With ws.Cells(row, 2)
        .Value = "สินทรัพย์หมุนเวียน"
        .Font.Bold = True
    End With
    With ws.Cells(row, 7)
        .Value = years(1)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    With ws.Cells(row, 9)
        .Value = years(2)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    row = row + 1

    ' Store current assets start row
    currentAssetsStartRow = row

    ' Add current assets
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินสดและรายการเทียบเท่าเงินสด", "1010", "1099", "", True, False, True, False)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น", "1140", "1299", "1141", True, False, True, False)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินให้กู้ยืมระยะสั้น", "1141", "1141", "", True, True, True, True)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "สินค้าคงเหลือ", "1510", "1530", "", True, False, True, False)

    ' Add total current assets with formulas
    With ws.Cells(row, 2)
        .Value = "รวมสินทรัพย์หมุนเวียน"
        .Font.Bold = True
    End With
    Dim currentAssetsRow As Long
    currentAssetsRow = row
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(currentAssetsStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(currentAssetsStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add "สินทรัพย์ไม่หมุนเวียน"
    With ws.Cells(row, 2)
        .Value = "สินทรัพย์ไม่หมุนเวียน"
        .Font.Bold = True
    End With
    row = row + 1

    ' Store non-current assets start row
    nonCurrentAssetsStartRow = row

    ' Add non-current assets
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "ที่ดิน อาคารและอุปกรณ์", "1600", "1659", "", True, False, True, False)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "สินทรัพย์ไม่หมุนเวียนอื่น", "1660", "1700", "", True, False, True, False)

    ' Add total non-current assets with formulas
    With ws.Cells(row, 2)
        .Value = "รวมสินทรัพย์ไม่หมุนเวียน"
        .Font.Bold = True
    End With
    Dim nonCurrentAssetsRow As Long
    nonCurrentAssetsRow = row
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(nonCurrentAssetsStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(nonCurrentAssetsStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add total assets with formulas
    With ws.Cells(row, 2)
        .Value = "รวมสินทรัพย์"
        .Font.Bold = True
    End With
    With ws.Cells(row, 7)
        .Formula = "=" & ws.Cells(currentAssetsRow, 7).Address & "+" & ws.Cells(nonCurrentAssetsRow, 7).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    With ws.Cells(row, 9)
        .Formula = "=" & ws.Cells(currentAssetsRow, 9).Address & "+" & ws.Cells(nonCurrentAssetsRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
End Sub

Sub CreateMultiPeriodLiabilityBalanceSheet(ws As Worksheet, trialBalanceSheets As Collection)
    Dim row As Long
    Dim currentLiabilitiesStartRow As Long
    Dim nonCurrentLiabilitiesStartRow As Long
    Dim equityStartRow As Long
    Dim partnersEquity As Collection
    Dim equityItem As Variant
    Dim tempResults As Variant
    Dim isLimitedPartnership As Boolean
    Dim liabilityAndEquityTerm As String
    Dim equityTerm As String
    Dim targetWorkbook As Workbook
    Dim years As Variant
    
    Set targetWorkbook = ws.Parent
    
    ' Get financial years
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
        Exit Sub
    End If
    
    ' Check if it's a limited partnership
    isLimitedPartnership = (targetWorkbook.Sheets("Info").Range("B2").Value = "ห้างหุ้นส่วนจำกัด")
    
    ' Set terms based on company type
    If isLimitedPartnership Then
        liabilityAndEquityTerm = "หนี้สินและส่วนของผู้เป็นหุ้นส่วน"
        equityTerm = "ส่วนของผู้เป็นหุ้นส่วน"
    Else
        liabilityAndEquityTerm = "หนี้สินและส่วนของผู้ถือหุ้น"
        equityTerm = "ส่วนของผู้ถือหุ้น"
    End If
    
    ' Initial setup
    CreateHeader ws, "Balance Sheet"
    row = 5
    AddBalanceSheetHeaderDetails ws, row
    row = row + 1

    ' Add main headers with years
    With ws.Cells(row, 2)
        .Value = liabilityAndEquityTerm
        .Font.Bold = True
    End With
    row = row + 1

    With ws.Cells(row, 2)
        .Value = "หนี้สินหมุนเวียน"
        .Font.Bold = True
    End With
    With ws.Cells(row, 7)
        .Value = years(1)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    With ws.Cells(row, 9)
        .Value = years(2)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    row = row + 1

    ' Store current liabilities start row
    currentLiabilitiesStartRow = row

    ' Add current liabilities
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน", "2001", "2009", "", True, False, True, False)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น", "2010", "2999", "2030,2045,2050,2051,2052,2100,2120,2121,2122,2123", True, False, True, False)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "ส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี", "0", "0", "", True, True, True, True)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินกู้ยืมระยะสั้น", "2030", "2030", "", True, True, True, True)
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "ภาษีเงินได้นิติบุคคลค้างจ่าย", "2045", "2045", "", True, True, True, True)

    ' Add total current liabilities with formulas
    With ws.Cells(row, 2)
        .Value = "รวมหนี้สินหมุนเวียน"
        .Font.Bold = True
    End With
    Dim currentLiabilitiesRow As Long
    currentLiabilitiesRow = row
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(currentLiabilitiesStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(currentLiabilitiesStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add non-current liabilities header
    With ws.Cells(row, 2)
        .Value = "หนี้สินไม่หมุนเวียน"
        .Font.Bold = True
    End With
    row = row + 1

    ' Store non-current liabilities start row
    nonCurrentLiabilitiesStartRow = row

    ' Add non-current liabilities
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินกู้ยืมระยะยาวจากสถาบันการเงิน", "2120", "2123", "2121", True, False, True, False)
    ' Subtract current portion
    With ws.Cells(row - 1, 7)
        .Value = .Value - gLoanCurrentPortion.CurrentYear
    End With
    With ws.Cells(row - 1, 9)
        .Value = .Value - gLoanCurrentPortion.PreviousYear
    End With
    
    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "เงินกู้ยืมระยะยาว", "2050", "2052", "", True, False, True, False)

    ' Add total non-current liabilities with formulas
    With ws.Cells(row, 2)
        .Value = "รวมหนี้สินไม่หมุนเวียน"
        .Font.Bold = True
    End With
    Dim nonCurrentLiabilitiesRow As Long
    nonCurrentLiabilitiesRow = row
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(nonCurrentLiabilitiesStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(nonCurrentLiabilitiesStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add total liabilities with formulas
    With ws.Cells(row, 2)
        .Value = "รวมหนี้สิน"
        .Font.Bold = True
    End With
    Dim totalLiabilitiesRow As Long
    totalLiabilitiesRow = row
    With ws.Cells(row, 7)
        .Formula = "=" & ws.Cells(currentLiabilitiesRow, 7).Address & "+" & ws.Cells(nonCurrentLiabilitiesRow, 7).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    With ws.Cells(row, 9)
        .Formula = "=" & ws.Cells(currentLiabilitiesRow, 9).Address & "+" & ws.Cells(nonCurrentLiabilitiesRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    row = row + 1

    ' Add equity section header
    With ws.Cells(row, 2)
        .Value = equityTerm
        .Font.Bold = True
    End With
    row = row + 1

    ' Store equity start row
    equityStartRow = row
    
    ' Declare paidUpCapitalStartRow here to use in both cases
    Dim paidUpCapitalStartRow As Long

    ' Add equity accounts
    If isLimitedPartnership Then
        ' For limited partnership, use equityStartRow
        paidUpCapitalStartRow = equityStartRow
        
        Set partnersEquity = GetPartnersEquity(targetWorkbook)
        For Each equityItem In partnersEquity
            ws.Cells(row, 3).Value = equityItem(0) ' Account name
            ws.Cells(row, 7).Value = equityItem(1) ' Current year
            ws.Cells(row, 9).Value = equityItem(1) ' Previous year (same as current)
            row = row + 1
        Next equityItem
    Else
        ' Add registered capital section
        ws.Cells(row, 3).Value = "ทุนจดทะเบียน"
        row = row + 1
        
        Dim shares As Long
        Dim shareValue As Double
        shares = CLng(targetWorkbook.Sheets("Info").Range("B6").Value)
        shareValue = CDbl(targetWorkbook.Sheets("Info").Range("B7").Value)
        
        ws.Cells(row, 4).Value = "หุ้นสามัญ " & Format(shares, "#,##0") & " หุ้น มูลค่าหุ้นละ " & Format(shareValue, "#,##0.00") & " บาท"
        ws.Cells(row, 7).Value = shares * shareValue
        ws.Cells(row, 9).Value = shares * shareValue
        With ws.Cells(row, 7)
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        With ws.Cells(row, 9)
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        row = row + 1
        
        ' Add paid-up capital section
        ws.Cells(row, 3).Value = "ทุนที่ออกและชำระแล้ว"
        row = row + 1
        ws.Cells(row, 4).Value = "หุ้นสามัญ " & Format(shares, "#,##0") & " หุ้น มูลค่าหุ้นละ " & Format(shareValue, "#,##0.00") & " บาท"
        paidUpCapitalStartRow = row
        tempResults = AddAccountGroup(ws, trialBalanceSheets, row, "", "3010", "3019", "", False, True, True, False)
    End If

    tempResults = AddAccountGroup(ws, trialBalanceSheets, row, GetRetainedEarningsText(targetWorkbook), "3020", "3020", "", False, True, True, True)

    ' Add total equity with formulas
    With ws.Cells(row, 2)
        .Value = "รวม" & equityTerm
        .Font.Bold = True
    End With
    Dim equityTotalRow As Long
    equityTotalRow = row
    
    ' Use paidUpCapitalStartRow which is set appropriately for both cases
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(paidUpCapitalStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(paidUpCapitalStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add total liabilities and equity with formulas
    With ws.Cells(row, 2)
        .Value = "รวม" & liabilityAndEquityTerm
        .Font.Bold = True
    End With
    With ws.Cells(row, 7)
        .Formula = "=" & ws.Cells(totalLiabilitiesRow, 7).Address & "+" & ws.Cells(equityTotalRow, 7).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    With ws.Cells(row, 9)
        .Formula = "=" & ws.Cells(totalLiabilitiesRow, 9).Address & "+" & ws.Cells(equityTotalRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
End Sub
Function AddAccountGroup(ws As Worksheet, trialBalanceSheets As Collection, ByRef row As Long, groupName As String, startCode As String, endCode As String, Optional excludeCode As String = "", Optional includeDecimal As Boolean = False, Optional listIndividualAccounts As Boolean = False, Optional isMultiPeriod As Boolean = True, Optional isSingleAccount As Boolean = False) As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim totalAmount As Double
    Dim totalAmountPrevious As Double
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    Dim amountPrevious As Double
    Dim result(1 To 2) As Double
    Dim foundMatch As Boolean
    
    ' Special handling for current portion of long-term loans
    If groupName = "ส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี" Then
        ws.Cells(row, 3).Value = groupName
        If isMultiPeriod Then
            ws.Cells(row, 7).Value = gLoanCurrentPortion.CurrentYear
            ws.Cells(row, 9).Value = gLoanCurrentPortion.PreviousYear
            result(1) = gLoanCurrentPortion.CurrentYear
            result(2) = gLoanCurrentPortion.PreviousYear
        Else
            ws.Cells(row, 9).Value = gLoanCurrentPortion.CurrentYear
            result(1) = gLoanCurrentPortion.CurrentYear
            result(2) = 0
        End If
        row = row + 1
        AddAccountGroup = result
        Exit Function
    End If
    
    lastRow = trialBalanceSheets(1).Cells(trialBalanceSheets(1).Rows.Count, 2).End(xlUp).row
    totalAmount = 0
    totalAmountPrevious = 0
    foundMatch = False
    
    If groupName <> "" Then
        ws.Cells(row, 3).Value = groupName
    End If
    
    For i = 2 To lastRow
        accountCode = trialBalanceSheets(1).Cells(i, 2).Value
        If (isSingleAccount And accountCode = startCode) Or _
           (Not isSingleAccount And accountCode >= startCode And accountCode <= endCode) Then
            If InStr(1, excludeCode, accountCode) = 0 And (includeDecimal Or InStr(1, accountCode, ".") = 0) Then
                foundMatch = True
                accountName = trialBalanceSheets(1).Cells(i, 1).Value
                amount = trialBalanceSheets(1).Cells(i, 6).Value
                totalAmount = totalAmount + amount
                If isMultiPeriod Then
                    amountPrevious = GetAmountFromPreviousPeriod(trialBalanceSheets(2), accountCode)
                    totalAmountPrevious = totalAmountPrevious + amountPrevious
                End If
                If listIndividualAccounts Or isSingleAccount Then
                    ws.Cells(row, 3).Value = groupName
                    If isMultiPeriod Then
                        ws.Cells(row, 7).Value = amount
                        ws.Cells(row, 9).Value = amountPrevious
                    Else
                        ws.Cells(row, 9).Value = amount
                    End If
                    row = row + 1
                End If
                If isSingleAccount Then Exit For
            End If
        End If
    Next i
    
    ' Handle case where no matching account codes were found
    If Not foundMatch Then
        If listIndividualAccounts Or isSingleAccount Then
            ws.Cells(row, 3).Value = groupName
            If isMultiPeriod Then
                ws.Cells(row, 7).Value = 0
                ws.Cells(row, 9).Value = 0
            Else
                ws.Cells(row, 9).Value = 0
            End If
            row = row + 1
        End If
    End If
    
    If Not listIndividualAccounts And Not isSingleAccount Then
        If isMultiPeriod Then
            ws.Cells(row, 7).Value = totalAmount
            ws.Cells(row, 9).Value = totalAmountPrevious
        Else
            ws.Cells(row, 9).Value = totalAmount
        End If
        row = row + 1
    End If
    
    ' Set the result values
    result(1) = IIf(foundMatch, totalAmount, 0)  ' Current year amount
    result(2) = IIf(isMultiPeriod, totalAmountPrevious, 0)  ' Previous year amount
    
    AddAccountGroup = result
End Function

Sub FormatWorksheet(ws As Worksheet)
    ' Apply Thai Sarabun font and font size 14 to the worksheet
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Set number format to use comma style for columns G and I
    ws.Columns("G").NumberFormatLocal = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    ws.Columns("I").NumberFormatLocal = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
    ' Format row 7 (years) separately
    With ws.Range("G7:I7")
        .NumberFormat = "General"
        .HorizontalAlignment = xlCenter
    End With
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 7
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D:F").ColumnWidth = 7
    ws.Columns("E").ColumnWidth = 28
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 2
    ws.Columns("I").ColumnWidth = 14
    
    ' Center align headers
    ws.Range("A1:I4").HorizontalAlignment = xlCenter
    
End Sub

Sub AddBalanceSheetHeaderDetails(ws As Worksheet, row As Long)
    With ws.Cells(row, 6)
        .Value = "หมายเหตุ"
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    With ws.Cells(row, 9)
        .Value = "หน่วย:บาท"
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .HorizontalAlignment = xlRight
    End With
End Sub

Function isLimitedPartnership(targetWorkbook As Workbook) As Boolean
    Dim infoSheet As Worksheet
    Set infoSheet = targetWorkbook.Sheets("Info")
    isLimitedPartnership = (infoSheet.Range("B2").Value = "ห้างหุ้นส่วนจำกัด")
End Function
Function GetRetainedEarningsText(targetWorkbook As Workbook) As String
    If isLimitedPartnership(targetWorkbook) Then
        GetRetainedEarningsText = "กำไร ( ขาดทุน ) สะสมยังไม่ได้แบ่ง"
    Else
        GetRetainedEarningsText = "กำไร ( ขาดทุน ) สะสมยังไม่ได้จัดสรร"
    End If
End Function


Function GetPartnersEquity(targetWorkbook As Workbook) As Collection
    Dim entityNumber As String
    Dim entityFilePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String
    Dim equityData As New Collection
    Dim accountName As String
    Dim amount As Double
    
    ' Get the entity number from the correct location in the target workbook
    entityNumber = targetWorkbook.Sheets("Info").Range("B4").Value
    
    ' Construct the file path using the target workbook's path
    entityFilePath = targetWorkbook.Path & "\ExtractWebDBD\" & entityNumber & ".csv"
    
    ' Check if the file exists
    If Dir(entityFilePath) = "" Then
        MsgBox "Entity file not found: " & entityFilePath, vbExclamation
        Exit Function
    End If
    
    ' Read the CSV file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(entityFilePath, 1)
    
    ' Skip the header
    ts.SkipLine
    
    ' Read data
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        parts = Split(line, ",")
        
        ' Check if we have enough columns for the account name
        If UBound(parts) >= 8 Then
            accountName = Trim(parts(8))
            If accountName <> "" And Left(accountName, 5) <> "ลงทุน" Then
                ' Read the next line for the amount
                If Not ts.AtEndOfStream Then
                    line = ts.ReadLine
                    parts = Split(line, ",")
                    If UBound(parts) >= 9 Then
                        If IsNumeric(Trim(parts(9))) Then
                            amount = CDbl(Trim(parts(9)))
                            equityData.Add Array(accountName, amount)
                        End If
                    End If
                End If
            End If
        End If
    Loop
    
    ts.Close
    Set GetPartnersEquity = equityData
End Function

Function CreateTrialBalanceCollection(trialBalanceSheet As Worksheet) As Collection
    Dim col As New Collection
    col.Add trialBalanceSheet
    Set CreateTrialBalanceCollection = col
End Function
