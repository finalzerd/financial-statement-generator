
Sub CreateSingleYearNotes(ws As Worksheet, trialBalanceSheet As Worksheet, trialPLSheet As Worksheet)
    ' Call CreateNote(ws, trialBalanceSheet, "เงินสดและรายการเทียบเท่าเงินสด", "1000", "1099")
    Call CreateFirstYearCashNote(ws, trialBalanceSheet)
    Call CreateNote(ws, trialBalanceSheet, "ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น", "1140", "1215", "1141")
    Call CreateNote(ws, trialBalanceSheet, "เงินให้กู้ยืมระยะสั้น", "1141", "1141")
    Call CreateFirstYearNoteForLandBuildingEquipment(ws, trialBalanceSheet)  ' New call for PPE note 1600-1659
    Call CreateNote(ws, trialBalanceSheet, "สินทรัพย์อื่น", "1660", "1700")
    Call CreateNote(ws, trialBalanceSheet, "เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน", "2001", "2009")
    Call CreateNote(ws, trialBalanceSheet, "เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น", "2010", "2999", "2030,2045,2050,2051,2052,2100,2120,2121,2122,2123")
    Call CreateNote(ws, trialBalanceSheet, "เงินกู้ยืมระยะสั้นจากบุคคลหรือกิจการที่เกี่ยวข้องกัน", "2030", "2030")
    ' Call CreateNote(ws, trialBalanceSheet, "เงินกู้ยืมระยะยาวจากสถาบันการเงิน", "2120", "2123", "2121")
    Call CreateFirstYearLongTermLoansNote(ws, trialBalanceSheet)
    Call CreateNote(ws, trialBalanceSheet, "เงินกู้ยืมระยะยาว", "2050", "2052")
    Call CreateNote(ws, trialBalanceSheet, "เงินกู้ยืมระยะยาวจากบุคคลหรือกิจการที่เกี่ยวข้องกัน", "2100", "2100")
    Call CreateNote(ws, trialPLSheet, "รายได้อื่น", "4020", "4999")
    Call CreateExpensesByNatureNote(ws)
    Call CreateFinancialApprovalNote(ws)
End Sub

Sub CreateMultiYearNotes(ws As Worksheet, trialBalanceSheets As Collection, trialPLSheets As Collection)
    ' Call CreateMultiPeriodNote(ws, trialBalanceSheets, trialPLSheets, "เงินสดและรายการเทียบเท่าเงินสด", "1000", "1099")
    Call CreateCashNote(ws, trialBalanceSheets)
    Call CreateMultiPeriodNote(ws, trialBalanceSheets, trialPLSheets, "ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น", "1140", "1215", "1141")
    Call CreateMultiPeriodNote(ws, trialBalanceSheets, trialPLSheets, "เงินให้กู้ยืมระยะสั้น", "1141", "1141")
    Call CreateNoteForLandBuildingEquipment(ws, trialBalanceSheets) ' New call to create ที่ดิน อาคารและอุปกรณ์ note
    Call CreateMultiPeriodNote(ws, trialBalanceSheets, trialPLSheets, "สินทรัพย์อื่น", "1660", "1700")
    Call CreateMultiPeriodNote(ws, trialBalanceSheets, trialPLSheets, "เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน", "2001", "2009")
    Call CreateMultiPeriodNote(ws, trialBalanceSheets, trialPLSheets, "เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น", "2010", "2999", "2030,2045,2050,2051,2052,2100,2120,2121,2122,2123")
    Call CreateMultiPeriodNote(ws, trialBalanceSheets, trialPLSheets, "เงินกู้ยืมระยะสั้นจากบุคคลหรือกิจการที่เกี่ยวข้องกัน", "2030", "2030")
    ' Call CreateMultiPeriodNote(ws, trialBalanceSheets, trialPLSheets, "เงินกู้ยืมระยะยาวจากสถาบันการเงิน", "2120", "2123", "2121")
    Call CreateLongTermLoansNote(ws, trialBalanceSheets, trialPLSheets)
    Call CreateMultiPeriodNote(ws, trialBalanceSheets, trialPLSheets, "เงินกู้ยืมระยะยาว", "2050", "2052")
    Call CreateMultiPeriodNote(ws, trialBalanceSheets, trialPLSheets, "เงินกู้ยืมระยะยาวจากบุคคลหรือกิจการที่เกี่ยวข้องกัน", "2100", "2100")
    Call CreateMultiPeriodNote(ws, trialPLSheets, trialPLSheets, "รายได้อื่น", "4020", "4999")
    Call CreateExpensesByNatureNote(ws)
    Call CreateFinancialApprovalNote(ws)
End Sub

Function CreateNote(ws As Worksheet, trialSheet As Worksheet, noteName As String, accountCodeStart As String, accountCodeEnd As String, Optional excludeAccountCodes As String = "") As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim totalAmount As Double
    Dim noteStartRow As Long
    Static noteOrder As Integer
    Dim infoSheet As Worksheet
    Dim year As Variant
    
    ' Get the financial year using the new function
    year = GetFinancialYears(ws)
    If IsError(year) Then
        MsgBox "Failed to get financial year: " & year, vbExclamation
    End If

    ' Initialize noteCreated to False
    noteCreated = False

    ' Initialize totalAmount to zero
    totalAmount = 0

    ' Get the last row of data in the trial sheet
    lastRow = trialSheet.Cells(trialSheet.Rows.Count, 1).End(xlUp).row

    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow

    ' Increment the note order
    noteOrder = noteOrder + 1

    ' Create the note header
    ws.Cells(noteRow, 1).Value = noteOrder + 2  ' Start from 3
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = noteName
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    ws.Cells(noteRow + 1, 9).Value = year
    noteRow = noteRow + 2

    ' Loop through the trial data and check for account codes
    For i = 2 To lastRow
        accountCode = trialSheet.Cells(i, 2).Value
        accountName = trialSheet.Cells(i, 1).Value
        amount = trialSheet.Cells(i, 6).Value

        If accountCode >= accountCodeStart And accountCode <= accountCodeEnd Then
            If InStr(1, excludeAccountCodes, accountCode) = 0 Then
                ' Account code is within the range and not in the exclude list
                totalAmount = totalAmount + amount

                ' Check if the amount is not zero or blank
                If amount <> 0 And Not IsEmpty(amount) Then
                    ' Add the account detail to the note
                    ws.Cells(noteRow, 3).Value = accountName
                    ws.Cells(noteRow, 9).Value = amount
                    noteRow = noteRow + 1

                    ' Set noteCreated to True
                    noteCreated = True
                End If
            End If
        End If
    Next i

    ' Check if any account details were added to the note
    If noteCreated Then
        ' Add the total amount to the note
        ws.Cells(noteRow, 3).Value = "รวม"
        ws.Cells(noteRow, 9).Value = totalAmount
        
        With ws.Cells(noteRow, 9)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        
        noteRow = noteRow + 1

        ' Add the "EndOfNote" mark to the final row of the note and color it white
        ws.Cells(noteRow, 1).Value = "EndOfNote"
        ws.Cells(noteRow, 1).Font.Color = vbWhite
    Else
        ' If no account details were added, remove the note header
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        noteOrder = noteOrder - 1  ' Decrement the note order if note was not created
    End If

    If noteRow > 34 Then
        ' Move the note to a new worksheet and recreate it
        Set ws = HandleNoteExceedingRow34(ws, noteName, noteStartRow, noteRow, trialSheet)
        noteCreated = True ' Set noteCreated to True since the note is recreated in the new worksheet
    End If

    ' Format the note
    FormatNote ws, noteStartRow, noteRow

    ' Return the value of noteCreated
    CreateNote = noteCreated
End Function

Function CreateMultiPeriodNote(ws As Worksheet, trialBalanceSheets As Collection, trialPLSheets As Collection, noteName As String, accountCodeStart As String, accountCodeEnd As String, Optional excludeAccountCodes As String = "") As Boolean
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long
    Dim accountCode As String
    Dim accountName As String
    Dim amount1 As Double, amount2 As Double
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim totalAmount1 As Double, totalAmount2 As Double
    Dim noteStartRow As Long
    Dim uniqueAccountCodes As New Collection
    Static noteOrder As Integer
    Dim infoSheet As Worksheet
    Dim years As Variant
    Dim targetWorkbook As Workbook
    
    ' Get the target workbook
    Set targetWorkbook = ws.Parent
    
    ' Get the financial years using the new function
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
    End If
    
    ' Initialize noteCreated to False
    noteCreated = False
    
    ' Initialize totalAmount1 and totalAmount2 to zero
    totalAmount1 = 0
    totalAmount2 = 0
    
    ' Get the last row of data in the trial sheets
    lastRow1 = trialBalanceSheets(1).Cells(trialBalanceSheets(1).Rows.Count, 1).End(xlUp).row
    lastRow2 = trialBalanceSheets(2).Cells(trialBalanceSheets(2).Rows.Count, 1).End(xlUp).row
    
    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow
    
    ' Increment the note order
    noteOrder = noteOrder + 1
    
    ' Create the note header
    ws.Cells(noteRow, 1).Value = noteOrder + 2  ' Start from 3
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = noteName
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    ws.Cells(noteRow + 1, 7).Value = years(1)
    ws.Cells(noteRow + 1, 9).Value = years(2)
    noteRow = noteRow + 2
    
    ' Loop through the trial data for both periods
    For i = 2 To lastRow1
        accountCode = trialBalanceSheets(1).Cells(i, 2).Value
        accountName = trialBalanceSheets(1).Cells(i, 1).Value
        amount1 = trialBalanceSheets(1).Cells(i, 6).Value
        
        If accountCode >= accountCodeStart And accountCode <= accountCodeEnd Then
            If InStr(1, excludeAccountCodes, accountCode) = 0 Then
                ' Account code is within the range and not in the exclude list
                If Not ContainsAccountCode(uniqueAccountCodes, accountCode) Then
                    ' Account code is unique, add it to the collection
                    uniqueAccountCodes.Add accountCode, CStr(accountCode)
                    
                    ' Check if the account code exists in the previous period
                    amount2 = GetAmountFromPreviousPeriod(trialBalanceSheets(2), accountCode)
                    
                    ' Only add the account detail if either amount is not zero or blank
                    If (amount1 <> 0 And Not IsEmpty(amount1)) Or (amount2 <> 0 And Not IsEmpty(amount2)) Then
                        ' Add the account detail to the note
                        ws.Cells(noteRow, 3).Value = accountName
                        ws.Cells(noteRow, 7).Value = amount1
                        ws.Cells(noteRow, 9).Value = amount2
                        noteRow = noteRow + 1
                        noteCreated = True
                    End If
                    
                    totalAmount1 = totalAmount1 + amount1
                    totalAmount2 = totalAmount2 + amount2
                End If
            End If
        End If
    Next i
    
    ' Loop through the trial data for the previous period to find any missing account codes
    For i = 2 To lastRow2
        accountCode = trialBalanceSheets(2).Cells(i, 2).Value
        accountName = trialBalanceSheets(2).Cells(i, 1).Value
        amount2 = trialBalanceSheets(2).Cells(i, 6).Value
        
        If accountCode >= accountCodeStart And accountCode <= accountCodeEnd Then
            If InStr(1, excludeAccountCodes, accountCode) = 0 Then
                ' Account code is within the range and not in the exclude list
                If Not ContainsAccountCode(uniqueAccountCodes, accountCode) Then
                    ' Account code is unique, add it to the collection
                    uniqueAccountCodes.Add accountCode, CStr(accountCode)
                    
                    ' Only add the account detail if the amount is not zero or blank
                    If amount2 <> 0 And Not IsEmpty(amount2) Then
                        ' Add the account detail to the note
                        ws.Cells(noteRow, 3).Value = accountName
                        ws.Cells(noteRow, 9).Value = amount2
                        noteRow = noteRow + 1
                        noteCreated = True
                    End If
                    
                    totalAmount2 = totalAmount2 + amount2
                End If
            End If
        End If
    Next i
    
    ' Check if any account details were added to the note
    If noteCreated Then
        ' Add the total amounts to the note
        ws.Cells(noteRow, 3).Value = "รวม"
        ws.Cells(noteRow, 7).Value = totalAmount1
        ws.Cells(noteRow, 9).Value = totalAmount2
        
        ' Add top and double bottom borders only to the amount cells (columns 7 and 9)
        With ws.Cells(noteRow, 7)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        With ws.Cells(noteRow, 9)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        
        noteRow = noteRow + 1
        
        ' Add the "EndOfNote" mark to the final row of the note and color it white
        ws.Cells(noteRow, 1).Value = "EndOfNote"
        ws.Cells(noteRow, 1).Font.Color = vbWhite
    Else
        ' If no account details were added, remove the note header
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        noteOrder = noteOrder - 1  ' Decrement the note order if note was not created
    End If
    
    If noteRow > 34 Then
        ' Move the note to a new worksheet and recreate it
        Set ws = HandleNoteExceedingRow34(ws, noteName, noteStartRow, noteRow, trialBalanceSheets(1))
        noteCreated = True ' Set noteCreated to True since the note is recreated in the new worksheet
    End If
    
    ' Format the note
    FormatNote ws, noteStartRow, noteRow
    
    ' Return the value of noteCreated
    CreateMultiPeriodNote = noteCreated
End Function

Function CreateFirstYearCashNote(ws As Worksheet, trialBalanceSheet As Worksheet) As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim accountCode As String
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim cashAmount As Double
    Dim bankAmount As Double
    Dim totalAmount As Double
    Dim noteStartRow As Long
    Static noteOrder As Integer
    Dim year As Variant
    
    ' Get the financial year
    year = GetFinancialYears(ws)
    If IsError(year) Then
        MsgBox "Failed to get financial year: " & year, vbExclamation
    End If

    ' Initialize
    noteCreated = False
    cashAmount = 0
    bankAmount = 0

    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow

    ' Increment the note order
    noteOrder = noteOrder + 1

    ' Create the note header
    ws.Cells(noteRow, 1).Value = noteOrder + 2
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "เงินสดและรายการเทียบเท่าเงินสด"
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    noteRow = noteRow + 1
    ws.Cells(noteRow, 9).Value = year
    noteRow = noteRow + 1

    ' Calculate amounts
    lastRow = trialBalanceSheet.Cells(trialBalanceSheet.Rows.Count, 2).End(xlUp).row
    For i = 2 To lastRow
        accountCode = trialBalanceSheet.Cells(i, 2).Value
        
        ' Cash (1010-1019)
        If accountCode >= "1010" And accountCode <= "1019" Then
            cashAmount = cashAmount + trialBalanceSheet.Cells(i, 6).Value
        End If
        
        ' Bank deposits (1020-1099)
        If accountCode >= "1020" And accountCode <= "1099" Then
            bankAmount = bankAmount + trialBalanceSheet.Cells(i, 6).Value
        End If
    Next i

    ' Add cash line if there's any amount
    If cashAmount <> 0 Then
        ws.Cells(noteRow, 3).Value = "เงินสด"
        ws.Cells(noteRow, 9).Value = cashAmount
        noteRow = noteRow + 1
        noteCreated = True
    End If

    ' Add bank deposits line if there's any amount
    If bankAmount <> 0 Then
        ws.Cells(noteRow, 3).Value = "เงินฝากธนาคาร"
        ws.Cells(noteRow, 9).Value = bankAmount
        noteRow = noteRow + 1
        noteCreated = True
    End If

    ' Add total if note was created
    If noteCreated Then
        totalAmount = cashAmount + bankAmount
        ws.Cells(noteRow, 3).Value = "รวม"
        With ws.Cells(noteRow, 9)
            .Value = totalAmount
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        noteRow = noteRow + 1

        ' Add the "EndOfNote" mark
        ws.Cells(noteRow, 1).Value = "EndOfNote"
        ws.Cells(noteRow, 1).Font.Color = vbWhite
    Else
        ' If no note was created, remove the header
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        noteOrder = noteOrder - 1
    End If

    ' Format the note
    FormatNote ws, noteStartRow, noteRow

    ' Return success/failure
    CreateFirstYearCashNote = noteCreated
End Function

Function CreateCashNote(ws As Worksheet, trialBalanceSheets As Collection) As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim accountCode As String
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim cashAmount1 As Double, cashAmount2 As Double
    Dim bankAmount1 As Double, bankAmount2 As Double
    Dim totalAmount1 As Double, totalAmount2 As Double
    Dim noteStartRow As Long
    Static noteOrder As Integer
    Dim years As Variant
    
    ' Get the financial years
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
    End If

    ' Initialize
    noteCreated = False
    cashAmount1 = 0: cashAmount2 = 0
    bankAmount1 = 0: bankAmount2 = 0

    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow

    ' Increment the note order
    noteOrder = noteOrder + 1

    ' Create the note header
    ws.Cells(noteRow, 1).Value = noteOrder + 2
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "เงินสดและรายการเทียบเท่าเงินสด"
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    noteRow = noteRow + 1
    ws.Cells(noteRow, 7).Value = years(1)
    ws.Cells(noteRow, 9).Value = years(2)
    noteRow = noteRow + 1

    ' Calculate amounts
    lastRow = trialBalanceSheets(1).Cells(trialBalanceSheets(1).Rows.Count, 2).End(xlUp).row
    For i = 2 To lastRow
        accountCode = trialBalanceSheets(1).Cells(i, 2).Value
        
        ' Cash (1010-1019)
        If accountCode >= "1010" And accountCode <= "1019" Then
            cashAmount1 = cashAmount1 + trialBalanceSheets(1).Cells(i, 6).Value
            cashAmount2 = cashAmount2 + GetAmountFromPreviousPeriod(trialBalanceSheets(2), accountCode)
        End If
        
        ' Bank deposits (1020-1099)
        If accountCode >= "1020" And accountCode <= "1099" Then
            bankAmount1 = bankAmount1 + trialBalanceSheets(1).Cells(i, 6).Value
            bankAmount2 = bankAmount2 + GetAmountFromPreviousPeriod(trialBalanceSheets(2), accountCode)
        End If
    Next i

    ' Add cash line if there's any amount
    If cashAmount1 <> 0 Or cashAmount2 <> 0 Then
        ws.Cells(noteRow, 3).Value = "เงินสด"
        ws.Cells(noteRow, 7).Value = cashAmount1
        ws.Cells(noteRow, 9).Value = cashAmount2
        noteRow = noteRow + 1
        noteCreated = True
    End If

    ' Add bank deposits line if there's any amount
    If bankAmount1 <> 0 Or bankAmount2 <> 0 Then
        ws.Cells(noteRow, 3).Value = "เงินฝากธนาคาร"
        ws.Cells(noteRow, 7).Value = bankAmount1
        ws.Cells(noteRow, 9).Value = bankAmount2
        noteRow = noteRow + 1
        noteCreated = True
    End If

    ' Add total if note was created
    If noteCreated Then
        totalAmount1 = cashAmount1 + bankAmount1
        totalAmount2 = cashAmount2 + bankAmount2
        ws.Cells(noteRow, 3).Value = "รวม"
        With ws.Cells(noteRow, 7)
            .Value = totalAmount1
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        With ws.Cells(noteRow, 9)
            .Value = totalAmount2
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        noteRow = noteRow + 1

        ' Add the "EndOfNote" mark
        ws.Cells(noteRow, 1).Value = "EndOfNote"
        ws.Cells(noteRow, 1).Font.Color = vbWhite
    Else
        ' If no note was created, remove the header
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        noteOrder = noteOrder - 1
    End If

    ' Format the note
    FormatNote ws, noteStartRow, noteRow

    ' Return success/failure
    CreateCashNote = noteCreated
End Function




Function CreateFirstYearNoteForLandBuildingEquipment(ws As Worksheet, trialBalanceSheet As Worksheet) As Boolean
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim noteStartRow As Long
    Static noteOrder As Integer
    Dim infoSheet As Worksheet
    Dim year As Variant
    Dim assetTotalRow As Long
    Dim accumulatedDepreciationTotalRow As Long
    
    ' Get the financial year using the new function
    year = GetFinancialYears(ws)
    If IsError(year) Then
        MsgBox "Failed to get financial year: " & year, vbExclamation
    End If

    ' Initialize noteCreated to False
    noteCreated = False

    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow

    ' Increment the note order
    noteOrder = noteOrder + 1

    ' Create the note header
    ws.Cells(noteRow, 1).Value = noteOrder + 2  ' Start from 3
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "ที่ดิน อาคารและอุปกรณ์"
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    noteRow = noteRow + 1

    ' Add column headers
    ws.Cells(noteRow, 4).Value = "ณ 31 ธ.ค. " & CStr(CLng(year) - 1)
    ws.Cells(noteRow, 6).Value = "ซื้อเพิ่ม"
    ws.Cells(noteRow, 7).Value = "จำหน่ายออก"
    ws.Cells(noteRow, 9).Value = "ณ 31 ธ.ค. " & year
    noteRow = noteRow + 1

    ' Add "ราคาทุนเดิม"
    ws.Cells(noteRow, 3).Value = "ราคาทุนเดิม"
    ws.Cells(noteRow, 3).Font.Bold = True
    noteRow = noteRow + 1

    ' Loop through the trial balance data for assets
    For i = 2 To trialBalanceSheet.Cells(trialBalanceSheet.Rows.Count, 2).End(xlUp).row
        accountCode = trialBalanceSheet.Cells(i, 2).Value
        
        ' Check if the account code is within range and doesn't contain a decimal
        If accountCode >= "1610" And accountCode <= "1659" And InStr(accountCode, ".") = 0 Then
            accountName = trialBalanceSheet.Cells(i, 1).Value
            amount = trialBalanceSheet.Cells(i, 5).Value
            
            ' Add the account detail to the note
            ws.Cells(noteRow, 3).Value = accountName
            ws.Cells(noteRow, 9).Value = amount
            
            ' For first year, assume all amounts are purchases
            ws.Cells(noteRow, 6).Value = amount
            
            noteRow = noteRow + 1
            noteCreated = True
        End If
    Next i

    ' Add total row for assets
    If noteCreated Then
        ws.Cells(noteRow, 3).Value = "รวม"
        For j = 6 To 9
            With ws.Cells(noteRow, j)
                .Formula = "=SUM(" & ws.Cells(noteStartRow + 3, j).Address & ":" & ws.Cells(noteRow - 1, j).Address & ")"
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
            End With
        Next j
        assetTotalRow = noteRow
        noteRow = noteRow + 2  ' Add two rows of space
    Else
        ' If no account details were added, remove the note header
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        noteOrder = noteOrder - 1  ' Decrement the note order if note was not created
        CreateFirstYearNoteForLandBuildingEquipment = False
        Exit Function
    End If

    ' Add "ค่าเสื่อมราคาสะสม" header
    ws.Cells(noteRow, 3).Value = "ค่าเสื่อมราคาสะสม"
    ws.Cells(noteRow, 3).Font.Bold = True
    noteRow = noteRow + 1

    ' Loop through the trial balance data for accumulated depreciation
    For i = 2 To trialBalanceSheet.Cells(trialBalanceSheet.Rows.Count, 2).End(xlUp).row
        accountCode = trialBalanceSheet.Cells(i, 2).Value
        
        ' Check if the account code is within range and contains a decimal
        If accountCode >= "1610" And accountCode <= "1659" And InStr(accountCode, ".") > 0 Then
            accountName = trialBalanceSheet.Cells(i, 1).Value
            amount = trialBalanceSheet.Cells(i, 5).Value
            
            ' Add the account detail to the note
            ws.Cells(noteRow, 3).Value = accountName
            ws.Cells(noteRow, 9).Value = amount
            
            ' For first year, assume all amounts are new depreciation
            ws.Cells(noteRow, 6).Value = amount
            
            noteRow = noteRow + 1
        End If
    Next i

    ' Add total row for accumulated depreciation
    ws.Cells(noteRow, 3).Value = "รวม"
    For j = 6 To 9
        If j <> 7 And j <> 8 Then  ' Skip columns G and H
            With ws.Cells(noteRow, j)
                .Formula = "=SUM(" & ws.Cells(assetTotalRow + 3, j).Address & ":" & ws.Cells(noteRow - 1, j).Address & ")"
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
            End With
        End If
    Next j
    accumulatedDepreciationTotalRow = noteRow
    noteRow = noteRow + 1

    ' Add "มูลค่าสุทธิ" row
    ws.Cells(noteRow, 3).Value = "มูลค่าสุทธิ"
    ws.Cells(noteRow, 3).Font.Bold = True
    
    With ws.Cells(noteRow, 9)
        .Formula = "=" & ws.Cells(assetTotalRow, 9).Address & "-" & ws.Cells(accumulatedDepreciationTotalRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    
    noteRow = noteRow + 1

    ' Add "ค่าเสื่อมราคา" row
    ws.Cells(noteRow, 3).Value = "ค่าเสื่อมราคา"
    ws.Cells(noteRow, 3).Font.Bold = True
    ws.Cells(noteRow, 9).Formula = "=" & ws.Cells(accumulatedDepreciationTotalRow, 6).Address
    noteRow = noteRow + 1

    ' Add the "EndOfNote" mark
    ws.Cells(noteRow, 1).Value = "EndOfNote"
    ws.Cells(noteRow, 1).Font.Color = vbWhite
    noteRow = noteRow + 1

    ' Format the note
    FormatNote ws, noteStartRow, noteRow

    ' Return True to indicate that the note was created
    CreateFirstYearNoteForLandBuildingEquipment = True
End Function

Function CreateNoteForLandBuildingEquipment(ws As Worksheet, trialBalanceSheets As Collection) As Boolean
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim accountCode As String
    Dim accountName As String
    Dim amountCurrent As Double, amountPrevious As Double
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim noteStartRow As Long
    Static noteOrder As Integer
    Dim infoSheet As Worksheet
    Dim years As Variant
    ' Dim previousYear As String
    Dim assetTotalRow As Long
    Dim accumulatedDepreciationTotalRow As Long
    Dim targetWorkbook As Workbook
    
    ' Get the target workbook
    Set targetWorkbook = ws.Parent
    
    ' Get the financial years using the new function
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
    End If

    ' Initialize noteCreated to False
    noteCreated = False

    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow

    ' Increment the note order
    noteOrder = noteOrder + 1

    ' Create the note header
    ws.Cells(noteRow, 1).Value = noteOrder + 2  ' Start from 3
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "ที่ดิน อาคารและอุปกรณ์"
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    noteRow = noteRow + 1

    ' Add column headers
    ws.Cells(noteRow, 4).Value = "ณ 31 ธ.ค. " & years(2)
    ws.Cells(noteRow, 6).Value = "ซื้อเพิ่ม"
    ws.Cells(noteRow, 7).Value = "จำหน่ายออก"
    ws.Cells(noteRow, 9).Value = "ณ 31 ธ.ค. " & years(1)
    noteRow = noteRow + 1

    ' Add "ราคาทุนเดิม"
    ws.Cells(noteRow, 3).Value = "ราคาทุนเดิม"
    ws.Cells(noteRow, 3).Font.Bold = True
    noteRow = noteRow + 1

    ' Loop through the trial balance data for assets
    For i = 2 To trialBalanceSheets(1).Cells(trialBalanceSheets(1).Rows.Count, 2).End(xlUp).row
        accountCode = trialBalanceSheets(1).Cells(i, 2).Value
        
        ' Check if the account code is within range and doesn't contain a decimal
        If accountCode >= "1610" And accountCode <= "1659" And InStr(accountCode, ".") = 0 Then
            accountName = trialBalanceSheets(1).Cells(i, 1).Value
            amountCurrent = trialBalanceSheets(1).Cells(i, 5).Value
            amountPrevious = GetAmountFromPreviousPeriodColumn5(trialBalanceSheets(2), accountCode)
            
            ' Add the account detail to the note
            ws.Cells(noteRow, 3).Value = accountName
            ws.Cells(noteRow, 4).Value = amountPrevious
            ws.Cells(noteRow, 9).Value = amountCurrent
            
            ' Calculate purchase or sale
            If amountCurrent > amountPrevious Then
                ws.Cells(noteRow, 6).Value = amountCurrent - amountPrevious
            ElseIf amountCurrent < amountPrevious Then
                ws.Cells(noteRow, 7).Value = amountPrevious - amountCurrent
            End If
            
            noteRow = noteRow + 1
            noteCreated = True
        End If
    Next i

    ' Add total row for assets
    If noteCreated Then
        ws.Cells(noteRow, 3).Value = "รวม"
        For j = 4 To 9
            If j <> 5 And j <> 8 Then  ' Skip columns E and H
                With ws.Cells(noteRow, j)
                    .Formula = "=SUM(" & ws.Cells(noteStartRow + 3, j).Address & ":" & ws.Cells(noteRow - 1, j).Address & ")"
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                End With
            End If
        Next j
        assetTotalRow = noteRow
        noteRow = noteRow + 2  ' Add two rows of space
    Else
        ' If no account details were added, remove the note header
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        noteOrder = noteOrder - 1  ' Decrement the note order if note was not created
        CreateNoteForLandBuildingEquipment = False
        Exit Function
    End If

    ' Add "ค่าเสื่อมราคาสะสม" header
    ws.Cells(noteRow, 3).Value = "ค่าเสื่อมราคาสะสม"
    ws.Cells(noteRow, 3).Font.Bold = True
    noteRow = noteRow + 1

    ' Loop through the trial balance data for accumulated depreciation
    For i = 2 To trialBalanceSheets(1).Cells(trialBalanceSheets(1).Rows.Count, 2).End(xlUp).row
        accountCode = trialBalanceSheets(1).Cells(i, 2).Value
        
        ' Check if the account code is within range and contains a decimal
        If accountCode >= "1610" And accountCode <= "1659" And InStr(accountCode, ".") > 0 Then
            accountName = trialBalanceSheets(1).Cells(i, 1).Value
            amountCurrent = trialBalanceSheets(1).Cells(i, 5).Value
            amountPrevious = GetAmountFromPreviousPeriodColumn5(trialBalanceSheets(2), accountCode)
            
            ' Add the account detail to the note
            ws.Cells(noteRow, 3).Value = accountName
            ws.Cells(noteRow, 4).Value = amountPrevious
            ws.Cells(noteRow, 9).Value = amountCurrent
            
            ' Calculate changes in accumulated depreciation
            ws.Cells(noteRow, 6).Value = amountCurrent - amountPrevious
            
            noteRow = noteRow + 1
        End If
    Next i

    ' Add total row for accumulated depreciation
    ws.Cells(noteRow, 3).Value = "รวม"
    For j = 4 To 9
        If j <> 5 And j <> 8 Then  ' Skip columns E, G, and H
            With ws.Cells(noteRow, j)
                .Formula = "=SUM(" & ws.Cells(assetTotalRow + 3, j).Address & ":" & ws.Cells(noteRow - 1, j).Address & ")"
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
            End With
        End If
    Next j
    
    accumulatedDepreciationTotalRow = noteRow
    noteRow = noteRow + 1

    ' Add "มูลค่าสุทธิ" row
    ws.Cells(noteRow, 3).Value = "มูลค่าสุทธิ"
    ws.Cells(noteRow, 3).Font.Bold = True

    ' Add borders to columns D, F, G, and I
    For Each col In Array(4, 6, 7, 9)
        With ws.Cells(noteRow, col)
            If col = 4 Or col = 9 Then
                .Formula = "=" & ws.Cells(assetTotalRow, col).Address & "-" & ws.Cells(accumulatedDepreciationTotalRow, col).Address
            End If
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
    Next col

    noteRow = noteRow + 1

    ' Add "ค่าเสื่อมราคา" row
    ws.Cells(noteRow, 3).Value = "ค่าเสื่อมราคา"
    ws.Cells(noteRow, 3).Font.Bold = True
    ws.Cells(noteRow, 9).Formula = "=" & ws.Cells(accumulatedDepreciationTotalRow, 6).Address & "-" & ws.Cells(accumulatedDepreciationTotalRow, 7).Address
    noteRow = noteRow + 1
    
    ' Add the "EndOfNote" mark
    ws.Cells(noteRow, 1).Value = "EndOfNote"
    ws.Cells(noteRow, 1).Font.Color = vbWhite
    noteRow = noteRow + 1

    ' Format the note
    FormatNote ws, noteStartRow, noteRow

    ' Return True to indicate that the note was created
    CreateNoteForLandBuildingEquipment = True
End Function

Function CreateFirstYearLongTermLoansNote(ws As Worksheet, trialSheet As Worksheet) As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim totalAmount As Double
    Dim noteStartRow As Long
    Static noteOrder As Integer
    Dim year As Variant
    Dim currentPortion As Double
    
    ' Get the financial year
    year = GetFinancialYears(ws)
    If IsError(year) Then
        MsgBox "Failed to get financial year: " & year, vbExclamation
        Exit Function
    End If

    ' Initialize variables
    noteCreated = False
    totalAmount = 0
    
    ' Get the last row of data in the trial sheet
    lastRow = trialSheet.Cells(trialSheet.Rows.Count, 1).End(xlUp).row

    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow

    ' Increment the note order
    noteOrder = noteOrder + 1

    ' Create the note header
    ws.Cells(noteRow, 1).Value = noteOrder + 2  ' Start from 3
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "เงินกู้ยืมระยะยาวจากสถาบันการเงิน"
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    ws.Cells(noteRow + 1, 9).Value = year
    noteRow = noteRow + 2

    ' Loop through the trial data and check for account codes
    For i = 2 To lastRow
        accountCode = trialSheet.Cells(i, 2).Value
        accountName = trialSheet.Cells(i, 1).Value
        amount = trialSheet.Cells(i, 6).Value

        If accountCode >= "2120" And accountCode <= "2123" And accountCode <> "2121" Then
            If amount <> 0 And Not IsEmpty(amount) Then
                ' Add the account detail to the note
                ws.Cells(noteRow, 3).Value = accountName
                ws.Cells(noteRow, 9).Value = amount
                noteRow = noteRow + 1
                noteCreated = True
            End If
            totalAmount = totalAmount + amount
        End If
    Next i

    ' Check if any account details were added to the note
    If noteCreated Then
        ' Add total row
        ws.Cells(noteRow, 3).Value = "รวม"
        With ws.Cells(noteRow, 9)
            .Value = totalAmount
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        noteRow = noteRow + 1
        
       ' Get current portion from user and store in global variable
        gLoanCurrentPortion.CurrentYear = CDbl(InputBox("กรุณากรอกส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี สำหรับปี " & year, "Current Portion"))
        
        ' Add current portion row
        ws.Cells(noteRow, 3).Value = "หัก ส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี"
        ws.Cells(noteRow, 9).Value = gLoanCurrentPortion.CurrentYear
        noteRow = noteRow + 1
        
        ' Add net amount row
        ws.Cells(noteRow, 3).Value = "เงินกู้ยืมระยะยาวสุทธิจากส่วนที่ถึงกำหนดชำระคืนภายในหนึ่งปี"
        With ws.Cells(noteRow, 9)
            .Value = totalAmount - gLoanCurrentPortion.CurrentYear
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        noteRow = noteRow + 1

        ' Add the "EndOfNote" mark
        ws.Cells(noteRow, 1).Value = "EndOfNote"
        ws.Cells(noteRow, 1).Font.Color = vbWhite
    Else
        ' If no account details were added, remove the note header
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        noteOrder = noteOrder - 1  ' Decrement the note order if note was not created
    End If

    If noteRow > 34 Then
        ' Move the note to a new worksheet and recreate it
        Set ws = HandleNoteExceedingRow34(ws, "เงินกู้ยืมระยะยาวจากสถาบันการเงิน", noteStartRow, noteRow, trialSheet)
        noteCreated = True
    End If

    ' Format the note
    FormatNote ws, noteStartRow, noteRow

    ' Return the value of noteCreated
    CreateFirstYearLongTermLoansNote = noteCreated
End Function

Function CreateLongTermLoansNote(ws As Worksheet, trialBalanceSheets As Collection, trialPLSheets As Collection) As Boolean
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long
    Dim accountCode As String
    Dim accountName As String
    Dim amount1 As Double, amount2 As Double
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim totalAmount1 As Double, totalAmount2 As Double
    Dim noteStartRow As Long
    Dim uniqueAccountCodes As New Collection
    Static noteOrder As Integer
    Dim years As Variant
    Dim targetWorkbook As Workbook
    Dim currentPortion1 As Double, currentPortion2 As Double
    
    ' Get the target workbook and years
    Set targetWorkbook = ws.Parent
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
            Exit Function
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
        Exit Function
    End If
    
    ' Initialize variables
    noteCreated = False
    totalAmount1 = 0: totalAmount2 = 0
    
    ' Set up note header and position
    lastRow1 = trialBalanceSheets(1).Cells(trialBalanceSheets(1).Rows.Count, 1).End(xlUp).row
    lastRow2 = trialBalanceSheets(2).Cells(trialBalanceSheets(2).Rows.Count, 1).End(xlUp).row
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow
    noteOrder = noteOrder + 1
    
    ' Create note header
    ws.Cells(noteRow, 1).Value = noteOrder + 2
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "เงินกู้ยืมระยะยาวจากสถาบันการเงิน"
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    ws.Cells(noteRow + 1, 7).Value = years(1)
    ws.Cells(noteRow + 1, 9).Value = years(2)
    noteRow = noteRow + 2
    
    ' Process current period accounts
    For i = 2 To lastRow1
        accountCode = trialBalanceSheets(1).Cells(i, 2).Value
        accountName = trialBalanceSheets(1).Cells(i, 1).Value
        amount1 = trialBalanceSheets(1).Cells(i, 6).Value
        
        If accountCode >= "2120" And accountCode <= "2123" And accountCode <> "2121" Then
            If Not ContainsAccountCode(uniqueAccountCodes, accountCode) Then
                uniqueAccountCodes.Add accountCode, CStr(accountCode)
                amount2 = GetAmountFromPreviousPeriod(trialBalanceSheets(2), accountCode)
                
                If (amount1 <> 0 And Not IsEmpty(amount1)) Or (amount2 <> 0 And Not IsEmpty(amount2)) Then
                    ws.Cells(noteRow, 3).Value = accountName
                    ws.Cells(noteRow, 7).Value = amount1
                    ws.Cells(noteRow, 9).Value = amount2
                    noteRow = noteRow + 1
                    noteCreated = True
                End If
                
                totalAmount1 = totalAmount1 + amount1
                totalAmount2 = totalAmount2 + amount2
            End If
        End If
    Next i
    
    ' Add total row
    If noteCreated Then
        With ws.Cells(noteRow, 3)
            .Value = "รวม"
        End With
        With ws.Cells(noteRow, 7)
            .Value = totalAmount1
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        With ws.Cells(noteRow, 9)
            .Value = totalAmount2
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        noteRow = noteRow + 1
        
       ' Get current portions from user and store in global variable
        gLoanCurrentPortion.CurrentYear = CDbl(InputBox("กรุณากรอกส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี สำหรับปี " & years(1), "Current Portion - Current Year"))
        gLoanCurrentPortion.PreviousYear = CDbl(InputBox("กรุณากรอกส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี สำหรับปี " & years(2), "Current Portion - Previous Year"))
        
        ' Add current portion row
        ws.Cells(noteRow, 3).Value = "หัก ส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี"
        ws.Cells(noteRow, 7).Value = gLoanCurrentPortion.CurrentYear
        ws.Cells(noteRow, 9).Value = gLoanCurrentPortion.PreviousYear
        noteRow = noteRow + 1
        
        ' Add net amount row
        With ws.Cells(noteRow, 3)
            .Value = "เงินกู้ยืมระยะยาวสุทธิจากส่วนที่ถึงกำหนดชำระคืนภายในหนึ่งปี"
        End With
        With ws.Cells(noteRow, 7)
            .Value = totalAmount1 - gLoanCurrentPortion.CurrentYear
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        With ws.Cells(noteRow, 9)
            .Value = totalAmount2 - gLoanCurrentPortion.PreviousYear
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        noteRow = noteRow + 1
        
        ' Add EndOfNote mark
        ws.Cells(noteRow, 1).Value = "EndOfNote"
        ws.Cells(noteRow, 1).Font.Color = vbWhite
    Else
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        noteOrder = noteOrder - 1
    End If
    
    ' Handle pagination and formatting
    If noteRow > 34 Then
        Set ws = HandleNoteExceedingRow34(ws, "เงินกู้ยืมระยะยาวจากสถาบันการเงิน", noteStartRow, noteRow, trialBalanceSheets(1))
        noteCreated = True
    End If
    
    FormatNote ws, noteStartRow, noteRow
    CreateLongTermLoansNote = noteCreated
End Function



Function GetAmountFromPreviousPeriodColumn5(trialBalanceSheet As Worksheet, accountCode As String) As Double
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = trialBalanceSheet.Cells(trialBalanceSheet.Rows.Count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If trialBalanceSheet.Cells(i, 2).Value = accountCode Then
            GetAmountFromPreviousPeriodColumn5 = trialBalanceSheet.Cells(i, 5).Value
            Exit Function
        End If
    Next i
    
    GetAmountFromPreviousPeriodColumn5 = 0
End Function

Function CreateExpensesByNatureNote(ws As Worksheet) As Boolean
    Dim noteRow As Long
    Dim noteStartRow As Long
    Dim infoSheet As Worksheet
    Dim years As Variant
    ' Dim previousYear As String
    Dim targetWorkbook As Workbook
    
    ' Get the target workbook
    Set targetWorkbook = ws.Parent
    
    ' Get the financial years using the new function
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
    End If

    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow

    ' Increment the note order
    noteOrder = noteOrder + 1

    ' Create the note header
    ws.Cells(noteRow, 1).Value = noteOrder + 2  ' Start from 3
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "ค่าใช้จ่ายจำแนกตามธรรมชาติของค่าใช้จ่าย"
    
    ' Highlight the main title in yellow
    ws.Cells(noteRow, 2).Interior.Color = RGB(255, 255, 0)  ' Yellow highlight
    
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    ws.Cells(noteRow + 1, 7).Value = years(1)
    ws.Cells(noteRow + 1, 9).Value = years(2)
    noteRow = noteRow + 2

    ' Add expense categories
    AddExpenseCategory ws, noteRow, "การเปลี่ยนแปลงในสินค้าสำเร็จรูปและงานระหว่างทำ"
    AddExpenseCategory ws, noteRow, "งานที่ทำโดยกิจการและบันทึกเป็นรายจ่ายฝ่ายทุน"
    AddExpenseCategory ws, noteRow, "วัตถุดิบและวัสดุสิ้นเปลืองใช้ไป"
    AddExpenseCategory ws, noteRow, "ค่าใช้จ่ายผลประโยชน์พนักงาน"
    AddExpenseCategory ws, noteRow, "ค่าเสื่อมราคาและค่าตัดจำหน่าย"
    AddExpenseCategory ws, noteRow, "ค่าใช้จ่ายอื่น"

    ' Add total
    ws.Cells(noteRow, 3).Value = "รวม"
    ws.Cells(noteRow, 3).Font.Bold = True
    noteRow = noteRow + 1

    ' Add the "EndOfNote" mark
    ws.Cells(noteRow, 1).Value = "EndOfNote"
    ws.Cells(noteRow, 1).Font.Color = vbWhite

    ' Check if note exceeds 34 rows
    If noteRow - noteStartRow > 34 Then
        ' Move the note to a new worksheet
        Set ws = HandleNoteExceedingRow34(ws, "ค่าใช้จ่ายแยกตามลักษณะของค่าใช้จ่าย", noteStartRow, noteRow, Nothing)
    End If

    ' Format the note
    FormatNote ws, noteStartRow, noteRow

    CreateExpensesByNatureNote = True
End Function

Sub AddExpenseCategory(ws As Worksheet, ByRef row As Long, categoryName As String)
    ws.Cells(row, 3).Value = categoryName
    row = row + 1
End Sub

Function IsLimitedCompany(targetWorkbook As Workbook) As Boolean
    Dim infoSheet As Worksheet
    Set infoSheet = targetWorkbook.Sheets("Info")
    IsLimitedCompany = (infoSheet.Range("B2").Value = "บริษัทจำกัด")
End Function

Function CreateFinancialApprovalNote(ws As Worksheet) As Boolean
    Dim noteRow As Long
    Dim noteStartRow As Long
    Static noteOrder As Integer
    Dim targetWorkbook As Workbook
    
    ' Get the target workbook
    Set targetWorkbook = ws.Parent
    
    ' Check if this is a limited company
    If Not IsLimitedCompany(targetWorkbook) Then
        CreateFinancialApprovalNote = False
        Exit Function
    End If
    
    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow
    
    ' Increment the note order
    noteOrder = noteOrder + 1
    
    ' Create the note header
    ws.Cells(noteRow, 1).Value = noteOrder + 2
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "การอนุมัติงบการเงิน"
    ws.Cells(noteRow, 2).Font.Bold = True
    noteRow = noteRow + 1
    
    ' Add approval text
    ws.Cells(noteRow, 3).Value = "งบการเงินนี้ได้การรับอนุมัติให้ออกงบการเงินโดยคณะกรรมการผู้มีอำนาจของบริษัทแล้ว"
    noteRow = noteRow + 1
    
    ' Add the "EndOfNote" mark
    ws.Cells(noteRow, 1).Value = "EndOfNote"
    ws.Cells(noteRow, 1).Font.Color = vbWhite
    
    CreateFinancialApprovalNote = True
End Function





Function HandleNoteExceedingRow34(ws As Worksheet, noteName As String, noteStartRow As Long, noteEndRow As Long, trialBalanceSheet As Worksheet) As Worksheet
    ' Create a new worksheet for the note
    Dim newWS As Worksheet
    Dim targetWorkbook As Workbook
    
    ' Get the target workbook
    Set targetWorkbook = ws.Parent
    
    ' Create the new worksheet in the target workbook
    Set newWS = targetWorkbook.Sheets.Add(After:=ws)
    ' newWS.Name = noteName
    
    ' Call the CreateHeader function to create the header in the new worksheet
    CreateHeader newWS
    
    ' Find the first empty row after the header in the new worksheet
    Dim newNoteStartRow As Long
    newNoteStartRow = newWS.Cells(newWS.Rows.Count, 1).End(xlUp).row + 1
    
    ' Copy the note content to the new worksheet
    ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteEndRow, 11)).Copy
    newWS.Cells(newNoteStartRow, 1).PasteSpecial xlPasteValues
    newWS.Cells(newNoteStartRow, 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    ' Remove the note from the original worksheet
    ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteEndRow, 11)).Delete Shift:=xlUp
    
    ' Apply Thai Sarabun font and font size 14 to the new worksheet
    newWS.Cells.Font.Name = "TH Sarabun New"
    newWS.Cells.Font.Size = 14
    
    ' Set number format to use comma style for columns J and K in the new worksheet
    newWS.Columns("J:K").NumberFormat = "#,##0.00"
    
    ' Return the new worksheet
    Set HandleNoteExceedingRow34 = newWS
End Function

Function ContainsAccountCode(uniqueAccountCodes As Collection, accountCode As String) As Boolean
    On Error Resume Next
    uniqueAccountCodes.Item CStr(accountCode)
    ContainsAccountCode = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub FormatNote(ws As Worksheet, startRow As Long, endRow As Long)
    Dim accountingFormat As String
    accountingFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    With ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 11))
        .Font.Name = "TH Sarabun New"
        .Font.Size = 14
        
        ' Set accounting format for columns D, F, G, and I if they contain values
        Dim column As Variant
        For Each column In Array("D", "F", "G", "I")
            If WorksheetFunction.CountA(.Columns(column)) > 0 Then
                .Columns(column).NumberFormatLocal = accountingFormat
            End If
        Next column
        
        ' Right align amount columns
        .Columns("D").HorizontalAlignment = xlRight
        .Columns("F").HorizontalAlignment = xlRight
        .Columns("G").HorizontalAlignment = xlRight
        .Columns("I").HorizontalAlignment = xlRight
    End With
    
    ' Bold the note name and total
    ws.Cells(startRow, 2).Font.Bold = True
    ws.Cells(endRow - 1, 3).Font.Bold = True
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 4
    ws.Columns("C").ColumnWidth = 25
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 2
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 2
    ws.Columns("I").ColumnWidth = 12
    
    ' Format year header row (assuming it's in the second row of the note)
    With ws.Rows(startRow + 1)
        .NumberFormat = "@"  ' Set to text format to prevent number formatting
        .HorizontalAlignment = xlCenter
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    ' Reapply accounting format to amount columns, excluding the year header
    If endRow > startRow + 2 Then
        Dim reapplyRange As Range
        Set reapplyRange = Union( _
            ws.Range(ws.Cells(startRow + 2, 4), ws.Cells(endRow, 4)), _
            ws.Range(ws.Cells(startRow + 2, 6), ws.Cells(endRow, 6)), _
            ws.Range(ws.Cells(startRow + 2, 7), ws.Cells(endRow, 7)), _
            ws.Range(ws.Cells(startRow + 2, 9), ws.Cells(endRow, 9)) _
        )
        reapplyRange.NumberFormatLocal = accountingFormat
    End If
End Sub



