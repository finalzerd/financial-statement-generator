Option Explicit

Private policyFilePath As String

Public Sub CreateGICContent(targetWorkbook As Workbook)
    Dim ws As Worksheet
    Dim currentRow As Long
    
    ' Create or get GIC sheet
    Set ws = GetOrCreateGICSheet(targetWorkbook)
    
    ' Create the header
    CreateHeader ws, "General Information"
    
    ' Add General Information
    currentRow = AddGeneralInformation(ws, targetWorkbook)
    
    ' Add Basis of Preparation
    currentRow = AddBasisOfPreparation(targetWorkbook, ws, currentRow)
    
    ' Add Accounting Policy
    AddAccountingPolicy targetWorkbook, ws, currentRow
    
    ' Apply formatting at the end
    FormatWorksheet ws
End Sub

Private Function GetOrCreateGICSheet(targetWorkbook As Workbook) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = targetWorkbook.Sheets("GIC")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        ws.Name = "GIC"
    End If
    
    Set GetOrCreateGICSheet = ws
End Function

Private Function AddGeneralInformation(ws As Worksheet, targetWorkbook As Workbook) As Long
    Dim infoSheet As Worksheet
    Dim entityNumber As String
    Dim entityFilePath As String
    Dim entityData As Object
    Dim currentRow As Long
    Dim mergedRange As Range
    
    On Error GoTo ErrorHandler
    
    Set infoSheet = targetWorkbook.Sheets("Info")
    entityNumber = infoSheet.Range("B4").Value
    entityFilePath = targetWorkbook.Path & "\ExtractWebDBD\" & entityNumber & ".csv"
    
    If Dir(entityFilePath) = "" Then
        MsgBox "Entity file not found: " & entityFilePath, vbExclamation
        Exit Function
    End If
    
    Set entityData = ReadCSV(entityFilePath)
    
    ' Add section 1: General Information
    currentRow = 5 ' Start after header
    
    ' Add main title
    ws.Cells(currentRow, 1).Value = "1"
    ws.Cells(currentRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(currentRow, 2).Value = "ข้อมูลทั่วไป"
    ws.Cells(currentRow, 2).Font.Bold = True
    
    currentRow = currentRow + 1 ' Add spacing
    
    ' Add sections 1.1, 1.2, 1.3
    currentRow = AddGeneralInfoSections(ws, currentRow, entityData)
    
    ' Add Company Status section
    currentRow = AddCompanyStatus(ws, currentRow, entityData)
    
    AddGeneralInformation = currentRow + 1 ' Return next available row
    Exit Function

ErrorHandler:
    MsgBox "Error in AddGeneralInformation: " & Err.Description
    AddGeneralInformation = currentRow
End Function

Private Function AddGeneralInfoSections(ws As Worksheet, startRow As Long, entityData As Object) As Long
    Dim currentRow As Long
    Dim mergedRange As Range
    
    currentRow = startRow
    
    ' Section 1.1
    AddInfoSection ws, currentRow, "1.1", "สถานะทางกฎหมาย", "เป็นนิติบุคคลจัดตั้งตามกฎหมายไทย"
    currentRow = currentRow + 1
    
    ' Section 1.2
    AddInfoSection ws, currentRow, "1.2", "สถานที่ตั้ง", entityData("D2")
    currentRow = currentRow + 1
    
    ' Section 1.3
    AddInfoSection ws, currentRow, "1.3", "ลักษณะธุรกิจและการดำเนินงาน", entityData("E2")
    currentRow = currentRow + 1
    
    AddGeneralInfoSections = currentRow
End Function

Private Sub AddInfoSection(ws As Worksheet, row As Long, sectionNum As String, title As String, content As String)
    ws.Cells(row, 3).Value = sectionNum
    ws.Cells(row, 3).HorizontalAlignment = xlCenter
    
    ws.Range(ws.Cells(row, 4), ws.Cells(row, 5)).Merge
    ws.Cells(row, 4).Value = title
    
    With ws.Range(ws.Cells(row, 6), ws.Cells(row, 8))
        .Merge
        .Value = content
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Function AddCompanyStatus(ws As Worksheet, startRow As Long, entityData As Object) As Long
    Dim currentRow As Long
    Dim statusText As String
    
    currentRow = startRow + 1
    
    ws.Cells(currentRow, 1).Value = "2"
    ws.Cells(currentRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(currentRow, 2).Value = "ฐานะการดำเนินงานของบริษัท"
    ws.Cells(currentRow, 2).Font.Bold = True
    
    currentRow = currentRow + 1
    
    statusText = entityData("G2") & " ได้จดทะเบียนตามประมวลกฎหมายแพ่งและพาณิชย์เป็นนิติบุคคล ประเภท " & _
                entityData("A2") & " เมื่อวันที่ " & entityData("B2") & " ทะเบียนเลขที่ " & entityData("H2")
    
    With ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 8))
        .Merge
        .Value = statusText
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With
    
    ' Set the row height to 48
    ws.Rows(currentRow).rowHeight = 48
    
    AddCompanyStatus = currentRow
End Function

Private Function AddBasisOfPreparation(targetWorkbook As Workbook, ws As Worksheet, currentRow As Long) As Long
    Dim basisFilePath As String
    Dim basisWorkbook As Workbook
    Dim basisSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rowHeight As Double
    
    On Error GoTo ErrorHandler
    
    basisFilePath = targetWorkbook.Path & "\AccountingPolicy\basis_of_preparation.xlsx"
    
    If Dir(basisFilePath) = "" Then
        MsgBox "Basis of preparation file not found: " & basisFilePath, vbExclamation
        Exit Function
    End If
    
    Set basisWorkbook = Workbooks.Open(basisFilePath)
    Set basisSheet = basisWorkbook.Sheets(1)
    
    lastRow = basisSheet.Cells(basisSheet.Rows.Count, "A").End(xlUp).row
    
    currentRow = currentRow + 1
    
    ' Add the main title
    ws.Cells(currentRow, 1).Value = "3"
    ws.Cells(currentRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(currentRow, 2).Value = "เกณฑ์ในการจัดทำและนำเสนองบการเงิน"
    ws.Cells(currentRow, 2).Font.Bold = True
    currentRow = currentRow + 1
    
    For i = 2 To lastRow
        If basisSheet.Cells(i, 1).Value <> "" Then
            ws.Cells(currentRow, 2).Value = basisSheet.Cells(i, 1).Value
            ws.Cells(currentRow, 2).HorizontalAlignment = xlCenter
            
            With ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 9))
                .Merge
                .Value = basisSheet.Cells(i, 2).Value
                .WrapText = True
                .VerticalAlignment = xlTop
                .HorizontalAlignment = xlLeft
            End With
            
            ' Set row height from column C of source file
            rowHeight = basisSheet.Cells(i, 3).Value
            If rowHeight > 0 Then
                ws.Rows(currentRow).rowHeight = rowHeight
            End If
            
            currentRow = currentRow + 1
        End If
    Next i
    
    basisWorkbook.Close SaveChanges:=False
    AddBasisOfPreparation = currentRow
    Exit Function
    
ErrorHandler:
    MsgBox "Error in AddBasisOfPreparation: " & Err.Description
    On Error Resume Next
    If Not basisWorkbook Is Nothing Then basisWorkbook.Close SaveChanges:=False
    On Error GoTo 0
    AddBasisOfPreparation = currentRow
End Function

Private Sub AddAccountingPolicy(targetWorkbook As Workbook, ws As Worksheet, startRow As Long)
    Dim policyWorkbook As Workbook
    Dim policySheet As Worksheet
    Dim trialBalanceSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim accountCodes As Collection
    Dim topic As String, detail1 As String, detail2 As String
    Dim codeRange As String, matchFound As Boolean
    Dim orderNum As Long
    Dim currentRow As Long
    Dim code As Variant
    Dim rowHeight As Double
    
    On Error GoTo ErrorHandler
    
    currentRow = startRow + 1
    policyFilePath = targetWorkbook.Path & "\AccountingPolicy\accounting_policy.xlsx"
    
    If Dir(policyFilePath) = "" Then
        MsgBox "Accounting policy file not found: " & policyFilePath, vbExclamation
        Exit Sub
    End If
    
    Set policyWorkbook = Workbooks.Open(policyFilePath)
    Set policySheet = policyWorkbook.Sheets(1)
    Set trialBalanceSheet = targetWorkbook.Sheets("Trial Balance 1")
    Set accountCodes = New Collection
    
    ' Get account codes
    lastRow = trialBalanceSheet.Cells(trialBalanceSheet.Rows.Count, 2).End(xlUp).row
    For i = 2 To lastRow
        On Error Resume Next
        accountCodes.Add trialBalanceSheet.Cells(i, 2).Value, CStr(trialBalanceSheet.Cells(i, 2).Value)
        On Error GoTo 0
    Next i
    
    ' Add main topic header
    ws.Cells(currentRow, 1).Value = "4"
    ws.Cells(currentRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(currentRow, 2).Value = "สรุปนโยบายการบัญชีที่สำคัญ"
    ws.Cells(currentRow, 2).Font.Bold = True
    ws.Rows(currentRow).AutoFit ' Autofit the main header
    currentRow = currentRow + 1
    
    ' Process policies
    lastRow = policySheet.Cells(policySheet.Rows.Count, 1).End(xlUp).row
    orderNum = 1
    
    For i = 2 To lastRow
        codeRange = policySheet.Cells(i, 1).Value
        topic = policySheet.Cells(i, 2).Value
        detail1 = policySheet.Cells(i, 3).Value
        detail2 = policySheet.Cells(i, 4).Value
        
        matchFound = (codeRange = "0")
        
        If Not matchFound Then
            For Each code In accountCodes
                If IsCodeInRange(CStr(code), codeRange) Then
                    matchFound = True
                    Exit For
                End If
            Next code
        End If
        
        If matchFound Then
            If topic <> "" Then
                ' Topic header
                ws.Cells(currentRow, 2).Value = "4." & orderNum
                ws.Cells(currentRow, 2).HorizontalAlignment = xlCenter
                
                With ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 9))
                    .Merge
                    .Value = topic
                    .Font.Bold = True
                    .WrapText = True
                    .VerticalAlignment = xlTop
                    .HorizontalAlignment = xlLeft
                End With
                
                ' Autofit the topic header row
                ws.Rows(currentRow).AutoFit
                
                currentRow = currentRow + 1
                orderNum = orderNum + 1
            End If
            
            If detail2 <> "" Then
                ' First detail as subheader with AutoFit
                With ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 9))
                    .Merge
                    .Value = detail1
                    .Font.Bold = True
                    .Font.Italic = True
                    .WrapText = True
                    .VerticalAlignment = xlTop
                    .HorizontalAlignment = xlLeft
                End With
                
                ' Autofit the subheader row
                ws.Rows(currentRow).AutoFit
                
                currentRow = currentRow + 1
                
                ' Second detail with specified height from source
                With ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 9))
                    .Merge
                    .Value = detail2
                    .WrapText = True
                    .VerticalAlignment = xlTop
                    .HorizontalAlignment = xlLeft
                End With
                
                ' Use height from source for detail row
                rowHeight = policySheet.Cells(i, 5).Value
                If rowHeight > 0 Then
                    ws.Rows(currentRow).rowHeight = rowHeight
                End If
                
            Else
                ' Single detail with specified height from source
                With ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 9))
                    .Merge
                    .Value = detail1
                    .WrapText = True
                    .VerticalAlignment = xlTop
                    .HorizontalAlignment = xlLeft
                End With
                
                ' Use height from source for detail row
                rowHeight = policySheet.Cells(i, 5).Value
                If rowHeight > 0 Then
                    ws.Rows(currentRow).rowHeight = rowHeight
                End If
            End If
            currentRow = currentRow + 1
        End If
    Next i
    
    policyWorkbook.Close SaveChanges:=False
    Exit Sub

ErrorHandler:
    MsgBox "Error in AddAccountingPolicy: " & Err.Description, vbCritical
    On Error Resume Next
    If Not policyWorkbook Is Nothing Then
        policyWorkbook.Close SaveChanges:=False
    End If
    On Error GoTo 0
End Sub

Private Function ReadCSV(filePath As String) As Object
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String
    Dim dict As Object
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Read header
    line = ts.ReadLine
    parts = Split(line, ",")
    
    ' Read data
    line = ts.ReadLine
    parts = Split(line, ",")
    
    ' Populate dictionary
    For i = 0 To UBound(parts)
        dict.Add Chr(65 + i) & "2", parts(i)
    Next i
    
    ts.Close
    Set ReadCSV = dict
End Function

Private Function IsCodeInRange(code As String, codeRange As String) As Boolean
    Dim rangeParts() As String
    Dim lowerBound As String, upperBound As String
    
    If InStr(codeRange, "-") > 0 Then
        rangeParts = Split(codeRange, "-")
        lowerBound = rangeParts(0)
        upperBound = rangeParts(1)
        IsCodeInRange = (code >= lowerBound And code <= upperBound)
    Else
        IsCodeInRange = (code = codeRange)
    End If
End Function

Private Sub FormatWorksheet(ws As Worksheet)
    ' Apply Thai Sarabun font and font size 14 to the worksheet
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 5.3
    ws.Columns("C").ColumnWidth = 8.33
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 14.11
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G").ColumnWidth = 16.44
    ws.Columns("H").ColumnWidth = 22.4
    ws.Columns("I").ColumnWidth = 2.33
    
    ' Set top alignment for columns B and C
    ws.Columns("B:C").VerticalAlignment = xlTop
End Sub

