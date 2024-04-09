Sub CombinedQBMacro()
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Delete rows above the first occurrence of the headers
    DeleteRowsAboveHeaders ws

    Dim debitsCol As Long
    Dim creditsCol As Long
    Dim amountCol As Long
    debitsCol = GetColumnNumber(ws, "Debit")
    creditsCol = GetColumnNumber(ws, "Credit")
    amountCol = GetColumnNumber(ws, "Amount")

    If debitsCol > 0 And creditsCol > 0 Then
        QBFormatDCB
    ElseIf amountCol > 0 Then
        QBFormatAmtBal
    Else
        MsgBox "The worksheet format is not recognized. Please check the column headers.", vbExclamation
    End If

    Application.ScreenUpdating = True
End Sub


Private Sub DeleteRowsAboveHeaders(sheet As Worksheet)
    Dim headers() As Variant
    headers = Array("Debit", "Credit", "Amount")
    
    Dim header As Variant
    Dim rng As Range
    Dim firstRow As Long
    firstRow = sheet.Rows.Count

    For Each header In headers
        Set rng = sheet.Cells.Find(What:=header, LookIn:=xlValues, LookAt:=xlPart)
        If Not rng Is Nothing Then
            If rng.Row < firstRow Then firstRow = rng.Row
        End If
    Next header

    If firstRow > 1 Then
        sheet.Rows("1:" & firstRow - 1).Delete
    End If
End Sub

Private Function GetColumnNumber(sheet As Worksheet, columnName As String) As Long
    Dim col As Long
    For col = 1 To sheet.Cells(1, sheet.Columns.Count).End(xlToLeft).Column
        If Trim(sheet.Cells(1, col).Value) = columnName Then
            GetColumnNumber = col
            Exit Function
        End If
    Next col
    GetColumnNumber = 0
End Function


Sub QBFormatAmtBal()

Application.ScreenUpdating = False

'Remove QB Export Tips sheet (if exists)
Application.DisplayAlerts = False
For Each sheet In ActiveWorkbook.Worksheets
    If sheet.Name Like "*Tips" Then
        sheet.Delete
    End If
Next sheet
Application.DisplayAlerts = True

'Make whole sheet normal formatting
With ActiveSheet.Cells
    .Style = "Normal"
    .UnMerge
End With

    'Remove any rows above the headers (search for "Date")
    Dim rng As Range
    Dim dateCell As Range
    Dim rowsToDelete As Range
    Dim i As Long
    Dim lastRow As Long

    With ActiveSheet
        lastRow = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        Set rng = .Range("A1:AZ" & lastRow)

        Set dateCell = rng.Find(What:="Date", LookAt:=xlPart, MatchCase:=False) 'non case-sensitive, and match part of cell content

        If Not dateCell Is Nothing Then ' If "Date" was found
            For i = dateCell.Row - 1 To 1 Step -1 ' Loop upwards from "Date" to the first row
                .Rows(i).Delete ' Delete the row
            Next i
        End If
    End With

'Remove all blank columns
Dim ws As Worksheet
Dim lastColumn As Long
Dim columnNumber As Long

Set ws = ActiveSheet
lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

For columnNumber = lastColumn To 1 Step -1
    If WorksheetFunction.CountA(ws.Columns(columnNumber)) = 0 Then
        ws.Columns(columnNumber).Delete
    End If
Next columnNumber

'Set xCol as first column with a header

Dim xRg As Range
Dim xStr As String
Dim xCol As Integer
Dim vRg As Range
Dim firstRow As Long
Dim lrow As Long

xCol = Range("A1").End(xlToRight).Column

'A1 = "Account ref. number"

Range("A1").Value = "Account ref. number"

If xCol > 2 Then
    Dim colRange As Range
    Set colRange = Columns(1).Resize(, xCol - 2)

    Dim blankCells As Range
    On Error Resume Next
    Set blankCells = colRange.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    If Not blankCells Is Nothing Then
        blankCells.FormulaR1C1 = "=RC[1]"
    End If
End If

Set ws = ActiveSheet

With ws.UsedRange
    firstRow = .Cells(1).Row
    lastRow = .Cells(.Cells.Count).Row
End With

Set rng = ws.Range("A1:A" & lastRow)
For lrow = lastRow To firstRow Step -1
    Set vRg = rng.Cells(lrow - firstRow + 1)
    If vRg.HasFormula Then
        vRg.Value = vRg.Value
    End If
    If vRg.Value = 0 Then
        vRg.Clear
    End If
Next lrow

On Error Resume Next
Set rng = rng.SpecialCells(xlCellTypeBlanks)
On Error GoTo 0

If Not rng Is Nothing Then
    rng.FormulaR1C1 = "=R[-1]C"
End If

For Each cell In rng
    cell.Value = Trim(cell)
Next cell

'Delete columns B thru xCol

If xCol > 2 Then
    Columns(2).Resize(, xCol - 2).EntireColumn.Select
    Selection.Delete
End If
Columns("A:A").EntireColumn.AutoFit

' Find the column number of Date
xCol = WorksheetFunction.Match("Date", Range("1:1"), 0)

' Trim the values in the range directly
Set rng = Range(Cells(1, xCol), Cells(lastRow, xCol))

For Each cell In rng
    cell.Value = Trim(cell.Value)
Next cell

' Collect the rows to be deleted in a range
For lrow = lastRow To firstRow Step -1
    Set vRg = Cells(lrow, xCol)
    If vRg.Value = "" Or vRg.Value = "Beginning Balance" Then
        If rowsToDelete Is Nothing Then
            Set rowsToDelete = vRg.EntireRow
        Else
            Set rowsToDelete = Union(rowsToDelete, vRg.EntireRow)
        End If
    End If
Next lrow

' Delete the collected rows
If Not rowsToDelete Is Nothing Then
    rowsToDelete.Delete
End If

' Format the Date Column
Columns(xCol).NumberFormat = "m/d/yyyy"

'Remove balance column

xStr = "Balance"
Set xRg = Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
xCol = xRg.Column
Columns(xCol).EntireColumn.Select
Selection.Delete

' Modify headers in the sheet

xStr = "Type"
Set xRg = Range("A1:AZ1").Find(What:=xStr, LookIn:=xlValues, LookAt:=xlPart)

If Not xRg Is Nothing Then
    xCol = xRg.Column
    Cells(1, xCol).Value = "Source"
Else
    ' Handle the case when a column header containing "Type" is not found
    MsgBox "No column header containing 'Type' was found.", vbExclamation
End If

xStr = "Date"
Set xRg = Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
xCol = xRg.Column
Cells(1, xCol).Value = "Posted Date"

xStr = "Num"
Set xRg = Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
xCol = xRg.Column
Cells(1, xCol).Value = "Posssible Journal ref. number"

'-------------------------------------------------Section to modify--------------------------------------------------------------------------------------------------------------------------- 

' Change "Amount" to "Original Amount" and add "Amount" column
Dim amountCol As Long', amountAdjustedCol As Long
Dim rowNum As Long, firstInstanceRow As Long

amountCol = WorksheetFunction.Match("Amount", Range("1:1"), 0)
'amountAdjustedCol = Cells(1, Columns.Count).End(xlToLeft).Column + 1
ws.Cells(1, amountCol).Value = "Amount"
'ws.Cells(1, amountAdjustedCol).Value = "Amount"
'ws.Cells(1, amountAdjustedCol).Interior.Color = RGB(255, 255, 0)

' Insert "Copy of Date" column
dateCol = GetColumnNumber(ws, "Posted Date")

If dateCol > 0 Then
    ws.Columns(dateCol + 1).Insert Shift:=xlToRight
    ws.Cells(1, dateCol + 1).Value = "Copy of Date"
    ws.Range(ws.Cells(2, dateCol + 1), ws.Cells(lastRow, dateCol + 1)).Value = ws.Range(ws.Cells(2, dateCol), ws.Cells(lastRow, dateCol)).Value
Else
    MsgBox "Column 'Posted Date' not found."
End If

' Insert "Comments" Column
Dim headers As Range, hdr As Range
Dim lastCol As Long
Dim memoCol As Long, splitCol As Long, commentsCol As Long
Dim headersToFind As Variant
Dim foundHeaders As Collection
Dim j As Long

Set ws = ActiveSheet  ' Set ws to the active sheet
headersToFind = Array("Memo", "Description", "Name", "Class")  ' Column headers to find
Set foundHeaders = New Collection

' Get the header row range
lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
Set headers = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))

' Initialize memoCol and splitCol
memoCol = 0
splitCol = 0

' Find the columns
For Each hdr In headers
    For i = 0 To UBound(headersToFind)
        If hdr.Value Like "*" & headersToFind(i) & "*" Then
            foundHeaders.Add hdr.Column
            If hdr.Value Like "*Memo*" Then
                memoCol = hdr.Column
            End If
        End If
    Next i
    If hdr.Value = "Split" Then
        splitCol = hdr.Column
    End If
Next hdr

' Determine where to insert the "Comments" column
If splitCol > 0 Then
    commentsCol = splitCol
ElseIf memoCol > 0 Then
    commentsCol = memoCol + 1
Else
    ' If neither "Split" nor "Memo" is found, handle accordingly (e.g., prompt the user or insert at a default location)
    MsgBox "Neither 'Split' nor 'Memo' column was found. Cannot insert 'Comments' column.", vbExclamation
    Exit Sub
End If

' Insert the "Comments" column
ws.Columns(commentsCol).Insert Shift:=xlToRight
ws.Cells(1, commentsCol).Value = "Comments"

    ' Concatenate the cell values
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Dim comment As String
        comment = ""
        For j = 1 To foundHeaders.Count
            comment = comment & " " & ws.Cells(i, foundHeaders(j)).Value
        Next j
        ws.Cells(i, commentsCol).Value = Trim(comment)
    Next i

' Update column numbers after adding new column
amountCol = GetColumnNumber(ws, "Amount")
'amountAdjustedCol = GetColumnNumber(ws, "Amount")

' Copy the values from "Original Amount" column to "Amount" column
'Range(Cells(2, amountCol), Cells(lastRow, amountCol)).Copy Destination:=Cells(2, amountAdjustedCol)

' Find the first instance of Account ref. number starting with "2", "Accounts Payable" or contains "A/P"
' For rowNum = 2 To lastRow
'     If Left(Cells(rowNum, 1).Value, 1) = "2" Or Cells(rowNum, 1).Value = "Accounts Payable" Or InStr(Cells(rowNum, 1).Value, "A/P") > 0 Then
'         firstInstanceRow = rowNum
'         Exit For
'     End If
' Next rowNum

' If the first instance is found, inverse the value and highlight the cell yellow
' If firstInstanceRow > 0 Then
'     For rowNum = firstInstanceRow To lastRow
'         If Left(Cells(rowNum, 1).Value, 1) >= "5" Or _
'         InStr(UCase(Cells(rowNum, 1).Value), "WAGES") > 0 Or _
'         InStr(UCase(Cells(rowNum, 1).Value), "COST OF GOODS") > 0 Or _
'         InStr(UCase(Cells(rowNum, 1).Value), "COGS") > 0 Then
'             Exit For
'         End If
'         Cells(rowNum, amountAdjustedCol).Value = -Cells(rowNum, amountCol).Value
'         Cells(rowNum, amountAdjustedCol).Interior.Color = RGB(255, 255, 0)
'     Next rowNum
' End If

' Format the "Amount" columns
Columns(amountCol).NumberFormat = "#,##0.00"
'Columns(amountAdjustedCol).NumberFormat = "#,##0.00"

' Summarize the Amount and Adjusted Amount Columns
' lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column + 1
' Cells(1, lastColumn).Value = "Total:"
' Cells(1, lastColumn + 1).Formula = "=SUM(" & Cells(2, amountCol).Address & ":" & Cells(lastRow + 1, amountCol).Address & ")"
' Cells(1, lastColumn + 2).Formula = "=SUM(" & Cells(2, amountAdjustedCol).Address & ":" & Cells(lastRow + 1, amountAdjustedCol).Address & ")"

' Apply Visual-formatting to the new results
' With Cells(1, lastColumn).Resize(, 3)
'     .Style = "Total"
'     .Font.Bold = True
' End With
' Cells(1, lastColumn + 2).Interior.Color = RGB(255, 255, 0)
' Range(Cells(1, 1), Cells(1, lastColumn - 1)).Font.Bold = True

' Check if the total of the Adjusted Amount column is close enough to zero
' If Abs(Cells(1, lastColumn + 2).Value) > 0.00001 Then
'     MsgBox "Adjusted Amount column does not balance out, please manually adjust for Completeness", vbExclamation
'     Exit Sub
' End If

' Summarize the accounts for Completeness
'     Dim accountRange As Range
'     Dim amountRange As Range
'     Dim summaryRange As Range
'     Dim summaryTable As ListObject
'     Dim uniqueAccounts As Collection
'     Dim account As Variant
'         Set ws = ActiveSheet
'     lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
' ' Set the ranges for Account and Amount columns
' Dim accountCol As Long

' accountCol = GetColumnNumber(ws, "Account ref. number")
' amountCol = GetColumnNumber(ws, "Amount")

' If accountCol > 0 And amountCol > 0 Then
'     Set accountRange = ws.Range(ws.Cells(2, accountCol), ws.Cells(lastRow, accountCol))
'     Set amountRange = ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol))
' Else
'     MsgBox "Could not find 'Account' or 'Amount' column. Please check the column headers.", vbExclamation
'     Exit Sub
' End If
    
'     ' Create a collection of Account Numberss
'     Set uniqueAccounts = New Collection
'         On Error Resume Next
'         For Each account In accountRange
'             uniqueAccounts.Add account.Value, CStr(account.Value)
'         Next account
'         On Error GoTo 0
    
'   ' Create Summary Worksheet
' Dim summarySheet As Worksheet
' Set summarySheet = ActiveWorkbook.Sheets.Add(After:=ActiveSheet)
' summarySheet.Name = "Summary"
' Set summaryRange = summarySheet.Range("A1:B" & uniqueAccounts.Count + 1)

'     ' Copy Account Numbers for total-column calculation
'     For i = 1 To uniqueAccounts.Count
'         summaryRange.Cells(i + 1, 1).Value = uniqueAccounts(i)
'         summaryRange.Cells(i + 1, 2).Formula = WorksheetFunction.SumIf(accountRange, uniqueAccounts(i), amountRange)
'     Next i
    
' ' Format summary table
' Set summaryTable = summarySheet.ListObjects.Add(xlSrcRange, summaryRange, , xlYes)
' summaryTable.Name = "AccountSummary"
' With summaryTable
'     .HeaderRowRange.Cells(1, 1).Value = "Account"
'     .HeaderRowRange.Cells(1, 2).Value = "Total"
'     .ListColumns(2).Range.NumberFormat = "#,##0.00"
'     .HeaderRowRange.Interior.ThemeColor = xlThemeColorAccent1
'     .HeaderRowRange.Interior.PatternColorIndex = xlAutomatic
' End With

' summaryTable.Range.Columns.AutoFit

' ' Convert table to a range
' summaryTable.Unlist

' ' Get_Account_Info Macro

'     Set ws = ActiveWorkbook.Worksheets("Summary") ' Set the target worksheet
'     ws.Visible = xlSheetVisible

'     'Insert row for Account Number
'     ws.Columns("A:A").Insert Shift:=xlToRight
'     ws.Range("A1").FormulaR1C1 = "Account Number"
'     wsLastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row 'get last row of data
'     ws.Range("A2:A" & wsLastRow).FormulaR1C1 = "=LEFT(RC[1],FIND("" "",RC[1],1)-1)"
'     ws.Range("A2:A" & wsLastRow).Value = ws.Range("A2:A" & wsLastRow).Value 'replace formulas with values
    
'    'Insert row for Account Description
' ws.Columns("B:B").Insert Shift:=xlToRight
' ws.Range("B1").FormulaR1C1 = "Account Description"
' wsLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row 'get last row of data
' ws.Range("B2:B" & wsLastRow).FormulaR1C1 = "=TRIM(RIGHT(RC[1],LEN(RC[1])-SEARCH("" "",RC[1])))"
' ws.Range("B2:B" & wsLastRow).Value = ws.Range("B2:B" & wsLastRow).Value 'replace formulas with values

' 'Checks for abnormal characters
' Dim descCell As Range
' For Each descCell In ws.Range("B2:B" & wsLastRow)
'     If Left(descCell.Value, 2) = "· " Or Left(descCell.Value, 2) = "- " Then
'         descCell.Value = Right(descCell.Value, Len(descCell.Value) - 2)
'     End If
' Next descCell

' 'Removes any instances of Multiple spaces
' For Each descCell In ws.Range("B2:B" & wsLastRow)
'     While InStr(descCell.Value, "  ") > 0
'         descCell.Value = Replace(descCell.Value, "  ", " ")
'     Wend
' Next descCell

' 'Removes Beginning and ending spaces
' For Each descCell In ws.Range("B2:B" & wsLastRow)
'     descCell.Value = Trim(descCell.Value)
' Next descCell
    
'     'Copy format from column C to columns A and B
'     ws.Columns("C:C").Copy
'     ws.Columns("A:A").PasteSpecial Paste:=xlPasteFormats
'     Application.CutCopyMode = False
    
'     ws.Columns("C:C").Copy
'     ws.Columns("B:B").PasteSpecial Paste:=xlPasteFormats
'     Application.CutCopyMode = False

'     ' Swap columns C and D
'     Dim tempCol As Range
'     ws.Columns("E:E").Insert Shift:=xlToRight
'     ws.Columns("C:C").Copy Destination:=ws.Columns("E:E")
'     ws.Columns("D:D").Copy Destination:=ws.Columns("C:C")
'     ws.Columns("E:E").Copy Destination:=ws.Columns("D:D")
'     ws.Columns("E:E").Delete
'     Application.CutCopyMode = False

'     Application.ScreenUpdating = True

' 'Paste the GL summary into the Completeness test template
' Dim GLSummary As Range
' lastRow = summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).Row
' Set GLSummary = summarySheet.Range("A2:C" & lastRow)
' GLSummary.Copy

' ' Open the target workbook
' Dim targetWorkbook As Workbook
' Dim originalWorkbook As Workbook

' Set originalWorkbook = ThisWorkbook
' Set targetWorkbook = Workbooks.Open("C:\*YourFilePath*)")
' On Error Resume Next

' Dim targetWs As Worksheet
' Set targetWs = targetWorkbook.Worksheets("GL")

' targetWs.Range("GL_Input").ClearContents
' targetWs.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
' GLSummary.Copy
' targetWs.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

' ' Save Workbook after opening completeness
' Dim TicketNo As String
' Dim refFile As String
' Dim text As String
' Dim textline As String
' Dim Slocation As String

' ' Prompt for ticket number
' TicketNo = Trim(InputBox("What is the ticket number?"))

' ' Reference file with S drive locations
' refFile = "S:\Fileapth\" & TicketNo & "\" & TicketNo & ".txt"

' On Error GoTo ErrMsg1
' Open refFile For Input As #1
'     Do Until EOF(1)
'         Line Input #1, textline
'         text = text & textline
'     Loop
' Close #1

' Dim locationstring As Integer
' locationstring = InStr(text, "SDrivePath:")

' If locationstring <> 0 Then
'     Slocation = Mid(text, locationstring + 12, InStr(text, "ClientName:") - InStr(text, "SDrivePath:") - 12)

'     On Error Resume Next
'     originalWorkbook.SaveAs Filename:= _
'         Slocation & originalWorkbook.Name & " - " & TicketNo & " (Format)" & ".xlsx" _
'         , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

' Else
'     MsgBox "Ticket No. is too old to be saved to the ticket folder. Please save manually"
' End If

' Exit Sub

' ErrMsg1:
'     MsgBox "Possible causes of this error:" & vbCrLf & "Ticket No. entered is not valid" & vbCrLf & "Can't connect to the S drive"

End Sub

Sub QBFormatDCB()

Application.ScreenUpdating = False

'Remove QB Export Tips sheet (if exists)
Application.DisplayAlerts = False
For Each sheet In ActiveWorkbook.Worksheets
    If sheet.Name Like "*Tips" Then
        sheet.Delete
    End If
Next sheet
Application.DisplayAlerts = True

'Make whole sheet normal formatting
With ActiveSheet.Cells
    .Style = "Normal"
    .UnMerge
End With

    'Remove any rows above the headers (search for "Date")
    Dim rng As Range
    Dim dateCell As Range
    Dim rowsToDelete As Range
    Dim i As Long
    Dim lastRow As Long

    With ActiveSheet
        lastRow = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        Set rng = .Range("A1:AZ" & lastRow)

        Set dateCell = rng.Find(What:="Date", LookAt:=xlPart, MatchCase:=False) 'non case-sensitive, and match part of cell content

        If Not dateCell Is Nothing Then ' If "Date" was found
            For i = dateCell.Row - 1 To 1 Step -1 ' Loop upwards from "Date" to the first row
                .Rows(i).Delete ' Delete the row
            Next i
        End If
    End With

'Remove all blank columns
Dim ws As Worksheet
Dim lastColumn As Long
Dim columnNumber As Long

Set ws = ActiveSheet
lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

For columnNumber = lastColumn To 1 Step -1
    If WorksheetFunction.CountA(ws.Columns(columnNumber)) = 0 Then
        ws.Columns(columnNumber).Delete
    End If
Next columnNumber

'Set xCol as first column with a header

Dim xRg As Range
Dim xStr As String
Dim xCol As Integer
Dim vRg As Range
Dim firstRow As Long
Dim lrow As Long

xCol = Range("A1").End(xlToRight).Column

'A1 = "Account ref. number"

Range("A1").Value = "Account ref. number"

If xCol > 2 Then
    Dim colRange As Range
    Set colRange = Columns(1).Resize(, xCol - 2)

    Dim blankCells As Range
    On Error Resume Next
    Set blankCells = colRange.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    If Not blankCells Is Nothing Then
        blankCells.FormulaR1C1 = "=RC[1]"
    End If
End If

Set ws = ActiveSheet

With ws.UsedRange
    firstRow = .Cells(1).Row
    lastRow = .Cells(.Cells.Count).Row
End With

Set rng = ws.Range("A1:A" & lastRow)
For lrow = lastRow To firstRow Step -1
    Set vRg = rng.Cells(lrow - firstRow + 1)
    If vRg.HasFormula Then
        vRg.Value = vRg.Value
    End If
    If vRg.Value = 0 Then
        vRg.Clear
    End If
Next lrow

On Error Resume Next
Set rng = rng.SpecialCells(xlCellTypeBlanks)
On Error GoTo 0

If Not rng Is Nothing Then
    rng.FormulaR1C1 = "=R[-1]C"
End If

For Each cell In rng
    cell.Value = Trim(cell)
Next cell

'Delete columns B thru xCol

If xCol > 2 Then
    Columns(2).Resize(, xCol - 2).EntireColumn.Select
    Selection.Delete
End If
Columns("A:A").EntireColumn.AutoFit

' Find the column number of Date
xCol = WorksheetFunction.Match("Date", Range("1:1"), 0)

' Trim the values in the range directly
Set rng = Range(Cells(1, xCol), Cells(lastRow, xCol))

For Each cell In rng
    cell.Value = Trim(cell.Value)
Next cell

' Collect the rows to be deleted in a range
For lrow = lastRow To firstRow Step -1
    Set vRg = Cells(lrow, xCol)
    If vRg.Value = "" Or vRg.Value = "Beginning Balance" Then
        If rowsToDelete Is Nothing Then
            Set rowsToDelete = vRg.EntireRow
        Else
            Set rowsToDelete = Union(rowsToDelete, vRg.EntireRow)
        End If
    End If
Next lrow

' Delete the collected rows
If Not rowsToDelete Is Nothing Then
    rowsToDelete.Delete
End If

' Format the Date Column
Columns(xCol).NumberFormat = "m/d/yyyy"

'Remove balance column

xStr = "Balance"
Set xRg = Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
xCol = xRg.Column
Columns(xCol).EntireColumn.Select
Selection.Delete

' Create "Amount" column
Dim amountCol As Long
amountCol = Cells(1, Columns.Count).End(xlToLeft).Column + 1
Cells(1, amountCol).Value = "Amount"
Cells(2, amountCol).FormulaR1C1 = "=RC[-2]-RC[-1]"
lastRow = ActiveSheet.Cells.SpecialCells(xlLastCell).Row
Range(Cells(2, amountCol), Cells(lastRow, amountCol)).FillDown
Columns(amountCol - 2).Resize(, 3).NumberFormat = "#,##0.00"

' Summarize the Amount column
lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column + 1
Cells(1, lastColumn).Value = "Total:"
Cells(2, lastColumn).Formula = "=SUM(" & Range(Cells(2, amountCol), Cells(lastRow, amountCol)).Address & ")"

' Modify headers in the sheet
xStr = "Type"
Set xRg = Range("A1:AZ1").Find(What:=xStr, LookIn:=xlValues, LookAt:=xlPart)

If Not xRg Is Nothing Then
    xCol = xRg.Column
    Cells(1, xCol).Value = "Source"
Else
    ' Handle the case when a column header containing "Type" is not found
    MsgBox "No column header containing 'Type' was found.", vbExclamation
End If

xStr = "Date"
Set xRg = Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
xCol = xRg.Column
Cells(1, xCol).Value = "Posted Date"

xStr = "Num"
Set xRg = Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
xCol = xRg.Column
Cells(1, xCol).Value = "Possible Journal ref. number"

' Insert "Copy of Date" column
dateCol = GetColumnNumber(ws, "Posted Date")
ws.Columns(dateCol + 1).Insert Shift:=xlToRight
ws.Cells(1, dateCol + 1).Value = "Copy of Date"
ws.Range(ws.Cells(2, dateCol + 1), ws.Cells(lastRow, dateCol + 1)).Value = ws.Range(ws.Cells(2, dateCol), ws.Cells(lastRow, dateCol)).Value

' Insert "Comments" Column
Dim headers As Range, hdr As Range
Dim lastCol As Long
Dim memoCol As Long, splitCol As Long, commentsCol As Long
Dim headersToFind As Variant
Dim foundHeaders As Collection
Dim j As Long

Set ws = ActiveSheet  ' Set ws to the active sheet
headersToFind = Array("Memo", "Description", "Name", "Class")  ' Column headers to find
Set foundHeaders = New Collection

' Get the header row range
lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
Set headers = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))

' Initialize memoCol and splitCol
memoCol = 0
splitCol = 0

' Find the columns
For Each hdr In headers
    For i = 0 To UBound(headersToFind)
        If hdr.Value Like "*" & headersToFind(i) & "*" Then
            foundHeaders.Add hdr.Column
            If hdr.Value Like "*Memo*" Then
                memoCol = hdr.Column
            End If
        End If
    Next i
    If hdr.Value = "Split" Then
        splitCol = hdr.Column
    End If
Next hdr

' Determine where to insert the "Comments" column
If splitCol > 0 Then
    commentsCol = splitCol
ElseIf memoCol > 0 Then
    commentsCol = memoCol + 1
Else
    ' If neither "Split" nor "Memo" is found, handle accordingly (e.g., prompt the user or insert at a default location)
    MsgBox "Neither 'Split' nor 'Memo' column was found. Cannot insert 'Comments' column.", vbExclamation
    Exit Sub
End If

' Insert the "Comments" column
ws.Columns(commentsCol).Insert Shift:=xlToRight
ws.Cells(1, commentsCol).Value = "Comments"
    ' Concatenate the cell values
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Dim comment As String
        comment = ""
        For j = 1 To foundHeaders.Count
            comment = comment & " " & ws.Cells(i, foundHeaders(j)).Value
        Next j
        ws.Cells(i, commentsCol).Value = Trim(comment)
    Next i

' Check if the total of the Amount column is close enough to zero
If Abs(Cells(2, lastColumn + 2).Value) > 0.001 Then
    MsgBox "Amount column does not balance out. Current total is " & Cells(2, lastColumn).Value & ". Please manually adjust for Completeness", vbExclamation
    Exit Sub
End If

' Summarize the accounts for Completeness
    Dim accountRange As Range
    Dim amountRange As Range
    Dim summaryRange As Range
    Dim summaryTable As ListObject
    Dim uniqueAccounts As Collection
    Dim account As Variant
        Set ws = ActiveSheet
   lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
' Set the ranges for Account and Amount columns
Dim accountCol As Long

accountCol = GetColumnNumber(ws, "Account ref. number")
amountCol = GetColumnNumber(ws, "Amount")

If accountCol > 0 And amountCol > 0 Then
    Set accountRange = ws.Range(ws.Cells(2, accountCol), ws.Cells(lastRow, accountCol))
    Set amountRange = ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol))
Else
    MsgBox "Could not find 'Account' or 'Amount' column. Please check the column headers.", vbExclamation
    Exit Sub
End If

    ' Create a collection tof Account Numberss
    Set uniqueAccounts = New Collection
        On Error Resume Next
        For Each account In accountRange
            uniqueAccounts.Add account.Value, CStr(account.Value)
        Next account
        On Error GoTo 0
    
    ' Create Summary Worksheet
    Dim summarySheet As Worksheet
    Set summarySheet = ThisWorkbook.Sheets.Add(After:=ws)
    summarySheet.Name = "Summary"
    Set summaryRange = summarySheet.Range("A1:B" & uniqueAccounts.Count + 1)
    
    ' Copy Account Numbers for total-column calculation
    For i = 1 To uniqueAccounts.Count
        summaryRange.Cells(i + 1, 1).Value = uniqueAccounts(i)
        summaryRange.Cells(i + 1, 2).Formula = WorksheetFunction.SumIf(accountRange, uniqueAccounts(i), amountRange)
    Next i
    
    ' Format summary table
    Set summaryTable = summarySheet.ListObjects.Add(xlSrcRange, summaryRange, , xlYes)
    summaryTable.Name = "AccountSummary"
    With summaryTable
        .HeaderRowRange.Cells(1, 1).Value = "Account"
        .HeaderRowRange.Cells(1, 2).Value = "Total"
      .ListColumns(2).Range.NumberFormat = "#,##0.00"
        .HeaderRowRange.Interior.ThemeColor = xlThemeColorAccent1
        .HeaderRowRange.Interior.PatternColorIndex = xlAutomatic
    End With

    summaryTable.Range.Columns.AutoFit

' Convert table to a range
    summaryTable.Unlist

' Get_Account_Info Macro

    Set ws = ActiveWorkbook.Worksheets("Summary") ' Set the target worksheet
    ws.Visible = xlSheetVisible

    'Insert row for Account Number
    ws.Columns("A:A").Insert Shift:=xlToRight
    ws.Range("A1").FormulaR1C1 = "Account Number"
    wsLastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row 'get last row of data
    ws.Range("A2:A" & wsLastRow).FormulaR1C1 = "=LEFT(RC[1],FIND("" "",RC[1],1)-1)"
    ws.Range("A2:A" & wsLastRow).Value = ws.Range("A2:A" & wsLastRow).Value 'replace formulas with values
    
'Insert row for Account Description
ws.Columns("B:B").Insert Shift:=xlToRight
ws.Range("B1").FormulaR1C1 = "Account Description"
wsLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row 'get last row of data
ws.Range("B2:B" & wsLastRow).FormulaR1C1 = "=TRIM(RIGHT(RC[1],LEN(RC[1])-SEARCH("" "",RC[1])))"
ws.Range("B2:B" & wsLastRow).Value = ws.Range("B2:B" & wsLastRow).Value 'replace formulas with values

'Checks for abnormal characters
Dim descCell As Range
For Each descCell In ws.Range("B2:B" & wsLastRow)
    If Left(descCell.Value, 2) = "· " Or Left(descCell.Value, 2) = "- " Then
        descCell.Value = Right(descCell.Value, Len(descCell.Value) - 2)
    End If
Next descCell

'Removes any instances of Multiple spaces
For Each descCell In ws.Range("B2:B" & wsLastRow)
    While InStr(descCell.Value, "  ") > 0
        descCell.Value = Replace(descCell.Value, "  ", " ")
    Wend
Next descCell

'Removes Beginning and ending spaces
For Each descCell In ws.Range("B2:B" & wsLastRow)
    descCell.Value = Trim(descCell.Value)
Next descCell
    
    'Copy format from column C to columns A and B
    ws.Columns("C:C").Copy
    ws.Columns("A:A").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ws.Columns("C:C").Copy
    ws.Columns("B:B").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Swap columns C and D
    Dim tempCol As Range
    ws.Columns("E:E").Insert Shift:=xlToRight
    ws.Columns("C:C").Copy Destination:=ws.Columns("E:E")
    ws.Columns("D:D").Copy Destination:=ws.Columns("C:C")
    ws.Columns("E:E").Copy Destination:=ws.Columns("D:D")
    ws.Columns("E:E").Delete
    Application.CutCopyMode = False

    Application.ScreenUpdating = True

'Paste the GL summary into the Completeness test template
Dim GLSummary As Range
lastRow = summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).Row
Set GLSummary = summarySheet.Range("A2:C" & lastRow)
GLSummary.Copy

' Open the target workbook
Dim targetWorkbook As Workbook
Dim originalWorkbook As Workbook

Set originalWorkbook = ThisWorkbook
Set targetWorkbook = Workbooks.Open("C:\*YourFilePath*)")
On Error Resume Next

Dim targetWs As Worksheet
Set targetWs = targetWorkbook.Worksheets("GL")

targetWs.Range("GL_Input").ClearContents
targetWs.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
GLSummary.Copy
targetWs.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

' Save Workbook after opening completeness
Dim TicketNo As String
Dim refFile As String
Dim text As String
Dim textline As String
Dim Slocation As String

' Prompt for ticket number
TicketNo = Trim(InputBox("What is the ticket number?"))

' Reference file with S drive locations
refFile = "S:\Fileapth\" & TicketNo & "\" & TicketNo & ".txt"

On Error GoTo ErrMsg1
Open refFile For Input As #1
    Do Until EOF(1)
        Line Input #1, textline
        text = text & textline
    Loop
Close #1

Dim locationstring As Integer
locationstring = InStr(text, "SDrivePath:")

If locationstring <> 0 Then
    Slocation = Mid(text, locationstring + 12, InStr(text, "ClientName:") - InStr(text, "SDrivePath:") - 12)

    On Error Resume Next
    originalWorkbook.SaveAs Filename:= _
        Slocation & originalWorkbook.Name & " - " & TicketNo & " (Format)" & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

Else
    MsgBox "Ticket No. is too old to be saved to the ticket folder. Please save manually"
End If

Exit Sub

ErrMsg1:
    MsgBox "Possible causes of this error:" & vbCrLf & "Ticket No. entered is not valid" & vbCrLf & "Can't connect to the S drive"

End Sub