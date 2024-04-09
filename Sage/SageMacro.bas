' V 1.0.0 by Matthew Chirichella
' This Macro is designed to easily format and prepare ledgers produced by Sage Accounting Software for further analysis.
Sub SageMacro()
    Dim keywords As Variant

    ' Optimization
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    ActiveWindow.DisplayGridlines = True  ' Display gridlines

    ' Define the keywords to search for
    keywords = Array("Posted dt.", "Doc dt.", "Doc", "Memo/Description", "Department", "JNL", "Debit", "Credit")

    ' Call the DeleteRowsAboveKeywords
    Call DeleteRowsAboveKeywords(keywords)

    ' Call the FormatWorksheet subroutine
    Call FormatWorksheet

    ' Call the FilterAndFillData subroutine
    Call FilterAndFillData

    ' Call the MakeFinal subroutine
    Call MakeFinal

    ' Create the summary sheet
    Call MakeSummarySheet

    ' Call GetAccountInfo subroutine to modify the Summary sheet format
    Call GetAccountInfo 

    ' Turn on Screen Updating
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

Private Sub DeleteRowsAboveKeywords(keywords As Variant)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim firstInstanceRow As Long
    firstInstanceRow = ws.Rows.Count
    
    Dim keyword As Variant
    Dim rng As Range, foundCell As Range
    Dim searchRange As Range
    Set searchRange = ws.UsedRange
    
    For Each keyword In keywords
        Set foundCell = searchRange.Find(What:=keyword, LookIn:=xlValues, LookAt:=xlPart)
        If Not foundCell Is Nothing Then
            If foundCell.row < firstInstanceRow Then firstInstanceRow = foundCell.row
        End If
    Next keyword
    
    If firstInstanceRow < ws.Rows.Count Then
        ws.Rows("1:" & firstInstanceRow - 1).Delete
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    'MsgBox "Finished DeleteAboveRows"
End Sub

Sub FormatWorksheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' Define and set the worksheet object

    Dim cell As Range
    Dim blankCols As Range, blankRows As Range

    ' Insert a new column at the beginning of the worksheet
    ws.Columns("A:A").Insert Shift:=xlToRight

    ' Set the value of the first cell in the new column to "Account ref. number"
    ws.Range("A1").Value = "Account ref. number"
    
    ' Consistent Header Formats
    ws.Range("C1").Copy
    ws.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Iterate through headers and modify name
    For Each cell In ws.Rows("1:1").Cells
        Select Case cell.Value
            Case "Posted dt."
                cell.Value = "Posted Date"
            Case "Doc dt."
                cell.Value = "Journal ref. number"
            Case "Doc"
                cell.Value = "Possible Journal ref."
            Case "Memo/Description"
                cell.Value = "Comments"
            Case "JNL"
                cell.Value = "Source"
            ' Additional Cases for future
        End Select
    Next cell

    ' Implement Batch Delete Sub (Optimized by 450% compared to individual row deletion)
    DeleteBlankRowsAndColumns ws
End Sub

Sub DeleteBlankRowsAndColumns(ws As Worksheet)
    Dim blankCols As Range, blankRows As Range
    
    ' Find all blank columns and rows
    Set blankCols = FindBlankColumns(ws.UsedRange)
    Set blankRows = FindBlankRows(ws.UsedRange)
    
    ' Disable Screen Updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Perform batch deletion
    If Not blankCols Is Nothing Then blankCols.EntireColumn.Delete
    If Not blankRows Is Nothing Then blankRows.EntireRow.Delete
    
    ' Re-enable Updating
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub FilterAndFillData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Disable Screen Updating
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    ' Assuming headers are in the first row and data starts from the second row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row ' Define Journal ref. number column as the reference
    
    ' Filter range for blanks and "total" in the "Journal ref. number" and "Posted Date" columns
    ws.Range("A1:C" & lastRow).AutoFilter Field:=3, Criteria1:="=" 
    ws.Range("A1:C" & lastRow).AutoFilter Field:=2, Criteria1:="<>*total*", Operator:=xlAnd

    'Remove "(Balance forward As of DD/MM/YYYY)"
    Dim rng As Range
    Set rng = ws.Range("B2:B" & lastRow)
    For Each cell In rng
        If InStr(cell.Value, " (Balance forward As of ") > 0 Then
            cell.Value = Left(cell.Value, InStr(cell.Value, " (Balance forward As of ") - 1)
        End If
    Next cell

    ' Fill Visible Values
    Dim visibleCells As Range
    On Error Resume Next
    Set visibleCells = ws.Range("B2:B" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleCells Is Nothing Then
        visibleCells.Copy
        ws.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        ' Clear Originally-filled Cells
        visibleCells.ClearContents
    End If
    
    ' Remove filters
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' Fill down all Accounts
    ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
    ws.Range("A2:A" & lastRow).Value = ws.Range("A2:A" & lastRow).Value ' Convert formulas to values

    ' Screen Updating
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

Sub MakeFinal()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow As Long, debitCol As Long, creditCol As Long, balanceCol As Long, amountCol As Long
    
    ' Screen Updating
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    ' Delete Excess Rows
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    ws.Range("C1:C" & lastRow).AutoFilter Field:=1, Criteria1:="="
    ws.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ' Find/Define column positions
    debitCol = 0
    creditCol = 0
    balanceCol = 0
    
    On Error Resume Next ' Error Handling
    debitCol = ws.Rows(1).Find("Debit", LookIn:=xlValues, LookAt:=xlWhole).Column
    creditCol = ws.Rows(1).Find("Credit", LookIn:=xlValues, LookAt:=xlWhole).Column
    balanceCol = ws.Rows(1).Find("Balance", LookIn:=xlValues, LookAt:=xlWhole).Column
    On Error GoTo 0 
    
    ' Logic to Insert Amount Column
    If creditCol > 0 Then
        amountCol = creditCol + 1
    ElseIf balanceCol > 0 Then
        amountCol = balanceCol
    Else
        amountCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    End If
    
    ws.Columns(amountCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Cells(1, amountCol).Value = "Amount"
    
    ' Copy Header Format
    ws.Cells(1, amountCol - 1).Copy
    ws.Cells(1, amountCol).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ' Define end of where to Fill Amount Column
    Dim lastRowDebit As Long
    Dim lastRowCredit As Long
    lastRowDebit = ws.Cells(ws.Rows.Count, debitCol).End(xlUp).row
    lastRowCredit = ws.Cells(ws.Rows.Count, creditCol).End(xlUp).row
        ' Use the larger of the two rows to ensure all data is covered
    Dim loopLastRow As Long
    loopLastRow = Application.WorksheetFunction.Max(lastRowDebit, lastRowCredit)

    If debitCol > 0 And creditCol > 0 Then ' Ensure both Debit and Credit columns were found
        For i = 2 To loopLastRow
            ' Check if either Debit or Credit cell is empty, treat as 0 for calculation
            Dim debitVal As Variant, creditVal As Variant
            debitVal = IIf(IsEmpty(ws.Cells(i, debitCol).Value), 0, ws.Cells(i, debitCol).Value)
            creditVal = IIf(IsEmpty(ws.Cells(i, creditCol).Value), 0, ws.Cells(i, creditCol).Value)

            ' Directly calculate and set the value
            ws.Cells(i, amountCol).Value = debitVal - creditVal
        Next i
    Else
        MsgBox "Debit or Credit column not found.", vbExclamation
    End If

    ' Convert formulas to values
    With ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol))
        .Value = .Value
    End With
    
    ' Screen Updating
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

Sub MakeSummarySheet()
    Dim ws As Worksheet, summarySheet As Worksheet
    Dim lastRow As Long, accountCol As Long, amountCol As Long
    Dim accountRange As Range, amountRange As Range, summaryRange As Range
    Dim summaryTable As ListObject
    Dim uniqueAccounts As Collection, account As Variant
    Dim i As Integer

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' Set the ranges for Account and Amount columns using the GetColumnNumber function
    accountCol = GetColumnNumber(ws, "Account ref. number")
    amountCol = GetColumnNumber(ws, "Amount")

    If accountCol > 0 And amountCol > 0 Then
        Set accountRange = ws.Range(ws.Cells(2, accountCol), ws.Cells(lastRow, accountCol))
        Set amountRange = ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol))
    Else
        MsgBox "Could not find 'Account ref. number' or 'Amount' column. Please check the column headers.", vbExclamation
        Exit Sub
    End If

    ' Create a collection of unique Accounts
    Set uniqueAccounts = New Collection
    On Error Resume Next ' Ignore duplicates
    For Each account In accountRange
        uniqueAccounts.Add account.Value, CStr(account.Value)
    Next account
    On Error GoTo 0

    ' Create Summary Worksheet
    Set summarySheet = ActiveWorkbook.Sheets.Add(After:=ActiveSheet)
    summarySheet.Name = "Summary"
    Set summaryRange = summarySheet.Range("A1:B" & uniqueAccounts.Count + 1)

   ' Copy Account Numbers for total-column calculation
    summarySheet.Cells(1, 1).Value = "Account"
    summarySheet.Cells(1, 2).Value = "Total"

    ' Copy Account Numbers for total-column calculation
    For i = 1 To uniqueAccounts.Count
        summaryRange.Cells(i + 1, 1).Value = uniqueAccounts(i)
        summaryRange.Cells(i + 1, 2).Formula = WorksheetFunction.SumIf(accountRange, uniqueAccounts(i), amountRange)
    Next i

    ' Format summary table
    Set summaryTable = summarySheet.ListObjects.Add(xlSrcRange, summaryRange, , xlYes)
    summaryTable.Name = "AccountSummary"
    With summaryTable
        .ListColumns(2).Range.NumberFormat = "#,##0.00"
        .HeaderRowRange.Interior.ThemeColor = xlThemeColorAccent1
        .HeaderRowRange.Interior.PatternColorIndex = xlAutomatic
    End With

    summaryTable.Range.Columns.AutoFit

    summaryTable.Unlist ' Convert to range from Table
End Sub

Sub GetAccountInfo()
    Dim ws As Worksheet, wsLastRow As Long
    Dim descCell As Range

    Application.ScreenUpdating = False ' Screen Updating
    
    Set ws = ActiveWorkbook.Worksheets("Summary") ' Target worksheet
    ws.Visible = xlSheetVisible

    ' Insert column for Account Number and set header
    ws.Columns("A:A").Insert Shift:=xlToRight
    ws.Range("A1").Value = "Account Number"
    wsLastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' Get last row of adjusted data
    ws.Range("A2:A" & wsLastRow).FormulaR1C1 = "=LEFT(RC[1], FIND("" - "",RC[1])-1)"
    ws.Range("A2:A" & wsLastRow).Value = ws.Range("A2:A" & wsLastRow).Value ' Convert formulas to values

    ' Insert column for Account Description and set header
    ws.Columns("B:B").Insert Shift:=xlToRight
    ws.Range("B1").Value = "Account Description"

    ' Assuming description might be a problem
    For Each descCell In ws.Range("C2:C" & wsLastRow)
        Dim accountInfo As String
        accountInfo = descCell.Value
        If InStr(1, accountInfo, "-") > 0 Then
            descCell.Offset(0, -1).Value = Trim(Mid(accountInfo, InStr(1, accountInfo, "-") + 1))
        Else
            descCell.Offset(0, -1).Value = accountInfo ' Fallback in case no hyphen is found
        End If
    Next descCell

    ' Post-processing for Account Description
    For Each descCell In ws.Range("B2:B" & wsLastRow)
        If Left(descCell.Value, 2) = "· " Or Left(descCell.Value, 2) = "- " Then
            descCell.Value = Right(descCell.Value, Len(descCell.Value) - 2)
        End If
        ' Remove extra spaces
        While InStr(descCell.Value, "  ") > 0
            descCell.Value = Replace(descCell.Value, "  ", " ")
        Wend
        ' Trim leading and trailing spaces
        descCell.Value = Trim(descCell.Value)
    Next descCell

    ' Copy format
    ws.Columns("C:C").Copy
    ws.Range("A:A").PasteSpecial Paste:=xlPasteFormats
    ws.Range("B:B").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ' Additional steps (e.g., swapping columns) will be Included

    Application.ScreenUpdating = True ' Restore screen updating
End Sub

Private Function GetColumnNumber(ws As Worksheet, columnName As String) As Long
    Dim rng As Range
    On Error Resume Next ' In case the column is not found
    Set rng = ws.Rows(1).Find(What:=columnName, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    On Error GoTo 0 ' Turn back on regular error handling

    If Not rng Is Nothing Then
        GetColumnNumber = rng.Column
    Else
        GetColumnNumber = 0 ' Indicate that the column was not found
    End If
End Function