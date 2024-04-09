' QB Combined Macros Version 2.1.2 by Matthew Chirichella

' Disgusting combined version of the macros For Amt & DC variants
' Shared Drive could not store Macros because reasons & it all has to be contained in one Sub :(
    ' Sorry to my Software Engineering professors

Sub CompleteQBMacro()
    Application.ScreenUpdating = False

    ' Define variables for both subvariants
    Dim ws As Worksheet, rng As Range, dateCell As Range, rowsToDelete As Range
    Dim debitsCol As Long, creditsCol As Long, amountCol As Long, columnNumber As Long
    Dim xRg As Range, xStr As String, xCol As Integer, vRg As Range
    Dim firstRow As Long, lastRow As Long, lastColumn As Long, memoCol As Long, splitCol As Long, commentsCol As Long
    Dim headers() As Variant, header As Variant, i As Long, j As Long, hdr As Range
    Dim accountRange As Range, amountRange As Range, summaryRange As Range, summaryTable As ListObject, uniqueAccounts As Collection, account As Variant
    Dim summarySheet As Worksheet, descCell As Range, tempCol As Range, GLSummary As Range, targetWorkbook As Workbook, originalWorkbook As Workbook
    Dim targetWs As Worksheet, TicketNo As String, refFile As String, text As String, textline As String, Slocation As String, locationstring As Integer
    'Ones I missed
    Dim lastCol As Long, headersToFind As Variant, foundHeaders As Collection, headersRange As Range, accountCol As Long, amountAdjustedCol As Long
    Dim firstInstanceRow As Long, rowNum As Long
    Dim comment As String
    ' Set the worksheet
    Set ws = ActiveSheet

    ' Initialize column numbers
    debitsCol = 0
    creditsCol = 0
    amountCol = 0

' Define the headers to find
headers = Array("Debit", "Credit", "Amount")

' First, ensure headers are in the first row by moving them if necessary
Dim headerRow As Long
headerRow = ws.Rows.Count

For Each header In headers
    Set rng = ws.Cells.Find(What:=header, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        If rng.Row <> 1 Then
            ' Move header to the first row
            ws.Rows(1).Cells(1, rng.Column).Value = header
            ' Clear the original header cell
            rng.Value = ""
        End If
    End If
Next header

' Refresh firstRow as headers are now in the first row
firstRow = 1

' Get column numbers again in case they were moved
For columnNumber = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If Trim(ws.Cells(1, columnNumber).Value) = "Debit" Then debitsCol = columnNumber
    If Trim(ws.Cells(1, columnNumber).Value) = "Credit" Then creditsCol = columnNumber
    If Trim(ws.Cells(1, columnNumber).Value) = "Amount" Then amountCol = columnNumber
Next columnNumber

    ' Check the format and run appropriate formatting logic
    If debitsCol > 0 And creditsCol > 0 Then
        ' Logic from QBFormatDCB
        Application.DisplayAlerts = False
        For Each Sheet In ActiveWorkbook.Worksheets
            If Sheet.Name Like "*Tips" Then
                Sheet.Delete
            End If
        Next Sheet
        Application.DisplayAlerts = True

        ' Make whole sheet normal formatting
        With ws.Cells
            .Style = "Normal"
            .UnMerge
        End With

        ' Remove any rows above the headers (search for "Date")
        lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        Set rng = ws.Range("A1:AZ" & lastRow)

        Set dateCell = rng.Find(What:="Date", LookAt:=xlPart, MatchCase:=False) ' Non case-sensitive, and match part of cell content

        If Not dateCell Is Nothing Then ' If "Date" was found
            For i = dateCell.Row - 1 To 1 Step -1 ' Loop upwards from "Date" to the first row
                ws.Rows(i).Delete ' Delete the row
            Next i
        End If

        ' Remove all blank columns
        lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

        For columnNumber = lastColumn To 1 Step -1
            If WorksheetFunction.CountA(ws.Columns(columnNumber)) = 0 Then
                ws.Columns(columnNumber).Delete
            End If
        Next columnNumber

        ' Set xCol as first column with a header
        xCol = ws.Range("A1").End(xlToRight).Column

        ' A1 = "Account ref. number"
        ws.Range("A1").Value = "Account ref. number"

         If xCol > 2 Then
            Set colRange = ws.Columns(1).Resize(, xCol - 2)
            On Error Resume Next
            Set blankCells = colRange.SpecialCells(xlCellTypeBlanks)
            On Error GoTo 0
            If Not blankCells Is Nothing Then
                blankCells.FormulaR1C1 = "=RC[1]"
            End If
        End If

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
            cell.Value = Trim(cell.Value)
        Next cell

        ' Delete columns B thru xCol
        If xCol > 2 Then
            ws.Columns(2).Resize(, xCol - 2).EntireColumn.Delete
        End If
        ws.Columns("A:A").EntireColumn.AutoFit

        ' Find the column number of Date
        xCol = WorksheetFunction.Match("Date", ws.Range("1:1"), 0)

        ' Trim the values in the range directly
        Set rng = ws.Range(ws.Cells(1, xCol), ws.Cells(lastRow, xCol))
        For Each cell In rng
            cell.Value = Trim(cell.Value)
        Next cell

        ' Collect the rows to be deleted in a range
        For lrow = lastRow To firstRow Step -1
            Set vRg = ws.Cells(lrow, xCol)
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
        ws.Columns(xCol).NumberFormat = "m/d/yyyy"

        ' Remove balance column
        xStr = "Balance"
        Set xRg = ws.Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
        xCol = xRg.Column
        ws.Columns(xCol).EntireColumn.Delete

        ' Create "Amount" column
        amountCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, amountCol).Value = "Amount"
        ws.Cells(2, amountCol).FormulaR1C1 = "=RC[-2]-RC[-1]"
        lastRow = ws.Cells.SpecialCells(xlLastCell).Row
        ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol)).FillDown
        ws.Columns(amountCol - 2).Resize(, 3).NumberFormat = "#,##0.00"

        ' Summarize the Amount column
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, lastColumn).Value = "Total:"
        ws.Cells(2, lastColumn).Formula = "=SUM(" & ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol)).Address & ")"

        ' Modify headers in the sheet
        xStr = "Type"
        Set xRg = ws.Range("A1:AZ1").Find(What:=xStr, LookIn:=xlValues, LookAt:=xlPart)

        If Not xRg Is Nothing Then
            xCol = xRg.Column
            ws.Cells(1, xCol).Value = "Source"
        Else
            ' Handle the case when a column header containing "Type" is not found
            MsgBox "No column header containing 'Type' was found.", vbExclamation
        End If

        xStr = "Date"
        Set xRg = ws.Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
        xCol = xRg.Column
        ws.Cells(1, xCol).Value = "Posted Date"

        xStr = "Num"
        Set xRg = ws.Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
        xCol = xRg.Column
        ws.Cells(1, xCol).Value = "Possible Journal ref. number"

        ' Insert "Copy of Date" column
        dateCol = 0
        For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If Trim(ws.Cells(1, col).Value) = "Posted Date" Then
                dateCol = col
                Exit For
            End If
        Next col

        If dateCol > 0 Then
            ws.Columns(dateCol + 1).Insert Shift:=xlToRight
            ws.Cells(1, dateCol + 1).Value = "Copy of Date"
            ws.Range(ws.Cells(2, dateCol + 1), ws.Cells(lastRow, dateCol + 1)).Value = ws.Range(ws.Cells(2, dateCol), ws.Cells(lastRow, dateCol)).Value
        ' Else
        '     MsgBox "Column 'Posted Date' not found.", vbExclamation
        End If

        'Insert "Comments" Column

        Set ws = ActiveSheet  ' Set ws to the active sheet
        headersToFind = Array("Memo", "Description", "Name", "Class")  ' Column headers to find
        Set foundHeaders = New Collection

        ' Get the header row range
        lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        Set headersRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))

        ' Initialize memoCol and splitCol
        memoCol = 0
        splitCol = 0

        For Each hdr In headersRange
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
            ' If neither "Split" nor "Memo" is found, handle accordingly
            MsgBox "Neither 'Split' nor 'Memo' column was found. Cannot insert 'Comments' column.", vbExclamation
            Exit Sub
        End If

        ' Check if the total of the Amount column is close enough to zero
        If Abs(ws.Cells(2, lastColumn + 2).Value) > 0.001 Then
            MsgBox "Amount column does not balance out. Current total is " & ws.Cells(2, lastColumn).Value & ". Please manually adjust for Completeness", vbExclamation
            Exit Sub
        End If

        ' Summarize the accounts for Completeness
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Set the ranges for Account and Amount columns
        accountCol = 0
        amountCol = 0
        For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If Trim(ws.Cells(1, col).Value) = "Account ref. number" Then
                accountCol = col
            ElseIf Trim(ws.Cells(1, col).Value) = "Amount" Then
                amountCol = col
            End If
        Next col

        If accountCol = 0 Or amountCol = 0 Then
            'MsgBox "Could not find 'Account ref. number' or 'Amount' column. Please check the column headers.", vbExclamation
            Exit Sub
        End If


        If accountCol > 0 And amountCol > 0 Then
            Set accountRange = ws.Range(ws.Cells(2, accountCol), ws.Cells(lastRow, accountCol))
            Set amountRange = ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol))
        Else
            MsgBox "Could not find 'Account' or 'Amount' column. Please check the column headers.", vbExclamation
            Exit Sub
        End If

        ' Create a collection of Account Numbers
        Set uniqueAccounts = New Collection
        On Error Resume Next
        For Each account In accountRange
            uniqueAccounts.Add account.Value, CStr(account.Value)
        Next account
        On Error GoTo 0

        ' Create Summary Worksheet
        Set summarySheet = ws.Parent.Sheets.Add(After:=ws.Parent.Sheets(ws.Parent.Sheets.Count))
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
        ws.Columns("E:E").Insert Shift:=xlToRight
        ws.Columns("C:C").Copy Destination:=ws.Columns("E:E")
        ws.Columns("D:D").Copy Destination:=ws.Columns("C:C")
        ws.Columns("E:E").Copy Destination:=ws.Columns("D:D")
        ws.Columns("E:E").Delete
        Application.CutCopyMode = False

        Application.ScreenUpdating = True

        'Paste the GL summary into the Completeness test template
        lastRow = summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).Row
        Set GLSummary = summarySheet.Range("A2:C" & lastRow)
        GLSummary.Copy

        ' Open the target workbook
        Set originalWorkbook = ThisWorkbook
        Set targetWorkbook = Workbooks.Open("C:\*YourFilePath*)")
        On Error Resume Next

        Set targetWs = targetWorkbook.Worksheets("GL")

        targetWs.Range("GL_Input").ClearContents
        targetWs.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        GLSummary.Copy
        targetWs.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ' Save Workbook after opening completeness

        ' Prompt for ticket number
        TicketNo = Trim(InputBox("What is the ticket number?"))

        ' Reference file with S drive locations
        refFile = "S:\*WholeFilePath*\" & TicketNo & "\" & TicketNo & ".txt"

        On Error GoTo ErrMsg1
        Open refFile For Input As #1
            Do Until EOF(1)
                Line Input #1, textline
                text = text & textline
            Loop
        Close #1

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

    ElseIf amountCol > 0 Then
        ' Logic from QBFormatAmtBal
        ' Remove QB Export Tips sheet (if exists)
        Application.DisplayAlerts = False
        For Each Sheet In ActiveWorkbook.Worksheets
            If Sheet.Name Like "*Tips" Then
                Sheet.Delete
            End If
        Next Sheet
        Application.DisplayAlerts = True

        ' Make whole sheet normal formatting
        With ws.Cells
            .Style = "Normal"
            .UnMerge
        End With

        ' Remove any rows above the headers (search for "Date")
        lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        Set rng = ws.Range("A1:AZ" & lastRow)

        Set dateCell = rng.Find(What:="Date", LookAt:=xlPart, MatchCase:=False) ' Non case-sensitive, and match part of cell content

        If Not dateCell Is Nothing Then ' If "Date" was found
            For i = dateCell.Row - 1 To 1 Step -1 ' Loop upwards from "Date" to the first row
                ws.Rows(i).Delete ' Delete the row
            Next i
        End If

        ' Remove all blank columns
        Set ws = ActiveSheet
        lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

        For columnNumber = lastColumn To 1 Step -1
            If WorksheetFunction.CountA(ws.Columns(columnNumber)) = 0 Then
                ws.Columns(columnNumber).Delete
            End If
        Next columnNumber

        ' Set xCol as first column with a header
        xCol = ws.Range("A1").End(xlToRight).Column
        ws.Range("A1").Value = "Account ref. number"

        If xCol > 2 Then
            Set colRange = ws.Columns(1).Resize(, xCol - 2)
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

        ' Trim the values in the range directly
        Set rng = ws.Range("A1:A" & lastRow)
        For lrow = lastRow To firstRow Step -1
            Set vRg = rng.Cells(lrow - firstRow + 1)
            If vRg.HasFormula Then  ' If the cell has a formula
                vRg.Value = vRg.Value
            End If
            If vRg.Value = 0 Then   ' If the cell is 0
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

        ' Delete columns B thru xCol
        If xCol > 2 Then
            ws.Columns(2).Resize(, xCol - 2).EntireColumn.Delete
        End If
        ws.Columns("A:A").EntireColumn.AutoFit

        ' Find and format the Date Column
        xCol = WorksheetFunction.Match("Date", ws.Range("1:1"), 0)
        ws.Columns(xCol).NumberFormat = "m/d/yyyy"

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

        ' Remove balance column
        xStr = "Balance"
        Set xRg = ws.Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
        If Not xRg Is Nothing Then
            xCol = xRg.Column
            ws.Columns(xCol).EntireColumn.Delete
        End If

        ' Modify headers in the sheet
        xStr = "Type"
        Set xRg = ws.Range("A1:AZ1").Find(What:=xStr, LookIn:=xlValues, LookAt:=xlPart)
        If Not xRg Is Nothing Then
            xCol = xRg.Column
            ws.Cells(1, xCol).Value = "Source"
        End If

        xStr = "Date"
        Set xRg = ws.Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
        If Not xRg Is Nothing Then
            xCol = xRg.Column
            ws.Cells(1, xCol).Value = "Posted date"
        End If

        xStr = "Num"
        Set xRg = ws.Range("A1:AZ1").Find(xStr, , xlValues, xlWhole, , , True)
        If Not xRg Is Nothing Then
            xCol = xRg.Column
            ws.Cells(1, xCol).Value = "Possible Journal ref. number"
        End If

        ' Change "Amount" to "Original Amount" and add "Amount" column
        amountCol = WorksheetFunction.Match("Amount", ws.Range("1:1"), 0)
        amountAdjustedCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, amountCol).Value = "Original Amount"
        ws.Cells(1, amountAdjustedCol).Value = "Amount"
        ws.Cells(1, amountAdjustedCol).Interior.Color = RGB(255, 255, 0)

        ' Insert "Copy of Date" column
        dateCol = WorksheetFunction.Match("Posted date", ws.Range("1:1"), 0)
        ws.Columns(dateCol + 1).Insert Shift:=xlToRight
        ws.Cells(1, dateCol + 1).Value = "Copy of Date"
        ws.Range(ws.Cells(2, dateCol + 1), ws.Cells(lastRow, dateCol + 1)).Value = ws.Range(ws.Cells(2, dateCol), ws.Cells(lastRow, dateCol)).Value

        ' Insert "Comments" Column
        headersToFind = Array("Memo", "Description", "Name", "Class")
        Set foundHeaders = New Collection

        'Get the header row range
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Set headersRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))

        ' Initialize memoCol and splitCol
        memoCol = 0
        splitCol = 0

       ' Find the columns
       For Each hdr In headersRange
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
            MsgBox "Neither 'Split' nor 'Memo' column was found. Cannot insert 'Comments' column.", vbExclamation
            Exit Sub
        End If

        ' Insert the "Comments" column
        ws.Columns(commentsCol).Insert Shift:=xlToRight
        ws.Cells(1, commentsCol).Value = "Comments"

        ' Concatenate the values from the found headers
        For i = 2 To lastRow
            comment = ""
            For j = 1 To foundHeaders.Count
                comment = comment & " " & ws.Cells(i, foundHeaders(j)).Value
            Next j
            ws.Cells(i, commentsCol).Value = Trim(comment)
        Next i

        ' Update column numbers after adding new column
        amountCol = 0
        amountAdjustedCol = 0
        For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If Trim(ws.Cells(1, col).Value) = "Original Amount" Then
                amountCol = col
            ElseIf Trim(ws.Cells(1, col).Value) = "Amount" Then
                amountAdjustedCol = col
            End If
        Next col

        If amountCol = 0 Or amountAdjustedCol = 0 Then
            'MsgBox "Could not find 'Original Amount' or 'Amount' column. Please check the column headers.", vbExclamation
            Exit Sub
        End If


        ' Copy the values from "Original Amount" to "Amount" column
        ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol)).Copy Destination:=ws.Cells(2, amountAdjustedCol)

        ' Find the first instance of Account ref. number starting with "2", "Accounts Payable" or contains "A/P"
        firstInstanceRow = 0
        For rowNum = 2 To lastRow
            If Left(ws.Cells(rowNum, 1).Value, 1) = "2" Or ws.Cells(rowNum, 1).Value = "Accounts Payable" Or InStr(ws.Cells(rowNum, 1).Value, "A/P") > 0 Then
                firstInstanceRow = rowNum
                Exit For
            End If
        Next rowNum

        ' If the first instance is found, inverse the value and highlight the cell yellow
        If firstInstanceRow > 0 Then
            For rowNum = firstInstanceRow To lastRow
                If Left(ws.Cells(rowNum, 1).Value, 1) >= "5" Or _
                InStr(UCase(ws.Cells(rowNum, 1).Value), "WAGES") > 0 Or _
                InStr(UCase(ws.Cells(rowNum, 1).Value), "COST OF GOODS") > 0 Or _
                InStr(UCase(ws.Cells(rowNum, 1).Value), "COGS") > 0 Then
                    Exit For
                End If
                ws.Cells(rowNum, amountAdjustedCol).Value = -ws.Cells(rowNum, amountCol).Value
                ws.Cells(rowNum, amountAdjustedCol).Interior.Color = RGB(255, 255, 0)
            Next rowNum
        End If

        ' Format the "Amount" columns
        ws.Columns(amountCol).NumberFormat = "#,##0.00"
        ws.Columns(amountAdjustedCol).NumberFormat = "#,##0.00"

        ' Summarize the Amount and Adjusted Amount Columns
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, lastColumn).Value = "Total:"
        ws.Cells(1, lastColumn + 1).Formula = "=SUM(" & ws.Cells(2, amountCol).Address & ":" & ws.Cells(lastRow, amountCol).Address & ")"
        ws.Cells(1, lastColumn + 2).Formula = "=SUM(" & ws.Cells(2, amountAdjustedCol).Address & ":" & ws.Cells(lastRow, amountAdjustedCol).Address & ")"

        ' Apply Visual-formatting to the new results
        With ws.Cells(1, lastColumn).Resize(, 3)
            .Style = "Total"
            .Font.Bold = True
        End With
        Cells(1, lastColumn + 2).Interior.Color = RGB(255, 255, 0)
        Range(Cells(1, 1), Cells(1, lastColumn - 1)).Font.Bold = True

        ' Check if the total of the Adjusted Amount column is close enough to zero
        If Abs(ws.Cells(1, lastColumn + 2).Value) > 0.00001 Then
            MsgBox "Adjusted Amount column does not balance out, please manually adjust for Completeness", vbExclamation
            Exit Sub
        End If

        ' Summarize the accounts for Completeness
            Set ws = ActiveSheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
        ' Set the ranges for Account and Amount columns
        accountCol = 0
        amountCol = 0
        For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If Trim(ws.Cells(1, col).Value) = "Account ref. number" Then
                accountCol = col
            ElseIf Trim(ws.Cells(1, col).Value) = "Amount" Then
                amountCol = col
            End If
        Next col

        If accountCol = 0 Or amountCol = 0 Then
            'MsgBox "Could not find 'Account ref. number' or 'Amount' column. Please check the column headers.", vbExclamation
            Exit Sub
        End If


        If accountCol > 0 And amountCol > 0 Then
            Set accountRange = ws.Range(ws.Cells(2, accountCol), ws.Cells(lastRow, accountCol))
            Set amountRange = ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol))
        Else
            MsgBox "Could not find 'Account' or 'Amount' column. Please check the column headers.", vbExclamation
            Exit Sub
        End If
            
            ' Create a collection of Account Numberss
            Set uniqueAccounts = New Collection
                On Error Resume Next
                For Each account In accountRange
                    uniqueAccounts.Add account.Value, CStr(account.Value)
                Next account
                On Error GoTo 0
            
        ' Create Summary Worksheet
        Set summarySheet = ActiveWorkbook.Sheets.Add(After:=ActiveSheet)
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
            ws.Columns("E:E").Insert Shift:=xlToRight
            ws.Columns("C:C").Copy Destination:=ws.Columns("E:E")
            ws.Columns("D:D").Copy Destination:=ws.Columns("C:C")
            ws.Columns("E:E").Copy Destination:=ws.Columns("D:D")
            ws.Columns("E:E").Delete
            Application.CutCopyMode = False

            Application.ScreenUpdating = True

        'Paste the GL summary into the Completeness test template
        lastRow = summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).Row
        Set GLSummary = summarySheet.Range("A2:C" & lastRow)
        GLSummary.Copy

        ' Open the target workbook
        Set originalWorkbook = ThisWorkbook
        Set targetWorkbook = Workbooks.Open("C:\*YourFilePath*)")
        On Error Resume Next

        Set targetWs = targetWorkbook.Worksheets("GL")

        targetWs.Range("GL_Input").ClearContents
        targetWs.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        GLSummary.Copy
        targetWs.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ' Save Workbook after opening completeness
        ' Prompt for ticket number
        TicketNo = Trim(InputBox("What is the ticket number?"))

        ' Reference file with S drive locations
        refFile = "S:\*WholeFilePath*\" & TicketNo & "\" & TicketNo & ".txt"

        On Error GoTo ErrMsg1
        Open refFile For Input As #1
            Do Until EOF(1)
                Line Input #1, textline
                text = text & textline
            Loop
        Close #1

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

    Else
        MsgBox "The worksheet format is not recognized. Please check the column headers.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = True
End Sub