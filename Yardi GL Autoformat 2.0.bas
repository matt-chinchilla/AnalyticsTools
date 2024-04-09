' Yardi Macro Version 2.0.0 by Matthew Chirichella

' Designed to be optimized for the Yardi GL export into Excel and rapidly launch into Completeness Testing Template
Option Explicit

Sub YardiGLAutoFormat2()

Application.ScreenUpdating = False

'Duplicate the active worksheet and rename it to "Reformatted"

    ActiveSheet.Select
    ActiveSheet.Copy After:=ActiveSheet
    ActiveSheet.Name = "Reformatted"
    
'Remove any rows above the headers

    Dim rng As Range
    Dim findCell As Range
    
    Set rng = ActiveSheet.Columns("A:AZ")
    Set findCell = rng.Find(What:="= Beginning Balance =").Offset(-1, 0)
    findCell.Select
    If findCell.Row > 1 Then
        Range("A1:A" & ActiveCell.Row - 1).EntireRow.Delete
    End If
    
'Fill down the Account number and description
    Columns("A:B").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    Set findCell = rng.Find(What:="Remarks")
    Cells.Select
    Selection.AutoFilter
    Selection.AutoFilter Field:=findCell.Column, Criteria1:= _
            "== Beginning Balance ="
    
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=RC[2]"
    Range("B2").Select
    Dim indx As Long
    indx = Range("B2").End(xlToRight).End(xlToRight).Column - 2
    ActiveCell.FormulaR1C1 = "=RC[" & indx & "]"
    
    indx = Range("C2").End(xlDown).Row
    Range("A2:B2").Select
    Selection.Copy
    Range("A2:B" & indx).Select
    ActiveSheet.Paste
    ActiveSheet.ShowAllData
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Acct No."
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Account Description"
    Columns("A:B").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("A:B").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    

'Delete Beginning and Ending Balance rows
    indx = Columns(findCell.Column).Find(What:="= Ending Balance =", searchdirection:=xlPrevious).Row
    Rows(indx & ":" & indx + 10).Delete
    
    Cells.Select
    Selection.AutoFilter Field:=findCell.Column, Criteria1:= _
        "== Beginning Balance =", Operator:=xlOr, Criteria2:="== Ending Balance ="
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    
'Delete rows that are blank except under the first 2 columns

    For indx = Selection.Rows.Count To 1 Step -1
        If WorksheetFunction.CountA(Rows(indx)) = 2 Then
            Rows(indx).EntireRow.Delete
        End If
    Next indx
    ActiveSheet.UsedRange
        
'Format column A and B
    Columns("C:C").Select
    Selection.Copy
    Columns("A:B").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone
    
'Add Account Ref. Number column
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Range("C1") = "Account Ref. Number"
    Columns("C:C").Select
    Range("C2").Activate
    ActiveCell.FormulaR1C1 = "=CONCAT(RC[-2],"" "",RC[-1])"
    indx = Range("A1").End(xlDown).Row
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & indx)
    
'Convert the range into a table and format the header row
    Range("A1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        ActiveSheet.ListObjects.Add(xlSrcRange, , , xlYes).Name = _
            "Reformatted"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
    End With
'Add Comments Column
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Range("H1") = "Comments"
    Range("H2").Activate
    ActiveCell.FormulaR1C1 = _
        "=TEXTJOIN("" "",TRUE,[@[Person/Description]],[@Reference],[@Remarks])"
    
'Add Amount column
    Set rng = ActiveSheet.Rows("1:1")
    Set findCell = rng.Find(What:="Credit")
    Columns(findCell.Column + 1).Select
    Selection.Insert Shift:=xlToRight
    ActiveCell.FormulaR1C1 = "Amount"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=[@Debit]-[@Credit]"
    
'Rename the Control and Date column header
    Range("Reformatted[[#Headers],[Date]]").Select
    ActiveCell.FormulaR1C1 = "Posted Date"
    Range("Reformatted[[#Headers],[Control]]").Select
    ActiveCell.FormulaR1C1 = "Journal Ref. Number"
    
'Convert the table back into a range
    ActiveSheet.ListObjects("Reformatted").Unlist
    ActiveWindow.ScrollRow = 1
    
'Summarize the GL in a different worksheet
    Columns("A:B").Select
    Selection.Copy
    Range("A1").Select
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "GL Summary"
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Range("$A:$B").RemoveDuplicates Columns:=Array(1, 2), _
        Header:=xlYes
    Range("A1").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$B$18"), , xlYes).Name = _
        "TableWow"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(Reformatted!C[-2],RC[-2],Reformatted!C[11])"
    
'Paste the GL summary into the Completeness test template
    Dim GLSummary As Range
    Range("tablewow").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Set GLSummary = Selection
    Workbooks.Open "C:\Users\*InsertYourFilepath*
    \Completeness Testing Macro Template 7.11.xlsm"
    On Error Resume Next
    ActiveWorkbook.Sheets("GL").Activate
    Range("GL_Input").ClearContents
    GLSummary.Copy
    ActiveSheet.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Application.ScreenUpdating = True


End Sub
