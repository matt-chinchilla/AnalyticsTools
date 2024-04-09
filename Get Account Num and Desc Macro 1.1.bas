Sub Get_Account_Info()
'
' Get_Account_Info Macro
' Inserts a row and splits the values in the first column into Account Number and Account Description
'

'
    Dim LastRow As Long
    
    'Insert row for Account Number
    Columns("A:A").Insert Shift:=xlToRight
    Range("A1").FormulaR1C1 = "Account Number"
    LastRow = Cells(Rows.Count, "B").End(xlUp).Row 'get last row of data
    Range("A2:A" & LastRow).FormulaR1C1 = "=LEFT(RC[1],FIND("" "",RC[1],1)-1)"
    Range("A2:A" & LastRow).Value = Range("A2:A" & LastRow).Value 'replace formulas with values
    
    'Insert row for Account Description
    Columns("B:B").Insert Shift:=xlToRight
    Range("B1").FormulaR1C1 = "Account Description"
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row 'get last row of data
    Range("B2:B" & LastRow).FormulaR1C1 = "=TRIM(RIGHT(RC[1],LEN(RC[1])-SEARCH("" "",RC[1])))"
    Range("B2:B" & LastRow).Value = Range("B2:B" & LastRow).Value 'replace formulas with values
    
    'AutoFill for Account Number
    LastRow = Cells(Rows.Count, "B").End(xlUp).Row 'get last row of data
    If LastRow > 2 Then 'if there's more than 1 row of data
        Range("A2:A" & LastRow).AutoFill Destination:=Range("A2:A" & LastRow)
        Resume NextLine1 'continue with next line of code
    End If
    
NextLine1:
    'AutoFill for Account Description
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row 'get last row of data
    If LastRow > 2 Then 'if there's more than 1 row of data
        Range("B2:B" & LastRow).AutoFill Destination:=Range("B2:B" & LastRow)
    End If
End Sub
