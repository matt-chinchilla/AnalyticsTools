Sub FormatWorksheet()

    ' Turn on gridlines
    ActiveWindow.DisplayGridlines = True

    ' Select the entire worksheet
    Cells.Select

    ' Unmerge all cells
    Cells.UnMerge

    ' Find and delete rows above the header row
    Dim headerRow As Range
    Set headerRow = Rows.Find(What:="Date", LookIn:=xlValues, LookAt:=xlWhole)
    If headerRow Is Nothing Then
        Set headerRow = Rows.Find(What:="Account", LookIn:=xlValues, LookAt:=xlWhole)
    End If
    If Not headerRow Is Nothing Then
        If headerRow.Row > 2 Then
            Rows("2:" & headerRow.Row - 2).Delete
        End If
    End If

    ' Set font, font size, and bold formatting for header row
    Rows(1).Font.Name = "Calibri"
    Rows(1).Font.Size = 11
    Rows(1).Font.Bold = True

    ' Set font and font size for all other cells
    Cells.Font.Name = "Calibri"
    Cells.Font.Size = 11

    ' Turn off wrap text for all cells
    Cells.WrapText = False

    ' Delete blank columns
    Dim blankColumns As Range
    Set blankColumns = FindBlankColumns(ActiveSheet.UsedRange)
    If Not blankColumns Is Nothing Then
        blankColumns.Delete
    End If

    ' Delete blank rows
    Dim blankRows As Range
    Set blankRows = FindBlankRows(ActiveSheet.UsedRange)
    If Not blankRows Is Nothing Then
        blankRows.Delete
    End If

    ' Auto-fit columns and rows
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

    ' Turn off screen updating
    Application.ScreenUpdating = True

End Sub

Function FindBlankColumns(r As Range) As Range
    Dim blankColumns As Range
    Dim c As Range

    For Each c In r.Columns
        If WorksheetFunction.CountA(c) = 0 Then
            If blankColumns Is Nothing Then
                Set blankColumns = c
            Else
                Set blankColumns = Union(blankColumns, c)
            End If
        End If
    Next c

    Set FindBlankColumns = blankColumns
End Function

Function FindBlankRows(r As Range) As Range
    Dim blankRows As Range
    Dim c As Range

    For Each c In r.Rows
        If WorksheetFunction.CountA(c) = 0 Then
            If blankRows Is Nothing Then
                Set blankRows = c
            Else
                Set blankRows = Union(blankRows, c)
            End If
        End If
    Next c

    Set FindBlankRows = blankRows
End Function
