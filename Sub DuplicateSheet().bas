Sub DuplicateSheet()
Dim x As Integer
x = InputBox("Enter number of times to copy the Active Sheet")
For numtimes = 1 To x
ActiveSheet.Copy After:=Worksheets(Worksheets.Count)
Next
End Sub