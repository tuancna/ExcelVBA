Sub Combine()
    Dim i As Integer
    Dim headerCount As Variant
    Dim rsRow As Worksheet
    On Error Resume Next
LInput:
    headerCount = Application.InputBox("Enter number of header row", "", "1")
    If TypeName(headerCount) = "Boolean" Then Exit Sub
    If Not IsNumeric(headerCount) Then
        MsgBox "Number only", , "Msgbox"
        GoTo LInput
    End If
    Set rsRow = ActiveWorkbook.Worksheets.Add(Sheets(1))
    rsRow.Name = "Combined"
    Worksheets(2).Range("A1").EntireRow.Copy Destination:=rsRow.Range("A1")
    For i = 2 To Worksheets.Count
        Worksheets(i).Range("A1").CurrentRegion.Offset(CInt(headerCount), 0).Copy _
               Destination:=rsRow.Cells(rsRow.UsedRange.Cells(rsRow.UsedRange.Count).Row + 1, 1)
    Next
End Sub
