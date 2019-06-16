Sub DeleteEmpty()
Application.ScreenUpdating = False
Worksheets("CMSPull").Activate
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim ActualStart As Long
ActualStart = WorksheetFunction.Match("Actual Start", Sheets("CMSPull").Rows(1), 0)
Dim eventStart As String

For i = 2 To lastrow
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
If i - 1 = lastrow Then Exit For
eventStart = Cells(i, ActualStart).Value
If Trim(eventStart & vbNullString) = vbNullString Then
    Rows(i).EntireRow.Delete
    i = i - 1
End If
Next i

End Sub
