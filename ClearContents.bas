Sub ClearContents()
Dim answer As Integer
Dim sht As Worksheet
Dim sht2 As Worksheet
Set sht = ThisWorkbook.Worksheets("Output")
Set sht2 = ThisWorkbook.Worksheets("CMSPull")
Dim lastrow As Long
Dim lastrow2 As Long
lastrow = sht.Cells(Rows.Count, 9).End(xlUp).Row
lastrow2 = sht2.Cells(Rows.Count, 9).End(xlUp).Row
Dim range1 As String
Dim range2 As String
range1 = "A" & 2 & ":" & "L" & lastrow + 1
range2 = "A" & 1 & ":" & "AZ" & lastrow2
answer = MsgBox("Are you sure you want to clean the sheets?", vbYesNo + vbQuestion)
If answer = vbYes Then
    sht.Range(range1).Clear
    sht2.Range(range2).Clear
Else
    'do nothing
End If

End Sub
