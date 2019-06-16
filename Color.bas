Sub Color()
Application.ScreenUpdating = False
Worksheets("Output").Activate
Dim sht As Worksheet
Set sht = ThisWorkbook.Worksheets("Output")
Dim lastrow As Long
lastrow = sht.Cells(Rows.Count, 9).End(xlUp).Row
Dim counter As Long
Dim opName As String
Dim startRange As Long
Dim endRange As Long
counter = 1
Dim sortRange As String

For i = 2 To lastrow
opName = sht.Cells(i, 9).Value
If opName <> sht.Cells(i - 1, 9).Value Then
startRange = i
counter = 1
Else
counter = counter + 1
End If

If opName <> sht.Cells(i + 1, 9).Value Then
endRange = i
sortRange = "A" & startRange & ":" & "L" & endRange
If counter = 3 Then
sht.Range(sortRange).Interior.Color = RGB(255, 255, 0)
ElseIf counter = 4 Then
sht.Range(sortRange).Interior.Color = RGB(112, 173, 71)
ElseIf counter = 5 Then
sht.Range(sortRange).Interior.Color = RGB(237, 125, 49)
ElseIf counter = 6 Then
sht.Range(sortRange).Interior.Color = RGB(117, 113, 113)
ElseIf counter = 7 Then
sht.Range(sortRange).Interior.Color = RGB(94, 247, 255)
End If
End If

Next i

Application.ScreenUpdating = True
End Sub
