Sub Sort()
Application.ScreenUpdating = False
Worksheets("CMSPull").Activate
Dim sht As Worksheet
Set sht = ThisWorkbook.Worksheets("CMSPull")
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim LastColumn As Long
LastColumn = sht.Cells(1, sht.Columns.Count).End(xlToLeft).Column

Dim starttime As Long
Dim operator As Long
Dim startrow As Long
Dim endrow As Long

Dim startval As String
Dim nextstartval As String
Dim sortRange As String
Dim initRange As String

Dim ActualStart As Long
Dim ActualStop As Long
Dim title As Long
Dim externalID As Long
Dim customerID  As Long
Dim account As Long
Dim idxOfSlash As Integer
Dim idxOfSlash2 As Integer
Dim idxOfNSlash As Integer
Dim idxOfNSlash2 As Integer
Dim newstartval As String
Dim newstartval2 As String


startrow = 2


title = WorksheetFunction.Match("Title", Sheets("CMSPull").Rows(1), 0)
account = WorksheetFunction.Match("Owner", Sheets("CMSPull").Rows(1), 0)
externalID = WorksheetFunction.Match("External ID", Sheets("CMSPull").Rows(1), 0)
customerID = WorksheetFunction.Match("CustomerID", Sheets("CMSPull").Rows(1), 0)
starttime = WorksheetFunction.Match("Scheduled Start", Sheets("CMSPull").Rows(1), 0)
ActualStart = WorksheetFunction.Match("Actual Start", Sheets("CMSPull").Rows(1), 0)
ActualStop = WorksheetFunction.Match("Actual Stop", Sheets("CMSPull").Rows(1), 0)
endtime = WorksheetFunction.Match("Scheduled Stop", Sheets("CMSPull").Rows(1), 0)
operator = WorksheetFunction.Match("Operator", Sheets("CMSPull").Rows(1), 0)

initRange = "A" & 2 & ":" & "AW" & lastrow

For i = 2 To lastrow
    startval = Cells(i, starttime).Value
    idxOfSlash = InStr(1, startval, "/")
    idxOfSlash2 = InStr(idxOfSlash + 1, startval, "/")
    newstartval = Left(startval, idxOfSlash2 - 1)
    
    nextstartval = Cells(i + 1, starttime).Value
    
    
    If nextstartval <> "" Then
    idxOfNSlash = InStr(1, nextstartval, "/")
    idxOfNSlash2 = InStr(idxOfSlash + 1, nextstartval, "/")
    newstartval2 = Left(nextstartval, idxOfNSlash2 - 1)
        If newstartval <> newstartval2 Then
            endrow = i
            sortRange = "A" & startrow & ":" & "AM" & endrow
            Range(sortRange).Sort key1:=Cells(2, operator), Order1:=xlAscending, Header:=xlNo
            startrow = i + 1
        End If
    Else
        endrow = i
        sortRange = "A" & startrow & ":" & "AM" & endrow
        Range(sortRange).Sort key1:=Cells(2, operator), Order1:=xlAscending, Header:=xlNo
    End If
Next i
End Sub
