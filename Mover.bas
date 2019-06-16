Sub Move()
Application.ScreenUpdating = False
Worksheets("CMSPull").Activate
Dim starttime As Long
Dim endtime As Long
Dim operator As Long
Dim ActualStart As Long
Dim ActualStop As Long
Dim title As Long
Dim externalID As Long
Dim customerID  As Long
Dim account As Long

title = WorksheetFunction.Match("Title", Sheets("CMSPull").Rows(1), 0)
account = WorksheetFunction.Match("Owner", Sheets("CMSPull").Rows(1), 0)
externalID = WorksheetFunction.Match("External ID", Sheets("CMSPull").Rows(1), 0)
customerID = WorksheetFunction.Match("CustomerID", Sheets("CMSPull").Rows(1), 0)
starttime = WorksheetFunction.Match("Scheduled Start", Sheets("CMSPull").Rows(1), 0)
ActualStart = WorksheetFunction.Match("Actual Start", Sheets("CMSPull").Rows(1), 0)
ActualStop = WorksheetFunction.Match("Actual Stop", Sheets("CMSPull").Rows(1), 0)
endtime = WorksheetFunction.Match("Scheduled Stop", Sheets("CMSPull").Rows(1), 0)
operator = WorksheetFunction.Match("Operator", Sheets("CMSPull").Rows(1), 0)

Sheets("CMSPull").Columns(account).Copy Destination:=Sheets("Output").Columns(1)
Sheets("CMSPull").Columns(title).Copy Destination:=Sheets("Output").Columns(2)
Sheets("CMSPull").Columns(externalID).Copy Destination:=Sheets("Output").Columns(3)
Sheets("CMSPull").Columns(starttime).Copy Destination:=Sheets("Output").Columns(4)
Sheets("CMSPull").Columns(endtime).Copy Destination:=Sheets("Output").Columns(5)
Sheets("CMSPull").Columns(ActualStart).Copy Destination:=Sheets("Output").Columns(6)
Sheets("CMSPull").Columns(ActualStop).Copy Destination:=Sheets("Output").Columns(7)
Sheets("CMSPull").Columns(operator).Copy Destination:=Sheets("Output").Columns(9)
Sheets("CMSPull").Columns(customerID).Copy Destination:=Sheets("Output").Columns(10)
Worksheets("Output").Columns("A").AutoFit
Worksheets("Output").Range("1:1").HorizontalAlignment = xlCenter
Worksheets("Output").Range("1:1").VerticalAlignment = xlCenter
Worksheets("Output").Range("B:B").ColumnWidth = 30
Worksheets("Output").Columns("C:J").AutoFit
Worksheets("Output").Range("1:1").Font.Bold = True
Worksheets("Output").Activate
End Sub
