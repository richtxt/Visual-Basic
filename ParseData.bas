Sub Parse()
'TO DO: Put Regular Expression instead of Trim
Application.ScreenUpdating = False
Worksheets("CMSPull").Activate
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim actualStartIdx As Long
Dim actualStopIdx As Long
Dim operatorIdx As Long

Dim opEmail As String
Dim opEmailPrev As String
Dim eventStart As Date
Dim eventStop As Date
Dim currentOperator As cOperator
Dim opArray As New Collection
Dim eventsToRemove As New Collection

actualStartIdx = WorksheetFunction.Match("Actual Start", Sheets("CMSPull").Rows(1), 0)
actualStopIdx = WorksheetFunction.Match("Actual Stop", Sheets("CMSPull").Rows(1), 0)
operatorIdx = WorksheetFunction.Match("Operator", Sheets("CMSPull").Rows(1), 0)
    
For j = 2 To lastrow
    eventStart = CDate(Left(Cells(j, actualStartIdx).Value, Len(Cells(j, actualStartIdx).Value) - 4))
    eventStop = CDate(Left(Cells(j, actualStopIdx).Value, Len(Cells(j, actualStopIdx).Value) - 4))
    opEmailPrev = Cells(j - 1, operatorIdx).Value
    opEmail = Cells(j, operatorIdx).Value
    
    If opEmail <> opEmailPrev Then
        Set currentOperator = New cOperator
        Dim eventsList As Collection
        Set eventsList = New Collection
        Set currentOperator.OperatorEvents = eventsList
        currentOperator.OperatorEmail = opEmail
        opArray.Add currentOperator
    End If
    
    Dim operatorEvent As cOperatorEvent
    Set operatorEvent = New cOperatorEvent
    operatorEvent.ActualStart = eventStart
    operatorEvent.ActualStop = eventStop
    operatorEvent.RowID = j
    Set operatorEvent.RowRef = Worksheets("CMSPull").Rows(j).EntireRow
    currentOperator.OperatorEvents.Add operatorEvent
Next j
    
For Each operator In opArray
    For Each opEvent In operator.OperatorEvents
        For Each opEventToCompare In operator.OperatorEvents
            If opEvent.RowID <> opEventToCompare.RowID Then
                Dim detectedCollision As Boolean
                If opEvent.ActualStart > opEventToCompare.ActualStart And opEvent.ActualStart < opEventToCompare.ActualStop Then
                    detectedCollision = True
                End If
                If opEvent.ActualStop > opEventToCompare.ActualStart And opEvent.ActualStop < opEventToCompare.ActualStop Then
                    detectedCollision = True
                End If
                If detectedCollision = True Then
                    opEvent.Collision = True
                    opEventToCompare.Collision = True
                End If
                detectedCollision = False
            End If
        Next opEventToCompare
    Next opEvent
Next operator

For Each operator In opArray
    For Each opEventToRemove In operator.OperatorEvents
        If opEventToRemove.Collision = False Then
            eventsToRemove.Add opEventToRemove
        End If
    Next opEventToRemove
Next operator

For Each rowToRemove In eventsToRemove
    rowToRemove.RowRef.Delete
Next rowToRemove

Application.ScreenUpdating = True
End Sub
