Attribute VB_Name = "AddMarker"
Sub AddMarkerAndScoringTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim markerName As String
    Dim scoringTableName As String
    Dim lastCol As Long
    Dim scoringTable As ListObject
    Dim newMarkerCell As Range
    Dim freeCol As Long
    Dim rowNum As Long
    Dim tblRange As Range

    ' Set the worksheet and main table
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects("MarkersTable")
    
    ' Prompt the user for the new marker name
    markerName = Application.InputBox("Enter the new marker name:", "Add Marker", Type:=2)

    If markerName = "" Or markerName = "False" Then Exit Sub
    
    ' Generate the scoring table name
    scoringTableName = CleanMarkerName(markerName)
    
    ' Check if the scoring table already exists
    On Error Resume Next
    Set scoringTable = ws.ListObjects(scoringTableName)
    On Error GoTo 0
    If Not scoringTable Is Nothing Then
        MsgBox "The scoring table '" & scoringTableName & "' already exists.", vbExclamation
        Exit Sub
    End If
    
    ' Add the new marker to the MarkersTable
    With tbl
        Set newMarkerCell = .ListColumns("Markers").DataBodyRange.Cells(.ListRows.Count + 1, 1)
        newMarkerCell.value = markerName
    End With

    ' Find the first empty column starting from 50
    freeCol = 150
    Do While ws.Cells(1, freeCol).value <> ""
        freeCol = freeCol + 2
    Loop
    
    ws.Cells(1, freeCol).value = markerName
    

    rowNum = 2
    Set tblRange = ws.Range(ws.Cells(1, freeCol), ws.Cells(rowNum, freeCol))

    Set scoringTable = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    scoringTable.Name = scoringTableName
    scoringTable.TableStyle = "TableStyleMedium9"
    

    ws.Cells(1, freeCol).Interior.Color = RGB(0, 51, 102)
    ws.Cells(1, freeCol).Font.Color = RGB(255, 255, 255)
    
    tblRange.Interior.Color = RGB(173, 216, 230)
    

    With tbl.ListColumns("Scoring").DataBodyRange
        .Cells(.Rows.Count, 1).Formula = _
            "=IFERROR(TEXTJOIN(""|"", TRUE, " & scoringTableName & "), ""N/A"")"
    End With

    MsgBox "Marker '" & markerName & "' and its empty scoring table '" & scoringTableName & "' have been added.", vbInformation
End Sub


Function CleanMarkerName(markerName As String) As String
    Dim cleanName As String
    
    cleanName = markerName
    cleanName = Replace(cleanName, " ", "")
    cleanName = Replace(cleanName, "-", "")
    cleanName = Replace(cleanName, "(", "")
    cleanName = Replace(cleanName, ")", "")
    cleanName = Replace(cleanName, "/", "")
    
    CleanMarkerName = cleanName & "Scoring"
End Function




