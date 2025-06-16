VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MarkerDeleteForm 
   Caption         =   "Delete Marker"
   ClientHeight    =   6852
   ClientLeft      =   105
   ClientTop       =   455
   ClientWidth     =   6328
   OleObjectBlob   =   "MarkerDeleteForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MarkerDeleteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim selectedMarker As String
    Dim ws As Worksheet
    Dim markersTable As ListObject
    Dim scoringTable As ListObject
    Dim tableName As String
    Dim i As Long

    ' Check if a marker is selected
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Please select a marker to delete.", vbExclamation
        Exit Sub
    End If

    ' Get the selected marker
    selectedMarker = Me.ListBox1.value

    ' Debugging: Output the selected marker
    Debug.Print "Deleting Marker: " & selectedMarker

    ' Set the worksheet containing the tables
    Set ws = SettingWS

    ' Find the Markers table
    On Error Resume Next
    Set markersTable = ws.ListObjects(MarkersTableName)
    On Error GoTo 0

    If markersTable Is Nothing Then
        MsgBox "Markers table not found. Cannot delete marker.", vbExclamation
        Exit Sub
    End If

    ' Loop through the Markers table to find and delete the marker
    For i = markersTable.ListRows.Count To 1 Step -1
        If markersTable.ListRows(i).Range.Cells(1, 1).value = selectedMarker Then
            markersTable.ListRows(i).Delete
            MsgBox "Marker '" & selectedMarker & "' deleted from the Markers table.", vbInformation
            Exit For
        End If
    Next i

    ' Construct the corresponding scoring table name
    tableName = Replace(selectedMarker, " ", "") ' Remove spaces
    tableName = Replace(tableName, "-", "") ' Remove dashes
    tableName = Replace(tableName, "(", "") ' Remove opening parenthesis
    tableName = Replace(tableName, ")", "")
    tableName = Replace(tableName, "/", "")
    tableName = tableName & "Scoring" ' Append "Scoring"

    ' Find and delete the scoring table
    On Error Resume Next
    Set scoringTable = ws.ListObjects(tableName)
    On Error GoTo 0

    If Not scoringTable Is Nothing Then
        scoringTable.Delete
    Else
        MsgBox "No scoring table found for marker '" & selectedMarker & "'.", vbExclamation
    End If

    ' Refresh the ListBox to reflect the remaining markers
    PopulateMarkersList
End Sub

' Method to populate the ListBox with markers
Private Sub PopulateMarkersList()
    Dim ws As Worksheet
    Dim markersTable As ListObject
    Dim i As Long

    ' Set the worksheet containing the markers table
    Set ws = SettingWS

    ' Find the Markers table
    On Error Resume Next
    Set markersTable = ws.ListObjects(MarkersTableName)
    On Error GoTo 0

    ' Clear the ListBox
    Me.ListBox1.Clear

    ' Populate the ListBox if the table exists
    If Not markersTable Is Nothing Then
        For i = 1 To markersTable.ListRows.Count
            Me.ListBox1.AddItem markersTable.ListRows(i).Range.Cells(1, 1).value
        Next i
    End If
End Sub

' Initialize the UserForm
Private Sub UserForm_Initialize()
    SetVariables
    PopulateMarkersList ' Populate the ListBox with markers on form load
End Sub

