VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteScoringForm 
   Caption         =   "Delete Scoring "
   ClientHeight    =   7236
   ClientLeft      =   105
   ClientTop       =   455
   ClientWidth     =   13986
   OleObjectBlob   =   "DeleteScoringForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteScoringForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Triggered when CommandButton1 is clicked
Private Sub CommandButton1_Click()
    Dim selectedMarker As String
    Dim selectedScoring As String
    Dim ws As Worksheet
    Dim scoringTable As ListObject
    Dim tableName As String
    Dim i As Long

    ' Check if a marker is selected
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Please select a marker first.", vbExclamation
        Exit Sub
    End If

    ' Get the selected marker
    selectedMarker = Me.ListBox1.value

    ' Check if a scoring is selected
    If Me.ListBoxScore.ListIndex = -1 Then
        MsgBox "Please select a scoring to delete.", vbExclamation
        Exit Sub
    End If

    ' Get the selected scoring
    selectedScoring = Me.ListBoxScore.value

    ' Debugging: Output selected marker and scoring
    Debug.Print "Selected Marker: " & selectedMarker
    Debug.Print "Selected Scoring: " & selectedScoring

    ' Set the worksheet containing the tables
    Set ws = SettingWS

    ' Construct the table name
    tableName = Replace(selectedMarker, " ", "") ' Remove spaces
    tableName = Replace(tableName, "-", "") ' Remove dashes
    tableName = Replace(tableName, "(", "") ' Remove opening parenthesis
    tableName = Replace(tableName, ")", "") ' Remove closing parenthesis
    tableName = tableName & "Scoring" ' Append "Scoring"

    ' Try to find the table corresponding to the constructed table name
    On Error Resume Next
    Set scoringTable = ws.ListObjects(tableName)
    On Error GoTo 0

    ' If the table exists, find and delete the selected scoring
    If Not scoringTable Is Nothing Then
        For i = scoringTable.ListRows.Count To 1 Step -1
            If scoringTable.ListRows(i).Range.Cells(1, 1).value = selectedScoring Then
                scoringTable.ListRows(i).Delete
                MsgBox "Scoring '" & selectedScoring & "' deleted successfully.", vbInformation
                Exit For
            End If
        Next i

        ' Refresh the ListBoxScore to reflect the changes
        PopulateScoringList selectedMarker
    Else
        MsgBox "No scoring table found for marker: " & selectedMarker, vbExclamation
    End If
End Sub


' Triggered when the user selects a marker in ListBox1
Private Sub ListBox1_Change()
    Dim selectedMarker As String

    ' Get the selected marker
    selectedMarker = Me.ListBox1.value

    ' Debugging: Output the selected marker
    Debug.Print "Selected Marker: " & selectedMarker

    ' Populate scores in ListBoxScore for the selected marker
    If selectedMarker <> "" Then
        PopulateScoringList selectedMarker
    Else
        MsgBox "No marker selected.", vbExclamation
    End If
End Sub

' Method to populate ListBoxScore with scores for the selected marker
Private Sub PopulateScoringList(Marker As String)
    Dim ws As Worksheet
    Dim scoringTable As ListObject
    Dim i As Long
    Dim tableName As String

    ' Set the worksheet containing the tables
    Set ws = SettingWS

    ' Clear any previous items in ListBoxScore
    Me.ListBoxScore.Clear

    ' Construct the table name by removing spaces, dashes, and parentheses, then adding "Scoring"
    tableName = Replace(Marker, " ", "") ' Remove spaces
    tableName = Replace(tableName, "-", "") ' Remove dashes
    tableName = Replace(tableName, "(", "") ' Remove opening parenthesis
    tableName = Replace(tableName, "/", "")
    tableName = Replace(tableName, ")", "") ' Remove closing parenthesis
    tableName = tableName & "Scoring" ' Append "Scoring"

    ' Try to find the table corresponding to the constructed table name
    On Error Resume Next
    Set scoringTable = ws.ListObjects(tableName)
    On Error GoTo 0

    ' If the table is found, add scores to ListBoxScore
    If Not scoringTable Is Nothing Then
        For i = 1 To scoringTable.ListRows.Count
            Me.ListBoxScore.AddItem scoringTable.ListRows(i).Range.Cells(1, 1).value
        Next i
    Else
        MsgBox "No scoring table found for marker: " & Marker, vbExclamation
    End If
End Sub

