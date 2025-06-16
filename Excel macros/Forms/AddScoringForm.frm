VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddScoringForm 
   Caption         =   "Add New Scoring"
   ClientHeight    =   6708
   ClientLeft      =   105
   ClientTop       =   455
   ClientWidth     =   10612
   OleObjectBlob   =   "AddScoringForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddScoringForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()
    Dim newScoring As String
    Dim selectedMarker As String
    Dim ws As Worksheet
    Dim scoringTable As ListObject
    Dim tableName As String

    ' Check if a marker is selected
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Please select a marker first.", vbExclamation
        Exit Sub
    End If

    ' Get the selected marker
    selectedMarker = Me.ListBox1.value

    ' Get the new scoring from the TextBox
    newScoring = Me.TextBoxNewScoring.value

    ' Check if new scoring is empty
    If Trim(newScoring) = "" Then
        MsgBox "Please enter a valid scoring to add.", vbExclamation
        Exit Sub
    End If

    ' Debugging: Output the new scoring
    Debug.Print "Adding New Scoring: " & newScoring & " for Marker: " & selectedMarker

    ' Set the worksheet containing the tables
    Set ws = SettingWS

    ' Construct the table name
    tableName = Replace(selectedMarker, " ", "") ' Remove spaces
    tableName = Replace(tableName, "-", "") ' Remove dashes
    tableName = Replace(tableName, "(", "") ' Remove opening parenthesis
    tableName = Replace(tableName, ")", "") ' Remove closing parenthesis
    tableName = Replace(tableName, "/", "")
    tableName = tableName & "Scoring" ' Append "Scoring"

    ' Try to find the table corresponding to the constructed table name
    On Error Resume Next
    Set scoringTable = ws.ListObjects(tableName)
    On Error GoTo 0

    ' If the table exists, add the new scoring
    If Not scoringTable Is Nothing Then
        scoringTable.ListRows.Add
        scoringTable.ListRows(scoringTable.ListRows.Count).Range.Cells(1, 1).value = newScoring
        MsgBox "New scoring '" & newScoring & "' added successfully.", vbInformation

    Else
        MsgBox "No scoring table found for marker: " & selectedMarker, vbExclamation
    End If
End Sub

