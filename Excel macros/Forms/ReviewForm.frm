VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReviewForm 
   Caption         =   "Send Block In Review"
   ClientHeight    =   7200
   ClientLeft      =   -14
   ClientTop       =   -21
   ClientWidth     =   5677
   OleObjectBlob   =   "ReviewForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReviewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim markerTable As ListObject
    Dim markerColumn As ListColumn
    Dim markerCell As Range
    Dim Marker As Variant

    ' Set reference to the "Settings" sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(settingsSheet)
    Set markerTable = ws.ListObjects("MarkersTable")
    On Error GoTo 0

    If markerTable Is Nothing Then
        MsgBox "MarkerTable not found in the Settings sheet. ListBox1 cannot be populated.", vbExclamation
        Exit Sub
    End If

    ' Set reference to the first column of the MarkerTable
    On Error Resume Next
    Set markerColumn = markerTable.ListColumns(1)
    On Error GoTo 0

    If markerColumn Is Nothing Then
        MsgBox "The first column in MarkerTable is not available. ListBox1 cannot be populated.", vbExclamation
        Exit Sub
    End If

    ' Populate ListBox1 with marker values from MarkerTable
    Me.ListBox1.Clear
    For Each markerCell In markerColumn.DataBodyRange
        If Not IsEmpty(markerCell.value) Then
            Me.ListBox1.AddItem markerCell.value
        End If
    Next markerCell
End Sub

Private Sub CommandButton1_Click()
    Dim BlockName As String
    Dim ListMarkers As Collection
    Dim blockType As String

    ' Initialize a collection for markers
    Set ListMarkers = New Collection

    ' Get selected markers
    Dim i As Long
    For i = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) Then
            ListMarkers.Add Me.ListBox1.List(i)
        End If
    Next i

    ' Get the block name
    BlockName = Me.TextBox1.value

    ' Determine the block type based on the selected option button
    If Me.OptionButtonParent.value Then
        blockType = "Parent"
    ElseIf Me.OptionButtonChild.value Then
        blockType = "Child"
    Else
        MsgBox "Please select either Parent or Child.", vbExclamation
        Exit Sub
    End If

    ' Call the module function with the block type
    MoveReviewBlock BlockName, ListMarkers, blockType


    Unload Me
End Sub

