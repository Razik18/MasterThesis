VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResultTMAForm 
   Caption         =   "Scoring"
   ClientHeight    =   7908
   ClientLeft      =   -133
   ClientTop       =   -469
   ClientWidth     =   12229
   OleObjectBlob   =   "ResultTMAForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ResultTMAForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SelectedScores As Collection
Private CurrentMarker As String
Private blockType As String
Private isParent As Boolean
Private MarkerScoresDict As Object




Private Sub UserForm_Initialize()
    SetVariables
    ' Initialize the collection to store selected scores and their values
    Set SelectedScores = New Collection
    Set MarkerScoresDict = CreateObject("Scripting.Dictionary")
End Sub

Private Sub CommandButton3_Click()
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
    newScoring = Me.TextBoxScoring.value

    ' Check if the new scoring is empty
    If Trim(newScoring) = "" Then
        MsgBox "Please enter a valid scoring to add.", vbExclamation
        Exit Sub
    End If

    ' Debugging: Output the new scoring
    Debug.Print "Adding New Scoring: " & newScoring & " for Marker: " & selectedMarker

    ' Set the worksheet containing the tables
    Set ws = SettingWS

    ' Construct the table name
    selectedMarker = Replace(selectedMarker, "(in Review)", "")
    tableName = Replace(selectedMarker, " ", "") ' Remove spaces
    tableName = Replace(tableName, "-", "") ' Remove dashes
    tableName = Replace(tableName, "(", "") ' Remove opening parenthesis
    tableName = Replace(tableName, ")", "") ' Remove closing parenthesis
    tableName = Replace(tableName, "/", "") ' Remove forward slashes
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

        ' Refresh the scoring ListBox by calling ListBox1_Change
        Call ListBox1_Change
    Else
        MsgBox "No scoring table found for marker: " & selectedMarker, vbExclamation
    End If
End Sub
Private Sub ListBox1_Change()
    ' Clear the Scoring ListBox
    Me.ListBoxScoring.Clear
    
    ' For each selected marker in ListBox1, populate ListBoxScoring with scoring options
    Dim Marker As String
    Dim i As Long
    
    For i = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) Then
            Marker = Me.ListBox1.List(i)
            PopulateScoringList Marker ' Populate scoring options for the marker
            
            ' If no previous scores exist for this marker, initialize it in the dictionary
            If Not MarkerScoresDict.Exists(Marker) Then
                MarkerScoresDict.Add Marker, New Collection
            End If
        End If
    Next i
End Sub
Private Sub CommandButton1_Click()
    Dim selectedItem As String
    Dim value As String
    Dim ResultString As String
    Dim blockWs As Worksheet
    Dim blockTable As ListObject
    Dim foundRow As ListRow
    Dim resultCell As Range
    Dim i As Long
    Dim Marker As String
    Dim MarkerList() As String
    Dim j As Long
    Dim markerUsed As String
    Dim isMarkerUpdated As Boolean
    Dim isTMA As Boolean

    ' Determine if the block is a Parent or Child
    If Me.OptionButtonParent.value Then
        isParent = True
    ElseIf Me.OptionButtonChild.value Then
        isParent = False
    Else
        MsgBox "Please select either Parent or Child block type.", vbExclamation
        Exit Sub
    End If

    ' Check if a marker is selected
    If ListBox1.ListIndex = -1 Then
        MsgBox "Please select a marker first.", vbExclamation
        Exit Sub
    End If

    ' Check if the TMA checkbox is checked
    isTMA = Me.CheckBoxTMA.value

    ' Get the selected marker and set it
    Marker = ListBox1.List(ListBox1.ListIndex)
    Call SetMarker(Marker)

    ' Ensure SelectedScores is tied to the current marker
    If Not MarkerScoresDict.Exists(CurrentMarker) Then
        MarkerScoresDict.Add CurrentMarker, New Collection
    End If

    ' Clear any existing scores for the current marker
    Set MarkerScoresDict(CurrentMarker) = New Collection

    ' Loop through each selected item in the ListBoxScoring
    For i = 0 To Me.ListBoxScoring.ListCount - 1
        If Me.ListBoxScoring.Selected(i) Then
            selectedItem = Me.ListBoxScoring.List(i)

            ' Prompt the user for the value of the selected score
            value = InputBox("Enter the value for " & selectedItem & " for marker " & CurrentMarker, "Input Value")

            ' Add the score to the dictionary
            If value = "" Then
                ' Handle empty value: add only the selected item
                MarkerScoresDict(CurrentMarker).Add selectedItem
            Else
                ' Add the selected item with "{TMA}" before the value, if the checkbox is checked
                If isTMA Then
                    MarkerScoresDict(CurrentMarker).Add selectedItem & "{TMA}:" & value
                Else
                    MarkerScoresDict(CurrentMarker).Add selectedItem & ":" & value
                End If
            End If
        End If
    Next i

    ' Concatenate all the scoring values into one result string
    ResultString = ""
    For Each item In MarkerScoresDict(CurrentMarker)
        If ResultString <> "" Then
            ResultString = ResultString & " | " & item
        Else
            ResultString = item
        End If
    Next item

    ' Update the BlockTable with the result in the Parent's or Child's "Result" column
    Set blockWs = ThisWorkbook.Sheets(blocksSheet)
    Set blockTable = blockWs.ListObjects("BlocksTable")

    ' Find the correct row in the BlockTable for the block
    Set foundRow = Nothing
    If isParent Then
        ' Search for Vendor Block ID
        For Each row In blockTable.ListRows
            If row.Range.Cells(1, blockTable.ListColumns(ParentBlockColName).Index).value = Me.TextBox1.value Then
                Set foundRow = row
                Exit For
            End If
        Next row
    Else
        ' Search for Labcorp Block ID
        For Each row In blockTable.ListRows
            If row.Range.Cells(1, blockTable.ListColumns(ChildBlockColName).Index).value = Me.TextBox1.value Then
                Set foundRow = row
                Exit For
            End If
        Next row
    End If

    ' If the block row is found, update the Result column
    If Not foundRow Is Nothing Then
        Set resultCell = foundRow.Range.Cells(1, blockTable.ListColumns(ScoreColName).Index)
        If IsEmpty(resultCell.value) Then
            resultCell.value = "[" & CurrentMarker & "]" & ResultString
        Else
            resultCell.value = resultCell.value & " | [" & CurrentMarker & "]" & ResultString
        End If
    Else
        MsgBox "No matching block found for " & Me.TextBox1.value, vbExclamation
        Exit Sub
    End If

    ' Now, remove "(in Review)" from the selected marker in the BlocksData table
    Dim markerColumn As Range
    Dim markerCell As Range
    Set markerColumn = blockTable.ListColumns(MarkerUsedColName).DataBodyRange
    isMarkerUpdated = False

    ' Loop through the rows in the Marker Used column
    For Each markerCell In markerColumn
        MarkerList = Split(markerCell.value, "|")

        ' Check if the current marker cell contains the selected marker
        For j = LBound(MarkerList) To UBound(MarkerList)
            If Trim(MarkerList(j)) = Marker & "(in Review)" Then
                ' Remove "(in Review)" from the selected marker
                MarkerList(j) = Replace(MarkerList(j), "(in Review)", "")
                isMarkerUpdated = True
            End If
        Next j

        ' If we updated the marker, join the markers back and update the cell
        If isMarkerUpdated Then
            markerCell.value = Join(MarkerList, " | ")
            isMarkerUpdated = False
        End If
    Next markerCell

    ' Success message with the marker and results
    MsgBox "Successfully saved the results for block: " & Me.TextBox1.value & vbCrLf & _
           "Marker: " & CurrentMarker & vbCrLf & _
           "Scores: " & ResultString, vbInformation
End Sub




Sub UpdateBlockState(BlockName As String, isParent As Boolean)
    Dim blockWs As Worksheet
    Dim blockTable As ListObject
    Dim markerColumn As Range
    Dim markerCell As Range
    Dim MarkerList() As String
    Dim isInReviewRemaining As Boolean
    Dim blockStateCell As Range
    Dim row As ListRow
    Dim j As Long

    Set blockWs = ThisWorkbook.Sheets(blocksSheet)
    Set blockTable = blockWs.ListObjects("BlocksTable")

    isInReviewRemaining = False

    ' Get the "Marker Used" column
    Set markerColumn = blockTable.ListColumns(MarkerUsedColName).DataBodyRange

    ' Loop through each marker cell in the column
    For Each markerCell In markerColumn
        ' Split the markers in the cell
        MarkerList = Split(markerCell.value, "|")
        
        ' Check for any "(in Review)" markers
        For j = LBound(MarkerList) To UBound(MarkerList)
            If InStr(MarkerList(j), "(in Review)") > 0 Then
                isInReviewRemaining = True
                Exit For
            End If
        Next j

        ' Exit early if "(in Review)" is found
        If isInReviewRemaining Then Exit For
    Next markerCell

    ' If no "(in Review)" remains, update the Block State
    If Not isInReviewRemaining Then
        For Each row In blockTable.ListRows
            If row.Range.Cells(1, blockTable.ListColumns(ParentBlockColName).Index).value = BlockName Or _
               row.Range.Cells(1, blockTable.ListColumns(ChildBlockColName).Index).value = BlockName Then
               
                ' Get the Block State column and update it
                Set blockStateCell = row.Range.Cells(1, blockTable.ListColumns(BlockStateColName).Index)
                If isParent Then
                    blockStateCell.value = "3-CharacterizedParent"
                Else
                    blockStateCell.value = "6-ValidatedChild"
                End If
                Exit For
            End If
        Next row
    End If
End Sub

Private Sub PopulateScoringList(Marker As String)
    Dim ws As Worksheet
    Dim scoringTable As ListObject
    Dim i As Long
    Dim tableName As String

    ' Set the worksheet containing the tables
    Set ws = SettingWS

    ' Construct the table name
    Marker = Replace(Marker, "(in Review)", "")
    tableName = Replace(Marker, " ", "") ' Remove spaces
    tableName = Replace(tableName, "-", "") ' Remove dashes
    tableName = Replace(tableName, "(", "") ' Remove opening parenthesis
    tableName = Replace(tableName, ")", "") ' Remove closing parenthesis
    tableName = Replace(tableName, "/", "")
    tableName = tableName & "Scoring" ' Append "Scoring"

    ' Try to find the table corresponding to the constructed table name
    On Error Resume Next
    Set scoringTable = ws.ListObjects(tableName)
    On Error GoTo 0

    ' If the table is found, add scoring options to the ListBox
    If Not scoringTable Is Nothing Then
        For i = 1 To scoringTable.ListRows.Count
            Me.ListBoxScoring.AddItem scoringTable.ListRows(i).Range.Cells(1, 1).value
        Next i
    Else
        MsgBox "No scoring table found for marker: " & Marker, vbExclamation
        ' Close the form if no scoring options are available
    End If
End Sub

Public Sub SetMarker(Marker As String)
    Marker = Replace(Marker, "(in Review)", "")
    CurrentMarker = Marker
End Sub

Private Sub CommandButton2_Click()
    Dim BlockName As String
    Dim blockRow As Long
    Dim MarkerText As String
    Dim MarkerList() As String
    Dim markerUsedCol As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim BlocksTable As ListObject
    Dim blockType As String

    SetVariables

    ' Retrieve Block Name from the TextBox
    BlockName = Me.TextBox1.value

    ' Check if Block Name is empty
    If BlockName = "" Then
        MsgBox "Block Name cannot be empty.", vbExclamation
        Exit Sub
    End If

    ' Determine the block type based on selected option button
    If Me.OptionButtonParent.value Then
        blockType = "Parent"
    ElseIf Me.OptionButtonChild.value Then
        blockType = "Child"
    Else
        MsgBox "Please select either Parent or Child.", vbExclamation
        Exit Sub
    End If

    ' Set the worksheet and table containing blocks
    Set ws = BlocksWS
    Set BlocksTable = ws.ListObjects(BlocksTableName)

    ' Determine the correct column name based on block type
    Dim BlockColName As String
    If blockType = "Parent" Then
        BlockColName = ParentBlockColName
    ElseIf blockType = "Child" Then
        BlockColName = ChildBlockColName
    End If

    ' Find the row corresponding to the Block Name
    blockRow = Get_ParentBlock_Rows(ws, BlocksTableName, BlockColName, BlockName)
    If blockRow = -1 Then
        MsgBox "This Block Name is not found: " & BlockName, vbExclamation
        Exit Sub
    End If

    ' Get the column for markers in review
    markerUsedCol = BlocksTable.HeaderRowRange.Cells.Find(MarkerUsedColName).column

    ' Get the marker text for this Block
    MarkerText = ws.Cells(blockRow, markerUsedCol).value

    ' Check if there are markers in review
    If MarkerText = "" Then
        MsgBox "This Block has no markers in review.", vbExclamation
        Exit Sub
    End If

    ' Split the marker text into a list of markers
    MarkerList = Split(MarkerText, "|")

    ' Populate the ListBox with markers
    Me.ListBox1.Clear
    For i = LBound(MarkerList) To UBound(MarkerList)
        Me.ListBox1.AddItem Trim(MarkerList(i))
    Next i
End Sub

Private Sub TextBox1_Change()
    ' Clear the ListBoxMarkers
    Me.ListBox1.Clear
    Me.ListBoxScoring.Clear
End Sub

