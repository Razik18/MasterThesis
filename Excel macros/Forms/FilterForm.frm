VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FilterForm 
   Caption         =   "Filters Form"
   ClientHeight    =   9300.001
   ClientLeft      =   105
   ClientTop       =   455
   ClientWidth     =   16191
   OleObjectBlob   =   "FilterForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FilterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    SetVariables

    ' Populate each list box with data from respective tables
    PopulateListBox Me.ListBoxBlock, "BlockStateTable"
    PopulateListBox Me.ListBoxAnatomic, "AnatomicSiteTable"
    PopulateListBox Me.ListBoxTumor, "TumorTypeTable"
    PopulateListBox Me.ListBoxMarker, "MarkersTable"
    PopulateListBox Me.ListBoxProcess, "ProcessTable"
    PopulateListBox Me.ListBoxSite, "SitesTable"
    PopulateListBox Me.ListBoxFixative, "FixativeTable"

    ' Clear the scoring list box
    Me.ListBoxScore.Clear
End Sub

Private Sub PopulateListBox(lb As MSForms.ListBox, tableName As String)
    Dim ws As Worksheet
    Dim table As ListObject
    Dim i As Long

    ' Set the worksheet containing the tables
    Set ws = SettingWS

    ' Try to find the table by name
    On Error Resume Next
    Set table = ws.ListObjects(tableName)
    On Error GoTo 0

    ' If the table exists, populate the ListBox
    If Not table Is Nothing Then
        lb.Clear ' Clear previous items
        For i = 1 To table.ListRows.Count
            lb.AddItem table.ListRows(i).Range.Cells(1, 1).value
        Next i
    Else
        MsgBox "Table '" & tableName & "' not found.", vbExclamation
    End If
End Sub

' Event handler for list box selection changes
Private Sub ListBoxBlock_Change()
    ApplyDynamicFilters
End Sub

Private Sub ListBoxAnatomic_Change()
    ApplyDynamicFilters
End Sub

Private Sub ListBoxTumor_Change()
    ApplyDynamicFilters
End Sub

Private Sub ListBoxMarker_Change()
    PopulateScoringList Me.ListBoxMarker, Me.ListBoxScore
    ApplyDynamicFilters
End Sub

Private Sub ListBoxScore_Change()
    ApplyDynamicFilters
End Sub

Private Sub ListBoxProcess_Change()
    ApplyDynamicFilters
End Sub

Private Sub ListBoxSite_Change()
    ApplyDynamicFilters
End Sub

Private Sub ListBoxFixative_Change()
    ApplyDynamicFilters
End Sub

Private Sub PopulateScoringList(markerBox As MSForms.ListBox, scoreBox As MSForms.ListBox)
    Dim ws As Worksheet
    Dim scoringTable As ListObject
    Dim tableName As String
    Dim Marker As String
    Dim i As Long, j As Long
    Dim scores As Collection
    Dim score As Variant
    Dim formattedScore As String

    Set scores = New Collection
    Set ws = SettingWS

    For i = 0 To markerBox.ListCount - 1
        If markerBox.Selected(i) Then
            Marker = markerBox.List(i)

            tableName = Replace(Marker, " ", "")
            tableName = Replace(tableName, "-", "")
            tableName = Replace(tableName, "(", "")
            tableName = Replace(tableName, ")", "")
            tableName = Replace(tableName, "/", "")
            tableName = tableName & "Scoring"

            On Error Resume Next
            Set scoringTable = ws.ListObjects(tableName)
            On Error GoTo 0

            If Not scoringTable Is Nothing Then
                For j = 1 To scoringTable.ListRows.Count
                    On Error Resume Next
                    formattedScore = "[" & Marker & "]" & scoringTable.ListRows(j).Range.Cells(1, 1).value
                    scores.Add formattedScore, CStr(formattedScore)
                    On Error GoTo 0
                Next j
            End If
        End If
    Next i

    scoreBox.Clear
    For Each score In scores
        scoreBox.AddItem score
    Next score
End Sub

Private Sub ApplyDynamicFilters()
    Dim ws As Worksheet
    Dim BlocksTable As ListObject
    Dim wasProtected As Boolean
    Const SheetPassword As String = "qc"

    ' Set the worksheet and BlocksTable
    Set ws = BlocksWS
    On Error Resume Next
    Set BlocksTable = ws.ListObjects("BlocksTable")
    On Error GoTo 0

    If BlocksTable Is Nothing Then
        MsgBox "BlocksTable not found.", vbExclamation
        Exit Sub
    End If

    ' Check if the sheet is protected
    wasProtected = ws.ProtectContents

    ' Unprotect the sheet if it was protected
    If wasProtected Then ws.Unprotect password:=SheetPassword

    ' Clear existing filters
    On Error Resume Next
    BlocksTable.AutoFilter.ShowAllData
    On Error GoTo 0

    ' Apply filters based on ListBox selections
    If Me.ListBoxBlock.ListIndex <> -1 Then
        ApplyFilter BlocksTable, BlockStateColName, GetSelectedItems(Me.ListBoxBlock)
    End If

    If Me.ListBoxAnatomic.ListIndex <> -1 Then
        ApplyFilter BlocksTable, AnatomicSiteColName, GetSelectedItems(Me.ListBoxAnatomic)
    End If

    If Me.ListBoxTumor.ListIndex <> -1 Then
        ApplyFilter BlocksTable, TumorTypeColName, GetSelectedItems(Me.ListBoxTumor)
    End If

    If Me.ListBoxProcess.ListIndex <> -1 Then
        ApplyFilter BlocksTable, ProcessColName, GetSelectedItems(Me.ListBoxProcess)
    End If

    If Me.ListBoxSite.ListIndex <> -1 Then
        ApplyFilter BlocksTable, SiteColName, GetSelectedItems(Me.ListBoxSite)
    End If

    If Me.ListBoxFixative.ListIndex <> -1 Then
        ApplyFilter BlocksTable, FixativeColName, GetSelectedItems(Me.ListBoxFixative)
    End If

    If Me.ListBoxMarker.ListIndex <> -1 Then
        ApplyContainsFilterWithOrLogic BlocksTable, MarkerUsedColName, GetSelectedItems(Me.ListBoxMarker)
    End If

    If Me.ListBoxScore.ListIndex <> -1 Then
        ApplyContainsFilterWithOrLogic BlocksTable, ScoreColName, GetSelectedItems(Me.ListBoxScore)
    End If

    ' Re-protect the sheet if it was protected
    If wasProtected Then
        ws.Protect password:=SheetPassword, AllowSorting:=True, AllowFiltering:=True
    End If
End Sub


Private Sub CommandButton1_Click()
    Dim i As Long

    ' Clear filters
    On Error Resume Next
    BlocksWS.ListObjects("BlocksTable").AutoFilter.ShowAllData
    On Error GoTo 0

    ' Clear list box selections
    For i = 0 To Me.Controls.Count - 1
        If TypeName(Me.Controls(i)) = "ListBox" Then
            ' Temporarily change MultiSelect mode to clear selections
            Me.Controls(i).MultiSelect = fmMultiSelectSingle
            Me.Controls(i).value = Null
            Me.Controls(i).MultiSelect = fmMultiSelectMulti
        End If
    Next i
End Sub


' Get selected items from a ListBox
Private Function GetSelectedItems(lb As MSForms.ListBox) As Collection
    Dim selectedItems As Collection
    Dim i As Long

    Set selectedItems = New Collection
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            selectedItems.Add lb.List(i)
        End If
    Next i

    Set GetSelectedItems = selectedItems
End Function

Private Sub ApplyFilter(tbl As ListObject, colName As String, filterValues As Collection)
    Dim colIndex As Long
    Dim criteria() As Variant
    Dim i As Long

    ' Find the column index by name
    On Error Resume Next
    colIndex = tbl.ListColumns(colName).Index
    On Error GoTo 0

    If colIndex = 0 Then Exit Sub ' Column not found

    ' Exit if there are no selected filter values
    If filterValues.Count = 0 Then Exit Sub

    ' Convert collection to array for filtering
    ReDim criteria(filterValues.Count - 1)
    For i = 1 To filterValues.Count
        criteria(i - 1) = filterValues(i)
    Next i

    ' Apply the filter
    tbl.Range.AutoFilter Field:=colIndex, Criteria1:=criteria, Operator:=xlFilterValues
End Sub

Private Sub ApplyContainsFilterWithOrLogic(tbl As ListObject, colName As String, filterValues As Collection)
    Dim colIndex As Long
    Dim criteria() As String
    Dim i As Long
    Dim cleanValue As String

    ' Find the column index by name
    On Error Resume Next
    colIndex = tbl.ListColumns(colName).Index
    On Error GoTo 0

    If colIndex = 0 Then Exit Sub ' Column not found

    ' Exit if no filter values
    If filterValues.Count = 0 Then Exit Sub

    ' Prepare the criteria array for the filter
    ReDim criteria(filterValues.Count - 1)

    ' Apply special logic for "Result" column only
    For i = 1 To filterValues.Count
        If colName = ScoreColName Then
            ' Extract the part after the closing bracket ]
            cleanValue = Mid(filterValues(i), InStr(filterValues(i), "]") + 1)
            criteria(i - 1) = "*" & cleanValue & "*"
        Else
            ' Standard contains behavior for other columns
            criteria(i - 1) = "*" & filterValues(i) & "*"
        End If
    Next i

    ' Apply the filter with multiple criteria (contains) using OR logic
    tbl.Range.AutoFilter Field:=colIndex, Criteria1:=criteria, Operator:=xlFilterValues
End Sub





' Generate "contains" filter criteria string
Private Function GetContainsFilter(lb As MSForms.ListBox) As String
    Dim filterString As String
    Dim i As Long

    filterString = ""
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            If filterString <> "" Then filterString = filterString & ","
            filterString = filterString & "*" & lb.List(i) & "*"
        End If
    Next i

    GetContainsFilter = filterString
End Function


