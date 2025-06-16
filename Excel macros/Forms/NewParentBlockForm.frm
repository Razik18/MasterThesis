VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewParentBlockForm 
   Caption         =   "Create New Parent Block"
   ClientHeight    =   9576.001
   ClientLeft      =   -231
   ClientTop       =   -861
   ClientWidth     =   17115
   OleObjectBlob   =   "NewParentBlockForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewParentBlockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    SetVariables
    ' Initialize list boxes with data from corresponding tables
    PopulateListBox Me.ListBoxAnatomic, "AnatomicSiteTable"
    PopulateListBox Me.ListBoxTumor, "TumorTypeTable"
    PopulateListBox Me.ListBoxVendor, "VendorsTable"
    PopulateListBox Me.ListBoxMarker, "MarkersTable"
    PopulateListBox Me.ListBoxProcess, "ProcessTable"
    PopulateListBox Me.ListBoxSite, "SitesTable"
    PopulateListBox Me.ListBoxFixative, "FixativeTable"
    PopulateComboBoxSampleType Me.ComboBoxSampleType
    Me.ComboBoxSampleType.value = "Tissue block"
    
    ' Clear the scoring list box as it depends on marker selection
    Me.ListBoxScore.Clear
End Sub

' Populate a ListBox from a given table
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

' Event handler for Marker ListBox selection
Private Sub ListBoxMarker_Change()
    ' Populate the scoring list box based on selected markers
    PopulateScoringList Me.ListBoxMarker, Me.ListBoxScore
End Sub
' Populate scoring ListBox based on selected markers
Private Sub PopulateScoringList(markerBox As MSForms.ListBox, scoreBox As MSForms.ListBox)
    Dim ws As Worksheet
    Dim scoringTable As ListObject
    Dim tableName As String
    Dim Marker As String
    Dim i As Long, j As Long
    Dim scores As Collection
    Set scores = New Collection

    ' Set the worksheet containing the scoring tables
    Set ws = SettingWS

    ' Iterate through selected markers
    For i = 0 To markerBox.ListCount - 1
        If markerBox.Selected(i) Then
            Marker = markerBox.List(i)

            ' Construct the scoring table name for the marker
            tableName = Replace(Marker, " ", "")
            tableName = Replace(tableName, "-", "")
            tableName = Replace(tableName, "(", "")
            tableName = Replace(tableName, ")", "")
            tableName = Replace(tableName, "/", "")
            tableName = tableName & "Scoring"

            ' Try to find the scoring table by name
            On Error Resume Next
            Set scoringTable = ws.ListObjects(tableName)
            On Error GoTo 0

            ' If the scoring table exists, add its scores to the collection
            If Not scoringTable Is Nothing Then
                For j = 1 To scoringTable.ListRows.Count
                    On Error Resume Next
                    ' Store the marker and score in the format [marker]score
                    scores.Add Array(Marker, scoringTable.ListRows(j).Range.Cells(1, 1).value), _
                               CStr(Marker & scoringTable.ListRows(j).Range.Cells(1, 1).value)
                    On Error GoTo 0
                Next j
            End If
        End If
    Next i

    ' Populate the scoring ListBox with unique scores
    scoreBox.Clear
    For Each score In scores
        scoreBox.AddItem score(0) & ":" & score(1)
    Next
End Sub



Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim BlocksTable As ListObject
    Dim NewRow As ListRow
    Dim ParentBlockName As String
    Dim additionalInfo As String
    Dim anatomicSite As String
    Dim tumorType As String
    Dim vendor As String
    Dim sampletype As String
    Dim markers As Collection
    Dim scores As Collection
    Dim biomarkerChar As String
    Dim MarkerScoresDict As Object
    Dim Marker As Variant
    Dim score As Variant
    Dim i As Long, j As Long
    Dim FolderPath As String
    Dim process As String
    Dim ParentBlockHyperlink As String
    Dim ParentBlockCell As Range
    Dim ChildBlockHyperlink As String
    Dim ChildBlockCell As Range
    Dim creationdate As String
    Dim letterIndex As Integer
    Dim generatedName As String
    Dim labcorpNameCell As Range
    Dim labcorpIDColIndex As Integer
    Dim acronym As String
    Dim newID As String
    Dim settingsWs As Worksheet
    Dim anatomicTable As ListObject
    Dim anatomicRow As ListRow
    Dim anatomicColIndex As Long
    Dim acronymColIndex As Long
    Dim counterColIndex As Long
    Dim counterRow As ListRow

    ' Initialize variables
    letterIndex = 1 ' Start with the first letter (A)

    ' Set worksheet and BlocksTable
    Set ws = ThisWorkbook.Sheets(blocksSheet)
    On Error Resume Next
    Set BlocksTable = ws.ListObjects("BlocksTable")
    On Error GoTo 0

    If BlocksTable Is Nothing Then
        MsgBox "BlocksTable not found.", vbExclamation
        Exit Sub
    End If

    ' Retrieve inputs from the form
    ParentBlockName = Me.TextBox1.Text
    additionalInfo = Me.TextBox2.Text
    fixtime = Me.TextBox3.Text
    creationdate = Me.TextBox4.Text
    anatomicSite = GetSingleSelection(Me.ListBoxAnatomic)
    tumorType = GetSingleSelection(Me.ListBoxTumor)
    vendor = GetSingleSelection(Me.ListBoxVendor)
    process = GetSingleSelection(Me.ListBoxProcess)
    site = GetSingleSelection(Me.ListBoxSite)
    fixative = GetSingleSelection(Me.ListBoxFixative)
    Set markers = GetSelectedItems(Me.ListBoxMarker)
    Set scores = GetSelectedItems(Me.ListBoxScore)
    sampletype = Me.ComboBoxSampleType.value

    ' Validate required fields
    If ParentBlockName = "" Then
        MsgBox "Vendor Block ID is required.", vbExclamation
        Exit Sub
    End If

    If anatomicSite = "" Then
        MsgBox "Anatomic Site is required.", vbExclamation
        Exit Sub
    End If

    If process = "" Then
        MsgBox "Process is required.", vbExclamation
        Exit Sub
    End If

    If site = "" Then
        MsgBox "Site is required.", vbExclamation
        Exit Sub
    End If

    ' Check if the Vendor Block ID already exists
    If Get_Rows_Table(ws, "BlocksTable", ParentBlockColName, ParentBlockName) <> -1 Then
        MsgBox "This Vendor Block ID already exists: " & ParentBlockName, vbExclamation
        Exit Sub
    End If

    ' Initialize MarkerScoresDict
    Set MarkerScoresDict = CreateObject("Scripting.Dictionary")
    For Each Marker In markers
        MarkerScoresDict.Add Marker, New Collection
    Next Marker

    ' Distribute scores across markers dynamically
    For Each score In scores
        Dim scoreParts() As String
        scoreParts = Split(score, ":")
        Dim markerName As String
        markerName = scoreParts(0)
        Dim scoreValue As String
        scoreValue = scoreParts(1)

        ' Prompt for score value
        Dim userInput As String
        userInput = Application.InputBox("Enter a value for marker '" & markerName & "' and score '" & scoreValue & "':", "Scoring Value", Type:=2)

        ' Handle empty score
        If userInput = "False" Or userInput = "" Then
            MarkerScoresDict(markerName).Add "[" & markerName & "]" & scoreValue ' Only add the score label
        Else
            MarkerScoresDict(markerName).Add "[" & markerName & "]" & scoreValue & ":" & userInput ' Add label with value
        End If
    Next score

    ' Construct Biomarker Characterisation
    biomarkerChar = ""
    For Each Marker In MarkerScoresDict.Keys
        If biomarkerChar <> "" Then
            biomarkerChar = biomarkerChar & "|"
        End If
        biomarkerChar = biomarkerChar & JoinCollection(MarkerScoresDict(Marker), "|")
    Next Marker

    ' Add a new row to the table
    Set NewRow = BlocksTable.ListRows.Add
    With NewRow
        .Range(BlocksTable.ListColumns(BlockStateColName).Index).value = "1-StockParent"
        .Range(BlocksTable.ListColumns(ParentBlockColName).Index).value = ParentBlockName
        .Range(BlocksTable.ListColumns(VendorInfoColName).Index).value = additionalInfo
        .Range(BlocksTable.ListColumns(AnatomicSiteColName).Index).value = anatomicSite
        .Range(BlocksTable.ListColumns(TumorTypeColName).Index).value = tumorType
        .Range(BlocksTable.ListColumns(VendorColName).Index).value = vendor
        .Range(BlocksTable.ListColumns(VendorBiomarkerColName).Index).value = biomarkerChar
        .Range(BlocksTable.ListColumns(ProcessColName).Index).value = process
        .Range(BlocksTable.ListColumns(SiteColName).Index).value = site
        .Range(BlocksTable.ListColumns(FixationtimeColName).Index).value = fixtime
        .Range(BlocksTable.ListColumns(FixativeColName).Index).value = fixative
        .Range(BlocksTable.ListColumns(SampleTypeColName).Index).value = sampletype
        .Range(BlocksTable.ListColumns(CreationDateColName).Index).value = creationdate
    End With

    ' Find the Labcorp ID column index
    labcorpIDColIndex = BlocksTable.ListColumns(ChildBlockColName).Index

    ' Identify the cell for the Labcorp ID in the new row
    Set labcorpNameCell = NewRow.Range(labcorpIDColIndex)

    ' Unlock the Settings sheet
    Set settingsWs = ThisWorkbook.Sheets(settingsSheet)
    settingsWs.Unprotect password:="settingsqc"

    ' Get the acronym and counter for the selected anatomic site
    Set anatomicTable = settingsWs.ListObjects(AcronymTable)
    anatomicColIndex = anatomicTable.ListColumns(Acronymfirstcolumn).Index
    acronymColIndex = anatomicTable.ListColumns("Acronym").Index
    counterColIndex = anatomicTable.ListColumns("Counter").Index

    For Each anatomicRow In anatomicTable.ListRows
        If anatomicRow.Range(anatomicColIndex).value = tumorType Then
            acronym = anatomicRow.Range(acronymColIndex).value
            anatomicRow.Range(counterColIndex).value = anatomicRow.Range(counterColIndex).value + 1
            newID = acronym & Format(anatomicRow.Range(counterColIndex).value, "0000")
            Exit For
        End If
    Next anatomicRow

    ' Re-lock the Settings sheet
    settingsWs.Protect password:="settingsqc"

    ' Assign the new ID to the Labcorp ID cell
    labcorpNameCell.value = newID

    ' Construct the hyperlink URL
    ChildBlockHyperlink = "https://labcorp.concentriq.proscia.com/dashboard/search?facets=[{%22field%22:{%22text%22:%22Image+name%22,%22searchType%22:%22name%22,%22facetGroup%22:%22image%22,%22resourceType%22:%22image%22},%22value%22:%22" & ParentBlockName & "%22},{%22field%22:{%22text%22:%22Image+name%22,%22searchType%22:%22name%22,%22facetGroup%22:%22image%22,%22resourceType%22:%22image%22},%22value%22:%22" & newID & "%22}]"

    ' Identify the cell where Parent Block Name is stored
    Set ChildBlockCell = NewRow.Range(BlocksTable.ListColumns(ChildBlockColName).Index)

    ' Add the hyperlink
    ws.Hyperlinks.Add Anchor:=ChildBlockCell, Address:=ChildBlockHyperlink, TextToDisplay:=newID

    ' Create folder and add hyperlink
    On Error GoTo HandleError
    FolderPath = MainFolderPath + "\" + anatomicSite + "\"
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath ' Create folder for anatomic site
    End If

    FolderPath = FolderPath + "\" + ParentBlockName
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If

    ' Identify the cell where Parent Block Name is stored
    Set ParentBlockCell = NewRow.Range(BlocksTable.ListColumns(ParentBlockColName).Index)

    ' Add the hyperlink
    ws.Hyperlinks.Add Anchor:=ParentBlockCell, Address:=FolderPath, TextToDisplay:=ParentBlockName

    MsgBox "Parent block added successfully!", vbInformation
    Exit Sub

HandleError:
    MsgBox "Error occurred while creating folder or hyperlink.", vbExclamation
    Unload Me
End Sub










' Function to get a single selected item from a ListBox
Private Function GetSingleSelection(lb As MSForms.ListBox) As String
    If lb.ListIndex <> -1 Then
        GetSingleSelection = lb.List(lb.ListIndex)
    Else
        GetSingleSelection = ""
    End If
End Function

' Function to get selected items from a multi-select ListBox
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

' Function to check if a Parent Block already exists in a table
Private Function Get_Rows_Table(ws As Worksheet, tableName As String, colName As String, searchValue As String) As Long
    Dim tbl As ListObject
    Dim colIndex As Long
    Dim foundRow As Long
    Dim i As Long

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If tbl Is Nothing Then
        Get_Rows_Table = -1
        Exit Function
    End If

    colIndex = tbl.ListColumns(colName).Index
    foundRow = -1

    For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange.Cells(i, colIndex).value = searchValue Then
            foundRow = i
            Exit For
        End If
    Next i

    Get_Rows_Table = foundRow
End Function
' Function to join a collection into a string separated by a delimiter
Private Function JoinCollection(col As Collection, delimiter As String) As String
    Dim result As String
    Dim item As Variant

    result = ""
    For Each item In col
        If result <> "" Then
            result = result & delimiter
        End If
        result = result & item
    Next item

    JoinCollection = result
End Function

Sub PopulateComboBoxSampleType(comboBox As MSForms.comboBox)
    Dim ws As Worksheet
    Dim sampleTypeTable As ListObject
    Dim row As ListRow

    ' Set the worksheet and table
    Set ws = ThisWorkbook.Sheets(settingsSheet)
    On Error Resume Next
    Set sampleTypeTable = ws.ListObjects("SampleType")
    On Error GoTo 0

    ' Exit if the table is not found
    If sampleTypeTable Is Nothing Then
        MsgBox "SampleType table not found on the Settings sheet.", vbExclamation
        Exit Sub
    End If

    ' Clear the ComboBox to remove any existing items
    comboBox.Clear

    ' Loop through each row in the table and add the first column value to the ComboBox
    For Each row In sampleTypeTable.ListRows
        comboBox.AddItem row.Range.Cells(1, 1).value
    Next row
End Sub




