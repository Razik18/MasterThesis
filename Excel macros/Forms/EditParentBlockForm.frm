VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditParentBlockForm 
   Caption         =   "Create New Parent Block"
   ClientHeight    =   9576.001
   ClientLeft      =   -231
   ClientTop       =   -861
   ClientWidth     =   17115
   OleObjectBlob   =   "EditParentBlockForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditParentBlockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    SetVariables
    On Error GoTo ErrHandler
    
    ' Populate all listboxes and combobox with master data
    PopulateListBox Me.ListBoxAnatomic, "AnatomicSiteTable"
    PopulateListBox Me.ListBoxTumor, "TumorTypeTable"
    PopulateListBox Me.ListBoxVendor, "VendorsTable"
    PopulateListBox Me.ListBoxMarker, "MarkersTable"
    PopulateListBox Me.ListBoxProcess, "ProcessTable"
    PopulateListBox Me.ListBoxSite, "SitesTable"
    PopulateListBox Me.ListBoxFixative, "FixativeTable"
    PopulateComboBoxSampleType Me.ComboBoxSampleType
    
    ' Clear the scoring list box initially
    Me.ListBoxScore.Clear
    
    If SelectedRowIndex = 0 Then
        MsgBox "No row selected. Closing form.", vbExclamation
        Unload Me
        Exit Sub
    End If
    
    ' Access the BlocksTable row
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(blocksSheet)
    Dim BlocksTable As ListObject: Set BlocksTable = ws.ListObjects("BlocksTable")
    
    If SelectedRowIndex < 1 Or SelectedRowIndex > BlocksTable.ListRows.Count Then
        MsgBox "Invalid row selection. Closing form.", vbExclamation
        Unload Me
        Exit Sub
    End If
    
    Dim targetRow As ListRow
    Set targetRow = BlocksTable.ListRows(SelectedRowIndex)
    
    ' Read existing values from that row
    Dim ParentBlockName As String
    Dim additionalInfo As String
    Dim fixtime As String
    Dim creationdate As String
    Dim anatomicSite As String
    Dim tumorType As String
    Dim vendor As String
    Dim process As String
    Dim site As String
    Dim fixative As String
    Dim sampletype As String
    Dim biomarkerChar As String
    
    ParentBlockName = targetRow.Range(BlocksTable.ListColumns(ParentBlockColName).Index).value
    additionalInfo = targetRow.Range(BlocksTable.ListColumns(VendorInfoColName).Index).value
    fixtime = targetRow.Range(BlocksTable.ListColumns(FixationtimeColName).Index).value
    creationdate = targetRow.Range(BlocksTable.ListColumns(CreationDateColName).Index).value
    anatomicSite = targetRow.Range(BlocksTable.ListColumns(AnatomicSiteColName).Index).value
    tumorType = targetRow.Range(BlocksTable.ListColumns(TumorTypeColName).Index).value
    vendor = targetRow.Range(BlocksTable.ListColumns(VendorColName).Index).value
    process = targetRow.Range(BlocksTable.ListColumns(ProcessColName).Index).value
    site = targetRow.Range(BlocksTable.ListColumns(SiteColName).Index).value
    fixative = targetRow.Range(BlocksTable.ListColumns(FixativeColName).Index).value
    sampletype = targetRow.Range(BlocksTable.ListColumns(SampleTypeColName).Index).value
    biomarkerChar = targetRow.Range(BlocksTable.ListColumns(VendorBiomarkerColName).Index).value
    
    ' Fill the form controls
    Me.TextBox1.value = ParentBlockName            ' Vendor Block ID
    Me.TextBox2.value = additionalInfo
    Me.TextBox3.value = fixtime
    Me.TextBox4.value = creationdate
    
    ' Preselect single-select ListBoxes
    PreselectListBox Me.ListBoxAnatomic, anatomicSite
    PreselectListBox Me.ListBoxTumor, tumorType
    PreselectListBox Me.ListBoxVendor, vendor
    PreselectListBox Me.ListBoxProcess, process
    PreselectListBox Me.ListBoxSite, site
    PreselectListBox Me.ListBoxFixative, fixative
    
    ' ComboBox for SampleType
    Me.ComboBoxSampleType.value = sampletype
    
    ' Parse existing biomarkers to preselect Markers,
    ' then populate scoring list and preselect existing scores.
    ParseAndPreselectMarkers biomarkerChar, Me.ListBoxMarker
    Me.ListBoxScore.Clear
    PopulateScoringList Me.ListBoxMarker, Me.ListBoxScore
    PreselectScores biomarkerChar, Me.ListBoxScore

Exit Sub

ErrHandler:
    MsgBox "Error in UserForm_Initialize: " & Err.Description, vbExclamation
    Unload Me
End Sub

Private Sub ListBoxMarker_Change()
    Dim savedSelections As New Collection
    Dim i As Long, sItem As Variant
    ' Save currently selected scoring items
    For i = 0 To Me.ListBoxScore.ListCount - 1
        If Me.ListBoxScore.Selected(i) Then
            savedSelections.Add Me.ListBoxScore.List(i)
        End If
    Next i
    
    Me.ListBoxScore.Clear
    PopulateScoringList Me.ListBoxMarker, Me.ListBoxScore
    
    ' Re-select any scoring item that was previously selected if it exists
    For Each sItem In savedSelections
        For i = 0 To Me.ListBoxScore.ListCount - 1
            If Me.ListBoxScore.List(i) = sItem Then
                Me.ListBoxScore.Selected(i) = True
            End If
        Next i
    Next sItem
End Sub


Private Sub CommandButtonOK_Click()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(blocksSheet)
    Dim BlocksTable As ListObject: Set BlocksTable = ws.ListObjects("BlocksTable")
    
    If SelectedRowIndex < 1 Or SelectedRowIndex > BlocksTable.ListRows.Count Then
        MsgBox "Invalid row selection.", vbExclamation
        Exit Sub
    End If
    
    Dim targetRow As ListRow
    Set targetRow = BlocksTable.ListRows(SelectedRowIndex)
    
    ' Gather form data
    Dim ParentBlockName As String: ParentBlockName = Me.TextBox1.Text
    Dim additionalInfo As String: additionalInfo = Me.TextBox2.Text
    Dim fixtime As String: fixtime = Me.TextBox3.Text
    Dim creationdate As String: creationdate = Me.TextBox4.Text
    Dim anatomicSite As String: anatomicSite = GetSingleSelection(Me.ListBoxAnatomic)
    Dim tumorType As String: tumorType = GetSingleSelection(Me.ListBoxTumor)
    Dim vendor As String: vendor = GetSingleSelection(Me.ListBoxVendor)
    Dim process As String: process = GetSingleSelection(Me.ListBoxProcess)
    Dim site As String: site = GetSingleSelection(Me.ListBoxSite)
    Dim fixative As String: fixative = GetSingleSelection(Me.ListBoxFixative)
    Dim sampletype As String: sampletype = Me.ComboBoxSampleType.value
    
    ' Markers + Scores
    Dim markers As Collection
    Set markers = GetSelectedItems(Me.ListBoxMarker)
    Dim scores As Collection
    Set scores = GetSelectedItems(Me.ListBoxScore)
    
    If ParentBlockName = "" Then
        MsgBox "Vendor Block ID is required.", vbExclamation
        Exit Sub
    End If
    
    Dim MarkerScoresDict As Object
    Set MarkerScoresDict = CreateObject("Scripting.Dictionary")
    
    Dim mk As Variant
    For Each mk In markers
        MarkerScoresDict.Add mk, New Collection
    Next mk
    
    Dim existingScores As Object
    Dim oldBiomarkerChar As String
    oldBiomarkerChar = targetRow.Range(BlocksTable.ListColumns(VendorBiomarkerColName).Index).value
    Set existingScores = ParseExistingScores(oldBiomarkerChar)
    
    Dim i As Long
    For Each mk In markers
        For i = 1 To scores.Count
            Dim s As String
            s = scores(i)
            If InStr(s, "[" & mk & "]") = 1 Then
                Dim tempKey As String
                tempKey = mk & "|" & s
                
                Dim defaultVal As String
                If existingScores.Exists(tempKey) Then
                    defaultVal = existingScores(tempKey)
                Else
                    defaultVal = ""
                End If
                
                Dim userVal As Variant
                userVal = Application.InputBox( _
                            Prompt:="Enter a value for marker '" & mk & "' and score '" & s & "':", _
                            Title:="Scoring Value", _
                            Default:=defaultVal, _
                            Type:=2)
                
                If userVal = False Then userVal = ""
                
                If userVal = "" Then
                    MarkerScoresDict(mk).Add s
                Else
                    MarkerScoresDict(mk).Add s & ":" & userVal
                End If
            End If
        Next i
    Next mk
    
    Dim biomarkerChar As String
    biomarkerChar = ""
    Dim dictKey As Variant
    For Each dictKey In MarkerScoresDict.Keys
        Dim joinedScores As String
        joinedScores = JoinCollection(MarkerScoresDict(dictKey), "|")
        
        If biomarkerChar <> "" Then biomarkerChar = biomarkerChar & "|"
        biomarkerChar = biomarkerChar & joinedScores
    Next dictKey
    
    With targetRow
        .Range(BlocksTable.ListColumns(ParentBlockColName).Index).value = ParentBlockName
        .Range(BlocksTable.ListColumns(VendorInfoColName).Index).value = additionalInfo
        .Range(BlocksTable.ListColumns(AnatomicSiteColName).Index).value = anatomicSite
        .Range(BlocksTable.ListColumns(TumorTypeColName).Index).value = tumorType
        .Range(BlocksTable.ListColumns(VendorColName).Index).value = vendor
        .Range(BlocksTable.ListColumns(ProcessColName).Index).value = process
        .Range(BlocksTable.ListColumns(SiteColName).Index).value = site
        .Range(BlocksTable.ListColumns(FixationtimeColName).Index).value = fixtime
        .Range(BlocksTable.ListColumns(FixativeColName).Index).value = fixative
        .Range(BlocksTable.ListColumns(SampleTypeColName).Index).value = sampletype
        .Range(BlocksTable.ListColumns(CreationDateColName).Index).value = creationdate
        .Range(BlocksTable.ListColumns(VendorBiomarkerColName).Index).value = biomarkerChar
        
        ' If the block has no Labcorp Block ID, generate one; otherwise leave it
        Dim labcorpIDColIndex As Long
        labcorpIDColIndex = BlocksTable.ListColumns(ChildBlockColName).Index
        If .Range(labcorpIDColIndex).value = "" Then
            .Range(labcorpIDColIndex).value = GenerateLabcorpID(anatomicSite)
        End If
    End With
    

    Dim FolderPath As String
    FolderPath = MainFolderPath & "\" & anatomicSite & "\"
    If Dir(FolderPath, vbDirectory) = "" Then MkDir FolderPath
    FolderPath = FolderPath & ParentBlockName
    If Dir(FolderPath, vbDirectory) = "" Then MkDir FolderPath
    
    ' Add a hyperlink to the Vendor Block ID cell (folder hyperlink)
    Dim VendorCell As Range
    Set VendorCell = targetRow.Range(BlocksTable.ListColumns(ParentBlockColName).Index)
    On Error Resume Next
    VendorCell.Hyperlinks.Delete
    On Error GoTo 0
    ws.Hyperlinks.Add Anchor:=VendorCell, Address:=FolderPath, TextToDisplay:=ParentBlockName
    
    ' Get Labcorp Block ID from the row and add its hyperlink if available
    Dim LabcorpID As String
    LabcorpID = targetRow.Range(BlocksTable.ListColumns(ChildBlockColName).Index).value
    
    If Trim(LabcorpID) <> "" Then
        Dim ParentBlockHyperlink As String
        ParentBlockHyperlink = "https://labcorp.concentriq.proscia.com/dashboard/search?facets=[{%22field%22:{%22text%22:%22Image+name%22,%22searchType%22:%22name%22,%22facetGroup%22:%22image%22,%22resourceType%22:%22image%22},%22value%22:%22" & ParentBlockName & "%22},{%22field%22:{%22text%22:%22Image+name%22,%22searchType%22:%22name%22,%22facetGroup%22:%22image%22,%22resourceType%22:%22image%22},%22value%22:%22" & LabcorpID & "%22}]"
        
        Dim LabcorpCell As Range
        Set LabcorpCell = targetRow.Range(BlocksTable.ListColumns(ChildBlockColName).Index)
        On Error Resume Next
        LabcorpCell.Hyperlinks.Delete
        On Error GoTo 0
        ws.Hyperlinks.Add Anchor:=LabcorpCell, Address:=ParentBlockHyperlink, TextToDisplay:=LabcorpID
    End If
    
    MsgBox "Parent block updated successfully!", vbInformation
    Unload Me
    Exit Sub
    
ErrHandler:
    MsgBox "Error in CommandButtonOK_Click: " & Err.Description, vbExclamation
End Sub


Private Sub CommandButtonCancel_Click()
    Unload Me
End Sub


Private Sub PopulateListBox(lb As MSForms.ListBox, tableName As String)
    Dim ws As Worksheet, tbl As ListObject, lr As ListRow
    Set ws = ThisWorkbook.Sheets(settingsSheet)
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    If Not tbl Is Nothing Then
        lb.Clear
        For Each lr In tbl.ListRows
            lb.AddItem lr.Range.Cells(1, 1).value
        Next lr
    End If
End Sub

Private Sub PopulateComboBoxSampleType(cmb As MSForms.comboBox)
    Dim ws As Worksheet, sampleTypeTable As ListObject, row As ListRow
    Set ws = ThisWorkbook.Sheets(settingsSheet)
    On Error Resume Next
    Set sampleTypeTable = ws.ListObjects("SampleType")
    On Error GoTo 0
    If sampleTypeTable Is Nothing Then Exit Sub
    cmb.Clear
    For Each row In sampleTypeTable.ListRows
        cmb.AddItem row.Range.Cells(1, 1).value
    Next row
End Sub

Private Sub PreselectListBox(lb As MSForms.ListBox, valueToFind As String)
    Dim i As Long
    For i = 0 To lb.ListCount - 1
        If lb.List(i) = valueToFind Then
            lb.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Function GetSingleSelection(lb As MSForms.ListBox) As String
    If lb.ListIndex <> -1 Then
        GetSingleSelection = lb.List(lb.ListIndex)
    Else
        GetSingleSelection = ""
    End If
End Function

Private Function GetSelectedItems(lb As MSForms.ListBox) As Collection
    Dim selectedItems As New Collection, i As Long
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then selectedItems.Add lb.List(i)
    Next i
    Set GetSelectedItems = selectedItems
End Function

Private Sub ParseAndPreselectMarkers(biomarkerChar As String, markerBox As MSForms.ListBox)
    If Trim(biomarkerChar) = "" Then Exit Sub
    Dim markerGroups() As String
    markerGroups = Split(biomarkerChar, "|")
    Dim dict As Object, group As Variant, items() As String, it As Variant, s As String, markerName As String, i As Long
    Set dict = CreateObject("Scripting.Dictionary")
    For Each group In markerGroups
        items = Split(group, ",")
        For Each it In items
            s = Trim(it)
            markerName = ExtractMarkerName(s)
            If markerName <> "" Then
                If Not dict.Exists(markerName) Then dict(markerName) = True
            End If
        Next it
    Next group
    For i = 0 To markerBox.ListCount - 1
        If dict.Exists(markerBox.List(i)) Then markerBox.Selected(i) = True
    Next i
End Sub

Private Function ExtractMarkerName(s As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(s, "[")
    endPos = InStr(s, "]")
    If startPos > 0 And endPos > startPos Then
        ExtractMarkerName = Mid(s, startPos + 1, endPos - startPos - 1)
    End If
End Function

Private Sub PopulateScoringList(markerBox As MSForms.ListBox, scoreBox As MSForms.ListBox)
    Dim ws As Worksheet, scores As New Collection, i As Long
    Set ws = ThisWorkbook.Sheets(settingsSheet)
    Dim Marker As String, tableName As String, scoringTable As ListObject, j As Long, scoreVal As String
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
                    scoreVal = scoringTable.ListRows(j).Range.Cells(1, 1).value
                    On Error Resume Next
                    scores.Add "[" & Marker & "]" & scoreVal, CStr("[" & Marker & "]" & scoreVal)
                    On Error GoTo 0
                Next j
            End If
        End If
    Next i
    scoreBox.Clear
    Dim sc As Variant
    For Each sc In scores
        scoreBox.AddItem sc
    Next sc
End Sub

Private Function ParseExistingScores(biomarkerChar As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If Trim(biomarkerChar) = "" Then
        Set ParseExistingScores = dict
        Exit Function
    End If
    Dim markerGroups() As String, group As Variant, items() As String, it As Variant, s As String, markerName As String, afterBracket As String, scoreLabel As String, scoreValue As String, key As String
    markerGroups = Split(biomarkerChar, "|")
    For Each group In markerGroups
        items = Split(group, ",")
        For Each it In items
            s = Trim(it)
            markerName = ExtractMarkerName(s)
            If markerName <> "" Then
                afterBracket = Replace(s, "[" & markerName & "]", "")
                If InStr(afterBracket, ":") > 0 Then
                    scoreLabel = Split(afterBracket, ":")(0)
                    scoreValue = Split(afterBracket, ":")(1)
                Else
                    scoreLabel = afterBracket
                    scoreValue = ""
                End If
                key = markerName & "|" & "[" & markerName & "]" & scoreLabel
                dict(key) = scoreValue
            End If
        Next it
    Next group
    Set ParseExistingScores = dict
End Function

Private Function JoinCollection(col As Collection, delimiter As String) As String
    Dim result As String, item As Variant
    For Each item In col
        If result <> "" Then result = result & delimiter
        result = result & item
    Next item
    JoinCollection = result
End Function

Private Sub PreselectScores(biomarkerChar As String, scoreBox As MSForms.ListBox)
    Dim existingScores As Object, i As Long, scoreItem As String, mkName As String, key As String
    Set existingScores = ParseExistingScores(biomarkerChar)
    For i = 0 To scoreBox.ListCount - 1
        scoreItem = scoreBox.List(i)
        mkName = ExtractMarkerName(scoreItem)
        If mkName <> "" Then
            key = mkName & "|" & scoreItem
            If existingScores.Exists(key) Then
                scoreBox.Selected(i) = True
            End If
        End If
    Next i
End Sub

Private Function GenerateLabcorpID(anatomicSite As String) As String
    Dim settingsWs As Worksheet
    Set settingsWs = ThisWorkbook.Sheets(settingsSheet)
    settingsWs.Unprotect password:="settingsqc"
    Dim anatomicTable As ListObject
    Set anatomicTable = settingsWs.ListObjects("AnatomicSiteTable")
    Dim anatomicColIndex As Long, acronymColIndex As Long, counterColIndex As Long
    anatomicColIndex = anatomicTable.ListColumns("Anatomique Site").Index
    acronymColIndex = anatomicTable.ListColumns("Acronym").Index
    counterColIndex = anatomicTable.ListColumns("Counter").Index
    Dim acronym As String, newID As String, anatomicRow As ListRow
    For Each anatomicRow In anatomicTable.ListRows
        If anatomicRow.Range(anatomicColIndex).value = anatomicSite Then
            acronym = anatomicRow.Range(acronymColIndex).value
            anatomicRow.Range(counterColIndex).value = anatomicRow.Range(counterColIndex).value + 1
            newID = acronym & Format(anatomicRow.Range(counterColIndex).value, "0000")
            Exit For
        End If
    Next anatomicRow
    settingsWs.Protect password:="settingsqc"
    GenerateLabcorpID = newID
End Function


