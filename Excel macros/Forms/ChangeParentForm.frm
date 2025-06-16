VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeParentForm 
   Caption         =   "Change TMA Parent"
   ClientHeight    =   6684
   ClientLeft      =   35
   ClientTop       =   168
   ClientWidth     =   7924
   OleObjectBlob   =   "ChangeParentForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChangeParentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ' Initialize global variables
    SetVariables
    
    ' Remove old Parent Block Names from the list box
    Me.ListBox1.Clear
    
    ' Retrieve the TMA Block Name from TextBox1
    TMABlockName = Me.TextBox1.value
    
    ' Check if the TMA Block Name is empty
    If TMABlockName = "" Then
        MsgBox "Le TMA Block Name est vide."
        Exit Sub
    End If
    
    ' Check that this TMA Block exists in the base worksheet (TMA Table)
    TMABlockRow = Get_Rows_Table(TmaWS, TmaTableName, TMABlockColName, TMABlockName)
    If TMABlockRow = -1 Then
        MsgBox "This TMA Block Name is not found: " & TMABlockName
        Exit Sub
    End If
    
    ''' Retrieve the list of Parent Block Names from the TMA table
    ' Get the column of Parent Block Names
    ParentBlockCol = GetColTable(TmaWS, TmaTableName, TMAParentColName)
    
    ' Retrieve the Parent Block Names text for the TMA
    ParentBlockNamesText = TmaWS.Cells(TMABlockRow, ParentBlockCol).value
    
    ' Fill the list box by splitting the text
    Dim ParentBlockNamesList() As String
    ParentBlockNamesList = Split(ParentBlockNamesText, "|")
    For Each ParentBlockName In ParentBlockNamesList
        Me.ListBox1.AddItem ParentBlockName
    Next
End Sub

Private Sub CommandButton2_Click()
    Dim ParentBlockNamesText As String
    Dim ParentBlockName As String
    Dim tbl As ListObject

    ''' Add a Parent Block to the list and update the TMA record

    ' Initialize global variables
    SetVariables
    
    ' Retrieve the TMA Block Name from TextBox1
    TMABlockName = Me.TextBox1.value
    
    ' Check if the TMA Block Name is empty
    If TMABlockName = "" Then
        MsgBox "Le TMA Block Name est vide."
        Exit Sub
    End If
    
    ' Check that this TMA Block exists in the TMA table
    TMABlockRow = Get_Rows_Table(TmaWS, TmaTableName, TMABlockColName, TMABlockName)
    If TMABlockRow = -1 Then
        MsgBox "This TMA Block Name is not found: " & TMABlockName
        Exit Sub
    End If
    
    ' Retrieve the Parent Block Name
    ParentBlockName = Me.TextBox2.value
    
    ' Check if the Parent Block Name is empty
    If ParentBlockName = "" Then
        MsgBox "Le Parent Block Name est vide."
        Exit Sub
    End If
    
    ' Verify that this Parent Block exists in the Blocks table
    ParentBlockRow = Get_ParentBlock_Rows(BlocksWS, BlocksTableName, ParentBlockColName, ParentBlockName)
    If ParentBlockRow = -1 Then
        MsgBox "This Parent Block Name is not found: " & ParentBlockName
        Exit Sub
    End If
    
    ''' Retrieve the current list of Parent Block Names from the TMA table
    ParentBlockCol = GetColTable(TmaWS, TmaTableName, TMAParentColName)
    ParentBlockNamesText = TmaWS.Cells(TMABlockRow, ParentBlockCol).value
    
    ' Add the Parent Block Name to the list box if it is not already there
    AlreadyAdded = False
    Dim i As Integer
    For i = 0 To Me.ListBox1.ListCount - 1
        If CStr(Me.ListBox1.List(i)) = CStr(ParentBlockName) Then
            AlreadyAdded = True
            Exit For
        End If
    Next i
    
    If Not AlreadyAdded Then
        Me.ListBox1.AddItem ParentBlockName
    End If

End Sub

Private Sub CommandButton3_Click()
    ''' Remove Parent Block from the list box without updating the original TMA row

    ' Initialize global variables
    SetVariables
    
    ' Retrieve the TMA Block Name from TextBox1
    TMABlockName = Me.TextBox1.value
    
    ' Check if the TMA Block Name is empty
    If TMABlockName = "" Then
        MsgBox "Le TMA Block Name est vide."
        Exit Sub
    End If
    
    ' Verify that the TMA Block exists in the TMA table
    TMABlockRow = Get_Rows_Table(TmaWS, TmaTableName, TMABlockColName, TMABlockName)
    If TMABlockRow = -1 Then
        MsgBox "This TMA Block Name is not found: " & TMABlockName
        Exit Sub
    End If
    
    ' Check if any items are selected in the list box
    Dim SelectedCount As Integer
    SelectedCount = 0
    Dim i As Integer
    For i = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) = True Then
            SelectedCount = SelectedCount + 1
        End If
    Next i
    
    If SelectedCount = 0 Then
        MsgBox "Select at least one Parent Block to remove."
        Exit Sub
    End If

    ' Remove selected items from the list box
    Dim idx As Integer
    For i = Me.ListBox1.ListCount - 1 To 0 Step -1
        If Me.ListBox1.Selected(i) = True Then
            Me.ListBox1.RemoveItem i
        End If
    Next i

    
End Sub


Private Sub CommandButton4_Click()
    Dim NewTMABlockName As String
    Dim BaseTMABlockName As String
    Dim ParentBlockName As Variant
    Dim ParentBlockNamesText As String
    Dim bis_text As String
    Dim AnatomicSitesText As String, AnatomicSiteText As String
    Dim i As Integer, j As Integer
    Dim lastRow As Long, NewTMARow As Long
    Dim TMABlockRow As Long, ParentBlockRow As Long
    Dim SearchTMABlockRow As Long
    Dim TMATable As ListObject
    Dim TMABlockCol As Integer

    ' Retrieve the TMA Block Name from TextBox1
    TMABlockName = Me.TextBox1.value
    TMABlockCol = TmaWS.ListObjects(TmaTableName).HeaderRowRange.Cells.Find(TMABlockColName).column

    ' Check if the TMA Block Name is empty
    If TMABlockName = "" Then
        MsgBox "Le TMA Block Name est vide."
        Exit Sub
    End If
    
    ' Verify that the TMA Block exists in the TMA table
    TMABlockRow = Get_Rows_Table(TmaWS, TmaTableName, TMABlockColName, TMABlockName)
    If TMABlockRow = -1 Then
        MsgBox "This TMA Block Name is not found: " & TMABlockName
        Exit Sub
    End If
    
    ' Create a collection of Parent Block Names from the list box
    Dim ListParentBlockNames As Collection
    Set ListParentBlockNames = New Collection
    For i = 0 To Me.ListBox1.ListCount - 1
        ListParentBlockNames.Add Me.ListBox1.List(i)
    Next i
    
    If ListParentBlockNames.Count = 0 Then
        MsgBox "Select at least one Parent Block."
        Exit Sub
    End If
    
    ' Build the Parent Block Names text from the collection
    ParentBlockNamesText = ""
    For Each ParentBlockName In ListParentBlockNames
        ParentBlockNamesText = GetMarkerText(CStr(ParentBlockName), ParentBlockNamesText)
    Next
    
    ''' Generate a new TMA Block Name with "bis" appended
    ' Extract the base TMA (if it already contains "bis", remove that part)
    BaseTMABlockName = ExtractBaseTMA(CStr(TMABlockName))

    bis_text = "bis"
    j = 1
    Do
        NewTMABlockName = BaseTMABlockName & bis_text & CStr(j)
        SearchTMABlockRow = Get_Rows_Table(TmaWS, TmaTableName, TMABlockColName, NewTMABlockName)
        If SearchTMABlockRow = -1 Then Exit Do
        j = j + 1
    Loop
    
    ' Copy the original TMA row to create a new one
    Set TMATable = TmaWS.ListObjects(TmaTableName)
    lastRow = GetLastRowTable(TmaWS, TmaTableName)
    NewTMARow = lastRow + 1
    TmaWS.Cells(NewTMARow, TMABlockCol).value = "temp"
    TmaWS.Rows(TMABlockRow).Copy TmaWS.Rows(NewTMARow)
    
    ' Update the new row with the new TMA Block Name and Parent Block Names
    TmaWS.Cells(NewTMARow, TMABlockCol).value = NewTMABlockName
    ParentBlockCol = GetColTable(TmaWS, TmaTableName, TMAParentColName)
    TmaWS.Cells(NewTMARow, ParentBlockCol).value = ParentBlockNamesText
    
    ' Update the Anatomic Site for the new TMA Block (retrieved from the Blocks table)
    Dim ParentBlockNamesList() As String
    ParentBlockNamesList = Split(ParentBlockNamesText, "|")
    AnatomicSitesText = ""
    For Each ParentBlockName In ParentBlockNamesList
        ParentBlockRow = Get_ParentBlock_Rows(BlocksWS, BlocksTableName, ParentBlockColName, CStr(ParentBlockName))
        If ParentBlockRow = -1 Then
            MsgBox "This Parent Block Name is not found: " & ParentBlockName
            Exit Sub
        End If
        AnatomicSiteText = BlocksWS.Cells(ParentBlockRow, GetColTable(BlocksWS, BlocksTableName, AnatomicSiteColName)).value
        AnatomicSitesText = GetMarkerText(CStr(AnatomicSiteText), CStr(AnatomicSitesText))
    Next
    TmaWS.Cells(NewTMARow, GetColTable(TmaWS, TmaTableName, AnatomicSiteColName)).value = AnatomicSitesText
    
    ' Mark the original TMA Block as exhausted (or non-validated)
    TmaWS.Cells(TMABlockRow, GetColTable(TmaWS, TmaTableName, BlockStateColName)).value = ExhaustedBlockText
    
    ' Create the folder and assign the hyperlink to the new TMA Block
    On Error GoTo HandleError
    Dim FolderPath As String
    FolderPath = MainFolderPath & "\TMA\" & NewTMABlockName & "\"
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If
    SetHyperlinkToCell TmaWS.Cells(NewTMARow, TMABlockCol), FolderPath, NewTMABlockName
    Exit Sub
    
HandleError:
    MsgBox "Problem to link to folder: " & FolderPath
    Exit Sub
End Sub

'function to extract the base TMA name (removing any "bis" part)
Function ExtractBaseTMA(TMAName As String)
    If ExtractBefore(TMAName, "bis") = "" Then
        ExtractBaseTMA = TMAName
    Else
        ExtractBaseTMA = ExtractBefore(TMAName, "bis")
    End If
End Function

'function to extract text before a target string
Function ExtractBefore(inputString As String, targetString As String) As Variant
    Dim pos As Long
    pos = InStr(1, inputString, targetString)
    If pos > 0 Then
        ExtractBefore = Left(inputString, pos - 1)
    Else
        ExtractBefore = ""
    End If
End Function


