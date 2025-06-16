Attribute VB_Name = "NewTMABlockModule"
Sub OpenNewTMABlockForm()
    Dim MyUserForm As NewTMABlockForm
    
    ' Initialiser les variables global
    SetVariables
    
    ' Créer un nouveau UserForm (pas encore afficher)
    Set MyUserForm = New NewTMABlockForm
    
    ' Afficher le UserForm
    MyUserForm.Show 0
End Sub
Function NewTMABlock(NumTMAStr As String, ListParentBlockNames As Collection)
    Dim ParentBlockNamesText As String
    Dim AnatomicSitesText As String
    Dim NumTMA As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    Dim firstColumn As Integer
    Dim TMABlockRow As Integer
    Dim TMABlockCol As Integer
    Dim BlockStateCol As Integer
    Dim ParentBlockCol As Integer
    Dim ChildBlockCol As Integer
    Dim AnatomicSiteCol As Integer
    Dim result As Variant
    Dim i As Integer, j As Integer
    Dim TMABlockNameExist As Boolean
    Dim ParentBlockRow As Long
    Dim AnatomicSiteText As String
    Dim FolderPath As String
    Dim RangeTMA As Range

    ' Initialize global variables
    SetVariables

    ' Ensure at least one Vendor Block ID is provided
    If ListParentBlockNames.Count = 0 Then
        MsgBox "Choose at least one Block ID."
        Exit Function
    End If

    ' Validate NumTMAStr is numeric
    If IsNumeric(NumTMAStr) Then
        NumTMA = CInt(NumTMAStr)
    Else
        MsgBox "Enter a number for the TMA count."
        Exit Function
    End If

    If NumTMA <= 0 Then
        MsgBox "Enter a positive number for the TMA count."
        Exit Function
    End If

    ' Determine the last row and column of the TMA table
    lastRow = GetLastRowFromTable(TmaWS, TmaTableName)
    lastColumn = GetLastColTable(TmaWS, TmaTableName)
    firstColumn = TmaWS.ListObjects(TmaTableName).Range.column

    ' Retrieve column indices
    TMABlockCol = GetColTable(TmaWS, TmaTableName, TMABlockColName)
    BlockStateCol = GetColTable(TmaWS, TmaTableName, BlockStateColName)
    ParentBlockCol = GetColTable(TmaWS, TmaTableName, TMAParentColName)
    AnatomicSiteCol = GetColTable(TmaWS, TmaTableName, AnatomicSiteColName)

    ' Create TMA blocks
    For i = 1 To NumTMA
        j = 0
        Do
            TMABlockRow = Get_Rows_Table(TmaWS, TmaTableName, TMABlockColName, GetDateWithLetter(i + j))
            TMABlockNameExist = TMABlockRow <> -1
            If TMABlockNameExist Then j = j + 1
        Loop Until Not TMABlockNameExist

        ' Define the range for the new TMA block
        Set RangeTMA = TmaWS.Range(TmaWS.Cells(lastRow + i, firstColumn), TmaWS.Cells(lastRow + i, lastColumn))

        ' Assign a name to the TMA block
        TmaWS.Cells(lastRow + i, TMABlockCol).value = GetDateWithLetter(i + j)

        ' Change the block state to stock
        TmaWS.Cells(lastRow + i, BlockStateCol).value = StockTMAText

        ' Fetch Vendor Block IDs and anatomic sites
        ParentBlockNamesText = ""
        AnatomicSitesText = ""
        For Each ParentBlockName In ListParentBlockNames
            ParentBlockRow = Get_Rows_Table(BlocksWS, BlocksTableName, ChildBlockColName, ParentBlockName)
            If ParentBlockRow = -1 Then
                MsgBox "Vendor Block ID not found: " & ParentBlockName
                Exit Function
            End If

            ParentBlockNamesText = GetMarkerText(CStr(ParentBlockName), ParentBlockNamesText)
            AnatomicSiteText = BlocksWS.Cells(ParentBlockRow, GetColTable(BlocksWS, BlocksTableName, AnatomicSiteColName)).value
            AnatomicSitesText = GetMarkerText(CStr(AnatomicSiteText), AnatomicSitesText)
        Next ParentBlockName

        ' Add Vendor Block IDs and anatomic sites to the TMA table
        TmaWS.Cells(lastRow + i, ParentBlockCol).value = ParentBlockNamesText
        TmaWS.Cells(lastRow + i, AnatomicSiteCol).value = AnatomicSitesText

        ' Create the folder and add the hyperlink
        On Error GoTo HandleError
        FolderPath = MainFolderPath + "\TMA\" + GetDateWithLetter(i + j) + "\"
        If Dir(FolderPath, vbDirectory) = "" Then
            ' Create the folder for the TMA if it doesn't exist
            MkDir FolderPath
        End If

        ' Add a hyperlink to the TMA name
        TmaWS.Hyperlinks.Add Anchor:=TmaWS.Cells(lastRow + i, TMABlockCol), _
            Address:=FolderPath, _
            TextToDisplay:=TmaWS.Cells(lastRow + i, TMABlockCol).value

    Next i

    TmaWS.Cells(lastRow + NumTMA, 1).Select
    Exit Function

HandleError:
    MsgBox "Error linking folder."
    Exit Function
End Function



Function AddParentToList(ParentBlockName As String, MyUserForm As Object)
    ' Initialize global variables
    SetVariables

    ' Check if the Vendor Block ID is empty
    If ParentBlockName = "" Then
        MsgBox "The Vendor Block ID is empty."
        Exit Function
    End If

    ' Temporarily remove any filters on BlocksWS
    Dim tbl As ListObject
    Set tbl = BlocksWS.ListObjects(BlocksTableName)
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0

    ' Verify that this Parent Block exists in the source worksheet
    ParentBlockRow = Get_ParentBlock_Rows(BlocksWS, BlocksTableName, ChildBlockColName, ParentBlockName)
    If ParentBlockRow = -1 Then
        MsgBox "This Block ID is not found: " & ParentBlockName
        ' Reapply the filter if it was removed
        tbl.AutoFilter.ApplyFilter
        Exit Function
    End If

    ' Retrieve the list of Vendor Block IDs
    Dim ListParentBlockName As Collection
    Set ListParentBlockName = New Collection ' Create an empty list

    Dim ParentExist As Boolean
    ParentExist = False

    ' Check if the parent block is already in the listbox
    For i = 0 To MyUserForm.ListBox1.ListCount - 1
        If MyUserForm.ListBox1.List(i) = ParentBlockName Then
            ParentExist = True
        End If
    Next i

    ' Add the Vendor Block ID to the listbox if it doesn't exist
    If Not ParentExist Then
        MyUserForm.ListBox1.AddItem ParentBlockName
    End If

    ' Reapply the filter if it was removed
    tbl.AutoFilter.ApplyFilter
End Function



Function GetLastRowFromTable(ws As Worksheet, tableName As String) As Integer
    Dim lo As ListObject
    Set lo = ws.ListObjects(tableName)
    GetLastRowFromTable = lo.ListRows.Count + lo.HeaderRowRange.row
End Function




