Attribute VB_Name = "ContextualMenuModule"
Dim CustomContextMenu As CommandBarControl


Sub AddCustomContextMenu()
    Dim contextMenu As CommandBar

    ' Get the List Range Popup (table row) context menu
    On Error Resume Next
    Set contextMenu = Application.CommandBars("List Range Popup")
    On Error GoTo 0

    ' Remove any previous custom items if they exist
    RemoveCustomContextMenu

    If Not contextMenu Is Nothing Then
        ' Insert "Open Folder" menu item at the top
        Set CustomContextMenu = contextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True, Before:=1)
        With CustomContextMenu
            .Caption = "Open Folder"
            .OnAction = "GoToFolder"
            .BeginGroup = True
        End With

        ' Insert "Open Result Form" menu item at the top
        Set CustomContextMenu = contextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True, Before:=1)
        With CustomContextMenu
            .Caption = "Open Result Form"
            .OnAction = "OpenResultsForm"
            .BeginGroup = True
        End With

        ' Insert "Send Block in Review" menu item at the top
        Set CustomContextMenu = contextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True, Before:=1)
        With CustomContextMenu
            .Caption = "Send Block in Review"
            .OnAction = "OpenReviewsForm"
            .BeginGroup = True
        End With

        ' Insert "Create Child Block" menu item at the top
        Set CustomContextMenu = contextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True, Before:=1)
        With CustomContextMenu
            .Caption = "Create Child Block"
            .OnAction = "OpenNewChildsBlockForm"
        End With
        
        ' Insert "Edit Parent Block" menu item at the top
        Set CustomContextMenu = contextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True, Before:=1)
        With CustomContextMenu
            .Caption = "Edit Parent Block"
            .OnAction = "OpenEditParentBlockForm"
            .BeginGroup = True
        End With
    End If
End Sub

Sub RemoveCustomContextMenu()
    On Error Resume Next
    If Not CustomContextMenu Is Nothing Then
        CustomContextMenu.Delete
    End If
    On Error GoTo 0
End Sub

Sub OpenEditParentBlockForm()
    Dim ws As Worksheet
    Dim BlocksTable As ListObject
    Dim selectedRow As Range
    Dim parentBlockColumn As ListColumn
    Dim ParentBlockName As String
    SetVariables
    ' Set the worksheet and table
    Set ws = ThisWorkbook.Sheets(blocksSheet)
    Set BlocksTable = ws.ListObjects("BlocksTable")

    ' Get the selected cell
    If TypeName(Application.Selection) = "Range" Then
        Set selectedRow = Application.Selection
    Else
        MsgBox "Please select a valid cell in the table.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set parentBlockColumn = BlocksTable.ListColumns(ParentBlockColName)
    On Error GoTo 0

    If parentBlockColumn Is Nothing Then
        MsgBox "'Vendor Block ID' column not found in the table.", vbExclamation
        Exit Sub
    End If

    ParentBlockName = selectedRow.EntireRow.Cells(1, parentBlockColumn.Index).value

    If Trim(ParentBlockName) = "" Then
        MsgBox "Block has no VendorID.", vbExclamation
        Exit Sub
    End If

    ' Set the global SelectedRowIndex (must be declared in a standard module)
    SelectedRowIndex = selectedRow.row - BlocksTable.DataBodyRange.row + 1

    ' Open the EditParentBlockForm
    EditParentBlockForm.Show
End Sub

Sub GoToFolder()
    Dim currentRow As Long
    Dim ws As Worksheet
    Dim anatomicSite As String
    Dim ParentBlockName As String
    Dim FolderPath As String
    SetVariables
    ' Get the active worksheet and current row
    Set ws = ActiveSheet
    currentRow = ActiveCell.row

    ' Get Anatomic Site and Vendor Block ID from the respective columns
    On Error Resume Next
    anatomicSite = ws.Cells(currentRow, ws.ListObjects("BlocksTable").ListColumns(AnatomicSiteColName).Index).value
    ParentBlockName = ws.Cells(currentRow, ws.ListObjects("BlocksTable").ListColumns(ParentBlockColName).Index).value
    On Error GoTo 0

    ' Validate required fields
    If anatomicSite = "" Or ParentBlockName = "" Then
        MsgBox "Anatomic Site or Vendor Block ID is missing in the selected row.", vbExclamation
        Exit Sub
    End If

    ' Construct the folder path
    FolderPath = MainFolderPath & "\" & anatomicSite & "\" & ParentBlockName

    ' Check if the folder exists
    If Dir(FolderPath, vbDirectory) <> "" Then
        ' Open the folder
        Shell "explorer.exe " & FolderPath, vbNormalFocus
    Else
        MsgBox "The folder does not exist: " & FolderPath, vbExclamation
    End If
End Sub



Sub OpenResultsForm()
    Dim ws As Worksheet
    Dim BlocksTable As ListObject
    Dim selectedRow As Range
    Dim parentBlockColumn As ListColumn
    Dim childBlockColumn As ListColumn
    Dim ParentBlockName As String
    Dim ChildBlockName As String

    ' Set the worksheet and table
    Set ws = ThisWorkbook.Sheets(blocksSheet)
    Set BlocksTable = ws.ListObjects("BlocksTable")

    ' Get the selected cell
    On Error Resume Next
    Set selectedRow = Application.Selection
    On Error GoTo 0

    If selectedRow Is Nothing Then Exit Sub

    ' Get the "Vendor Block ID" and "Labcorp Block ID" columns
    On Error Resume Next
    Set parentBlockColumn = BlocksTable.ListColumns(ParentBlockColName)
    Set childBlockColumn = BlocksTable.ListColumns(ChildBlockColName)
    On Error GoTo 0

    If parentBlockColumn Is Nothing Or childBlockColumn Is Nothing Then Exit Sub

    ' Retrieve the "Vendor Block ID" and "Labcorp Block ID" from the selected row
    ParentBlockName = selectedRow.EntireRow.Cells(1, parentBlockColumn.Index).value
    ChildBlockName = selectedRow.EntireRow.Cells(1, childBlockColumn.Index).value

    ' Open the ResultForm and prefill TextBox1
    If Trim(ChildBlockName) = "" Then
        ' If Labcorp Block ID is empty, use Vendor Block ID
        If ParentBlockName <> "" Then
            ResultForm.TextBox1.value = ParentBlockName
            ResultForm.OptionButtonParent.value = True
            ResultForm.Show
        Else
            MsgBox "No valid 'Vendor Block ID' found for this row.", vbExclamation
        End If
    Else
        ' If Labcorp Block ID is not empty, use Labcorp Block ID
        ResultForm.TextBox1.value = ChildBlockName
        ResultForm.OptionButtonChild.value = True
        ResultForm.Show
    End If
End Sub


Sub OpenReviewsForm()
    Dim ws As Worksheet
    Dim BlocksTable As ListObject
    Dim selectedRow As Range
    Dim parentBlockColumn As ListColumn
    Dim childBlockColumn As ListColumn
    Dim BlockName As String

    ' Set the worksheet and table
    Set ws = ThisWorkbook.Sheets(blocksSheet)
    Set BlocksTable = ws.ListObjects("BlocksTable")

    ' Get the selected cell
    On Error Resume Next
    Set selectedRow = Application.Selection
    On Error GoTo 0

    If selectedRow Is Nothing Then Exit Sub

    ' Get the "Vendor Block ID" and "Labcorp Block ID" columns
    On Error Resume Next
    Set parentBlockColumn = BlocksTable.ListColumns(ParentBlockColName)
    Set childBlockColumn = BlocksTable.ListColumns(ChildBlockColName)
    On Error GoTo 0

    If parentBlockColumn Is Nothing Or childBlockColumn Is Nothing Then Exit Sub

    ' Retrieve the block name and set the appropriate option button
    BlockName = selectedRow.EntireRow.Cells(1, childBlockColumn.Index).value
    If Trim(BlockName) = "" Then
        BlockName = selectedRow.EntireRow.Cells(1, parentBlockColumn.Index).value
        ReviewForm.OptionButtonParent.value = True
    Else
        ReviewForm.OptionButtonChild.value = True
    End If

    ' Open the ReviewForm and prefill TextBox1
    If BlockName <> "" Then
        ReviewForm.TextBox1.value = BlockName
        ReviewForm.Show
    Else
        MsgBox "No valid block name found for this row.", vbExclamation
    End If
End Sub



Sub OpenNewChildsBlockForm()
    Dim ws As Worksheet
    Dim BlocksTable As ListObject
    Dim selectedRow As Range
    Dim childBlockColumn As ListColumn
    Dim ParentBlockName As String

    ' Set the worksheet and table
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(blocksSheet)
    If ws Is Nothing Then
        MsgBox "Worksheet 'BlocksData' not found.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    On Error Resume Next
    Set BlocksTable = ws.ListObjects("BlocksTable")
    If BlocksTable Is Nothing Then
        MsgBox "Table 'BlocksTable' not found in the worksheet.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Get the selected cell
    If TypeName(Application.Selection) = "Range" Then
        Set selectedRow = Application.Selection
    Else
        MsgBox "Please select a valid cell in the table.", vbExclamation
        Exit Sub
    End If

    ' Get the "Vendor Block ID" column
    On Error Resume Next
    Set parentBlockColumn = BlocksTable.ListColumns(ChildBlockColName)
    If parentBlockColumn Is Nothing Then
        MsgBox "'Vendor Block ID' column not found in the table.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Retrieve the "Vendor Block ID" from the selected row
    On Error Resume Next
    ParentBlockName = selectedRow.EntireRow.Cells(1, parentBlockColumn.Index).value
    On Error GoTo 0

    If ParentBlockName = "" Then
        MsgBox "No valid 'Vendor Block ID' found for this row.", vbExclamation
        Exit Sub
    End If

    ' Open the NewChildBlockForm and prefill TextBox1
    NewChildBlockForm.LabelBlockID = ParentBlockName
    NewChildBlockForm.Show
End Sub


