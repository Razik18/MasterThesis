VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LinkForm 
   Caption         =   "Add image to block"
   ClientHeight    =   3552
   ClientLeft      =   -133
   ClientTop       =   -469
   ClientWidth     =   7679
   OleObjectBlob   =   "LinkForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LinkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




' Function to find the row for a given value in a specified column
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
Private Sub CommandButtonAddLink_Click()
    Dim scanSheet As Worksheet
    Dim scanTable As ListObject
    Dim blocksSheet As Worksheet
    Dim BlocksTable As ListObject
    Dim blockID As String
    Dim linkFromTextBox As String
    Dim blockRow As Long
    Dim parentBlockNameColumn As Long
    Dim childBlockNameColumn As Long
    Dim isParent As Boolean

    ' Set scan sheet and table
    Set scanSheet = ThisWorkbook.Sheets("Scan")
    Set blocksSheet = ThisWorkbook.Sheets(blocksSheet)
    On Error Resume Next
    Set scanTable = scanSheet.ListObjects("ScanTable")
    Set BlocksTable = blocksSheet.ListObjects("BlocksTable")
    On Error GoTo 0

    If scanTable Is Nothing Or BlocksTable Is Nothing Then
        MsgBox "ScanTable or BlocksTable not found.", vbExclamation
        Exit Sub
    End If

    ' Retrieve Block ID from the form
    blockID = Me.TextBox1.Text
    If blockID = "" Then
        MsgBox "Please enter a Block ID.", vbExclamation
        Exit Sub
    End If

    ' Retrieve the link from TextBox2
    linkFromTextBox = Me.TextBox2.Text
    If linkFromTextBox = "" Then
        MsgBox "Please enter a link.", vbExclamation
        Exit Sub
    End If

    ' Determine if the block is a parent or a child based on the option buttons
    If Me.OptionButtonParent.value Then
        isParent = True
    ElseIf Me.OptionButtonChild.value Then
        isParent = False
    Else
        MsgBox "Please select whether the block is a Parent or a Child.", vbExclamation
        Exit Sub
    End If

    ' Add the data to the ScanTable
    Dim NewRow As ListRow
    Set NewRow = scanTable.ListRows.Add
    NewRow.Range.Cells(1, 1).value = blockID
    NewRow.Range.Cells(1, 2).value = "Link"
    NewRow.Range.Cells(1, 3).value = linkFromTextBox

    ' Create a hyperlink to the URL in the ScanTable
    On Error Resume Next
    scanSheet.Hyperlinks.Add _
        Anchor:=NewRow.Range.Cells(1, 3), _
        Address:=linkFromTextBox, _
        TextToDisplay:="Open Link"
    On Error GoTo 0

    ' Find the row in BlocksTable where Block ID exists
    blockRow = FindBlockRow(blocksSheet, "BlocksTable", blockID)
    If blockRow > 0 Then
        If isParent Then
            ' Apply the hyperlink to the Vendor Block ID column
            parentBlockNameColumn = BlocksTable.ListColumns(ParentBlockColName).Index
            BlocksTable.DataBodyRange.Cells(blockRow, parentBlockNameColumn).Hyperlinks.Add _
                Anchor:=BlocksTable.DataBodyRange.Cells(blockRow, parentBlockNameColumn), _
                Address:=linkFromTextBox, _
                TextToDisplay:=blockID
        Else
            ' Apply the hyperlink to the Labcorp Block ID column
            childBlockNameColumn = BlocksTable.ListColumns(ChildBlockColName).Index
            BlocksTable.DataBodyRange.Cells(blockRow, childBlockNameColumn).Hyperlinks.Add _
                Anchor:=BlocksTable.DataBodyRange.Cells(blockRow, childBlockNameColumn), _
                Address:=linkFromTextBox, _
                TextToDisplay:=blockID
        End If
    Else
        MsgBox "Block ID not found in BlocksTable.", vbExclamation
    End If

    MsgBox "Link has been successfully added!", vbInformation
End Sub



Private Function FindBlockRow(ws As Worksheet, tableName As String, blockID As String) As Long
    Dim BlocksTable As ListObject
    Dim dataRange As Range
    Dim parentBlockColumn As ListColumn
    Dim childBlockColumn As ListColumn
    Dim Cell As Range

    On Error Resume Next
    Set BlocksTable = ws.ListObjects(tableName)
    On Error GoTo 0

    If BlocksTable Is Nothing Then Exit Function

    Set parentBlockColumn = BlocksTable.ListColumns(ParentBlockColName)
    Set childBlockColumn = BlocksTable.ListColumns(ChildBlockColName)

    If parentBlockColumn Is Nothing Or childBlockColumn Is Nothing Then Exit Function

    Set dataRange = Union(parentBlockColumn.DataBodyRange, childBlockColumn.DataBodyRange)

    For Each Cell In dataRange
        If Cell.value = blockID Then
            FindBlockRow = Cell.row - BlocksTable.HeaderRowRange.row
            Exit Function
        End If
    Next Cell

    FindBlockRow = 0
End Function





