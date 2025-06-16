Attribute VB_Name = "ClearFilterModule"
Sub ClearFiltersInBlocksTable()
    Dim blockWs As Worksheet
    Dim blockTable As ListObject
    Dim wasProtected As Boolean
    Const SheetPassword As String = "qc"

    ' Set the worksheet and table
    Set blockWs = ThisWorkbook.Sheets(blocksSheet)
    Set blockTable = blockWs.ListObjects("BlocksTable")

    ' Check if the sheet is protected
    wasProtected = blockWs.ProtectContents

    ' Unprotect the sheet if it was protected
    If wasProtected Then blockWs.Unprotect password:=SheetPassword

    ' Check if the table has any filters applied and clear them
    If blockTable.AutoFilter.FilterMode Then
        blockTable.AutoFilter.ShowAllData
    End If

    ' Re-protect the sheet if it was protected
    If wasProtected Then
        blockWs.Protect password:=SheetPassword, AllowSorting:=True, AllowFiltering:=True
    End If
End Sub

