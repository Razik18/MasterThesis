Attribute VB_Name = "PreselectedViewsModule"
Const SheetPassword As String = "qc"
Const SettingsPassword As String = "settingsqc"


Sub ApplyView_QCIHC()
    Dim processFilter As New Collection
    SetVariables
    processFilter.Add "QCIHC" ' Set the process filter to "QCIHC"
    If UnprotectSheetWithPassword() Then
        ApplyViewAndFilter "QCIHC", processFilter
    End If
End Sub

Sub ApplyView_Validation()
    Dim processFilter As New Collection
    SetVariables
    processFilter.Add "Validation" ' Set the process filter to "Validation"
    If UnprotectSheetWithPassword() Then
        ApplyViewAndFilter "Validation", processFilter
    End If
End Sub

Sub ApplyView_Pathologist()
    SetVariables
    If UnprotectSheetWithPassword() Then
        ApplyViewAndDisplay "Complete"
        ClearFiltersInBlocksTable
    End If
End Sub

Sub ApplyView_Complete()
    SetVariables
    If UnprotectSheetWithPassword() Then
        ApplyViewAndDisplay "Complete"
        ClearFiltersInBlocksTable
    End If
End Sub
Sub ApplyView_External()
    SetVariables
    Sheets(blocksSheet).Activate
End Sub

Sub Gotosettings()
    SetVariables
    If UnprotectSettingsWithPassword() Then
        ' Navigate to the "Settings" sheet
        Dim settingSheet As Worksheet
        On Error Resume Next
        Set settingSheet = ThisWorkbook.Sheets(settingsSheet)
        On Error GoTo 0
        

        settingSheet.Activate

    End If
End Sub


Function UnprotectSheetWithPassword() As Boolean
    On Error GoTo HandleError
    Dim userPassword As String
    Dim ws As Worksheet
    
    ' Reference to BlocksData sheet
    Set ws = ThisWorkbook.Sheets(blocksSheet)
    
    ' Prompt the user for a password
    userPassword = InputBox("Enter the password to unlock the BlocksData sheet:", "Password Required")
    
    ' Check the entered password
    If userPassword = SheetPassword Then
        ws.Unprotect password:=SheetPassword
        UnprotectSheetWithPassword = True
    Else
        MsgBox "Incorrect password. The sheet remains protected.", vbExclamation
        UnprotectSheetWithPassword = False
    End If
    Exit Function

HandleError:
    MsgBox "An error occurred while attempting to unprotect the sheet.", vbCritical
    UnprotectSheetWithPassword = False
End Function
Function UnprotectSettingsWithPassword() As Boolean
    On Error GoTo HandleError
    Dim userPassword As String
    Dim ws As Worksheet
    
    ' Reference to the Settings sheet
    Set ws = ThisWorkbook.Sheets(settingsSheet)
    
    ' Prompt the user for a password
    userPassword = InputBox("Enter the password to unlock the Settings sheet:", "Password Required")
    
    ' Check if the user canceled the input box
    If userPassword = vbNullString Then
        MsgBox "Password entry canceled. The sheet remains protected.", vbInformation
        UnprotectSettingsWithPassword = False
        Exit Function
    End If
    
    ' Validate the entered password
    If userPassword = SettingsPassword Then
        ws.Unprotect password:=SettingsPassword
        UnprotectSettingsWithPassword = True
    Else
        MsgBox "Incorrect password. The sheet remains protected.", vbExclamation
        UnprotectSettingsWithPassword = False
    End If
    Exit Function

HandleError:
    MsgBox "An error occurred while attempting to unprotect the sheet. " & _
           "Please ensure the Settings sheet exists and try again.", vbCritical
    UnprotectSettingsWithPassword = False
End Function

Sub ApplyViewAndFilter(viewName As String, processFilter As Collection)
    Dim ws As Worksheet
    Dim BlocksTable As ListObject
    Dim viewTable As ListObject
    Dim foundRow As ListRow
    Dim relatedColumns() As String
    Dim columnsString As String
    Dim allColumns As Collection
    Dim visibleColumns As Collection
    Dim col As ListColumn
    Dim i As Long
    
    SetVariables

    ' Set worksheets and tables
    Set ws = ThisWorkbook.Sheets(blocksSheet)
    Set BlocksTable = ws.ListObjects("BlocksTable")
    Set viewTable = ThisWorkbook.Sheets(settingsSheet).ListObjects("ViewTable")

    ' Find the corresponding row in the ViewTable
    Set foundRow = Nothing
    For Each row In viewTable.ListRows
        If row.Range.Cells(1, viewTable.ListColumns("View").Index).value = viewName Then
            Set foundRow = row
            Exit For
        End If
    Next row

    ' Exit if no matching view found
    If foundRow Is Nothing Then
        MsgBox "View not found: " & viewName, vbExclamation
        Exit Sub
    End If

    ' Extract related columns from the found row
    columnsString = foundRow.Range.Cells(1, viewTable.ListColumns("Columns").Index).value
    relatedColumns = Split(columnsString, "|")

    ' Create collections for all columns and visible columns
    Set allColumns = New Collection
    Set visibleColumns = New Collection

    ' Populate allColumns with all BlockTable columns
    For Each col In BlocksTable.ListColumns
        allColumns.Add col
    Next col

    ' Populate visibleColumns with the columns from the selected view
    For i = LBound(relatedColumns) To UBound(relatedColumns)
        For Each col In allColumns
            If col.Name = Trim(relatedColumns(i)) Then
                visibleColumns.Add col
            End If
        Next col
    Next i

    ' Hide all columns in the table
    For Each col In allColumns
        col.Range.EntireColumn.Hidden = True
    Next col

    ' Show only the visible columns for the selected view
    For Each col In visibleColumns
        col.Range.EntireColumn.Hidden = False
    Next col

    ' Apply additional filters for the specific view
    If Not processFilter Is Nothing Then
        ApplyFilter BlocksTable, ProcessColName, processFilter
    End If

    ' Display the BlocksData sheet
    ws.Activate

    
End Sub

Sub ApplyViewAndDisplay(viewName As String)
    ' Reuse ApplyViewAndFilter without filtering
    ApplyViewAndFilter viewName, Nothing
End Sub

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

Sub GoToTMAData()
    If UnprotectBothSheets() Then
        ' Apply the QCIHC view and filter to the BlocksData sheet
        ApplyMultiplexViewAndFilter

        ' Navigate to the "TMAData" sheet
        Dim tmaSheet As Worksheet
        On Error Resume Next
        Set tmaSheet = ThisWorkbook.Sheets(MultiplexSheet)
        On Error GoTo 0

        If Not tmaSheet Is Nothing Then
            tmaSheet.Activate
        Else
            MsgBox "TMAData sheet not found.", vbExclamation
        End If
    End If
End Sub

Sub ApplyMultiplexViewAndFilter()
    Dim processFilter As New Collection
    processFilter.Add "Multiplex" ' Set the process filter to "Multiplex"
    ApplyViewAndFilter "QCIHC", processFilter
End Sub

Function UnprotectBothSheets() As Boolean
    On Error GoTo HandleError
    Dim userPassword As String
    Dim wsBlocks As Worksheet
    Dim wsTMA As Worksheet

    ' Reference to the BlocksData sheet
    Set wsBlocks = ThisWorkbook.Sheets(blocksSheet)

    ' Reference to the TMAData sheet
    Set wsTMA = ThisWorkbook.Sheets(MultiplexSheet)

    ' Prompt the user for a password
    userPassword = InputBox("Enter the password to unlock the sheets:", "Password Required")

    ' Check if the user canceled the input box
    If userPassword = vbNullString Then
        MsgBox "Password entry canceled. The sheets remain protected.", vbInformation
        UnprotectBothSheets = False
        Exit Function
    End If

    ' Validate the entered password and unprotect both sheets
    If userPassword = SheetPassword Then
        wsBlocks.Unprotect password:=SheetPassword
        wsTMA.Unprotect password:=SheetPassword
        UnprotectBothSheets = True
    Else
        MsgBox "Incorrect password. The sheets remain protected.", vbExclamation
        UnprotectBothSheets = False
    End If
    Exit Function

HandleError:
    MsgBox "An error occurred while attempting to unprotect the sheets. " & _
           "Please ensure both sheets exist and try again.", vbCritical
    UnprotectBothSheets = False
End Function


Function UnprotectTMADataWithPassword() As Boolean
    On Error GoTo HandleError
    Dim userPassword As String
    Dim ws As Worksheet

    ' Reference to the TMAData sheet
    Set ws = ThisWorkbook.Sheets(MultiplexSheet)

    ' Prompt the user for a password
    userPassword = InputBox("Enter the password to unlock the TMAData sheet:", "Password Required")

    ' Check if the user canceled the input box
    If userPassword = vbNullString Then
        MsgBox "Password entry canceled. The sheet remains protected.", vbInformation
        UnprotectTMADataWithPassword = False
        Exit Function
    End If

    ' Validate the entered password
    If userPassword = SheetPassword Then
        ws.Unprotect password:=SheetPassword
        UnprotectTMADataWithPassword = True
    Else
        MsgBox "Incorrect password. The sheet remains protected.", vbExclamation
        UnprotectTMADataWithPassword = False
    End If
    Exit Function

HandleError:
    MsgBox "An error occurred while attempting to unprotect the sheet. " & _
           "Please ensure the TMAData sheet exists and try again.", vbCritical
    UnprotectTMADataWithPassword = False
End Function


