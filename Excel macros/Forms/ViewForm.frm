VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ViewForm 
   Caption         =   "Select a View"
   ClientHeight    =   1464
   ClientLeft      =   105
   ClientTop       =   455
   ClientWidth     =   4585
   OleObjectBlob   =   "ViewForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ViewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    PopulateViews
End Sub

Private Sub PopulateViews()
    Dim ws As Worksheet
    Dim viewTable As ListObject
    Dim row As ListRow

    Set ws = ThisWorkbook.Sheets(settingsSheet)
    Set viewTable = ws.ListObjects("ViewTable")
     
    Me.ComboBox1.Clear
    
    For Each row In viewTable.ListRows
        Me.ComboBox1.AddItem row.Range.Cells(1, viewTable.ListColumns("View").Index).value
    Next row
End Sub

Private Sub ComboBox1_Change()
    ' Trigger when a view is selected in the ComboBox
    Dim ws As Worksheet
    Dim BlocksTable As ListObject
    Dim selectedView As String
    Dim relatedColumns() As String
    Dim columnsString As String
    Dim viewTable As ListObject
    Dim foundRow As ListRow
    Dim col As ListColumn
    Dim allColumns As Collection
    Dim visibleColumns As Collection
    Dim i As Long
    
    ' Set worksheets and tables
    Set ws = ThisWorkbook.Sheets(blocksSheet)
    Set BlocksTable = ws.ListObjects("BlocksTable")
    Set viewTable = ThisWorkbook.Sheets(settingsSheet).ListObjects("ViewTable")
    
    ' Get selected view from ComboBox
    selectedView = Me.ComboBox1.value
    
    ' Find the corresponding row in the ViewTable
    Set foundRow = Nothing
    For Each row In viewTable.ListRows
        If row.Range.Cells(1, viewTable.ListColumns("View").Index).value = selectedView Then
            Set foundRow = row
            Exit For
        End If
    Next row
    
    ' Exit if no matching view found
    If foundRow Is Nothing Then
        MsgBox "View not found: " & selectedView, vbExclamation
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
End Sub

