VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateViewForm 
   Caption         =   "Create a View"
   ClientHeight    =   6060
   ClientLeft      =   105
   ClientTop       =   455
   ClientWidth     =   5712
   OleObjectBlob   =   "CreateViewForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateViewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Populate the ListBox with the column names from the BlocksTable
    PopulateListBoxWithColumns
End Sub

Private Sub PopulateListBoxWithColumns()
    Dim ws As Worksheet
    Dim blockTable As ListObject
    Dim column As ListColumn
    
    ' Set the worksheet and table containing the columns
    Set ws = ThisWorkbook.Sheets(blocksSheet)
    Set blockTable = ws.ListObjects("BlocksTable")
    
    ' Clear the ListBox before populating
    Me.ListBox1.Clear
    
    ' Add the column names to the ListBox
    For Each column In blockTable.ListColumns
        Me.ListBox1.AddItem column.Name
    Next column
End Sub

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim viewTable As ListObject
    Dim NewRow As ListRow
    Dim selectedColumns As String
    Dim i As Long
    
    ' Check if the TextBox is empty
    If Me.TextBox1.value = "" Then
        MsgBox "Please enter a value in the View field.", vbExclamation
        Exit Sub
    End If
    
    ' Gather the selected columns from the ListBox
    selectedColumns = ""
    For i = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) Then
            If selectedColumns = "" Then
                selectedColumns = Me.ListBox1.List(i)
            Else
                selectedColumns = selectedColumns & "|" & Me.ListBox1.List(i)
            End If
        End If
    Next i
    
    ' Check if at least one column was selected
    If selectedColumns = "" Then
        MsgBox "Please select at least one column.", vbExclamation
        Exit Sub
    End If
    
    ' Set the worksheet and table where the data will be stored
    Set ws = ThisWorkbook.Sheets(settingsSheet)
    Set viewTable = ws.ListObjects("ViewTable")
    
    ' Add a new row to the table
    Set NewRow = viewTable.ListRows.Add
    With NewRow.Range
        .Cells(1, viewTable.ListColumns("View").Index).value = Me.TextBox1.value
        .Cells(1, viewTable.ListColumns("Columns").Index).value = selectedColumns
    End With
    
    ' Success message
    MsgBox "Data successfully added to ViewTable!" & vbNewLine & _
           "View: " & Me.TextBox1.value & vbNewLine & _
           "Columns: " & selectedColumns, vbInformation
    
    ' Clear the form fields for the next entry
    Me.TextBox1.value = ""
    Me.ListBox1.MultiSelect = fmMultiSelectSingle
    Me.ListBox1.value = ""
    Me.ListBox1.MultiSelect = fmMultiSelectMulti
End Sub


