Attribute VB_Name = "NewChildBlockModule"
Sub OpenNewChildBlockForm()
    Dim MyUserForm As NewChildBlockForm
    
    ' Initialiser les variables global :)
    SetVariables
    
    ' Créer un nouveau UserForm (pas encore afficher)
    Set MyUserForm = New NewChildBlockForm
    
    ' Afficher le UserForm
    MyUserForm.Show 0
End Sub
Function NewChildBlock(ParentBlockName As String, NumChildStr As String, KeepParent As Boolean, Marker As String)
    Dim NumChild As Integer
    Dim lastRow As Integer
    Dim ParentBlockRow As Integer
    Dim ParentBlockCol As Integer
    Dim LabcorpBlockID As String
    Dim result As Variant

    Dim RangeParent As Range
    Dim RangeChild As Range

    ' Initialize global variables
    SetVariables

    ' Check if the Vendor Block ID is empty
    If ParentBlockName = "" Then
        MsgBox "The Block ID is empty."
        Exit Function
    End If

    ' Check if NumChild is a number and convert it
    If IsNumeric(NumChildStr) Then
        NumChild = CInt(NumChildStr)
    Else
        MsgBox "Enter a valid number for the number of children."
        Exit Function
    End If

    ' Check if NumChild is greater than 0
    If NumChild = 0 Then
        MsgBox "Enter a number greater than 0 for the number of children."
        Exit Function
    End If

    ' Check if the ParentBlock exists in the main worksheet
    ParentBlockRow = Get_ParentBlock_Rows(BlocksWS, BlocksTableName, ChildBlockColName, ParentBlockName)
    If ParentBlockRow = -1 Then
        MsgBox "The specified Vendor Block ID is not found: " & ParentBlockName
        Exit Function
    End If

    ' Retrieve the Labcorp Block ID for the parent
    Dim LabcorpBlockIDCol As Integer
    LabcorpBlockIDCol = BlocksWS.ListObjects(BlocksTableName).HeaderRowRange.Cells.Find(ChildBlockColName).column
    LabcorpBlockID = BlocksWS.Cells(ParentBlockRow, LabcorpBlockIDCol).value

    If LabcorpBlockID = "" Then
        MsgBox "The Labcorp Block ID for the parent block is empty."
        Exit Function
    End If

    ' Determine the range of the last row and columns
    lastRow = GetLastRowTable(BlocksWS, BlocksTableName)
    lastColumn = GetLastColTable(BlocksWS, BlocksTableName)
    firstColumn = BlocksWS.ListObjects(BlocksTableName).Range.column

    ' Determine the range for the Parent Block
    Set RangeParent = BlocksWS.Range(BlocksWS.Cells(ParentBlockRow, firstColumn), BlocksWS.Cells(ParentBlockRow, lastColumn))
    RangeParent.Select

    ' Retrieve the columns for child and block state
    Dim ChildBlockCol As Integer, BlockStateCol As Integer
    ChildBlockCol = BlocksWS.ListObjects(BlocksTableName).HeaderRowRange.Cells.Find(ChildBlockColName).column
    BlockStateCol = BlocksWS.ListObjects(BlocksTableName).HeaderRowRange.Cells.Find(BlockStateColName).column

    ' Retrieve the columns for score and marker used
    Dim ScoreCol As Integer, markerUsedCol As Integer, HECol As Integer
    ScoreCol = BlocksWS.ListObjects(BlocksTableName).HeaderRowRange.Cells.Find(ScoreColName).column
    markerUsedCol = BlocksWS.ListObjects(BlocksTableName).HeaderRowRange.Cells.Find(MarkerUsedColName).column
    HECol = BlocksWS.ListObjects(BlocksTableName).HeaderRowRange.Cells.Find(HEColName).column

    Dim i As Integer

    ' Create as many children as the number specified
    For i = 1 To NumChild
        Dim ChildBlockName As String
        Dim suffix As String
        Dim maxSuffix As Integer
        maxSuffix = 0

        ' Generate child block name based on the ending of the parent block name
        If IsNumeric(Right(LabcorpBlockID, 1)) Then
            ' Ends with a number, append a letter
            Dim charSuffix As String
            charSuffix = Chr(Asc("A") + maxSuffix)
            Do
                ChildBlockName = LabcorpBlockID & charSuffix
                maxSuffix = maxSuffix + 1
                charSuffix = Chr(Asc("A") + maxSuffix)
            Loop While Not IsUniqueBlockName(ChildBlockName)
        Else
            ' Ends with a letter, append a number
            maxSuffix = 1
            Do
                ChildBlockName = LabcorpBlockID & "." & maxSuffix
                maxSuffix = maxSuffix + 1
            Loop While Not IsUniqueBlockName(ChildBlockName)
        End If

        ' Determine the range for the new Child block
        BlocksWS.Range(BlocksWS.Cells(lastRow + i, firstColumn), BlocksWS.Cells(lastRow + i, lastColumn)).Select
        Set RangeChild = BlocksWS.Range(BlocksWS.Cells(lastRow + i, firstColumn), BlocksWS.Cells(lastRow + i, lastColumn))

        ' Copy the Parent block data to the Child
        RangeParent.Copy RangeChild

        ' Assign the new Child block name
        BlocksWS.Cells(lastRow + i, ChildBlockCol).value = ChildBlockName

        ' Add the marker with "(in Review)" to the "Marker Used" column if the marker is not empty
        If Trim(Marker) <> "" Then
            BlocksWS.Cells(lastRow + i, markerUsedCol).value = Marker & " (in Review)"
        Else
            ' Clear the cell if the marker is empty
            BlocksWS.Cells(lastRow + i, markerUsedCol).ClearContents
        End If

        ' Clear any existing values for the child
        BlocksWS.Cells(lastRow + i, ScoreCol).value = ""
        BlocksWS.Cells(lastRow + i, HECol).value = ""

        ' Change the Block State to "Stock"
        BlocksWS.Cells(lastRow + i, BlockStateCol).value = StockChildText

        anatomicSiteColumn = BlocksWS.ListObjects("BlocksTable").ListColumns(AnatomicSiteColName).Range.column
        anatomicSite = BlocksWS.Cells(lastRow + i, anatomicSiteColumn).value

        vendorNameColumn = BlocksWS.ListObjects("BlocksTable").ListColumns(ParentBlockColName).Range.column
        vendorName = BlocksWS.Cells(lastRow + i, vendorNameColumn).value

        ' Create folder and add hyperlink
        FolderPath = MainFolderPath + "\" + anatomicSite + "\" + "\" + vendorName + "\" + ChildBlockName
        If Dir(FolderPath, vbDirectory) = "" Then
            MkDir FolderPath
        End If

        ' Construct the hyperlink URL
        ParentBlockHyperlink = "https://labcorp.concentriq.proscia.com/dashboard/search?facets=[{%22field%22:{%22text%22:%22Image+name%22,%22searchType%22:%22name%22,%22facetGroup%22:%22image%22,%22resourceType%22:%22image%22},%22value%22:%22" & ChildBlockName & "%22}"

        ' Identify the cell where Parent Block Name is stored
        Set ChildBlockCell = BlocksWS.Cells(lastRow + i, ChildBlockCol)

        ' Add the hyperlink
        BlocksWS.Hyperlinks.Add Anchor:=ChildBlockCell, Address:=ParentBlockHyperlink, TextToDisplay:=ChildBlockName
    Next i

    ' If "Keep Parent" is not checked, move the parent to the "Exhausted" state
    If Not KeepParent Then
        BlocksWS.Cells(ParentBlockRow, BlockStateCol).value = ExhaustedBlockText
    Else
        BlocksWS.Cells(ParentBlockRow, BlockStateCol).value = StockBlockText
    End If

    ' Move the selection to the newly added child block
    BlocksWS.Cells(lastRow + i - 1, 1).Select

End Function

Private Function IsUniqueBlockName(BlockName As String) As Boolean
    Dim checkRow As Integer
    For checkRow = 2 To GetLastRowTable(BlocksWS, BlocksTableName)
        If BlocksWS.Cells(checkRow, BlocksWS.ListObjects(BlocksTableName).HeaderRowRange.Cells.Find(ChildBlockColName).column).value = BlockName Then
            IsUniqueBlockName = False
            Exit Function
        End If
    Next checkRow
    IsUniqueBlockName = True
End Function





