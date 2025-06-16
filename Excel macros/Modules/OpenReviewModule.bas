Attribute VB_Name = "OpenReviewModule"
Sub OpenReviewForm()
    Dim MyUserForm As ReviewForm
    
    ' Initialiser les variables global
    SetVariables
    
    ' Définir les cellules où se trouve les Markers
    Set rangeMarkers = SettingWS.ListObjects(MarkersTableName).DataBodyRange
    
    ' Créer un nouveau UserForm (pas encore afficher)
    Set MyUserForm = New ReviewForm
    

    ' Afficher le UserForm
    MyUserForm.Show 0
    
End Sub

Function MoveReviewBlock(BlockName As String, ListMarkers As Collection, blockType As String)
    Dim blockRow As Long
    Dim BlockColName As String
    Dim BlockStateCol As Long
    Dim MarkerReviewCol As Long
    Dim MarkerText As String
    Dim markerName As Variant

    ' Initialize global variables
    SetVariables

    ' Check if the block name is empty
    If BlockName = "" Then
        MsgBox "The block name cannot be empty."
        Exit Function
    End If

    ' Determine which column to use based on block type
    If blockType = "Parent" Then
        BlockColName = ParentBlockColName
    ElseIf blockType = "Child" Then
        BlockColName = ChildBlockColName
    Else
        MsgBox "Invalid block type specified."
        Exit Function
    End If

    ' Find the row corresponding to the block name
    blockRow = Get_ParentBlock_Rows(BlocksWS, BlocksTableName, BlockColName, BlockName)
    If blockRow = -1 Then
        MsgBox "The specified block name was not found: " & BlockName
        Exit Function
    End If

    ' Get the Block State column
    BlockStateCol = GetColTable(BlocksWS, BlocksTableName, BlockStateColName)

    ' Change the Block State to InReviewParent or InReviewChild based on the block type
    If blockType = "Parent" Then
        BlocksWS.Cells(blockRow, BlockStateCol).value = ParentInReviewText
    ElseIf blockType = "Child" Then
        BlocksWS.Cells(blockRow, BlockStateCol).value = ChildInReviewText
    End If

    ' Get the Marker Review column
    MarkerReviewCol = BlocksWS.ListObjects(BlocksTableName).HeaderRowRange.Cells.Find(MarkerUsedColName).column

    ' Retrieve existing markers
    MarkerText = BlocksWS.Cells(blockRow, MarkerReviewCol).value

    ' Add the markers to the text
    For Each markerName In ListMarkers
        If MarkerText = "" Then
            MarkerText = markerName & "(in Review)"
        Else
            ' Add marker only if it doesn't already exist
            If InStr(MarkerText, markerName) = 0 Then
                MarkerText = MarkerText & "|" & markerName & "(in Review)"
            End If
        End If
    Next markerName

    ' Update the Marker Review column
    BlocksWS.Cells(blockRow, MarkerReviewCol).value = MarkerText
End Function

