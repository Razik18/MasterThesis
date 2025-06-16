Attribute VB_Name = "Functions"
Public BlocksWS As Worksheet
Public SettingWS As Worksheet
Public TmaWS As Worksheet


Public SelectedRowIndex As Long


Public AnatomicSiteColName As String
Public BlockStateColName As String
Public ParentBlockColName As String
Public ChildBlockColName As String
Public ProjectsColName As String
Public ScoreColName As String
Public MarkerUsedColName As String
Public HEColName As String
Public TMABlockColName As String
Public mIFPannelColName As String

Public AnatomicSiteTableName As String
Public ProjectsTableName As String
Public TumorTypeTableName As String
Public MarkersTableName As String
Public CloseDateColName As String
Public OpenDateColName As String
Public SearchedMarkerTableName As String
Public TmaTableName As String
Public AcronymTable As String
Public Acronymfirstcolumn As String
Public VendorInfoColName As String
Public TumorTypeColName As String
Public VendorColName As String
Public VendorBiomarkerColName As String
Public ProcessColName As String
Public SiteColName As String
Public FixationtimeColName As String
Public FixativeColName As String
Public SampleTypeColName As String
Public CreationDateColName As String
Public settingsSheet As String
Public blocksSheet As String
Public MultiplexSheet As String

Public TMAParentColName As String


Public StockTMAText As String
Public TMAInReviewText As String
Public ValidatedTMAText As String
Public TMANonValidatedText As String

Public BlocksTableName As String

Public StockBlockText As String
Public ParentInReviewText As String
Public CharacterizedBlockText As String
Public StockChildText As String
Public ChildInReviewText As String
Public ValidatedChildText As String
Public InUseBlockText As String
Public ExhaustedBlockText As String

Public MainFolderPath As String

Public password As String

Function SetVariables()
    ' Les différents paramètre à changer si des changement sont fait dans le worsheet
    
    
    ' Le nom du table pour les blocks
    AnatomicSiteTableName = "AnatomicSiteTable"
    AcronymTable = "TumorTypeTable"
    Acronymfirstcolumn = "Tumor Type"
    
    ' Les noms des colonnes dans le table
    BlockStateColName = "Block State"
    AnatomicSiteColName = "Anatomic Site"
    ParentBlockColName = "Vendor Block ID"
    ChildBlockColName = "Labcorp Block ID"
    ProjectsColName = "Projet Associeted"
    CloseDateColName = "date de fermeture"
    OpenDateColName = "date d'ouverture"
    ScoreColName = "Result" 'new
    TMABlockColName = "TMA Block Name (MMJJAA)"
    mIFPannelColName = "mIF pannel" ' new
    MarkerUsedColName = "Marker Used" ' new
    HEColName = "H&E State"
    TMAParentColName = "Parent Block Name"
    VendorInfoColName = "Additional Informations from Vendor"
    TumorTypeColName = "Clinical Indication"
    VendorColName = "Vendor"
    VendorBiomarkerColName = "Biomarker Characterisation (Vendor)"
    ProcessColName = "Process"
    SiteColName = "Site"
    FixationtimeColName = "Fixation time"
    FixativeColName = "Fixative"
    SampleTypeColName = "Sample Type"
    CreationDateColName = "Creation Date"
    
    ' Les noms des tables dans settings worksheet
    BlocksTableName = "BlocksTable"
    ProjectsTableName = "ProjectsTable"
    TumorTypeTableName = "TumorTypeTable"
    MarkersTableName = "MarkersTable"
    SearchedMarkerTableName = "SearchMarkerTable"
    TmaTableName = "TMATable"
    
    ' Les textes à mettre pour les block state
    StockBlockText = "1-StockParent"
    ParentInReviewText = "2-InReviewParent" ' new
    CharacterizedBlockText = "3-CharacterizedParent" ' new
    StockChildText = "4-StockChild" ' new
    ChildInReviewText = "5-InReviewChild" ' new
    ValidatedChildText = "6-ValidatedChild" ' new
    InUseBlockText = "7-In Use"
    ExhaustedBlockText = "8-Exhausted"
    StockTMAText = "4-StockTMA" ' new
    TMAInReviewText = "5-InReviewTMA" ' new
    ValidatedTMAText = "6-ValidatedTMA" ' new
    TMANonValidatedText = "TMANonValidated" ' new
    
    'Sheets name
    settingsSheet = "Settings"
    blocksSheet = "BlocksData"
    MultiplexSheet = "TMAData"
    
    Set BlocksWS = ThisWorkbook.Sheets(blocksSheet)
    Set SettingWS = ThisWorkbook.Sheets(settingsSheet)
    Set TmaWS = ThisWorkbook.Sheets(MultiplexSheet)
    ' le folder parent où se trouve les block charactérisation
    MainFolderPath = "\\gvafps05\Lab_AllSites\HISTOLOGY\GVA- Tech\QC BLOCKS\Test"
    
End Function

Function Get_Rows_Table(ws As Worksheet, tableName As String, colName As String, SearchedValue As Variant)
    ' Prendre la colonne dans le tableau qui a le nom du colonne cherché
    column = ws.ListObjects(tableName).HeaderRowRange.Cells.Find(colName).column
    
    ' Chercher la ligne qui contient la valeur dans la colonne
    If Not ws.Columns(column).Find(SearchedValue, LookIn:=xlValues, LookAt:=xlWhole) Is Nothing Then
        Get_Rows_Table = ws.Columns(column).Find(SearchedValue, LookIn:=xlValues, LookAt:=xlWhole).row
    Else
        ' Retouner une valeurs négative si on trouve pas la valeur
        Get_Rows_Table = -1
    End If
End Function

Function Get_ParentBlock_Rows(ws As Worksheet, tableName As String, ColumnName As String, ParentBlockName As String) As Long
    Dim table As ListObject
    Dim SearchColumn As Range
    Dim Cell As Range

    ' Set the table and search column
    Set table = ws.ListObjects(tableName)
    Set SearchColumn = table.ListColumns(ColumnName).DataBodyRange

    ' Find the Vendor Block ID in the search column
    Set Cell = SearchColumn.Find(ParentBlockName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not Cell Is Nothing Then
        Get_ParentBlock_Rows = Cell.row
    Else
        Get_ParentBlock_Rows = -1 ' Return -1 if not found
    End If
End Function


Function GetFirstRowTable(ws As Worksheet, tableName As String)
    GetFirstRowTable = ws.ListObjects(tableName).Range.Rows(1).row
End Function

Function GetLastRow(ws As Worksheet)
    ' Récuperer la dernière ligne de worksheet
    GetLastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
End Function

Function GetLastRowTable(ws As Worksheet, tableName As String)
    ' Récuperer la dernière ligne de Table
    GetLastRowTable = ws.ListObjects(tableName).Range.row + ws.ListObjects(tableName).Range.Rows.Count - 1
End Function

Function GetLastColTable(ws As Worksheet, tableName As String)
    ' Récuperer la dernière colonne de Table
    GetLastColTable = ws.ListObjects(tableName).Range.column + ws.ListObjects(tableName).Range.Columns.Count - 1
End Function

Function GetColTable(ws As Worksheet, tableName As String, colName As String)
    ' Prendre la colonne dans le tableau qui a le nom du colonne cherché
    GetColTable = ws.ListObjects(tableName).HeaderRowRange.Cells.Find(colName).column
End Function

Function SetHyperlinkToCell(CellRange As Range, FolderPath As String, CellValue As String)
    ' Mettre le hyperlink
    CellRange.Hyperlinks.Add Anchor:=CellRange, Address:=FolderPath, SubAddress:="", ScreenTip:="it is a Hyperlink", TextToDisplay:=CellValue
End Function

Function GetDatemmddyy()
    ' Récupérer la date de jour en mm dd yy
    Dim FormattedDate As String
    
    ' Récupérer la date de jour
    Dim currentDate As Date
    currentDate = Date
    
    ' Formatter la date en mm dd yy
    FormattedDate = Format(currentDate, "mmddyy")
    
    ' Mettre en majuscule les date
    FormattedDate = UCase(FormattedDate)
    GetDatemmddyy = FormattedDate
End Function

Function GetDateddmmmyy()
    ' Récupérer la date de jour en mmm dd yy
    Dim FormattedDate As String
    
    ' Récupérer la date de jour
    Dim currentDate As Date
    currentDate = Date
    
    ' Formatter la date en mmm dd yy
    FormattedDate = Format(currentDate, "ddmmmyy")
    
    ' Mettre en majuscule les date
    FormattedDate = UCase(FormattedDate)
    GetDateddmmmyy = FormattedDate
End Function

Function GetDateWithLetter(Num As Integer)
    ' Ajouter une lettre dependement du nombre Num
    GetDateWithLetter = GetDatemmddyy() + Chr(Num + 64)
End Function

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Function GetMarkerText(markerName As String, MarkerText As String)
    Dim MarkerList() As String
    MarkerList = Split(MarkerText, "|")
    
    If MarkerText = "" Then
        ' Si il n'y a aucun marqueur mettre le premier marqueur
        MarkerText = markerName
    Else
        ' Si il existe déjà des marqueur, ajouter le marqueur si est seulement s'il n'existe pas déjà parmis les marqueurs
        If Not IsInArray(markerName, MarkerList) Then
            ' Ajouter le marqueur au texte déjà mis
            MarkerText = MarkerText + "|" + markerName
            MarkerList = Split(MarkerText, "|")
        End If
    End If
    
    GetMarkerText = MarkerText
    
End Function

Function ClearUserFormValidate(MyUserForm As Object)
    For Each cont In MyUserForm.Controls
        If InStr(cont.Name, "CheckBoxPos_") = 1 Then
            MyUserForm.Controls.Remove cont.Name
            GoTo EndLoop
        End If
        
        If InStr(cont.Name, "CheckBoxNeg_") = 1 Then
            MyUserForm.Controls.Remove cont.Name
            GoTo EndLoop
        End If
        
        If InStr(cont.Name, "CheckBoxRejected_") = 1 Then
            MyUserForm.Controls.Remove cont.Name
            GoTo EndLoop
        End If
        
        If InStr(cont.Name, "Label_") = 1 Then
            MyUserForm.Controls.Remove cont.Name
            GoTo EndLoop
        End If
        
EndLoop:
    Next cont
End Function
Function RemoveOneElementText(TextToRemove As String, Text As String)
    ''' Return TextList without TextUnit. Text List is separated by "|"
    Dim TextList() As String
    TextList = Split(Text, "|")
    NewText = ""
    
    For Each TextListUnit In TextList
        If CStr(TextListUnit) <> CStr(TextToRemove) Then
            NewText = GetMarkerText(CStr(TextListUnit), CStr(NewText))
        End If
    Next
    
    RemoveOneElementText = NewText
    
End Function

