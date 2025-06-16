Attribute VB_Name = "CreateNewScoringModule"
Sub OpenAddScoringForm()
    Dim MyUserForm As AddScoringForm
    Dim rangeMarkers As Range

    ' Initialiser les variables globales
    SetVariables

    ' Définir les cellules où se trouvent les Markers
    Set rangeMarkers = SettingWS.ListObjects(MarkersTableName).DataBodyRange

    ' Créer un nouveau UserForm (pas encore affiché)
    Set MyUserForm = New AddScoringForm

    ' Remplir la ListBox avec les Markers
    MyUserForm.ListBox1.RowSource = rangeMarkers.Address(External:=True)
    MyUserForm.ListBox1.value = ""

    ' Afficher le UserForm
    MyUserForm.Show 0
End Sub

