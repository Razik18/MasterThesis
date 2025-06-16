Attribute VB_Name = "DeleteScoring"
Sub OpenDeleteScoringForm()
    Dim MyUserForm As DeleteScoringForm
    Dim rangeMarkers As Range

    ' Initialiser les variables globales
    SetVariables

    ' D�finir les cellules o� se trouvent les Markers
    Set rangeMarkers = SettingWS.ListObjects(MarkersTableName).DataBodyRange

    ' Cr�er un nouveau UserForm (pas encore affich�)
    Set MyUserForm = New DeleteScoringForm

    ' Remplir la ListBox avec les Markers
    MyUserForm.ListBox1.RowSource = rangeMarkers.Address(External:=True)
    MyUserForm.ListBox1.value = ""

    ' Afficher le UserForm
    MyUserForm.Show 0
End Sub
