Attribute VB_Name = "OpenResultModule"
Sub OpenResultForm()
    Dim MyUserForm As ResultForm
    
    ' Initialiser les variables global
    SetVariables
    
    ' Créer un nouveau UserForm (pas encore afficher)
    Set MyUserForm = New ResultForm
    
    ' Afficher le UserForm
    MyUserForm.Show 0
    
End Sub
