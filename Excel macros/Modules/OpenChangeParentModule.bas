Attribute VB_Name = "OpenChangeParentModule"
Sub OpenChangeParentForm()
    Dim MyUserForm As ChangeParentForm
    
    ' Initialiser les variables global
    SetVariables
    
    ' Créer un nouveau UserForm (pas encore afficher)
    Set MyUserForm = New ChangeParentForm
    
    ' Afficher le UserForm
    MyUserForm.Show 0
End Sub
