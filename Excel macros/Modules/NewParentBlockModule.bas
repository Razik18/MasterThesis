Attribute VB_Name = "NewParentBlockModule"

Sub OpenNewParentBlockForm()
    Dim MyUserForm As NewParentBlockForm
    
    ' Initialiser les variables global
    SetVariables
       
    ' Créer un nouveau UserForm (pas encore afficher)
    Set MyUserForm = New NewParentBlockForm
    
    ' Afficher le UserForm
    MyUserForm.Show 0
    
End Sub

