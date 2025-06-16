VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewTMABlockForm 
   Caption         =   "Create New TMA Block"
   ClientHeight    =   6276
   ClientLeft      =   -987
   ClientTop       =   -4557
   ClientWidth     =   7098
   OleObjectBlob   =   "NewTMABlockForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewTMABlockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Dim NumTMA As String
    
    ' Récuperer le Number of Child
    NumTMA = Me.TextBox4.value
    
    ' Récupérer la liste des projects
    Dim ListParentBlockName As Collection
    Set ListParentBlockName = New Collection 'Créer une liste des marquers (vide encore)
    
    For i = 0 To Me.ListBox1.ListCount - 1
        ListParentBlockName.Add Me.ListBox1.List(i)
    Next i
    
    ' Créer le(s) nouveaux TMA block(s)
    NewTMABlock NumTMA, ListParentBlockName
    
    ' Fermer le UserForm si tout a bien marché
    Unload Me
    
End Sub

Private Sub CommandButton2_Click()
    Dim ParentBlockName As String

    ' Récuperer le Vendor Block ID
    ParentBlockName = Me.TextBox5.value
    
    ' Créer le Parent Block
    AddParentToList ParentBlockName, Me
End Sub
