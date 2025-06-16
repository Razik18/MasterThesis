VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewChildBlockForm 
   Caption         =   "Create New Child Block"
   ClientHeight    =   3960
   ClientLeft      =   -686
   ClientTop       =   -2688
   ClientWidth     =   6160
   OleObjectBlob   =   "NewChildBlockForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewChildBlockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Dim ParentBlockName As String
    Dim NumChild As String
    Dim KeepParent As Boolean
    Dim Marker As String ' Add the Marker variable

    ' Retrieve the Vendor Block ID
    ParentBlockName = Me.LabelBlockID

    ' Retrieve the Marker
    Marker1 = Me.LabelMarker
    Marker = Replace(Marker1, "(in Review)", "")

    ' Retrieve the Number of Child
    NumChild = Me.TextBox3.value

    ' Retrieve the boolean on whether to keep the Parent Block
    KeepParent = Me.CheckBox1.value

    ' Create the new child block(s) and pass the Marker
    NewChildBlock ParentBlockName, NumChild, KeepParent, Marker

    ' Close the UserForm if everything works
    Unload Me
End Sub

