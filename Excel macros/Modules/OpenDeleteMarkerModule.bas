Attribute VB_Name = "OpenDeleteMarkerModule"
Sub OpenDeleteMarkerForm()
    Dim MyUserForm As MarkerDeleteForm
    
    ' Create a new instance of the UserForm
    Set MyUserForm = New MarkerDeleteForm
    ' Show the form
    MyUserForm.Show
End Sub

