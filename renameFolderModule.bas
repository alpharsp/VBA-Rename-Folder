Attribute VB_Name = "Module1"
Public Function renameFolder() As Boolean
    'Need to add "Microsoft Scripting Runtime" From Tools/References
    Dim fd As FileDialog
    Dim fs As FileSystemObject
    Dim folderPath As String
    Dim newName As String
    Dim cFolder As Folder
    Dim newpath As String
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    Set fs = New FileSystemObject
    
    renameFolder = False
    fd.AllowMultiSelect = False
    fd.Show
    If fd.SelectedItems.Count > 0 Then
        folderPath = fd.SelectedItems(1)
        If fs.FolderExists(folderPath) Then
            Set cFolder = fs.GetFolder(folderPath)
enterName:
            newName = InputBox("New name for the folder")
            If newName <> "" Then
                newpath = cFolder.ParentFolder.Path & "/" & newName
                If Not fs.FolderExists(newpath) Then
                    cFolder.Name = newName
                    renameFolder = True
                Else
                    MsgBox "Folder already exists, enter a new name"
                    GoTo enterName
                End If
            Else
                MsgBox "You either pressed cancel or didn't enter a new name, operation will be aborted"
            End If
        Else
            MsgBox "Folder doesn't exist, operation will be aborted"
        End If
    Else
        MsgBox "No folder selected, operation will be aborted"
    End If
End Function
