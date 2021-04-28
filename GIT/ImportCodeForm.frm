VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportCodeForm 
   Caption         =   "Select Addin to Export"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3780
   OleObjectBlob   =   "ImportCodeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImportCodeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("GIT")
Option Explicit


Private Sub UserForm_Initialize()
    Dim CurrentVBProject As VBIDE.VBProject
    
    
    For Each CurrentVBProject In Application.VBE.VBProjects
        Debug.Print CurrentVBProject.Name
        If CurrentVBProject.Name <> "VBAProject" Then AddinSelection.AddItem pvargItem:=CurrentVBProject.Name
    Next CurrentVBProject
    AddinSelection.AddItem pvargItem:="End of Add-ins"
    AddinSelection.ListIndex = 0
    
End Sub


Private Sub OkButton_Click()
    Dim CurrentVBProject As VBIDE.VBProject
    Set CurrentVBProject = Application.VBE.VBProjects.Item(AddinSelection.Value)
    
    Dim FolderPath As String: FolderPath = GetFolderPath(DefaultPath:=ThisWorkbook.Path)
    If Right$(FolderPath, Len(FolderPath) - InStrRev(FolderPath, "\", , vbTextCompare)) <> CurrentVBProject.Name Then
        If MsgBox(prompt:="Folder name and VB project are not the same. Are you sure you want to import the code into this project?", Buttons:=vbYesNo) = vbNo Then
            Debug.Print "Process stoped by the user."
            Exit Sub
        End If
    End If

    
    ImportAllModulesintoVBProject CurrentVBProject:=CurrentVBProject, FolderPath:=FolderPath
    
    Unload Me
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub


Private Sub ImportAllModulesintoVBProject(ByVal CurrentVBProject As VBIDE.VBProject, ByVal FolderPath As String)
    
    
    DeleteVBAModulesandUserForms CurrentVBProject:=CurrentVBProject
    LoopThrougFolderandImportCode CurrentVBProject:=CurrentVBProject, FolderPath:=FolderPath

End Sub

Private Sub LoopThrougFolderandImportCode(ByVal CurrentVBProject As VBIDE.VBProject, ByVal FolderPath As String)
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    Dim Folder As Folder
    Set Folder = FSO.GetFolder(FolderPath:=FolderPath)
    
    Dim SubFolder As Folder
    
    For Each SubFolder In Folder.SubFolders
        LoopThrougFolderandImportCode CurrentVBProject:=CurrentVBProject, FolderPath:=SubFolder.Path
    Next SubFolder
    
    Dim CurrentFile As File
    
    For Each CurrentFile In Folder.Files
        If Right$(CurrentFile.Path, 4) = ".bas" Or Right$(CurrentFile.Path, 4) = ".cls" Or Right$(CurrentFile.Path, 4) = ".frm" Then
            CurrentVBProject.VBComponents.Import fileName:=CurrentFile.Path
            Debug.Print "Imported: ", CurrentFile.Path
        End If
    Next CurrentFile
    

End Sub

Private Function GetFolderPath(ByVal DefaultPath As String) As String
    Dim DefaultPathLocal As String: DefaultPathLocal = DefaultPath
     
    With Application.FileDialog(msoFileDialogFolderPicker)
        If DefaultPathLocal <> vbNullString Then
            If Right$(DefaultPathLocal, 1) = "\" Then DefaultPathLocal = Left$(DefaultPathLocal, Len(DefaultPathLocal))
            .InitialFileName = DefaultPathLocal
        End If
        If .Show <> 0 Then GetFolderPath = .SelectedItems.Item(1)
    End With
End Function

Private Sub DeleteVBAModulesandUserForms(CurrentVBProject As VBIDE.VBProject)
    Dim CurrentVBAPComponent As VBIDE.VBComponent
    
    For Each CurrentVBAPComponent In CurrentVBProject.VBComponents
        If CurrentVBAPComponent.Type <> vbext_ct_Document Then CurrentVBProject.VBComponents.Remove CurrentVBAPComponent
    Next CurrentVBAPComponent

End Sub
