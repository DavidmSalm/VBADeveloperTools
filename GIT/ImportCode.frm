VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportCode 
   Caption         =   "Select Addin to Export"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3780
   OleObjectBlob   =   "ImportCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImportCode"
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
        If CurrentVBProject.Name <> "VBAProject" Then ComboBox1.AddItem pvargItem:=CurrentVBProject.Name
    Next CurrentVBProject
    ComboBox1.AddItem pvargItem:="End of Add-ins"
    ComboBox1.ListIndex = 0
    
End Sub


Private Sub OkButton_Click()
    Dim CurrentVBProject As VBIDE.VBProject
    Set CurrentVBProject = Application.VBE.VBProjects.Item(ComboBox1.Value)
    
    Dim FolderPath As String: FolderPath = GetFolderPath(DefaultPath:=ThisWorkbook.path)
    If Right$(FolderPath, InStrRev(FolderPath, "\", , vbTextCompare)) <> CurrentVBProject.Name Then
        If MsgBox(prompt:="Folder name and VB project are not the same. Are you sure you want to import the code into this project?", Buttons:=vbYesNo) = vbNo Then
            Debug.Print "Process stoped by the user."
            Exit Sub
        End If
    End If

    
    ImportAllModulesintoVBProject CurrentVBProject:=CurrentVBProject, FolderPath:=FolderPath
    
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub


Private Sub ImportAllModulesintoVBProject(ByVal CurrentVBProject As VBIDE.VBProject, ByVal FolderPath As String)



End Sub



Private Function GetFolderPath(ByVal DefaultPath As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If DefaultPath <> vbNullString Then
            If Right$(DefaultPath, 1) = "\" Then DefaultPath = Left$(DefaultPath, Len(DefaultPath))
            .InitialFileName = DefaultPath
        End If
        If .Show <> 0 Then GetFolderPath = .SelectedItems(1)
    End With
End Function
