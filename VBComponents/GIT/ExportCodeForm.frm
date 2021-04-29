VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportCodeForm 
   Caption         =   "Select Addin to Export"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3780
   OleObjectBlob   =   "ExportCodeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExportCodeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("GIT")
Option Explicit

Private CurrentVBProject As VBIDE.VBProject

Private Sub UserForm_Initialize()
    Dim CurrentVBProjectLocal As VBIDE.VBProject
    For Each CurrentVBProjectLocal In Application.VBE.VBProjects
        Debug.Print CurrentVBProjectLocal.Name
        If CurrentVBProjectLocal.Name <> "VBAProject" Then AddinSelection.AddItem pvargItem:=CurrentVBProjectLocal.Name
    Next CurrentVBProjectLocal
    AddinSelection.AddItem pvargItem:="End of Add-ins"
    AddinSelection.ListIndex = 0

End Sub


Private Sub OkButton_Click()
    Dim CurrentVBProjectLocal As VBIDE.VBProject
    Set CurrentVBProjectLocal = Application.VBE.VBProjects.Item(AddinSelection.Value)
    Set CurrentVBProject = CurrentVBProjectLocal
    Dim FolderPath As String
    FolderPath = ThisWorkbook.Path & "\" & CurrentVBProjectLocal.Name & "\VBComponents"
    FolderPath = GetFolderPath(DefaultPath:=FolderPath)
    If FolderPath = vbNullString Then
        Debug.Print "Exited the export process, because folder path did not exist."
        Exit Sub
    End If
    
    DeleteExistingModulesinFolder
    ExportAllModulesinVBProject CurrentVBProjectLocal:=CurrentVBProjectLocal, FolderPath:=FolderPath
    
    Unload Me
End Sub


Private Sub ExportAllModulesinVBProject(ByVal CurrentVBProjectLocal As VBIDE.VBProject, ByVal FolderPath As String)

    Dim CurrentModule As VBIDE.VBComponent
    For Each CurrentModule In CurrentVBProjectLocal.VBComponents
        If CurrentModule.Type <> vbext_ct_Document Then
            Dim fileName As String
            fileName = CreateFileNameforModule(FolderPath:=FolderPath, CurrentModule:=CurrentModule)
            CurrentModule.Export fileName:=fileName
        End If
    Next CurrentModule

End Sub

Private Sub DeleteExistingModulesinFolder(ByVal FolderPath As String)
    Dim FSO                            As FileSystemObject
    Set FSO = New FileSystemObject
    Dim Folder                         As Folder
    Set Folder = FSO.GetFolder(FolderPath:=FolderPath)
    
    Dim SubFolder                      As Folder
    
    For Each SubFolder In Folder.SubFolders
        DeleteExistingModulesinFolder FolderPath:=SubFolder.Path
    Next SubFolder
    
    Dim CurrentFile                    As File
    
    For Each CurrentFile In Folder.Files
        If Right$(CurrentFile.Path, 4) = ".bas" Or Right$(CurrentFile.Path, 4) = ".cls" Or Right$(CurrentFile.Path, 4) = ".frm" Then
            Kill CurrentFile.Path
            Debug.Print "Killed: ", CurrentFile.Path
        End If
    Next CurrentFile
    
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Function CreateFileNameforModule(ByVal FolderPath As String, ByVal CurrentModule As VBIDE.VBComponent) As String
    Dim SubFolderIndicator As String: SubFolderIndicator = ModuleFolderIndicator(CurrentModule:=CurrentModule)
    If SubFolderIndicator <> vbNullString Then
        CreateFileNameforModule = FolderPath & "\" & SubFolderIndicator & "\" & CurrentModule.Name
        MakeDirectory DirectoryPath:=FolderPath & "\" & SubFolderIndicator
    Else
        CreateFileNameforModule = FolderPath & "\" & CurrentModule.Name
    End If
    
    Select Case CurrentModule.Type
        Case vbext_ct_ClassModule: CreateFileNameforModule = CreateFileNameforModule & ".cls"
        Case vbext_ct_MSForm: CreateFileNameforModule = CreateFileNameforModule & ".frm"
        Case vbext_ct_StdModule: CreateFileNameforModule = CreateFileNameforModule & ".bas"
    End Select
    
    Debug.Print CurrentModule.Name, ModuleFolderIndicator(CurrentModule:=CurrentModule)
End Function

Private Function ModuleFolderIndicator(ByVal CurrentModule As VBIDE.VBComponent) As String
    
    With CurrentModule.CodeModule
        Dim LineofCode As Long
        
        For LineofCode = 1 To .CountOfLines
            Debug.Print .Lines(LineofCode, 1)
            If InStr(1, .Lines(LineofCode, 1), "@Folder", vbTextCompare) Then
                Dim FirstQuotePosition As Long: FirstQuotePosition = InStr(1, .Lines(LineofCode, 1), """", vbTextCompare)
                Dim SecondQuotePOsition As Long: SecondQuotePOsition = InStr(FirstQuotePosition + 1, .Lines(LineofCode, 1), """", vbTextCompare)
                ModuleFolderIndicator = Mid$(.Lines(LineofCode, 1), FirstQuotePosition + 1, SecondQuotePOsition - FirstQuotePosition - 1)
                ModuleFolderIndicator = Replace$(ModuleFolderIndicator, ".", "\")
                Exit Function
            End If
        Next LineofCode
    End With
    ModuleFolderIndicator = CurrentVBProject.Name

End Function


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


Private Sub MakeDirectory(ByVal DirectoryPath As String)

Dim FSO As FileSystemObject
Set FSO = New FileSystemObject

If Not FSO.FolderExists(DirectoryParent(DirectoryPath:=DirectoryPath)) Then MakeDirectory DirectoryPath:=DirectoryParent(DirectoryPath:=DirectoryPath)
If Not FSO.FolderExists(DirectoryPath) Then FSO.CreateFolder Path:=DirectoryPath

End Sub

Private Function DirectoryParent(ByVal DirectoryPath As String) As String
    DirectoryParent = Left$(DirectoryPath, InStrRev(DirectoryPath, "\", , vbTextCompare) - 1)
End Function
