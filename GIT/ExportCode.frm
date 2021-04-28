VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportCode 
   Caption         =   "Select Addin to Export"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3780
   OleObjectBlob   =   "ExportCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExportCode"
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
    
    Dim FolderPath As String
    FolderPath = ThisWorkbook.path & "\" & CurrentVBProject.Name
    FolderPath = GetFolderPath(DefaultPath:=FolderPath)
    If FolderPath = vbNullString Then
        Debug.Print "Exited the export process, because folder path did not exist."
        Exit Sub
    End If
    
    ExportAllModulesinVBProject CurrentVBProject:=CurrentVBProject, FolderPath:=FolderPath
    
End Sub


Private Sub ExportAllModulesinVBProject(ByVal CurrentVBProject As VBIDE.VBProject, ByVal FolderPath As String)

    Dim CurrentModule As VBIDE.VBComponent
    For Each CurrentModule In CurrentVBProject.VBComponents
        If CurrentModule.Type <> vbext_ct_Document Then
            Dim FileName As String
            FileName = CreateFileNameforModule(FolderPath:=FolderPath, CurrentModule:=CurrentModule)
            CurrentModule.Export FileName:=FileName
        End If
    Next CurrentModule

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
        Dim LineofCode As Integer
        
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
    

End Function


Private Function GetFolderPath(ByVal DefaultPath As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If DefaultPath <> vbNullString Then
            If Right$(DefaultPath, 1) = "\" Then DefaultPath = Left$(DefaultPath, Len(DefaultPath))
            .InitialFileName = DefaultPath
        End If
        If .Show <> 0 Then GetFolderPath = .SelectedItems(1)
    End With
End Function


Private Sub MakeDirectory(ByVal DirectoryPath As String)

Dim FSO As New FileSystemObject

If Not FSO.FolderExists(DirectoryPath) Then
    FSO.CreateFolder path:=DirectoryPath
End If

End Sub
