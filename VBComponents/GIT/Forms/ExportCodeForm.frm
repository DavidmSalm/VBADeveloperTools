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

Private ThisWorkbookPath               As String
Private BaseFolderPath                 As String
Private VBComponentBaseFolderPath      As String
Private VBProjectCompiledFilePath          As String

Private CurrentVBProject               As VBIDE.VBProject

Private Sub UserForm_Initialize()
    Dim CurrentVBProject               As VBIDE.VBProject
    
    For Each CurrentVBProject In Application.VBE.VBProjects
        Debug.Print CurrentVBProject.Name
        If CurrentVBProject.Name <> "VBAProject" Then AddinSelection.AddItem pvargItem:=CurrentVBProject.Name
    Next CurrentVBProject
    AddinSelection.AddItem pvargItem:="(End of Add-ins)"
    AddinSelection.ListIndex = 0

End Sub

Private Sub OkButton_Click()
    If AddinSelection.Value = "(End of Add-ins)" Then
        Debug.Print "please select an add-in"
        End
    End If

    Set CurrentVBProject = Application.VBE.VBProjects.Item(AddinSelection.Value)
    
    PopulateGlobalStrings
    
    If Not FoldersAndFiles.FolderExists(strDataFolder:=VBComponentBaseFolderPath) Then
        BaseFolderPath = FoldersAndFiles.GetUserSelectedPath(DefaultPath:=BaseFolderPath, FileType:=msoFileDialogFolderPicker)
    End If
    
    LoopThrougFolderandDeleteCode FolderPath:=VBComponentBaseFolderPath
    ExportAllModulesinVBProject CurrentVBProjectLocal:=CurrentVBProject, FolderPath:=VBComponentBaseFolderPath
    GitRestoreUnchangedForms RepositoryPath:=BaseFolderPath
    
    FoldersAndFiles.FileCreateCopy Source:=CurrentVBProject.fileName, Destination:=VBProjectCompiledFilePath
    Unload Me
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub PopulateGlobalStrings()
    ThisWorkbookPath = ThisWorkbook.Path
    BaseFolderPath = ThisWorkbook.Path & Application.PathSeparator & CurrentVBProject.Name
    VBComponentBaseFolderPath = BaseFolderPath & Application.PathSeparator & "VBComponents"
    VBProjectCompiledFilePath = BaseFolderPath & Application.PathSeparator & CurrentVBProject.Name & ".xlam"
End Sub

Private Sub LoopThrougFolderandDeleteCode(ByVal FolderPath As String)
    Dim FSO                            As FileSystemObject
    Set FSO = New FileSystemObject
    Dim Folder                         As Folder
    Set Folder = FSO.GetFolder(FolderPath:=FolderPath)
    
    Dim SubFolder                      As Folder
    
    For Each SubFolder In Folder.SubFolders
        LoopThrougFolderandDeleteCode FolderPath:=SubFolder.Path
    Next SubFolder
    
    Dim CurrentFile                    As File
    
    For Each CurrentFile In Folder.Files
        If Right$(CurrentFile.Path, 4) = ".bas" Or Right$(CurrentFile.Path, 4) = ".cls" Or Right$(CurrentFile.Path, 4) = ".frm" Or Right$(CurrentFile.Path, 4) = ".frx" Then
            Debug.Print "Killed: ", CurrentFile.Path
            Kill CurrentFile.Path
        End If
    Next CurrentFile
    

End Sub

Private Sub ExportAllModulesinVBProject(ByVal CurrentVBProjectLocal As VBIDE.VBProject, ByVal FolderPath As String)

    Dim CurrentModule                  As VBIDE.VBComponent
    For Each CurrentModule In CurrentVBProjectLocal.VBComponents
        If CurrentModule.Type <> vbext_ct_Document Then
            Dim fileName               As String
            fileName = CreateFileNameforModule(FolderPath:=FolderPath, CurrentModule:=CurrentModule)
            FoldersAndFiles.MakeDirectory DirectoryPath:=FoldersAndFiles.DirectoryParent(DirectoryPath:=fileName)
            CurrentModule.Export fileName:=fileName
        End If
    Next CurrentModule

End Sub

Private Function CreateFileNameforModule(ByVal FolderPath As String, ByVal CurrentModule As VBIDE.VBComponent) As String
    Dim SubFolderIndicator             As String: SubFolderIndicator = ModuleFolderIndicator(CurrentModule:=CurrentModule)
    If SubFolderIndicator <> vbNullString Then
        CreateFileNameforModule = FolderPath & Application.PathSeparator & SubFolderIndicator & Application.PathSeparator
        FoldersAndFiles.MakeDirectory DirectoryPath:=FolderPath & Application.PathSeparator & SubFolderIndicator
    Else
        CreateFileNameforModule = FolderPath & Application.PathSeparator
    End If
    
    Select Case CurrentModule.Type
        Case vbext_ct_ClassModule: CreateFileNameforModule = CreateFileNameforModule & CurrentModule.Name & ".cls"
        Case vbext_ct_MSForm: CreateFileNameforModule = CreateFileNameforModule & "Forms" & Application.PathSeparator & CurrentModule.Name & ".frm"
        Case vbext_ct_StdModule: CreateFileNameforModule = CreateFileNameforModule & CurrentModule.Name & ".bas"
    End Select
    
    'Debug.Print CurrentModule.Name, ModuleFolderIndicator(CurrentModule:=CurrentModule)
End Function

Private Function ModuleFolderIndicator(ByVal CurrentModule As VBIDE.VBComponent) As String
    
    With CurrentModule.CodeModule
        Dim LineofCode                 As Long
        
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

Private Function GitRestoreUnchangedForms(ByVal RepositoryPath As String) As String
   
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1

    'run command
    Dim oExec As Object
    Dim oOutput As Object
     oShell.CurrentDirectory = RepositoryPath
    Set oExec = oShell.Exec("git status")
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    Dim FRXDictionary As Dictionary
    Set FRXDictionary = New Dictionary
    Dim OutputDictionary As Dictionary
    Set OutputDictionary = New Dictionary
    
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then
        s = s & sLine & vbCrLf
        If Not OutputDictionary.Exists(Key:=sLine) Then
        OutputDictionary.Add Key:=sLine, Item:=sLine
        If InStr(1, sLine, "modified:") <> 0 And InStr(1, sLine, ".frx") <> 0 Then
            Debug.Print sLine
            FRXDictionary.Add Key:=sLine, Item:=sLine
            End If
            End If
        End If
    Wend

    Dim frxkey As Variant
    For Each frxkey In FRXDictionary.Keys()
        Dim frmkey As String
        frmkey = Replace(expression:=frxkey, Find:=".frx", Replace:=".frm")
        If Not OutputDictionary.Exists(Key:=frmkey) Then
            Dim frxlocalfilepath As String
            frxlocalfilepath = frxkey
            frxlocalfilepath = Right(frxlocalfilepath, Len(frxlocalfilepath) - InStrRev(frxlocalfilepath, "modified:") - 9)
            frxlocalfilepath = Trim(frxlocalfilepath)
            Dim gitcommand As String
            gitcommand = "git restore """ & frxlocalfilepath & """"
            oShell.Run gitcommand, , waitOnReturn
            Debug.Print "removed:", gitcommand
        Else
            Debug.Print "Kept:", frmkey
        End If
        
    Next frxkey

    GitRestoreUnchangedForms = s

End Function
