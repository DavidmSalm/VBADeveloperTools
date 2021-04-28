Attribute VB_Name = "FoldersAndFiles"
Option Explicit
'@Folder("General Tools")

'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function Unzip(ByVal FilePath As String) As String

    Dim UnzipFolderPath                As String
    UnzipFolderPath = FilePath & " Unzip\"
    MakeDirectory DirectoryPath:=UnzipFolderPath

    Dim ShellApplication               As Object:        Set ShellApplication = CreateObject("Shell.Application")
    ShellApplication.Namespace(CVar(UnzipFolderPath)).CopyHere ShellApplication.Namespace(FilePath & "\").Items

    Unzip = UnzipFolderPath
End Function

Public Sub MakeDirectory(ByVal DirectoryPath As String)

    Dim FSO                            As FileSystemObject
    Set FSO = New FileSystemObject

    If Not FSO.FolderExists(DirectoryParent(DirectoryPath:=DirectoryPath)) Then MakeDirectory DirectoryPath:=DirectoryParent(DirectoryPath:=DirectoryPath)
    If Not FSO.FolderExists(DirectoryPath) Then FSO.CreateFolder Path:=DirectoryPath

End Sub

Private Function DirectoryParent(ByVal DirectoryPath As String) As String
    DirectoryParent = Left$(DirectoryPath, InStrRev(DirectoryPath, "\", , vbTextCompare) - 1)
End Function

Public Function FolderExists(FolderPath As String) As Boolean
    FolderExists = (Dir(FolderPath, vbDirectory) <> vbNullString)
End Function
