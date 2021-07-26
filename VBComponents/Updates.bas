Attribute VB_Name = "Updates"
#Const Debugging = True
Option Explicit

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
      Alias "URLDownloadToFileA" ( _
        ByVal pCaller As LongPtr, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As LongPtr, _
        ByVal lpfnCB As LongPtr) As Long

Private JSON As Object

Public Sub CheckforUpdates()

    GetRepositoryInformation
    If NeedsUpdate Then Update

End Sub


Private Sub GetRepositoryInformation()

    Dim WebRequest As MSXML2.XMLHTTP60
    Set WebRequest = New MSXML2.XMLHTTP60
    WebRequest.Open bstrMethod:="GET", bstrURL:="https://api.github.com/repos/DavidmSalm/VBADeveloperTools/releases/latest"
    WebRequest.send
    
    'Parse the response
    Set JSON = JSONConverter.ParseJson(JsonString:=WebRequest.responseText)
        
    #If Debugging Then
        Debug.Print JSON.Item("tag_name")
        Debug.Print JSON.Item("zipball_url")
    #End If
End Sub

Private Function NeedsUpdate() As Boolean

    NeedsUpdate = JSON.Item("tag_name") > ThisWorkbook.CustomDocumentProperties.Item("Version")
    If NeedsUpdate Then NeedsUpdate = UserSelectsUpdate

    #If Debugging Then
        Debug.Print "tag_name:", JSON.Item("tag_name")
        Debug.Print "ThisWorkbook Version:", ThisWorkbook.CustomDocumentProperties.Item("Version")
        Debug.Print "NeedsUpdate:", NeedsUpdate
    #End If

End Function

Private Function UserSelectsUpdate() As Boolean
    Select Case MsgBox(Prompt:="Update is available for the VBA developer tools. Would you like to update", Buttons:=vbYesNo)
        Case vbYes: UserSelectsUpdate = True
        Case vbNo: UserSelectsUpdate = False
    End Select
End Function

Private Sub Update()
    Dim UpdatedZIPFileLocalPath As String
    UpdatedZIPFileLocalPath = ThisWorkbook.Path & "\UpdatedFiles.zip"
    DownloadFiles DownloadLocation:=UpdatedZIPFileLocalPath
            
    Dim UpdatedFileLocalPath As String
    UpdatedFileLocalPath = ThisWorkbook.Path & "\UpdatedFiles"
    FoldersAndFiles.FolderUnzip FolderPath:=UpdatedZIPFileLocalPath, UnzipFolderPath:=UpdatedFileLocalPath

    Dim Files As Collection: Set Files = New Collection
    Set Files = FoldersAndFiles.FindFilesbyName(FileName:="VBADeveloperTools.xlam", ParentFolderPath:=UpdatedFileLocalPath)

    If Files.Count = 1 Then
        Dim Currentfile As File
        Set Currentfile = Files.Item(1)
        Dim ThisworkbookFullName As String
        ThisworkbookFullName = ThisWorkbook.FullName
        ThisWorkbook.ChangeFileAccess Mode:=xlReadOnly
        Kill PathName:=ThisworkbookFullName
        FoldersAndFiles.FileCreateCopy Source:=Currentfile.Path, Destination:=ThisworkbookFullName
    Else
        Debug.Print "ERROR"
        
    End If

    Kill PathName:=UpdatedZIPFileLocalPath
    FoldersAndFiles.FolderDelete FolderPath:=UpdatedFileLocalPath
    
    ThisWorkbook.SaveAs FileName:=ThisWorkbook.Path & Application.PathSeparator & CreateObject("Scripting.FileSystemObject").GetTempName()
    
    Workbooks.Open FileName:=ThisworkbookFullName
    
    ThisWorkbook.Close
    
    #If Debugging Then
        Debug.Print "Updated"
    #End If
End Sub

Private Sub DownloadFiles(ByVal DownloadLocation As String)
    URLDownloadToFile pCaller:=0, szURL:=JSON.Item("zipball_url"), szFileName:=DownloadLocation, dwReserved:=0, lpfnCB:=0
End Sub


