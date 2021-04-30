Attribute VB_Name = "FoldersAndFiles"
'@Folder("General Tools")
Option Explicit

'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function FolderUnzip(ByVal FolderPath As String, Optional ByVal UnzipFolderPath As String) As String

    If UnzipFolderPath = vbNullString Then UnzipFolderPath = FolderPath & " Unzip\"
    MakeDirectory DirectoryPath:=UnzipFolderPath

    Dim ShellApplication               As Object:        Set ShellApplication = CreateObject("Shell.Application")
    ShellApplication.Namespace(CVar(UnzipFolderPath)).CopyHere ShellApplication.Namespace(FolderPath & "\").Items

    FolderUnzip = UnzipFolderPath
End Function

Public Function FolderZip(ByVal FolderPathSource As String, Optional ByVal ZipPathDestination As String) As String
    

        If ZipPathDestination = vbNullString Then ZipPathDestination = DirectoryParent(DirectoryPath:=FolderPathSource) & "Temporary.zip"
        FolderZip = ZipPathDestination
        'Create empty Zip File
        ZipCreateNewEmptyFile (ZipPathDestination)
    
        Dim ShellApplication               As Object:        Set ShellApplication = CreateObject("Shell.Application")
        ShellApplication.Namespace(CStr(ZipPathDestination)).CopyHere ShellApplication.Namespace(CStr(FolderPathSource)).Items
    
    
        'Keep script waiting until Compressing is done
        On Error Resume Next
        Do Until ShellApplication.Namespace(CStr(ZipPathDestination)).Items.Count = ShellApplication.Namespace(CStr(FolderPathSource)).Items.Count
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        On Error GoTo 0

End Function

Private Function ZipCreateNewEmptyFile(ByVal FilePath As String) As String
    If Len(Dir(FilePath)) > 0 Then Kill FilePath
    Open FilePath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Function

Public Sub FileCreateCopy(ByVal Source As String, ByVal Destination As String, Optional ByVal Overwritefiles As Boolean = True)

    Dim FSO                            As FileSystemObject
    Set FSO = New FileSystemObject
    MakeDirectory DirectoryPath:=DirectoryParent(DirectoryPath:=Destination)
    FSO.CopyFile Source:=Source, Destination:=Destination, Overwritefiles:=Overwritefiles

End Sub

Public Sub MakeDirectory(ByVal DirectoryPath As String)

    Dim FSO                            As FileSystemObject
    Set FSO = New FileSystemObject

    If Not FSO.FolderExists(DirectoryParent(DirectoryPath:=DirectoryPath)) Then MakeDirectory DirectoryPath:=DirectoryParent(DirectoryPath:=DirectoryPath)
    If Not FSO.FolderExists(DirectoryPath) Then FSO.CreateFolder Path:=GetDocLocalPath(docPath:=DirectoryPath)

End Sub

Public Function DirectoryParent(ByVal DirectoryPath As String) As String
    DirectoryPath = GetDocLocalPath(docPath:=DirectoryPath)
    DirectoryParent = Left$(DirectoryPath, InStrRev(DirectoryPath, "\", , vbTextCompare) - 1)
End Function

Public Function FileExists(strFileName As String) As Boolean


    If strFileName = "" Then
        GoTo Exit_Point
    End If

    On Error Resume Next
    Dim DirTest                        As String
    DirTest = Dir(PathName:=strFileName, Attributes:=vbNormal)

    If DirTest <> "" Then
        FileExists = True
    End If

Exit_Point:

    Exit Function

End Function

Public Function FolderExists(strDataFolder As String) As Boolean

    On Error Resume Next
    If Dir(PathName:=strDataFolder, Attributes:=vbDirectory) <> "" Then
        If Err.Number = 0 Then
            FolderExists = True
        End If
    End If


End Function

Public Function GetUserSelectedPath(Optional ByVal DefaultPath As String, Optional FileType As MsoFileDialogType = msoFileDialogOpen) As String
    If DefaultPath = "" Then DefaultPath = Application.DefaultFilePath
     
    With Application.FileDialog(FileType)
        If DefaultPath <> vbNullString Then
            If Right$(DefaultPath, 1) = "\" Then DefaultPath = Left$(DefaultPath, Len(DefaultPath))
            .InitialFileName = DefaultPath
        End If
        If .Show <> 0 Then
            GetUserSelectedPath = .SelectedItems.Item(1)
        Else
            Debug.Print "Process cancelled by user. "
            End
        End If
    End With
    
End Function

'__________________________________________________________________________________________________________________________________
'https://www.thespreadsheetguru.com/blog/vba-guide-text-files

'Some Terminology
'When we are working with text files, there will be some terminology that you probably haven't seen or used before when writing VBA code.
'Let's walk through some of the pieces you will see throughout the code in this guide.
'
'For Output - When you are opening the text file with this command, you are wanting to create or modify the text file.
'           You will not be able to pull anything from the text file while opening with this mode.
'
'For Input - When you are opening the text file with this command, you are wanting to extract information from the text file.
'           You will not be able to modify the text file while opening it with this mode.
'
'For Append - Add new text to the bottom of your text file content.
'
'FreeFile - Is used to supply a file number that is not already in use. This is similar to referencing Workbook(1) vs. Workbook(2).
'           By using FreeFile, the function will automatically return the next available reference number for your text file.
'
'Write - This writes a line of text to the file surrounding it with quotations
'
'Print - This writes a line of text to the file without quotations


Sub TextFileCreate(ByVal FilePath As String, ByVal FileContent As String)

    Dim TextFile                       As Integer          'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile
    
    Open FilePath For Output As TextFile
    Print #TextFile, FileContent
    Close TextFile

End Sub

Function TextFileGetContent(ByVal FilePath As String) As String


    Dim TextFile                       As Integer
    TextFile = FreeFile                                    'Determine the next file number available for use by the FileOpen function

    Open FilePath For Input As TextFile
    TextFileGetContent = Input(LOF(TextFile), TextFile)
    Close TextFile
    Debug.Print TextFileGetContent
End Function

Sub TextFileFindReplace(ByVal FilePath As String, ByVal FindString As String, ByVal ReplaceString As String)

    Dim FileContent                    As String
    FileContent = TextFileGetContent(FilePath:=FilePath)
    FileContent = Replace(FileContent, FindString, ReplaceString)

    TextFileCreate FilePath:=FilePath, FileContent:=FileContent

End Sub

Sub TextFileAppend(ByVal FilePath As String, ByVal FileContent As String)

    Dim TextFile                       As Integer
    TextFile = FreeFile

    Open FilePath For Append As TextFile
    Print #TextFile, FileContent
    Close TextFile

End Sub

Sub TextFileToArray(ByVal FilePath As String, Optional ByVal Delimiter As String)
    'PURPOSE: Load an Array variable with data from a delimited text file

    Dim rw                             As Long
    Dim col                            As Long

    rw = 0
    col = 0

    Dim FileContent                    As String
    FileContent = TextFileGetContent(FilePath:=FilePath)

    Dim LineArray()                    As String
    
    LineArray() = Split(FileContent, vbCrLf)               'Separate Out lines of data

    Dim DataArray()                    As String
    Dim TempArray()                    As String
    
    'Read Data into an Array Variable
    Dim x                              As Long
    For x = LBound(LineArray) To UBound(LineArray)
        If Len(Trim(LineArray(x))) <> 0 Then
            'Split up line of text by delimiter
            TempArray = Split(LineArray(x), Delimiter)
      
            'Determine how many columns are needed
            col = UBound(TempArray)
      
            'Re-Adjust Array boundaries
            ReDim Preserve DataArray(col, rw)
      
            'Load line of data into Array variable
            Dim y                      As Long
            For y = LBound(TempArray) To UBound(TempArray)
                DataArray(y, rw) = TempArray(y)
            Next y
        End If
    
        'Next line
        rw = rw + 1
    
    Next x

End Sub

'__________________________________________________________________________________________________________________________________

Public Function GetTempFolder() As String
    GetTempFolder = CreateObject("scripting.filesystemobject").GetSpecialFolder(2)
End Function

Public Function GetDocLocalPath(docPath As String) As String
'Gel Local Path NOT URL to Onedrive
Const strcOneDrivePart As String = "https://d.docs.live.net/"
Dim strRetVal As String, bytSlashPos As Byte

  strRetVal = docPath & "\"
  If Left(LCase(docPath), Len(strcOneDrivePart)) = strcOneDrivePart Then 'yep, it's the OneDrive path
    'locate and remove the "remote part"
    bytSlashPos = InStr(Len(strcOneDrivePart) + 1, strRetVal, "/")
    strRetVal = Mid(docPath, bytSlashPos)
    'read the "local part" from the registry and concatenate
    strRetVal = RegKeyRead("HKEY_CURRENT_USER\Environment\OneDrive") & strRetVal
    strRetVal = Replace(strRetVal, "/", "\") 'slashes in the right direction
    strRetVal = Replace(strRetVal, "%20", " ") 'a space is a space once more
End If
If Right(strRetVal, 1) = Application.PathSeparator Then strRetVal = Left(strRetVal, Len(strRetVal) - 1)
GetDocLocalPath = strRetVal

End Function

Private Function RegKeyRead(i_RegKey As String) As String
Dim myWS As Object

  On Error Resume Next
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'read key from registry
  RegKeyRead = myWS.RegRead(i_RegKey)
End Function
