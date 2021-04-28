Attribute VB_Name = "FoldersAndFiles"
'@Folder("General Tools")
Option Explicit

'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function Unzip(ByVal Filepath As String) As String

    Dim UnzipFolderPath                As String
    UnzipFolderPath = Filepath & " Unzip\"
    MakeDirectory DirectoryPath:=UnzipFolderPath

    Dim ShellApplication               As Object:        Set ShellApplication = CreateObject("Shell.Application")
    ShellApplication.Namespace(CVar(UnzipFolderPath)).CopyHere ShellApplication.Namespace(Filepath & "\").Items

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


Sub TextFileCreate(ByVal Filepath As String, ByVal FileContent As String)

    Dim TextFile                       As Integer          'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile
    
    Open Filepath For Output As TextFile
    Print #TextFile, FileContent
    Close TextFile

End Sub

Function TextFileGetContent(ByVal Filepath As String) As String


    Dim TextFile                       As Integer
    TextFile = FreeFile                                    'Determine the next file number available for use by the FileOpen function

    Open Filepath For Input As TextFile
    TextFileGetContent = Input(LOF(TextFile), TextFile)
    Close TextFile

End Function

Sub TextFileFindReplace(ByVal Filepath As String, ByVal FindString As String, ByVal ReplaceString As String)

    Dim FileContent                    As String
    FileContent = TextFileGetContent(Filepath:=TextFileGetContent)
    FileContent = Replace(FileContent, FindString, ReplaceString)

    TextFileCreate Filepath:=Filepath, FileContent:=FileContent

End Sub

Sub TextFileAppend(ByVal Filepath As String, ByVal FileContent As String)

    Dim TextFile                       As Integer
    TextFile = FreeFile

    Open Filepath For Append As TextFile
    Print #TextFile, FileContent
    Close TextFile

End Sub

Sub TextFileToArray(ByVal Filepath As String, Optional ByVal Delimiter As String)
    'PURPOSE: Load an Array variable with data from a delimited text file

    Dim rw                             As Long
    Dim col                            As Long

    rw = 0
      col = 0

    Dim FileContent                    As String
    FileContent = TextFileGetContent(Filepath:=Filepath)

    Dim LineArray()                    As String
    
    LineArray() = Split(FileContent, vbCrLf)               'Separate Out lines of data

    Dim DataArray()                    As String
    Dim TempArray()                    As String
    
    'Read Data into an Array Variable
    For x = LBound(LineArray) To UBound(LineArray)
        If Len(Trim(LineArray(x))) <> 0 Then
            'Split up line of text by delimiter
            TempArray = Split(LineArray(x), Delimiter)
      
            'Determine how many columns are needed
            col = UBound(TempArray)
      
            'Re-Adjust Array boundaries
            ReDim Preserve DataArray(col, rw)
      
            'Load line of data into Array variable
            For y = LBound(TempArray) To UBound(TempArray)
                DataArray(y, rw) = TempArray(y)
            Next y
        End If
    
        'Next line
        rw = rw + 1
    
    Next x

End Sub
'__________________________________________________________________________________________________________________________________

