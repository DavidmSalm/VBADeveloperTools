VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportCodeForm 
   Caption         =   "Select Addin to Import"
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

Private SourceWorkbookPath             As String
Private WorkbookBackupPath             As String
Private WorkbookName                   As String
Private WorkbookZippedPath             As String
Private WorkbookZippedPatheWithTimestamp As String
Private WorkbookUnzippedFolderPath     As String
Private WorkbookXMLFolderPath          As String
Private CustomUIFilePath               As String
Private CustomUI14FilePath             As String
Private WorkbookInstalledAddin         As Excel.AddIn
Private RelationshipFilePath           As String

Private SourceFolderPath               As String
Private SourceFolderPathXustomUIXMLPath As String
Private SourceFolderPathXustomUI14XMLPath As String

Private CurrentVBProject As VBIDE.VBProject

Private Enum XMLType
    XMLType_customUI = 1
    XMLType_customUI14 = 2
End Enum

Const CustomUIRelType                  As String = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
Const CustomUI14RelType                As String = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

Private Sub PopulateGlobalStrings()
    SourceWorkbookPath = CurrentVBProject.fileName
    WorkbookName = Mid(SourceWorkbookPath, InStrRev(SourceWorkbookPath, "\") + 1, Len(SourceWorkbookPath))
    WorkbookBackupPath = ThisWorkbook.Path & Application.PathSeparator & "zArchive" & Application.PathSeparator & CurrentVBProject.Name & ".Archive." & GetTimestamp
    WorkbookZippedPath = SourceWorkbookPath & ".zip"
    WorkbookUnzippedFolderPath = Mid(SourceWorkbookPath, 1, InStrRev(SourceWorkbookPath, "\")) & "Unzipped " & WorkbookName & ".zip" & Application.PathSeparator
    WorkbookXMLFolderPath = WorkbookUnzippedFolderPath & "customUI"
    CustomUIFilePath = WorkbookXMLFolderPath & Application.PathSeparator & "CustomUI.xml"
    CustomUI14FilePath = WorkbookXMLFolderPath & Application.PathSeparator & "CustomUI14.xml"
    RelationshipFilePath = WorkbookUnzippedFolderPath & "_rels" & Application.PathSeparator & ".rels"
    WorkbookZippedPatheWithTimestamp = SourceWorkbookPath & GetTimestamp & ".zip"
    
    SourceFolderPathXustomUIXMLPath = SourceFolderPath & "\XML\CustomUI.xml"
    SourceFolderPathXustomUI14XMLPath = SourceFolderPath & "\XML\CustomUI14.xml"
    
End Sub

Private Sub UserForm_Initialize()
    Dim CurrentVBProject               As VBIDE.VBProject
    
    
    For Each CurrentVBProject In Application.VBE.VBProjects
        Debug.Print CurrentVBProject.Name
        If CurrentVBProject.Name <> "VBAProject" Then AddinSelection.AddItem pvargItem:=CurrentVBProject.Name
    Next CurrentVBProject
    AddinSelection.AddItem pvargItem:="End of Add-ins"
    AddinSelection.ListIndex = 0
    
End Sub

Private Sub OkButton_Click()
    Set CurrentVBProject = Application.VBE.VBProjects.Item(AddinSelection.Value)
    If CurrentVBProject.fileName = ThisWorkbook.FullName Then
        Dim ThisworkbookFullName As String: ThisworkbookFullName = ThisWorkbook.FullName
        ThisWorkbook.SaveAs fileName:=ThisWorkbook.Path & Application.PathSeparator & GetTimestamp & "Temp.xlsm"
        Dim CurrentWorkbook As Workbook
        Set CurrentWorkbook = Workbooks.Open(fileName:=ThisworkbookFullName)
        Set CurrentVBProject = CurrentWorkbook.VBProject
    End If
    
    
    SourceFolderPath = FoldersAndFiles.GetUserSelectedPath(DefaultPath:=ThisWorkbook.Path, FileType:=msoFileDialogFolderPicker)
    
    If Right$(SourceFolderPath, Len(SourceFolderPath) - InStrRev(SourceFolderPath, "\", , vbTextCompare)) <> CurrentVBProject.Name Then
        If MsgBox(prompt:="Folder name and VB project are not the same. Are you sure you want to import the code into this project?", Buttons:=vbYesNo) = vbNo Then
            Debug.Print "Process stoped by the user."
            Exit Sub
        End If
    End If
    
    PopulateGlobalStrings
    FoldersAndFiles.FileCreateCopy Source:=SourceWorkbookPath, Destination:=WorkbookBackupPath 'Create Backup
    
    UpdateSelectedVBProjectWithFileComponents
    If FoldersAndFiles.FileExists(strFileName:=SourceFolderPathXustomUIXMLPath) Or FoldersAndFiles.FileExists(strFileName:=SourceFolderPathXustomUI14XMLPath) Then UpdateXML
    
    DeleteFilesandFolders
    Unload Me
End Sub

Private Sub DeleteFilesandFolders()


End Sub



Private Sub CancelButton_Click()
    Unload Me
End Sub



Private Sub UpdateSelectedVBProjectWithFileComponents()
    DeleteVBAModulesandUserForms
    LoopThrougFolderandImportCode FolderPath:=SourceFolderPath
End Sub

Private Sub DeleteVBAModulesandUserForms()
    Dim CurrentVBComponent           As VBIDE.VBComponent
    
    For Each CurrentVBComponent In CurrentVBProject.VBComponents
        If CurrentVBComponent.Type <> vbext_ct_Document Then CurrentVBProject.VBComponents.Remove CurrentVBComponent
    Next CurrentVBComponent

    For Each CurrentVBComponent In CurrentVBProject.VBComponents
        If CurrentVBComponent.Type <> vbext_ct_Document Then
            Debug.Print "Not all VB Components were deleted. Please review and rerun."
            Debug.Assert False
        End If
    Next CurrentVBComponent


End Sub

Private Sub LoopThrougFolderandImportCode(ByVal FolderPath As String)
    Dim FSO                            As FileSystemObject
    Set FSO = New FileSystemObject
    Dim Folder                         As Folder
    Set Folder = FSO.GetFolder(FolderPath:=FolderPath)
    
    Dim SubFolder                      As Folder
    
    For Each SubFolder In Folder.SubFolders
        LoopThrougFolderandImportCode FolderPath:=SubFolder.Path
    Next SubFolder
    
    Dim CurrentFile                    As File
    
    For Each CurrentFile In Folder.Files
        If Right$(CurrentFile.Path, 4) = ".bas" Or Right$(CurrentFile.Path, 4) = ".cls" Or Right$(CurrentFile.Path, 4) = ".frm" Then
            CurrentVBProject.VBComponents.Import fileName:=CurrentFile.Path
            Debug.Print "Imported: ", CurrentFile.Path
        End If
    Next CurrentFile
    

End Sub

Private Sub UpdateXML()
    

    FoldersAndFiles.FileCreateCopy Source:=SourceWorkbookPath, Destination:=WorkbookZippedPath
    FoldersAndFiles.FolderUnzip FolderPath:=WorkbookZippedPath, UnzipFolderPath:=WorkbookUnzippedFolderPath
    
    If FoldersAndFiles.FileExists(strFileName:=SourceFolderPathXustomUIXMLPath) Then
        AddRelationshipifNeeded CurrentXMLType:=XMLType_customUI
        FoldersAndFiles.FileCreateCopy Source:=SourceFolderPathXustomUIXMLPath, Destination:=CustomUIFilePath
    End If
    If FoldersAndFiles.FileExists(strFileName:=SourceFolderPathXustomUI14XMLPath) Then
        AddRelationshipifNeeded CurrentXMLType:=XMLType_customUI14
        FoldersAndFiles.FileCreateCopy Source:=SourceFolderPathXustomUI14XMLPath, Destination:=CustomUI14FilePath
    End If
    
    Kill PathName:=WorkbookZippedPath
    FoldersAndFiles.FolderZip FolderPathSource:=WorkbookUnzippedFolderPath, ZipPathDestination:=WorkbookZippedPath
    
    Workbooks.Open(fileName:=SourceWorkbookPath).Close
    Kill PathName:=SourceWorkbookPath
    FoldersAndFiles.FileCreateCopy Source:=WorkbookZippedPath, Destination:=SourceWorkbookPath
    Workbooks.Open fileName:=SourceWorkbookPath
    
End Sub

Private Sub AddRelationshipifNeeded(ByVal CurrentXMLType As XMLType)
    If Not CustomUIRelationshipExists(RelationshipsFilePath:=RelationshipFilePath, RelationshipXMLType:=CurrentXMLType) Then
        AddCustomUIToRels RelationshipsFilePath:=RelationshipFilePath, RelationshipXMLType:=CurrentXMLType
    End If

End Sub

Private Function GetTimestamp() As String
    GetTimestamp = Format(Now(), "yyyymmddhhmmss") & Right(Format(Timer, "#0.00"), 2)
End Function

Private Sub SetUpWorkbookforRibbonUI(ByVal UnzippedWorkbookFolder As String)
    
    Dim RelationshipsFilePath          As String
    RelationshipsFilePath = UnzippedWorkbookFolder & "\_rels\.rels"
    
    
    If Not CustomUIRelationshipExists(RelationshipsFilePath:=RelationshipsFilePath, RelationshipXMLType:=XMLType_customUI14) Then
        AddCustomUIToRels RelationshipsFilePath:=RelationshipsFilePath, RelationshipXMLType:=XMLType_customUI14
    End If
    
    '    Dim RelationshipsFileContent       As String
    '    RelationshipsFileContent = FoldersAndFiles.TextFileGetContent(FilePath:=RelationshipsFilePath)
    '    Dim XMLRequiredRelationship        As String
    '    XMLRequiredRelationship = "<Relationship Id=""cuID14"" Type=""http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"" Target=""customUI/customUI14.xml"" />"
    '
    '    If InStr(1, RelationshipsFileContent, XMLRequiredRelationship, vbTextCompare) = 0 Then FoldersAndFiles.TextFileFindReplace FilePath:=RelationshipsFilePath, FindString:="</Relationships>", ReplaceString:=XMLRequiredRelationship & "</Relationships>"

End Sub

Private Sub AddCustomUIToRels(RelationshipsFilePath As String, RelationshipXMLType As XMLType)
    'Date Created : 5/14/2009 23:29
    'Author       : Ken Puls (www.excelguru.ca)
    'Modified by  : Doug Glancy 11/20/2017
    'Macro Purpose: Add the customUI relationship to the rels file

    Dim XMLDoc                        As MSXML2.DOMDocument60
    Dim XMLElement                    As MSXML2.IXMLDOMNode
    Dim XMLAttrib                     As MSXML2.IXMLDOMAttribute
    Dim NamedNodeMap                  As MSXML2.IXMLDOMNamedNodeMap

    'Create a new XML document
    Set XMLDoc = New MSXML2.DOMDocument60
    'Attach to the root element of the .rels file
    XMLDoc.Load RelationshipsFilePath

    'Create a new relationship element in the .rels file
    Set XMLElement = XMLDoc.createNode(1, "Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")
    Set NamedNodeMap = XMLElement.Attributes

    'Create ID attribute for the element
    Set XMLAttrib = XMLDoc.createAttribute("Id")
    Select Case RelationshipXMLType
        Case XMLType_customUI
            XMLAttrib.NodeValue = "cuID"
        Case XMLType_customUI14
            XMLAttrib.NodeValue = "cuID14"
    End Select
    NamedNodeMap.setNamedItem XMLAttrib

    'Create Type attribute for the element
    Set XMLAttrib = XMLDoc.createAttribute("Type")
    Select Case RelationshipXMLType
        Case XMLType_customUI
            XMLAttrib.NodeValue = CustomUIRelType
        Case XMLType_customUI14
            XMLAttrib.NodeValue = CustomUI14RelType
    End Select
    NamedNodeMap.setNamedItem XMLAttrib

    'Create Target element for the attribute
    Set XMLAttrib = XMLDoc.createAttribute("Target")
    Select Case RelationshipXMLType
        Case XMLType_customUI
            XMLAttrib.NodeValue = "customUI/customUI.xml"
        Case XMLType_customUI14
            XMLAttrib.NodeValue = "customUI/customUI14.xml"
    End Select
    NamedNodeMap.setNamedItem XMLAttrib

    'Now insert the new node at the proper location
    XMLDoc.ChildNodes(1).appendChild XMLElement
    'Save the .rels file
    XMLDoc.Save RelationshipsFilePath
    Set XMLAttrib = Nothing
    Set XMLElement = Nothing
    Set XMLDoc = Nothing
End Sub

Private Function CustomUIRelationshipExists(ByVal RelationshipsFilePath As String, ByVal RelationshipXMLType As XMLType)
    Dim oXMLDoc                        As MSXML2.DOMDocument60
    Dim XmlRelsNamespace               As String
    Dim RelType                        As String

    Select Case RelationshipXMLType
        Case XMLType.XMLType_customUI
            RelType = CustomUIRelType
        Case XMLType.XMLType_customUI14
            RelType = CustomUI14RelType
    End Select

    XmlRelsNamespace = "xmlns:rels='http://schemas.openxmlformats.org/package/2006/relationships'"
    Set oXMLDoc = New MSXML2.DOMDocument60
    oXMLDoc.SetProperty "SelectionNamespaces", XmlRelsNamespace
    oXMLDoc.Load RelationshipsFilePath
    CustomUIRelationshipExists = Not oXMLDoc.SelectSingleNode("//rels:Relationship[@Type='" & RelType & "']") Is Nothing
End Function

