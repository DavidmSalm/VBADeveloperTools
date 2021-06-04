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
'@IgnoreModule ArgumentWithIncompatibleObjectType
#Const Debugging = True
Option Explicit

Private ImportVBProject                     As VBIDE.VBProject
Private WorkbookFullName                    As String
Private WorkbookName                        As String

Private ImportFolder                        As String
Private ImportVBComponentFolder             As String

Private WorkbookArchiveFullName             As String
Private ZipFullName                         As String
Private UnzippedFolderPath                  As String
Private UnzippedFolderCustomUIFullName      As String
Private UnzippedFolderCustomUI14FullName    As String
Private UnzippedFolderRelationshipFullName  As String

Private ImportCustomUIFullName              As String
Private ImportCustomUI14FullName            As String

Private Enum XMLType
    XMLTypeCustomUI = 1
    '@Ignore UseMeaningfulName
    XMLTypeCustomUI14 = 2
End Enum

Const CustomUIRelType                       As String = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
Const CustomUI14RelType                     As String = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

Private Sub PopulateGlobalStrings()
    Dim TimeStamp                           As String
    TimeStamp = Format$(Now, "yyyymmddhhmmss") & Right$(Format$(Timer, "#0.00"), 2)

    WorkbookFullName = ImportVBProject.fileName
    WorkbookName = Mid$(WorkbookFullName, InStrRev(WorkbookFullName, "\") + 1, Len(WorkbookFullName))

    WorkbookArchiveFullName = ThisWorkbook.Path & Application.PathSeparator & "zArchive" & Application.PathSeparator & WorkbookName & "." & TimeStamp

    ZipFullName = WorkbookFullName & TimeStamp & ".zip"
    UnzippedFolderPath = Mid$(WorkbookFullName, 1, InStrRev(WorkbookFullName, "\")) & WorkbookName & TimeStamp & "Unzipped"

    UnzippedFolderCustomUIFullName = UnzippedFolderPath & Application.PathSeparator & "customUI" & Application.PathSeparator & "CustomUI.xml"
    UnzippedFolderCustomUI14FullName = UnzippedFolderPath & Application.PathSeparator & "customUI" & Application.PathSeparator & "CustomUI14.xml"
    UnzippedFolderRelationshipFullName = UnzippedFolderPath & Application.PathSeparator & "_rels" & Application.PathSeparator & ".rels"

    ImportFolder = ThisWorkbook.Path & Application.PathSeparator & ImportVBProject.Name
    ImportVBComponentFolder = ImportFolder & Application.PathSeparator & "VBComponents"
    ImportCustomUIFullName = ImportFolder & "\XML\CustomUI.xml"
    ImportCustomUI14FullName = ImportFolder & "\XML\CustomUI14.xml"

    #If Debugging Then
        Debug.Print "Global Strings:"
        Debug.Print "TimeStamp                          ", TimeStamp
        Debug.Print "WorkbookFullName                   ", WorkbookFullName
        Debug.Print "WorkbookName                       ", WorkbookName
        Debug.Print "WorkbookArchiveFullName            ", WorkbookArchiveFullName
        Debug.Print "ZipFullName                        ", ZipFullName
        Debug.Print "UnzippedFolderPath                 ", UnzippedFolderPath
        Debug.Print "UnzippedFolderCustomUIFullName     ", UnzippedFolderCustomUIFullName
        Debug.Print "UnzippedFolderCustomUI14FullName   ", UnzippedFolderCustomUI14FullName
        Debug.Print "UnzippedFolderRelationshipFullName ", UnzippedFolderRelationshipFullName
        Debug.Print "ImportFolder                       ", ImportFolder
        Debug.Print "ImportVBComponentFolder            ", ImportVBComponentFolder
        Debug.Print "ImportCustomUIFullName             ", ImportCustomUIFullName
        Debug.Print "ImportCustomUI14FullName           ", ImportCustomUI14FullName
    #End If
End Sub

Private Sub UserForm_Initialize()

    For Each ImportVBProject In Application.VBE.VBProjects
        Debug.Print ImportVBProject.Name
        If ImportVBProject.Name <> "VBAProject" Then AddinSelection.AddItem pvargItem:=ImportVBProject.Name
    Next ImportVBProject
    AddinSelection.AddItem pvargItem:="End of Add-ins"
    AddinSelection.ListIndex = 0

End Sub

Private Sub OkButton_Click()

    Set ImportVBProject = Application.VBE.VBProjects.Item(AddinSelection.Value)

    PopulateGlobalStrings
    If Not FoldersAndFiles.FolderExists(strDataFolder:=ImportVBComponentFolder) Then
        ImportFolder = FoldersAndFiles.GetUserSelectedPath(DefaultPath:=ImportFolder, FileType:=msoFileDialogFolderPicker)
    End If

    If Right$(ImportFolder, Len(ImportFolder) - InStrRev(ImportFolder, "\", , vbTextCompare)) <> ImportVBProject.Name Then
        If MsgBox(prompt:="Folder name and VB project are not the same. Are you sure you want to import the code into this project?", Buttons:=vbYesNo) = vbNo Then
            Debug.Print "Process stoped by the user."
            Exit Sub
        End If
    End If


    FoldersAndFiles.FileCreateCopy Source:=WorkbookFullName, Destination:=WorkbookArchiveFullName 'Create Backup

    IfThisFileThenCreateTemp

    UpdateSelectedVBProjectWithFileComponents
    Application.Workbooks(WorkbookName).Save
    If FoldersAndFiles.FileExists(strFileName:=ImportCustomUIFullName) Or FoldersAndFiles.FileExists(strFileName:=ImportCustomUI14FullName) Then UpdateXML

    DeleteFilesandFolders
    If FoldersAndFiles.FileGetExtension(FilePath:=ThisWorkbook.Name) = "tmp" Then DeleteThisWorkbook
    Unload Me
End Sub

Private Sub IfThisFileThenCreateTemp()

    If ImportVBProject.fileName = ThisWorkbook.FullName Then
        Dim ThisWorkbookFullName            As String: ThisWorkbookFullName = ThisWorkbook.FullName
        ThisWorkbook.SaveAs fileName:=ThisWorkbook.Path & Application.PathSeparator & CreateObject("Scripting.FileSystemObject").GetTempName()
        Dim CurrentWorkbook                 As Workbook
        Set CurrentWorkbook = Workbooks.Open(fileName:=ThisWorkbookFullName)
        Set ImportVBProject = CurrentWorkbook.VBProject
    End If
End Sub

Private Sub DeleteFilesandFolders()
    Kill PathName:=ZipFullName
    FoldersAndFiles.FolderDelete FolderPath:=UnzippedFolderPath
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub DeleteThisWorkbook()

    With ThisWorkbook
        .Saved = True
        .ChangeFileAccess xlReadOnly
        Kill .FullName
        .Close False
    End With

End Sub

Private Sub UpdateSelectedVBProjectWithFileComponents()
    DeleteVBAModulesandUserForms
    LoopThrougFolderandImportCode FolderPath:=ImportFolder
End Sub

Private Sub DeleteVBAModulesandUserForms()
    Dim CurrentVBComponent                  As VBIDE.VBComponent

    For Each CurrentVBComponent In ImportVBProject.VBComponents
        If CurrentVBComponent.Type <> vbext_ct_Document Then ImportVBProject.VBComponents.Remove CurrentVBComponent
    Next CurrentVBComponent

    For Each CurrentVBComponent In ImportVBProject.VBComponents
        If CurrentVBComponent.Type <> vbext_ct_Document Then
            Debug.Print "Not all VB Components were deleted. Please review and rerun."
            Debug.Assert False
        End If
    Next CurrentVBComponent


End Sub

Private Sub LoopThrougFolderandImportCode(ByVal FolderPath As String)
    Dim FSO                                 As FileSystemObject
    Set FSO = New FileSystemObject
    Dim Folder                              As Folder
    Set Folder = FSO.GetFolder(FolderPath:=FolderPath)

    Dim SubFolder                           As Folder

    For Each SubFolder In Folder.SubFolders
        LoopThrougFolderandImportCode FolderPath:=SubFolder.Path
    Next SubFolder

    Dim CurrentFile                         As File

    For Each CurrentFile In Folder.Files
        If Right$(CurrentFile.Path, 4) = ".bas" Or Right$(CurrentFile.Path, 4) = ".cls" Or Right$(CurrentFile.Path, 4) = ".frm" Then
            Dim CurrentModule As VBIDE.VBComponent
            Set CurrentModule = ImportVBProject.VBComponents.Import(fileName:=CurrentFile.Path)
            RemoveBlankLinesAtBeginningofVBComponent CurrentModule:=CurrentModule
            Debug.Print "Imported: ", CurrentFile.Path
        End If
    Next CurrentFile
    
End Sub


Private Sub RemoveBlankLinesAtBeginningofVBComponent(CurrentModule As VBIDE.VBComponent)
    
    With CurrentModule.CodeModule
        Dim LineofCode                      As Long
        For LineofCode = 1 To .CountOfLines
            Debug.Print .Lines(LineofCode, 1)
            If Not .Lines(LineofCode, 1) = vbNullString Then Exit For
        Next LineofCode
        If LineofCode <> 1 Then .DeleteLines StartLine:=1, Count:=LineofCode - 1
    End With
    

End Sub



Private Sub UpdateXML()


    FoldersAndFiles.FileCreateCopy Source:=WorkbookFullName, Destination:=ZipFullName
    FoldersAndFiles.FolderUnzip FolderPath:=ZipFullName, UnzipFolderPath:=UnzippedFolderPath

    If FoldersAndFiles.FileExists(strFileName:=ImportCustomUIFullName) Then
        AddRelationshipifNeeded CurrentXMLType:=XMLTypeCustomUI
        FoldersAndFiles.FileCreateCopy Source:=ImportCustomUIFullName, Destination:=UnzippedFolderCustomUIFullName
    End If
    If FoldersAndFiles.FileExists(strFileName:=ImportCustomUI14FullName) Then
        AddRelationshipifNeeded CurrentXMLType:=XMLTypeCustomUI14
        FoldersAndFiles.FileCreateCopy Source:=ImportCustomUI14FullName, Destination:=UnzippedFolderCustomUI14FullName
    End If

    Kill PathName:=ZipFullName
    FoldersAndFiles.FolderZip FolderPathSource:=UnzippedFolderPath, ZipPathDestination:=ZipFullName

    Workbooks.Open(fileName:=WorkbookFullName).Close
    Kill PathName:=WorkbookFullName
    FoldersAndFiles.FileCreateCopy Source:=ZipFullName, Destination:=WorkbookFullName
    Workbooks.Open fileName:=WorkbookFullName

End Sub

Private Sub AddRelationshipifNeeded(ByVal CurrentXMLType As XMLType)
    If Not CustomUIRelationshipExists(RelationshipsFilePath:=UnzippedFolderRelationshipFullName, RelationshipXMLType:=CurrentXMLType) Then
        AddCustomUIToRels RelationshipsFilePath:=UnzippedFolderRelationshipFullName, RelationshipXMLType:=CurrentXMLType
    End If

End Sub

Private Sub SetUpWorkbookforRibbonUI(ByVal UnzippedWorkbookFolder As String)

    Dim RelationshipsFilePath               As String
    RelationshipsFilePath = UnzippedWorkbookFolder & "\_rels\.rels"


    If Not CustomUIRelationshipExists(RelationshipsFilePath:=RelationshipsFilePath, RelationshipXMLType:=XMLTypeCustomUI14) Then
        AddCustomUIToRels RelationshipsFilePath:=RelationshipsFilePath, RelationshipXMLType:=XMLTypeCustomUI14
    End If

    '    Dim RelationshipsFileContent       As String
    '    RelationshipsFileContent = FoldersAndFiles.TextFileGetContent(FilePath:=RelationshipsFilePath)
    '    Dim XMLRequiredRelationship        As String
    '    XMLRequiredRelationship = "<Relationship Id=""cuID14"" Type=""http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"" Target=""customUI/customUI14.xml"" />"
    '
    '    If InStr(1, RelationshipsFileContent, XMLRequiredRelationship, vbTextCompare) = 0 Then FoldersAndFiles.TextFileFindReplace FilePath:=RelationshipsFilePath, FindString:="</Relationships>", ReplaceString:=XMLRequiredRelationship & "</Relationships>"

End Sub

Private Sub AddCustomUIToRels(ByVal RelationshipsFilePath As String, ByVal RelationshipXMLType As XMLType)
    'Date Created : 5/14/2009 23:29
    'Author       : Ken Puls (www.excelguru.ca)
    'Modified by  : Doug Glancy 11/20/2017
    'Macro Purpose: Add the customUI relationship to the rels file

    Dim XMLDoc                              As MSXML2.DOMDocument60
    Dim XMLElement                          As MSXML2.IXMLDOMNode
    Dim XMLAttrib                           As MSXML2.IXMLDOMAttribute
    Dim NamedNodeMap                        As MSXML2.IXMLDOMNamedNodeMap

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
        Case XMLTypeCustomUI
            XMLAttrib.NodeValue = "cuID"
        Case XMLTypeCustomUI14
            XMLAttrib.NodeValue = "cuID14"
    End Select

    NamedNodeMap.setNamedItem XMLAttrib

    'Create Type attribute for the element
    Set XMLAttrib = XMLDoc.createAttribute("Type")
    Select Case RelationshipXMLType
        Case XMLTypeCustomUI
            XMLAttrib.NodeValue = CustomUIRelType
        Case XMLTypeCustomUI14
            XMLAttrib.NodeValue = CustomUI14RelType
    End Select
    NamedNodeMap.setNamedItem XMLAttrib

    'Create Target element for the attribute
    Set XMLAttrib = XMLDoc.createAttribute("Target")
    Select Case RelationshipXMLType
        Case XMLTypeCustomUI
            XMLAttrib.NodeValue = "customUI/customUI.xml"
        Case XMLTypeCustomUI14
            XMLAttrib.NodeValue = "customUI/customUI14.xml"
    End Select
    NamedNodeMap.setNamedItem XMLAttrib

    'Now insert the new node at the proper location
    XMLDoc.ChildNodes.Item(1).appendChild XMLElement
    'Save the .rels file
    XMLDoc.Save RelationshipsFilePath
    Set XMLAttrib = Nothing
    Set XMLElement = Nothing
    Set XMLDoc = Nothing
End Sub

Private Function CustomUIRelationshipExists(ByVal RelationshipsFilePath As String, ByVal RelationshipXMLType As XMLType) As Boolean
    '@Ignore HungarianNotation
    Dim oXMLDoc                             As MSXML2.DOMDocument60
    Dim XmlRelsNamespace                    As String
    Dim RelType                             As String

    Select Case RelationshipXMLType
        Case XMLType.XMLTypeCustomUI
            RelType = CustomUIRelType
        Case XMLType.XMLTypeCustomUI14
            RelType = CustomUI14RelType
    End Select

    XmlRelsNamespace = "xmlns:rels='http://schemas.openxmlformats.org/package/2006/relationships'"
    Set oXMLDoc = New MSXML2.DOMDocument60
    oXMLDoc.SetProperty "SelectionNamespaces", XmlRelsNamespace
    oXMLDoc.Load RelationshipsFilePath
    CustomUIRelationshipExists = Not oXMLDoc.SelectSingleNode("//rels:Relationship[@Type='" & RelType & "']") Is Nothing
End Function


