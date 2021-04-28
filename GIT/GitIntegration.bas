Attribute VB_Name = "GitIntegration"
Option Explicit
'@Folder("GIT")


Public Sub ExportCodeSelection()
    ExportCodeForm.Show
End Sub

Public Sub ImportCodeSelection()
    ImportCodeForm.Show
End Sub


Public Sub UpdateXML()

    ThisWorkbook.SaveCopyAs fileName:=ThisWorkbook.FullName & ".zip"
    Unzip FilePath:=ThisWorkbook.FullName & ".zip"

End Sub


Public Function ValidateRibbonUIXMLFile(ByVal FilePath As String) As Boolean
    Dim FSO                            As FileSystemObject
    Set FSO = New FileSystemObject
    Dim RibbonUIXML                            As String
    RibbonUIXML = FSO.OpenTextFile(fileName:=FilePath).ReadAll

    Dim MSXMLDocument                  As MSXML2.DOMDocument60
    Set MSXMLDocument = New MSXML2.DOMDocument60
    MSXMLDocument.LoadXML bstrxml:=RibbonUIXML
    If Not IsValidRibbonUIXML(XMLDoc:=MSXMLDocument) Then
        MsgBox prompt:="XML for this file is not valid please refer to the immediate window for more details."
        ValidateRibbonUIXMLFile = False
    Else
        ValidateRibbonUIXMLFile = True
    End If
End Function

Private Function IsValidRibbonUIXML(ByVal XMLDoc As MSXML2.DOMDocument60) As Boolean

    Dim SchemaLocation                 As String: SchemaLocation = ThisWorkbook.Path & "\customui14.xsd"
    Const SchemaNamespace              As String = "http://schemas.microsoft.com/office/2009/07/cutomui"

    Dim SchemaCache                    As xmlschemacache60: Set SchemaCache = New xmlschemacache60


    SchemaCache.Add NamespaceURI:=SchemaNamespace, Var:=LoadXMLFile(SchemaLocation:=SchemaLocation)

    With XMLDoc
        Set .Schemas = SchemaCache
        .SetProperty Name:="MultipleErrorMessages", Value:=True
        Dim ParseError                 As MSXML2.IXMLDOMParseError2
        Set ParseError = .Validate()
    End With
    If ParseError.allErrors.Length = 0 Then
        IsValidRibbonUIXML = True
    Else
        Debug.Print ParseError.allErrors.Length
        Dim CurrentError               As Object
        For Each CurrentError In ParseError.allErrors
            Debug.Print "Parse error: ", CurrentError.ErrorCode, CurrentError.reason
        Next CurrentError
    End If

End Function

Private Function LoadXMLFile(ByVal SchemaLocation As String) As MSXML2.DOMDocument60

    Set LoadXMLFile = New MSXML2.DOMDocument60

    With LoadXMLFile
        .async = False
        .validateOnParse = False
        .resolveExternals = False
        .Load SchemaLocation
    End With

End Function

