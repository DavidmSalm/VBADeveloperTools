Attribute VB_Name = "GitIntegration"
'@Folder("GIT")


Sub ExportCodeSelection(ByRef control As IRibbonControl)
    ExportCodeForm.Show
End Sub

Sub ImportCodeSelection(ByRef control As IRibbonControl)
    ImportCodeForm.Show
End Sub


Public Function ValidateRibbonUIXMLFile(ByVal FilePath As String) As Boolean
    Dim FSO                            As FileSystemObject
    Set FSO = New FileSystemObject
    Dim XML                            As String
    XML = FSO.OpenTextFile(FileName:=FilePath).ReadAll

    Dim MSXMLDocument                  As MSXML2.DOMDocument60
    Set MSXMLDocument = New MSXML2.DOMDocument60
    MSXMLDocument.LoadXML bstrxml:=XML
    If Not IsValidRibbonUIXML(XMLDoc:=MSXMLDocument) Then
        MsgBox Prompt:="XML for this file is not valid please refer to the immediate window for more details."
        ValidateRibbonUIXMLFile = False
    Else
        ValidateRibbonUIXMLFile = True
    End If
End Function

Private Function IsValidRibbonUIXML(ByVal XMLDoc As MSXML2.DOMDocument60) As Boolean

    Dim SchemaLocation                 As String: SchemaLocation = ThisWorkbook.Path & "\customui14.xsd"
    Const SchemaNamespace              As String = "http://schemas.microsoft.com/office/2009/07/cutomui"

    Dim SchemaCache                    As New xmlschemacache60


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

Private Function LoadXMLFile(SchemaLocation As String) As MSXML2.DOMDocument60

    Set LoadXMLFile = New MSXML2.DOMDocument60

    With LoadXMLFile
        .async = False
        .validateOnParse = False
        .resolveExternals = False
        .Load SchemaLocation
    End With

End Function
