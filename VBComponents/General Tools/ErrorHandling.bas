Attribute VB_Name = "ErrorHandling"
'@Folder "General Tools"
'@IgnoreModule
Option Explicit

Private Const LINE_NO_TEXT As String = "Line no: "
Dim AlreadyUsed As Boolean

' ExcelMacroMastery.com Error handling code
' https://excelmacromastery.com/vba-error-handling/

' Example of using:
'
' 1. Place DisplayError in the topmost sub at the bottom. Replace the third paramter
'    with the name of the sub.
' DisplayError Err.source, Err.Description, "Module1.Topmost", Erl
'
' 2. Place RaiseError in all the other subs at the bottom of each. Replace the third paramter
'    with the name of the sub.
' RaiseError Err.Number, Err.source, "Module1.Level1", Err.Description, Erl
'
'
' 3. The error handling in each sub should look like this
'
'  Sub subName()
'
'    On Error Goto eh
'
'    The main code of the sub here!!!!!
'
'  done:
'      Exit Sub
'  eh:
'      DisplayError Err.Source, Err.Description, "Module1.Topmost", Erl
'  End Sub
'

' Reraises an error and adds line number and current procedure name
Sub RaiseError(ByVal errorNo As Long _
                , ByVal src As String _
                , ByVal proc As String _
                , ByVal desc As String _
                , ByVal lineNo As Long)

    Dim sSource As String

    ' If called for the first time then add line number
    If AlreadyUsed = False Then
        
        ' Add error line number if present
        If lineNo <> 0 Then
            sSource = vbNewLine & LINE_NO_TEXT & lineNo & " "
        End If

        ' Add procedure to source
        sSource = sSource & vbNewLine & proc
        AlreadyUsed = True
        
    Else
        ' If error has already been raised simply add on procedure name
        sSource = src & vbNewLine & proc
    End If
    
    ' Pause the code here when debugging
    '(To Debug: "Tools->VBA Properties" from the menu.
    ' Add "Debugging=1" to the     ' "Conditional Compilation Arguments.)
#If Debugging = 1 Then
    Debug.Assert False
#End If

    ' Reraise the error so it will be caught in the caller procedure
    ' (Note: If the code stops here, make sure DisplayError has been
    ' placed in the topmost procedure)
    Err.Raise errorNo, sSource, desc

End Sub

' Displays the error when it reaches the topmost sub
' Note: You can add a call to logging from this sub
Sub DisplayError(ByVal src As String, ByVal desc As String _
                    , ByVal sProcname As String, lineNo As Long)

    ' Check If the error happens in topmost sub
    If AlreadyUsed = False Then
        ' Reset string to remove "VBAProject" and add line number if it exists
        src = IIf(lineNo = 0, "", vbNewLine & LINE_NO_TEXT & lineNo)
    End If

    ' Build the final message
    Dim sMsg As String
    sMsg = "The following error occurred: " & vbNewLine & Err.Description _
                    & vbNewLine & vbNewLine & "Error Location is: "
    sMsg = sMsg & src & vbNewLine & sProcname

    ' Display the message
    MsgBox sMsg, Title:="Error"

    ' reset the boolean value
    AlreadyUsed = False

End Sub

