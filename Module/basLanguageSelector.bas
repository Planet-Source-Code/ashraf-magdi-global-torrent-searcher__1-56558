Attribute VB_Name = "basLAnguageSelector"
Option Explicit

''I've found this sub in an Example in the Visual Basic Help
''and changed it to perform some more function
Public Sub LoadResStrings(frm As Form, strSourceFile As String)
On Error Resume Next
Dim ctl As Control
Dim obj As Object
Dim sCtlType As String
Dim nVal As Integer

    ''Set the Form's Caption
    frm.Caption = LoadString(CInt(frm.Tag), strSourceFile)

    ''Set the Caption of all the Controls
    For Each ctl In frm.Controls

        ''Get the Type of the Control...
        sCtlType = TypeName(ctl)

        ''... and set all the Caption by Type
        Select Case (sCtlType)
            Case "Label"
                ''For a Caption just look at the Tag
                ctl.Caption = LoadString(CInt(ctl.Tag), strSourceFile)

            Case "Menu"
                ''For a Menu we look at the Index
                ctl.Caption = LoadString(CInt(ctl.Index), strSourceFile)

            Case "TabStrip"
                For Each obj In ctl.Tabs
                    obj.Caption = LoadString(CInt(obj.Tag), strSourceFile)
                    ''For a ToolTipText we look at the Tag and the Index of the Control
                    obj.ToolTipText = LoadString(CInt(obj.Tag & obj.Index), strSourceFile)
                Next

            Case "Toolbar"
                For Each obj In ctl.Buttons
                    obj.ToolTipText = LoadString(CInt(obj.Tag & obj.Index), strSourceFile)
                    obj.Caption = LoadString(CInt(obj.Tag), strSourceFile)
                Next

            Case "ListView"
                For Each obj In ctl.ColumnHeaders
                    obj.Text = LoadString(CInt(obj.Tag), strSourceFile)
                Next

            Case "TextBox"
                ctl.Text = LoadString(CInt(ctl.Tag), strSourceFile)

            Case Else
                nVal = 0
                nVal = Val(ctl.Tag)
                If nVal > 0 Then ctl.Caption = LoadString(CInt(nVal), strSourceFile)
                nVal = 0
                nVal = Val(ctl.Tag & ctl.Index)
                If nVal > 0 Then ctl.ToolTipText = LoadString(CInt(nVal), strSourceFile)
        End Select
    Next
End Sub

''This is my own Function
''It load the Strings from a Language File (*.lan)
Public Function LoadString(lngIdent As Long, strSourceFile As String) As String
On Error GoTo ERRORE
''Declare the Variables
Dim strPath As String
Dim strTextLine As String
Dim lngSource As Long
Dim strString As String

    ''Set the Path of the Source File
    If Right(App.Path, 1) <> "\" Then
        strPath = App.Path & "\Lang\" & strSourceFile
    Else
        strPath = App.Path & strSourceFile
    End If

    ''Open the Source File
    Open strPath For Input As #1
        Do While Not EOF(1)
            ''Input a line from File
            Line Input #1, strTextLine
            ''Check if not EMPTY
            If strTextLine <> "" Then
                ''Get the Identificator from the File
                lngSource = CInt(Left(strTextLine, InStr(1, strTextLine, " ")))
                ''Search for the right Identificator
                If lngSource = lngIdent Then
                    ''Get the String
                    strString = Mid$(strTextLine, 6)
                    ''If the String is Empty put some '??????'
                    If strString = "" Then
                        LoadString = "??????"
                    Else
                        ''Else get the String
                        LoadString = strString
                    End If
                End If
            End If
        Loop
    Close #1
    Exit Function

ERRORE:

    If Err.Number = 53 Then MsgBox "File Not Found." & vbCrLf & _
        "Must Have at least one Language File!", _
        vbCritical + vbApplicationModal, "Errore While Loading"
    End
End Function

Public Sub SetLanguageFile(strFile As String)
''Declare the variables
Dim strPath As String
Dim strFound As String

    ''Set the path of the Application
    If Right(App.Path, 1) <> "\" Then
        strPath = App.Path & "\" & strFile
    Else
        strPath = App.Path & strFile
    End If

    ''Search for ALL avaibles Languages
    strFound = Dir(strPath)

    Do Until strFound = ""
        ''Add the language to the List and cut the ".lan"
        Main.lstLanguages.AddItem Mid$(strFound, 1, Len(strFound) - 4)
        strFound = Dir
    Loop
End Sub

