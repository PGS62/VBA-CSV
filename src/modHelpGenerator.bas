Attribute VB_Name = "modHelpGenerator"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CodeToRegister
' Purpose    : Generate VBA code to register a function.
' -----------------------------------------------------------------------------------------------------------------------
Function CodeToRegister(FunctionName, FunctionDescription As String, ArgDescriptions)

    Dim code As String
    Const DQ = """"
    Dim i As Long
    
    If TypeName(ArgDescriptions) = "Range" Then ArgDescriptions = ArgDescriptions.value

    On Error GoTo ErrHandler

    If Len(FunctionDescription) > 255 Then Throw "FunctionDescription " + CStr(i) + " is of length " + CStr(Len(FunctionDescription)) + " but must be of length 255 or less."

    code = "Sub Register" + FunctionName + "()" + vbLf
    code = code + "    Const FnDesc = " + DQ + Replace(FunctionDescription, DQ, DQ + DQ) + DQ + vbLf
    code = code + "    Dim " + "ArgDescs() As String" + vbLf
    code = code + "    ReDim " + "ArgDescs(" + CStr(LBound(ArgDescriptions, 1)) + " To " & CStr(UBound(ArgDescriptions, 1)) + ")" + vbLf

    For i = LBound(ArgDescriptions, 1) To UBound(ArgDescriptions, 1)
        If Len(ArgDescriptions(i, 1)) > 255 Then Throw "ArgDescriptions element " + CStr(i) + " is of length " + CStr(Len(ArgDescriptions(i, 1))) + " but must be of length 255 or less."
        code = code + "    " + "ArgDescs(" & CStr(i) & ") = " & DQ + Replace(ArgDescriptions(i, 1), DQ, DQ + DQ) + DQ + vbLf
    Next i

    code = code + "    Application.MacroOptions """ + FunctionName + """, FnDesc, , , , , , , , , ArgDescs" + vbLf
    code = code + "End Sub"

    CodeToRegister = Application.WorksheetFunction.Transpose(VBA.Split(code, vbLf))

    Exit Function
ErrHandler:
    CodeToRegister = "#CodeToRegister: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HelpForVBE
' Purpose    : Generate a header to paste into the VBE. The header generated will be consistent with the registration
'              created by calling CodeToRegister.
' -----------------------------------------------------------------------------------------------------------------------
Function HelpForVBE(FunctionName As String, FunctionDescription As String, ArgNames, ArgDescriptions, Optional ExtraHelp As String, Optional Author As String, Optional DateWritten As Long)
    Dim i As Long
    Dim NumArgs As Long
    Dim Hlp As String
    
    On Error GoTo ErrHandler

    Hlp = Hlp & "'" & String(105, "-") & vbLf
    Hlp = Hlp & "' Procedure : " & FunctionName & vbLf
    If Len(Author) > 0 Then
        Hlp = Hlp & "' Author    : " & Author & "" & vbLf
    End If
    If DateWritten <> 0 Then
        Hlp = Hlp & "' Date      : " & Format$(DateWritten, "dd-mmm-yyyy") & vbLf
    End If

    Hlp = Hlp & "' Purpose   :" & InsertBreaks(FunctionDescription) & vbLf
    Hlp = Hlp & "' Arguments" & vbLf

    NumArgs = sNRows(ArgNames)
    For i = 1 To NumArgs
        Hlp = Hlp & "' " & ArgNames(i, 1)
        If Len(ArgNames(i, 1)) < 10 Then Hlp = Hlp & String(10 - Len(ArgNames(i, 1)), " ")
        Hlp = Hlp & ":" & InsertBreaks(ArgDescriptions(i, 1)) + vbLf
    Next
    If Len(ExtraHelp) > 0 Then
        Do While (Left$(ExtraHelp, 1) = vbLf Or Left$(ExtraHelp, 1) = vbCr)
            ExtraHelp = Right$(ExtraHelp, Len(ExtraHelp) - 1)
        Loop
        Hlp = Hlp & ("'" & vbLf)
        Hlp = Hlp & "' Notes     :"
        Hlp = Hlp & InsertBreaks(ExtraHelp)
        Hlp = Hlp & vbLf
    End If
    Hlp = Hlp & "'" & String(105, "-")
    HelpForVBE = Application.WorksheetFunction.Transpose(VBA.Split(Hlp, vbLf))
    Exit Function
ErrHandler:
    HelpForVBE = "#HelpVBE: " & Err.Description & "!"
End Function

Function InsertBreaks(ByVal TheString As String)

    Const FirstTab = 0
    Const NextTabs = 13
    Const Width = 90
    Dim DoNewLine As Boolean
    Dim i As Long
    Dim LineLength As Long
    Dim res As String
    Dim Words
    Dim WordsNLB
    
    On Error GoTo ErrHandler
    
    If InStr(TheString, " ") = 0 Then
        InsertBreaks = TheString
        Exit Function
    End If
    
    TheString = Replace(TheString, vbLf, vbLf + " ")
    TheString = Replace(TheString, vbLf + "  ", vbLf + " ")
    
    res = String(FirstTab, " ")
    LineLength = FirstTab

    Words = VBA.Split(TheString, " ")
    WordsNLB = Words
    For i = LBound(Words) To UBound(Words)
        WordsNLB(i) = Replace(WordsNLB(i), vbLf, vbNullString)
    Next

    For i = LBound(Words) To UBound(Words)
        DoNewLine = LineLength + Len(WordsNLB(i)) > Width
        If i > 1 Then
            If InStr(Words(i - 1), vbLf) > 0 Then
                DoNewLine = True
            End If
        End If

        If DoNewLine Then
            res = res + vbLf + "'" + String(NextTabs, " ") + WordsNLB(i)
            LineLength = 1 + NextTabs + Len(WordsNLB(i))
        Else
            res = res + " " + WordsNLB(i)
            LineLength = LineLength + 1 + Len(WordsNLB(i))
        End If
    Next
    InsertBreaks = res

    Exit Function
ErrHandler:
    Throw "#InsertBreaks: " & Err.Description & "!"
End Function

