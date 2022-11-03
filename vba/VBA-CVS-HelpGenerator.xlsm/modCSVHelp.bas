Attribute VB_Name = "modCSVHelp"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CodeToRegister
' Purpose    : Generate VBA code to register a function.
' -----------------------------------------------------------------------------------------------------------------------
Function CodeToRegister(FunctionName, Description As String, ArgDescs)

    Const DQ = """"
    Dim code As String
    Dim i As Long
    
    If TypeName(ArgDescs) = "Range" Then ArgDescs = ArgDescs.value

    On Error GoTo ErrHandler

    If Len(Description) > 255 Then Throw "Description " & CStr(i) & " is of length " & CStr(Len(Description)) & " but must be of length 255 or less."
    
    code = code & "' " & String(119, "-") & vbLf
    code = code & "' Procedure  : Register" & FunctionName & vbLf
    code = code & "' Purpose    : Register the function " & FunctionName & " with the Excel function wizard. Suggest this function is called from a" & vbLf
    code = code & "'              WorkBook_Open event." & vbLf
    code = code & "' " & String(119, "-") & vbLf

    code = code & "Public Sub Register" & FunctionName & "()" & vbLf
    code = code & "    Const Description As String = " & InsertBreaksInStringLiteral(DQ & Replace(Description, DQ, DQ & DQ) & DQ, 34) & vbLf
    code = code & "    Dim " & "ArgDescs() As String" & vbLf & vbLf
    code = code & "    On Error GoTo ErrHandler" & vbLf & vbLf
    
    code = code & "    ReDim " & "ArgDescs(" & CStr(LBound(ArgDescs, 1)) & " To " & CStr(UBound(ArgDescs, 1)) & ")" & vbLf

    For i = LBound(ArgDescs, 1) To UBound(ArgDescs, 1)
        If Len(ArgDescs(i, 1)) > 255 Then Throw "ArgDescs element " & CStr(i) & " is of length " & CStr(Len(ArgDescs(i, 1))) & " but must be of length 255 or less."
        code = code & "    " & "ArgDescs(" & CStr(i) & ") = " & InsertBreaksInStringLiteral(DQ & Replace(ArgDescs(i, 1), DQ, DQ & DQ) & DQ, IIf(i < 10, 18, 19)) & vbLf
    Next i

    code = code & "    Application.MacroOptions """ & FunctionName & """, Description, , , , , , , , , ArgDescs" & vbLf
    
    code = code & "    Exit Sub" & vbLf & vbLf
    
    code = code & "ErrHandler:" & vbLf
    code = code & "    Debug.Print ""Warning: Registration of function " & FunctionName & " failed with error: "" & Err.Description" & vbLf
    code = code & "End Sub"

    CodeToRegister = Application.WorksheetFunction.Transpose(VBA.Split(code, vbLf))

    Exit Function
ErrHandler:
    CodeToRegister = "#CodeToRegister: " & Err.Description & "!"
End Function

Function InsertBreaksInStringLiteral(ByVal TheString As String, Optional FirstRowShorterBy As Long)

    Const FirstTab = 0
    Dim NextTabs As Long
    Const Width = 114
    Dim DoNewLine As Boolean
    Dim i As Long
    Dim LineLength As Long
    Dim Res As String
    Dim Words
    Dim WordsNLB
    
    On Error GoTo ErrHandler
    
    NextTabs = FirstRowShorterBy
    
    If InStr(TheString, " ") = 0 Then
        InsertBreaksInStringLiteral = TheString
        Exit Function
    End If
    
    Res = String(FirstTab, " ")
    LineLength = FirstTab + FirstRowShorterBy

    Words = VBA.Split(TheString, " ")
    WordsNLB = Words

    For i = LBound(Words) To UBound(Words)
        DoNewLine = LineLength + Len(WordsNLB(i)) > Width

        If DoNewLine Then
            Res = Res + " "" & _" + vbLf + String(NextTabs, " ") + """" + WordsNLB(i)
            LineLength = 1 + NextTabs + Len(WordsNLB(i))
        Else
            Res = Res + " " + WordsNLB(i)
            LineLength = LineLength + 1 + Len(WordsNLB(i))
        End If
    Next
    InsertBreaksInStringLiteral = Trim(Res)

    Exit Function
ErrHandler:
    Throw "#InsertBreaksInStringLiteral: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HelpForVBE
' Purpose    : Generate a header to paste into the VBE. The header generated will be consistent with the registration
'              created by calling CodeToRegister.
' -----------------------------------------------------------------------------------------------------------------------
Function HelpForVBE(FunctionName As String, FunctionDescription As String, ArgNames, ArgDescriptions, _
        Optional ExtraHelp As String, Optional Author As String, Optional DateWritten As Long)

          Dim Hlp As String
          Dim i As Long
          Dim NumArgs As Long
          Dim Spacers As String
          
1         On Error GoTo ErrHandler
          
2         Hlp = Hlp & "' " & String(119, "-") & vbLf
3         Hlp = Hlp & "' Procedure : " & FunctionName & vbLf
4         If Len(Author) > 0 Then
5             Hlp = Hlp & "' Author    : " & Author & "" & vbLf
6         End If
7         If DateWritten <> 0 Then
8             Hlp = Hlp & "' Date      : " & Format$(DateWritten, "dd-mmm-yyyy") & vbLf
9         End If

10        Hlp = Hlp & "' Purpose   :" & InsertBreaks(FunctionDescription, Len("Len(ArgNames(i, 1))")) & vbLf
11        Hlp = Hlp & "' Arguments" & vbLf

12        If TypeName(ArgNames) = "Range" Then ArgNames = ArgNames.value

13        NumArgs = UBound(ArgNames, 1) - LBound(ArgNames, 1) + 1
14        For i = 1 To NumArgs
              '    If InStr(ArgNames(i, 1), "EOL") > 0 Then Stop
15            Hlp = Hlp & "' " & ArgNames(i, 1)
16            If Len(ArgNames(i, 1)) < 10 Then
17                Spacers = String(10 - Len(ArgNames(i, 1)), " ")
18            Else
19                Spacers = ""
20            End If
              
21            Hlp = Hlp & Spacers
22            Hlp = Hlp & ":" & InsertBreaks(ArgDescriptions(i, 1), Len(ArgNames(i, 1)) + Len(Spacers) + 2) + vbLf
23        Next
24        If Len(ExtraHelp) > 0 Then
25            Do While (Left$(ExtraHelp, 1) = vbLf Or Left$(ExtraHelp, 1) = vbCr)
26                ExtraHelp = Right$(ExtraHelp, Len(ExtraHelp) - 1)
27            Loop
28            Hlp = Hlp & ("'" & vbLf)
29            Hlp = Hlp & "' Notes     :"
30            Hlp = Hlp & InsertBreaks(ExtraHelp)
31            Hlp = Hlp & vbLf
32        End If
33        Hlp = Hlp & "' " & String(119, "-")
34        HelpForVBE = Application.WorksheetFunction.Transpose(VBA.Split(Hlp, vbLf))
35        Exit Function
ErrHandler:
36        HelpForVBE = "#HelpVBE: " & Err.Description & "!"
End Function

Function InsertBreaks(ByVal TheString As String, Optional FirstRowShorterBy As Long)

        Const FirstTab = 0
        Const NextTabs = 13
        Const Width = 106
        Dim DoNewLine As Boolean
        Dim i As Long
        Dim LineLength As Long
        Dim Res As String
        Dim Words
        Dim WordsNLB
    
        On Error GoTo ErrHandler
    
        If InStr(TheString, " ") = 0 Then
              InsertBreaks = TheString
              Exit Function
        End If
    
        TheString = Replace(TheString, vbLf, vbLf + " ")
        TheString = Replace(TheString, vbLf + "  ", vbLf + " ")
    
        Res = String(FirstTab, " ")
        LineLength = FirstTab + FirstRowShorterBy

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
                  Res = Res + vbLf + "'" + String(NextTabs, " ") + WordsNLB(i)
                  LineLength = 1 + NextTabs + Len(WordsNLB(i))
              Else
                  Res = Res + " " + WordsNLB(i)
                  LineLength = LineLength + 1 + Len(WordsNLB(i))
              End If
        Next
        InsertBreaks = Res

        Exit Function
ErrHandler:
        Throw "#InsertBreaks: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MarkdownHelp
' Purpose    : Formats the help as a markdown table.
' -----------------------------------------------------------------------------------------------------------------------
Function MarkdownHelp(SourceFile As String, FunctionName As String, ByVal FunctionDescription As String, _
        ArgNames, ArgDescriptions, Optional Replacements)

          Dim Declaration As String
          Dim Hlp As String
          Dim i As Long
          Dim j As Long
          Dim LeftString As String
          Dim RightString As String
          Dim SourceCode As String
          Dim StringsToEncloseInBackTicks
          Dim ThisArgDescription As String

1         On Error GoTo ErrHandler
2         SourceCode = RawFileContents(SourceFile)
3         SourceCode = Replace(SourceCode, vbCrLf, vbLf)

4         LeftString = "Public Function " & FunctionName & "("

5         If InStr(SourceCode, LeftString) = 0 Then
6             LeftString = "Private Function " & FunctionName & "("
7             If InStr(SourceCode, LeftString) = 0 Then
8                 LeftString = "Function " & FunctionName & "("
9                 If InStr(SourceCode, LeftString) = 0 Then Throw "Cannot find function declaration in SourceFile"
10            End If
11        End If

12        Declaration = StringBetweenStrings(SourceCode, LeftString, ")", True, True)

          'Bodge - get the "As VarType"
          Dim matchPoint As Long
          Dim NextChars As String
13        matchPoint = InStr(SourceCode, Declaration)
14        NextChars = Mid$(SourceCode, matchPoint + Len(Declaration), 100)
15        If Left$(NextChars, 4) = " As " Then
16            NextChars = StringBetweenStrings(NextChars, " As ", vbLf, True, False)
17            NextChars = " " & Trim(NextChars)
18            Declaration = Declaration & NextChars
19        End If

20        Hlp = "#### _" & FunctionName & "_" & vbLf

21        StringsToEncloseInBackTicks = VStack(ArgNames, "CSVRead", "CSVWrite", "#", "!")

22        For j = 1 To sNRows(StringsToEncloseInBackTicks)
23            FunctionDescription = sRegExReplace(FunctionDescription, "\b" & StringsToEncloseInBackTicks(j, 1) & "\b", "`" & StringsToEncloseInBackTicks(j, 1) & "`", True)
24        Next j

25        Hlp = Hlp & FunctionDescription & vbLf
          
26        Hlp = Hlp & "```vba" & vbLf & _
              Declaration & vbLf & _
              "```" & vbLf & vbLf & _
              "|Argument|Description|" & vbLf & _
              "|:-------|:----------|"
          
27        For i = 1 To sNRows(ArgNames)
28            ThisArgDescription = ArgDescriptions(i, 1)
29            For j = 1 To sNRows(StringsToEncloseInBackTicks)
30                ThisArgDescription = sRegExReplace(ThisArgDescription, "\b" & StringsToEncloseInBackTicks(j, 1) & "\b", "`" & StringsToEncloseInBackTicks(j, 1) & "`", True)
31            Next j
32            ThisArgDescription = Replace(ThisArgDescription, vbCrLf, vbLf)
33            ThisArgDescription = Replace(ThisArgDescription, vbCr, vbLf)
34            ThisArgDescription = Replace(ThisArgDescription, vbLf, "<br/>")
35            Hlp = Hlp & vbLf & "|`" & ArgNames(i, 1) & "`|" & ThisArgDescription & "|"
36        Next i

37        If Not IsMissing(Replacements) Then
38            If TypeName(Replacements) = "Range" Then
39                Replacements = Replacements.value
40            End If
41            For i = 1 To sNRows(Replacements)
42                Hlp = Replace(Hlp, Replacements(i, 1), Replacements(i, 2))
43            Next i
44        End If

45        MarkdownHelp = Application.Transpose(VBA.Split(Hlp, vbLf))

46        Exit Function
ErrHandler:
47        MarkdownHelp = "#MarkdownHelp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function RawFileContents(FileName As String)
    Dim F As Scripting.File
    Dim FSO As New FileSystemObject
    Dim T As Scripting.TextStream
    On Error GoTo ErrHandler
    Set F = FSO.GetFile(FileName)
    Set T = F.OpenAsTextStream()
    RawFileContents = T.ReadAll
    T.Close

    Exit Function
ErrHandler:
   ' Throw "#RawFileContents (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

