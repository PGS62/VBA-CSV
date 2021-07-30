Attribute VB_Name = "modHelpGenerator"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CodeToRegister
' Author     : Philip Swannell
' Date       : 30-Jul-2021
' Purpose    : Generate VBA code to register a function.
' -----------------------------------------------------------------------------------------------------------------------
Function CodeToRegister(FunctionName, FunctionDescription As String, ArgDescriptions)

          Dim Code As String
          Const DQ = """"
          Dim i As Long
          
1         If TypeName(ArgDescriptions) = "Range" Then ArgDescriptions = ArgDescriptions.Value

2         On Error GoTo ErrHandler

3         If Len(FunctionDescription) > 255 Then Err.Raise vbObjectError + 1, , "FunctionDescription " + CStr(i) + " is of length " + CStr(Len(FunctionDescription)) + " but must be of length 255 or less."

4         Code = "Sub Register" + FunctionName + "()" + vbLf
5         Code = Code + "    Const FnDesc = " + DQ + Replace(FunctionDescription, DQ, DQ + DQ) + DQ + vbLf
6         Code = Code + "    Dim " + "ArgDescs() As String" + vbLf
7         Code = Code + "    ReDim " + "ArgDescs(" + CStr(LBound(ArgDescriptions, 1)) + " To " & CStr(UBound(ArgDescriptions, 1)) + ")" + vbLf

8         For i = LBound(ArgDescriptions, 1) To UBound(ArgDescriptions, 1)
9             If Len(ArgDescriptions(i, 1)) > 255 Then Err.Raise vbObjectError + 1, , "ArgDescriptions element " + CStr(i) + " is of length " + CStr(Len(ArgDescriptions(i, 1))) + " but must be of length 255 or less."
10            Code = Code + "    " + "ArgDescs(" & CStr(i) & ") = " & DQ + Replace(ArgDescriptions(i, 1), DQ, DQ + DQ) + DQ + vbLf
11        Next i

12        Code = Code + "    Application.MacroOptions """ + FunctionName + """, FnDesc, , , , , , , , , ArgDescs" + vbLf
13        Code = Code + "End Sub"

14        CodeToRegister = Application.WorksheetFunction.Transpose(VBA.Split(Code, vbLf))

15        Exit Function
ErrHandler:
16        CodeToRegister = "#CodeToRegister (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HelpForVBE
' Author     : Philip Swannell
' Date       : 30-Jul-2021
' Purpose    : Generate a header to paste into the VBE. The header generated will be consistent with the registration
'              created by calling CodeToRegister.
' -----------------------------------------------------------------------------------------------------------------------
Function HelpForVBE(FunctionName As String, FunctionDescription As String, Author As String, DateWritten As Long, ArgNames, ArgDescriptions, Optional ExtraHelp As String)
          Dim i As Long
          Dim NumArgs As Long
          Dim RowNum As Variant
          Dim Hlp As String
          
1         On Error GoTo ErrHandler

2         Hlp = Hlp & "'" & String(105, "-") & vbLf
3         Hlp = Hlp & "' Procedure : " & FunctionName & vbLf
4         Hlp = Hlp & "' Author    : Philip Swannell" & vbLf

5         Hlp = Hlp & "' Date      : "
6         Hlp = Hlp & Format$(DateWritten, "dd-mmm-yyyy") & vbLf

7         Hlp = Hlp & "' Purpose   :" & InsertBreaks(FunctionDescription) & vbLf
8         Hlp = Hlp & "' Arguments" & vbLf

9         NumArgs = sNRows(ArgNames)
10        For i = 1 To NumArgs
11            Hlp = Hlp & "' " & ArgNames(i, 1)
12            If Len(ArgNames(i, 1)) < 10 Then Hlp = Hlp & String(10 - Len(ArgNames(i, 1)), " ")
13            Hlp = Hlp & ":" & InsertBreaks(ArgDescriptions(i, 1)) + vbLf
14        Next
15        If Len(ExtraHelp) > 0 Then
16            Do While (Left$(ExtraHelp, 1) = vbLf Or Left$(ExtraHelp, 1) = vbCr)
17                ExtraHelp = Right$(ExtraHelp, Len(ExtraHelp) - 1)
18            Loop
19            Hlp = Hlp & ("'" & vbLf)
20            Hlp = Hlp & "' Notes     :"
21            Hlp = Hlp & InsertBreaks(ExtraHelp)
22            Hlp = Hlp & vbLf
23        End If
24        Hlp = Hlp & "'" & String(105, "-")
25        HelpForVBE = Application.WorksheetFunction.Transpose(VBA.Split(Hlp, vbLf))
26        Exit Function
ErrHandler:
27        HelpForVBE = "#HelpVBE (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


Function InsertBreaks(ByVal TheString As String)

          Const FirstTab = 0
          Const NextTabs = 13
          Const Width = 90
          Dim DoNewLine As Boolean
          Dim i As Long
          Dim LineLength As Long
          Dim Res As String
          Dim Words
          Dim WordsNLB
          
1         On Error GoTo ErrHandler
          
2         If InStr(TheString, " ") = 0 Then
3             InsertBreaks = TheString
4             Exit Function
5         End If
          
6         TheString = Replace(TheString, vbLf, vbLf + " ")
7         TheString = Replace(TheString, vbLf + "  ", vbLf + " ")
          
8         Res = String(FirstTab, " ")
9         LineLength = FirstTab

10        Words = VBA.Split(TheString, " ")
11        WordsNLB = Words
12        For i = LBound(Words) To UBound(Words)
13            WordsNLB(i) = Replace(WordsNLB(i), vbLf, vbNullString)
14        Next

15        For i = LBound(Words) To UBound(Words)
16            DoNewLine = LineLength + Len(WordsNLB(i)) > Width
17            If i > 1 Then
18                If InStr(Words(i - 1), vbLf) > 0 Then
19                    DoNewLine = True
20                End If
21            End If

22            If DoNewLine Then
23                Res = Res + vbLf + "'" + String(NextTabs, " ") + WordsNLB(i)
24                LineLength = 1 + NextTabs + Len(WordsNLB(i))
25            Else
26                Res = Res + " " + WordsNLB(i)
27                LineLength = LineLength + 1 + Len(WordsNLB(i))
28            End If
29        Next
30        InsertBreaks = Res

31        Exit Function
ErrHandler:
32        Err.Raise vbObjectError + 1, , "#InsertBreaks (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

