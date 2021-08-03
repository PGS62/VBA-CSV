Attribute VB_Name = "modSplitNew"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SplitNew
' Author     : Philip Swannell
' Date       : 02-Aug-2021
' Purpose    : Drop-in replacement for VBA.Split, but rather than splitting Expression at all occurrences of Delimiter,
'              splits occur at only those instances that are preceded by an even number of quote characters (DQ).
 'Notes
 '            1) DQ is assumed to have one character and defaults to the double-quote character.
 '            2) DQ must not be contained in the Delimiter string, which is allowe to have more than one character.
 '            3) AltDelim, should be either omitted or passed as a string which is 1) the same length as Delimiter;
 '               and b) not a sub-string of Expression. If AltDelim is omitted and if both DQ and Delimiter are
 '               sub-strings of Expression then AltDelim is set to a string satisfying conditions a) and b)
' -----------------------------------------------------------------------------------------------------------------------
Function SplitNew(ByVal Expression As String, Optional Delimiter As String = vbCrLf, Optional DQ As String = """", Optional Limit As Long = -1, Optional ByRef AltDelim As String)

          'DQ assumed to have length 1
          'Delimiter may have length > 1
          'Also assume DQ is not a character of Delimiter

          Dim Ret() As String
          Dim DQPos As Long
          Dim DelimPos As Long
          Dim LDelim As Long
          Dim EvenDQs As Boolean
          Dim NDelims As Long

1         On Error GoTo ErrHandler
2         If Len(Expression) = 0 Then
3             ReDim Ret(0 To 0)
4             SplitNew = Ret
5             Exit Function
6         End If

7         DQPos = 0
8         LDelim = Len(Delimiter)
9         DelimPos = 1 - LDelim

10        DQPos = InStr(DQPos + 1, Expression, DQ)
11        If DQPos = 0 Then
12            SplitNew = VBA.Split(Expression, Delimiter, Limit)
13            Exit Function
14        End If

15        DelimPos = InStr(DelimPos + LDelim, Expression, Delimiter)
16        If DelimPos = 0 Then
17            ReDim Ret(0 To 0)
18            Ret(0) = Expression
19            SplitNew = Ret
20            Exit Function
21        End If

22        EvenDQs = True
23        If Len(AltDelim) = 0 Then
24            AltDelim = String(Len(Delimiter), CharNotInString(Expression))
25        End If

26        While DelimPos > 0 And Limit = -1 Or NDelims < Limit - 1
27            While (DQPos < DelimPos) And DQPos > 0
28                EvenDQs = Not EvenDQs
29                DQPos = InStr(DQPos + 1, Expression, DQ)
30            Wend
31            If EvenDQs Then
32                Mid$(Expression, DelimPos, LDelim) = AltDelim
33                NDelims = NDelims + 1
34            End If
35            DelimPos = InStr(DelimPos + LDelim, Expression, Delimiter)
36        Wend

37        SplitNew = VBA.Split(Expression, AltDelim, Limit)

38        Exit Function
ErrHandler:
39        Throw "#SplitNew (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
