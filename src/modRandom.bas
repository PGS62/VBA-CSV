Attribute VB_Name = "modRandom"
Option Explicit

Private Function RandomString(AllowLineFeed As Boolean, Unicode As Boolean, EOL As String)
          Dim length As Long
          Dim i As Long
          Dim Res As String
          Const MAXLEN = 20
1         On Error GoTo ErrHandler
2         length = CLng(1 + Rnd() * MAXLEN)
3         Res = String(length, " ")

4         For i = 1 To length
5             If Unicode Then
6                 Mid(Res, i, 1) = ChrW(33 + Rnd() * 370)
7             Else
8                 Mid(Res, i, 1) = Chr(34 + Rnd() * 88)
9             End If

10            If Not AllowLineFeed Then
11                If Mid(Res, i, 1) = vbLf Or Mid(Res, i, 1) = vbCr Then
12                    Mid(Res, i, 1) = " "
13                End If
14            End If
15        Next

16        If AllowLineFeed Then
17            If length > 5 Then
18                If Rnd() < 0.2 Then
19                    Mid(Res, length / 2, Len(EOL)) = EOL
20                End If
21            End If
22        End If



23        RandomString = Res

24        Exit Function
ErrHandler:
25        Err.Raise vbObjectError + 1, , "#RandomString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function RandomStrings(NumRows As Long, NumCols As Long, Unicode As Boolean, AllowLineFeed As Boolean, EOL As String)
          Dim Result() As String, i As Long, j As Long
1         On Error GoTo ErrHandler
2         ReDim Result(1 To NumRows, 1 To NumCols)
3         For i = 1 To NumRows
4             For j = 1 To NumCols
5                 Result(i, j) = RandomString(AllowLineFeed, Unicode, EOL)
6             Next j
7         Next i
8         If AllowLineFeed Then
9             Result(1, 1) = "Here" & EOL & "be" & EOL & "line" & EOL & "feeds"
10        End If
11        RandomStrings = Result
12        Exit Function
ErrHandler:
13        Err.Raise vbObjectError + 1, , "#RandomStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function RandomLong() As Long
1         RandomLong = CLng((Rnd() - 0.5) * 2000000)
End Function

Function RandomLongs(NumRows As Long, NumCols As Long)
          Dim Result() As Long, i As Long, j As Long
1         On Error GoTo ErrHandler
2         ReDim Result(1 To NumRows, 1 To NumCols)
3         For i = 1 To NumRows
4             For j = 1 To NumCols
5                 Result(i, j) = RandomLong()
6             Next j
7         Next i
8         RandomLongs = Result
9         Exit Function
ErrHandler:
10        Err.Raise vbObjectError + 1, , "#RandomLongs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function RandomDouble() As Double
1         On Error GoTo ErrHandler
          'Trick! - Generate a Double that has an exact representation as String. Avoids rounding errors when we write to disk and read back
2         RandomDouble = CDbl(CStr((Rnd() - 0.5) * 2 * 10 ^ ((Rnd() - 0.5) * 20)))
3         Exit Function
ErrHandler:
4         Err.Raise vbObjectError + 1, , "#RandomDouble (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function RandomDoubles(NumRows As Long, NumCols As Long)
          Dim Result() As Double, i As Long, j As Long
1         On Error GoTo ErrHandler
2         ReDim Result(1 To NumRows, 1 To NumCols)
3         For i = 1 To NumRows
4             For j = 1 To NumCols
5                 Result(i, j) = RandomDouble()
6             Next j
7         Next i
8         RandomDoubles = Result
9         Exit Function
ErrHandler:
10        Err.Raise vbObjectError + 1, , "#RandomDoubles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function RandomBoolean() As Boolean
1         RandomBoolean = Rnd() < 0.5
End Function

Function RandomBooleans(NumRows As Long, NumCols As Long)
          Dim Result() As Boolean, i As Long, j As Long
1         On Error GoTo ErrHandler
2         ReDim Result(1 To NumRows, 1 To NumCols)
3         For i = 1 To NumRows
4             For j = 1 To NumCols
5                 Result(i, j) = RandomBoolean()
6             Next j
7         Next i
8         RandomBooleans = Result
9         Exit Function
ErrHandler:
10        Err.Raise vbObjectError + 1, , "#RandomBooleans (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function RandomDate()
1         On Error GoTo ErrHandler
2         RandomDate = CDate(CLng(25569 + Rnd() * 36525)) 'Date in range 1 Jan 1970 to 1 Jan 2070
3         Exit Function
ErrHandler:
4         Err.Raise vbObjectError + 1, , "#RandomDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function RandomDates(NumRows As Long, NumCols As Long)
          Dim Result() As Date, i As Long, j As Long
1         On Error GoTo ErrHandler
2         ReDim Result(1 To NumRows, 1 To NumCols)
3         For i = 1 To NumRows
4             For j = 1 To NumCols
5                 Result(i, j) = RandomDate()
6             Next j
7         Next i
8         RandomDates = Result
9         Exit Function
ErrHandler:
10        Err.Raise vbObjectError + 1, , "#RandomDates (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function RandomErrorValue()
          Dim n As Long
1         On Error GoTo ErrHandler
2         n = CLng(0.5 + Rnd() * 14)
3         RandomErrorValue = CVErr(Choose(n, xlErrBlocked, xlErrCalc, xlErrConnect, xlErrDiv0, xlErrField, xlErrGettingData, xlErrNA, xlErrName, xlErrNull, xlErrNum, xlErrRef, xlErrSpill, xlErrUnknown, xlErrValue))
4         Exit Function
ErrHandler:
5         Err.Raise vbObjectError + 1, , "#RandomErrorValue (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function RandomErrorValues(NumRows As Long, NumCols As Long)
          Dim Result() As Variant, i As Long, j As Long
1         On Error GoTo ErrHandler
2         ReDim Result(1 To NumRows, 1 To NumCols)
3         For i = 1 To NumRows
4             For j = 1 To NumCols
5                 Result(i, j) = RandomErrorValue()
6             Next j
7         Next i
8         RandomErrorValues = Result
9         Exit Function
ErrHandler:
10        Err.Raise vbObjectError + 1, , "#RandomErrorValues (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function RandomVariant(DateFormat As String, AllowLineFeed As Boolean, Unicode As Boolean, EOL As String)

          Dim n As Long
          Const NUMTYPES = 11

1         On Error GoTo ErrHandler
2         n = CLng(0.5 + NUMTYPES * Rnd())

3         Select Case n
              Case 1
4                 RandomVariant = RandomBoolean()
5             Case 2
6                 RandomVariant = RandomLong()
7             Case 3
8                 RandomVariant = RandomDouble()
9             Case 4
10                RandomVariant = RandomString(AllowLineFeed, Unicode, EOL)
11            Case 5
12                RandomVariant = RandomDate()
13            Case 6
14                RandomVariant = vbNullString
15            Case 7
                  'String that looks like a number
16                RandomVariant = CStr(RandomDouble())
17            Case 8
                  'String that looks like a date
18                RandomVariant = Format(CLng(RandomDate()), DateFormat)
19            Case 9
                  'String that looks like Boolean
20                RandomVariant = CStr(RandomBoolean())
21            Case 10
22                RandomVariant = Empty
23            Case 11
24                RandomVariant = RandomErrorValue()
25        End Select

26        Exit Function
ErrHandler:
27        Err.Raise vbObjectError + 1, , "#RandomVariant (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function




Function RandomVariants(NRows As Long, NCols As Long, AllowLineFeed As Boolean, Unicode As Boolean, ByVal EOL As String)

          Const DateFormat = "yyyy-mmm-dd"
          Const MAXCOLS = 5
          Const MAXROWS = 50
          Dim i As Long
          Dim j As Long
          Dim Res() As Variant

1         On Error GoTo ErrHandler
2         EOL = OStoEOL(EOL, "EOL")
3         ReDim Res(1 To NRows, 1 To NCols)

4         For i = 1 To NRows
5             For j = 1 To NCols
6                 Res(i, j) = RandomVariant(DateFormat, AllowLineFeed, Unicode, EOL)
7             Next j
8         Next i
9         If AllowLineFeed Then
10            Res(1, 1) = "Here" & EOL & "be" & EOL & "line" & EOL & "feeds"
11        End If

12        RandomVariants = Res

13        Exit Function
ErrHandler:
14        Err.Raise vbObjectError + 1, , "#RandomVariants (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
