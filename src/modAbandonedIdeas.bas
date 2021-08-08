Attribute VB_Name = "modAbandonedIdeas"
Option Explicit

Private Sub ParseCharByChar(ByVal CSVContents As String, DQ As String, Delimiter As String, ByRef NumRows As Long, ByRef NumCols As Long, _
          ByRef NumFields As Long, ByRef Starts() As Long, ByRef Lengths() As Long, ByRef IsLasts() As Boolean)

    Dim ColNum As Long
    Dim EvenQuotes As Boolean
    Dim ThisChar As String
    Dim PrevChar As String

    Dim i As Long 'Index to read CSVContents
    Dim j As Long 'Index to write to Starts, Lengths etc


    On Error GoTo ErrHandler
    'Ensure terminates with vbCrLf
    If Right(CSVContents, 1) <> vbCr And Right(CSVContents, 1) <> vbLf Then
        CSVContents = CSVContents + vbCrLf
    ElseIf Right(CSVContents, 1) = vbCr Then
        CSVContents = CSVContents + vbLf
    End If

    ReDim Starts(1 To 8)
    ReDim Lengths(1 To 8)
    ReDim IsLasts(1 To 8)
    ReDim QuoteCounts(1 To 8)

    NumRows = 0
    j = 1
    ColNum = 1
    EvenQuotes = True
    Starts(1) = 1
    For i = 1 To Len(CSVContents)
        If j + 1 > UBound(Starts) Then
            ReDim Preserve Starts(1 To UBound(Starts) * 2)
            ReDim Preserve Lengths(1 To UBound(Lengths) * 2)
            ReDim Preserve IsLasts(1 To UBound(IsLasts) * 2)
            ReDim Preserve QuoteCounts(1 To UBound(QuoteCounts) * 2)
        End If

        ThisChar = Mid(CSVContents, i, 1)
        If ThisChar = DQ Then
            EvenQuotes = Not EvenQuotes
        ElseIf EvenQuotes Then
            Select Case ThisChar
                Case Delimiter
                    Lengths(j) = i - Starts(j)
                    Starts(j + 1) = i + 1
                    ColNum = ColNum + 1
                    j = j + 1
                    NumFields = NumFields + 1
                Case vbLf
                    If PrevChar = vbCr Then
                        'Windows line ending
                        Lengths(j) = i - Starts(j) - 1
                    Else
                        'Unix line ending
                        Lengths(j) = i - Starts(j)
                    End If
                    Starts(j + 1) = i + 1
                    If ColNum > NumCols Then NumCols = ColNum
                    ColNum = 1
                    IsLasts(j) = True
                    j = j + 1
                    NumRows = NumRows + 1
                    NumFields = NumFields + 1
                Case vbCr
                    If Mid(CSVContents, i + 1, 1) <> vbLf Then
                        'Mac line ending (Mac pre OSX)
                        Lengths(j) = i - Starts(j)
                        Starts(j + 1) = i + 1
                        If ColNum > NumCols Then NumCols = ColNum
                        ColNum = 1
                        IsLasts(j) = True
                        j = j + 1
                        NumRows = NumRows + 1
                        NumFields = NumFields + 1
                    End If
            End Select
        End If
        PrevChar = ThisChar
    Next i

    Exit Sub
ErrHandler:
    Throw "#ParseCharByChar (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub




'Version that handles a specific line feed
Function ParseWithInStr2(ByVal CSVContents As String, DQ As String, Delimiter As String, EOL As String, ByRef NumRows As Long, ByRef NumCols As Long, _
    ByRef NumFields As Long, ByRef Starts() As Long, ByRef Lengths() As Long, ByRef IsLasts() As Boolean, QuoteCounts() As Long)


    Dim ColNum As Long
    Dim EvenQuotes As Boolean
    Dim i As Long 'Index to read CSVContents
    Dim j As Long 'Index to write to Starts, Lengths and IsLasts
    Dim LenP1 As Long
    Dim OrigLen As Long
    Dim PosDL As Long
    Dim PosDQ As Long
    Dim PosEOL As Long
    Dim QuoteCount As Long
    Dim Which As Long
    Dim LEOL As Long


    On Error GoTo ErrHandler

    LEOL = Len(EOL)

    ReDim Starts(1 To 8)
    ReDim Lengths(1 To 8)
    ReDim IsLasts(1 To 8)
    ReDim QuoteCounts(1 To 8)

    
    'Ensure CSVContents terminates with vbCrLf. Need to consider if this section is correct!
    OrigLen = Len(CSVContents)
    If Right(CSVContents, 1) <> vbCr And Right(CSVContents, 1) <> vbLf Then
        CSVContents = CSVContents + vbCrLf
    ElseIf Right(CSVContents, 1) = vbCr Then
        CSVContents = CSVContents + vbLf
    End If
    LenP1 = Len(CSVContents) + 1

    j = 1
    ColNum = 1
    EvenQuotes = True
    Starts(1) = 1

    Do
        If EvenQuotes Then
            If PosDL <= i Then PosDL = InStr(i + 1, CSVContents, Delimiter): If PosDL = 0 Then PosDL = LenP1
            If PosEOL <= i Then PosEOL = InStr(i + 1, CSVContents, EOL): If PosEOL = 0 Then PosEOL = LenP1
            '    If PosLF <= i Then PosLF = InStr(i + 1, CSVContents, vbLf): If PosLF = 0 Then PosLF = LenP1
            '    If PosCR <= i Then PosCR = InStr(i + 1, CSVContents, vbCr): If PosCR = 0 Then PosCR = LenP1
            If PosDQ <= i Then PosDQ = InStr(i + 1, CSVContents, DQ): If PosDQ = 0 Then PosDQ = LenP1
            i = Min3(PosDL, PosEOL, PosDQ, Which)
            
            If i >= LenP1 Then Exit Do

            If j + 1 > UBound(Starts) Then
                ReDim Preserve Starts(1 To UBound(Starts) * 2)
                ReDim Preserve Lengths(1 To UBound(Lengths) * 2)
                ReDim Preserve IsLasts(1 To UBound(IsLasts) * 2)
                ReDim Preserve QuoteCounts(1 To UBound(QuoteCounts) * 2)
            End If

            Select Case Which
                Case 1 'Found Delimiter
                    Lengths(j) = i - Starts(j)
                    Starts(j + 1) = i + 1
                    ColNum = ColNum + 1
                    QuoteCounts(j) = QuoteCount
                    QuoteCount = 0
                    j = j + 1
                    NumFields = NumFields + 1
                Case 2 'Found EOL
                    Lengths(j) = i - Starts(j)
                    Starts(j + 1) = i + LEOL
                    i = i + LEOL - 1
                    If ColNum > NumCols Then NumCols = ColNum
                    ColNum = 1
                    IsLasts(j) = True
                    QuoteCounts(j) = QuoteCount
                    QuoteCount = 0
                    j = j + 1
                    NumRows = NumRows + 1
                    NumFields = NumFields + 1

                Case 3 'Found DQ
                    EvenQuotes = False
                    QuoteCount = QuoteCount + 1
            End Select
        Else
            PosDQ = InStr(i + 1, CSVContents, DQ)
            If PosDQ = 0 Then
                'Malformed CSVContents. There should always be an even number of double quotes. _
                 If there are an odd number then all text after the last double quote will be _
                 (part of) the last field in the last line.
                Lengths(j) = OrigLen - Starts(j) + 1
                If ColNum > NumCols Then NumCols = ColNum
                NumRows = NumRows + 1
                NumFields = NumFields + 1
                Exit Do
            Else
                i = PosDQ
                EvenQuotes = True
                QuoteCount = QuoteCount + 1
            End If
        End If
    Loop


    Exit Function
ErrHandler:
    Throw "#ParseWithInStr (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function Min3(N1 As Long, N2 As Long, N3 As Long, ByRef Which As Long) As Long

    If N1 < N2 Then
        Min3 = N1
        Which = 1
    Else
        Min3 = N2
        Which = 2
    End If

    If N3 < Min3 Then
        Min3 = N3
        Which = 3
    End If

End Function

