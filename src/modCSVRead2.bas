Attribute VB_Name = "modCSVRead2"
Option Explicit

Function CSVRead2(FileName As String, Delimiter As String)
    Dim FSO As New Scripting.FileSystemObject

    Dim CSVContents As String
    Dim T As TextStream
    Dim NumRows As Long
    Dim NumCols As Long
    Dim NumFields As Long
    Dim Starts() As Long
    Dim Lengths() As Long
    Dim IsLasts() As Boolean
    Dim Contents() As String
    Dim QuoteCounts() As Long
    Dim i As Long, j As Long, k As Long
    Const DQ = """"
    Const DQ2 = """"""

    On Error GoTo ErrHandler
    Set T = FSO.GetFile(FileName).OpenAsTextStream

    CSVContents = T.ReadAll
    T.Close

    ParseWithInStr CSVContents, DQ, Delimiter, NumRows, NumCols, NumFields, Starts, Lengths, IsLasts, QuoteCounts
   '  ParseWithInStr2 CSVContents, DQ, Delimiter, vbCrLf, NumRows, NumCols, NumFields, Starts, Lengths, IsLasts, QuoteCounts

    ReDim Contents(1 To NumRows, 1 To NumCols)

    i = 1: j = 1
    For k = 1 To NumFields
        If QuoteCounts(k) = 0 Then
            Contents(i, j) = Mid(CSVContents, Starts(k), Lengths(k))
        ElseIf Mid(CSVContents, Starts(k), 1) = DQ And Mid(CSVContents, Starts(k) + Lengths(k) - 1, 1) = DQ Then
            Contents(i, j) = Mid(CSVContents, Starts(k) + 1, Lengths(k) - 2)
            If QuoteCounts(k) > 2 Then
                Contents(i, j) = Replace(Contents(i, j), DQ2, DQ)
            End If

        Else
            Contents(i, j) = Mid(CSVContents, Starts(k), Lengths(k))
        End If
          

        If QuoteCounts(j) > 0 Then
            If Left(Contents(i, j), 1) = """" Then
                If Right(Contents(i, j), 1) = """" Then
                    Contents(i, j) = Mid(Contents(i, j), 2, Len(Contents(i, j)) - 2)
                    If QuoteCounts(j) > 2 Then
                        Contents(i, j) = Replace(Contents(i, j), """""", """")
                    End If
                End If
            End If
        End If
        If IsLasts(k) Then
            i = i + 1: j = 1
        Else
            j = j + 1
        End If
    Next k

    CSVRead2 = Contents


    Exit Function
ErrHandler:
    CSVRead2 = "#CSVRead2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


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



' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Min4
' Author     : Philip Swannell
' Date       : 07-Aug-2021
' Purpose    : Returns the minimum of four inputs and an indicator of which of the four was the minimum
' -----------------------------------------------------------------------------------------------------------------------
Function Min4(N1 As Long, N2 As Long, N3 As Long, N4 As Long, ByRef Which As Long) As Long

    If N1 < N2 Then
        Min4 = N1
        Which = 1
    Else
        Min4 = N2
        Which = 2
    End If

    If N3 < Min4 Then
        Min4 = N3
        Which = 3
    End If

    If N4 < Min4 Then
        Min4 = N4
        Which = 4
    End If

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





'Version that handles mixed line feeds
Sub ParseWithInStr(ByVal CSVContents As String, DQ As String, Delimiter As String, ByRef NumRows As Long, ByRef NumCols As Long, _
          ByRef NumFields As Long, ByRef Starts() As Long, ByRef Lengths() As Long, ByRef IsLasts() As Boolean, QuoteCounts() As Long)


    Dim ColNum As Long
    Dim EvenQuotes As Boolean
    Dim i As Long 'Index to read CSVContents
    Dim j As Long 'Index to write to Starts, Lengths and IsLasts
    Dim LenP1 As Long
    Dim OrigLen As Long
    Dim PosCR As Long
    Dim PosDL As Long
    Dim PosDQ As Long
    Dim PosLF As Long
    Dim QuoteCount As Long
    Dim Which As Long
    Dim LDlm As Long

    On Error GoTo ErrHandler

    ReDim Starts(1 To 8)
    ReDim Lengths(1 To 8)
    ReDim IsLasts(1 To 8)
    ReDim QuoteCounts(1 To 8)

    
    'Ensure CSVContents terminates with vbCrLf
    
    LDlm = Len(Delimiter)
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
            If PosLF <= i Then PosLF = InStr(i + 1, CSVContents, vbLf): If PosLF = 0 Then PosLF = LenP1
            If PosCR <= i Then PosCR = InStr(i + 1, CSVContents, vbCr): If PosCR = 0 Then PosCR = LenP1
            If PosDQ <= i Then PosDQ = InStr(i + 1, CSVContents, DQ): If PosDQ = 0 Then PosDQ = LenP1
            i = Min4(PosDL, PosLF, PosCR, PosDQ, Which)
            
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
                    Starts(j + 1) = i + LDlm
                    ColNum = ColNum + 1
                    QuoteCounts(j) = QuoteCount
                    QuoteCount = 0
                    j = j + 1
                    NumFields = NumFields + 1
                    i = i + LDlm - 1
                Case 2 'Found LF, Unix line ending
                    Lengths(j) = i - Starts(j)
                    Starts(j + 1) = i + 1
                    'i = i + 1
                    If ColNum > NumCols Then NumCols = ColNum
                    ColNum = 1
                    IsLasts(j) = True
                    QuoteCounts(j) = QuoteCount
                    QuoteCount = 0
                    j = j + 1
                    NumRows = NumRows + 1
                    NumFields = NumFields + 1

                Case 3 'Found CR. Either Windows or (old) Mac ending
                    
                    Lengths(j) = i - Starts(j)
                    If Mid(CSVContents, i + 1, 1) = vbLf Then 'Safe to look one character ahead since CSVContents terminates with vbCrLf
                        'Windows line ending
                        Starts(j + 1) = i + 2
                        i = i + 1
                    Else
                        'Mac line ending (Mac pre OSX)
                        Starts(j + 1) = i + 1
                    End If

                    If ColNum > NumCols Then NumCols = ColNum
                    ColNum = 1
                    IsLasts(j) = True
                    QuoteCounts(j) = QuoteCount
                    QuoteCount = 0
                    j = j + 1
                    NumRows = NumRows + 1
                    NumFields = NumFields + 1


                Case 4 'Found DQ
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

    Exit Sub
ErrHandler:
    Throw "#ParseWithInStr (line " & CStr(Erl) + "): " & Err.Description & "!"
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

    'For debugging line below is helpful
    'ParseWithInStr = sArrayRange(sArrayStack(NumRows, NumCols, NumFields), sArrayTranspose(Starts), sArrayTranspose(Lengths), sArrayTranspose(IsLasts))



    Exit Function
ErrHandler:
    Throw "#ParseWithInStr (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
































