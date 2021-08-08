Attribute VB_Name = "modCSV_V2"
Option Explicit
Private Const Err_EmptyFile = "File is empty"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead_V2
' Purpose    : Register the function CSVRead_V2 with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Sub RegisterCSVRead_V2()
    Const FnDesc = "Returns the contents of a comma-separated file on disk as an array."
    Dim ArgDescs() As String
    ReDim ArgDescs(1 To 11)
    ArgDescs(1) = "The full name of the file, including the path."
    ArgDescs(2) = "TRUE to convert Numbers, Dates, Booleans and Excel Errors into their typed values, or FALSE to leave as strings. For more control enter a string containing the letters N,D,B, and E eg ""NB"" to convert just numbers and Booleans, not dates or errors."
    ArgDescs(3) = "Delimiter character. Defaults to the first instance of comma, tab, semi-colon, colon or vertical bar found in the file outside quoted regions. Enter FALSE to to see the file's raw contents as would be displayed in a text editor."
    ArgDescs(4) = "The format of dates in the file such as D-M-Y, M-D-Y or Y/M/D. If omitted a value is read from Windows regional settings. Repeated D's (or M's or Y's) are equivalent to single instances, so that d-m-y and DD-MMM-YYYY are equivalent."
    ArgDescs(5) = "The row in the file at which reading starts. Optional and defaults to 1 to read from the first row."
    ArgDescs(6) = "The column in the file at which reading starts. Optional and defaults to 1 to read from the first column."
    ArgDescs(7) = "The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the end of the file are read."
    ArgDescs(8) = "The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol are read."
    ArgDescs(9) = "Enter TRUE if the file is unicode, FALSE if the file is ascii. Omit to infer from the file's contents."
    ArgDescs(10) = "Value to represent empty fields (successive delimiters) in the file. May be a string or an Empty value. Optional and defaults to the zero-length string."
    ArgDescs(11) = "The character that represents a decimal point. If omitted, then the value from Windows regional settings is used."
    Application.MacroOptions "CSVRead_V2", FnDesc, , , , , , , , , ArgDescs
End Sub

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVRead_V2
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path.
' ConvertTypes: TRUE to convert Numbers, Dates, Booleans and Excel Errors into their typed values, or
'             FALSE to leave as strings. For more control enter a string containing the
'             letters N,D,B, and E eg "NB" to convert just numbers and Booleans, not dates
'             or errors.
' Delimiter : Delimiter character. Defaults to the first instance of comma, tab, semi-colon, colon or
'             vertical bar found in the file outside quoted regions. Enter FALSE to to see
'             the file's raw contents as would be displayed in a text editor.
' DateFormat: The format of dates in the file such as D-M-Y, M-D-Y or Y/M/D. If omitted a value is read
'             from Windows regional settings. Repeated D's (or M's or Y's) are equivalent
'             to single instances, so that d-m-y and DD-MMM-YYYY are equivalent.
' SkipToRow : The row in the file at which reading starts. Optional and defaults to 1 to read from the
'             first row.
' SkipToCol : The column in the file at which reading starts. Optional and defaults to 1 to read from
'             the first column.
' NumRows   : The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to
'             the end of the file are read.
' NumCols   : The number of columns to read from the file. If omitted (or zero), all columns from
'             SkipToCol are read.
' Unicode   : Enter TRUE if the file is unicode, FALSE if the file is ascii. Omit to infer from the
'             file's contents.
' ShowMissingsAs: Value to represent empty fields (successive delimiters) in the file. May be a string or an
'             Empty value. Optional and defaults to the zero-length string.
' DecimalSeparator: The character that represents a decimal point. If omitted, then the value from Windows
'             regional settings is used.
'
' Notes     : See also CSVWrite for which this function is the inverse.
'
'             The function handles all csv files that conform to the standards described in
'             RFC4180  https://www.rfc-editor.org/rfc/rfc4180.txt including files with
'             quoted fields.
'
'             In addition the function handles files which break some of those standards:
'             * Not all lines of the file need have the same number of fields. The function
'             "pads" with the value given by ShowMissingsAs.
'             * Fields which start with a double quote but do not end with a double quote
'             are handled by being returned unchanged. Necessarily such fields have an even
'             number of double quotes, or otherwise the field will be treated as the last
'             field in the file.
'             * The standard states that csv files should have Windows-style line endings,
'             but the function supports Windows, Unix and (old) Mac line endings.
'---------------------------------------------------------------------------------------------------------
Function CSVRead_V2(FileName As String, Optional ConvertTypes As Variant = False, Optional ByVal Delimiter As Variant, _
    Optional DateFormat As String, Optional ByVal SkipToRow As Long = 1, Optional ByVal SkipToCol As Long = 1, _
    Optional ByVal NumRows As Long = 0, Optional ByVal NumCols As Long = 0, _
    Optional ByVal Unicode As Variant, Optional ByVal ShowMissingsAs As Variant = "", _
    Optional DecimalSeparator As String = vbNullString)

    Const Err_Delimiter = "Delimiter character must be passed as a string, FALSE for no delimiter, or else omitted to infer from the file's contents"
    Const Err_FileIsUniCode = "Unicode must be passed as TRUE or FALSE, or omitted to infer from the file's contents"
    Const Err_InFuncWiz = "#Disabled in Function Dialog!"
    Const Err_NumCols = "NumCols must be positive to read a given number or columns, or zero or omitted to read all columns from SkipToCol to the maximum column encountered."
    Const Err_NumRows = "NumRows must be positive to read a given number or rows, or zero or omitted to read all rows from SkipToRow to the end of the file."
    Const Err_Seps = "DecimalSeparator must be different from Delimiter"
    Const Err_SkipToCol = "SkipToCol must be at least 1."
    Const Err_SkipToRow = "SkipToRow must be at least 1."
    
    'The lower bounds for the array returned by this function is set by the constant LB rather than _
     by any "Option Base 1" or "Option Base 0" that might appear at the start of the module.
    Const LB = 1
    Const DQ = """"
    Const DQ2 = """"""
    
    Dim AnyConversion As Boolean
    Dim CSVS As clsCSVStream
    Dim DateOrder As Long
    Dim DateSeparator As String
    Dim F As Scripting.File
    Dim FSO As New Scripting.FileSystemObject
    Dim i As Long
    Dim j As Long
    Dim NotDelimited As Boolean
    Dim NumColsInReturn As Long
    Dim NumRowsInReturn As Long
    Dim CSVContents As String
    Dim RemoveQuotes As Boolean
    Dim ReturnArray() As Variant
    Dim SepsStandard As Boolean
    Dim ShowDatesAsDates As Boolean
    Dim ShowErrorsAsErrors As Boolean
    Dim ShowLogicalsAsLogicals As Boolean
    Dim ShowMissingAsNullString As Boolean
    Dim ShowNumbersAsNumbers As Boolean
    Dim strDelimiter As String
    Dim SysDateOrder As Long
    Dim SysDateSeparator As String
    Dim SysDecimalSeparator As String
    Dim T As Scripting.TextStream
    
    On Error GoTo ErrHandler

    If FunctionWizardActive() Then
        CSVRead_V2 = Err_InFuncWiz
        Exit Function
    End If

    'Parse and validate inputs...
    If IsEmpty(Unicode) Or IsMissing(Unicode) Then
        Unicode = IsUnicodeFile(FileName)
    ElseIf VarType(Unicode) <> vbBoolean Then
        Throw Err_FileIsUniCode
    End If

    If VarType(Delimiter) = vbBoolean Then
        If Not Delimiter Then
            NotDelimited = True
        Else
            Throw Err_Delimiter
        End If
    ElseIf VarType(Delimiter) = vbString Then
        strDelimiter = Delimiter
    ElseIf IsEmpty(Delimiter) Or IsMissing(Delimiter) Then
        strDelimiter = InferDelimiter(FileName, CBool(Unicode))
    Else
        Throw Err_Delimiter
    End If

    ParseTypeConversion ConvertTypes, ShowNumbersAsNumbers, _
        ShowDatesAsDates, ShowLogicalsAsLogicals, ShowErrorsAsErrors, RemoveQuotes

    If ShowNumbersAsNumbers Then
        If ((DecimalSeparator = Application.DecimalSeparator) Or DecimalSeparator = vbNullString) Then
            SepsStandard = True
        ElseIf DecimalSeparator = strDelimiter Then
            Throw Err_Seps
        End If
    End If

    If ShowDatesAsDates Then
        ParseDateFormat DateFormat, DateOrder, DateSeparator
        SysDateOrder = Application.International(xlDateOrder)
        SysDateSeparator = Application.International(xlDateSeparator)
    End If

    If SkipToRow < 1 Then Throw Err_SkipToRow
    If SkipToCol < 1 Then Throw Err_SkipToCol
    If NumRows < 0 Then Throw Err_NumRows
    If NumCols < 0 Then Throw Err_NumCols

    If TypeName(ShowMissingsAs) = "Range" Then
        ShowMissingsAs = ShowMissingsAs.value
    End If
    If Not (IsEmpty(ShowMissingsAs) Or VarType(ShowMissingsAs) = vbString) Then
        Throw "ShowMissingsAs must be Empty or a string"
    End If
    If VarType(ShowMissingsAs) = vbString Then
        ShowMissingAsNullString = ShowMissingsAs = ""
    End If
    'End of input validation
          
    If NotDelimited Then
        CSVRead_V2 = ShowTextFile(FileName, SkipToRow, NumRows, CBool(Unicode))
        Exit Function
    End If
          
    If SkipToRow = 1 And SkipToCol = 1 And NumRows = 0 And NumCols = 0 Then
    
        Set F = FSO.GetFile(FileName)
        Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))
        If T.atEndOfStream Then
            T.Close: Set T = Nothing: Set F = Nothing
            Throw Err_EmptyFile
        End If

        CSVContents = T.ReadAll
        T.Close: Set T = Nothing: Set F = Nothing
    
    Else
        'TODO get this section working again, Need to populate CSVContents with contents of relevant lines from file
        '        Set CSVS = CreateCSVStream(FileName, EOL, Unicode)
        '        For i = 1 To SkipToRow - 1
        '            CSVS.ReadLine
        '        Next i
        '        CSVS.StartRecording
        '        If NumRows > 0 Then
        '            For i = 1 To NumRows
        '                CSVS.ReadLine
        '            Next
        '        Else
        '            While Not CSVS.atEndOfStream
        '                CSVS.ReadLine
        '            Wend
        '        End If
        '
        '        Lines = CSVS.ReportAllLinesRead()
        '        If Not CSVS.QuotesEncountered Then
        '            RemoveQuotes = False
        '        End If
        '        Set CSVS = Nothing
    End If
    
    Dim NumFields As Long
    Dim Starts() As Long
    Dim Lengths() As Long
    Dim IsLasts() As Boolean
    Dim QuoteCounts() As Long
    Dim k As Long, m As Long
    Dim ThisField As String
    
    AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
        ShowLogicalsAsLogicals Or ShowErrorsAsErrors Or (Not ShowMissingAsNullString)
        
    Call ParseCSVContents(CSVContents, DQ, strDelimiter, NumRowsInReturn, NumColsInReturn, NumFields, Starts, Lengths, IsLasts, QuoteCounts)
        
    ReDim ReturnArray(LB To LB + NumRowsInReturn - 1, LB To LB + NumColsInReturn - 1)
        
    'TODO Handle SkipToCol and NumCols

    i = LB: j = LB
    For k = 1 To NumFields
        If QuoteCounts(k) = 0 Or Not RemoveQuotes Then
            ThisField = Mid(CSVContents, Starts(k), Lengths(k))
        ElseIf Mid(CSVContents, Starts(k), 1) = DQ And Mid(CSVContents, Starts(k) + Lengths(k) - 1, 1) = DQ Then
            ThisField = Mid(CSVContents, Starts(k) + 1, Lengths(k) - 2)
            If QuoteCounts(k) > 2 Then
                ThisField = Replace(ThisField, DQ2, DQ)
            End If
        Else 'Field which does not start with quote but contains quotes, so not RFC4180 compliant. We do not replace DQ2 by DQ in this case.
            ThisField = Mid(CSVContents, Starts(k), Lengths(k))
        End If
        
        If AnyConversion And QuoteCounts(k) = 0 Then
                ReturnArray(i, j) = CastToVariant(ThisField, _
                    ShowNumbersAsNumbers, SepsStandard, DecimalSeparator, SysDecimalSeparator, _
                    ShowDatesAsDates, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, _
                    ShowLogicalsAsLogicals, ShowErrorsAsErrors, ShowMissingAsNullString, ShowMissingsAs)
        Else
            ReturnArray(i, j) = ThisField
        End If
        
        If IsLasts(k) Then
            'Handle case of file having non-constant number of fields per line.
            If j < NumColsInReturn Then
                If Not IsEmpty(ShowMissingsAs) Then
                    For m = j + 1 To NumColsInReturn
                        ReturnArray(i, m) = ShowMissingsAs
                    Next m
                End If
            End If
            i = i + 1: j = LB
        Else
            j = j + 1
        End If
    Next k

    CSVRead_V2 = ReturnArray

    Exit Function

ErrHandler:
    CSVRead_V2 = "#CSVRead_V2 (line " & CStr(Erl) + "): " & Err.Description & "!"
    If Not CSVS Is Nothing Then
        Set CSVS = Nothing
    End If
    If Not T Is Nothing Then
        T.Close
        Set T = Nothing
    End If

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseTypeConversion
' Purpose    : Parse the input ConvertTypes to set five Boolean flags which are passed by reference
' Parameters :
'  ConvertTypes        :
'  ShowNumbersAsNumbers  : Should fields in the file that look like numbers be returned as Numbers? (Doubles)
'  ShowDatesAsDates      : Should fields in the file that look like dates with the specified DateFormat be returned as Dates?
'  ShowLogicalsAsLogicals: Should fields in the file that are TRUE or FALSE (case insensitive) be returned as Booleans?
'  ShowErrorsAsErrors    : Should fields in the file that look like Excel errors (#N/A #REF! etc) be returned as errors?
'  RemoveQuotes          : Should quoted fields be unquoted?
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseTypeConversion(ByVal ConvertTypes As Variant, ByRef ShowNumbersAsNumbers As Boolean, _
    ByRef ShowDatesAsDates As Boolean, ByRef ShowLogicalsAsLogicals As Boolean, _
    ByRef ShowErrorsAsErrors As Boolean, ByRef RemoveQuotes As Boolean)

    Const Err_ConvertTypes = "ConvertTypes must be TRUE (convert all types), FALSE (no conversion) or a string of letter: 'N' to show numbers as numbers, 'D' to show dates as dates, 'L' to show logicals as logicals, `E` to show Excel errors as errors, Q to show quoted fields with their quotes."
    Dim i As Long

    On Error GoTo ErrHandler
    If TypeName(ConvertTypes) = "Range" Then
        ConvertTypes = ConvertTypes.value
    End If

    If VarType(ConvertTypes) = vbBoolean Then
        If ConvertTypes Then
            ShowNumbersAsNumbers = True
            ShowDatesAsDates = True
            ShowLogicalsAsLogicals = True
            ShowErrorsAsErrors = True
            RemoveQuotes = True
        Else
            ShowNumbersAsNumbers = False
            ShowDatesAsDates = False
            ShowLogicalsAsLogicals = False
            ShowErrorsAsErrors = False
            RemoveQuotes = True
        End If
    ElseIf VarType(ConvertTypes) = vbString Then
        ShowNumbersAsNumbers = False
        ShowDatesAsDates = False
        ShowLogicalsAsLogicals = False
        ShowErrorsAsErrors = False
        RemoveQuotes = True
        For i = 1 To Len(ConvertTypes)
            Select Case UCase(Mid(ConvertTypes, i, 1))
                Case "N"
                    ShowNumbersAsNumbers = True
                Case "D"
                    ShowDatesAsDates = True
                Case "L", "B" 'Logicals aka Booleans
                    ShowLogicalsAsLogicals = True
                Case "E"
                    ShowErrorsAsErrors = True
                Case "Q"
                    RemoveQuotes = False
                Case Else
                    Throw "Unrecognised character '" + Mid(ConvertTypes, i, 1) + "' in ConvertTypes."
            End Select
        Next i
    Else
        Throw Err_ConvertTypes
    End If

    Exit Sub
ErrHandler:
    Throw "#ParseTypeConversion: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Min4
' Author     : Philip Swannell
' Date       : 07-Aug-2021
' Purpose    : Returns the minimum of four inputs and an indicator of which of the four was the minimum
' -----------------------------------------------------------------------------------------------------------------------
Private Function Min4(N1 As Long, N2 As Long, N3 As Long, N4 As Long, ByRef Which As Long) As Long

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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCSVContents
' Purpose    : Parse the contents of a CSV file.
' Parameters :
'  CSVContents: The contents of a CSV file as a string
'  QuoteChar  : The quote character, usually ascii 34 ("), which allow fields to contain characters that would otherwise
'               be significant to parsing, such as delimiters or new line characters.
'  Delimiter  : The character that separates fields within each line.
'  NumRows    : Set to the number of rows in the file.
'  NumCols    : Set to the number of columns in the file, i.e. the maximum number of fields in any single line.
'  NumFields  : Set to the number of fields in the file.  May be less than NumRows times NumCols if not all lines have
'               the same number of fields.
'  Starts     : Set to an array of size at least NumFields. Element k gives the point in CSVContents at which the kth
'               field starts.
'  Lengths    : Set to an array of size at least NumFields. Element k gives the length of the kth field.
'  IsLasts    : Set to an array of size at least NumFields. Element k indicates whether the kth field is the last field
'               in its line.
'  QuoteCounts: Set to an array of size at least NumFields. Element k gives the number of QuoteChars that appear in the
'               kth field.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseCSVContents(ByVal CSVContents As String, QuoteChar As String, Delimiter As String, ByRef NumRows As Long, ByRef NumCols As Long, _
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
            If PosDQ <= i Then PosDQ = InStr(i + 1, CSVContents, QuoteChar): If PosDQ = 0 Then PosDQ = LenP1
            i = Min4(PosDL, PosLF, PosCR, PosDQ, Which)
            
            If i >= LenP1 Then Exit Do

            If j + 1 > UBound(Starts) Then
                ReDim Preserve Starts(1 To UBound(Starts) * 2)
                ReDim Preserve Lengths(1 To UBound(Lengths) * 2)
                ReDim Preserve IsLasts(1 To UBound(IsLasts) * 2)
                ReDim Preserve QuoteCounts(1 To UBound(QuoteCounts) * 2)
            End If

            Select Case Which
                Case 1
                    'Found Delimiter
                    Lengths(j) = i - Starts(j)
                    Starts(j + 1) = i + LDlm
                    ColNum = ColNum + 1
                    QuoteCounts(j) = QuoteCount
                    QuoteCount = 0
                    j = j + 1
                    NumFields = NumFields + 1
                    i = i + LDlm - 1
                Case 2
                    'Found vbLf, Unix line ending
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
                Case 3
                    'Found vbCr. Either Windows or (old) Mac line ending
                    Lengths(j) = i - Starts(j)
                    'It is safe to look one character ahead since CSVContents terminates with vbCrLf
                    If Mid(CSVContents, i + 1, 1) = vbLf Then
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
                Case 4
                    'Found QuoteChar
                    EvenQuotes = False
                    QuoteCount = QuoteCount + 1
            End Select
        Else
            PosDQ = InStr(i + 1, CSVContents, QuoteChar)
            If PosDQ = 0 Then
                'Malformed CSVContents (not RFC4180 compliant). There should always be an even number of double quotes. _
                 If there are an odd number then all text after the last double quote in the file will be (part of) _
                 the last field in the last line.
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
    Throw "#ParseCSVContents (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToVariant
' Purpose    : Convert a string to a variable of another type, or return the string unchanged if conversion not possible.
'              Always unquotes quoted strings.
' Parameters :
'  strIn                    : The input string.
'Numbers
'  ShowNumbersAsNumbers     : Boolean - should conversion to Double be attempted?
'  SepsStandard             : Is the decimal separator the same as the system defaults? If true then the next two
'                             arguments are ignored.
'  DecimalSeparator         : The decimal separator used in the input string.
'  SysDecimalSeparator      : The default decimal separator on the system.
'Dates
'  ShowDatesAsDates         : Should strings resembling dates be converted to Dates?
'  DateOrder                : The date order respected by the contents of strIn. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  DateSeparator            : The date separator used by the input.
'  SysDateOrder             : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  SysDateSeparator         : The Windows system date separator.
'  ShowDatesAsDatesToLongs: If TRUE and ShowDatesAsDates is also true then date-like strings are converted to Longs.
'Booleans
'  ShowLogicalsAsLogicals   : Should strings "TRUE" and "FALSE" (case insensitive) be converted to Booleans.
'Errors
'  ShowErrorsAsErrors       : Should strings that match how errors are represented in Excel worksheets be converted to
'                             those errors values?
'Missings
'  ShowMissingAsNullString  : Should Missings (consecutive delimiters in the file) be represented as zero-length strings?
'  ShowMissingsAs           : Ignored if ShowMissingAsNullString is true, otherwise a variant into which missings are
'                             converted.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CastToVariant(strIn As String, ShowNumbersAsNumbers As Boolean, SepsStandard As Boolean, _
    DecimalSeparator As String, SysDecimalSeparator As String, _
    ShowDatesAsDates As Boolean, DateOrder As Long, DateSeparator As String, SysDateOrder As Long, _
    SysDateSeparator As String, ShowLogicalsAsLogicals As Boolean, _
    ShowErrorsAsErrors As Boolean, ShowMissingAsNullString As Boolean, ShowMissingsAs As Variant)

    Dim Converted As Boolean
    Dim dblResult As Double
    Dim dtResult As Date
    Dim bResult As Boolean
    Dim eResult As Variant

    If ShowNumbersAsNumbers Then
        CastToDouble strIn, dblResult, SepsStandard, DecimalSeparator, SysDecimalSeparator, Converted
        If Converted Then
            CastToVariant = dblResult
            Exit Function
        End If
    End If

    If ShowDatesAsDates Then
        CastToDate strIn, dtResult, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
        If Converted Then
            CastToVariant = dtResult
            Exit Function
        End If
    End If

    If ShowLogicalsAsLogicals Then
        CastToBool strIn, bResult, Converted
        If Converted Then
            CastToVariant = bResult
            Exit Function
        End If
    End If

    If ShowErrorsAsErrors Then
        CastToError strIn, eResult, Converted
        If Converted Then
            CastToVariant = eResult
            Exit Function
        End If
    End If

    If Not ShowMissingAsNullString Then
        If Len(strIn) = 0 Then
            CastToVariant = ShowMissingsAs
            Exit Function
        End If
    End If

    CastToVariant = strIn
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble
' Purpose    : Casts strIn to double where strIn has specified decimals separator.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDouble(strIn As String, ByRef dblOut As Double, SepsStandard As Boolean, DecimalSeparator As String, _
    SysDecimalSeparator As String, ByRef Converted As Boolean)
    
    On Error GoTo ErrHandler
    If SepsStandard Then
        dblOut = CDbl(strIn)
    Else
        dblOut = CDbl(Replace(strIn, DecimalSeparator, SysDecimalSeparator))
    End If
    Converted = True
ErrHandler:
    'Do nothing - strIn was not a string representing a number.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDate
' Purpose    : In-place conversion of a string that looks like a date into a Long or Date. No error if string cannot be
'              converted to date.
' Parameters :
'  strIn           : String
'  dtOut           : Result of cast
'  DateOrder       : The date order respected by the contents of strIn. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D
'  DateSeparator   : The date separator used by the input
'  SysDateOrder    : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D
'  SysDateSeparator: The Windows system date separator
'  Converted       : Boolean flipped to TRUE if conversion takes place
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDate(strIn As String, ByRef dtOut As Date, DateOrder As Long, DateSeparator As String, _
    SysDateOrder As Long, SysDateSeparator As String, ByRef Converted As Boolean)
    
    Dim d As String
    Dim m As String
    Dim pos1 As Long
    Dim pos2 As Long
    Dim y As String
    
    On Error GoTo ErrHandler
    pos1 = InStr(strIn, DateSeparator)
    If pos1 = 0 Then Exit Sub
    pos2 = InStr(pos1 + 1, strIn, DateSeparator)
    If pos2 = 0 Then Exit Sub

    If DateOrder = 0 Then
        m = Left$(strIn, pos1 - 1)
        d = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        y = Mid$(strIn, pos2 + 1)
    ElseIf DateOrder = 1 Then
        d = Left$(strIn, pos1 - 1)
        m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        y = Mid$(strIn, pos2 + 1)
    ElseIf DateOrder = 2 Then
        y = Left$(strIn, pos1 - 1)
        m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        d = Mid$(strIn, pos2 + 1)
    Else
        Throw "DateOrder must be 0, 1, or 2"
    End If
    If SysDateOrder = 0 Then
        dtOut = CDate(m + SysDateSeparator + d + SysDateSeparator + y)
        Converted = True
    ElseIf SysDateOrder = 1 Then
        dtOut = CDate(d + SysDateSeparator + m + SysDateSeparator + y)
        Converted = True
    ElseIf SysDateOrder = 2 Then
        dtOut = CDate(y + SysDateSeparator + m + SysDateSeparator + d)
        Converted = True
    End If

    Exit Sub
ErrHandler:
    'Do nothing - was not a string representing a date with the specified date order and date separator.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToBool
' Purpose    : Convert string to Boolean, case insensitive.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CastToBool(strIn As String, ByRef bOut As Boolean, ByRef Converted)
    Dim l As Long
    If VarType(strIn) = vbString Then
        l = Len(strIn)
        If l = 4 Then
            If StrComp(strIn, "true", vbTextCompare) = 0 Then
                bOut = True
                Converted = True
            End If
        ElseIf l = 5 Then
            If StrComp(strIn, "false", vbTextCompare) = 0 Then
                bOut = False
                Converted = True
            End If
        End If
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToError
' Purpose    : Convert the string representation of Excel errors back to Excel errors.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToError(strIn As String, ByRef eOut As Variant, ByRef Converted As Boolean)
    On Error GoTo ErrHandler
    If Left(strIn, 1) = "#" Then
        Converted = True
        Select Case strIn 'Editing this function? Then its inverse function Encode!!!!
            Case "#DIV/0!"
                eOut = CVErr(xlErrDiv0)
            Case "#NAME?"
                eOut = CVErr(xlErrName)
            Case "#REF!"
                eOut = CVErr(xlErrRef)
            Case "#NUM!"
                eOut = CVErr(xlErrNum)
            Case "#NULL!"
                eOut = CVErr(xlErrNull)
            Case "#N/A"
                eOut = CVErr(xlErrNA)
            Case "#VALUE!"
                eOut = CVErr(xlErrValue)
            Case "#SPILL!"
                eOut = CVErr(2045)    'CVErr(xlErrNoSpill)'These constants introduced in Excel 2016
            Case "#BLOCKED!"
                eOut = CVErr(2047)    'CVErr(xlErrBlocked)
            Case "#CONNECT!"
                eOut = CVErr(2046)    'CVErr(xlErrConnect)
            Case "#UNKNOWN!"
                eOut = CVErr(2048)    'CVErr(xlErrUnknown)
            Case "#GETTING_DATA!"
                eOut = CVErr(2043)    'CVErr(xlErrGettingData)
            Case "#FIELD!"
                eOut = CVErr(2049)    'CVErr(xlErrField)
            Case "#CALC!"
                eOut = CVErr(2050)    'CVErr(xlErrField)
            Case Else
                Converted = False
        End Select
    End If

    Exit Sub
ErrHandler:
    Throw "#CastToError: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsUnicodeFile
' Purpose    : Tests if a file is unicode by reading the byte-order-mark. Return is True, False or an error is raised.
'              Adapted from
'              https://stackoverflow.com/questions/36188224/vba-test-encoding-of-a-text-file
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsUnicodeFile(FilePath As String)
    Static FSO As Scripting.FileSystemObject
    Dim T As Scripting.TextStream

    Dim intAsc1Chr As Long
    Dim intAsc2Chr As Long

    On Error GoTo ErrHandler
    If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
    If (FSO.FileExists(FilePath) = False) Then
        IsUnicodeFile = "#File not found!"
        Exit Function
    End If

    ' 1=Read-only, False=do not create if not exist, -1=Unicode 0=ASCII
    Set T = FSO.OpenTextFile(FilePath, 1, False, 0)
    If T.atEndOfStream Then
        T.Close: Set T = Nothing
        IsUnicodeFile = False
        Exit Function
    End If
    intAsc1Chr = Asc(T.Read(1))
    If T.atEndOfStream Then
        T.Close: Set T = Nothing
        IsUnicodeFile = False
        Exit Function
    End If
    intAsc2Chr = Asc(T.Read(1))
    T.Close
    If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
        IsUnicodeFile = True
    Else
        IsUnicodeFile = False
    End If

    Exit Function
ErrHandler:
    Throw "#IsUnicodeFile: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FunctionWizardActive
' Purpose    : Test if the Function wizard is active to allow early exit in slow functions.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FunctionWizardActive() As Boolean
    
    On Error GoTo ErrHandler
    If TypeName(Application.Caller) = "Range" Then
        If Not Application.CommandBars("Standard").Controls(1).Enabled Then
            FunctionWizardActive = True
        End If
    End If

    Exit Function
ErrHandler:
    Throw "#FunctionWizardActive: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferDelimiter
' Purpose    : Infer the delimiter in a file by looking for first occurrence outside quoted regions of comma, tab,
'              semi-colon, colon or vertical bar (|)
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferDelimiter(FileName As String, Unicode As Boolean)
    
    Const CHUNK_SIZE = 1000
    Dim Contents As String
    Dim CopyOfErr As String
    Dim EvenQuotes As Boolean
    Dim F As Scripting.File
    Dim FoundInEven As Boolean
    Dim FSO As Scripting.FileSystemObject
    Dim i As Long
    Dim T As TextStream
    Const QuoteChar As String = """"

    On Error GoTo ErrHandler

    Set FSO = New FileSystemObject
    Set F = FSO.GetFile(FileName)
    Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))

    If T.atEndOfStream Then
        T.Close: Set T = Nothing: Set F = Nothing
        Throw "File is empty"
    End If

    EvenQuotes = True
    While Not T.atEndOfStream
        Contents = T.Read(CHUNK_SIZE)
        For i = 1 To Len(Contents)
            Select Case Mid$(Contents, i, 1)
                Case QuoteChar
                    EvenQuotes = Not EvenQuotes
                Case ",", vbTab, "|", ";", ":"
                    If EvenQuotes Then
                        InferDelimiter = Mid$(Contents, i, 1)
                        T.Close: Set T = Nothing: Set F = Nothing
                        Exit Function
                    Else
                        FoundInEven = True
                    End If
            End Select
        Next i
    Wend

    'No commonly-used delimiter found in the file outside quoted regions. There are two possibilities: _
    either the file has only one column or some other character has been used, returning comma is _
        equivalent to assuming the former.

    InferDelimiter = ","

    Exit Function
ErrHandler:
    CopyOfErr = "#InferDelimiter: " & Err.Description & "!"
    If Not T Is Nothing Then
        T.Close
        Set T = Nothing: Set F = Nothing: Set FSO = Nothing
    End If
    Throw CopyOfErr
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDateFormat
' Purpose    :
' Parameters :
'  DateFormat   : String such as D/M/Y or Y-M-D
'  DateOrder    : ByRef argument is set to DateFormat using same convention as Application.International(xlDateOrder)
'                 (0 = MDY, 1 = DMY, 2 = YMD)
'  DateSeparator: ByRef argument is set to the DateSeparator, typically "-" or "/"
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseDateFormat(ByVal DateFormat As String, ByRef DateOrder As Long, ByRef DateSeparator As String)

    Const Err_DateFormat = "DateFormat should be 'M-D-Y', 'D-M-Y' or 'Y-M-D'. A character other " + _
        "than '-' is allowed as the separator. Omit to use the Windows default, which on this PC is "

    On Error GoTo ErrHandler

    'Replace repeated D's with a single D, etc since sParseDateCore only needs _
     to know the order in which the three parts of the date appear.
    If Len(DateFormat) > 5 Then
        DateFormat = UCase(DateFormat)
        ReplaceRepeats DateFormat, "D"
        ReplaceRepeats DateFormat, "M"
        ReplaceRepeats DateFormat, "Y"
    End If

    If Len(DateFormat) = 0 Then
        DateOrder = Application.International(xlDateOrder)
        DateSeparator = Application.International(xlDateSeparator)
    ElseIf Len(DateFormat) <> 5 Then
        Throw Err_DateFormat + WindowsDefaultDateFormat
    ElseIf Mid$(DateFormat, 2, 1) <> Mid$(DateFormat, 4, 1) Then
        Throw Err_DateFormat + WindowsDefaultDateFormat
    Else
        DateSeparator = Mid$(DateFormat, 2, 1)
        Select Case UCase$(Left$(DateFormat, 1) + Mid$(DateFormat, 3, 1) + Right$(DateFormat, 1))
            Case "MDY"
                DateOrder = 0
            Case "DMY"
                DateOrder = 1
            Case "YMD"
                DateOrder = 2
            Case Else
                Throw Err_DateFormat + WindowsDefaultDateFormat
        End Select
    End If

    Exit Sub
ErrHandler:
    Throw "#ParseDateFormat: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReplaceRepeats
' Purpose    : Replace repeated instances of a character in a string with a single instance.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ReplaceRepeats(ByRef TheString As String, TheChar As String)
    Dim ChCh As String
    ChCh = TheChar & TheChar
    While InStr(TheString, ChCh) > 0
        TheString = Replace(TheString, ChCh, TheChar)
    Wend
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WindowsDefaultDateFormat
' Purpose    : Returns a description of the system date format, used only for error message generation.
' -----------------------------------------------------------------------------------------------------------------------
Private Function WindowsDefaultDateFormat() As String
    Dim DS As String
    On Error GoTo ErrHandler
    DS = Application.International(xlDateSeparator)
    Select Case Application.International(xlDateOrder)
        Case 0
            WindowsDefaultDateFormat = "M" + DS + "D" + DS + "Y"
        Case 1
            WindowsDefaultDateFormat = "D" + DS + "M" + DS + "Y"
        Case 2
            WindowsDefaultDateFormat = "Y" + DS + "M" + DS + "D"
    End Select

    Exit Function
ErrHandler:
    WindowsDefaultDateFormat = "Cannot determine!"
End Function



' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ShowTextFile
' Purpose    : Parse any text file to a 1-column two-dimensional array of strings. No splitting into columns and no
'              casting.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ShowTextFile(FileName, StartRow As Long, NumRows As Long, _
    FileIsUnicode As Boolean)

    Dim FSO As Scripting.FileSystemObject
    Dim F As Scripting.File
    Dim T As Scripting.TextStream
    Dim ReadAll As String
    Dim Contents1D() As String
    Dim Contents2D() As String
    Dim i As Long

    On Error GoTo ErrHandler
    Set FSO = New FileSystemObject
    Set F = FSO.GetFile(FileName)

    Set T = F.OpenAsTextStream(ForReading, IIf(FileIsUnicode, TristateTrue, TristateFalse))
    For i = 1 To StartRow - 1
        T.SkipLine
    Next

    If NumRows = 0 Then
        ReadAll = T.ReadAll
        T.Close: Set T = Nothing: Set F = Nothing: Set FSO = Nothing

        ReadAll = Replace(ReadAll, vbCrLf, vbLf)
        ReadAll = Replace(ReadAll, vbCr, vbLf)

        'Text files may or may not be terminated by EOL characters...
        If Right$(ReadAll, 1) = vbLf Then
            ReadAll = Left$(ReadAll, Len(ReadAll) - 1)
        End If

        If Len(ReadAll) = 0 Then
            ReDim Contents1D(0 To 0)
        Else
            Contents1D = VBA.Split(ReadAll, vbLf)
        End If
        ReDim Contents2D(1 To UBound(Contents1D) - LBound(Contents1D) + 1, 1 To 1)
        For i = LBound(Contents1D) To UBound(Contents1D)
            Contents2D(i + 1, 1) = Contents1D(i)
        Next i
    Else
        ReDim Contents2D(1 To NumRows, 1 To 1)

        For i = 1 To NumRows 'BUG, won't work for Mac files. TODO Fix this?
        
            Contents2D(i, 1) = T.ReadLine
        Next i

        T.Close: Set T = Nothing: Set F = Nothing: Set FSO = Nothing
    End If

    ShowTextFile = Contents2D

    Exit Function
ErrHandler:
    Throw "#ShowTextFile: " & Err.Description & "!"
End Function


