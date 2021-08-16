Attribute VB_Name = "modCSVReadWrite"

' VBA-CSV

' Copyright (C) 2021 - PGS62 (https://github.com/PGS62/VBA-CSV )
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme

Option Explicit

'---------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead
' Purpose    : Register the function CSVRead with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
'---------------------------------------------------------------------------------------------------------
Sub RegisterCSVRead()
    Const Description = "Returns the contents of a comma-separated file on disk as an array."
    Dim ArgumentDescriptions() As String
    ReDim ArgumentDescriptions(1 To 13)
    ArgumentDescriptions(1) = "The full name of the file, including the path."
    ArgumentDescriptions(2) = "TRUE to convert Numbers, Dates, Booleans and Errors into their typed values; FALSE to leave as strings. Or string of letters N, D, B, E, Q, R. E.g. ""BN"" to convert just Booleans and numbers. R = quoted strings retain quotes. Q = convert quoted fields."
    ArgumentDescriptions(3) = "Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe found outside quoted regions within the first 10,000 characters. Enter FALSE to  see the file's contents as would be displayed in a text editor."
    ArgumentDescriptions(4) = "Whether delimiters which appear immediately after another delimiter, or at the start of a line, should be ignored while parsing; useful-for fixed-width files with delimiter padding between fields."
    ArgumentDescriptions(5) = "The format of dates in the file such as ""D-M-Y"", ""M-D-Y"" or ""Y/M/D"". If omitted a value is read from Windows regional settings. Repeated D's (or M's or Y's) are equivalent to single instances, so that ""d-m-y"" and ""DD-MMM-YYYY"" are equivalent."
    ArgumentDescriptions(6) = "Rows that start with this string will be skipped while parsing."
    ArgumentDescriptions(7) = "The row in the file at which reading starts. Optional and defaults to 1 to read from the first (not commented) row."
    ArgumentDescriptions(8) = "The column in the file at which reading starts. Optional and defaults to 1 to read from the first column."
    ArgumentDescriptions(9) = "The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the end of the file are read."
    ArgumentDescriptions(10) = "The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol are read."
    ArgumentDescriptions(11) = "Fields which are missing in the file (consecutive delimiters) are represented by ShowMissingsAs. Defaults to the null string, but can be any string or Empty."
    ArgumentDescriptions(12) = "Enter TRUE if the file is Unicode, FALSE otherwise. Omit to guess from the file's contents."
    ArgumentDescriptions(13) = "The character that represents a decimal point. If omitted, then the value from Windows regional settings is used."
    Application.MacroOptions "CSVRead", Description, , , , , , , , , ArgumentDescriptions
End Sub
'---------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVWrite
' Purpose    : Register the function CSVWrite with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
'---------------------------------------------------------------------------------------------------------
Sub RegisterCSVWrite()
    Const Description = "Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an ""error string"" (starts with #, ends with !) describing what went wrong."
    Dim ArgumentDescriptions() As String
    ReDim ArgumentDescriptions(1 To 8)
    ArgumentDescriptions(1) = "The full name of the file, including the path."
    ArgumentDescriptions(2) = "An array of data. Elements may be strings, numbers, dates, Booleans, empty, Excel errors or null values."
    ArgumentDescriptions(3) = "If TRUE (the default) then all strings in Data are quoted before being written to file. If FALSE only strings containing Delimiter, line feed, carriage return or double quote are quoted. Double quotes are always escaped by another double quote."
    ArgumentDescriptions(4) = "A format string that determine how dates, including cells formatted as dates, appear in the file. If omitted, defaults to ""yyyy-mm-dd""."
    ArgumentDescriptions(5) = "A format string that determines how dates with non-zero time part appear in the file. If omitted defaults to ""yyyy-mm-dd hh:mm:ss"".The companion function CSVRead is not capable of converting fields written in DateTime format back from strings into Dates."
    ArgumentDescriptions(6) = "The delimiter string, if omitted defaults to a comma. Delimiter may have more than one character."
    ArgumentDescriptions(7) = "If FALSE (the default) the file written will be encoded UTF-8. If TRUE the file written will be encoded UTF-16 LE BOM. An error will result if this argument is FALSE but Data contains strings with characters whose code points exceed 255."
    ArgumentDescriptions(8) = "Controls the line endings of the file written. Enter ""Windows"" (the default), ""Unix"" or ""Mac"". Also supports the line-ending characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13) or the strings ""CRLF"", ""LF"" or ""CR""."
    Application.MacroOptions "CSVWrite", Description, , , , , , , , , ArgumentDescriptions
End Sub

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVRead
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path.
' ConvertTypes: ConvertTypes provides control over whether fields in the file are converted to typed values
'             in the return or remain as strings, and also sets the treatment of "quoted fields", i.e.
'             fields with double-quotes as both their first and last characters.
'
'             ConvertTypes may take values FALSE (the default), TRUE, or a string of zero or more letters
'             from "NDBEQR".
'
'             If ConvertTypes is:
'             * FALSE then no conversion takes place other than quoted fields being unquoted.
'             * TRUE then unquoted numbers, dates, Booleans and errors are converted.
'
'             If ConvertTypes is a string including:
'             1) "N" then fields that represent numbers are converted to numbers (Doubles).
'             2) "D" then fields that represent dates (respecting DateFormat) are converted to Dates.
'             3) "B" then fields that read true or false are converted to Booleans. The match is not case
'             sensitive so TRUE, FALSE, True and False are also converted.
'             4) "E" then fields that match Excel"s representation of error values are converted to error
'             values. There are fourteen such strings, including #N/A, #NAME?, #VALUE! and #DIV/0!.
'             5) "Q" then conversion happens for both quoted and unquoted fields; otherwise only unquoted
'             fields are converted.
'             6) "R" then quoted fields retain their quotes, otherwise they are "unquoted" i.e. have their
'             leading and trailing characters removed and consecutive pairs of double-quotes replaced by a
'             single double quote.
' Delimiter : By default, CSVRead will try to detect a file's delimiter as the first instance of comma, tab,
'             semi-colon, colon or pipe found outside quoted regions in the first 10,000 characters of the
'             file. If it can't auto-detect the delimiter, it will assume comma. If your file includes a
'             different character or string delimiter you should pass that as the Delimiter argument.
'
'             Alternatively, enter FALSE as the delimiter to treat the file as "not a delimited file". In
'             this case the return will mimic how the file would appear in a text editor such as NotePad.
'             The file will by split into lines at all line breaks (irrespective of double-quotes) and each
'             element of the return will be a line of the file.
' IgnoreRepeated: Whether delimiters which appear immediately after another delimiter, or at the start of a
'             line, should be ignored while parsing; useful-for fixed-width files with delimiter padding
'             between fields.
' DateFormat: The format of dates in the file such as "D-M-Y", "M-D-Y" or "Y/M/D". If omitted a value is
'             read from Windows regional settings. Repeated D's (or M's or Y's) are equivalent to single
'             instances, so that "d-m-y" and "DD-MMM-YYYY" are equivalent.
' Comment   : Rows that start with this string will be skipped while parsing.
' SkipToRow : The row in the file at which reading starts. Optional and defaults to 1 to read from the first
'             (not commented) row.
' SkipToCol : The column in the file at which reading starts. Optional and defaults to 1 to read from the
'             first column.
' NumRows   : The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the
'             end of the file are read.
' NumCols   : The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol
'             are read.
' ShowMissingsAs: Fields which are missing in the file (consecutive delimiters) are represented by
'             ShowMissingsAs. Defaults to the null string, but can be any string or Empty. If NumRows is
'             greater than the number of rows in the file then the return is "padded" with the value of
'             ShowMissingsAs. Likewise if NumCols is greater than the number of columns in the file.
' Unicode   : In most cases, this argument can be omitted, in which case CSVRead will examine the file's
'             byte order mark to guess how the file should be opened - i.e. the correct "format" argument
'             to pass to VBA's method OpenAsTextStream, which requires an argument indicating whether the
'             file is Unicode or ASCII. Alternatively, enter TRUE for a Unicode file or FALSE for an ASCII
'             file.
' DecimalSeparator: In many places in the world, floating point number decimals are separated with a comma
'             instead of a period (3,14 vs. 3.14). CSVRead can correctly parse these numbers by passing in
'             the DecimalSeparator as a comma, in which case comma ceases to be a candidate if the parser
'             needs to guess the Delimiter.
'
' Notes     : See also companion function CSVRead.
'
'             For definition of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
'---------------------------------------------------------------------------------------------------------
Public Function CSVRead(FileName As String, Optional ConvertTypes As Variant = False, _
    Optional ByVal Delimiter As Variant, Optional IgnoreRepeated As Boolean, _
    Optional DateFormat As String, Optional Comment As String, Optional ByVal SkipToRow As Long = 1, _
    Optional ByVal SkipToCol As Long = 1, Optional ByVal NumRows As Long = 0, _
    Optional ByVal NumCols As Long = 0, Optional ByVal ShowMissingsAs As Variant = "", _
    Optional ByVal Unicode As Variant, Optional DecimalSeparator As String = vbNullString)

    Const DQ = """"
    Const Err_Delimiter = "Delimiter character must be passed as a string, FALSE for no delimiter. Omit to guess from file contents"
    Const Err_Delimiter2 = "Delimiter must have at least one character and cannot start with a double quote, line feed or carriage return"
    Const Err_FileEmpty = "File is empty"
    Const Err_Unicode = "Unicode must be passed as TRUE or FALSE to indicate whether the file is Unicode encoded. Omit to guess from the file's contents"
    Const Err_FunctionWizard = "Disabled in Function Wizard"
    Const Err_NumCols = "NumCols must be positive to read a given number of columns, or zero or omitted to read all columns from SkipToCol to the maximum column encountered."
    Const Err_NumRows = "NumRows must be positive to read a given number of rows, or zero or omitted to read all rows from SkipToRow to the end of the file."
    Const Err_Seps1 = "DecimalSeparator must be a single character"
    Const Err_Seps2 = "DecimalSpearator must not be equal to the first character of Delimiter or to a line-feed or carriage-return"
    Const Err_ShowMissings = "ShowMissingsAs has an illegal value, such as an array or an object"
    Const Err_SkipToCol = "SkipToCol must be at least 1."
    Const Err_SkipToRow = "SkipToRow must be at least 1."
    Const Err_Comment = "Comment must not contain double-quote, line feed or carriage return"
    
    Dim AnyConversion As Boolean
    Dim ColIndexes() As Long
    Dim ConvertQuoted As Boolean
    Dim CSVContents As String
    Dim DateOrder As Long
    Dim DateSeparator As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Lengths() As Long
    Dim m As Long
    Dim NeedToFill As Boolean
    Dim NotDelimited As Boolean
    Dim NumColsFound As Long
    Dim NumColsInReturn As Long
    Dim NumFields As Long
    Dim NumRowsFound As Long
    Dim NumRowsInReturn As Long
    Dim QuoteCounts() As Long
    Dim Ragged As Boolean
    Dim RetainQuotes As Boolean
    Dim ReturnArray() As Variant
    Dim RowIndexes() As Long
    Dim SepStandard As Boolean
    Dim SF As Scripting.File
    Dim SFSO As New Scripting.FileSystemObject
    Dim ShowBooleansAsBooleans As Boolean
    Dim ShowDatesAsDates As Boolean
    Dim ShowErrorsAsErrors As Boolean
    Dim ShowMissingsAsEmpty As Boolean
    Dim ShowNumbersAsNumbers As Boolean
    Dim Starts() As Long
    Dim strDelimiter As String
    Dim STS As Scripting.TextStream
    Dim SysDateOrder As Long
    Dim SysDateSeparator As String
    Dim SysDecimalSeparator As String
    
    On Error GoTo ErrHandler

    'Parse and validate inputs...
    If IsEmpty(Unicode) Or IsMissing(Unicode) Then
        Unicode = IsFileUnicode(FileName)
    ElseIf VarType(Unicode) <> vbBoolean Then
        Throw Err_Unicode
    End If

    SysDecimalSeparator = Application.DecimalSeparator
    If DecimalSeparator = vbNullString Then DecimalSeparator = SysDecimalSeparator
    If DecimalSeparator = SysDecimalSeparator Then
        SepStandard = True
    ElseIf Len(DecimalSeparator) <> 1 Then
        Throw Err_Seps1
    ElseIf DecimalSeparator = strDelimiter Or DecimalSeparator = vbLf Or DecimalSeparator = vbCr Then
        Throw Err_Seps2
    End If

    If VarType(Delimiter) = vbBoolean Then
        If Not Delimiter Then
            NotDelimited = True
        Else
            Throw Err_Delimiter
        End If
    ElseIf VarType(Delimiter) = vbString Then
        If Len(Delimiter) = 0 Then
            strDelimiter = InferDelimiter(FileName, CBool(Unicode), DecimalSeparator)
        ElseIf Left(Delimiter, 1) = DQ Or Left(Delimiter, 1) = vbLf Or Left(Delimiter, 1) = vbCr Then
            Throw Err_Delimiter2
        Else
            strDelimiter = Delimiter
        End If
    ElseIf IsEmpty(Delimiter) Or IsMissing(Delimiter) Then
        strDelimiter = InferDelimiter(FileName, CBool(Unicode), DecimalSeparator)
    Else
        Throw Err_Delimiter
    End If

    ParseConvertTypes ConvertTypes, ShowNumbersAsNumbers, _
        ShowDatesAsDates, ShowBooleansAsBooleans, ShowErrorsAsErrors, ConvertQuoted, RetainQuotes

    If ShowDatesAsDates Then
        ParseDateFormat DateFormat, DateOrder, DateSeparator
        SysDateOrder = Application.International(xlDateOrder)
        SysDateSeparator = Application.International(xlDateSeparator)
    End If

    If SkipToRow < 1 Then Throw Err_SkipToRow
    If SkipToCol < 1 Then Throw Err_SkipToCol
    If NumRows < 0 Then Throw Err_NumRows
    If NumCols < 0 Then Throw Err_NumCols
    
    Select Case VarType(ShowMissingsAs)
        Case vbEmpty
            ShowMissingsAsEmpty = True
        Case Is < vbArray
        Case Else
            Throw Err_ShowMissings
    End Select
    
    If InStr(Comment, DQ) > 0 Or InStr(Comment, vbLf) Or InStr(Comment, crlf) > 0 Then Throw Err_Comment
    
    'End of input validation
          
    Set SF = SFSO.GetFile(FileName)
    If FunctionWizardActive() Then
        If SF.Size > 1000000 Then
            CSVRead = "#" & Err_FunctionWizard & "!"
            Exit Function
        End If
    End If
    
    Set STS = SF.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))
    
    If STS.AtEndOfStream Then Throw Err_FileEmpty
    
    If NotDelimited Then
        CSVRead = ShowTextFile(STS, SkipToRow, NumRows)
        Exit Function
    End If

    If STS.AtEndOfStream Then
        STS.Close: Set STS = Nothing: Set SF = Nothing
        Throw Err_FileEmpty
    End If
          
    If SkipToRow = 1 And NumRows = 0 Then
        CSVContents = STS.ReadAll
        STS.Close: Set STS = Nothing: Set SF = Nothing
        Call ParseCSVContents(CSVContents, DQ, strDelimiter, Comment, IgnoreRepeated, SkipToRow, NumRows, NumRowsFound, NumColsFound, NumFields, Ragged, _
            Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts)
    Else
        CSVContents = ParseCSVContents(STS, DQ, strDelimiter, Comment, IgnoreRepeated, SkipToRow, NumRows, NumRowsFound, NumColsFound, NumFields, Ragged, _
            Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts)
        STS.Close
    End If
    
    'Useful for debugging, TODO remove this block in due course
    ' Dim Chars() As String, Numbers() As Long
    ' ReDim Numbers(1 To Len(CSVContents))
    ' ReDim Chars(1 To Len(CSVContents))
    ' For i = 1 To Len(CSVContents)
    '     Chars(i, 1) = Mid(CSVContents, i, 1)
    '     Numbers(i, 1) = i
    ' Next i
    ' CSVRead = HStack(VStack(NumRowsFound, NumColsFound, NumFields, strDelimiter), Transpose(Starts), _
    '     Transpose(Lengths), Transpose(RowIndexes), Transpose(ColIndexes), Transpose(QuoteCounts), Numbers, Chars)
    ' Exit Function
        
    If NumCols = 0 Then
        NumColsInReturn = NumColsFound - SkipToCol + 1
        If NumColsInReturn <= 0 Then
            Throw "SkipToCol (" + CStr(SkipToCol) + ") exceeds the number of columns in the file (" + CStr(NumColsFound) + ")"
        End If
    Else
        NumColsInReturn = NumCols
    End If
    If NumRows = 0 Then
        NumRowsInReturn = NumRowsFound
    Else
        NumRowsInReturn = NumRows
    End If
        
    AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
        ShowBooleansAsBooleans Or ShowErrorsAsErrors
        
    ReDim ReturnArray(1 To NumRowsInReturn, 1 To NumColsInReturn)
        
    For k = 1 To NumFields
        i = RowIndexes(k)
        j = ColIndexes(k) - SkipToCol + 1
        If j >= 1 And j <= NumColsInReturn Then
        
            ReturnArray(i, j) = ConvertField(Mid$(CSVContents, Starts(k), Lengths(k)), AnyConversion, _
                Lengths(k), DQ, QuoteCounts(k), RetainQuotes, ConvertQuoted, _
                ShowNumbersAsNumbers, SepStandard, DecimalSeparator, SysDecimalSeparator, _
                ShowDatesAsDates, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, _
                ShowBooleansAsBooleans, ShowErrorsAsErrors, ShowMissingsAs)
            
            'File is ragged with variable number of fields per line...
            If Ragged Then
                If Not ShowMissingsAsEmpty Then
                    If k = NumFields Then
                        NeedToFill = j < NumColsInReturn
                    ElseIf RowIndexes(k + 1) > RowIndexes(k) Then
                        NeedToFill = j < NumColsInReturn
                    Else
                        NeedToFill = False
                    End If
                    If NeedToFill Then
                        For m = j + 1 To NumColsInReturn
                            ReturnArray(i, m) = ShowMissingsAs
                        Next m
                    End If
                End If
            End If
        End If
    Next k

    'Pad if necessary
    If Not ShowMissingsAsEmpty Then
        If NumColsInReturn > NumColsFound - SkipToCol + 1 Then
            For i = 1 To NumRowsInReturn
                For j = NumColsFound - SkipToCol + 2 To NumColsInReturn
                    ReturnArray(i, j) = ShowMissingsAs
                Next j
            Next i
        End If
        If NumRowsInReturn > NumRowsFound Then
            For i = NumRowsFound + 1 To NumRowsInReturn
                For j = 1 To NumColsInReturn
                    ReturnArray(i, j) = ShowMissingsAs
                Next j
            Next i
        End If
    End If

    CSVRead = ReturnArray

    Exit Function

ErrHandler:
    CSVRead = "#CSVRead: " & Err.Description & "!"
    If Not STS Is Nothing Then STS.Close
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseConvertTypes
' Purpose    : Parse the input ConvertTypes to set five Boolean flags which are passed by reference
' Parameters :
'  ConvertTypes        :
'  ShowNumbersAsNumbers  : Should fields in the file that look like numbers be returned as Numbers? (Doubles)
'  ShowDatesAsDates      : Should fields in the file that look like dates with the specified DateFormat be returned as Dates?
'  ShowBooleansAsBooleans: Should fields in the file that are TRUE or FALSE (case insensitive) be returned as Booleans?
'  ShowErrorsAsErrors    : Should fields in the file that look like Excel errors (#N/A #REF! etc) be returned as errors?
'  ConvertQuoted         : Should the four conversion rules above apply even to quoted fields?
'  RetainQuotes          : Should quotes in quoted fields be retained?
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseConvertTypes(ByVal ConvertTypes As Variant, ByRef ShowNumbersAsNumbers As Boolean, _
    ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
    ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef RetainQuotes As Boolean)

    Const Err_ConvertTypes = "ConvertTypes must be Boolean or string with allowed letters NDBEQR. " & _
        "'N' show numbers as numbers, 'D' show dates as dates, 'B' show Booleans " & _
        "as Booleans, 'E' show Excel errors as errors, 'Q' rules NDBE apply even to quoted fields, 'R' quoted fields retain their quotes, TRUE = NDBE (convert unquoted numbers, dates, Booleans and errors), FALSE = no conversion"
    Const Err_Quoted = "ConvertTypes cannot contain both 'Q' and 'R' since 'Q' indicates that type conversion should apply to quoted fields, but 'R' indicates that quoted fields should retain their quotes."
    Const Err_Quoted2 = "ConvertTypes is incorrect, 'Q' indicates that conversion should apply even to quoted fields, but none of 'N', 'D', 'B' or 'E' are present to indicate which type conversion to apply"
    Dim i As Long

    On Error GoTo ErrHandler
    
    If TypeName(ConvertTypes) = "Range" Then
        ConvertTypes = ConvertTypes.value
    End If
    
    If IsEmpty(ConvertTypes) Then ConvertTypes = False

    If VarType(ConvertTypes) = vbBoolean Then
        If ConvertTypes Then
            ShowNumbersAsNumbers = True
            ShowDatesAsDates = True
            ShowBooleansAsBooleans = True
            ShowErrorsAsErrors = True
            RetainQuotes = False
            ConvertQuoted = False
        Else
            ShowNumbersAsNumbers = False
            ShowDatesAsDates = False
            ShowBooleansAsBooleans = False
            ShowErrorsAsErrors = False
            RetainQuotes = False
            ConvertQuoted = False
        End If
    ElseIf VarType(ConvertTypes) = vbString Then
        ShowNumbersAsNumbers = False
        ShowDatesAsDates = False
        ShowBooleansAsBooleans = False
        ShowErrorsAsErrors = False
        RetainQuotes = False
        ConvertQuoted = False
        For i = 1 To Len(ConvertTypes)
            Select Case UCase(Mid(ConvertTypes, i, 1))
                Case "N"
                    ShowNumbersAsNumbers = True
                Case "D"
                    ShowDatesAsDates = True
                Case "B", "L" 'Booleans aka Logicals
                    ShowBooleansAsBooleans = True
                Case "E"
                    ShowErrorsAsErrors = True
                Case "Q"
                    ConvertQuoted = True
                Case "R"
                    RetainQuotes = True
                Case Else
                    Throw Err_ConvertTypes + " Found unrecognised character '" _
                        + Mid(ConvertTypes, i, 1) + "'"
            End Select
        Next i
    Else
        Throw Err_ConvertTypes
    End If
    
    If RetainQuotes And ConvertQuoted Then
        Throw Err_Quoted
    ElseIf ConvertQuoted And Not (ShowNumbersAsNumbers Or ShowDatesAsDates Or _
        ShowBooleansAsBooleans Or ShowErrorsAsErrors) Then
        Throw Err_Quoted2
    End If

    Exit Sub
ErrHandler:
    Throw "#ParseConvertTypes: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Min4
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
' Procedure  : IsFileUnicode
' Purpose    : Guesses whether a file needs to be opened with the "format" argument to File.OpenAsTextStream set to
'              TriStateTrue or TriStateFalse.
'              The documentation at
'              https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/openastextstream-method
'              is limited but I believe that:
'            * TriStateTrue needs to passed for files which (as reported by NotePad++) are encoded as either
'              "UTF-16 LE BOM" or "UTF-16 BE BOM"
'            * TristateFalse needs to be passed for files encoded as "UTF-8" or "ANSI"
'            * Files encoded "UTF-8 BOM" are not correctly handled by OpenAsTextStream: the characters of the byte order
'              mark (BOM) are returned as part of the file's contents.
'              This method is adapted from
'              https://stackoverflow.com/questions/36188224/vba-test-encoding-of-a-text-file
'              but with changes following experimentation using NotePad++ to create files with various encodings.
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsFileUnicode(FilePath As String) As Boolean

    Dim intAsc1Chr As Long
    Dim intAsc2Chr As Long
    Dim T As Scripting.TextStream
    Static FSO As Scripting.FileSystemObject

    On Error GoTo ErrHandler
    If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
    If (FSO.FileExists(FilePath) = False) Then
        Throw "File not found!"
    End If

    ' 1=Read-only, False=do not create if not exist, -1=Unicode 0=ASCII
    Set T = FSO.OpenTextFile(FilePath, 1, False, 0)
    If T.AtEndOfStream Then
        T.Close: Set T = Nothing
        IsFileUnicode = False
        Exit Function
    End If
    intAsc1Chr = Asc(T.Read(1))
    If T.AtEndOfStream Then
        T.Close: Set T = Nothing
        IsFileUnicode = False
        Exit Function
    End If
    intAsc2Chr = Asc(T.Read(1))
    T.Close
    
    If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
        'File is probably encoded UTF-16 LE BOM (little endian, with Byte Option Marker)
        IsFileUnicode = True
    ElseIf (intAsc1Chr = 254) And (intAsc2Chr = 255) Then
        'File is probably encoded UTF-16 BE BOM (big endian, with Byte Option Marker)
        IsFileUnicode = True
    Else
        'File is probably encoded UTF-8. Hopefully not UTF-8-BOM since VBA's File.OpenAsTextStream does not cope well with that.
        IsFileUnicode = False
    End If

    Exit Function
ErrHandler:
    Throw "#IsFileUnicode: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferDelimiter
' Purpose    : Infer the delimiter in a file by looking for first occurrence outside quoted regions of comma, tab,
'              semi-colon, colon or pipe (|). Only look in the first 10,000 characters, Would prefer to look at the first
'              10 lines, but that presents a problem for files with Mac line endings as T.ReadLine doesn't work for them...
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferDelimiter(FileName As String, Unicode As Boolean, DecimalSeparator As String)
    
    Const CHUNK_SIZE = 1000
    Const MAX_CHUNKS = 10
    Const QuoteChar As String = """"
    Dim Contents As String
    Dim CopyOfErr As String
    Dim EvenQuotes As Boolean
    Dim F As Scripting.File
    Dim FSO As Scripting.FileSystemObject
    Dim i As Long, j As Long
    Dim T As TextStream

    On Error GoTo ErrHandler

    Set FSO = New FileSystemObject
    Set F = FSO.GetFile(FileName)
    Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))

    If T.AtEndOfStream Then
        T.Close: Set T = Nothing: Set F = Nothing
        Throw "File is empty"
    End If

    EvenQuotes = True
    While Not T.AtEndOfStream And j <= MAX_CHUNKS
        j = j + 1
        Contents = T.Read(CHUNK_SIZE)
        For i = 1 To Len(Contents)
            Select Case Mid$(Contents, i, 1)
                Case QuoteChar
                    EvenQuotes = Not EvenQuotes
                Case ",", vbTab, "|", ";", ":"
                    If Mid$(Contents, i, 1) <> DecimalSeparator Then
                        If EvenQuotes Then
                            InferDelimiter = Mid$(Contents, i, 1)
                            T.Close: Set T = Nothing: Set F = Nothing
                            Exit Function
                        End If
                    End If
            End Select
        Next i
    Wend

    'No commonly-used delimiter found in the file outside quoted regions _
     and in the first MAX_CHUNKS * CHUNK_SIZE characters. Assume comma in this case.

    T.Close
    
    If DecimalSeparator = "," Then
        InferDelimiter = ";"
    Else
        InferDelimiter = ","
    End If

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
' Purpose    : Populate DateOrder and DateSeparator by parsing DateFormat.
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
                Throw Err_DateFormat + WindowsDefaultDateFormat()
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
' Purpose    : Returns a description of the system date format, used only for error string generation.
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
Private Function ShowTextFile(T As TextStream, StartRow As Long, NumRows As Long)

    Dim Contents1D() As String
    Dim Contents2D() As String
    Dim i As Long
    Dim ReadAll As String

    On Error GoTo ErrHandler

    For i = 1 To StartRow - 1
        T.SkipLine
    Next

    If NumRows = 0 Then
        ReadAll = T.ReadAll
        T.Close

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
            If T.AtEndOfStream Then Exit For
            Contents2D(i, 1) = T.ReadLine
        Next i

        T.Close
    End If

    ShowTextFile = Contents2D

    Exit Function
ErrHandler:
    Throw "#ShowTextFile: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCSVContents
' Purpose    : Parse the contents of a CSV file. Returns a string Buffer together with arrays which assist splitting Buffer
'              into a two-dimensional array.
' Parameters :
'  ContentsOrStream: The contents of a CSV file as a string, or else a Scripting.TextStream.
'  QuoteChar       : The quote character, usually ascii 34 ("), which allow fields to contain characters that would
'                    otherwise be significant to parsing, such as delimiters or new line characters.
'  Delimiter       : The string that separates fields within each line. Typically a single character, but needn't be.
'  SkipToRow       : Rows in the file prior to SkipToRow are ignored.
'  NumRows         : The number of rows to parse. 0 for all rows from SkipToRow to the end of the file.
'  NumRowsFound    : Set to the number of rows in the file that are on or after SkipToRow.
'  NumColsFound    : Set to the number of columns in the file, i.e. the maximum number of fields in any single line.
'  NumFields       : Set to the number of fields in the file that are on or after SkipToRow.  May be less than
'                    NumRowsFound times NumColsFound if not all lines have the same number of fields.
'  Ragged          : Set to True if not all rows of the file have the same number of fields.
'  Starts          : Set to an array of size at least NumFields. Element k gives the point in Buffer at which the
'                    kth field starts.
'  Lengths         : Set to an array of size at least NumFields. Element k gives the length of the kth field.
'  RowIndexes      : Set to an array of size at least NumFields. Element k gives the row at which the kth field should
'                    appear in the return from CSVRead.
'  ColIndexes      : Set to an array of size at least NumFields. Element k gives the column at which the kth field would
'                    appear in the return from CSVRead under the assumption that argument SkipToCol is 1.
'  QuoteCounts     : Set to an array of size at least NumFields. Element k gives the number of QuoteChars that appear in the
'                    kth field.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseCSVContents(ContentsOrStream As Variant, QuoteChar As String, Delimiter As String, Comment As String, IgnoreRepeated As Boolean, SkipToRow As Long, _
    NumRows As Long, ByRef NumRowsFound As Long, ByRef NumColsFound As Long, ByRef NumFields As Long, ByRef Ragged As Boolean, ByRef Starts() As Long, _
    ByRef Lengths() As Long, RowIndexes() As Long, ColIndexes() As Long, QuoteCounts() As Long) As String

    Const Err_ContentsOrStream = "ContentsOrStream must either be a string or a TextStream"
    Const Err_Delimiter = "Delimiter must not be the null string"
    Dim Buffer As String
    Dim BufferUpdatedTo As Long
    Dim CheckForComments As Boolean
    Dim ColNum As Long
    Dim EvenQuotes As Boolean
    Dim HaveReachedSkipToRow As Boolean
    Dim i As Long 'Index to read from Buffer
    Dim j As Long 'Index to write to Starts, Lengths, RowIndexes and ColIndexes
    Dim LComment As Long
    Dim LDlm As Long
    Dim NumRowsInFile As Long
    Dim OrigLen As Long
    Dim PosCR As Long
    Dim PosDL As Long
    Dim PosLF As Long
    Dim PosQC As Long
    Dim QuoteArray() As String
    Dim QuoteCount As Long
    Dim RowNum As Long
    Dim SearchFor() As String
    Dim Streaming As Boolean
    Dim T As Scripting.TextStream
    Dim tmp As Long
    Dim Which As Long

    On Error GoTo ErrHandler
    
    If VarType(ContentsOrStream) = vbString Then
        Buffer = ContentsOrStream
        Streaming = False
    ElseIf TypeName(ContentsOrStream) = "TextStream" Then
        Set T = ContentsOrStream
        If NumRows = 0 Then
            Buffer = T.ReadAll
            T.Close
            Streaming = False
        Else
            Call GetMoreFromStream(T, Delimiter, QuoteChar, Buffer, BufferUpdatedTo)
            Streaming = True
        End If
    Else
        Throw Err_ContentsOrStream
    End If
       
    LComment = Len(Comment)
    If LComment > 0 Then
        CheckForComments = True
    End If
       
    If Streaming Then
        ReDim SearchFor(1 To 4)
        SearchFor(1) = Delimiter
        SearchFor(2) = vbLf
        SearchFor(3) = vbCr
        SearchFor(4) = QuoteChar
        ReDim QuoteArray(1 To 1)
        QuoteArray(1) = QuoteChar
    End If

    ReDim Starts(1 To 8)
    ReDim Lengths(1 To 8)
    ReDim RowIndexes(1 To 8)
    ReDim ColIndexes(1 To 8)
    ReDim QuoteCounts(1 To 8)
    
    LDlm = Len(Delimiter)
    If LDlm = 0 Then Throw Err_Delimiter 'Avoid infinite loop!
    OrigLen = Len(Buffer)
    If Not Streaming Then
        'Ensure Buffer terminates with vbCrLf
        If Right(Buffer, 1) <> vbCr And Right(Buffer, 1) <> vbLf Then
            Buffer = Buffer + vbCrLf
        ElseIf Right(Buffer, 1) = vbCr Then
            Buffer = Buffer + vbLf
        End If
        BufferUpdatedTo = Len(Buffer)
    End If
    
    j = 1: i = 0
    
    If CheckForComments Then
        SkipComment Streaming, Comment, LComment, T, Delimiter, Buffer, i, QuoteChar, PosLF, PosCR, BufferUpdatedTo
    End If
    
    If IgnoreRepeated Then
        Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
            i = i + LDlm
        Loop
    End If
    
    ColNum = 1: RowNum = 1
    EvenQuotes = True
    Starts(1) = i + 1
    If SkipToRow = 1 Then HaveReachedSkipToRow = True

    Do
        If EvenQuotes Then
            If Not Streaming Then
                If PosDL <= i Then PosDL = InStr(i + 1, Buffer, Delimiter): If PosDL = 0 Then PosDL = BufferUpdatedTo + 1
                If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
                If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
                If PosQC <= i Then PosQC = InStr(i + 1, Buffer, QuoteChar): If PosQC = 0 Then PosQC = BufferUpdatedTo + 1
                i = Min4(PosDL, PosLF, PosCR, PosQC, Which)
            Else
                i = SearchInBuffer(SearchFor, i + 1, T, Delimiter, QuoteChar, Which, Buffer, BufferUpdatedTo)
            End If

            If i >= BufferUpdatedTo + 1 Then
                Exit Do
            End If

            If j + 1 > UBound(Starts) Then
                ReDim Preserve Starts(1 To UBound(Starts) * 2)
                ReDim Preserve Lengths(1 To UBound(Lengths) * 2)
                ReDim Preserve RowIndexes(1 To UBound(RowIndexes) * 2)
                ReDim Preserve ColIndexes(1 To UBound(ColIndexes) * 2)
                ReDim Preserve QuoteCounts(1 To UBound(QuoteCounts) * 2)
            End If

            Select Case Which
                Case 1
                    'Found Delimiter
                    Lengths(j) = i - Starts(j)
                    If IgnoreRepeated Then
                        Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
                            i = i + LDlm
                        Loop
                    End If
                    
                    Starts(j + 1) = i + LDlm
                    ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                    ColNum = ColNum + 1
                    QuoteCounts(j) = QuoteCount: QuoteCount = 0
                    j = j + 1
                    NumFields = NumFields + 1
                    i = i + LDlm - 1
                Case 2, 3
                    Lengths(j) = i - Starts(j) 'Line Ending
                    If Mid(Buffer, i + 1, 1) = vbLf Then
                        'Ending is Windows rather than Mac or Unix.
                        i = i + 1
                    End If
                    
                    If CheckForComments Then
                        SkipComment Streaming, Comment, LComment, T, Delimiter, Buffer, i, QuoteChar, PosLF, PosCR, BufferUpdatedTo
                    End If
                    
                    If IgnoreRepeated Then
                        Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
                            i = i + LDlm
                        Loop
                    End If
                    Starts(j + 1) = i + 1

                    If ColNum > NumColsFound Then
                        If NumColsFound > 0 Then
                            Ragged = True
                        End If
                        NumColsFound = ColNum
                    ElseIf ColNum < NumColsFound Then
                        Ragged = True
                    End If
                    
                    ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                    ColNum = 1: RowNum = RowNum + 1
                    QuoteCounts(j) = QuoteCount: QuoteCount = 0
                    j = j + 1
                    NumFields = NumFields + 1
                    
                    If HaveReachedSkipToRow Then
                        If RowNum = NumRows + 1 Then
                            Exit Do
                        End If
                    Else
                        If RowNum = SkipToRow Then
                            HaveReachedSkipToRow = True
                            tmp = Starts(j)
                            ReDim Starts(1 To 8): ReDim Lengths(1 To 8): ReDim RowIndexes(1 To 8)
                            ReDim ColIndexes(1 To 8): ReDim QuoteCounts(1 To 8)
                            RowNum = 1: j = 1: NumFields = 0
                            Starts(1) = tmp
                        End If
                    End If
                Case 4
                    'Found QuoteChar
                    EvenQuotes = False
                    QuoteCount = QuoteCount + 1
            End Select
        Else
            If Not Streaming Then
                PosQC = InStr(i + 1, Buffer, QuoteChar)
            Else
                If PosQC <= i Then PosQC = SearchInBuffer(QuoteArray(), i + 1, T, Delimiter, QuoteChar, 0, Buffer, BufferUpdatedTo)
            End If
            
            If PosQC = 0 Then
                'Malformed Buffer (not RFC4180 compliant). There should always be an even number of double quotes. _
                 If there are an odd number then all text after the last double quote in the file will be (part of) _
                 the last field in the last line.
                Lengths(j) = OrigLen - Starts(j) + 1
                ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                
                RowNum = RowNum + 1
                If ColNum > NumColsFound Then NumColsFound = ColNum
                NumFields = NumFields + 1
                Exit Do
            Else
                i = PosQC
                EvenQuotes = True
                QuoteCount = QuoteCount + 1
            End If
        End If
    Loop

    NumRowsFound = RowNum - 1
    
    If HaveReachedSkipToRow Then
        NumRowsInFile = SkipToRow - 1 + RowNum - 1
    Else
        NumRowsInFile = RowNum - 1
    End If
    
    If SkipToRow > NumRowsInFile Then
        If NumRows = 0 Then 'Attempting to read from SkipToRow to the end of the file, but that would be zero or _
                             a negative number of rows. So throw an error.
            Throw "SkipToRow (" + CStr(SkipToRow) + ") exceeds the number of rows in the file (" + CStr(NumRowsInFile) + ")"
        Else
            'Attempting to read a set number of rows, function CSVRead will return an array of Empty values.
            NumFields = 0
            NumRowsFound = 0
        End If
    End If

    ParseCSVContents = Buffer

    Exit Function
ErrHandler:
    Throw "#ParseCSVContents: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SkipComment
' Author     : Philip Swannell
' Date       : 16-Aug-2021
' Purpose    : Skip a commented row by incrementing i to the position of the line feed just before the next not-commented
'              line.
' -----------------------------------------------------------------------------------------------------------------------
Sub SkipComment(Streaming As Boolean, Comment As String, LComment As Long, T As Scripting.TextStream, _
    Delimiter As String, ByRef Buffer As String, ByRef i As Long, QuoteChar As String, ByVal PosLF As Long, _
    ByVal PosCR As Long, ByRef BufferUpdatedTo As Long)
    Do
        If Streaming Then
            If i + LComment > BufferUpdatedTo Then
                If Not T.AtEndOfStream Then
                    Call GetMoreFromStream(T, Delimiter, QuoteChar, Buffer, BufferUpdatedTo)
                End If
            End If
        End If

        If Mid$(Buffer, i + 1, LComment) = Comment Then
            If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
            If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
            If PosLF < PosCR Then
                i = PosLF
            ElseIf PosLF = PosCR + 1 Then
                i = PosLF
            Else
                i = PosCR
            End If
        Else
            Exit Do
        End If
    Loop
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SearchInBuffer
' Purpose    : Returns the location in the buffer of the first-encountered string amongst the elements of SearchFor,
'              starting the search at point SearchFrom and finishing the search at point BufferUpdatedTo. If none found in
'              that region returns BufferUpdatedTo + 1. Otherwise returns the location of the first found and sets the
'              by-reference argument Which to indicate which element of SearchFor was the first to be found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SearchInBuffer(SearchFor() As String, StartingAt As Long, T As Scripting.TextStream, Delimiter As String, _
    QuoteChar As String, ByRef Which As Long, ByRef Buffer As String, ByRef BufferUpdatedTo As Long)

    Dim InstrRes As Long
    Dim PrevBufferUpdatedTo As Long

    On Error GoTo ErrHandler

    'in this call only search as far as BufferUpdatedTo
    InstrRes = InStrMulti(SearchFor, Buffer, StartingAt, BufferUpdatedTo, Which)
    If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
        SearchInBuffer = InstrRes
        Exit Function
    ElseIf T.AtEndOfStream Then
        SearchInBuffer = BufferUpdatedTo + 1
        Exit Function
    End If

    Do
        PrevBufferUpdatedTo = BufferUpdatedTo
        GetMoreFromStream T, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
        InstrRes = InStrMulti(SearchFor, Buffer, PrevBufferUpdatedTo + 1, BufferUpdatedTo, Which)
        If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
            SearchInBuffer = InstrRes
            Exit Function
        ElseIf T.AtEndOfStream Then
            SearchInBuffer = BufferUpdatedTo + 1
            Exit Function
        End If
    Loop
    Exit Function
ErrHandler:
    Throw "#SearchInBuffer: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InStrMulti
' Purpose    : Returns the first point in SearchWithin at which one of the elements of SearchFor is found, search is
'              restricted to region [StartingAt, EndingAt] and Which is updated with the index into SearchFor of the
'              first string found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function InStrMulti(SearchFor() As String, SearchWithin As String, StartingAt As Long, EndingAt As Long, _
    ByRef Which As Long)

    Const Inf = 2147483647
    Dim i As Long
    Dim InstrRes() As Long
    Dim LB As Long, UB As Long
    Dim Result As Long

    On Error GoTo ErrHandler
    LB = LBound(SearchFor): UB = UBound(SearchFor)

    Result = Inf

    ReDim InstrRes(LB To UB)
    For i = LB To UB
        InstrRes(i) = InStr(StartingAt, SearchWithin, SearchFor(i))
        If InstrRes(i) > 0 Then
            If InstrRes(i) <= EndingAt Then
                If InstrRes(i) < Result Then
                    Result = InstrRes(i)
                    Which = i
                End If
            End If
        End If
    Next
    InStrMulti = IIf(Result = Inf, 0, Result)

    Exit Function
ErrHandler:
    Throw "#InStrMulti: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetMoreFromStream
' Purpose    : Write CHUNKSIZE characters from the TextStream T into the buffer, modifying the passed-by-reference arguments
'              Buffer, BufferUpdatedTo and Streaming.
'              Complexities:
'           a) We have to be careful not to update the buffer to a point part-way through a two-character end-of-line or a
'              multi-character delimiter, otherwise calling method SearchInBuffer might give the wrong result.
'           b) We update a few characters of the buffer beyond the BufferUpdatedTo point with the delimiter, the QuoteChar
'              and vbCrLf. This ensures that the calls to Instr that search the buffer for these strings do not needlessly
'              scan the unupdated part of the buffer.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub GetMoreFromStream(T As Scripting.TextStream, Delimiter As String, QuoteChar As String, ByRef Buffer As String, _
    ByRef BufferUpdatedTo As Long)
    Const CHUNKSIZE = 5000  ' The number of characters to read from the stream on each call. _
                              Set to a small number for testing logic and a bigger number for _
                              performance, but not too high since a common use case is reading _
                              just the first line of a file. Suggest 5000? Note that when reading _
                              an entire file (NumRows argument to CSVRead is zero) function _
                              GetMoreFromStream is not called, instead the entire file is read _
                              with a single call to T.ReadAll.
    Dim ExpandBufferBy As Long
    Dim FirstPass As Boolean
    Dim i As Long
    Dim NCharsToWriteToBuffer As Long
    Dim NewChars
    Dim OKToExit

    On Error GoTo ErrHandler
    FirstPass = True
    Do
        NewChars = T.Read(IIf(FirstPass, CHUNKSIZE, 1))
        FirstPass = False
        If T.AtEndOfStream Then
            'Ensure NewChars terminates with vbCrLf
            If Right(NewChars, 1) <> vbCr And Right(NewChars, 1) <> vbLf Then
                NewChars = NewChars + vbCrLf
            ElseIf Right(NewChars, 1) = vbCr Then
                NewChars = NewChars + vbLf
            End If
        End If

        NCharsToWriteToBuffer = Len(NewChars) + Len(Delimiter) + 3

        If BufferUpdatedTo + NCharsToWriteToBuffer > Len(Buffer) Then
            ExpandBufferBy = MaxLngs(Len(Buffer), NCharsToWriteToBuffer)
            Buffer = Buffer & String(ExpandBufferBy, "?")
        End If
        
        Mid(Buffer, BufferUpdatedTo + 1, Len(NewChars)) = NewChars
        BufferUpdatedTo = BufferUpdatedTo + Len(NewChars)

        OKToExit = True
        'Ensure we don't leave the buffer updated to part way through a two-character end of line marker.
        If Right(NewChars, 1) = vbCr Then
            OKToExit = False
        End If
        'Ensure we don't leave the buffer updated to a point part-way through a multi-character delimiter
        If Len(Delimiter) > 1 Then
            For i = 1 To Len(Delimiter) - 1
                If Mid$(Buffer, BufferUpdatedTo - i + 1, i) = Left(Delimiter, i) Then
                    OKToExit = False
                    Exit For
                End If
            Next i
            If Mid(Buffer, BufferUpdatedTo - Len(Delimiter) + 1, Len(Delimiter)) = Delimiter Then
                OKToExit = True
            End If
        End If
        If OKToExit Then Exit Do
    Loop

    'Line below arranges that when calling Instr(Buffer,....) we don't pointlessly scan the space characters _
     we can be sure that there is space in the buffer to write the extra characters thanks to
    Mid(Buffer, BufferUpdatedTo + 1, Len(Delimiter) + 3) = vbCrLf & QuoteChar & Delimiter

    Exit Sub
ErrHandler:
    Throw "#GetMoreFromStream: " & Err.Description & "!"
End Sub

Private Function MaxLngs(x As Long, y As Long) As Long
    If x > y Then
        MaxLngs = x
    Else
        MaxLngs = y
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConvertField
' Purpose    : Process one field of the file, applying type conversion, quote removal etc.
' Parameters :
'  Field                    : The input string.
'Quote handling
'  QuoteCount               : How many quote characters does Field contain
'  RetainQuotes             : Should quoted fields keep their quote characters?
'  ConvertQuoted            : Should quoted fields (after quote removal) be converted according to args
'                             ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans and ShowErrorsAsErrors?
'Numbers
'  ShowNumbersAsNumbers     : If inStr is a string representation of a number should the function return that number?
'  SepStandard              : Is the decimal separator the same as the system defaults? If True then the next two
'                             arguments are ignored.
'  DecimalSeparator         : The decimal separator used in the input string.
'  SysDecimalSeparator      : The default decimal separator on the system.
'Dates
'  ShowDatesAsDates         : If inStr represents a date should the function return taht date?
'  DateOrder                : If inStr is a string representaiton of a date it will use this date order.
'                             0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  DateSeparator            : The date separator used by inStr, typically "-" or "/".
'  SysDateOrder             : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  SysDateSeparator         : The Windows system date separator.
'Booleans
'  ShowBooleansAsBooleans   : If inStr matches "TRUE" and "FALSE" (case insensitive) should the Boolean value be returned?
'Errors
'  ShowErrorsAsErrors       : Should strings that match how errors are represented in Excel worksheets be converted to
'                             those errors values?
' -----------------------------------------------------------------------------------------------------------------------
Private Function ConvertField(Field As String, AnyConversion As Boolean, FieldLength As Long, QuoteChar As String, QuoteCount As Long, RetainQuotes As Boolean, ConvertQuoted As Boolean, _
    ShowNumbersAsNumbers As Boolean, SepStandard As Boolean, _
    DecimalSeparator As String, SysDecimalSeparator As String, _
    ShowDatesAsDates As Boolean, DateOrder As Long, DateSeparator As String, SysDateOrder As Long, _
    SysDateSeparator As String, ShowBooleansAsBooleans As Boolean, _
    ShowErrorsAsErrors As Boolean, ShowMissingsAs As Variant)

    Dim bResult As Boolean
    Dim Converted As Boolean
    Dim dblResult As Double
    Dim dtResult As Date
    Dim eResult As Variant
    Dim isQuoted As Boolean

    If FieldLength = 0 Then
        ConvertField = ShowMissingsAs
        Exit Function
    ElseIf Left(Field, 1) = QuoteChar Then 'NOTE definition of quoted is both first and last characters must be quote characters
        If Right(QuoteChar, 1) = QuoteChar Then
            isQuoted = True
            If Not RetainQuotes Then
                Field = Mid$(Field, 2, FieldLength - 2)
                If QuoteCount > 2 Then
                    Field = Replace(Field, QuoteChar + QuoteChar, QuoteChar) 'TODO QuoteCharTwice arg
                End If
            End If
        End If
    End If

    If Not AnyConversion Then
        ConvertField = Field
        Exit Function
    ElseIf Not ConvertQuoted Then
        If QuoteCount > 0 Then
            ConvertField = Field
            Exit Function
        End If
    End If

    If ShowNumbersAsNumbers Then
        CastToDouble Field, dblResult, SepStandard, DecimalSeparator, SysDecimalSeparator, Converted
        If Converted Then
            ConvertField = dblResult
            Exit Function
        End If
    End If

    If ShowDatesAsDates Then
        CastToDate Field, dtResult, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
        If Converted Then
            ConvertField = dtResult
            Exit Function
        End If
    End If

    If ShowBooleansAsBooleans Then
        CastToBool Field, bResult, Converted
        If Converted Then
            ConvertField = bResult
            Exit Function
        End If
    End If

    If ShowErrorsAsErrors Then
        CastToError Field, eResult, Converted
        If Converted Then
            ConvertField = eResult
            Exit Function
        End If
    End If
    
    If Len(Field) = 0 Then
        ConvertField = ShowMissingsAs
        Exit Function
    End If

    ConvertField = Field
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble
' Purpose    : Casts strIn to double where strIn has specified decimals separator.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDouble(strIn As String, ByRef dblOut As Double, SepStandard As Boolean, DecimalSeparator As String, _
    SysDecimalSeparator As String, ByRef Converted As Boolean)
    
    On Error GoTo ErrHandler
    If SepStandard Then
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
    
    Dim D As String
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
        D = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        y = Mid$(strIn, pos2 + 1)
    ElseIf DateOrder = 1 Then
        D = Left$(strIn, pos1 - 1)
        m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        y = Mid$(strIn, pos2 + 1)
    ElseIf DateOrder = 2 Then
        y = Left$(strIn, pos1 - 1)
        m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        D = Mid$(strIn, pos2 + 1)
    Else
        Throw "DateOrder must be 0, 1, or 2"
    End If
    If SysDateOrder = 0 Then
        dtOut = CDate(m + SysDateSeparator + D + SysDateSeparator + y)
        Converted = True
    ElseIf SysDateOrder = 1 Then
        dtOut = CDate(D + SysDateSeparator + m + SysDateSeparator + y)
        Converted = True
    ElseIf SysDateOrder = 2 Then
        dtOut = CDate(y + SysDateSeparator + m + SysDateSeparator + D)
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
' Procedure  : OStoEOL
' Purpose    : Convert text describing an operating system to the end-of-line marker employed. Note that "Mac" converts
'              to vbCr but Apple operating systems since OSX use vbLf, matching Unix.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OStoEOL(OS As String, ArgName As String) As String

    Const Err_Invalid = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associented end of line characters."

    On Error GoTo ErrHandler
    Select Case LCase(OS)
        Case "windows", vbCrLf, "crlf"
            OStoEOL = vbCrLf
        Case "unix", "linux", vbLf, "lf"
            OStoEOL = vbLf
        Case "mac", vbCr, "cr"
            OStoEOL = vbCr
        Case Else
            Throw ArgName + Err_Invalid
    End Select

    Exit Function
ErrHandler:
    Throw "#OStoEOL: " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVWrite
' Purpose   : Creates a comma-separated file on disk containing Data. Any existing file of the same
'             name is overwritten. If successful, the function returns FileName, otherwise an "error
'             string" (starts with #, ends with !) describing what went wrong.
' Arguments
' FileName  : The full name of the file, including the path.
' Data      : An array of data. Elements may be strings, numbers, dates, Booleans, empty, Excel errors or
'             null values.
' QuoteAllStrings: If TRUE (the default) then all strings in Data are quoted before being written to file.
'             If FALSE only strings containing Delimiter, line feed, carriage return or double quote are
'             quoted. Double quotes are always escaped by another double quote.
' DateFormat: A format string that determine how dates, including cells formatted as dates, appear in the
'             file. If omitted, defaults to "yyyy-mm-dd".
' DateTimeFormat: A format string that determines how dates with non-zero time part appear in the file. If
'             omitted defaults to "yyyy-mm-dd hh:mm:ss".The companion function CSVRead is not capable of
'             converting fields written in DateTime format back from strings into Dates.
' Delimiter : The delimiter string, if omitted defaults to a comma. Delimiter may have more than one
'             character.
' Unicode   : If FALSE (the default) the file written will be encoded UTF-8. If TRUE the file written will
'             be encoded UTF-16 LE BOM. An error will result if this argument is FALSE but Data contains
'             strings with characters whose code points exceed 255.
' EOL       : Controls the line endings of the file written. Enter "Windows" (the default), "Unix" or "Mac".
'             Also supports the line-ending characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13)
'             or the strings "CRLF", "LF" or "CR". The last line of the file is written with a line ending.
'
' Notes     : See also companion function CSVRead.
'
'             For definition of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
'---------------------------------------------------------------------------------------------------------
Public Function CSVWrite(FileName As String, ByVal Data As Variant, _
    Optional QuoteAllStrings As Boolean = True, Optional DateFormat As String = "yyyy-mm-dd", _
    Optional DateTimeFormat As String = "yyyy-mm-dd hh:mm:ss", _
    Optional Delimiter As String = ",", Optional Unicode As Boolean, _
    Optional ByVal EOL As String = vbCrLf)
Attribute CSVWrite.VB_Description = "Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an ""error string"" (starts with #, ends with !) describing what went wrong."
Attribute CSVWrite.VB_ProcData.VB_Invoke_Func = " \n14"

    Const DQ = """"
    Dim EOLIsWindows As Boolean
    Dim FSO As Scripting.FileSystemObject
    Dim i As Long
    Dim j As Long
    Dim OneLine() As String
    Dim OneLineJoined As String
    Dim T As Scripting.TextStream
    
    Const Err_Delimiter = "Delimiter must have at least one character and cannot start with a double quote, line feed or carriage return"
    Const Err_Dimensions = "Data must be a range or a 2-dimensional array"

    On Error GoTo ErrHandler

    EOL = OStoEOL(EOL, "EOL")
    EOLIsWindows = EOL = vbCrLf

    If Len(Delimiter) = 0 Or Left(Delimiter, 1) = DQ Or Left(Delimiter, 1) = vbLf Or Left(Delimiter, 1) = vbCr Then
        Throw Err_Delimiter
    End If

    If TypeName(Data) = "Range" Then
        'Preserve elements of type Date by using .Value, not .Value2
        Data = Data.value
    End If
    
    If NumDimensions(Data) <> 2 Then Throw Err_Dimensions
    
    Set FSO = New FileSystemObject
    Set T = FSO.CreateTextFile(FileName, True, Unicode)

    ReDim OneLine(LBound(Data, 2) To UBound(Data, 2))

    For i = LBound(Data) To UBound(Data)
        For j = LBound(Data, 2) To UBound(Data, 2)
            OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat)
        Next j
        OneLineJoined = VBA.Join(OneLine, Delimiter)
        WriteLineWrap T, OneLineJoined, EOLIsWindows, EOL, Unicode
    Next i

    T.Close: Set T = Nothing: Set FSO = Nothing
    CSVWrite = FileName
    Exit Function
ErrHandler:
    CSVWrite = "#CSVWrite: " & Err.Description & "!"
    If Not T Is Nothing Then Set T = Nothing: Set FSO = Nothing

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WriteLineWrap
' Purpose    : Wrapper to TextStream.Write[Line] to give more informative error message than "invalid procedure call or
'              argument" if the error is caused by attempting to write characters with code>255 to a stream opened with TriStateFalse.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub WriteLineWrap(T As TextStream, text As String, EOLIsWindows As Boolean, EOL As String, Unicode As Boolean)

    Dim ErrDesc As String
    Dim ErrNum As Long
    Dim i As Long

    On Error GoTo ErrHandler
    If EOLIsWindows Then
        T.WriteLine text
    Else
        T.Write text
        T.Write EOL
    End If

    Exit Sub

ErrHandler:
    ErrNum = Err.Number
    ErrDesc = Err.Description
    If Not Unicode Then
        If ErrNum = 5 Then
            For i = 1 To Len(text)
                If AscW(Mid(text, i, 1)) > 255 Then
                    ErrDesc = "Data contains characters with code points above 255 (first found has code " & CStr(AscW(Mid(text, i, 1))) & _
                        ") which cannot be written to an ascii file. Try calling CSVWrite with argument Unicode as True"
                    Exit For
                End If
            Next i
        End If
    End If
    Throw "#WriteLineWrap: " & ErrDesc & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Encode
' Purpose    : Encode arbitrary value as a string, sub-routine of CSVWrite.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Encode(x As Variant, QuoteAllStrings As Boolean, DateFormat As String, DateTimeFormat As String) As String
    Const DQ = """"
    Const DQ2 = """"""

    On Error GoTo ErrHandler
    Select Case VarType(x)

        Case vbString
            If InStr(x, DQ) > 0 Then
                Encode = DQ + Replace(x, DQ, DQ2) + DQ
            ElseIf QuoteAllStrings Then
                Encode = DQ + x + DQ
            ElseIf InStr(x, vbCr) > 0 Then
                Encode = DQ + x + DQ
            ElseIf InStr(x, vbLf) > 0 Then
                Encode = DQ + x + DQ
            ElseIf InStr(x, ",") > 0 Then
                Encode = DQ + x + DQ
            Else
                Encode = x
            End If
        Case vbBoolean, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbLongLong, vbEmpty
            Encode = CStr(x)
        Case vbDate
            If CLng(x) = CDbl(x) Then
                Encode = Format$(x, DateFormat)
            Else
                Encode = Format$(x, DateTimeFormat)
            End If
        Case vbNull
            Encode = "NULL"
        Case vbError
            Select Case CStr(x) 'Editing this case statement? Edit also its inverse - CastToError
                Case "Error 2000"
                    Encode = "#NULL!"
                Case "Error 2007"
                    Encode = "#DIV/0!"
                Case "Error 2015"
                    Encode = "#VALUE!"
                Case "Error 2023"
                    Encode = "#REF!"
                Case "Error 2029"
                    Encode = "#NAME?"
                Case "Error 2036"
                    Encode = "#NUM!"
                Case "Error 2042"
                    Encode = "#N/A"
                Case "Error 2043"
                    Encode = "#GETTING_DATA!"
                Case "Error 2045"
                    Encode = "#SPILL!"
                Case "Error 2046"
                    Encode = "#CONNECT!"
                Case "Error 2047"
                    Encode = "#BLOCKED!"
                Case "Error 2048"
                    Encode = "#UNKNOWN!"
                Case "Error 2049"
                    Encode = "#FIELD!"
                Case "Error 2050"
                    Encode = "#CALC!"
                Case Else
                    Encode = CStr(x)        'should never hit this line...
            End Select
        Case Else
            Throw "Cannot convert variant of type " + TypeName(x) + " to String"
    End Select
    Exit Function
ErrHandler:
    Throw "#Encode: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Throw
' Purpose    : Simple error handling.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Throw(ByVal ErrorString As String)
    Err.Raise vbObjectError + 1, , ErrorString
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ThrowIfError
' Purpose   : In the event of an error, methods intended to be callable from spreadsheets
'             return an error string (starts with "#", ends with "!"). ThrowIfError allows such
'             methods to be used from VBA code while keeping error handling robust
'             MyVariable = ThrowIfError(MyFunctionThatReturnsAStringIfAnErrorHappens(...))
'---------------------------------------------------------------------------------------
Public Function ThrowIfError(Data As Variant)
    ThrowIfError = Data
    If VarType(Data) = vbString Then
        If Left$(Data, 1) = "#" Then
            If Right$(Data, 1) = "!" Then
                Throw CStr(Data)
            End If
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : NumDimensions
' Purpose   : Returns the number of dimensions in an array variable, or 0 if the variable
'             is not an array.
'---------------------------------------------------------------------------------------
Private Function NumDimensions(x As Variant) As Long
    Dim i As Long
    Dim y As Long
    If Not IsArray(x) Then
        NumDimensions = 0
        Exit Function
    Else
        On Error GoTo ExitPoint
        i = 1
        Do While True
            y = LBound(x, i)
            i = i + 1
        Loop
    End If
ExitPoint:
    NumDimensions = i - 1
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FunctionWizardActive
' Purpose    : Test if Excel's Function Wizard is active to allow early exit in slow functions.
'              https://stackoverflow.com/questions/20866484/can-i-disable-a-vba-udf-calculation-when-the-insert-function-function-arguments
' -----------------------------------------------------------------------------------------------------------------------
Private Function FunctionWizardActive() As Boolean
    
    On Error GoTo ErrHandler
    If Not Application.CommandBars("Standard").Controls(1).Enabled Then
        FunctionWizardActive = True
    End If

    Exit Function
ErrHandler:
    Throw "#FunctionWizardActive: " & Err.Description & "!"
End Function

