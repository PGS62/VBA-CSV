Attribute VB_Name = "modCSVReadWrite"
' VBA-CSV

' Copyright (C) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme
' This version at: https://github.com/PGS62/VBA-CSV/releases/tag/v0.2

Option Explicit

Private m_FSO As Scripting.FileSystemObject

#If VBA7 And Win64 Then
    'for 64-bit Excel
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As Long
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#Else
    'for 32-bit Excel
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

Private Enum enmErrorStyle
    es_ReturnString = 0
    es_RaiseError = 1
End Enum

Private Const m_ErrorStyle = es_ReturnString

Private Enum enmSourceType
    st_File = 0
    st_URL = 1
    st_String = 2
End Enum

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVRead
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path, or else a URL of a file, or else a string in
'             CSV format.
' ConvertTypes: Controls whether fields in the file are converted to typed values or remain as strings, and
'             sets the treatment of "quoted fields" and space characters.
'
'             ConvertTypes should be a string of zero or more letters from allowed characters "NDBETQ".
'
'             The most commonly useful letters are:
'             1) `N` number fields are returned as numbers (Doubles).
'             2) `D` date fields (that respect DateFormat) are returned as Dates.
'             3) `B` fields matching TrueStrings or FalseStrings are returned as Booleans.
'
'             ConvertTypes is optional and defaults to the null string for no type conversion. `TRUE is`
'             equivalent to `NDB` and `FALSE` to the null string.
'
'             Three further options are available:
'             4) `E` fields that match Excel errors are converted to error values. There are fourteen of
'             these, including `#N/A`, `#NAME?`, `#VALUE!` and `#DIV/0!`.
'             5) `T` leading and trailing spaces are trimmed from fields. In the case of quoted fields,
'             this will not remove spaces between the quotes.
'             6) `Q` conversion happens for both quoted and unquoted fields; otherwise only unquoted fields
'             are converted.
'
'             For most files, correct type conversion can be achieved with ConvertTypes as a string which
'             applies for all columns, but type conversion can also be specified on a per-column basis.
'
'             Enter an array (or range) with two columns or two rows, column numbers on the left/top and
'             type conversion (subset of `NDBETQ`) on the right/bottom. Instead of column numbers, you can
'             enter strings matching the contents of the header row, and a column number of zero applies to
'             all columns not otherwise referenced.
'
'             For convenience when calling from VBA, you can pass an array of two element arrays such as
'             `Array(Array(0,"N"),Array(3,""),Array("Phone",""))` to convert all numbers in a file into
'             numbers in the return except for those in column 3 and in the column(s) headed "Phone".
' Delimiter : By default, CSVRead will try to detect a file's delimiter as the first instance of comma, tab,
'             semi-colon, colon or pipe found outside quoted regions in the first 10,000 characters of the
'             file. If it can't auto-detect the delimiter, it will assume comma. If your file includes a
'             different character or string delimiter you should pass that as the Delimiter argument.
'
'             Alternatively, enter `FALSE` as the delimiter to treat the file as "not a delimited file". In
'             this case the return will mimic how the file would appear in a text editor such as NotePad.
'             The file will be split into lines at all line breaks (irrespective of double quotes) and each
'             element of the return will be a line of the file.
' IgnoreRepeated: Whether delimiters which appear at the start of a line, the end of a line or immediately
'             after another delimiter should be ignored while parsing; useful-for fixed-width files with
'             delimiter padding between fields.
' DateFormat: The format of dates in the file such as `Y-M-D`, `M-D-Y` or `Y/M/D`. Also supports `ISO` for
'             ISO8601 (e.g., 2021-08-26T09:11:30) or `ISOZ` (time zone given e.g.
'             2021-08-26T13:11:30+05:00), in which case dates-with-time are returned in UTC time.
' Comment   : Rows that start with this string will be skipped while parsing.
' IgnoreEmptyLines: Whether empty rows/lines in the file should be skipped while parsing (if `FALSE`, each
'             column will be assigned ShowMissingsAs for that empty row).
' HeaderRowNum: The row in the file containing headers. This argument is most useful when calling from VBA,
'             with SkipToRow set to one more than HeaderRowNum. In that case the function returns the "data
'             rows", and the header rows are returned via the by-reference argument HeaderRow. Optional and
'             defaults to 0.
' SkipToRow : The first row in the file that's included in the return. Optional and defaults to one more
'             than HeaderRowNum.
' SkipToCol : The column in the file at which reading starts. Optional and defaults to 1 to read from the
'             first column.
' NumRows   : The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the
'             end of the file are read.
' NumCols   : The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol
'             are read.
' TrueStrings: Indicates how `TRUE` values are represented in the file. May be a string, an array of strings
'             or a range containing strings; by default, `TRUE`, `True` and `true` are recognised.
' FalseStrings: Indicates how `FALSE` values are represented in the file. May be a string, an array of
'             strings or a range containing strings; by default, `FALSE`, `False` and `false` are
'             recognised.
' MissingStrings: Indicates how missing values are represented in the file. May be a string, an array of
'             strings or a range containing strings. By default, only an empty field (consecutive
'             delimiters) is considered missing.
' ShowMissingsAs: Fields which are missing in the file (consecutive delimiters) or match one of the
'             MissingStrings are returned in the array as ShowMissingsAs. Defaults to Empty, but the null
'             string or `#N/A!` error value can be good alternatives.
'
'             If NumRows is greater than the number of rows in the file then the return is "padded" with
'             the value of ShowMissingsAs. Likewise, if NumCols is greater than the number of columns in
'             the file.
' Encoding  : Allowed entries are `UTF-16`, `UTF-8`, `UTF-8-BOM`, and `ANSI`, but for most files this
'             argument can be omitted and CSVRead will detect the file's encoding. If auto-detection does
'             not work then it's possible that the file is encoded `UTF-16` but without a byte option mark,
'             so try entering Encoding as `UTF-16`.
' DecimalSeparator: In many places in the world, floating point number decimals are separated with a comma
'             instead of a period (3,14 vs. 3.14). CSVRead can correctly parse these numbers by passing in
'             the DecimalSeparator as a comma, in which case comma ceases to be a candidate if the parser
'             needs to guess the Delimiter.
' HeaderRow : This by-reference argument is for use from VBA (as opposed to from Excel formulas). It is
'             populated with the contents of the header row, with no type conversion except that quoted
'             fields are unquoted.
'
' Notes     : See also companion function CSVRead.
'
'             For discussion of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
'---------------------------------------------------------------------------------------------------------
Public Function CSVRead(FileName As String, Optional ConvertTypes As Variant = False, _
    Optional ByVal Delimiter As Variant, Optional IgnoreRepeated As Boolean, _
    Optional DateFormat As String, Optional Comment As String, _
    Optional IgnoreEmptyLines As Boolean = False, Optional ByVal HeaderRowNum As Long = 0, _
    Optional ByVal SkipToRow As Long = 1, Optional ByVal SkipToCol As Long = 1, _
    Optional ByVal NumRows As Long = 0, Optional ByVal NumCols As Long = 0, _
    Optional TrueStrings As Variant, Optional FalseStrings As Variant, _
    Optional MissingStrings As Variant, Optional ByVal ShowMissingsAs As Variant, _
    Optional ByVal Encoding As Variant, Optional DecimalSeparator As String = vbNullString, _
    Optional ByRef HeaderRow)
Attribute CSVRead.VB_Description = "Returns the contents of a comma-separated file on disk as an array."
Attribute CSVRead.VB_ProcData.VB_Invoke_Func = " \n14"

    Const DQ = """"
    Const Err_Delimiter = "Delimiter character must be passed as a string, FALSE for no delimiter. Omit to guess from file contents"
    Const Err_Delimiter2 = "Delimiter must have at least one character and cannot start with a double quote, line feed or carriage return"
    Const Err_FileEmpty = "File is empty"
    
    Const Err_FunctionWizard = "Disabled in Function Wizard"
    Const Err_NumCols = "NumCols must be positive to read a given number of columns, or zero or omitted to read all columns from SkipToCol to the maximum column encountered."
    Const Err_NumRows = "NumRows must be positive to read a given number of rows, or zero or omitted to read all rows from SkipToRow to the end of the file."
    Const Err_Seps1 = "DecimalSeparator must be a single character"
    Const Err_Seps2 = "DecimalSeparator must not be equal to the first character of Delimiter or to a line-feed or carriage-return"
    Const Err_SkipToCol = "SkipToCol must be at least 1."
    Const Err_SkipToRow = "SkipToRow must be at least 1."
    Const Err_Comment = "Comment must not contain double-quote, line feed or carriage return"
    Const Err_UTF8BOM = "Argument Encoding specifies that the file is UTF 8 encoded with Byte Option Mark, but the file has some other encoding"
    Const Err_HeaderRowNum = "HeaderRowNum must be greater than or equal to zero and less than or equal to SkipToRow"
    
    Dim AcceptWithoutTimeZone As Boolean
    Dim AcceptWithTimeZone As Boolean
    Dim AnyConversion As Boolean
    Dim AnySentinels As Boolean
    Dim ColByColFormatting As Boolean
    Dim ColIndexes() As Long
    Dim ConvertQuoted As Boolean
    Dim CSVContents As String
    Dim CTDict As New Scripting.Dictionary
    Dim DateOrder As Long
    Dim DateSeparator As String
    Dim ErrRet As String
    Dim FirstThreeChars As String
    Dim i As Long
    Dim ISO8601 As Boolean
    Dim IsUTF8BOM As Boolean
    Dim j As Long
    Dim k As Long
    Dim Lengths() As Long
    Dim m As Long
    Dim MaxSentinelLength As Long
    Dim NeedToFill As Boolean
    Dim NotDelimited As Boolean
    Dim NumColsFound As Long
    Dim NumColsInReturn As Long
    Dim NumFields As Long
    Dim NumRowsFound As Long
    Dim NumRowsInReturn As Long
    Dim QuoteCounts() As Long
    Dim Ragged As Boolean
    Dim ReturnArray() As Variant
    Dim RowIndexes() As Long
    Dim Sentinels As New Scripting.Dictionary
    Dim SepStandard As Boolean
    Dim SF As Scripting.File
    Dim ShowBooleansAsBooleans As Boolean
    Dim ShowDatesAsDates As Boolean
    Dim ShowErrorsAsErrors As Boolean
    Dim ShowMissingsAsEmpty As Boolean
    Dim ShowNumbersAsNumbers As Boolean
    Dim SourceType As enmSourceType
    Dim Starts() As Long
    Dim strDelimiter As String
    Dim STS As Scripting.TextStream
    Dim SysDateOrder As Long
    Dim SysDateSeparator As String
    Dim SysDecimalSeparator As String
    Dim TempFile As String
    Dim TrimFields As Boolean
    Dim TriState As Long
    
    On Error GoTo ErrHandler

    SourceType = InferSourceType(FileName)

    'Download file from internet to local temp folder
    If SourceType = st_URL Then
        TempFile = Environ("Temp") & "\VBA-CSV\Downloads\DownloadedFile.csv"
        FileName = Download(FileName, TempFile)
        SourceType = st_File
    End If

    'Parse and validate inputs...
    If SourceType <> st_String Then
        ParseEncoding FileName, Encoding, TriState, IsUTF8BOM
    End If

    If VarType(Delimiter) = vbBoolean Then
        If Not Delimiter Then
            NotDelimited = True
        Else
            Throw Err_Delimiter
        End If
    ElseIf VarType(Delimiter) = vbString Then
        If Len(Delimiter) = 0 Then
            strDelimiter = InferDelimiter(SourceType, FileName, TriState, DecimalSeparator)
        ElseIf Left$(Delimiter, 1) = DQ Or Left$(Delimiter, 1) = vbLf Or Left$(Delimiter, 1) = vbCr Then
            Throw Err_Delimiter2
        Else
            strDelimiter = Delimiter
        End If
    ElseIf IsEmpty(Delimiter) Or IsMissing(Delimiter) Then
        strDelimiter = InferDelimiter(SourceType, FileName, TriState, DecimalSeparator)
    Else
        Throw Err_Delimiter
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

    ParseConvertTypes ConvertTypes, ShowNumbersAsNumbers, _
        ShowDatesAsDates, ShowBooleansAsBooleans, ShowErrorsAsErrors, _
        ConvertQuoted, TrimFields, ColByColFormatting, HeaderRowNum, CTDict

    MakeSentinels Sentinels, MaxSentinelLength, AnySentinels, ShowBooleansAsBooleans, _
        ShowErrorsAsErrors, ShowMissingsAs, TrueStrings, FalseStrings, MissingStrings
    
    If ShowDatesAsDates Then
        ParseDateFormat DateFormat, DateOrder, DateSeparator, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone
        SysDateOrder = Application.International(xlDateOrder)
        SysDateSeparator = Application.International(xlDateSeparator)
    End If

    If HeaderRowNum < 0 Then Throw Err_HeaderRowNum
    If SkipToRow = 0 Then SkipToRow = HeaderRowNum + 1
    If HeaderRowNum > SkipToRow Then Throw Err_HeaderRowNum
    If SkipToCol = 0 Then SkipToCol = 1
    If SkipToRow < 1 Then Throw Err_SkipToRow
    If SkipToCol < 1 Then Throw Err_SkipToCol
    If NumRows < 0 Then Throw Err_NumRows
    If NumCols < 0 Then Throw Err_NumCols

    If HeaderRowNum > SkipToRow Then Throw Err_HeaderRowNum
       
    If InStr(Comment, DQ) > 0 Or InStr(Comment, vbLf) > 0 Or InStr(Comment, vbCrLf) > 0 Then Throw Err_Comment
    'End of input validation
          
    If SourceType = st_String Then
        CSVContents = FileName
        
        If NotDelimited Then
            CSVRead = SplitCSVContents(CSVContents, SkipToRow, NumRows)
            Exit Function
        End If
        
        Call ParseCSVContents(CSVContents, DQ, strDelimiter, Comment, IgnoreEmptyLines, IgnoreRepeated, SkipToRow, _
            HeaderRowNum, NumRows, NumRowsFound, NumColsFound, NumFields, Ragged, Starts, Lengths, RowIndexes, _
            ColIndexes, QuoteCounts, HeaderRow)
    Else
        If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
          
        Set SF = m_FSO.GetFile(FileName)
        If FunctionWizardActive() Then
            If SF.Size > 1000000 Then
                CSVRead = "#" & Err_FunctionWizard & "!"
                Exit Function
            End If
        End If
    
        Set STS = SF.OpenAsTextStream(ForReading, TriState)
    
        If STS.AtEndOfStream Then Throw Err_FileEmpty
    
        If NotDelimited Then
            CSVRead = ShowTextFile(STS, SkipToRow, NumRows)
            Exit Function
        End If

        If STS.AtEndOfStream Then
            STS.Close: Set STS = Nothing: Set SF = Nothing
            Throw Err_FileEmpty
        End If
        If IsUTF8BOM Then
            FirstThreeChars = STS.Read(3)
            If FirstThreeChars <> Chr(239) & Chr(187) & Chr(191) Then Throw Err_UTF8BOM
        End If
        If SkipToRow = 1 And NumRows = 0 Then
            CSVContents = STS.ReadAll
            STS.Close: Set STS = Nothing: Set SF = Nothing
            Call ParseCSVContents(CSVContents, DQ, strDelimiter, Comment, IgnoreEmptyLines, IgnoreRepeated, SkipToRow, _
                HeaderRowNum, NumRows, NumRowsFound, NumColsFound, NumFields, Ragged, Starts, Lengths, RowIndexes, _
                ColIndexes, QuoteCounts, HeaderRow)
        Else
            CSVContents = ParseCSVContents(STS, DQ, strDelimiter, Comment, IgnoreEmptyLines, IgnoreRepeated, SkipToRow, _
                HeaderRowNum, NumRows, NumRowsFound, NumColsFound, NumFields, Ragged, Starts, Lengths, RowIndexes, _
                ColIndexes, QuoteCounts, HeaderRow)
            STS.Close
        End If
    End If
           
    If NumCols = 0 Then
        NumColsInReturn = NumColsFound - SkipToCol + 1
        If NumColsInReturn <= 0 Then
            Throw "SkipToCol (" & CStr(SkipToCol) & ") exceeds the number of columns in the file (" & CStr(NumColsFound) & ")"
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
        ShowBooleansAsBooleans Or ShowErrorsAsErrors Or TrimFields
        
    ReDim ReturnArray(1 To NumRowsInReturn, 1 To NumColsInReturn)
        
    For k = 1 To NumFields
        i = RowIndexes(k)
        j = ColIndexes(k) - SkipToCol + 1
        If j >= 1 And j <= NumColsInReturn Then
            If ColByColFormatting Then
                ReturnArray(i, j) = Mid$(CSVContents, Starts(k), Lengths(k))
            Else
                ReturnArray(i, j) = ConvertField(Mid$(CSVContents, Starts(k), Lengths(k)), AnyConversion, _
                    Lengths(k), TrimFields, DQ, QuoteCounts(k), ConvertQuoted, _
                    ShowNumbersAsNumbers, SepStandard, DecimalSeparator, SysDecimalSeparator, _
                    ShowDatesAsDates, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, _
                    AnySentinels, Sentinels, MaxSentinelLength, ShowMissingsAs)
            End If
            
            'File has variable number of fields per line...
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

    If Ragged Then
        If Not IsEmpty(HeaderRow) Then
            If sNCols(HeaderRow) < sNCols(ReturnArray) + SkipToCol - 1 Then
                ReDim Preserve HeaderRow(1 To 1, 1 To sNCols(ReturnArray) + SkipToCol - 1)
            End If
        End If
    End If
    If SkipToCol > 1 Then
        If Not IsEmpty(HeaderRow) Then
            Dim HeaderRowTruncated() As String
            ReDim HeaderRowTruncated(1 To 1, 1 To NumColsInReturn)
            For i = 1 To NumColsInReturn
                HeaderRowTruncated(1, i) = HeaderRow(1, i + SkipToCol - 1)
            Next i
            HeaderRow = HeaderRowTruncated
        End If
    End If

    If ColByColFormatting Then
        Dim NR As Long, NC As Long, CT As Variant, UnQuotedHeader As String, NCH As Long, Field As String, QC As Long
        NR = sNRows(ReturnArray)
        NC = sNCols(ReturnArray)
        If IsEmpty(HeaderRow) Then
            NCH = 0
        Else
            NCH = sNCols(HeaderRow) 'possible that headers has fewer than expected columns if file is ragged
        End If

        For j = 1 To NC
            If j + SkipToCol - 1 <= NCH Then
                UnQuotedHeader = HeaderRow(1, j + SkipToCol - 1)
            Else
                UnQuotedHeader = -1 'Guaranteed not to be a key of the Dictionary
            End If
            If CTDict.Exists(j + SkipToCol - 1) Then
                CT = CTDict(j + SkipToCol - 1)
            ElseIf CTDict.Exists(UnQuotedHeader) Then
                CT = CTDict(UnQuotedHeader)
            ElseIf CTDict.Exists(0) Then
                CT = CTDict(0)
            Else
                CT = False
            End If
            ParseCTString CT, ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, ShowErrorsAsErrors, ConvertQuoted, TrimFields
            AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
                ShowBooleansAsBooleans Or ShowErrorsAsErrors
                
            Set Sentinels = New Scripting.Dictionary
            MakeSentinels Sentinels, MaxSentinelLength, AnySentinels, ShowBooleansAsBooleans, ShowErrorsAsErrors, ShowMissingsAs, TrueStrings, FalseStrings, MissingStrings

            For i = 1 To NR
                If Not IsEmpty(ReturnArray(i, j)) Then
                    Field = CStr(ReturnArray(i, j))
                    QC = CountQuotes(Field, DQ)
                    ReturnArray(i, j) = ConvertField(Field, AnyConversion, _
                        Len(ReturnArray(i, j)), TrimFields, DQ, QC, ConvertQuoted, _
                        ShowNumbersAsNumbers, SepStandard, DecimalSeparator, SysDecimalSeparator, _
                        ShowDatesAsDates, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, _
                        AnySentinels, Sentinels, MaxSentinelLength, ShowMissingsAs)
                End If
            Next i
        Next j
    End If

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
    ErrRet = "#CSVRead: " & Err.Description & "!"
    If Not STS Is Nothing Then STS.Close
    If m_ErrorStyle = es_ReturnString Then
        CSVRead = ErrRet
    Else
        Throw ErrRet
    End If
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead
' Purpose    : Register the function CSVRead with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
'---------------------------------------------------------------------------------------------------------
Sub RegisterCSVRead()
    Const Description = "Returns the contents of a comma-separated file on disk as an array."
    Dim ArgumentDescriptions() As String

    On Error GoTo ErrHandler

    ReDim ArgumentDescriptions(1 To 19)
    ArgumentDescriptions(1) = "The full name of the file, including the path, or else a URL of a file, or else a string in CSV format."
    ArgumentDescriptions(2) = "Type conversion: Boolean or string, subset of letters NDBETQ. N = convert Numbers, D = convert Dates, B = Convert Booleans, E = convert Excel errors, T = trim leading & trailing spaces, Q = quoted fields also converted. TRUE = NDB, FALSE = no conversion."
    ArgumentDescriptions(3) = "Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe found outside quoted regions within the first 10,000 characters. Enter FALSE to  see the file's contents as would be displayed in a text editor."
    ArgumentDescriptions(4) = "Whether delimiters which appear at the start of a line, the end of a line or immediately after another delimiter should be ignored while parsing; useful-for fixed-width files with delimiter padding between fields."
    ArgumentDescriptions(5) = "The format of dates in the file such as `Y-M-D`, `M-D-Y` or `Y/M/D`. Also supports `ISO` for ISO8601 (e.g., 2021-08-26T09:11:30) or `ISOZ` (time zone given e.g. 2021-08-26T13:11:30+05:00), in which case dates-with-time are returned in UTC time."
    ArgumentDescriptions(6) = "Rows that start with this string will be skipped while parsing."
    ArgumentDescriptions(7) = "Whether empty rows/lines in the file should be skipped while parsing (if `FALSE`, each column will be assigned ShowMissingsAs for that empty row)."
    ArgumentDescriptions(8) = "The row in the file containing headers. Optional and defaults to 0."
    ArgumentDescriptions(9) = "The first row in the file that's included in the return. Optional and defaults to one more than HeaderRowNum."
    ArgumentDescriptions(10) = "The column in the file at which reading starts. Optional and defaults to 1 to read from the first column."
    ArgumentDescriptions(11) = "The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the end of the file are read."
    ArgumentDescriptions(12) = "The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol are read."
    ArgumentDescriptions(13) = "Indicates how `TRUE` values are represented in the file. May be a string, an array of strings or a range containing strings; by default, `TRUE`, `True` and `true` are recognised."
    ArgumentDescriptions(14) = "Indicates how `FALSE` values are represented in the file. May be a string, an array of strings or a range containing strings; by default, `FALSE`, `False` and `false` are recognised."
    ArgumentDescriptions(15) = "Indicates how missing values are represented in the file. May be a string, an array of strings or a range containing strings. By default, only an empty field (consecutive delimiters) is considered missing."
    ArgumentDescriptions(16) = "Fields which are missing in the file (consecutive delimiters) or match one of the MissingStrings are returned in the array as ShowMissingsAs. Defaults to Empty, but the null string or `#N/A!` error value can be good alternatives."
    ArgumentDescriptions(17) = "Allowed entries are `UTF-16`, `UTF-8`, `UTF-8-BOM`, and `ANSI`, but for most files this argument can be omitted and CSVRead will detect the file's encoding."
    ArgumentDescriptions(18) = "The character that represents a decimal point. If omitted, then the value from Windows regional settings is used."
    ArgumentDescriptions(19) = "For use from VBA only."
    Application.MacroOptions "CSVRead", Description, , , , , , , , , ArgumentDescriptions
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function CSVRead failed with error: " + Err.Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferSourceType
' Purpose    : Guess whether FileName is in fact a file, a URL or a string in CSV format
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferSourceType(FileName As String) As enmSourceType

    On Error GoTo ErrHandler
    If Mid$(FileName, 2, 2) = ":\" Then
        InferSourceType = st_File
    ElseIf Left$(FileName, 2) = "\\" Then
        InferSourceType = st_File
    ElseIf Left$(FileName, 8) = "https://" Then
        InferSourceType = st_URL
    ElseIf Left$(FileName, 7) = "http://" Then
        InferSourceType = st_URL
    ElseIf InStr(FileName, vbLf) > 0 Then
        InferSourceType = st_String
    ElseIf InStr(FileName, vbCr) > 0 Then
        InferSourceType = st_String
    Else
        InferSourceType = st_File
    End If

    Exit Function
ErrHandler:
    Throw "#InferSourceType: " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Download
' Purpose   : Downloads bits from the Internet and saves them to a file.
'             See https://msdn.microsoft.com/en-us/library/ms775123(v=vs.85).aspx
'---------------------------------------------------------------------------------------
Private Function Download(URLAddress As String, ByVal FileName As String)
    Dim ErrString As String
    Dim Res
    Dim TargetFolder As String

    On Error GoTo ErrHandler
    
    TargetFolder = FileFromPath(FileName, False)
    CreatePath TargetFolder
    If FileExists(FileName) Then FileDelete FileName
    Res = URLDownloadToFile(0, URLAddress, FileName, 0, 0)
    If Res <> 0 Then
        ErrString = ParseDownloadError(CLng(Res))
        Throw "Windows API function URLDownloadToFile returned error code " & CStr(Res) & " with description '" & ErrString & "'"
    End If
    If Not FileExists(FileName) Then Throw "Windows API function URLDownloadToFile did not report an error, but appears to have not successfuly downloaded a file from " & URLAddress & " to " & FileName
    Download = FileName

    Exit Function
ErrHandler:
    Throw "#Download: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDownloadError, sub of Download
'              https://www.vbforums.com/showthread.php?882757-URLDownloadToFile-error-codes
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseDownloadError(ErrNum As Long)
    Dim ErrString
    Select Case ErrNum
        Case &H80004004
            ErrString = "Aborted"
        Case &H800C0001
            ErrString = "Destination File Exists"
        Case &H800C0002
            ErrString = "Invalid Url"
        Case &H800C0003
            ErrString = "No Session"
        Case &H800C0004
            ErrString = "Cannot Connect"
        Case &H800C0005
            ErrString = "Resource Not Found"
        Case &H800C0006
            ErrString = "Object Not Found"
        Case &H800C0007
            ErrString = "Data Not Available"
        Case &H800C0008
            ErrString = "Download Failure"
        Case &H800C0009
            ErrString = "Authentication Required"
        Case &H800C000A
            ErrString = "No Valid Media"
        Case &H800C000B
            ErrString = "Connection Timeout"
        Case &H800C000C
            ErrString = "Invalid Request"
        Case &H800C000D
            ErrString = "Unknown Protocol"
        Case &H800C000E
            ErrString = "Security Problem"
        Case &H800C000F
            ErrString = "Cannot Load Data"
        Case &H800C0010
            ErrString = "Cannot Instantiate Object"
        Case &H800C0014
            ErrString = "Redirect Failed"
        Case &H800C0015
            ErrString = "Redirect To Dir"
        Case &H800C0016
            ErrString = "Cannot Lock Request"
        Case Else
            ErrString = "Unknown"
    End Select
    ParseDownloadError = ErrString
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileExists
' Purpose    : Returns True if FileName exists on disk, False o.w.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileExists(FileName As String) As Boolean
    Dim F As Scripting.File
    On Error GoTo ErrHandler
    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    Set F = m_FSO.GetFile(FileName)
    FileExists = True
    Exit Function
ErrHandler:
    FileExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileDelete
' Purpose    : Delete a file, returns True or error string.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileDelete(FileName As String) As Boolean
    Dim F As Scripting.File
    On Error GoTo ErrHandler

    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    Set F = m_FSO.GetFile(FileName)
    F.Delete
    FileDelete = True

    Exit Function
ErrHandler:
    Throw "#FileDelete: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileFromPath
' Purpose    : Split file-with-path to file name (if ReturnFileName is True) or path otherwise.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileFromPath(FullFileName As String, Optional ReturnFileName As Boolean = True) As Variant
    Dim SlashPos As Long
    Dim SlashPos2 As Long

    On Error GoTo ErrHandler

    SlashPos = InStrRev(FullFileName, "\")
    SlashPos2 = InStrRev(FullFileName, "/")
    If SlashPos2 > SlashPos Then SlashPos = SlashPos2
    If SlashPos = 0 Then Throw "Neither '\' nor '/' found"

    If ReturnFileName Then
        FileFromPath = Mid$(FullFileName, SlashPos + 1)
    Else
        FileFromPath = Left$(FullFileName, SlashPos - 1)
    End If

    Exit Function
ErrHandler:
    Throw "#FileFromPath: " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : FolderExists
' Purpose   : Returns True or False. Does not matter if FolderPath has a terminating backslash or not.
'---------------------------------------------------------------------------------------
Private Function FolderExists(ByVal FolderPath As String)
    Dim F As Scripting.Folder
    
    On Error GoTo ErrHandler
    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    
    Set F = m_FSO.GetFolder(FolderPath)
    FolderExists = True
    Exit Function
ErrHandler:
    FolderExists = False
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : CreatePath
' Purpose   : Creates a folder on disk. FolderPath can be passed in as C:\This\That\TheOther even if the
'             folder C:\This does not yet exist. If successful returns the name of the
'             folder. If not successful returns an error string.
' Arguments
' FolderPath: Path of the folder to be created. For example C:\temp\My_New_Folder. It does not matter if
'             this path has a terminating backslash or not.
'---------------------------------------------------------------------------------------------------------
Private Function CreatePath(ByVal FolderPath As String)

    Dim F As Scripting.Folder
    Dim i As Long
    Dim isUNC As Boolean
    Dim ParentFolderName
    Dim ThisFolderName As String

    On Error GoTo ErrHandler

    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject

    If Left$(FolderPath, 2) = "\\" Then
        isUNC = True
    ElseIf Mid$(FolderPath, 2, 2) <> ":\" Or Asc(UCase$(Left$(FolderPath, 1))) < 65 Or Asc(UCase$(Left$(FolderPath, 1))) > 90 Then
        Throw "First three characters of FolderPath must give drive letter followed by "":\"" or else be""\\"" for " & _
            "UNC folder name"
    End If

    FolderPath = Replace(FolderPath, "/", "\")

    If Right$(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If

    If FolderExists(FolderPath) Then
        GoTo EarlyExit
    End If

    'Go back until we find parent folder that does exist
    For i = Len(FolderPath) - 1 To 3 Step -1
        If Mid$(FolderPath, i, 1) = "\" Then
            If FolderExists(Left$(FolderPath, i)) Then
                Set F = m_FSO.GetFolder(Left$(FolderPath, i))
                ParentFolderName = Left$(FolderPath, i)
                Exit For
            End If
        End If
    Next i

    If F Is Nothing Then Throw "Cannot create folder " & Left$(FolderPath, 3)

    'now add folders one level at a time
    For i = Len(ParentFolderName) + 1 To Len(FolderPath)
        If Mid$(FolderPath, i, 1) = "\" Then
            
            ThisFolderName = Mid$(FolderPath, InStrRev(FolderPath, "\", i - 1) + 1, i - 1 - InStrRev(FolderPath, "\", i - 1))
            F.SubFolders.Add ThisFolderName
            Set F = m_FSO.GetFolder(Left$(FolderPath, i))
        End If
    Next i

EarlyExit:
    Set F = m_FSO.GetFolder(FolderPath)
    CreatePath = F.path
    Set F = Nothing

    Exit Function
ErrHandler:
    Throw "#CreatePath: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseEncoding
' Purpose    : Set Booleans Encoding and IsUTF8BOM by parsing user input or calling DetectEncoding.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseEncoding(FileName As String, Encoding As Variant, ByRef TriState As Long, ByRef IsUTF8BOM As Boolean)

    Const Err_Encoding = "Encoding argument can usually be omitted, but otherwise Encoding be either ""UTF-16"", ""UTF-8"", ""UTF-8-BOM"" or ""ANSI""."
    
    On Error GoTo ErrHandler
    If IsEmpty(Encoding) Or IsMissing(Encoding) Then
        DetectEncoding FileName, TriState, IsUTF8BOM
    ElseIf VarType(Encoding) = vbString Then
        Select Case UCase(Replace(Replace(Encoding, "-", ""), " ", ""))
            Case "ANSI", "UTF8"
                TriState = TristateFalse
                IsUTF8BOM = False
            Case "UTF16"
                TriState = TristateTrue
                IsUTF8BOM = False
            Case "UTF8BOM"
                TriState = TristateFalse
                IsUTF8BOM = True
            Case Else
                Throw Err_Encoding
        End Select
    Else
        Throw Err_Encoding
    End If

    Exit Sub
ErrHandler:
    Throw "#ParseEncoding: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsCTValid
' Purpose    : Is a "CT string" (which can in fact be either a string or a Boolean) valid?
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsCTValid(CT As Variant) As Boolean

    Static rx As VBScript_RegExp_55.RegExp

    On Error GoTo ErrHandler
    If rx Is Nothing Then
        Set rx = New RegExp
        With rx
            .IgnoreCase = True
            .Pattern = "^[NDBETQ]*$"
            .Global = False        'Find first match only
        End With
    End If

    If VarType(CT) = vbBoolean Then
        IsCTValid = True
    ElseIf VarType(CT) = vbString Then
        IsCTValid = rx.Test(CT)
    Else
        IsCTValid = False
    End If

    Exit Function
ErrHandler:
    Throw "#IsCTValid: " & Err.Description & "!"
End Function

Private Function TestParseConvertTypes()
    Dim ShowNumbersAsNumbers As Boolean
    Dim ShowDatesAsDates As Boolean
    Dim ShowBooleansAsBooleans As Boolean
    Dim ShowErrorsAsErrors As Boolean
    Dim ConvertQuoted As Boolean
    Dim TrimFields As Boolean
    Dim ColByColFormatting As Boolean
    Dim CTDict As New Scripting.Dictionary
    Dim ConvertTypes As Variant

    ConvertTypes = Array(Array(1, "B"), Array(2, "N"))

    On Error GoTo ErrHandler
    ParseConvertTypes ConvertTypes, ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, ShowErrorsAsErrors, ConvertQuoted, TrimFields, ColByColFormatting, 1, CTDict

    TestParseConvertTypes = True
    Exit Function
ErrHandler:
    TestParseConvertTypes = "#TestParseConvertTypes: " & Err.Description & "!"
    Debug.Print TestParseConvertTypes
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CTsEqual
' Purpose    : Test if two CT strings (strings to define type conversion) are equal, i.e. will have the same effect
' -----------------------------------------------------------------------------------------------------------------------
Private Function CTsEqual(CT1, CT2) As Boolean
    On Error GoTo ErrHandler
    If VarType(CT1) = VarType(CT2) Then
        If CT1 = CT2 Then
            CTsEqual = True
            Exit Function
        End If
    End If
    CTsEqual = StandardiseCT(CT1) = StandardiseCT(CT2)
    Exit Function
ErrHandler:
    Throw "#CTsEqual: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StandardiseCT
' Purpose    : Put a CT string into standard form so that two such can be compared.
' -----------------------------------------------------------------------------------------------------------------------
Private Function StandardiseCT(CT As Variant) As String
    On Error GoTo ErrHandler
    If VarType(CT) = vbBoolean Then
        If CT Then
            StandardiseCT = "BDN"
        Else
            StandardiseCT = ""
        End If
        Exit Function
    ElseIf VarType(CT) = vbString Then
        StandardiseCT = IIf(InStr(1, CT, "B", vbTextCompare), "B", "") & _
            IIf(InStr(1, CT, "D", vbTextCompare), "D", "") & _
            IIf(InStr(1, CT, "E", vbTextCompare), "E", "") & _
            IIf(InStr(1, CT, "N", vbTextCompare), "N", "") & _
            IIf(InStr(1, CT, "Q", vbTextCompare), "Q", "") & _
            IIf(InStr(1, CT, "T", vbTextCompare), "T", "")
    End If

    Exit Function
ErrHandler:
    Throw "#StandardiseCT: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OneDArrayToTwoDArray
' Purpose    : Convert 1-d array of 2-element 1-d arrays into a 1-based, two-column, 2-d array
' -----------------------------------------------------------------------------------------------------------------------
Private Function OneDArrayToTwoDArray(x As Variant)
    Const Err_1DArray = "If ConvertTypes is given as a 1-dimensional array, each element must be a 1-dimensional array with two elements"

    Dim TwoDArray()
    Dim i As Long, k As Long
    On Error GoTo ErrHandler
    ReDim TwoDArray(1 To UBound(x) - LBound(x) + 1, 1 To 2)
    For i = LBound(x) To UBound(x)
        k = k + 1
        If Not IsArray(x(i)) Then Throw Err_1DArray
        If NumDimensions(x(i)) <> 1 Then Throw Err_1DArray
        If UBound(x(i)) - LBound(x(i)) <> 1 Then Throw Err_1DArray
        TwoDArray(k, 1) = x(i)(LBound(x(i)))
        TwoDArray(k, 2) = x(i)(1 + LBound(x(i)))
    Next i
    OneDArrayToTwoDArray = TwoDArray
    Exit Function
ErrHandler:
    Throw "#OneDArrayToTwoDArray: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseConvertTypes
' Purpose    : There is flexibility in how the ConvertTypes argument is provided to CSVRead:
'              a) As a string or Boolean for the same type conversion rules for every field in the file; or
'              b) An array to define different type conversion rules by column, in which case ConvertTypes can be passed
'                 as a two-column or two-row array (convenient from Excel) or as an array of two-element arrays
'                 (convenient from VBA).
'              If an array, then the left col(or top row or first element) can contain either column numbers or strings
'              that match the elements of the SkipToRow row of the file
'
' Parameters :
'  ConvertTypes          :
'  ShowNumbersAsNumbers  : Set only if ConvertTypes is not an array
'  ShowDatesAsDates      : Set only if ConvertTypes is not an array
'  ShowBooleansAsBooleans: Set only if ConvertTypes is not an array
'  ShowErrorsAsErrors    : Set only if ConvertTypes is not an array
'  ConvertQuoted         : Set only if ConvertTypes is not an array
'  TrimFields            : Set only if ConvertTypes is not an array
'  ColByColFormatting    : Set to True if ConvertTypes is an array
'  CTDict                : Set to a dictionary keyed on the elements of the left column (top row) of ConvertTypes,
'                          each element containing the corresponding right (or bottom) element.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseConvertTypes(ByVal ConvertTypes As Variant, ByRef ShowNumbersAsNumbers As Boolean, _
    ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
    ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef TrimFields As Boolean, _
    ByRef ColByColFormatting As Boolean, HeaderRowNum As Long, ByRef CTDict As Scripting.Dictionary)
    
    Const Err_2D = "If ConvertTypes is given as a two dimensional array then the lower bounds in each dimension must be 1"
    Const Err_Ambiguous = "ConvertTypes is ambiguous, it can be interpreted as two rows, or as two columns"
    Const Err_BadColumnIdentifier = "Column identifiers in the left column (or top row) of ConvertTypes must be strings or non-negative whole numbers"
    Const Err_BadCT = "Type Conversion given in bottom row (or right column) of ConvertTypes must be Booleans or strings containing letters NDBETQ"
    Const Err_ConvertTypes = "ConvertTypes must be a Boolean, a string with allowed letters ""NDBETQ"" or an array"
    Const Err_HeaderRowNum = "ConvertTypes specifies columns by their header (instead of by number), but HeaderRowNum has not been specified"
    
    Dim ColIdentifier As Variant
    Dim CT As Variant
    Dim i As Long
    Dim LCN As Long 'Left column number
    Dim NC As Long 'Number of columns
    Dim ND As Long 'Number of dimensions
    Dim NR As Long 'Number of rows
    Dim RCN As Long 'Right Column Number
    Dim Transposed As Boolean
    
    On Error GoTo ErrHandler
    If VarType(ConvertTypes) = vbString Or VarType(ConvertTypes) = vbBoolean Or IsEmpty(ConvertTypes) Then
        ParseCTString CStr(ConvertTypes), ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, ShowErrorsAsErrors, ConvertQuoted, TrimFields
        ColByColFormatting = False
        Exit Sub
    End If

    If TypeName(ConvertTypes) = "Range" Then ConvertTypes = ConvertTypes.Value2
    ND = NumDimensions(ConvertTypes)
    If ND = 1 Then
        ConvertTypes = OneDArrayToTwoDArray(ConvertTypes)
    ElseIf ND = 2 Then
        If LBound(ConvertTypes, 1) <> 1 Or LBound(ConvertTypes, 2) <> 1 Then
            Throw Err_2D
        End If
    End If

    NR = sNRows(ConvertTypes)
    NC = sNCols(ConvertTypes)
    If NR = 2 And NC = 2 Then
        'Tricky - have we been given two rows or two columns?
        If Not IsCTValid(ConvertTypes(2, 2)) Then Throw Err_ConvertTypes
        If IsCTValid(ConvertTypes(1, 2)) And IsCTValid(ConvertTypes(2, 1)) Then
            If StandardiseCT(ConvertTypes(1, 2)) <> StandardiseCT(ConvertTypes(2, 1)) Then
                Throw Err_Ambiguous
            End If
        End If
        If IsCTValid(ConvertTypes(2, 1)) Then
            ConvertTypes = Transpose(ConvertTypes)
            Transposed = True
        End If
    ElseIf NR = 2 Then
        ConvertTypes = Transpose(ConvertTypes)
        Transposed = True
        NR = NC
    ElseIf NC <> 2 Then
        Throw Err_ConvertTypes
    End If
    LCN = LBound(ConvertTypes, 2)
    RCN = LCN + 1
    For i = LBound(ConvertTypes, 1) To UBound(ConvertTypes, 1)
        ColIdentifier = ConvertTypes(i, LCN)
        CT = ConvertTypes(i, RCN)
        If IsNumber(ColIdentifier) Then
            If ColIdentifier <> CLng(ColIdentifier) Then
                Throw Err_BadColumnIdentifier & " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & ") is " & CStr(ColIdentifier)
            ElseIf ColIdentifier < 0 Then
                Throw Err_BadColumnIdentifier & " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & ") is " & CStr(ColIdentifier)
            End If
        ElseIf VarType(ColIdentifier) <> vbString Then
            Throw Err_BadColumnIdentifier & " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & ") is of type " & TypeName(ColIdentifier)
        End If
        If Not IsCTValid(CT) Then
            If VarType(CT) = vbString Then
                Throw Err_BadCT & " but ConvertTypes(" & IIf(Transposed, "2," & CStr(i), CStr(i) & ",2") & ") is string """ & CStr(CT) & """"
            Else
                Throw Err_BadCT & " but ConvertTypes(" & IIf(Transposed, "2," & CStr(i), CStr(i) & ",2") & ") is of type " & TypeName(CT)
            End If
        End If

        If CTDict.Exists(ColIdentifier) Then
            If Not CTsEqual(CTDict(ColIdentifier), CT) Then
                Throw "ConvertTypes is contradictory. Column " & CStr(ColIdentifier) & " is specified to be converted using two different conversion rules: " & CStr(CT) & " and " & CStr(CTDict(ColIdentifier))
            End If
        Else
            CT = StandardiseCT(CT)
            'Need this line to ensure that we parse the DateFormat provided when doing Col-by-col type conversion
            If InStr(CT, "D") > 0 Then ShowDatesAsDates = True
            If VarType(ColIdentifier) = vbString Then
                If HeaderRowNum = 0 Then
                    Throw Err_HeaderRowNum
                End If
            End If
            CTDict.Add ColIdentifier, CT
        End If
    Next i
    ColByColFormatting = True
    Exit Sub
ErrHandler:
    Throw "#ParseConvertTypes: " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsNumber
' Purpose   : Is a singleton a number?
'---------------------------------------------------------------------------------------
Private Function IsNumber(x As Variant) As Boolean
    Select Case VarType(x)
        Case vbDouble, vbInteger, vbSingle, vbLong ', vbCurrency, vbDecimal
            IsNumber = True
    End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCTString
' Purpose    : Parse the input ConvertTypes to set seven Boolean flags which are passed by reference.
' Parameters :
'  ConvertTypes          : The argument to CSVRead
'  ShowNumbersAsNumbers  : Should fields in the file that look like numbers be returned as Numbers? (Doubles)
'  ShowDatesAsDates      : Should fields in the file that look like dates with the specified DateFormat be returned as Dates?
'  ShowBooleansAsBooleans: Should fields in the file that match one of the TrueStrings or FalseStrings be returned as Booleans?
'  ShowErrorsAsErrors    : Should fields in the file that look like Excel errors (#N/A #REF! etc) be returned as errors?
'  ConvertQuoted         : Should the four conversion rules above apply even to quoted fields?
'  TrimFields            : Should leading and trailing spaces be trimmed from fields?
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseCTString(ByVal ConvertTypes As String, ByRef ShowNumbersAsNumbers As Boolean, _
    ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
    ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef TrimFields As Boolean)

    Const Err_ConvertTypes = "ConvertTypes must be Boolean or string with allowed letters NDBETQ. " & _
        """N"" show numbers as numbers, ""D"" show dates as dates, ""B"" show Booleans " & _
        "as Booleans, ""E"" show Excel errors as errors, ""T"" to trim leading and trailing " & _
        "spaces from fields, ""Q"" rules NDBE apply even to quoted fields, TRUE = ""NDB"" " & _
        "(convert unquoted numbers, dates and Booleans), FALSE = no conversion"
    Const Err_Quoted = "ConvertTypes is incorrect, ""Q"" indicates that conversion should apply even to " & _
        "quoted fields, but none of ""N"", ""D"", ""B"" or ""E"" are present to indicate which type conversion to apply"
    Dim i As Long

    On Error GoTo ErrHandler

    If ConvertTypes = "True" Or ConvertTypes = "False" Then
        ConvertTypes = StandardiseCT(CBool(ConvertTypes))
    End If

    ShowNumbersAsNumbers = False
    ShowDatesAsDates = False
    ShowBooleansAsBooleans = False
    ShowErrorsAsErrors = False
    ConvertQuoted = False
    For i = 1 To Len(ConvertTypes)
        'Adding another letter? Also change method IsCTValid!
        Select Case UCase(Mid$(ConvertTypes, i, 1))
            Case "N"
                ShowNumbersAsNumbers = True
            Case "D"
                ShowDatesAsDates = True
            Case "B"
                ShowBooleansAsBooleans = True
            Case "E"
                ShowErrorsAsErrors = True
            Case "Q"
                ConvertQuoted = True
            Case "T"
                TrimFields = True
            Case Else
                Throw Err_ConvertTypes & " Found unrecognised character '" _
                    & Mid$(ConvertTypes, i, 1) & "'"
        End Select
    Next i
    
    If ConvertQuoted And Not (ShowNumbersAsNumbers Or ShowDatesAsDates Or _
        ShowBooleansAsBooleans Or ShowErrorsAsErrors) Then
        Throw Err_Quoted
    End If

    Exit Sub
ErrHandler:
    Throw "#ParseCTString: " & Err.Description & "!"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : sNCols
' Purpose   : Number of columns in an array. Missing has zero rows, 1-dimensional arrays
'             have one row and the number of columns returned by this function.
'---------------------------------------------------------------------------------------
Private Function sNCols(Optional TheArray) As Long
    If TypeName(TheArray) = "Range" Then
        sNCols = TheArray.Columns.Count
    ElseIf IsMissing(TheArray) Then
        sNCols = 0
    ElseIf VarType(TheArray) < vbArray Then
        sNCols = 1
    Else
        Select Case NumDimensions(TheArray)
            Case 1
                sNCols = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
            Case Else
                sNCols = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
        End Select
    End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : sNRows
' Purpose   : Number of rows in an array. Missing has zero rows, 1-dimensional arrays have one row.
'---------------------------------------------------------------------------------------
Private Function sNRows(Optional TheArray) As Long
    If TypeName(TheArray) = "Range" Then
        sNRows = TheArray.Rows.Count
    ElseIf IsMissing(TheArray) Then
        sNRows = 0
    ElseIf VarType(TheArray) < vbArray Then
        sNRows = 1
    Else
        Select Case NumDimensions(TheArray)
            Case 1
                sNRows = 1
            Case Else
                sNRows = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
        End Select
    End If
End Function
'---------------------------------------------------------------------------------------------------------
' Procedure : Transpose
' Purpose   : Returns the transpose of an array.
' Arguments
' TheArray  : An array of arbitrary values.
'             also converts 0-based to 1-based arrays
'---------------------------------------------------------------------------------------------------------
Private Function Transpose(ByVal TheArray As Variant)
    Dim Co As Long
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim N As Long
    Dim Result As Variant
    Dim Ro As Long
    On Error GoTo ErrHandler
    Force2DArrayR TheArray, N, m
    Ro = LBound(TheArray, 1) - 1
    Co = LBound(TheArray, 2) - 1
    ReDim Result(1 To m, 1 To N)
    For i = 1 To N
        For j = 1 To m
            Result(j, i) = TheArray(i + Ro, j + Co)
        Next j
    Next i
    Transpose = Result
    Exit Function
ErrHandler:
    Transpose = "#Transpose: " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Force2DArrayR
' Purpose   : When writing functions to be called from sheets, we often don't want to process
'             the inputs as Range objects, but instead as Arrays. This method converts the
'             input into a 2-dimensional 1-based array (even if it's a single cell or single row of cells)
'---------------------------------------------------------------------------------------
Private Sub Force2DArrayR(ByRef RangeOrArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
    If TypeName(RangeOrArray) = "Range" Then RangeOrArray = RangeOrArray.Value2
    Force2DArray RangeOrArray, NR, NC
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Force2DArray
' Purpose   : In-place amendment of singletons and one-dimensional arrays to two dimensions.
'             singletons and 1-d arrays are returned as 2-d 1-based arrays. Leaves two
'             two dimensional arrays untouched (i.e. a zero-based 2-d array will be left as zero-based).
'             See also Force2DArrayR that also handles Range objects.
'---------------------------------------------------------------------------------------
Private Sub Force2DArray(ByRef TheArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
    Dim TwoDArray As Variant

    On Error GoTo ErrHandler

    Select Case NumDimensions(TheArray)
        Case 0
            ReDim TwoDArray(1 To 1, 1 To 1)
            TwoDArray(1, 1) = TheArray
            TheArray = TwoDArray
            NR = 1: NC = 1
        Case 1
            Dim i As Long
            Dim LB As Long
            LB = LBound(TheArray, 1)
            NR = 1: NC = UBound(TheArray, 1) - LB + 1
            ReDim TwoDArray(1 To 1, 1 To NC)
            For i = 1 To UBound(TheArray, 1) - LBound(TheArray) + 1
                TwoDArray(1, i) = TheArray(LB + i - 1)
            Next i
            TheArray = TwoDArray
        Case 2
            NR = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
            NC = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
            'Nothing to do
        Case Else
            Throw "Cannot convert array of dimension greater than two"
    End Select

    Exit Sub
ErrHandler:
    Throw "#Force2DArray: " & Err.Description & "!"
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
' Procedure  : DetectEncoding
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
Private Sub DetectEncoding(FilePath As String, ByRef TriState As Long, ByRef IsUTF8BOM As Boolean)

    Dim intAsc1Chr As Long
    Dim intAsc2Chr As Long
    Dim intAsc3Chr As Long
    Dim T As Scripting.TextStream

    On Error GoTo ErrHandler
    
    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    
    If (m_FSO.FileExists(FilePath) = False) Then
        Throw "File not found!"
    End If

    ' 1=Read-only, False=do not create if not exist, -1=Unicode 0=ASCII
    Set T = m_FSO.OpenTextFile(FilePath, 1, False, 0)
    If T.AtEndOfStream Then
        T.Close: Set T = Nothing
        TriState = TristateFalse
        IsUTF8BOM = False
        Exit Sub
    End If
    intAsc1Chr = Asc(T.Read(1))
    If T.AtEndOfStream Then
        T.Close: Set T = Nothing
        TriState = TristateFalse
        IsUTF8BOM = False
        Exit Sub
    End If
    intAsc2Chr = Asc(T.Read(1))
    
    If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
        'File is probably encoded UTF-16 LE BOM (little endian, with Byte Option Marker)
        TriState = TristateTrue
        IsUTF8BOM = False
    ElseIf (intAsc1Chr = 254) And (intAsc2Chr = 255) Then
        'File is probably encoded UTF-16 BE BOM (big endian, with Byte Option Marker)
        TriState = TristateTrue
        IsUTF8BOM = False
    Else
        If T.AtEndOfStream Then
            TriState = TristateFalse
            IsUTF8BOM = False
            Exit Sub
        End If
        intAsc3Chr = Asc(T.Read(1))
        If (intAsc1Chr = 239) And (intAsc2Chr = 187) And (intAsc3Chr = 191) Then
            'File is probably encoded UTF-8 with BOM
            TriState = TristateFalse
            IsUTF8BOM = True
        Else
            'File is probably encoded UTF-8 without BOM
            TriState = TristateFalse
            IsUTF8BOM = False
        End If
    End If

    T.Close: Set T = Nothing
    Exit Sub
ErrHandler:
    Throw "#DetectEncoding: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferDelimiter
' Purpose    : Infer the delimiter in a file by looking for first occurrence outside quoted regions of comma, tab,
'              semi-colon, colon or pipe (|). Only look in the first 10,000 characters, Would prefer to look at the first
'              10 lines, but that presents a problem for files with Mac line endings as T.ReadLine doesn't work for them...
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferDelimiter(st As enmSourceType, FileNameOrContents As String, TriState As Long, DecimalSeparator As String)
    
    Const CHUNK_SIZE = 1000
    Const Err_SourceType = "Cannot infer delimiter directly from URL"
    Const MAX_CHUNKS = 10
    Const QuoteChar As String = """"
    Dim Contents As String
    Dim CopyOfErr As String
    Dim EvenQuotes As Boolean
    Dim F As Scripting.File
    Dim i As Long, j As Long
    Dim MaxChars
    Dim T As TextStream

    On Error GoTo ErrHandler

    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject

    EvenQuotes = True
    If st = st_File Then

        Set F = m_FSO.GetFile(FileNameOrContents)
        Set T = F.OpenAsTextStream(ForReading, TriState)

        If T.AtEndOfStream Then
            T.Close: Set T = Nothing: Set F = Nothing
            Throw "File is empty"
        End If

        Do While Not T.AtEndOfStream And j <= MAX_CHUNKS
            j = j + 1
            Contents = T.Read(CHUNK_SIZE)
            For i = 1 To Len(Contents)
                Select Case Mid$(Contents, i, 1)
                    Case QuoteChar
                        EvenQuotes = Not EvenQuotes
                    Case ",", vbTab, "|", ";", ":"
                        If EvenQuotes Then
                            If Mid$(Contents, i, 1) <> DecimalSeparator Then
                                InferDelimiter = Mid$(Contents, i, 1)
                                T.Close: Set T = Nothing: Set F = Nothing
                                Exit Function
                            End If
                        End If
                End Select
            Next i
        Loop
        T.Close
    ElseIf st = st_String Then
        Contents = FileNameOrContents
        MaxChars = MAX_CHUNKS * CHUNK_SIZE
        If MaxChars > Len(Contents) Then MaxChars = Len(Contents)

        For i = 1 To MaxChars
            Select Case Mid$(Contents, i, 1)
                Case QuoteChar
                    EvenQuotes = Not EvenQuotes
                Case ",", vbTab, "|", ";", ":"
                    If EvenQuotes Then
                        If Mid$(Contents, i, 1) <> DecimalSeparator Then
                            InferDelimiter = Mid$(Contents, i, 1)
                            Exit Function
                        End If
                    End If
            End Select
        Next i
    Else
        Throw Err_SourceType
    End If

    'No commonly-used delimiter found in the file outside quoted regions _
     and in the first MAX_CHUNKS * CHUNK_SIZE characters. Assume comma _
     unless that's the decimal separator.
    
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
        Set T = Nothing: Set F = Nothing
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
Private Sub ParseDateFormat(ByVal DateFormat As String, ByRef DateOrder As Long, ByRef DateSeparator As String, _
    ByRef ISO8601, ByRef AcceptWithoutTimeZone As Boolean, ByRef AcceptWithTimeZone As Boolean)

    Dim Err_DateFormat As String

    On Error GoTo ErrHandler
    
    If UCase(DateFormat) = "ISO" Then
        ISO8601 = True
        AcceptWithoutTimeZone = True
        AcceptWithTimeZone = False
        Exit Sub
    ElseIf UCase(DateFormat) = "ISOZ" Then
        ISO8601 = True
        AcceptWithoutTimeZone = False
        AcceptWithTimeZone = True
        Exit Sub
    End If
    
    Err_DateFormat = "DateFormat not valid should be one of 'ISO', 'ISOZ', 'M-D-Y', 'D-M-Y', 'Y-M-D', " & _
        "'M/D/Y', 'D/M/Y' or 'Y/M/D'" & ". Omit to use the default date format on this PC which is """ & WindowsDateFormat() & """"
     
    'Replace repeated D's with a single D, etc since CastToDate only needs _
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
        Throw Err_DateFormat
    ElseIf Mid$(DateFormat, 2, 1) <> Mid$(DateFormat, 4, 1) Then
        Throw Err_DateFormat
    Else
        DateSeparator = Mid$(DateFormat, 2, 1)
        If DateSeparator <> "/" And DateSeparator <> "-" Then Throw Err_DateFormat
        Select Case UCase$(Left$(DateFormat, 1) & Mid$(DateFormat, 3, 1) & Right$(DateFormat, 1))
            Case "MDY"
                DateOrder = 0
            Case "DMY"
                DateOrder = 1
            Case "YMD"
                DateOrder = 2
            Case Else
                Throw Err_DateFormat
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
    Do While InStr(TheString, ChCh) > 0
        TheString = Replace(TheString, ChCh, TheChar)
    Loop
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WindowsDateFormat
' Purpose    : Returns a description of the system date format, used only for error string generation.
' -----------------------------------------------------------------------------------------------------------------------
Private Function WindowsDateFormat() As String
    Dim DS As String
    On Error GoTo ErrHandler
    DS = Application.International(xlDateSeparator)
    Select Case Application.International(xlDateOrder)
        Case 0
            WindowsDateFormat = "M" & DS & "D" & DS & "Y"
        Case 1
            WindowsDateFormat = "D" & DS & "M" & DS & "Y"
        Case 2
            WindowsDateFormat = "Y" & DS & "M" & DS & "D"
    End Select

    Exit Function
ErrHandler:
    WindowsDateFormat = "Cannot determine!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SplitCSVContents
' Purpose    : Used when the CSVReads arg FileName is a CSVString and Delimiter is FALSE.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SplitCSVContents(CSVContents As String, StartRow As Long, NumRows As Long)
    Dim Contents1D() As String
    Dim Contents2D() As String
    Dim i As Long
    Dim LoopTo As Long
    Dim LB As Long
    Dim LastElement As String

    On Error GoTo ErrHandler
    If Len(CSVContents) = 1 Then
        ReDim Contents1D(0 To 0)
    Else
        CSVContents = Replace(CSVContents, vbCrLf, vbLf)
        CSVContents = Replace(CSVContents, vbCr, vbLf)
        Contents1D = VBA.Split(CSVContents, vbLf)
        LastElement = Contents1D(UBound(Contents1D))
        If LastElement = "" Then 'Because last line of CSV may or may not terminate with a line feed
            ReDim Preserve Contents1D(LBound(Contents1D) To UBound(Contents1D) - 1)
        End If
    End If
    LB = LBound(Contents1D)

    Dim NumRowsInReturn As Long
    If NumRows = 0 Then
        NumRowsInReturn = (UBound(Contents1D) - LBound(Contents1D) + 1) - (StartRow - 1)
        LoopTo = NumRowsInReturn
    Else
        NumRowsInReturn = NumRows
        LoopTo = MinLngs(NumRowsInReturn, (UBound(Contents1D) - LBound(Contents1D) + 1) - (StartRow - 1))
    End If

    ReDim Contents2D(1 To NumRowsInReturn, 1 To 1)

    For i = 1 To LoopTo
        Contents2D(i, 1) = Contents1D(LB - 1 + i + StartRow - 1)
    Next i

    SplitCSVContents = Contents2D

    Exit Function
ErrHandler:
    Throw "#SplitCSVContents: " & Err.Description & "!"
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

        If InStr(ReadAll, vbCr) > 0 Then
            ReadAll = Replace(ReadAll, vbCrLf, vbLf)
            ReadAll = Replace(ReadAll, vbCr, vbLf)
        End If

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
' Purpose    : Parse the contents of a CSV file. Returns a string Buffer together with arrays which assist splitting
'              Buffer into a two-dimensional array.
' Parameters :
'  ContentsOrStream: The contents of a CSV file as a string, or else a Scripting.TextStream.
'  QuoteChar       : The quote character, usually ascii 34 ("), which allow fields to contain characters that would
'                    otherwise be significant to parsing, such as delimiters or new line characters.
'  Delimiter       : The string that separates fields within each line. Typically a single character, but needn't be.
'  SkipToRow       : Rows in the file prior to SkipToRow are ignored.
'  Comment         : Lines in the file that start with these characters will be ignored, handled by method SkipLines.
'  IgnoreRepeated  : If true then parsing ignores delimiters at the start of lines, consecutive delimiters and delimiters
'                    at the end of lines.
'  SkipToRow       : The first line of the file to appear in the return from CSVRead. However, we need to parse earlier
'                    lines to identify where SkipToRow starts in the file - see variable HaveReachedSkipToRow.
'  HeaderRowNum    : The row number of the headers in the file, must be less than or equal to SkipToRow.
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
'  QuoteCounts     : Set to an array of size at least NumFields. Element k gives the number of QuoteChars that appear in
'                    the kth field.
'  HeaderRow       : Set equal to the contents of the header row in the file, no type conversion other than unquoting.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseCSVContents(ContentsOrStream As Variant, QuoteChar As String, Delimiter As String, Comment As String, _
    IgnoreEmptyLines As Boolean, IgnoreRepeated As Boolean, SkipToRow As Long, HeaderRowNum As Long, NumRows As Long, ByRef NumRowsFound As Long, _
    ByRef NumColsFound As Long, ByRef NumFields As Long, ByRef Ragged As Boolean, ByRef Starts() As Long, _
    ByRef Lengths() As Long, RowIndexes() As Long, ColIndexes() As Long, QuoteCounts() As Long, ByRef HeaderRow) As String

    Const Err_ContentsOrStream = "ContentsOrStream must either be a string or a TextStream"
    Const Err_Delimiter = "Delimiter must not be the null string"
    Dim Buffer As String
    Dim BufferUpdatedTo As Long
    Dim DoSkipping As Boolean
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
    HeaderRow = Empty
    
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
    If LComment > 0 Or IgnoreEmptyLines Then
        DoSkipping = True
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

    ReDim Starts(1 To 8): ReDim Lengths(1 To 8): ReDim RowIndexes(1 To 8)
    ReDim ColIndexes(1 To 8): ReDim QuoteCounts(1 To 8)
    
    LDlm = Len(Delimiter)
    If LDlm = 0 Then Throw Err_Delimiter 'Avoid infinite loop!
    OrigLen = Len(Buffer)
    If Not Streaming Then
        'Ensure Buffer terminates with vbCrLf
        If Right$(Buffer, 1) <> vbCr And Right$(Buffer, 1) <> vbLf Then
            Buffer = Buffer & vbCrLf
        ElseIf Right$(Buffer, 1) = vbCr Then
            Buffer = Buffer & vbLf
        End If
        BufferUpdatedTo = Len(Buffer)
    End If
    
    i = 0: j = 1
    
    If DoSkipping Then
        SkipLines Streaming, Comment, LComment, IgnoreEmptyLines, T, Delimiter, Buffer, i, QuoteChar, PosLF, PosCR, BufferUpdatedTo
    End If
    
    If IgnoreRepeated Then
        'IgnoreRepeated: Handle repeated delimiters at the start of the first line
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
                    'Found line ending
                    Lengths(j) = i - Starts(j)
                    If Which = 3 Then 'Found a vbCr
                        If Mid$(Buffer, i + 1, 1) = vbLf Then
                            'Ending is Windows rather than Mac or Unix.
                            i = i + 1
                        End If
                    End If
                    
                    If DoSkipping Then
                        SkipLines Streaming, Comment, LComment, IgnoreEmptyLines, T, Delimiter, Buffer, i, QuoteChar, PosLF, PosCR, BufferUpdatedTo
                    End If
                    
                    If IgnoreRepeated Then
                        'IgnoreRepeated: Handle repeated delimiters at the end of the line, all but one will have already been handled.
                        If Lengths(j) = 0 Then
                            If ColNum > 1 Then
                                j = j - 1
                                ColNum = ColNum - 1
                                NumFields = NumFields - 1
                            End If
                        End If
                        'IgnoreRepeated: handle delimiters at the start of the next line to be parsed
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
                    QuoteCounts(j) = QuoteCount: QuoteCount = 0
                    
                    If RowNum = 1 Then
                        If SkipToRow = 1 Then
                            If HeaderRowNum = 1 Then
                                HeaderRow = GetLastParsedRow(Buffer, Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts, j)
                            End If
                        End If
                    End If
                    If Not HaveReachedSkipToRow Then
                        If RowNum = HeaderRowNum Then
                            HeaderRow = GetLastParsedRow(Buffer, Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts, j)
                        End If
                    End If
                    
                    ColNum = 1: RowNum = RowNum + 1
                    
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
            Dim RowDescription As String
            If IgnoreEmptyLines And Len(Comment) > 0 Then
                RowDescription = "not commented, not empty "
            ElseIf IgnoreEmptyLines Then
                RowDescription = "not empty "
            ElseIf Len(Comment) > 0 Then
                RowDescription = "not commented "
            End If
                             
            Throw "SkipToRow (" & CStr(SkipToRow) & ") exceeds the number of " & RowDescription & "rows in the file (" & CStr(NumRowsInFile) & ")"
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
' Procedure  : GetLastParsedRow
' Purpose    : For use during parsing (fn ParseCSVContents) to grab the header row (which may or may not be part of the
'              function return). The argument j should point into the Starts, Lengths etc arrays, pointing to the last
'              field in the header row
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetLastParsedRow(Buffer As String, Starts() As Long, Lengths() As Long, RowIndexes() As Long, ColIndexes() As Long, QuoteCounts() As Long, j As Long)
    Dim NC As Long

    Dim RowNum As Long
    Dim i As Long
    Dim Field As String

    Dim Res() As String

    On Error GoTo ErrHandler
    RowNum = RowIndexes(j)
    NC = ColIndexes(j)

    ReDim Res(1 To 1, 1 To NC)
    For i = j To j - NC + 1 Step -1
        Field = Mid$(Buffer, Starts(i), Lengths(i))
        Res(1, NC + i - j) = Unquote(Field, """", QuoteCounts(i))
    Next i
    GetLastParsedRow = Res

    Exit Function
ErrHandler:
    Throw "#GetLastParsedRow: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SkipLines
' Purpose    : Sub-routine of ParseCSVContents. Skip a commented or empty row by incrementing i to the position of the line feed
'              just before the next not-commented line.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SkipLines(Streaming As Boolean, Comment As String, LComment As Long, IgnoreEmptyLines As Boolean, T As Scripting.TextStream, _
    Delimiter As String, ByRef Buffer As String, ByRef i As Long, QuoteChar As String, ByVal PosLF As Long, _
    ByVal PosCR As Long, ByRef BufferUpdatedTo As Long)
    
    Dim LookAheadBy As Long
    Dim SkipThisLine As Boolean
    Do
        If Streaming Then
            LookAheadBy = MaxLngs(LComment, 2)
        
            If i + LookAheadBy > BufferUpdatedTo Then
                If Not T.AtEndOfStream Then
                    Call GetMoreFromStream(T, Delimiter, QuoteChar, Buffer, BufferUpdatedTo)
                End If
            End If
        End If

        SkipThisLine = False
        If LComment > 0 Then
            If Mid$(Buffer, i + 1, LComment) = Comment Then
                SkipThisLine = True
            End If
        End If
        If IgnoreEmptyLines Then
            Select Case Mid$(Buffer, i + 1, 1)
                Case vbLf, vbCr
                    SkipThisLine = True
            End Select
        End If

        If SkipThisLine Then
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
' Purpose    : Sub-routine of ParseCSVContents. Returns the location in the buffer of the first-encountered string amongst
'              the elements of SearchFor, starting the search at point SearchFrom and finishing the search at point
'              BufferUpdatedTo. If none found in that region returns BufferUpdatedTo + 1. Otherwise returns the location
'              of the first found and sets the by-reference argument Which to indicate which element of SearchFor was the
'              first to be found.
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
' Purpose    : Sub-routine of ParseCSVContents. Returns the first point in SearchWithin at which one of the elements of
'              SearchFor is found, search is restricted to region [StartingAt, EndingAt] and Which is updated with the
'              index identifying which was the first of the strings to be found.
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
' Procedure  : GetMoreFromStream, Sub-routine of ParseCSVContents
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
            If Right$(NewChars, 1) <> vbCr And Right$(NewChars, 1) <> vbLf Then
                NewChars = NewChars & vbCrLf
            ElseIf Right$(NewChars, 1) = vbCr Then
                NewChars = NewChars & vbLf
            End If
        End If

        NCharsToWriteToBuffer = Len(NewChars) + Len(Delimiter) + 3

        If BufferUpdatedTo + NCharsToWriteToBuffer > Len(Buffer) Then
            ExpandBufferBy = MaxLngs(Len(Buffer), NCharsToWriteToBuffer)
            Buffer = Buffer & String(ExpandBufferBy, "?")
        End If
        
        Mid$(Buffer, BufferUpdatedTo + 1, Len(NewChars)) = NewChars
        BufferUpdatedTo = BufferUpdatedTo + Len(NewChars)

        OKToExit = True
        'Ensure we don't leave the buffer updated to part way through a two-character end of line marker.
        If Right$(NewChars, 1) = vbCr Then
            OKToExit = False
        End If
        'Ensure we don't leave the buffer updated to a point part-way through a multi-character delimiter
        If Len(Delimiter) > 1 Then
            For i = 1 To Len(Delimiter) - 1
                If Mid$(Buffer, BufferUpdatedTo - i + 1, i) = Left$(Delimiter, i) Then
                    OKToExit = False
                    Exit For
                End If
            Next i
            If Mid$(Buffer, BufferUpdatedTo - Len(Delimiter) + 1, Len(Delimiter)) = Delimiter Then
                OKToExit = True
            End If
        End If
        If OKToExit Then Exit Do
    Loop

    'Line below arranges that when calling Instr(Buffer,....) we don't pointlessly scan the space characters _
     we can be sure that there is space in the buffer to write the extra characters thanks to
    Mid$(Buffer, BufferUpdatedTo + 1, Len(Delimiter) + 3) = vbCrLf & QuoteChar & Delimiter

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

Private Function MinLngs(x As Long, y As Long) As Long
    If x > y Then
        MinLngs = y
    Else
        MinLngs = x
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CountQuotes
' Purpose    : Count the quotes in a string, only used when applying column-by-column type conversion, because in that case
'              it's not possible to use the count of quotes made at parsing time which is organised row-by-row.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CountQuotes(Str As String, QuoteChar As String)
    Dim N As Long
    Dim pos As Long

    Do
        pos = InStr(pos + 1, Str, QuoteChar)
        If pos = 0 Then
            CountQuotes = N
            Exit Function
        End If
        N = N + 1
    Loop
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConvertField
' Purpose    : Convert a field in the file into an element of the returned array.
' Parameters :
'General
'  Field                : Field, i.e. characters from the file between successive delimiters.
'  AnyConversion        : Is any type conversion to take place? i.e. processing other than trimming whitespace and
'                         unquoting.
'  FieldLength          : The length of Field.
'Whitespace and Quotes
'  TrimFields           : Should leading and trailing spaces be trimmed from fields? For quoted fields, this will not
'                         remove spaces between the quotes.
'  QuoteChar            : The quote character, typically ". No support for different opening and closing quote characters
'                         or different escape character.
'  QuoteCount           : How many quote characters does Field contain?
'  ConvertQuoted        : Should quoted fields (after quote removal) be converted according to args ShowNumbersAsNumbers
'                         ShowDatesAsDates, and the contents of Sentinels.
'Numbers
'  ShowNumbersAsNumbers : If Field is a string representation of a number should the function return that number?
'  SepStandard          : Is the decimal separator the same as the system defaults? If True then the next two arguments
'                         are ignored.
'  DecimalSeparator     : The decimal separator used in the input string.
'  SysDecimalSeparator  : The default decimal separator on the system.
'Dates
'  ShowDatesAsDates     : If Field is a string representation of a date should the function return that date?
'  ISO8601              : If field is a date, does it respect (a subset of) ISO8601?
'  AcceptWithoutTimeZone: In the case of ISO8601 dates, should conversion be dates-with-time that have no time zone
'                         information?
'  AcceptWithTimeZone   : In the case of ISO8601 dates, should conversion be dates-with-time that have time zone
'                         information?
'  DateOrder            : If Field is a string representation what order of parts must it respect (not relevant if
'                         ISO8601 is True) 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  DateSeparator        : The date separator, must be one of "-" or "/".
'  SysDateOrder         : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  SysDateSeparator     : The Windows system date separator.
'Booleans, Errors, Missings
'  AnySentinels         : Does the sentinel dictionary have any elements?
'  Sentinels            : A dictionary of Sentinels. If Sentinels.Exists(Field) Then ConvertField = Sentinels(Field)
'  MaxSentinelLength    : The maximum length of the keys of Sentinels.
'  ShowMissingsAs       : The value to which missing fields (consecutive delimiters) are converted. If CSVRead has a
'                         MissingStrings argument then values matching those strings are also converted to ShowMissingsAs,
'                         thanks to method MakeSentinels.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ConvertField(Field As String, AnyConversion As Boolean, FieldLength As Long, TrimFields As Boolean, _
    QuoteChar As String, QuoteCount As Long, ConvertQuoted As Boolean, _
    ShowNumbersAsNumbers As Boolean, SepStandard As Boolean, DecimalSeparator As String, SysDecimalSeparator As String, _
    ShowDatesAsDates As Boolean, ISO8601 As Boolean, AcceptWithoutTimeZone As Boolean, AcceptWithTimeZone As Boolean, _
    DateOrder As Long, DateSeparator As String, SysDateOrder As Long, SysDateSeparator As String, _
    AnySentinels As Boolean, Sentinels As Dictionary, MaxSentinelLength As Long, _
    ShowMissingsAs As Variant)

    Dim Converted As Boolean
    Dim dblResult As Double
    Dim dtResult As Date
    Dim isQuoted As Boolean

    If TrimFields Then
        If Left$(Field, 1) = " " Then
            Field = Trim$(Field)
            FieldLength = Len(Field)
        ElseIf Right$(Field, 1) = " " Then
            Field = Trim$(Field)
            FieldLength = Len(Field)
        End If
    End If

    If FieldLength = 0 Then
        ConvertField = ShowMissingsAs
        Exit Function
    End If

    If Not AnyConversion Then
        If QuoteCount = 0 Then
            ConvertField = Field
            Exit Function
        End If
    End If

    If QuoteCount > 0 Then
        If Left$(Field, 1) = QuoteChar Then 'NOTE definition of quoted is both first and last characters must be quote characters
            If Right$(QuoteChar, 1) = QuoteChar Then
                isQuoted = True
                Field = Mid$(Field, 2, FieldLength - 2)
                If QuoteCount > 2 Then
                    Field = Replace(Field, QuoteChar & QuoteChar, QuoteChar) 'TODO QuoteCharTwice arg
                End If
                If ConvertQuoted Then
                    FieldLength = Len(Field)
                Else
                    ConvertField = Field
                    Exit Function
                End If
            End If
        End If
    End If

    If AnySentinels Then
        If FieldLength <= MaxSentinelLength Then
            If Sentinels.Exists(Field) Then
                ConvertField = Sentinels(Field)
                Exit Function
            End If
        End If
    End If

    If Not ConvertQuoted Then
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
        If ISO8601 Then
            CastISO8601 Field, dtResult, Converted, AcceptWithoutTimeZone, AcceptWithTimeZone
        Else
            CastToDate Field, dtResult, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
        End If
        If Not Converted Then
            If InStr(Field, ":") > 0 Then
                CastToTime Field, dtResult, Converted
                If Not Converted Then
                    CastToTimeB Field, dtResult, Converted
                End If
            End If
        End If
        If Converted Then
            ConvertField = dtResult
            Exit Function
        End If
    End If

    ConvertField = Field
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Unquote
' Purpose    : Unquote a field.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Unquote(ByVal Field As String, QuoteChar As String, QuoteCount As Long)

    On Error GoTo ErrHandler
    If QuoteCount > 0 Then
        If Left$(Field, 1) = QuoteChar Then 'NOTE definition of quoted is both first and last characters must be quote characters
            If Right$(QuoteChar, 1) = QuoteChar Then
                Field = Mid$(Field, 2, Len(Field) - 2)
                If QuoteCount > 2 Then
                    Field = Replace(Field, QuoteChar & QuoteChar, QuoteChar) 'TODO QuoteCharTwice arg
                End If
            End If
        End If
    End If
    Unquote = Field

    Exit Function
ErrHandler:
    Throw "#Unquote: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble, sub-routine of ConvertField
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
' Procedure  : TestCastToDate
' Purpose    : Quick test of method CastToDate, See also RunTests() for much more comprehensive testing
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestCastToDate()
    Dim strIn As String
    Dim dtOut As Date
    Dim DateOrder As Long
    Dim DateSeparator As String
    Dim SysDateOrder As Long
    Dim SysDateSeparator As String
    Dim Converted As Boolean
    Dim i As Long
    
    strIn = "2021-08-30 4:1:2.123": DateOrder = 2
    
    DateSeparator = "-"
    SysDateOrder = Application.International(xlDateOrder)
    SysDateSeparator = Application.International(xlDateSeparator)

    CastToDate strIn, dtOut, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted

    Debug.Print "StrIn = """ & strIn & """", "DateOrder = " & DateOrder, "dtOut = " & Format(dtOut, "yyyy-mmm-dd hh:mm:ss")

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDate, sub-routine of ConvertField
' Purpose    : In-place conversion of a string that looks like a date into a Long or Date. No error if string cannot be
'              converted to date. Converts Dates, DateTimes and Times. Times in very simple format hh:mm:ss
'              Does not handle ISO8601 - see alternative function CastISO8601
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
    Dim TimePart As String
    Dim TimePartConverted As Date
    Dim Converted2 As Boolean
    
    On Error GoTo ErrHandler
    
    pos1 = InStr(strIn, DateSeparator)
    If pos1 = 0 Then Exit Sub
    pos2 = InStr(pos1 + 1, strIn, DateSeparator)
    If pos2 = 0 Then Exit Sub

    If DateOrder = 0 Then 'M-D-Y
        m = Left$(strIn, pos1 - 1)
        D = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        y = Mid$(strIn, pos2 + 1)
        SplitOutTime y, TimePart
    ElseIf DateOrder = 1 Then 'D-M-Y
        D = Left$(strIn, pos1 - 1)
        m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        y = Mid$(strIn, pos2 + 1)
        SplitOutTime y, TimePart
    ElseIf DateOrder = 2 Then 'Y-M-D
        y = Left$(strIn, pos1 - 1)
        m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        D = Mid$(strIn, pos2 + 1)
        SplitOutTime D, TimePart
    Else
        Throw "DateOrder must be 0, 1, or 2"
    End If
    If InStr(TimePart, ".") = 0 Then
        If SysDateOrder = 0 Then
            dtOut = CDate(m & SysDateSeparator & D & SysDateSeparator & y & TimePart)
            Converted = True
        ElseIf SysDateOrder = 1 Then
            dtOut = CDate(D & SysDateSeparator & m & SysDateSeparator & y & TimePart)
            Converted = True
        ElseIf SysDateOrder = 2 Then
            dtOut = CDate(y & SysDateSeparator & m & SysDateSeparator & D & TimePart)
            Converted = True
        End If
    Else
        If Left$(TimePart, 1) <> " " Then Exit Sub
        CastToTimeB Mid$(TimePart, 2), TimePartConverted, Converted2
        If Converted2 Then
            If SysDateOrder = 0 Then
                dtOut = CDate(m & SysDateSeparator & D & SysDateSeparator & y) + TimePartConverted
                Converted = True
            ElseIf SysDateOrder = 1 Then
                dtOut = CDate(D & SysDateSeparator & m & SysDateSeparator & y) + TimePartConverted
                Converted = True
            ElseIf SysDateOrder = 2 Then
                dtOut = CDate(y & SysDateSeparator & m & SysDateSeparator & D) + TimePartConverted
                Converted = True
            End If
    
        End If
    End If

    Exit Sub
ErrHandler:
    'Do nothing - was not a string representing a date with the specified date order and date separator.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SplitOutTime, sub-routine of ConvertField
' Purpose    : A string to represent DateTime comes in four parts, the first three being Y, M, D (though not necessarily
'             in that  order) with the (optional) fourth part being time. This method splits the third part that may
'             include time into the third part and the time part.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SplitOutTime(ByRef ThirdPart As String, ByRef TimePart)
    Dim i As Long
    Dim LoopTo As Long
    
    LoopTo = Len(ThirdPart)
    If LoopTo > 5 Then LoopTo = 5

    For i = 1 To LoopTo
        Select Case AscW(Mid$(ThirdPart, i, 1))
            Case 48 To 57
            Case Else
                TimePart = Mid$(ThirdPart, i)
                ThirdPart = Left$(ThirdPart, i - 1)
                Exit Function
        End Select
    Next i
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OStoEOL
' Purpose    : Convert text describing an operating system to the end-of-line marker employed. Note that "Mac" converts
'              to vbCr but Apple operating systems since OSX use vbLf, matching Unix.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OStoEOL(OS As String, ArgName As String) As String

    Const Err_Invalid = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associated end of line characters."

    On Error GoTo ErrHandler
    Select Case LCase(OS)
        Case "windows", vbCrLf, "crlf"
            OStoEOL = vbCrLf
        Case "unix", "linux", vbLf, "lf"
            OStoEOL = vbLf
        Case "mac", vbCr, "cr"
            OStoEOL = vbCr
        Case Else
            Throw ArgName & Err_Invalid
    End Select

    Exit Function
ErrHandler:
    Throw "#OStoEOL: " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVWrite
' Purpose   : Creates a comma-separated file on disk containing Data. Any existing file of the same
'             name is overwritten. If successful, the function returns FileName, otherwise an "error
'             string" (starts with `#`, ends with `!`) describing what went wrong.
' Arguments
' Data      : An array of data, or an Excel range. Elements may be strings, numbers, dates, Booleans, empty,
'             Excel errors or null values.
' FileName  : The full name of the file, including the path. Alternatively, if FileName is omitted, then the
'             function returns Data converted CSV-style to a string.
' QuoteAllStrings: If `TRUE` (the default) then elements of Data that are strings are quoted before being
'             written to file, other elements (Numbers, Booleans, Errors) are not quoted. If `FALSE` then
'             the only elements of Data that are quoted are strings containing Delimiter, line feed,
'             carriage return or double quote. In all cases, double quotes are escaped by another double
'             quote.
' DateFormat: A format string that determines how dates, including cells formatted as dates, appear in the
'             file. If omitted, defaults to `yyyy-mm-dd`.
' DateTimeFormat: Format for datetimes. Defaults to `ISO` which abbreviates `yyyy-mm-ddThh:mm:ss`. Use
'             `ISOZ` for ISO8601 format with time zone the same as the PC's clock. Use with care, daylight
'             saving may be inconsistent across the datetimes in data.
' Delimiter : The delimiter string, if omitted defaults to a comma. Delimiter may have more than one
'             character.
' Unicode   : If `FALSE` (the default) the file written will be encoded UTF-8. If TRUE the file written will
'             be encoded UTF-16 LE BOM. An error will result if this argument is `FALSE` but Data contains
'             strings with characters whose code points exceed 255.
' EOL       : Controls the line endings of the file written. Enter `Windows` (the default), `Unix` or `Mac`.
'             Also supports the line-ending characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13)
'             or the strings `CRLF`, `LF` or `C`. The last line of the file is written with a line ending.
'
' Notes     : See also companion function CSVRead.
'
'             For discussion of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
'---------------------------------------------------------------------------------------------------------
Public Function CSVWrite(ByVal Data As Variant, Optional FileName As String, _
    Optional QuoteAllStrings As Boolean = True, Optional DateFormat As String = "yyyy-mm-dd", _
    Optional ByVal DateTimeFormat As String = "ISO", _
    Optional Delimiter As String = ",", Optional Unicode As Boolean, _
    Optional ByVal EOL As String = "")
Attribute CSVWrite.VB_Description = "Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an ""error string"" (starts with `#`, ends with `!`) describing what went wrong."
Attribute CSVWrite.VB_ProcData.VB_Invoke_Func = " \n14"

    Const DQ = """"
    Const Err_Delimiter = "Delimiter must have at least one character and cannot start with a " & _
        "double  quote, line feed or carriage return"
    Const Err_Dimensions = "Data must be a range or a 2-dimensional array"
    
    Dim EOLIsWindows As Boolean
    Dim ErrRet As String
    Dim i As Long
    Dim j As Long
    Dim Lines() As String
    Dim OneLine() As String
    Dim OneLineJoined As String
    Dim T As Scripting.TextStream
    Dim WriteToFile As Boolean

    On Error GoTo ErrHandler

    WriteToFile = Len(FileName) > 0

    If EOL = "" Then
        If WriteToFile Then
            EOL = vbCrLf
        Else
            EOL = vbLf
        End If
    End If

    EOL = OStoEOL(EOL, "EOL")
    EOLIsWindows = EOL = vbCrLf

    If Len(Delimiter) = 0 Or Left$(Delimiter, 1) = DQ Or Left$(Delimiter, 1) = vbLf Or Left$(Delimiter, 1) = vbCr Then
        Throw Err_Delimiter
    End If
    
    Select Case UCase(DateTimeFormat)
        Case "ISO"
            DateTimeFormat = "yyyy-mm-ddThh:mm:ss"
        Case "ISOZ"
            DateTimeFormat = ISOZFormatString()
    End Select

    If TypeName(Data) = "Range" Then
        'Preserve elements of type Date by using .Value, not .Value2
        Data = Data.value
    End If
    
    If NumDimensions(Data) <> 2 Then Throw Err_Dimensions
    ReDim OneLine(LBound(Data, 2) To UBound(Data, 2))
    
    If WriteToFile Then
        If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
        Set T = m_FSO.CreateTextFile(FileName, True, Unicode)
        
        For i = LBound(Data) To UBound(Data)
            For j = LBound(Data, 2) To UBound(Data, 2)
                OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, ",")
            Next j
            OneLineJoined = VBA.Join(OneLine, Delimiter)
            WriteLineWrap T, OneLineJoined, EOLIsWindows, EOL, Unicode
        Next i

        T.Close: Set T = Nothing
        CSVWrite = FileName
    Else

        ReDim Lines(LBound(Data) To UBound(Data) + 1) 'add one to ensure that result has a terminating EOL
        
        For i = LBound(Data) To UBound(Data)
            For j = LBound(Data, 2) To UBound(Data, 2)
                OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, ",")
            Next j
            Lines(i) = VBA.Join(OneLine, Delimiter)
        Next i
        CSVWrite = VBA.Join(Lines, EOL)
        If Len(CSVWrite) >= 32768 Then
            If TypeName(Application.Caller) = "Range" Then
                Throw "Cannot return string of length " & Format(CStr(Len(CSVWrite)), "#,###") & " to a cell of an Excel worksheet"
            End If
        End If
    End If
    
    Exit Function
ErrHandler:
    ErrRet = "#CSVWrite: " & Err.Description & "!"
    If Not T Is Nothing Then
        T.Close
        Set T = Nothing
    End If
    If m_ErrorStyle = es_ReturnString Then
        CSVWrite = ErrRet
    Else
        Throw ErrRet
    End If
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVWrite
' Purpose    : Register the function CSVWrite with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
'---------------------------------------------------------------------------------------------------------
Sub RegisterCSVWrite()
    Const Description = "Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an ""error string"" (starts with `#`, ends with `!`) describing what went wrong."
    Dim ArgumentDescriptions() As String

    On Error GoTo ErrHandler

    ReDim ArgumentDescriptions(1 To 8)
    ArgumentDescriptions(1) = "An array of data, or an Excel range. Elements may be strings, numbers, dates, Booleans, empty, Excel errors or null values."
    ArgumentDescriptions(2) = "The full name of the file, including the path. Alternatively, if FileName is omitted, then the function returns Data converted CSV-style to a string."
    ArgumentDescriptions(3) = "If TRUE (the default) then all strings in Data are quoted before being written to file. If FALSE only strings containing Delimiter, line feed, carriage return or double quote are quoted. Double quotes are always escaped by another double quote."
    ArgumentDescriptions(4) = "A format string that determines how dates, including cells formatted as dates, appear in the file. If omitted, defaults to `yyyy-mm-dd`."
    ArgumentDescriptions(5) = "Format for datetimes. Defaults to `ISO` which abbreviates `yyyy-mm-ddThh:mm:ss`. Use `ISOZ` for ISO8601 format with time zone the same as the PC's clock. Use with care, daylight saving may be inconsistent across the datetimes in data."
    ArgumentDescriptions(6) = "The delimiter string, if omitted defaults to a comma. Delimiter may have more than one character."
    ArgumentDescriptions(7) = "If `FALSE` (the default) the file written will be encoded UTF-8. If TRUE the file written will be encoded UTF-16 LE BOM. An error will result if this argument is `FALSE` but Data contains strings with characters whose code points exceed 255."
    ArgumentDescriptions(8) = "Controls the line endings of the file written. Enter `Windows` (the default), `Unix` or `Mac`. Also supports the line-ending characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13) or the strings `CRLF`, `LF` or `C`."
    Application.MacroOptions "CSVWrite", Description, , , , , , , , , ArgumentDescriptions
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function CSVWrite failed with error: " + Err.Description
End Sub

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
                If AscW(Mid$(text, i, 1)) > 255 Then
                    ErrDesc = "Data contains characters with code points above 255 (first found has code " & CStr(AscW(Mid$(text, i, 1))) & _
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
Private Function Encode(x As Variant, QuoteAllStrings As Boolean, DateFormat As String, DateTimeFormat As String, Delim As String) As String
    Const DQ = """"
    Const DQ2 = """"""

    On Error GoTo ErrHandler
    Select Case VarType(x)

        Case vbString
            If InStr(x, DQ) > 0 Then
                Encode = DQ & Replace(x, DQ, DQ2) & DQ
            ElseIf QuoteAllStrings Then
                Encode = DQ & x & DQ
            ElseIf InStr(x, vbCr) > 0 Then
                Encode = DQ & x & DQ
            ElseIf InStr(x, vbLf) > 0 Then
                Encode = DQ & x & DQ
            ElseIf InStr(x, Delim) > 0 Then
                Encode = DQ & x & DQ
            Else
                Encode = x
            End If
        Case vbBoolean, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbEmpty 'vbLongLong - not available on 16 bit.
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
            Select Case CStr(x) 'Editing this case statement? Edit also its inverse, see method MakeSentinels
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
            Throw "Cannot convert variant of type " & TypeName(x) & " to String"
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

Sub g(Data)
    Application.Run "SolumAddin.xlam!g", Data
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeSentinels
' Purpose    : Returns a Dictionary keyed on strings for which if a key to the dictionary is a field of the CSV file then
'              that field should be converted to the associated item value. Handles Booleans, Missings and Excel errors.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub MakeSentinels(ByRef Sentinels As Scripting.Dictionary, ByRef MaxLength As Long, ByRef AnySentinels As Boolean, ShowBooleansAsBooleans As Boolean, _
    ShowErrorsAsErrors As Boolean, ByRef ShowMissingsAs As Variant, Optional TrueStrings As Variant, _
    Optional FalseStrings As Variant, Optional MissingStrings As Variant)

    Const Err_FalseStrings = "FalseStrings must be omitted or provided as a string or an array of strings that represent Boolean value False"
    Const Err_MissingStrings = "MissingStrings must be omitted or provided a string or an array of strings that represent missing values"
    Const Err_ShowMissings = "ShowMissingsAs has an illegal value, such as an array or an object"
    Const Err_TrueStrings = "TrueStrings must be omitted or provided as string or an array of strings that represent Boolean value True"
    Const Err_TrueStrings2 = "TrueStrings has been provided, but type conversion for Booleans is not switched on for any column"
    Const Err_FalseStrings2 = "FalseStrings has been provided, but type conversion for Booleans is not switched on for any column"

    On Error GoTo ErrHandler

    If IsMissing(ShowMissingsAs) Then ShowMissingsAs = Empty
    Select Case VarType(ShowMissingsAs)
        Case vbObject, vbArray, vbByte, vbDataObject, vbUserDefinedType, vbVariant
            Throw Err_ShowMissings
    End Select
    
    If Not IsMissing(MissingStrings) And Not IsEmpty(MissingStrings) Then
        AddKeysToDict Sentinels, MissingStrings, ShowMissingsAs, Err_MissingStrings
    End If

    If ShowBooleansAsBooleans Then
        If IsMissing(TrueStrings) Or IsEmpty(TrueStrings) Then
            AddKeysToDict Sentinels, Array("TRUE", "true", "True"), True, Err_TrueStrings
        Else
            AddKeysToDict Sentinels, TrueStrings, True, Err_TrueStrings
        End If
        If IsMissing(FalseStrings) Or IsEmpty(FalseStrings) Then
            AddKeysToDict Sentinels, Array("FALSE", "false", "False"), False, Err_FalseStrings
        Else
            AddKeysToDict Sentinels, FalseStrings, False, Err_FalseStrings
        End If
    Else
        If Not (IsMissing(TrueStrings) Or IsEmpty(TrueStrings)) Then
            Throw Err_TrueStrings2
        End If
        If Not (IsMissing(FalseStrings) Or IsEmpty(FalseStrings)) Then
            Throw Err_FalseStrings2
        End If
    End If
    
    If ShowErrorsAsErrors Then
        AddKeyToDict Sentinels, "#DIV/0!", CVErr(xlErrDiv0)
        AddKeyToDict Sentinels, "#NAME?", CVErr(xlErrName)
        AddKeyToDict Sentinels, "#REF!", CVErr(xlErrRef)
        AddKeyToDict Sentinels, "#NUM!", CVErr(xlErrNum)
        AddKeyToDict Sentinels, "#NULL!", CVErr(xlErrNull)
        AddKeyToDict Sentinels, "#N/A", CVErr(xlErrNA)
        AddKeyToDict Sentinels, "#VALUE!", CVErr(xlErrValue)
        AddKeyToDict Sentinels, "#SPILL!", CVErr(2045)
        AddKeyToDict Sentinels, "#BLOCKED!", CVErr(2047)
        AddKeyToDict Sentinels, "#CONNECT!", CVErr(2046)
        AddKeyToDict Sentinels, "#UNKNOWN!", CVErr(2048)
        AddKeyToDict Sentinels, "#GETTING_DATA!", CVErr(2043)
        AddKeyToDict Sentinels, "#FIELD!", CVErr(2049)
        AddKeyToDict Sentinels, "#CALC!", CVErr(2050)
    End If

    Dim k As Variant
    MaxLength = 0
    For Each k In Sentinels.Keys
        If Len(k) > MaxLength Then MaxLength = Len(k)
    Next
    AnySentinels = Sentinels.Count > 0

    Exit Sub
ErrHandler:
    Throw "#MakeSentinels: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddKeysToDict, Sub-routine of MakeSentinels
' Purpose    : Broadcast AddKeyToDict over an array of keys.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AddKeysToDict(ByRef Sentinels As Scripting.Dictionary, ByVal Keys As Variant, item As Variant, FriendlyErrorString As String)
    On Error GoTo ErrHandler

    Dim i As Long
    Dim j As Long
  
    If TypeName(Keys) = "Range" Then
        Keys = Keys.value
    End If
    
    If VarType(Keys) = vbString Then
        If InStr(Keys, ",") > 0 Then
            Keys = VBA.Split(Keys, ",")
        End If
    End If
    
    Select Case NumDimensions(Keys)
        Case 0
            AddKeyToDict Sentinels, Keys, item, FriendlyErrorString
        Case 1
            For i = LBound(Keys) To UBound(Keys)
                AddKeyToDict Sentinels, Keys(i), item, FriendlyErrorString
            Next i
        Case 2
            For i = LBound(Keys, 1) To UBound(Keys, 1)
                For j = LBound(Keys, 2) To UBound(Keys, 2)
                    AddKeyToDict Sentinels, Keys(i, j), item, FriendlyErrorString
                Next j
            Next i
        Case Else
            Throw FriendlyErrorString
    End Select
    Exit Sub
ErrHandler:
    Throw "#AddKeysToDict: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddKeyToDict, Sub-routine of MakeSentinels
' Purpose    : Wrap .Add method to have more helpful error message if things go awry.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AddKeyToDict(ByRef Sentinels As Scripting.Dictionary, Key As Variant, item As Variant, Optional FriendlyErrorString As String)

    Dim FoundRepeated As Boolean

    On Error GoTo ErrHandler

    If VarType(Key) <> vbString Then Throw FriendlyErrorString & " but '" & CStr(Key) & "' is of type " & TypeName(Key)
    If Len(Key) = 0 Then Exit Sub
    
    If Not Sentinels.Exists(Key) Then
        Sentinels.Add Key, item
    Else
        FoundRepeated = True
        If VarType(item) = VarType(Sentinels(Key)) Then
            If item = Sentinels(Key) Then
                FoundRepeated = False
            End If
        End If
    End If

    If FoundRepeated Then
        Throw "There is a conflicting definition of what the string '" & Key & _
            "' should be converted to, both the " & TypeName(item) & " value '" & CStr(item) & _
            "' and the " & TypeName(Sentinels(Key)) & " value '" & CStr(Sentinels(Key)) & _
            "' have been specified. Please check the TrueStrings, FalseStrings and MissingStrings arguments."
    End If

    Exit Sub
ErrHandler:
    Throw "#AddKeyToDict: " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------------------------
' Procedure : sElapsedTime
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
'
'             See http://msdn.microsoft.com/en-us/library/windows/desktop/ms644904(v=vs.85).aspx
'---------------------------------------------------------------------------------------------------------
Private Function sElapsedTime() As Double
    Dim a As Currency
    Dim b As Currency
    On Error GoTo ErrHandler

    QueryPerformanceCounter a
    QueryPerformanceFrequency b
    sElapsedTime = a / b

    Exit Function
ErrHandler:
    Throw "#sElapsedTime: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestSentinelSpeed
' Purpose    : Test speed of accessing the sentinels dictionary, using similar approach to that employed in method
'              ConvertField.
'
' Results:  On Surface Book 2, Intel(R) Core(TM) i7-8650U CPU @ 1.90GHz   2.11 GHz, 16GB RAM
'
'Running TestSentinelSpeed 2021-08-25T15:01:33
'Conversions per second = 90,346,968       Field = "This string is longer than the longest sentinel, which is 14"
'Conversions per second = 20,976,150       Field = "mini" (Not a sentinel, but shorter than the longest sentinel)
'Conversions per second = 9,295,050        Field = "True" (A sentinel, one of the elements of TrueStrings)
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestSentinelSpeed()
    Dim Sentinels As New Scripting.Dictionary

    Dim Field As String
    Dim t1 As Double, t2 As Double
    Dim i As Long
    Dim j As Long
    Const N = 10000000
    Dim Res As Variant
    Dim MaxLength As Long
    Dim AnySentinels As Boolean
    Dim Comment As String

    MakeSentinels Sentinels, MaxLength, AnySentinels, _
        ShowBooleansAsBooleans:=True, _
        ShowErrorsAsErrors:=True, _
        ShowMissingsAs:=Empty, _
        TrueStrings:=Array("True", "T"), _
        FalseStrings:=Array("False", "F"), _
        MissingStrings:=Array("NA", "-999")
    
    Dim Converted As Boolean
    
    Debug.Print "Running TestSentinelSpeed " & Format(Now(), "yyyy-mm-ddThh:mm:ss")
    
    For j = 1 To 3
    
        Select Case j
            Case 1
                Field = "This string is longer than the longest sentinel, which is 14"
            Case 2
                Field = "mini"
                Comment = "Not a sentinel, but shorter than the longest sentinel"
            Case 3
                Field = "True"
                Comment = "A sentinel, one of the elements of TrueStrings"
        End Select

        t1 = sElapsedTime()
        For i = 1 To N
            If Len(Field) <= MaxLength Then
                If Sentinels.Exists(Field) Then
                    Res = Sentinels(Field)
                    Converted = True
                End If
            End If
        Next i
        t2 = sElapsedTime()

        Debug.Print "Conversions per second = " & Format(N / (t2 - t1), "###,###"), _
            "Field = """ & Field & """" & IIf(Comment = "", "", " (" & Comment & ")")

    Next j

    Exit Sub
ErrHandler:
    MsgBox "#TestSentinelSpeed: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SpeedTestISO8601
' Purpose    : Testing speed of CastISO8601

'Example output: (Surface Book 2, Intel(R) Core(TM) i7-8650U CPU @ 1.90GHz   2.11 GHz, 16GB RAM)

'Running SpeedTestISO8601 2021-08-24T18:59:30
'Calls per second = 2,650,189              strIn = "Foo" Check = 1900-01-00T00:00:00.000
'Calls per second = 2,660,280              strIn = "Foo" Check = 1900-01-00T00:00:00.000
'Calls per second = 2,681,348              strIn = "Foo" Check = 1900-01-00T00:00:00.000
'Calls per second = 2,229,503              strIn = "xxxxxxxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 2,144,496              strIn = "xxxxxxxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 2,135,019              strIn = "xxxxxxxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 1,581,522              strIn = "xxxx-xxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 1,574,575              strIn = "xxxx-xxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 1,594,879              strIn = "xxxx-xxxxxxx"      Check = 1900-01-00T00:00:00.000
'Calls per second = 632,374  strIn = "2021-08-24T15:18:01.123+05:0x"   Check = 1900-01-00T00:00:00.000
'Calls per second = 633,033  strIn = "2021-08-24T15:18:01.123+05:0x"   Check = 1900-01-00T00:00:00.000
'Calls per second = 643,524  strIn = "2021-08-24T15:18:01.123+05:0x"   Check = 1900-01-00T00:00:00.000
'Calls per second = 469,506  strIn = "2021-08-23"        Check = 2021-08-23T00:00:00.000
'Calls per second = 471,985  strIn = "2021-08-23"        Check = 2021-08-23T00:00:00.000
'Calls per second = 465,819  strIn = "2021-08-23"        Check = 2021-08-23T00:00:00.000
'Calls per second = 368,133  strIn = "2021-08-24T15:18:01"             Check = 2021-08-24T15:18:01.000
'Calls per second = 367,303  strIn = "2021-08-24T15:18:01"             Check = 2021-08-24T15:18:01.000
'Calls per second = 350,648  strIn = "2021-08-24T15:18:01"             Check = 2021-08-24T15:18:01.000
'Calls per second = 230,480  strIn = "2021-08-23T08:47:21.123"         Check = 2021-08-23T08:47:21.000
'Calls per second = 225,836  strIn = "2021-08-23T08:47:21.123"         Check = 2021-08-23T08:47:21.000
'Calls per second = 246,625  strIn = "2021-08-23T08:47:21.123"         Check = 2021-08-23T08:47:21.000
'Calls per second = 235,492  strIn = "2021-08-24T15:18:01+05:00"       Check = 2021-08-24T10:18:01.000
'Calls per second = 230,097  strIn = "2021-08-24T15:18:01+05:00"       Check = 2021-08-24T10:18:01.000
'Calls per second = 231,845  strIn = "2021-08-24T15:18:01+05:00"       Check = 2021-08-24T10:18:01.000
'Calls per second = 215,734  strIn = "2021-08-24T15:18:01.123+05:00"   Check = 2021-08-24T10:18:01.000
'Calls per second = 208,329  strIn = "2021-08-24T15:18:01.123+05:00"   Check = 2021-08-24T10:18:01.000
'Calls per second = 211,625  strIn = "2021-08-24T15:18:01.123+05:00"   Check = 2021-08-24T10:18:01.000
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SpeedTestISO8601()

    Const N = 1000000
    Dim Converted As Boolean
    Dim dtOut As Date
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strIn As String
    Dim SysDateOrder As Long
    Dim t1 As Double
    Dim t2 As Double

    SysDateOrder = Application.International(xlDateOrder)

    Debug.Print "Running SpeedTestISO8601 " & Format(Now(), "yyyy-mm-ddThh:mm:ss")
    For k = 1 To 9
        For j = 1 To 3
            dtOut = 0
            Select Case k
                Case 1
                    strIn = "Foo" ' less than 10 in length
                Case 2
                    strIn = "xxxxxxxxxxxx" '5th character not "-"
                Case 3
                    strIn = "xxxx-xxxxxxx" 'rejected by RegEx
                Case 4
                    strIn = "2021-08-24T15:18:01.123+05:0x" ' rejected by regex
                Case 5
                    strIn = "2021-08-23"
                Case 6
                    strIn = "2021-08-24T15:18:01"
                Case 7
                    strIn = "2021-08-23T08:47:21.123"
                Case 8
                    strIn = "2021-08-24T15:18:01+05:00"
                Case 9
                    strIn = "2021-08-24T15:18:01.123+05:00"
            End Select

            t1 = sElapsedTime()
            For i = 1 To N
                Call CastISO8601(strIn, dtOut, Converted, True, True)
            Next i
            t2 = sElapsedTime
            Debug.Print "Calls per second = " & Format(N / (t2 - t1), "###,###"), "strIn = """ & strIn & """", "Check = " & Application.WorksheetFunction.text(dtOut, "yyyy-mm-ddThh:mm:ss.000")
            DoEvents 'kick Immediate window to life
        Next j
    Next k

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseISO8601
' Purpose    : Test harness for calling from spreadsheets
' -----------------------------------------------------------------------------------------------------------------------
Public Function ParseISO8601(strIn As String)
    Dim dtOut As Date
    Dim Converted As Boolean

    On Error GoTo ErrHandler
    CastISO8601 strIn, dtOut, Converted, True, True

    If Converted Then
        ParseISO8601 = dtOut
    Else
        ParseISO8601 = "#Not recognised as ISO8601 date!"
    End If
    Exit Function
ErrHandler:
    ParseISO8601 = "#ParseISO8601: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToTime
' Purpose    : Cast strings that represent a time to a date, no handling of TimeZone.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToTime(strIn As String, ByRef dtOut As Date, ByRef Converted As Boolean)
    Static rx As VBScript_RegExp_55.RegExp

    On Error GoTo ErrHandler
    
    dtOut = CDate(strIn)
    If dtOut <= 1 Then
        Converted = True
    End If
    
    Exit Sub
ErrHandler:
    'Do nothing, was not a valid time (e.g. h,m or s out of range)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToTimeB
' Purpose    : CDate does not correctly cope with times such as '04:20:10.123 am' or '04:20:10.123', i.e, times with a
'              fractional second, so this method is called after CastToTime
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToTimeB(strIn As String, ByRef dtOut As Date, ByRef Converted As Boolean)
    Static rx As VBScript_RegExp_55.RegExp
    Dim L As Long
    Dim DecPointAt As Long
    Dim SpaceAt
    Dim FractionalSecond As Double
    
    On Error GoTo ErrHandler
    If rx Is Nothing Then
        Set rx = New RegExp
        With rx
            .IgnoreCase = True
            .Pattern = "^[0-2]?[0-9]:[0-5]?[0-9]:[0-5]?[0-9](\.[0-9]+)( am| pm)?$"
            .Global = False        'Find first match only
        End With
    End If

    If Not rx.Test(strIn) Then Exit Sub
    DecPointAt = InStr(strIn, ".")
    If DecPointAt = 0 Then Exit Sub ' should never happen
    SpaceAt = InStr(strIn, " ")
    If SpaceAt = 0 Then SpaceAt = Len(strIn) + 1
    FractionalSecond = CDbl(Mid$(strIn, DecPointAt, SpaceAt - DecPointAt)) / 86400
    
    dtOut = CDate(Left$(strIn, DecPointAt - 1) + Mid(strIn, SpaceAt)) + FractionalSecond
    Converted = True
    Exit Sub
ErrHandler:
    'Do nothing, was not a valid time (e.g. h,m or s out of range)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastISO8601
' Purpose    : Convert ISO8601 formatted datestrings to UTC date. Handles the following formats. https://xkcd.com/1179/

'Always accepts dates without time
'Format                        Example
'yyyy-mm-dd                    2021-08-23

'If AcceptWithoutTimeZone is True:
'yyyy-mm-ddThh:mm:ss           2021-08-23T08:47:21
'yyyy-mm-ddThh:mm:ss.000       2021-08-23T08:47:20.920

'If AcceptWithTimeZone is True:
'yyyy-mm-ddThh:mm:ssZ          2021-08-23T08:47:21Z
'yyyy-mm-ddThh:mm:ss.000Z      2021-08-23T08:47:20.920Z
'yyyy-mm-ddThh:mm:ss+hh:mm     2021-08-23T08:47:21+05:00
'yyyy-mm-ddThh:mm:ss.000+hh:mm 2021-08-23T08:47:20.920+05:00

' Parameters :
'  StrIn                : The string to be converted
'  DtOut                : The date that the string converts to.
'  Converted            : Did the function convert (true) or reject as not a correctly formatted date (false)
'  AcceptWithoutTimeZone: Should the function accept datetime without time zone given?
'  AcceptWithTimeZone   : Should the function accept datetime with time zone given?

'       IMPORTANT:       WHEN TIMEZONE IS GIVEN THE FUNCTION RETURNS THE TIME IN UTC

' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastISO8601(ByVal strIn As String, dtOut As Date, ByRef Converted As Boolean, _
    AcceptWithoutTimeZone As Boolean, AcceptWithTimeZone As Boolean)

    Dim L As Long
    Dim LocalTime As Double
    Dim MilliPart As Double
    Dim MinusPos As Long
    Dim PlusPos As Long
    Dim Sign
    Dim ZAtEnd As Boolean
    
    Static rxNoNo As VBScript_RegExp_55.RegExp
    Static RxYesNo As VBScript_RegExp_55.RegExp
    Static RxNoYes As VBScript_RegExp_55.RegExp
    Static rxYesYes As VBScript_RegExp_55.RegExp
    Static rxExists As Boolean

    On Error GoTo ErrHandler
     
    If Not rxExists Then
        Set rxNoNo = New RegExp
        'Reject datetime
        With rxNoNo
            .IgnoreCase = False
            .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9]$"
            .Global = False
        End With
        
        'Accept datetime without time zone, reject datetime with timezone
        Set RxYesNo = New RegExp
        With RxYesNo
            .IgnoreCase = False
            .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?)?$"
            .Global = False
        End With
        
        'Reject datetime without time zone, accept datetime with timezone
        Set RxNoYes = New RegExp
        With RxNoYes
            .IgnoreCase = False
            .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?(Z|((\+|\-)[0-2][0-9]:[0-5][0-9])))?$"
            .Global = False
        End With
        
        'Accept datetime, both with and without timezone
        Set rxYesYes = New RegExp
        With rxYesYes
            .IgnoreCase = False
            .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?((Z|((\+|\-)[0-2][0-9]:[0-5][0-9])))?)?$"
            .Global = False
        End With
        rxExists = True
    End If

    Converted = False

    L = Len(strIn)

    If L < 10 Then Exit Sub
    If Mid$(strIn, 5, 1) <> "-" Then Exit Sub
    
    If AcceptWithoutTimeZone Then
        If AcceptWithTimeZone Then
            If Not rxYesYes.Test(strIn) Then Exit Sub
        Else
            If Not RxYesNo.Test(strIn) Then Exit Sub
        End If
    Else
        If AcceptWithTimeZone Then
            If Not RxNoYes.Test(strIn) Then Exit Sub
        Else
            If Not rxNoNo.Test(strIn) Then Exit Sub
        End If
    End If

    If L = 10 Then
        'This works irrespective of Windows regional settings
        dtOut = CDate(strIn)
        Converted = True
        Exit Sub
    End If
    
    'Replace the "T" separator
    Mid$(strIn, 11, 1) = " "
    
    If L = 19 Then
        dtOut = CDate(strIn)
        Converted = True
        Exit Sub
    End If

    If Right$(strIn, 1) = "Z" Then
        Sign = 0
        ZAtEnd = True
    Else
        PlusPos = InStr(20, strIn, "+")
        If PlusPos > 0 Then
            Sign = 1
        Else
            MinusPos = InStr(20, strIn, "-")
            If MinusPos > 0 Then
                Sign = -1
            End If
        End If
    End If

    If Mid$(strIn, 20, 1) = "." Then 'Have fraction of a second
        Select Case Sign
            Case 0
                'Example: "2021-08-23T08:47:20.920Z"
                MilliPart = CDbl(Mid$(strIn, 20, IIf(ZAtEnd, L - 20, L - 19)))
            Case 1
                'Example: "2021-08-23T08:47:20.920+05:00"
                MilliPart = CDbl(Mid$(strIn, 20, PlusPos - 20))
            Case -1
                'Example: "2021-08-23T08:47:20.920-05:00"
                MilliPart = CDbl(Mid$(strIn, 20, MinusPos - 20))
        End Select
    End If
    
    LocalTime = CDate(Left$(strIn, 19)) + MilliPart / 86400

    Dim Adjust As Date
    Select Case Sign
        Case 0
            dtOut = LocalTime
            Converted = True
            Exit Sub
        Case 1
            If L <> PlusPos + 5 Then Exit Sub
            Adjust = CDate(Right$(strIn, 5))
            dtOut = LocalTime - Adjust
            Converted = True
        Case -1
            If L <> MinusPos + 5 Then Exit Sub
            Adjust = CDate(Right$(strIn, 5))
            dtOut = LocalTime + Adjust
            Converted = True
    End Select

    Exit Sub
ErrHandler:
    'Was not recognised as ISO8601 date
End Sub

'See "gogeek"'s post at https://stackoverflow.com/questions/1600875/how-to-get-the-current-datetime-in-utc-from-an-excel-vba-macro
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetLocalOffsetToUTC
' Purpose    : Get the PC's offset to UTC.
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetLocalOffsetToUTC()
    Dim dt As Object, UTC As Date
    Dim TimeNow As Date
    On Error GoTo ErrHandler
    TimeNow = Now()

    Set dt = CreateObject("WbemScripting.SWbemDateTime")
    dt.SetVarDate TimeNow
    UTC = dt.GetVarDate(False)
    GetLocalOffsetToUTC = (TimeNow - UTC)

    Exit Function
ErrHandler:
    Throw "#GetLocalOffsetToUTC: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISOZFormatString
' Purpose    : Returns the format string required to save datetimes with timezone under the assumton that the offset to
'              UTC is the same as the curent offset on this PC - use with care, Daylight saving may mean that that's not
'              a correct assumption for all the dates in a set of data...
' -----------------------------------------------------------------------------------------------------------------------
Private Function ISOZFormatString()
    Dim TimeZone
    Dim RightChars

    On Error GoTo ErrHandler
    TimeZone = GetLocalOffsetToUTC()

    If TimeZone = 0 Then
        RightChars = "Z"
    ElseIf TimeZone > 0 Then
        RightChars = "+" & Format(TimeZone, "hh:mm")
    Else
        RightChars = "-" & Format(Abs(TimeZone), "hh:mm")
    End If
    ISOZFormatString = "yyyy-mm-ddT:hh:mm:ss" & RightChars

    Exit Function
ErrHandler:
    Throw "#ISOZFormatString: " & Err.Description & "!"
End Function

