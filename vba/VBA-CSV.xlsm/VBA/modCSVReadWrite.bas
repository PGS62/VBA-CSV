Attribute VB_Name = "modCSVReadWrite"
' VBA-CSV

' Copyright (C) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme
' This version at: https://github.com/PGS62/VBA-CSV/releases/tag/v0.30

'Installation:
'1) Import this module into your project (Open VBA Editor, Alt + F11; File > Import File).

'2) Add three references (In VBA Editor Tools > References)
'   Microsoft Scripting Runtime
'   Microsoft VBScript Regular Expressions 5.5 (or a later version if available)
'   Microsoft Active X Data Objects 6.1 Library (or later version if available)

'3) If you plan to call the functions from spreadsheet formulas then you might like to tell
'   Excel's Function Wizard about them by adding calls to RegisterCSVRead and RegisterCSVWrite
'   to the project's Workbook_Open event, which lives in the ThisWorkbook class module.

'        Private Sub Workbook_Open()
'            RegisterCSVRead
'            RegisterCSVWrite
'        End Sub

'4) An alternative (or additional) approach to providing help on CSVRead and CSVWrite is:
'   a) Install Excel-DNA Intellisense. See https://github.com/Excel-DNA/IntelliSense#getting-started
'   b) Copy the worksheet _Intellisense_ from
'      https://github.com/PGS62/VBA-CSV/releases/download/v0.30/VBA-CSV-Intellisense.xlsx
'      into the workbook that contains this VBA code.

'5) If you envisage calling CSVRead and CSVWrite only from VBA code and not from worksheet formulas
'   then consider changing constant m_ErrorStyle to be es_RaiseError.
'   https://github.com/PGS62/VBA-CSV#errors

'6) If you would prefer the arrays returned by CSVRead to be zero-based rather than one-based
'   then change constant m_LBound to 0.

Option Explicit

Private Enum enmErrorStyle
    es_ReturnString = 0
    es_RaiseError = 1
End Enum

Private Const m_ErrorStyle As Long = es_ReturnString

Private Const m_LBound As Long = 1

'Set to True when debugging for more informative error description (i.e. with call stack).
Private Const m_ErrorStringsEmbedCallStack As Boolean = False

Private m_FSO As Scripting.FileSystemObject
Private Const DQ = """"
Private Const DQ2 = """"""

'The two constants below set the range of years to which date strings with two-digit years are converted.
Private Const m_2DigitYearIsFrom As Long = 1930
Private Const m_2DigitYearIsTo As Long = 2029

#If VBA7 And Win64 Then
'for 64-bit Excel
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As Long
#Else
'for 32-bit Excel
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Private Enum enmSourceType
    st_File = 0
    st_URL = 1
    st_String = 2
End Enum

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CSVRead
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path, or else a URL of a file, or else a string in
'             CSV format.
' ConvertTypes: Controls whether fields in the file are converted to typed values or remain as strings, and
'             sets the treatment of "quoted fields" and space characters.
'
'             ConvertTypes should be a string of zero or more letters from allowed characters `NDBETQK`.
'
'             The most commonly useful letters are:
'             1) `N` number fields are returned as numbers (Doubles).
'             2) `D` date fields (that respect DateFormat) are returned as Dates.
'             3) `B` fields matching TrueStrings or FalseStrings are returned as Booleans.
'
'             ConvertTypes is optional and defaults to the null string for no type conversion. `TRUE` is
'             equivalent to `NDB` and `FALSE` to the null string.
'
'             Four further options are available:
'             4) `E` fields that match Excel errors are converted to error values. There are fourteen of
'             these, including `#N/A`, `#NAME?`, `#VALUE!` and `#DIV/0!`.
'             5) `T` leading and trailing spaces are trimmed from fields. In the case of quoted fields,
'             this will not remove spaces between the quotes.
'             6) `Q` conversion happens for both quoted and unquoted fields; otherwise only unquoted fields
'             are converted.
'             7) `K` quoted fields are returned with their quotes kept in place.
'
'             For most files, correct type conversion can be achieved with ConvertTypes as a string which
'             applies for all columns, but type conversion can also be specified on a per-column basis.
'
'             Enter an array (or range) with two columns or two rows, column numbers on the left/top and
'             type conversion (subset of `NDBETQK`) on the right/bottom. Instead of column numbers, you can
'             enter strings matching the contents of the header row, and a column number of zero applies to
'             all columns not otherwise referenced.
'
'             For convenience when calling from VBA, you can pass an array of two element arrays such as
'             `Array(Array(0,"N"),Array(3,""),Array("Phone",""))` to convert all numbers in a file into
'             numbers in the return except for those in column 3 and in the column(s) headed "Phone".
' Delimiter : By default, CSVRead will try to detect a file's delimiter as the first instance of comma, tab,
'             semi-colon, colon or pipe found in the first 10,000 characters of the file, searching only
'             outside of quoted regions and outside of date-with-time fields (since these contain colons).
'             If it can't auto-detect the delimiter, it will assume comma. If your file includes a
'             different character or string delimiter you should pass that as the Delimiter argument.
'
'             Alternatively, enter `FALSE` as the delimiter to treat the file as "not a delimited file". In
'             this case the return will mimic how the file would appear in a text editor such as NotePad.
'             The file will be split into lines at all line breaks (irrespective of double quotes) and each
'             element of the return will be a line of the file.
' IgnoreRepeated: Whether delimiters which appear at the start of a line, the end of a line or immediately
'             after another delimiter should be ignored while parsing; useful for fixed-width files with
'             delimiter padding between fields.
' DateFormat: The format of dates in the file such as `Y-M-D` (the default), `M-D-Y` or `Y/M/D`. Also `ISO`
'             for ISO8601 (e.g., 2021-08-26T09:11:30) or `ISOZ` (time zone given e.g.
'             2021-08-26T13:11:30+05:00), in which case dates-with-time are returned in UTC time.
' Comment   : Rows that start with this string will be skipped while parsing.
' IgnoreEmptyLines: Whether empty rows/lines in the file should be skipped while parsing (if `FALSE`, each
'             column will be assigned ShowMissingsAs for that empty row).
' HeaderRowNum: The row in the file containing headers. Type conversion is not applied to fields in the
'             header row, though leading and trailing spaces are trimmed.
'
'             This argument is most useful when calling from VBA, with SkipToRow set to one more than
'             HeaderRowNum. In that case the function returns the rows starting from SkipToRow, and the
'             header row is returned via the by-reference argument HeaderRow. Optional and defaults to 0.
' SkipToRow : The first row in the file that's included in the return. Optional and defaults to one more
'             than HeaderRowNum.
' SkipToCol : The column in the file at which reading starts, as a number or a string matching one of the
'             file's headers. Optional and defaults to 1 to read from the first column.
' NumRows   : The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the
'             end of the file are read.
' NumCols   : If a number, sets the number of columns to read from the file. If a string matching one of the
'             file's headers, sets the last column to be read. If omitted (or zero), all columns from
'             SkipToCol are read.
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
' Encoding  : Allowed entries are `ASCII`, `ANSI`, `UTF-8`, or `UTF-16`. For most files this argument can be
'             omitted and CSVRead will detect the file's encoding. If auto-detection does not work, then
'             it's possible that the file is encoded `UTF-8` or `UTF-16` but without a byte order mark to
'             identify the encoding. Experiment with Encoding as each of `UTF-8` and `UTF-16`.
'
'             `ANSI` is taken to mean `Windows-1252` encoding.
' DecimalSeparator: In many places in the world, floating point number decimals are separated with a comma
'             instead of a period (3,14 vs. 3.14). CSVRead can correctly parse these numbers by passing in
'             the DecimalSeparator as a comma, in which case comma ceases to be a candidate if the parser
'             needs to guess the Delimiter.
' HeaderRow : This by-reference argument is for use from VBA (as opposed to from Excel formulas). It is
'             populated with the contents of the header row, with no type conversion, though leading and
'             trailing spaces are removed.
'
' Notes     : See also companion function CSVWrite.
'
'             The function handles all csv files that conform to the standards described in RFC4180
'             https://www.rfc-editor.org/rfc/rfc4180.txt including files with quoted fields.
'
'             In addition the function handles files which break some of those standards:
'             * Not all lines of the file need have the same number of fields. The function "pads" with
'             ShowMissingsAs values.
'             * Fields which start with a double quote but do not end with a double quote are handled by
'             being returned unchanged. Necessarily such fields have an even number of double quotes, or
'             otherwise the field will be treated as the last field in the file.
'             * The standard states that csv files should have Windows-style line endings, but the function
'             supports files with Windows, Unix and (old) Mac line endings. Files may also have mixed line
'             endings.
' -----------------------------------------------------------------------------------------------------------------------
Public Function CSVRead(ByVal FileName As String, Optional ByVal ConvertTypes As Variant = False, _
          Optional ByVal Delimiter As Variant, Optional ByVal IgnoreRepeated As Boolean, _
          Optional ByVal DateFormat As String = "Y-M-D", Optional ByVal Comment As String, _
          Optional ByVal IgnoreEmptyLines As Boolean, Optional ByVal HeaderRowNum As Long, _
          Optional ByVal SkipToRow As Long, Optional ByVal SkipToCol As Variant = 1, _
          Optional ByVal NumRows As Long, Optional ByVal NumCols As Variant = 0, _
          Optional ByVal TrueStrings As Variant, Optional ByVal FalseStrings As Variant, _
          Optional ByVal MissingStrings As Variant, Optional ByVal ShowMissingsAs As Variant, _
          Optional ByVal Encoding As Variant, Optional ByVal DecimalSeparator As String, _
          Optional ByRef HeaderRow As Variant) As Variant
Attribute CSVRead.VB_Description = "Returns the contents of a comma-separated file on disk as an array."
Attribute CSVRead.VB_ProcData.VB_Invoke_Func = " \n14"

          Const Err_Delimiter As String = "Delimiter character must be passed as a string, FALSE for no delimiter. " & _
              "Omit to guess from file contents"
          Const Err_Delimiter2 As String = "Delimiter must have at least one character and cannot start with a double " & _
              "quote, line feed or carriage return"
          Const Err_FileEmpty As String = "File is empty"
          Const Err_FunctionWizard  As String = "Disabled in Function Wizard"
          Const Err_NumRows As String = "NumRows must be positive to read a given number of rows, or zero or omitted to " & _
              "read all rows from SkipToRow to the end of the file"
          Const Err_Seps1 As String = "DecimalSeparator must be a single character"
          Const Err_Seps2 As String = "DecimalSeparator must not be equal to the first character of Delimiter or to a " & _
              "line-feed or carriage-return"
          Const Err_SkipToRow As String = "SkipToRow must be at least 1"
          Const Err_Comment As String = "Comment must not contain double-quote, line feed or carriage return"
          Const Err_HeaderRowNum As String = "HeaderRowNum must be greater than or equal to zero and less than or equal to SkipToRow"
          
          Dim AcceptWithoutTimeZone As Boolean
          Dim AcceptWithTimeZone As Boolean
          Dim Adj As Long
          Dim AnyConversion As Boolean
          Dim AnySentinels As Boolean
          Dim AscSeparator As Long
          Dim CallingFromWorksheet As Boolean
          Dim CharSet As String
          Dim ColByColFormatting As Boolean
          Dim ColIndexes() As Long
          Dim ConvertQuoted As Boolean
          Dim CSVContents As String
          Dim CTDict As Scripting.Dictionary
          Dim DateOrder As Long
          Dim DateSeparator As String
          Dim Err_StringTooLong As String
          Dim EstNumChars As Long
          Dim FileSize As Long
          Dim i As Long
          Dim ISO8601 As Boolean
          Dim j As Long
          Dim k As Long
          Dim KeepQuotes As Boolean
          Dim Lengths() As Long
          Dim MaxSentinelLength As Long
          Dim MSLIA As Long
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
          Dim Sentinels As Scripting.Dictionary
          Dim SepStandard As Boolean
          Dim ShowBooleansAsBooleans As Boolean
          Dim ShowDatesAsDates As Boolean
          Dim ShowErrorsAsErrors As Boolean
          Dim ShowMissingsAsEmpty As Boolean
          Dim ShowNumbersAsNumbers As Boolean
          Dim SourceType As enmSourceType
          Dim Starts() As Long
          Dim strDelimiter As String
          Dim Stream As ADODB.Stream
          Dim SysDateSeparator As String
          Dim SysDecimalSeparator As String
          Dim TempFile As String
          Dim TrimFields As Boolean
          
1         On Error GoTo ErrHandler

2         SourceType = InferSourceType(FileName)

          'Download file from internet to local temp folder
3         If SourceType = st_URL Then
4             TempFile = Environ$("Temp") & "\VBA-CSV\Downloads\DownloadedFile.csv"
5             FileName = Download(FileName, TempFile)
6             SourceType = st_File
7         End If

          'Parse and validate inputs...
8         If SourceType <> st_String Then
9             FileSize = GetFileSize(FileName)
10            ParseEncoding FileName, Encoding, CharSet
11            EstNumChars = EstimateNumChars(FileSize, CStr(Encoding))
12        End If

13        If VarType(Delimiter) = vbBoolean Then
14            If Not Delimiter Then
15                NotDelimited = True
16            Else
17                Throw Err_Delimiter
18            End If
19        ElseIf VarType(Delimiter) = vbString Then
20            If Len(Delimiter) = 0 Then
21                strDelimiter = InferDelimiter(SourceType, FileName, DecimalSeparator, CharSet)
22            ElseIf Left$(Delimiter, 1) = DQ Or Left$(Delimiter, 1) = vbLf Or Left$(Delimiter, 1) = vbCr Then
23                Throw Err_Delimiter2
24            Else
25                strDelimiter = Delimiter
26            End If
27        ElseIf IsEmpty(Delimiter) Or IsMissing(Delimiter) Then
28            strDelimiter = InferDelimiter(SourceType, FileName, DecimalSeparator, CharSet)
29        Else
30            Throw Err_Delimiter
31        End If

32        SysDecimalSeparator = Application.DecimalSeparator
33        If DecimalSeparator = vbNullString Then DecimalSeparator = SysDecimalSeparator

34        If DecimalSeparator = SysDecimalSeparator Then
35            SepStandard = True
36        ElseIf Len(DecimalSeparator) <> 1 Then
37            Throw Err_Seps1
38        ElseIf DecimalSeparator = strDelimiter Or DecimalSeparator = vbLf Or DecimalSeparator = vbCr Then
39            Throw Err_Seps2
40        End If

41        Set CTDict = New Scripting.Dictionary

42        ParseConvertTypes ConvertTypes, ShowNumbersAsNumbers, _
              ShowDatesAsDates, ShowBooleansAsBooleans, ShowErrorsAsErrors, _
              ConvertQuoted, KeepQuotes, TrimFields, ColByColFormatting, HeaderRowNum, CTDict

43        Set Sentinels = New Scripting.Dictionary
44        MakeSentinels Sentinels, ConvertQuoted, strDelimiter, MaxSentinelLength, AnySentinels, ShowBooleansAsBooleans, _
              ShowErrorsAsErrors, ShowMissingsAs, TrueStrings, FalseStrings, MissingStrings
          
45        If ShowDatesAsDates Then
46            ParseDateFormat DateFormat, DateOrder, DateSeparator, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone
47            SysDateSeparator = Application.International(xlDateSeparator)
48        End If

49        If HeaderRowNum < 0 Then Throw Err_HeaderRowNum
50        If SkipToRow = 0 Then SkipToRow = HeaderRowNum + 1
51        If HeaderRowNum > SkipToRow Then Throw Err_HeaderRowNum

52        If Not IsOneAndZero(SkipToCol, NumCols) Then
53            AmendSkipToColAndNumCols FileName, SkipToCol, NumCols, Delimiter, IgnoreRepeated, Comment, IgnoreEmptyLines, HeaderRowNum
54        End If

55        If SkipToCol = 0 Then SkipToCol = 1
56        If SkipToRow < 1 Then Throw Err_SkipToRow
57        If NumRows < 0 Then Throw Err_NumRows

58        If HeaderRowNum > SkipToRow Then Throw Err_HeaderRowNum
             
59        If InStr(Comment, DQ) > 0 Or InStr(Comment, vbLf) > 0 Or InStr(Comment, vbCrLf) > 0 Then Throw Err_Comment
          'End of input validation
          
60        CallingFromWorksheet = TypeName(Application.Caller) = "Range"
          
61        If CallingFromWorksheet Then
62            If FunctionWizardActive() Then
63                CSVRead = "#" & Err_FunctionWizard & "!"
64                Exit Function
65            End If
66        End If
          
67        If NotDelimited Then
68            HeaderRow = Empty
69            CSVRead = ParseTextFile(FileName, SourceType <> st_String, CharSet, SkipToRow, NumRows, CallingFromWorksheet)
70            Exit Function
71        End If
                
72        If SourceType = st_String Then
73            CSVContents = FileName
74            ParseCSVContents CSVContents, DQ, strDelimiter, Comment, IgnoreEmptyLines, _
                  IgnoreRepeated, SkipToRow, HeaderRowNum, NumRows, NumRowsFound, NumColsFound, _
                  NumFields, Ragged, Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts, HeaderRow
75        Else
76            If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
                  
77            Set Stream = New ADODB.Stream
78            Stream.CharSet = CharSet
79            Stream.Open
80            Stream.LoadFromFile FileName
81            If Stream.EOS Then Throw Err_FileEmpty

82            If SkipToRow = 1 And NumRows = 0 Then
83                CSVContents = ReadAllFromStream(Stream, EstNumChars)
84                Stream.Close
85                ParseCSVContents CSVContents, DQ, strDelimiter, Comment, IgnoreEmptyLines, _
                      IgnoreRepeated, SkipToRow, HeaderRowNum, NumRows, NumRowsFound, NumColsFound, NumFields, _
                      Ragged, Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts, HeaderRow
86            Else
87                CSVContents = ParseCSVContents(Stream, DQ, strDelimiter, Comment, IgnoreEmptyLines, _
                      IgnoreRepeated, SkipToRow, HeaderRowNum, NumRows, NumRowsFound, NumColsFound, NumFields, _
                      Ragged, Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts, HeaderRow)
88                Stream.Close
89            End If
90        End If
                           
91        If NumCols = 0 Then
92            NumColsInReturn = NumColsFound - SkipToCol + 1
93            If NumColsInReturn <= 0 Then
94                Throw "SkipToCol (" & CStr(SkipToCol) & _
                      ") exceeds the number of columns in the file (" & CStr(NumColsFound) & ")"
95            End If
96        Else
97            NumColsInReturn = NumCols
98        End If
99        If NumRows = 0 Then
100           NumRowsInReturn = NumRowsFound
101       Else
102           NumRowsInReturn = NumRows
103       End If
              
104       AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
              ShowBooleansAsBooleans Or ShowErrorsAsErrors Or TrimFields
              
105       Adj = m_LBound - 1
106       ReDim ReturnArray(1 + Adj To NumRowsInReturn + Adj, 1 + Adj To NumColsInReturn + Adj)
107       MSLIA = MaxStringLengthInArray()
108       ShowMissingsAsEmpty = IsEmpty(ShowMissingsAs)
              
109       For k = 1 To NumFields
110           i = RowIndexes(k)
111           j = ColIndexes(k) - SkipToCol + 1
112           If j >= 1 And j <= NumColsInReturn Then
113               If CallingFromWorksheet Then
114                   If Lengths(k) > MSLIA Then
                          Dim UnquotedLength As Long
115                       UnquotedLength = Len(Unquote(Mid$(CSVContents, Starts(k), Lengths(k)), DQ, 4, KeepQuotes))
116                       If UnquotedLength > MSLIA Then
117                           Err_StringTooLong = "The file has a field (row " & CStr(i + SkipToRow - 1) & _
                                  ", column " & CStr(j + SkipToCol - 1) & ") of length " & Format$(UnquotedLength, "###,###")
118                           If MSLIA >= 32767 Then
119                               Err_StringTooLong = Err_StringTooLong & ". Excel cells cannot contain strings longer than " & Format$(MSLIA, "####,####")
120                           Else 'Excel 2013 and earlier
121                               Err_StringTooLong = Err_StringTooLong & _
                                      ". An array containing a string longer than " & Format$(MSLIA, "###,###") & _
                                      " cannot be returned from VBA to an Excel worksheet"
122                           End If
123                           Throw Err_StringTooLong
124                       End If
125                   End If
126               End If
              
127               If ColByColFormatting Then
128                   ReturnArray(i + Adj, j + Adj) = Mid$(CSVContents, Starts(k), Lengths(k))
129               Else
130                   ReturnArray(i + Adj, j + Adj) = ConvertField(Mid$(CSVContents, Starts(k), Lengths(k)), AnyConversion, _
                          Lengths(k), TrimFields, DQ, QuoteCounts(k), ConvertQuoted, KeepQuotes, ShowNumbersAsNumbers, SepStandard, _
                          DecimalSeparator, SysDecimalSeparator, ShowDatesAsDates, ISO8601, AcceptWithoutTimeZone, _
                          AcceptWithTimeZone, DateOrder, DateSeparator, SysDateSeparator, AnySentinels, _
                          Sentinels, MaxSentinelLength, ShowMissingsAs, AscSeparator)
131               End If
132           End If
133       Next k
          
134       If Ragged Then
135           If Not ShowMissingsAsEmpty Then
136               For i = 1 + Adj To NumRowsInReturn + Adj
137                   For j = 1 + Adj To NumColsInReturn + Adj
138                       If IsEmpty(ReturnArray(i, j)) Then
139                           ReturnArray(i, j) = ShowMissingsAs
140                       End If
141                   Next j
142               Next i
143           End If
144           If Not IsEmpty(HeaderRow) Then
145               If NCols(HeaderRow) < NCols(ReturnArray) + SkipToCol - 1 Then
146                   ReDim Preserve HeaderRow(1 To 1, 1 To NCols(ReturnArray) + SkipToCol - 1)
147               End If
148           End If
149       End If

150       If SkipToCol > 1 Then
151           If Not IsEmpty(HeaderRow) Then
                  Dim HeaderRowTruncated() As String
152               ReDim HeaderRowTruncated(1 To 1, 1 To NumColsInReturn)
153               For i = 1 To NumColsInReturn
154                   HeaderRowTruncated(1, i) = HeaderRow(1, i + SkipToCol - 1)
155               Next i
156               HeaderRow = HeaderRowTruncated
157           End If
158       End If
          
          'In this case no type conversion should be applied to the top row of the return
159       If HeaderRowNum = SkipToRow Then
160           If AnyConversion Then
161               For i = 1 To MinLngs(NCols(HeaderRow), NumColsInReturn)
162                   ReturnArray(1 + Adj, i + Adj) = HeaderRow(1, i)
163               Next
164           End If
165       End If

166       If ColByColFormatting Then
              Dim CT As Variant
              Dim Field As String
              Dim NC As Long
              Dim NCH As Long
              Dim NR As Long
              Dim QC As Long
              Dim UnQuotedHeader As String
167           NR = NRows(ReturnArray)
168           NC = NCols(ReturnArray)
169           If IsEmpty(HeaderRow) Then
170               NCH = 0
171           Else
172               NCH = NCols(HeaderRow) 'possible that headers has fewer than expected columns if file is ragged
173           End If

174           For j = 1 To NC
175               If j + SkipToCol - 1 <= NCH Then
176                   UnQuotedHeader = HeaderRow(1, j + SkipToCol - 1)
177               Else
178                   UnQuotedHeader = -1 'Guaranteed not to be a key of the Dictionary
179               End If
180               If CTDict.Exists(j + SkipToCol - 1) Then
181                   CT = CTDict.item(j + SkipToCol - 1)
182               ElseIf CTDict.Exists(UnQuotedHeader) Then
183                   CT = CTDict.item(UnQuotedHeader)
184               ElseIf CTDict.Exists(0) Then
185                   CT = CTDict.item(0)
186               Else
187                   CT = vbNullString
188               End If
                  
189               If VarType(CT) = vbBoolean Then CT = StandardiseCT(CT)
                  
190               ParseCTString CT, ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, _
                      ShowErrorsAsErrors, ConvertQuoted, KeepQuotes, TrimFields
                  
191               AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
                      ShowBooleansAsBooleans Or ShowErrorsAsErrors
                      
192               Set Sentinels = New Scripting.Dictionary
                  
193               MakeSentinels Sentinels, ConvertQuoted, strDelimiter, MaxSentinelLength, AnySentinels, ShowBooleansAsBooleans, _
                      ShowErrorsAsErrors, ShowMissingsAs, TrueStrings, FalseStrings, MissingStrings

194               For i = 1 To NR
195                   If Not IsEmpty(ReturnArray(i + Adj, j + Adj)) Then
196                       Field = CStr(ReturnArray(i + Adj, j + Adj))
197                       QC = CountQuotes(Field, DQ)
198                       ReturnArray(i + Adj, j + Adj) = ConvertField(Field, AnyConversion, _
                              Len(ReturnArray(i + Adj, j + Adj)), TrimFields, DQ, QC, ConvertQuoted, KeepQuotes, _
                              ShowNumbersAsNumbers, SepStandard, DecimalSeparator, SysDecimalSeparator, _
                              ShowDatesAsDates, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone, DateOrder, _
                              DateSeparator, SysDateSeparator, AnySentinels, Sentinels, _
                              MaxSentinelLength, ShowMissingsAs, AscSeparator)
199                   End If
200               Next i
201           Next j
202       End If

203       CSVRead = ReturnArray

204       Exit Function

ErrHandler:
205       CSVRead = ReThrow("CSVRead", Err, m_ErrorStyle = es_ReturnString)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferSourceType
' Purpose    : Guess whether FileName is in fact a file, a URL or a string in CSV format
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferSourceType(FileName As String) As enmSourceType

1         On Error GoTo ErrHandler
2         If InStr(FileName, vbLf) > 0 Then 'vbLf and vbCr are not permitted characters in file names or urls
3             InferSourceType = st_String
4         ElseIf InStr(FileName, vbCr) > 0 Then
5             InferSourceType = st_String
6         ElseIf Mid$(FileName, 2, 2) = ":\" Then
7             InferSourceType = st_File
8         ElseIf Left$(FileName, 2) = "\\" Then
9             InferSourceType = st_File
10        ElseIf Left$(FileName, 8) = "https://" Then
11            InferSourceType = st_URL
12        ElseIf Left$(FileName, 7) = "http://" Then
13            InferSourceType = st_URL
14        Else
              'Doesn't look like either file with path, url or string in CSV format
15            InferSourceType = st_String
16            If Len(FileName) < 1000 Then
17                If FileExists(FileName) Then 'file exists in current working directory
18                    InferSourceType = st_File
19                End If
20            End If
21        End If

22        Exit Function
ErrHandler:
23        ReThrow "InferSourceType", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReadAllFromStream
' Purpose    : Read an entire Stream into a string. Replaces use of ADODB.ReadText(-1) since that has _very_ poor
'              performance for large files. The solution is to read the stream in chunks. See Microsoft Knowledge Base
'              280067 at https://mskb.pkisolutions.com/kb/280067. The article suggests a chunk size of 131072 (2^17),
'              but my tests (on a 134Mb file) suggested 32768 (2^15). A further optimisation is to know the number of
'              characters in the stream, to avoid string concatenation inside the loop, hence the EstNumChars argument.
' Parameters :
'  Stream     : An ADODB.Stream
'  EstNumChars: An estimate of the number of characters in the stream. Performance is improved if this estimate is
'               accurate.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ReadAllFromStream(Stream As ADODB.Stream, Optional ByVal EstNumChars As Long) As String
            
          Const ChunkSize As Long = 32768
          Dim Chunk As String
          Dim Contents As String
          Dim i As Long

1         If EstNumChars = 0 Then EstNumChars = 10000

2         On Error GoTo ErrHandler
3         Contents = String(EstNumChars, " ")
4         i = 1
5         Do While Not Stream.EOS
6             Chunk = Stream.ReadText(ChunkSize)
7             If i - 1 + Len(Chunk) > Len(Contents) Then
                  'Increase length of Contents by a factor (at least) 2
8                 Contents = Contents & String(i - 1 + Len(Chunk), " ")
9             End If

10            Mid$(Contents, i, Len(Chunk)) = Chunk
11            i = i + Len(Chunk)
12        Loop

13        If (i - 1) < Len(Contents) Then
14            Contents = Left$(Contents, i - 1)
15        End If

16        ReadAllFromStream = Contents

17        Exit Function
ErrHandler:
18        ReThrow "ReadAllFromStream", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseEncoding
' Purpose    : Set by-ref arguments
' Parameters :
'  FileName:
'  Encoding: Optional argument passed in to CSVRead. If not passed, we delegate to DetectEncoding.
'            NB the encoding argument (which may be user input) is "standardised" by this method to one of
'            "ASCII", "ANSI", "UTF-8", "UTF-16"
'  CharSet : Set by reference.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseEncoding(FileName As String, ByRef Encoding As Variant, ByRef CharSet As String)
          Const Err_Encoding As String = "Encoding argument can usually be omitted, but otherwise Encoding must be " & _
              "either ""ASCII"", ""ANSI"", ""UTF-8"", or ""UTF-16"""
          
1         On Error GoTo ErrHandler
2         If IsEmpty(Encoding) Or IsMissing(Encoding) Then
3             Encoding = DetectEncoding(FileName)
4         End If
              
5         If VarType(Encoding) = vbString Then
6             Select Case UCase$(Replace$(Replace$(Encoding, "-", vbNullString), " ", vbNullString))
                  Case "ASCII"
7                     Encoding = "ASCII"
8                     CharSet = "us-ascii"
9                 Case "ANSI"
                      'Unfortunately "ANSI" is not well defined. See
                      'https://stackoverflow.com/questions/701882/what-is-ansi-format
                      'For a list of the character set names that are known by a system, see the subkeys of
                      'HKEY_CLASSES_ROOT\MIME\Database\Charset in the Windows Registry.
10                    Encoding = "ANSI"
11                    CharSet = "windows-1252"
12                Case "UTF8", "UTF8NOBOM"
13                    Encoding = "UTF-8"
14                    CharSet = "utf-8"
15                Case "UTF16", "UTF16NOBOM"
16                    Encoding = "UTF-16"
17                    CharSet = "utf-16"
18                Case Else
19                    Throw Err_Encoding
20            End Select
21        Else
22            Throw Err_Encoding
23        End If

24        Exit Sub
ErrHandler:
25        ReThrow "ParseEncoding", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsCTValid
' Purpose    : Is a "Convert Types string" (which can in fact be either a string or a Boolean) valid?
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsCTValid(CT As Variant) As Boolean

          Static rx As VBScript_RegExp_55.RegExp

1         On Error GoTo ErrHandler
2         If rx Is Nothing Then
3             Set rx = New RegExp
4             With rx
5                 .IgnoreCase = True
6                 .Pattern = "^[NDBETQK]*$"
7                 .Global = False        'Find first match only
8             End With
9         End If

10        If VarType(CT) = vbBoolean Then
11            IsCTValid = True
12        ElseIf VarType(CT) = vbString Then
13            IsCTValid = rx.Test(CT)
14        Else
15            IsCTValid = False
16        End If

17        Exit Function
ErrHandler:
18        ReThrow "IsCTValid", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CTsEqual
' Purpose    : Test if two CT strings (strings to define type conversion) are equal, i.e. will have the same effect
' -----------------------------------------------------------------------------------------------------------------------
Private Function CTsEqual(CT1 As Variant, CT2 As Variant) As Boolean
1         On Error GoTo ErrHandler
2         If VarType(CT1) = VarType(CT2) Then
3             If CT1 = CT2 Then
4                 CTsEqual = True
5                 Exit Function
6             End If
7         End If
8         CTsEqual = StandardiseCT(CT1) = StandardiseCT(CT2)
9         Exit Function
ErrHandler:
10        ReThrow "CTsEqual", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StandardiseCT
' Purpose    : Put a CT string into standard form so that two such can be compared.
' -----------------------------------------------------------------------------------------------------------------------
Private Function StandardiseCT(CT As Variant) As String
1         On Error GoTo ErrHandler
2         If VarType(CT) = vbBoolean Then
3             If CT Then
4                 StandardiseCT = "BDN"
5             Else
6                 StandardiseCT = vbNullString
7             End If
8             Exit Function
9         ElseIf VarType(CT) = vbString Then
10            StandardiseCT = IIf(InStr(1, CT, "B", vbTextCompare), "B", vbNullString) & _
                  IIf(InStr(1, CT, "D", vbTextCompare), "D", vbNullString) & _
                  IIf(InStr(1, CT, "E", vbTextCompare), "E", vbNullString) & _
                  IIf(InStr(1, CT, "N", vbTextCompare), "N", vbNullString) & _
                  IIf(InStr(1, CT, "Q", vbTextCompare), "Q", vbNullString) & _
                  IIf(InStr(1, CT, "T", vbTextCompare), "T", vbNullString) & _
                  IIf(InStr(1, CT, "K", vbTextCompare), "K", vbNullString)
11        End If
12        Exit Function
ErrHandler:
13        ReThrow "StandardiseCT", Err
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
'  KeepQuotes            : Set only if ConvertTypes is not an array
'  TrimFields            : Set only if ConvertTypes is not an array
'  ColByColFormatting    : Set to True if ConvertTypes is an array
'  HeaderRowNum          : As passed to CSVRead, used to throw an error if HeaderRowNum has not been specified when
'                          it needs to have been.
'  CTDict                : Set to a dictionary keyed on the elements of the left column (or top row) of ConvertTypes,
'                          each element containing the corresponding right (or bottom) element.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseConvertTypes(ByVal ConvertTypes As Variant, ByRef ShowNumbersAsNumbers As Boolean, _
          ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
          ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef KeepQuotes As Boolean, ByRef TrimFields As Boolean, _
          ByRef ColByColFormatting As Boolean, HeaderRowNum As Long, ByRef CTDict As Scripting.Dictionary)
          
          Const Err_2D As String = "If ConvertTypes is given as a two dimensional array then the " & _
              " lower bounds in each dimension must be 1"
          Const Err_Ambiguous As String = "ConvertTypes is ambiguous, it can be interpreted as two rows, or as two columns"
          Const Err_BadColumnIdentifier As String = "Column identifiers in the left column (or top row) of " & _
              "ConvertTypes must be strings or non-negative whole numbers"
          Const Err_BadCT As String = "Type Conversion given in bottom row (or right column) of ConvertTypes must be " & _
              "Booleans or strings containing letters NDBETQ"
          Const Err_ConvertTypes As String = "ConvertTypes must be a Boolean, a string with allowed letters ""NDBETQK"" or an array"
          Const Err_HeaderRowNum As String = "ConvertTypes specifies columns by their header (instead of by number), " & _
              "but HeaderRowNum has not been specified"
          
          Dim ColIdentifier As Variant
          Dim CT As Variant
          Dim i As Long
          Dim LCN As Long 'Left column number
          Dim NC As Long 'Number of columns
          Dim ND As Long 'Number of dimensions
          Dim NR As Long 'Number of rows
          Dim RCN As Long 'Right Column Number
          Dim Transposed As Boolean
          
1         On Error GoTo ErrHandler
2         If VarType(ConvertTypes) = vbBoolean Then ConvertTypes = StandardiseCT(ConvertTypes)

3         If VarType(ConvertTypes) = vbString Or IsEmpty(ConvertTypes) Then
4             ParseCTString CStr(ConvertTypes), ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, _
                  ShowErrorsAsErrors, ConvertQuoted, KeepQuotes, TrimFields
5             ColByColFormatting = False
6             Exit Sub
7         End If

8         If TypeName(ConvertTypes) = "Range" Then ConvertTypes = ConvertTypes.Value2
9         ND = NumDimensions(ConvertTypes)
10        If ND = 1 Then
11            ConvertTypes = OneDArrayToTwoDArray(ConvertTypes)
12        ElseIf ND = 2 Then
13            If LBound(ConvertTypes, 1) <> 1 Or LBound(ConvertTypes, 2) <> 1 Then
14                Throw Err_2D
15            End If
16        End If

17        NR = NRows(ConvertTypes)
18        NC = NCols(ConvertTypes)
19        If NR = 2 And NC = 2 Then
              'Tricky - have we been given two rows or two columns?
20            If Not IsCTValid(ConvertTypes(2, 2)) Then Throw Err_ConvertTypes
21            If IsCTValid(ConvertTypes(1, 2)) And IsCTValid(ConvertTypes(2, 1)) Then
22                If StandardiseCT(ConvertTypes(1, 2)) <> StandardiseCT(ConvertTypes(2, 1)) Then
23                    Throw Err_Ambiguous
24                End If
25            End If
26            If IsCTValid(ConvertTypes(2, 1)) Then
27                ConvertTypes = Transpose(ConvertTypes)
28                Transposed = True
29            End If
30        ElseIf NR = 2 Then
31            ConvertTypes = Transpose(ConvertTypes)
32            Transposed = True
33            NR = NC
34        ElseIf NC <> 2 Then
35            Throw Err_ConvertTypes
36        End If
37        LCN = LBound(ConvertTypes, 2)
38        RCN = LCN + 1
39        For i = LBound(ConvertTypes, 1) To UBound(ConvertTypes, 1)
40            ColIdentifier = ConvertTypes(i, LCN)
41            CT = ConvertTypes(i, RCN)
42            If IsNumber(ColIdentifier) Then
43                If ColIdentifier <> CLng(ColIdentifier) Then
44                    Throw Err_BadColumnIdentifier & _
                          " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & _
                          ") is " & CStr(ColIdentifier)
45                ElseIf ColIdentifier < 0 Then
46                    Throw Err_BadColumnIdentifier & " but ConvertTypes(" & _
                          IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & ") is " & CStr(ColIdentifier)
47                End If
48            ElseIf VarType(ColIdentifier) <> vbString Then
49                Throw Err_BadColumnIdentifier & " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & _
                      ") is of type " & TypeName(ColIdentifier)
50            End If
51            If Not IsCTValid(CT) Then
52                If VarType(CT) = vbString Then
53                    Throw Err_BadCT & " but ConvertTypes(" & IIf(Transposed, "2," & CStr(i), CStr(i) & ",2") & _
                          ") is string """ & CStr(CT) & """"
54                Else
55                    Throw Err_BadCT & " but ConvertTypes(" & IIf(Transposed, "2," & CStr(i), CStr(i) & ",2") & _
                          ") is of type " & TypeName(CT)
56                End If
57            End If
58            If CTDict.Exists(ColIdentifier) Then
59                If Not CTsEqual(CTDict.item(ColIdentifier), CT) Then
60                    Throw "ConvertTypes is contradictory. Column " & CStr(ColIdentifier) & _
                          " is specified to be converted using two different conversion rules: " & CStr(CT) & _
                          " and " & CStr(CTDict.item(ColIdentifier))
61                End If
62            Else
63                CT = StandardiseCT(CT)
                  'Need this line to ensure that we parse the DateFormat provided when doing Col-by-col type conversion
64                If InStr(CT, "D") > 0 Then ShowDatesAsDates = True
65                If VarType(ColIdentifier) = vbString Then
66                    If HeaderRowNum = 0 Then
67                        Throw Err_HeaderRowNum
68                    End If
69                End If
70                CTDict.Add ColIdentifier, CT
71            End If
72        Next i
73        ColByColFormatting = True
74        Exit Sub
ErrHandler:
75        ReThrow "ParseConvertTypes", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCTString
' Purpose    : Parse the input ConvertTypes to set seven Boolean flags which are passed by reference.
' Parameters :
'  ConvertTypes          : The argument to CSVRead
'  ShowNumbersAsNumbers  : Should fields in the file that look like numbers be returned as Numbers? (Doubles)
'  ShowDatesAsDates      : Should fields in the file that look like dates with the specified DateFormat be returned as
'                          Dates?
'  ShowBooleansAsBooleans: Should fields in the file that match one of the TrueStrings or FalseStrings be returned as
'                          Booleans?
'  ShowErrorsAsErrors    : Should fields in the file that look like Excel errors (#N/A #REF! etc) be returned as errors?
'  ConvertQuoted         : Should the four conversion rules above apply even to quoted fields?
'  KeepQuotes            : Should quotes be kept? If True then quotes fields are returned to Excel with their leading and trailing quotes..
'  TrimFields            : Should leading and trailing spaces be trimmed from fields?
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseCTString(ByVal ConvertTypes As String, ByRef ShowNumbersAsNumbers As Boolean, _
          ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
          ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef KeepQuotes As Boolean, ByRef TrimFields As Boolean)

          Const Err_ConvertTypes As String = "ConvertTypes must be Boolean or string with allowed letters NDBETQ. " & _
              """N"" show numbers as numbers, ""D"" show dates as dates, ""B"" show Booleans " & _
              "as Booleans, ""E"" show Excel errors as errors, ""T"" to trim leading and trailing " & _
              "spaces from fields, ""Q"" rules NDBE apply even to quoted fields, TRUE = ""NDB"" " & _
              "(convert unquoted numbers, dates and Booleans), FALSE = no conversion"
          Const Err_Quoted As String = "ConvertTypes is incorrect, ""Q"" indicates that conversion should apply even to " & _
              "quoted fields, but none of ""N"", ""D"", ""B"" or ""E"" are present to indicate which type conversion to apply"
          Const Err_KQ As String = "ConvertTypes is incorrect, since it contains both ""Q"" and ""K"" which specify incompatible treatment of quoted fields"
          Dim i As Long

1         On Error GoTo ErrHandler

2         ShowNumbersAsNumbers = False
3         ShowDatesAsDates = False
4         ShowBooleansAsBooleans = False
5         ShowErrorsAsErrors = False
6         ConvertQuoted = False
7         KeepQuotes = False
8         For i = 1 To Len(ConvertTypes)
              'Adding another letter? Also change methods IsCTValid and StandardiseCT.
9             Select Case UCase$(Mid$(ConvertTypes, i, 1))
                  Case "N"
10                    ShowNumbersAsNumbers = True
11                Case "D"
12                    ShowDatesAsDates = True
13                Case "B"
14                    ShowBooleansAsBooleans = True
15                Case "E"
16                    ShowErrorsAsErrors = True
17                Case "Q"
18                    ConvertQuoted = True
19                Case "T"
20                    TrimFields = True
21                Case "K"
22                    KeepQuotes = True
23                Case Else
24                    Throw Err_ConvertTypes & " Found unrecognised character '" _
                          & Mid$(ConvertTypes, i, 1) & "'"
25            End Select
26        Next i
          
27        If ConvertQuoted And KeepQuotes Then
28            Throw Err_KQ
29        End If
          
30        If ConvertQuoted And Not (ShowNumbersAsNumbers Or ShowDatesAsDates Or _
              ShowBooleansAsBooleans Or ShowErrorsAsErrors) Then
31            Throw Err_Quoted
32        End If

33        Exit Sub
ErrHandler:
34        ReThrow "ParseCTString", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Min4
' Purpose    : Returns the minimum of four inputs and an indicator of which of the four was the minimum
' -----------------------------------------------------------------------------------------------------------------------
Private Function Min4(N1 As Long, N2 As Long, N3 As Long, _
          N4 As Long, ByRef Which As Long) As Long

1         If N1 < N2 Then
2             Min4 = N1
3             Which = 1
4         Else
5             Min4 = N2
6             Which = 2
7         End If

8         If N3 < Min4 Then
9             Min4 = N3
10            Which = 3
11        End If

12        If N4 < Min4 Then
13            Min4 = N4
14            Which = 4
15        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DetectEncoding
' Purpose    : Attempt to detect the file's encoding by looking for a byte order mark
' -----------------------------------------------------------------------------------------------------------------------
Private Function DetectEncoding(FilePath As String)

          Dim intAsc1Chr As Long
          Dim intAsc2Chr As Long
          Dim intAsc3Chr As Long
          Dim T As Scripting.TextStream

1         On Error GoTo ErrHandler
          
2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
          
3         Set T = m_FSO.OpenTextFile(FilePath, 1, False, 0)
4         If T.AtEndOfStream Then
5             DetectEncoding = "ANSI"
6             GoTo EarlyExit
7         End If
          
8         intAsc1Chr = Asc(T.Read(1))
9         If T.AtEndOfStream Then
10            DetectEncoding = "ANSI"
11            GoTo EarlyExit
12        End If
          
13        intAsc2Chr = Asc(T.Read(1))
14        If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
              'File is probably encoded UTF-16 LE BOM (little endian, with Byte Order Mark)
15            DetectEncoding = "UTF-16"
16        ElseIf (intAsc1Chr = 254) And (intAsc2Chr = 255) Then
              'File is probably encoded UTF-16 BE BOM (big endian, with Byte Order Mark)
17            DetectEncoding = "UTF-16"
18        Else
19            If T.AtEndOfStream Then
20                DetectEncoding = "ANSI"
21                GoTo EarlyExit
22            End If
23            intAsc3Chr = Asc(T.Read(1))
24            If (intAsc1Chr = 239) And (intAsc2Chr = 187) And (intAsc3Chr = 191) Then
                  'File is probably encoded UTF-8 with BOM
25                DetectEncoding = "UTF-8"
26            Else
                  'We don't know, assume ANSI but that may be incorrect.
27                DetectEncoding = "ANSI"
28            End If
29        End If

EarlyExit:
30        T.Close: Set T = Nothing
31        Exit Function

ErrHandler:
32        If Not T Is Nothing Then T.Close
33        ReThrow "DetectEncoding", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : EstimateNumChars
' Purpose    : Estimate the number of characters in a file. For Ansii files and UTF files with BOM and containing only
'              ansi characters the estimate will be exact, otherwise when UTF files contain high codepoint characters
'              the return will be an overestimate
' -----------------------------------------------------------------------------------------------------------------------
Private Function EstimateNumChars(FileSize As Long, Encoding As String)
1         Select Case Encoding
              Case "ANSI", "ASCII"
                  'will be exact
2                 EstimateNumChars = FileSize
3             Case "UTF-16"
                  'Will be exact if the file has a BOM (2 bytes) and contains only
                  'ansi characters (2 bytes each). When file contains non-ansi characters
                  'this will overestimate the character count.
4                 EstimateNumChars = (FileSize - 2) / 2
5             Case "UTF-8"
                  'Will be exact if the file has a BOM (3 bytes) and contains only ansi characters (1 byte each).
6                 EstimateNumChars = (FileSize - 3)
7             Case Else
8                 Throw "Unrecognised encoding '" & Encoding & "'"
9         End Select

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AmendDelimiterIfFirstFieldIsDateTime
' Purpose    : Subroutine of InferDelimiter. When the first field in a file is a date-with-time this method allows us to
'              avoid interpreting the colons inside that first field as the file delimiter.
' Parameters :
'  FirstChunk: Beginning part of the file.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AmendDelimiterIfFirstFieldIsDateTime(FirstChunk As String, ByRef Delimiter As String)

          Dim Converted As Boolean
          Dim DateOrder As Long
          Dim DateSeparator As Variant
          Dim DelimAt As Long
          Dim DtOut As Date
          Dim FirstField As String
          Dim SysDateSeparator As String
          Dim TrialDelim As Variant
          
1         On Error GoTo ErrHandler
2         For Each TrialDelim In Array(",", vbTab, "|", ";", vbCr, vbLf)
3             DelimAt = InStr(FirstChunk, CStr(TrialDelim))
4             If DelimAt > 0 Then
5                 FirstField = Left$(FirstChunk, DelimAt - 1)
6                 If InStr(FirstField, "-") > 0 Or InStr(FirstField, "/") > 0 Or InStr(FirstField, " ") > 0 Then

7                     SysDateSeparator = Application.International(xlDateSeparator)

8                     For Each DateSeparator In Array("/", "-", " ")
9                         For DateOrder = 0 To 2
10                            Converted = False
11                            CastToDate FirstField, DtOut, DateOrder, CStr(DateSeparator), SysDateSeparator, Converted
12                            If Not Converted Then
13                                CastISO8601 FirstField, DtOut, Converted, True, True
14                            End If
15                            If Converted Then
16                                Select Case TrialDelim
                                      Case vbCr, vbLf
17                                        Delimiter = ","
18                                    Case Else
19                                        Delimiter = TrialDelim
20                                End Select
21                                Exit Sub
22                            End If
23                        Next DateOrder
24                    Next
25                End If
26            End If
27        Next TrialDelim
          
28        Exit Sub
ErrHandler:
29        ReThrow "AmendDelimiterIfFirstFieldIsDateTime", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferDelimiter
' Purpose    : Infer the delimiter in a file by looking for first occurrence outside quoted regions of comma, tab,
'              semi-colon, colon or pipe (|). Only look in the first 10,000 characters, Would prefer to look at the first
'              10 lines, but that presents a problem for files with Mac line endings as T.ReadLine doesn't work for them.
'              But see also sub-routine AmendDelimiterIfFirstFieldIsDateTime.
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferDelimiter(st As enmSourceType, FileNameOrContents As String, _
          DecimalSeparator As String, CharSet As String) As String
          
          Const CHUNK_SIZE As Long = 1000
          Const Err_SourceType As String = "Cannot infer delimiter directly from URL"
          Const MAX_CHUNKS As Long = 10
          Dim Contents As String
          Dim EvenQuotes As Boolean
          Dim F As Scripting.File
          Dim i As Long
          Dim j As Long
          Dim MaxChars As Long
          Dim Stream As ADODB.Stream
          Const Err_FileEmpty As String = "File is empty"

1         On Error GoTo ErrHandler

2         EvenQuotes = True
3         If st = st_File Then

4             Set Stream = New ADODB.Stream
5             Stream.CharSet = CharSet
6             Stream.Open
7             Stream.LoadFromFile FileNameOrContents
8             If Stream.EOS Then Throw Err_FileEmpty

9             Do While Not Stream.EOS And j <= MAX_CHUNKS
10                j = j + 1
11                Contents = Stream.ReadText(CHUNK_SIZE)
12                For i = 1 To Len(Contents)
13                    Select Case Mid$(Contents, i, 1)
                          Case DQ
14                            EvenQuotes = Not EvenQuotes
15                        Case ",", vbTab, "|", ";", ":"
16                            If EvenQuotes Then
17                                If Mid$(Contents, i, 1) <> DecimalSeparator Then
18                                    InferDelimiter = Mid$(Contents, i, 1)
19                                    If InferDelimiter = ":" Then
20                                        If j = 1 Then
21                                            AmendDelimiterIfFirstFieldIsDateTime Contents, InferDelimiter
22                                        End If
23                                    End If
24                                    Stream.Close: Set Stream = Nothing: Set F = Nothing
25                                    Exit Function
26                                End If
27                            End If
28                    End Select
29                Next i
30            Loop
31            Stream.Close: Set Stream = Nothing: Set F = Nothing
32        ElseIf st = st_String Then
33            Contents = FileNameOrContents
34            MaxChars = MAX_CHUNKS * CHUNK_SIZE
35            If MaxChars > Len(Contents) Then MaxChars = Len(Contents)

36            For i = 1 To MaxChars
37                Select Case Mid$(Contents, i, 1)
                      Case DQ
38                        EvenQuotes = Not EvenQuotes
39                    Case ",", vbTab, "|", ";", ":"
40                        If EvenQuotes Then
41                            If Mid$(Contents, i, 1) <> DecimalSeparator Then
42                                InferDelimiter = Mid$(Contents, i, 1)
43                                If InferDelimiter = ":" Then
44                                    If i < 100 Then
45                                        AmendDelimiterIfFirstFieldIsDateTime Contents, InferDelimiter
46                                    End If
47                                End If
48                                Exit Function
49                            End If
50                        End If
51                End Select
52            Next i
53        Else
54            Throw Err_SourceType
55        End If

          'No commonly-used delimiter found in the file outside quoted regions
          'and in the first MAX_CHUNKS * CHUNK_SIZE characters. Assume comma
          'unless that's the decimal separator.
          
56        If DecimalSeparator = "," Then
57            InferDelimiter = ";"
58        Else
59            InferDelimiter = ","
60        End If

61        Exit Function
ErrHandler:
62        If Not Stream Is Nothing Then
63            Stream.Close
64            Set Stream = Nothing: Set F = Nothing
65        End If
66        ReThrow "InferDelimiter", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDateFormat
' Purpose    : Populate DateOrder and DateSeparator by parsing DateFormat.
' Parameters :
'  DateFormat   : String such as "D/M/Y" or "Y-M-D" or "M D Y"
'  DateOrder    : ByRef argument is set to DateFormat using same convention as Application.International(xlDateOrder)
'                 (0 = MDY, 1 = DMY, 2 = YMD)
'  DateSeparator: ByRef argument is set to the DateSeparator, typically "-" or "/", but can also be space character.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseDateFormat(ByVal DateFormat As String, ByRef DateOrder As Long, ByRef DateSeparator As String, _
          ByRef ISO8601 As Boolean, ByRef AcceptWithoutTimeZone As Boolean, ByRef AcceptWithTimeZone As Boolean)

          Dim Err_DateFormat As String

1         On Error GoTo ErrHandler
          
2         If UCase$(DateFormat) = "ISO" Then
3             ISO8601 = True
4             AcceptWithoutTimeZone = True
5             AcceptWithTimeZone = False
6             Exit Sub
7         ElseIf UCase$(DateFormat) = "ISOZ" Then
8             ISO8601 = True
9             AcceptWithoutTimeZone = False
10            AcceptWithTimeZone = True
11            Exit Sub
12        End If
          
13        ISO8601 = False
          
14        Err_DateFormat = "DateFormat not valid should be one of 'ISO', 'ISOZ', 'M-D-Y', 'D-M-Y', 'Y-M-D', " & _
              "'M/D/Y', 'D/M/Y', 'Y/M/D', 'M D Y', 'D M Y' or 'Y M D'" & ". Omit to use the default date format of 'Y-M-D'"
              
15        DateFormat = UCase$(DateFormat)
          'Replace repeated D's with a single D, etc since CastToDate only needs
          'to know the order in which the three parts of the date appear.
16        If Len(DateFormat) > 5 Then
17            ReplaceRepeats DateFormat, "D"
18            ReplaceRepeats DateFormat, "M"
19            ReplaceRepeats DateFormat, "Y"
20        End If
             
21        If Len(DateFormat) = 0 Then 'use "Y-M-D"
22            DateOrder = 2
23            DateSeparator = "-"
24        ElseIf Len(DateFormat) <> 5 Then
25            Throw Err_DateFormat
26        ElseIf Mid$(DateFormat, 2, 1) <> Mid$(DateFormat, 4, 1) Then
27            Throw Err_DateFormat
28        Else
29            DateSeparator = Mid$(DateFormat, 2, 1)
30            If DateSeparator <> "/" And DateSeparator <> "-" And DateSeparator <> " " Then Throw Err_DateFormat
31            Select Case UCase$(Left$(DateFormat, 1) & Mid$(DateFormat, 3, 1) & Right$(DateFormat, 1))
                  Case "MDY"
32                    DateOrder = 0
33                Case "DMY"
34                    DateOrder = 1
35                Case "YMD"
36                    DateOrder = 2
37                Case Else
38                    Throw Err_DateFormat
39            End Select
40        End If

41        Exit Sub
ErrHandler:
42        ReThrow "ParseDateFormat", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReplaceRepeats
' Purpose    : Replace repeated instances of a character in a string with a single instance.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ReplaceRepeats(ByRef TheString As String, TheChar As String)
          Dim ChCh As String
1         ChCh = TheChar & TheChar
2         Do While InStr(TheString, ChCh) > 0
3             TheString = Replace$(TheString, ChCh, TheChar)
4         Loop
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCSVContents
' Purpose    : Parse the contents of a CSV file. Returns a string Buffer together with arrays which assist splitting
'              Buffer into a two-dimensional array.
' Parameters :
'  ContentsOrStream: The contents of a CSV file as either a string, a Scripting.TextStream or an ADODB.Stream.
'  QuoteChar       : The quote character, usually ascii 34 ("), which allow fields to contain characters that would
'                    otherwise be significant to parsing, such as delimiters or new line characters.
'  Delimiter       : The string that separates fields within each line. Typically a single character, but needn't be.
'  Comment         : Lines in the file that start with these characters will be ignored, handled by method SkipLines.
'  SkipToRow       : Rows in the file prior to SkipToRow are ignored.
'  IgnoreEmptyLines: If True, lines in the file with no characters will be ignored, handled by method SkipLines.
'  IgnoreRepeated  : If True then parsing ignores delimiters at the start of lines, consecutive delimiters and delimiters
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
'  HeaderRow       : Set equal to the contents of the header row in the file, no type conversion, but quoted fields are
'                    unquoted and leading and trailing spaces are removed.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseCSVContents(ContentsOrStream As Variant, QuoteChar As String, _
          Delimiter As String, Comment As String, IgnoreEmptyLines As Boolean, _
          IgnoreRepeated As Boolean, SkipToRow As Long, HeaderRowNum As Long, NumRows As Long, _
          ByRef NumRowsFound As Long, ByRef NumColsFound As Long, ByRef NumFields As Long, ByRef Ragged As Boolean, _
          ByRef Starts() As Long, ByRef Lengths() As Long, ByRef RowIndexes() As Long, ByRef ColIndexes() As Long, _
          ByRef QuoteCounts() As Long, ByRef HeaderRow As Variant) As String

          Const Err_Delimiter As String = "Delimiter must not be the null string"
          Dim Buffer As String
          Dim BufferUpdatedTo As Long
          Dim ColNum As Long
          Dim DoSkipping As Boolean
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
          Dim quoteCount As Long
          Dim RowNum As Long
          Dim SearchFor() As String
          Dim Stream As Object
          Dim Streaming As Boolean
          Dim Tmp As Long
          Dim Which As Long

1         On Error GoTo ErrHandler
2         HeaderRow = Empty
          
3         If VarType(ContentsOrStream) = vbString Then
4             Buffer = ContentsOrStream
5             Streaming = False
6         Else
7             Set Stream = ContentsOrStream
8             If NumRows = 0 Then
9                 Buffer = ReadAllFromStream(Stream)
10                Streaming = False
11            Else
12                GetMoreFromStream Stream, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
13                Streaming = True
14            End If
15        End If
             
16        LComment = Len(Comment)
17        If LComment > 0 Or IgnoreEmptyLines Then
18            DoSkipping = True
19        End If
             
20        If Streaming Then
21            ReDim SearchFor(1 To 4)
22            SearchFor(1) = Delimiter
23            SearchFor(2) = vbLf
24            SearchFor(3) = vbCr
25            SearchFor(4) = QuoteChar
26            ReDim QuoteArray(1 To 1)
27            QuoteArray(1) = QuoteChar
28        End If

29        ReDim Starts(1 To 8): ReDim Lengths(1 To 8): ReDim RowIndexes(1 To 8)
30        ReDim ColIndexes(1 To 8): ReDim QuoteCounts(1 To 8)
          
31        LDlm = Len(Delimiter)
32        If LDlm = 0 Then Throw Err_Delimiter 'Avoid infinite loop!
33        OrigLen = Len(Buffer)
34        If Not Streaming Then
              'Ensure Buffer terminates with vbCrLf
35            If Right$(Buffer, 1) <> vbCr And Right$(Buffer, 1) <> vbLf Then
36                Buffer = Buffer & vbCrLf
37            ElseIf Right$(Buffer, 1) = vbCr Then
38                Buffer = Buffer & vbLf
39            End If
40            BufferUpdatedTo = Len(Buffer)
41        End If
          
42        i = 0: j = 1
          
43        If DoSkipping Then
44            SkipLines Streaming, Comment, LComment, IgnoreEmptyLines, _
                  Stream, Delimiter, Buffer, i, QuoteChar, PosLF, PosCR, BufferUpdatedTo
45        End If
          
46        If IgnoreRepeated Then
              'IgnoreRepeated: Handle repeated delimiters at the start of the first line
47            Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
48                i = i + LDlm
49            Loop
50        End If
          
51        ColNum = 1: RowNum = 1
52        EvenQuotes = True
53        Starts(1) = i + 1
54        If SkipToRow = 1 Then HaveReachedSkipToRow = True

55        Do
56            If EvenQuotes Then
57                If Not Streaming Then
58                    If PosDL <= i Then PosDL = InStr(i + 1, Buffer, Delimiter): If PosDL = 0 Then PosDL = BufferUpdatedTo + 1
59                    If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
60                    If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
61                    If PosQC <= i Then PosQC = InStr(i + 1, Buffer, QuoteChar): If PosQC = 0 Then PosQC = BufferUpdatedTo + 1
62                    i = Min4(PosDL, PosLF, PosCR, PosQC, Which)
63                Else
64                    i = SearchInBuffer(SearchFor, i + 1, Stream, Delimiter, QuoteChar, Which, Buffer, BufferUpdatedTo)
65                End If

66                If i >= BufferUpdatedTo + 1 Then
67                    Exit Do
68                End If

69                If j + 1 > UBound(Starts) Then
70                    ReDim Preserve Starts(1 To UBound(Starts) * 2)
71                    ReDim Preserve Lengths(1 To UBound(Lengths) * 2)
72                    ReDim Preserve RowIndexes(1 To UBound(RowIndexes) * 2)
73                    ReDim Preserve ColIndexes(1 To UBound(ColIndexes) * 2)
74                    ReDim Preserve QuoteCounts(1 To UBound(QuoteCounts) * 2)
75                End If

76                Select Case Which
                      Case 1
                          'Found Delimiter
77                        Lengths(j) = i - Starts(j)
78                        If IgnoreRepeated Then
79                            Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
80                                i = i + LDlm
81                            Loop
82                        End If
                          
83                        Starts(j + 1) = i + LDlm
84                        ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
85                        ColNum = ColNum + 1
86                        QuoteCounts(j) = quoteCount: quoteCount = 0
87                        j = j + 1
88                        NumFields = NumFields + 1
89                        i = i + LDlm - 1
90                    Case 2, 3
                          'Found line ending
91                        Lengths(j) = i - Starts(j)
92                        If Which = 3 Then 'Found a vbCr
93                            If Mid$(Buffer, i + 1, 1) = vbLf Then
                                  'Ending is Windows rather than Mac or Unix.
94                                i = i + 1
95                            End If
96                        End If
                          
97                        If DoSkipping Then
98                            SkipLines Streaming, Comment, LComment, IgnoreEmptyLines, Stream, _
                                  Delimiter, Buffer, i, QuoteChar, PosLF, PosCR, BufferUpdatedTo
99                        End If
                          
100                       If IgnoreRepeated Then
                              'IgnoreRepeated: Handle repeated delimiters at the end of the line, _
                               all but one will have already been handled.
101                           If Lengths(j) = 0 Then
102                               If ColNum > 1 Then
103                                   j = j - 1
104                                   ColNum = ColNum - 1
105                                   NumFields = NumFields - 1
106                               End If
107                           End If
                              'IgnoreRepeated: handle delimiters at the start of the next line to be parsed
108                           Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
109                               i = i + LDlm
110                           Loop
111                       End If
112                       Starts(j + 1) = i + 1

113                       If ColNum > NumColsFound Then
114                           If NumColsFound > 0 Then
115                               Ragged = True
116                           End If
117                           NumColsFound = ColNum
118                       ElseIf ColNum < NumColsFound Then
119                           Ragged = True
120                       End If
                          
121                       ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
122                       QuoteCounts(j) = quoteCount: quoteCount = 0
                          
123                       If HaveReachedSkipToRow Then
124                           If RowNum + SkipToRow - 1 = HeaderRowNum Then
125                               HeaderRow = GetLastParsedRow(Buffer, Starts, Lengths, _
                                      ColIndexes, QuoteCounts, j, False) 'TODO need to consider whether passing KeepQuotes as false is correct behaviour in this case
126                           End If
127                       Else
128                           If RowNum = HeaderRowNum Then
129                               HeaderRow = GetLastParsedRow(Buffer, Starts, Lengths, _
                                      ColIndexes, QuoteCounts, j, False) 'TODO need to consider whether passing KeepQuotes as false is correct behaviour in this case
130                           End If
131                       End If
                          
132                       ColNum = 1: RowNum = RowNum + 1
                          
133                       j = j + 1
134                       NumFields = NumFields + 1
                          
135                       If HaveReachedSkipToRow Then
136                           If RowNum = NumRows + 1 Then
137                               Exit Do
138                           End If
139                       Else
140                           If RowNum = SkipToRow Then
141                               HaveReachedSkipToRow = True
142                               Tmp = Starts(j)
143                               ReDim Starts(1 To 8): ReDim Lengths(1 To 8): ReDim RowIndexes(1 To 8)
144                               ReDim ColIndexes(1 To 8): ReDim QuoteCounts(1 To 8)
145                               RowNum = 1: j = 1: NumFields = 0
146                               Starts(1) = Tmp
147                           End If
148                       End If
149                   Case 4
                          'Found QuoteChar
150                       EvenQuotes = False
151                       quoteCount = quoteCount + 1
152               End Select
153           Else
154               If Not Streaming Then
155                   PosQC = InStr(i + 1, Buffer, QuoteChar)
156               Else
157                   If PosQC <= i Then PosQC = SearchInBuffer(QuoteArray, i + 1, Stream, _
                          Delimiter, QuoteChar, 0, Buffer, BufferUpdatedTo)
158               End If
                  
159               If PosQC = 0 Then
                      'Malformed Buffer (not RFC4180 compliant). There should always be an even number of double quotes.
                      'If there are an odd number then all text after the last double quote in the file will be (part of)
                      'the last field in the last line.
160                   Lengths(j) = OrigLen - Starts(j) + 1
161                   ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                      
162                   RowNum = RowNum + 1
163                   If ColNum > NumColsFound Then NumColsFound = ColNum
164                   NumFields = NumFields + 1
165                   Exit Do
166               Else
167                   i = PosQC
168                   EvenQuotes = True
169                   quoteCount = quoteCount + 1
170               End If
171           End If
172       Loop

173       NumRowsFound = RowNum - 1
          
174       If HaveReachedSkipToRow Then
175           NumRowsInFile = SkipToRow - 1 + RowNum - 1
176       Else
177           NumRowsInFile = RowNum - 1
178       End If
          
179       If SkipToRow > NumRowsInFile Then
180           If NumRows = 0 Then 'Attempting to read from SkipToRow to the end of the file, but that would be zero or
                  'a negative number of rows. So throw an error.
                  Dim RowDescription As String
181               If IgnoreEmptyLines And Len(Comment) > 0 Then
182                   RowDescription = "not commented, not empty "
183               ElseIf IgnoreEmptyLines Then
184                   RowDescription = "not empty "
185               ElseIf Len(Comment) > 0 Then
186                   RowDescription = "not commented "
187               End If
                                   
188               Throw "SkipToRow (" & CStr(SkipToRow) & ") exceeds the number of " & RowDescription & _
                      "rows in the file (" & CStr(NumRowsInFile) & ")"
189           Else
                  'Attempting to read a set number of rows, function CSVRead will return an array of Empty values.
190               NumFields = 0
191               NumRowsFound = 0
192           End If
193       End If

194       ParseCSVContents = Buffer

195       Exit Function
ErrHandler:
196       ReThrow "ParseCSVContents", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetLastParsedRow
' Purpose    : For use during parsing (fn ParseCSVContents) to grab the header row (which may or may not be part of the
'              function return). The argument j should point into the Starts, Lengths etc arrays, pointing to the last
'              field in the header row
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetLastParsedRow(Buffer As String, Starts() As Long, Lengths() As Long, _
          ColIndexes() As Long, QuoteCounts() As Long, j As Long, KeepQuotes As Boolean) As Variant
          Dim NC As Long

          Dim Field As String
          Dim i As Long
          Dim Res() As String

1         On Error GoTo ErrHandler
2         NC = ColIndexes(j)

3         ReDim Res(1 To 1, 1 To NC)
4         For i = j To j - NC + 1 Step -1
5             Field = Mid$(Buffer, Starts(i), Lengths(i))
6             Res(1, NC + i - j) = Unquote(Trim$(Field), DQ, QuoteCounts(i), KeepQuotes)
7         Next i
8         GetLastParsedRow = Res

9         Exit Function
ErrHandler:
10        ReThrow "GetLastParsedRow", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SkipLines
' Purpose    : Sub-routine of ParseCSVContents. Skip a commented or empty row by incrementing i to the position of
'              the line feed just before the next not-commented line.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SkipLines(Streaming As Boolean, Comment As String, _
          LComment As Long, IgnoreEmptyLines As Boolean, Stream As Object, ByVal Delimiter As String, _
          ByRef Buffer As String, ByRef i As Long, QuoteChar As String, ByVal PosLF As Long, ByVal PosCR As Long, _
          ByRef BufferUpdatedTo As Long)
          
          Dim AtEndOfStream As Boolean
          Dim LookAheadBy As Long
          Dim SkipThisLine As Boolean
          
1         On Error GoTo ErrHandler
2         Do
3             If Streaming Then
4                 LookAheadBy = MaxLngs(LComment, 2)
5                 If i + LookAheadBy > BufferUpdatedTo Then

6                     AtEndOfStream = Stream.EOS
7                     If Not AtEndOfStream Then
8                         GetMoreFromStream Stream, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
9                     End If
10                End If
11            End If

12            SkipThisLine = False
13            If LComment > 0 Then
14                If Mid$(Buffer, i + 1, LComment) = Comment Then
15                    SkipThisLine = True
16                End If
17            End If
18            If IgnoreEmptyLines Then
19                Select Case Mid$(Buffer, i + 1, 1)
                      Case vbLf, vbCr
20                        SkipThisLine = True
21                End Select
22            End If

23            If SkipThisLine Then
24                If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
25                If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
26                If PosLF < PosCR Then
27                    i = PosLF
28                ElseIf PosLF = PosCR + 1 Then
29                    i = PosLF
30                Else
31                    i = PosCR
32                End If
33            Else
34                Exit Do
35            End If
36        Loop

37        Exit Sub
ErrHandler:
38        ReThrow "SkipLines", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SearchInBuffer
' Purpose    : Sub-routine of ParseCSVContents. Returns the location in the buffer of the first-encountered string
'              amongst the elements of SearchFor, starting the search at point SearchFrom and finishing the search at
'              point BufferUpdatedTo. If none found in that region returns BufferUpdatedTo + 1. Otherwise returns the
'              location of the first found and sets the by-reference argument Which to indicate which element of
'              SearchFor was the first to be found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SearchInBuffer(SearchFor() As String, StartingAt As Long, Stream As Object, _
          Delimiter As String, QuoteChar As String, ByRef Which As Long, _
          ByRef Buffer As String, ByRef BufferUpdatedTo As Long) As Long

          Dim InstrRes As Long
          Dim PrevBufferUpdatedTo As Long

1         On Error GoTo ErrHandler

          'in this call only search as far as BufferUpdatedTo
2         InstrRes = InStrMulti(SearchFor, Buffer, StartingAt, BufferUpdatedTo, Which)
3         If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
4             SearchInBuffer = InstrRes
5             Exit Function
6         Else

7             If Stream.EOS Then
8                 SearchInBuffer = BufferUpdatedTo + 1
9                 Exit Function
10            End If
11        End If

12        Do
13            PrevBufferUpdatedTo = BufferUpdatedTo
14            GetMoreFromStream Stream, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
15            InstrRes = InStrMulti(SearchFor, Buffer, PrevBufferUpdatedTo + 1, BufferUpdatedTo, Which)
16            If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
17                SearchInBuffer = InstrRes
18                Exit Function
19            ElseIf Stream.EOS Then
20                SearchInBuffer = BufferUpdatedTo + 1
21                Exit Function
22            End If
23        Loop
24        Exit Function

25        Exit Function
ErrHandler:
26        ReThrow "SearchInBuffer", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InStrMulti
' Purpose    : Sub-routine of ParseCSVContents. Returns the first point in SearchWithin at which one of the elements of
'              SearchFor is found, search is restricted to region [StartingAt, EndingAt] and Which is updated with the
'              index identifying which was the first of the strings to be found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function InStrMulti(SearchFor() As String, SearchWithin As String, StartingAt As Long, _
          EndingAt As Long, ByRef Which As Long) As Long

          Const Inf As Long = 2147483647
          Dim i As Long
          Dim InstrRes() As Long
          Dim LB As Long
          Dim Result As Long
          Dim UB As Long

1         On Error GoTo ErrHandler
2         LB = LBound(SearchFor): UB = UBound(SearchFor)

3         Result = Inf

4         ReDim InstrRes(LB To UB)
5         For i = LB To UB
6             InstrRes(i) = InStr(StartingAt, SearchWithin, SearchFor(i))
7             If InstrRes(i) > 0 Then
8                 If InstrRes(i) <= EndingAt Then
9                     If InstrRes(i) < Result Then
10                        Result = InstrRes(i)
11                        Which = i
12                    End If
13                End If
14            End If
15        Next
16        InStrMulti = IIf(Result = Inf, 0, Result)

17        Exit Function
ErrHandler:
18        ReThrow "InStrMulti", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetMoreFromStream, Sub-routine of ParseCSVContents
' Purpose    : Write CHUNKSIZE characters from the Stream into the buffer, modifying the passed-by-reference
'              arguments  Buffer, BufferUpdatedTo and Streaming.
'              Complexities:
'           a) We have to be careful not to update the buffer to a point part-way through a two-character end-of-line
'              or a multi-character delimiter, otherwise calling method SearchInBuffer might give the wrong result.
'           b) We update a few characters of the buffer beyond the BufferUpdatedTo point with the delimiter, the
'              QuoteChar and vbCrLf. This ensures that the calls to Instr that search the buffer for these strings do
'              not needlessly scan the unupdated part of the buffer.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub GetMoreFromStream(Stream As ADODB.Stream, Delimiter As String, QuoteChar As String, _
          ByRef Buffer As String, ByRef BufferUpdatedTo As Long)

          Const ChunkSize As Long = 5000  ' The number of characters to read from the stream on each call.
          ' Set to a small number for testing logic and a bigger number for
          ' performance, but not too high since a common use case is reading
          ' just the first line of a file. Suggest 5000? Note that when reading
          ' an entire file (NumRows argument to CSVRead is zero) function
          ' GetMoreFromStream is not called.
          Dim AtEndOfStream As Boolean
          Dim ExpandBufferBy As Long
          Dim FirstPass As Boolean
          Dim i As Long
          Dim NCharsToWriteToBuffer As Long
          Dim NewChars As String
          Dim OKToExit As Boolean

1         On Error GoTo ErrHandler
          
2         FirstPass = True
3         Do
4             NewChars = Stream.ReadText(IIf(FirstPass, ChunkSize, 1))
5             AtEndOfStream = Stream.EOS
6             FirstPass = False
7             If AtEndOfStream Then
                  'Ensure NewChars terminates with vbCrLf
8                 If Right$(NewChars, 1) <> vbCr And Right$(NewChars, 1) <> vbLf Then
9                     NewChars = NewChars & vbCrLf
10                ElseIf Right$(NewChars, 1) = vbCr Then
11                    NewChars = NewChars & vbLf
12                End If
13            End If

14            NCharsToWriteToBuffer = Len(NewChars) + Len(Delimiter) + 3

15            If BufferUpdatedTo + NCharsToWriteToBuffer > Len(Buffer) Then
16                ExpandBufferBy = MaxLngs(Len(Buffer), NCharsToWriteToBuffer)
17                Buffer = Buffer & String(ExpandBufferBy, "?")
18            End If
              
19            Mid$(Buffer, BufferUpdatedTo + 1, Len(NewChars)) = NewChars
20            BufferUpdatedTo = BufferUpdatedTo + Len(NewChars)

21            OKToExit = True
              'Ensure we don't leave the buffer updated to part way through a two-character end of line marker.
22            If Right$(NewChars, 1) = vbCr Then
23                OKToExit = False
24            End If
              'Ensure we don't leave the buffer updated to a point part-way through a multi-character delimiter
25            If Len(Delimiter) > 1 Then
26                For i = 1 To Len(Delimiter) - 1
27                    If Mid$(Buffer, BufferUpdatedTo - i + 1, i) = Left$(Delimiter, i) Then
28                        OKToExit = False
29                        Exit For
30                    End If
31                Next i
32                If Mid$(Buffer, BufferUpdatedTo - Len(Delimiter) + 1, Len(Delimiter)) = Delimiter Then
33                    OKToExit = True
34                End If
35            End If
36            If OKToExit Then Exit Do
37        Loop

          'Line below arranges that when calling Instr(Buffer,....) we don't pointlessly scan the space characters
          'we can be sure that there is space in the buffer to write the extra characters thanks to
38        Mid$(Buffer, BufferUpdatedTo + 1, 2 + Len(QuoteChar) + Len(Delimiter)) = vbCrLf & QuoteChar & Delimiter

39        Exit Sub
ErrHandler:
40        ReThrow "GetMoreFromStream", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CountQuotes
' Purpose    : Count the quotes in a string, only used when applying column-by-column type conversion, because in that
'              case it's not possible to use the count of quotes made at parsing time which is organised row-by-row.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CountQuotes(Str As String, QuoteChar As String) As Long
          Dim N As Long
          Dim pos As Long

1         Do
2             pos = InStr(pos + 1, Str, QuoteChar)
3             If pos = 0 Then
4                 CountQuotes = N
5                 Exit Function
6             End If
7             N = N + 1
8         Loop
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
'  ConvertQuoted        : Should quoted fields (after quote removal) be converted according to arguments
'                         ShowNumbersAsNumbers, ShowDatesAsDates, and the contents of Sentinels.
'  KeepQuotes           : Should quotes be kept instead of removed?
'Numbers
'  ShowNumbersAsNumbers : If Field is a string representation of a number should the function return that number?
'  SepStandard          : Is the decimal separator the same as the system defaults? If True then the next two arguments
'                         are ignored.
'  DecimalSeparator     : The decimal separator used in Field.
'  SysDecimalSeparator  : The default decimal separator on the system.
'Dates
'  ShowDatesAsDates     : If Field is a string representation of a date should the function return that date?
'  ISO8601              : If Field is a date, does it respect (a subset of) ISO8601?
'  AcceptWithoutTimeZone: In the case of ISO8601 dates, should conversion be applied to dates-with-time that have no time
'                         zone information?
'  AcceptWithTimeZone   : In the case of ISO8601 dates, should conversion be applied to dates-with-time that have time
'                         zone information?
'  DateOrder            : If Field is a string representation what order of parts must it respect (not relevant if
'                         ISO8601 is True) 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  DateSeparator        : The date separator, must be either "-" or "/".
'  SysDateSeparator     : The Windows system date separator.
'Booleans, Errors, Missings
'  AnySentinels         : Does the sentinel dictionary have any elements?
'  Sentinels            : A dictionary of Sentinels. If Sentinels.Exists(Field) Then ConvertField = Sentinels(Field)
'  MaxSentinelLength    : The maximum length of the keys of Sentinels.
'  ShowMissingsAs       : The value to which missing fields (consecutive delimiters) are converted. If CSVRead has a
'                         MissingStrings argument then values matching those strings are also converted to
'                         ShowMissingsAs, thanks to method MakeSentinels.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ConvertField(Field As String, AnyConversion As Boolean, FieldLength As Long, _
          TrimFields As Boolean, QuoteChar As String, quoteCount As Long, ConvertQuoted As Boolean, KeepQuotes As Boolean, _
          ShowNumbersAsNumbers As Boolean, SepStandard As Boolean, DecimalSeparator As String, _
          SysDecimalSeparator As String, ShowDatesAsDates As Boolean, ISO8601 As Boolean, _
          AcceptWithoutTimeZone As Boolean, AcceptWithTimeZone As Boolean, DateOrder As Long, _
          DateSeparator As String, SysDateSeparator As String, _
          AnySentinels As Boolean, Sentinels As Dictionary, MaxSentinelLength As Long, _
          ShowMissingsAs As Variant, AscSeparator As Long) As Variant

          Dim Converted As Boolean
          Dim dblResult As Double
          Dim dtResult As Date

1         If TrimFields Then
2             If Left$(Field, 1) = " " Then
3                 Field = Trim$(Field)
4                 FieldLength = Len(Field)
5             ElseIf Right$(Field, 1) = " " Then
6                 Field = Trim$(Field)
7                 FieldLength = Len(Field)
8             End If
9         End If

10        If FieldLength = 0 Then
11            ConvertField = ShowMissingsAs
12            Exit Function
13        End If

14        If Not AnyConversion Then
15            If quoteCount = 0 Then
16                ConvertField = Field
17                Exit Function
18            End If
19        End If

20        If AnySentinels Then
21            If FieldLength <= MaxSentinelLength Then
22                If Sentinels.Exists(Field) Then
23                    ConvertField = Sentinels.item(Field)
24                    Exit Function
25                End If
26            End If
27        End If

28        If quoteCount > 0 Then
29            If KeepQuotes Then
30                ConvertField = Field
31                Exit Function
32            End If
33            If Left$(Field, 1) = QuoteChar Then
34                If Right$(Field, 1) = QuoteChar Then
35                    Field = Mid$(Field, 2, FieldLength - 2)
36                    If quoteCount > 2 Then
37                        Field = Replace$(Field, QuoteChar & QuoteChar, QuoteChar)
38                    End If
39                    If ConvertQuoted Then
40                        FieldLength = Len(Field)
41                    Else
42                        ConvertField = Field
43                        Exit Function
44                    End If
45                End If
46            End If
47        End If

48        If Not ConvertQuoted Then
49            If quoteCount > 0 Then
50                ConvertField = Field
51                Exit Function
52            End If
53        End If

54        If ShowNumbersAsNumbers Then
55            CastToDouble Field, dblResult, SepStandard, DecimalSeparator, AscSeparator, SysDecimalSeparator, Converted
56            If Converted Then
57                ConvertField = dblResult
58                Exit Function
59            End If
60        End If

61        If ShowDatesAsDates Then
62            If ISO8601 Then
63                CastISO8601 Field, dtResult, Converted, AcceptWithoutTimeZone, AcceptWithTimeZone
64            Else
65                CastToDate Field, dtResult, DateOrder, DateSeparator, SysDateSeparator, Converted
66            End If
67            If Not Converted Then
68                If InStr(Field, ":") > 0 Then
69                    CastToTime Field, dtResult, Converted
70                    If Not Converted Then
71                        CastToTimeB Field, dtResult, Converted
72                    End If
73                End If
74            End If
75            If Converted Then
76                ConvertField = dtResult
77                Exit Function
78            End If
79        End If

80        ConvertField = Field
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Unquote
' Purpose    : Unquote a field.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Unquote(ByVal Field As String, QuoteChar As String, quoteCount As Long, KeepQuotes As Boolean) As String

1         On Error GoTo ErrHandler
2         If quoteCount > 0 Then
3             If Not KeepQuotes Then
4                 If Left$(Field, 1) = QuoteChar Then
5                     If Right$(QuoteChar, 1) = QuoteChar Then
6                         Field = Mid$(Field, 2, Len(Field) - 2)
7                         If quoteCount > 2 Then
8                             Field = Replace$(Field, QuoteChar & QuoteChar, QuoteChar)
9                         End If
10                    End If
11                End If
12            End If
13        End If
14        Unquote = Field

15        Exit Function
ErrHandler:
16        ReThrow "Unquote", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble, sub-routine of ConvertField
' Purpose    : Casts strIn to double where strIn has specified decimals separator.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub CastToDouble(strIn As String, ByRef DblOut As Double, SepStandard As Boolean, _
          DecimalSeparator As String, AscSeparator As Long, SysDecimalSeparator As String, ByRef Converted As Boolean)
          
1         On Error GoTo ErrHandler
          'Checking the first character makes this function approx 12 times faster at rejecting non-numeric strings at the cost of about a 20% slowdown in converting numeric strings so worth it


2         Select Case Asc(strIn)
              Case 32, 43, 45, 46, 48 To 57, AscSeparator 'Characters " +-.0123456789" and also the decimal separator that may be different from "."
3                 If SepStandard Then
4                     DblOut = CDbl(strIn)
5                 Else
6                     If InStr(strIn, DecimalSeparator) > 0 Then
7                         DblOut = CDbl(Replace$(strIn, DecimalSeparator, SysDecimalSeparator))
8                     Else
9                         DblOut = CDbl(strIn)
10                    End If
11                End If
12                Converted = True
13        End Select
ErrHandler:
          'Do nothing - strIn was not a string representing a number.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TwoToFourDigitYear
' Author     : Philip Swannell
' Date       : 23-Feb-2023
' Purpose    : Amend in-place a two-digit "year part" of a date, paying attention to module-level constants
'              m_2DigitYearIsFrom and m_2DigitYearIsTo
' -----------------------------------------------------------------------------------------------------------------------
Sub TwoToFourDigitYear(ByRef y As String)
          Dim y_lng As Long

          Static rx As VBScript_RegExp_55.RegExp

1         If rx Is Nothing Then
2             Set rx = New RegExp
3             With rx
4                 .IgnoreCase = True
5                 .Pattern = "^[0-9][0-9]$"
6                 .Global = False
7             End With
8         End If
9         If Not rx.Test(y) Then Throw "Bad year part"
10        y_lng = (m_2DigitYearIsFrom \ 100) * 100 + CLng(y)
11        If y_lng > m_2DigitYearIsTo Then y_lng = y_lng - 100
12        If y_lng < m_2DigitYearIsFrom Then y_lng = y_lng + 100
13        y = CStr(y_lng)
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
'  SysDateSeparator: The Windows system date separator
'  Converted       : Boolean flipped to TRUE if conversion takes place
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDate(strIn As String, ByRef DtOut As Date, DateOrder As Long, _
          DateSeparator As String, SysDateSeparator As String, _
          ByRef Converted As Boolean)

          Dim d As String
          Dim m As String
          Dim pos1 As Long 'First date separator
          Dim pos2 As Long 'Second date separator
          Dim pos3 As Long 'Space to separate date from time
          Dim pos4 As Long 'decimal point for fractions of a second
          Dim Converted2 As Boolean
          Dim HasFractionalSecond As Boolean
          Dim HasTimePart As Boolean
          Dim ld As Long
          Dim lm As Long
          Dim ly As Long
          Dim TimePart As String
          Dim TimePartConverted As Date
          Dim y As String
          
          'Can reject most input before switching on error handling - for speed
1         pos1 = InStr(strIn, DateSeparator)
2         If pos1 = 0 Then Exit Sub
3         pos2 = InStr(pos1 + 1, strIn, DateSeparator)
4         If pos2 = 0 Then Exit Sub

5         On Error GoTo ErrHandler

6         pos3 = InStr(pos2 + 1, strIn, " ")
          
7         HasTimePart = pos3 > 0
8         If HasTimePart Then
9             pos4 = InStr(pos3 + 1, strIn, ".")
10            HasFractionalSecond = pos4 > 0
11            TimePart = Mid$(strIn, pos3)
12            If HasFractionalSecond Then
13                CastToTimeB Mid$(TimePart, 2), TimePartConverted, Converted2
14                If Not Converted2 Then Exit Sub
15                TimePart = ""
16            End If
17        End If
          
18        If DateOrder = 2 Then 'Y-M-D
19            ly = pos1 - 1
20            y = Left$(strIn, ly)
21            lm = pos2 - pos1 - 1
22            m = Mid$(strIn, pos1 + 1, lm)
23            If HasTimePart Then
24                ld = pos3 - pos2 - 1
25                d = Mid$(strIn, pos2 + 1, ld)
26            Else
27                ld = Len(strIn) - pos2
28                d = Mid$(strIn, pos2 + 1)
29                If pos1 = 5 Then
30                    DtOut = MyCDate(y, m, d, lm, ld)
31                    Converted = True
32                    Exit Sub
33                End If
34            End If
35            If pos1 = 5 Then 'Len(y)=4
36                If Not HasFractionalSecond Then
37                    DtOut = CDate(strIn)
38                    Converted = True
39                    Exit Sub
40                End If
41            End If
42        ElseIf DateOrder = 1 Then 'D-M-Y
43            ld = pos1 - 1
44            d = Left$(strIn, ld)
45            lm = pos2 - pos1 - 1
46            m = Mid$(strIn, pos1 + 1, lm)
47            If HasTimePart Then
48                ly = pos3 - pos2 - 1
49                y = Mid$(strIn, pos2 + 1, ly)
50            Else
51                ly = Len(strIn) - pos2
52                y = Mid$(strIn, pos2 + 1)
53            End If
54        ElseIf DateOrder = 0 Then 'M-D-Y
55            lm = pos1 - 1
56            m = Left$(strIn, pos1 - 1)
57            ld = pos2 - pos1 - 1
58            d = Mid$(strIn, pos1 + 1, ld)
59            If HasTimePart Then
60                ly = pos3 - pos2 - 1
61                y = Mid$(strIn, pos2 + 1, ly)
62            Else
63                ly = Len(strIn) - pos2
64                y = Mid$(strIn, pos2 + 1)
65            End If
66        Else
67            Throw "DateOrder must be 0, 1, or 2"
68        End If

69        If Not IsNumeric(d) Then Exit Sub
70        If Not IsNumeric(y) Then Exit Sub
71        If ld > 2 Then Exit Sub
72        If ly = 2 Then
73            TwoToFourDigitYear y
74        ElseIf ly <> 4 Then
75            Exit Sub
76        End If

77        If HasTimePart Then
78            DtOut = CDate(y & SysDateSeparator & m & SysDateSeparator & d & TimePart)
79            If HasFractionalSecond Then
80                DtOut = DtOut + TimePartConverted
81            End If
82        Else
83            DtOut = MyCDate(y, m, d, lm, ld)
84        End If

85        Converted = True
86        Exit Sub
ErrHandler:
          'Do nothing - was not a string representing a date with the specified date order and date separator.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MyCDate
' Author     : Philip Swannell
' Date       : 24-Feb-2023
' Purpose    : Replacement for CDate. Where possible, avoids string concatenation and uses DateSerial for better speed.
' Parameters :
'  y     : The year - MUST BE OF LENGTH FOUR.
'  m     : The month
'  d     : The Day
'  lm    : The length of m
'  ld    : The length of d
' -----------------------------------------------------------------------------------------------------------------------
Private Function MyCDate(y As String, m As String, d As String, lm As Long, ld As Long)

1         If lm <= 2 Then
2             If ld <= 2 Then
3                 If m <= 12 Then
4                     If d <= 28 Then
5                         If m > 0 Then
6                             If d > 0 Then
7                                 MyCDate = DateSerial(y, m, d)
8                                 Exit Function
9                             End If
10                        End If
11                    End If
12                End If
13            End If
14        End If
15        MyCDate = CDate(y & "-" & m & "-" & d)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToTime
' Purpose    : Cast strings that represent a time to a date, no handling of TimeZone.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToTime(strIn As String, ByRef DtOut As Date, ByRef Converted As Boolean)

1         On Error GoTo ErrHandler
          
2         DtOut = CDate(strIn)
3         If DtOut <= 1 Then
4             Converted = True
5         End If
6         Exit Sub
ErrHandler:
          'Do nothing, was not a valid time (e.g. h,m or s out of range)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToTimeB
' Purpose    : CDate does not correctly cope with times such as '04:20:10.123 am' or '04:20:10.123', i.e, times with a
'              fractional second, so this method is called after CastToTime
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToTimeB(strIn As String, ByRef DtOut As Date, ByRef Converted As Boolean)
          Static rx As VBScript_RegExp_55.RegExp
          Dim DecPointAt As Long
          Dim FractionalSecond As Double
          Dim SpaceAt As Long
          
1         On Error GoTo ErrHandler
2         If rx Is Nothing Then
3             Set rx = New RegExp
4             With rx
5                 .IgnoreCase = True
6                 .Pattern = "^[0-2]?[0-9]:[0-5]?[0-9]:[0-5]?[0-9](\.[0-9]+)( am| pm)?$"
7                 .Global = False        'Find first match only
8             End With
9         End If

10        If Not rx.Test(strIn) Then Exit Sub
11        DecPointAt = InStr(strIn, ".")
12        If DecPointAt = 0 Then Exit Sub ' should never happen
13        SpaceAt = InStr(strIn, " ")
14        If SpaceAt = 0 Then SpaceAt = Len(strIn) + 1
15        FractionalSecond = CDbl(Mid$(strIn, DecPointAt, SpaceAt - DecPointAt)) / 86400
          
16        DtOut = CDate(Left$(strIn, DecPointAt - 1) + Mid$(strIn, SpaceAt)) + FractionalSecond
17        Converted = True
18        Exit Sub
ErrHandler:
          'Do nothing, was not a valid time (e.g. h,m or s out of range)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastISO8601
' Purpose    : Convert ISO8601 formatted datestrings to UTC date. https://xkcd.com/1179/

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
Private Sub CastISO8601(ByVal strIn As String, ByRef DtOut As Date, ByRef Converted As Boolean, _
          AcceptWithoutTimeZone As Boolean, AcceptWithTimeZone As Boolean)

          Dim d As Long
          Dim L As Long
          Dim LocalTime As Double
          Dim m As Long
          Dim MilliPart As Double
          Dim MinusPos As Long
          Dim PlusPos As Long
          Dim Sign As Long
          Dim y As Long
          Dim ZAtEnd As Boolean
          
          Static rxNoNo As VBScript_RegExp_55.RegExp
          Static RxYesNo As VBScript_RegExp_55.RegExp
          Static RxNoYes As VBScript_RegExp_55.RegExp
          Static rxYesYes As VBScript_RegExp_55.RegExp
          Static rxExists As Boolean

1         On Error GoTo ErrHandler
          
2         If Not rxExists Then
3             Set rxNoNo = New RegExp
              'Reject datetime
4             With rxNoNo
5                 .IgnoreCase = False
6                 .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9]$"
7                 .Global = False
8             End With
              
              'Accept datetime without time zone, reject datetime with timezone
9             Set RxYesNo = New RegExp
10            With RxYesNo
11                .IgnoreCase = False
12                .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?)?$"
13                .Global = False
14            End With
              
              'Reject datetime without time zone, accept datetime with timezone
15            Set RxNoYes = New RegExp
16            With RxNoYes
17                .IgnoreCase = False
18                .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?(Z|((\+|\-)[0-2][0-9]:[0-5][0-9])))?$"
19                .Global = False
20            End With
              
              'Accept datetime, both with and without timezone
21            Set rxYesYes = New RegExp
22            With rxYesYes
23                .IgnoreCase = False
24                .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?((Z|((\+|\-)[0-2][0-9]:[0-5][0-9])))?)?$"
25                .Global = False
26            End With
27            rxExists = True
28        End If
          
29        L = Len(strIn)

30        If L = 10 Then
31            If rxNoNo.Test(strIn) Then
32                y = Left$(strIn, 4)
33                m = Mid$(strIn, 6, 2)
34                d = Mid$(strIn, 9, 2)
                  'Use DateSerial for speed, but carefully
35                If m > 0 And m <= 12 And d > 0 And d <= 28 Then
36                    DtOut = DateSerial(y, m, d)
37                Else
38                    DtOut = CDate(strIn)
39                End If
40                Converted = True
41                Exit Sub
42            End If
43        ElseIf L < 10 Then
44            Converted = False
45            Exit Sub
46        ElseIf L > 40 Then
47            Converted = False
48            Exit Sub
49        End If

50        Converted = False
          
51        If AcceptWithoutTimeZone Then
52            If AcceptWithTimeZone Then
53                If Not rxYesYes.Test(strIn) Then Exit Sub
54            Else
55                If Not RxYesNo.Test(strIn) Then Exit Sub
56            End If
57        Else
58            If AcceptWithTimeZone Then
59                If Not RxNoYes.Test(strIn) Then Exit Sub
60            Else
61                If Not rxNoNo.Test(strIn) Then Exit Sub
62            End If
63        End If
          
          'Replace the "T" separator
64        Mid$(strIn, 11, 1) = " "
          
65        If L = 19 Then
              'Tests show that CDate is faster than DateSerial(Mid$(... + TimeSerial(Mid$(...
66            DtOut = CDate(strIn)
67            Converted = True
68            Exit Sub
69        End If

70        If Right$(strIn, 1) = "Z" Then
71            Sign = 0
72            ZAtEnd = True
73        Else
74            PlusPos = InStr(20, strIn, "+")
75            If PlusPos > 0 Then
76                Sign = 1
77            Else
78                MinusPos = InStr(20, strIn, "-")
79                If MinusPos > 0 Then
80                    Sign = -1
81                End If
82            End If
83        End If

84        If Mid$(strIn, 20, 1) = "." Then 'Have fraction of a second
85            Select Case Sign
                  Case 0
                      'Example: "2021-08-23T08:47:20.920Z"
86                    MilliPart = CDbl(Mid$(strIn, 20, IIf(ZAtEnd, L - 20, L - 19)))
87                Case 1
                      'Example: "2021-08-23T08:47:20.920+05:00"
88                    MilliPart = CDbl(Mid$(strIn, 20, PlusPos - 20))
89                Case -1
                      'Example: "2021-08-23T08:47:20.920-05:00"
90                    MilliPart = CDbl(Mid$(strIn, 20, MinusPos - 20))
91            End Select
92        End If
          
93        LocalTime = CDate(Left$(strIn, 19)) + MilliPart / 86400

          Dim Adjust As Date
94        Select Case Sign
              Case 0
95                DtOut = LocalTime
96                Converted = True
97                Exit Sub
98            Case 1
99                If L <> PlusPos + 5 Then Exit Sub
100               Adjust = CDate(Right$(strIn, 5))
101               DtOut = LocalTime - Adjust
102               Converted = True
103           Case -1
104               If L <> MinusPos + 5 Then Exit Sub
105               Adjust = CDate(Right$(strIn, 5))
106               DtOut = LocalTime + Adjust
107               Converted = True
108       End Select

109       Exit Sub
ErrHandler:
          'Was not recognised as ISO8601 date
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseISO8601
' Purpose    : Test harness for calling from spreadsheets
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseISO8601(strIn As String) As Variant
          Dim Converted As Boolean
          Dim DtOut As Date

1         On Error GoTo ErrHandler
2         CastISO8601 strIn, DtOut, Converted, True, True

3         If Converted Then
4             ParseISO8601 = DtOut
5         Else
6             ParseISO8601 = "#Not recognised as ISO8601 date!"
7         End If
8         Exit Function
ErrHandler:
9         ReThrow "ParseISO8601", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISOZFormatString
' Purpose    : Returns the format string required to save datetimes with timezone under the assumton that the offset to
'              UTC is the same as the curent offset on this PC - use with care, Daylight saving may mean that that's not
'              a correct assumption for all the dates in a set of data...
' -----------------------------------------------------------------------------------------------------------------------
Private Function ISOZFormatString() As String
          Dim RightChars As String
          Dim TimeZone As String

1         On Error GoTo ErrHandler
2         TimeZone = GetLocalOffsetToUTC()

3         If TimeZone = 0 Then
4             RightChars = "Z"
5         ElseIf TimeZone > 0 Then
6             RightChars = "+" & Format$(TimeZone, "hh:mm")
7         Else
8             RightChars = "-" & Format$(Abs(TimeZone), "hh:mm")
9         End If
10        ISOZFormatString = "yyyy-mm-ddT:hh:mm:ss" & RightChars

11        Exit Function
ErrHandler:
12        ReThrow "ISOZFormatString", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeSentinels
' Purpose    : Returns a Dictionary keyed on strings for which if a key to the dictionary is a field of the CSV file then
'              that field should be converted to the associated item value. Handles Booleans, Missings and Excel errors.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub MakeSentinels(ByRef Sentinels As Scripting.Dictionary, ConvertQuoted As Boolean, Delimiter As String, ByRef MaxLength As Long, _
          ByRef AnySentinels As Boolean, ShowBooleansAsBooleans As Boolean, ShowErrorsAsErrors As Boolean, _
          ByRef ShowMissingsAs As Variant, Optional TrueStrings As Variant, Optional FalseStrings As Variant, _
          Optional MissingStrings As Variant)

          Const Err_FalseStrings As String = "FalseStrings must be omitted or provided as a string or an array of " & _
              "strings that represent Boolean value False"
          Const Err_MissingStrings As String = "MissingStrings must be omitted or provided a string or an array of " & _
              "strings that represent missing values"
          Const Err_ShowMissings As String = "ShowMissingsAs has an illegal value, such as an array or an object"
          Const Err_TrueStrings As String = "TrueStrings must be omitted or provided as string or an array of " & _
              "strings that represent Boolean value True"
          Const Err_TrueStrings2 As String = "TrueStrings has been provided, but type conversion for Booleans is " & _
              "not switched on for any column"
          Const Err_FalseStrings2 As String = "FalseStrings has been provided, but type conversion for Booleans " & _
              "is not switched on for any column"

1         On Error GoTo ErrHandler

2         If IsMissing(ShowMissingsAs) Then
3             ShowMissingsAs = Empty
4         ElseIf TypeName(ShowMissingsAs) = "Range" Then
5             ShowMissingsAs = ShowMissingsAs.value
6         End If
          
7         Select Case VarType(ShowMissingsAs)
              Case vbEmpty, vbString, vbBoolean, vbError, vbLong, vbInteger, vbSingle, vbDouble
8             Case Else
9                 Throw Err_ShowMissings
10        End Select
          
11        If Not IsMissing(MissingStrings) And Not IsEmpty(MissingStrings) Then
12            AddKeysToDict Sentinels, MissingStrings, ShowMissingsAs, Err_MissingStrings, "MissingString", Delimiter
13        End If

14        If ShowBooleansAsBooleans Then
15            If IsMissing(TrueStrings) Or IsEmpty(TrueStrings) Then
16                AddKeysToDict Sentinels, Array("TRUE", "true", "True"), True, Err_TrueStrings, "TrueString", Delimiter
17            Else
18                AddKeysToDict Sentinels, TrueStrings, True, Err_TrueStrings, "TrueString", Delimiter
19            End If
20            If IsMissing(FalseStrings) Or IsEmpty(FalseStrings) Then
21                AddKeysToDict Sentinels, Array("FALSE", "false", "False"), False, Err_FalseStrings, "FalseString", Delimiter
22            Else
23                AddKeysToDict Sentinels, FalseStrings, False, Err_FalseStrings, "FalseString", Delimiter
24            End If
25        Else
26            If Not (IsMissing(TrueStrings) Or IsEmpty(TrueStrings)) Then
27                Throw Err_TrueStrings2
28            End If
29            If Not (IsMissing(FalseStrings) Or IsEmpty(FalseStrings)) Then
30                Throw Err_FalseStrings2
31            End If
32        End If
          
33        If ShowErrorsAsErrors Then
34            AddKeyToDict Sentinels, "#DIV/0!", CVErr(xlErrDiv0)
35            AddKeyToDict Sentinels, "#NAME?", CVErr(xlErrName)
36            AddKeyToDict Sentinels, "#REF!", CVErr(xlErrRef)
37            AddKeyToDict Sentinels, "#NUM!", CVErr(xlErrNum)
38            AddKeyToDict Sentinels, "#NULL!", CVErr(xlErrNull)
39            AddKeyToDict Sentinels, "#N/A", CVErr(xlErrNA)
40            AddKeyToDict Sentinels, "#VALUE!", CVErr(xlErrValue)
41            AddKeyToDict Sentinels, "#SPILL!", CVErr(2045)
42            AddKeyToDict Sentinels, "#BLOCKED!", CVErr(2047)
43            AddKeyToDict Sentinels, "#CONNECT!", CVErr(2046)
44            AddKeyToDict Sentinels, "#UNKNOWN!", CVErr(2048)
45            AddKeyToDict Sentinels, "#GETTING_DATA!", CVErr(2043)
46            AddKeyToDict Sentinels, "#FIELD!", CVErr(2049)
47            AddKeyToDict Sentinels, "#CALC!", CVErr(2050)
48        End If

          'Add "quoted versions" of the existing sentinels
49        If ConvertQuoted Then
              Dim i As Long
              Dim items
              Dim Keys
              Dim NewKey As String
50            Keys = Sentinels.Keys
51            items = Sentinels.items
52            For i = LBound(Keys) To UBound(Keys)
53                NewKey = DQ & Replace$(Keys(i), DQ, DQ2) & DQ
54                AddKeyToDict Sentinels, NewKey, items(i)
55            Next i
56        End If

          Dim k As Variant
57        MaxLength = 0
58        For Each k In Sentinels.Keys
59            If Len(k) > MaxLength Then MaxLength = Len(k)
60        Next
61        AnySentinels = Sentinels.count > 0

62        Exit Sub
ErrHandler:
63        ReThrow "MakeSentinels", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddKeysToDict, Sub-routine of MakeSentinels
' Purpose    : Broadcast AddKeyToDict over an array of keys.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AddKeysToDict(ByRef Sentinels As Scripting.Dictionary, ByVal Keys As Variant, item As Variant, _
          FriendlyErrorString As String, KeyType As String, Delimiter As String)

          Dim i As Long
          Dim j As Long
        
1         On Error GoTo ErrHandler
        
2         If TypeName(Keys) = "Range" Then
3             Keys = Keys.value
4         End If
          
5         If VarType(Keys) = vbString Then
6             If InStr(Keys, ",") > 0 Then
7                 Keys = VBA.Split(Keys, ",")
8             End If
9         End If
          
10        Select Case NumDimensions(Keys)
              Case 0
11                ValidateCSVField CStr(Keys), KeyType, Delimiter
12                AddKeyToDict Sentinels, Keys, item, FriendlyErrorString
13            Case 1
14                For i = LBound(Keys) To UBound(Keys)
15                    ValidateCSVField CStr(Keys(i)), KeyType, Delimiter
16                    AddKeyToDict Sentinels, Keys(i), item, FriendlyErrorString
17                Next i
18            Case 2
19                For i = LBound(Keys, 1) To UBound(Keys, 1)
20                    For j = LBound(Keys, 2) To UBound(Keys, 2)
21                        ValidateCSVField CStr(Keys(i, j)), KeyType, Delimiter
22                        AddKeyToDict Sentinels, Keys(i, j), item, FriendlyErrorString
23                    Next j
24                Next i
25            Case Else
26                Throw FriendlyErrorString
27        End Select
28        Exit Sub
ErrHandler:
29        ReThrow "AddKeysToDict", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddKeyToDict, Sub-routine of MakeSentinels
' Purpose    : Wrap .Add method to have more helpful error message if things go awry.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AddKeyToDict(ByRef Sentinels As Scripting.Dictionary, Key As Variant, item As Variant, _
          Optional FriendlyErrorString As String)

          Dim FoundRepeated As Boolean

1         On Error GoTo ErrHandler

2         If VarType(Key) <> vbString Then Throw FriendlyErrorString & " but '" & CStr(Key) & "' is of type " & TypeName(Key)
          
3         If Len(Key) = 0 Then Exit Sub
          
4         If Not Sentinels.Exists(Key) Then
5             Sentinels.Add Key, item
6         Else
7             FoundRepeated = True
8             If VarType(item) = VarType(Sentinels.item(Key)) Then
9                 If item = Sentinels.item(Key) Then
10                    FoundRepeated = False
11                End If
12            End If
13        End If
          
14        If FoundRepeated Then
15            Throw "There is a conflicting definition of what the string '" & Key & _
                  "' should be converted to, both the " & TypeName(item) & " value '" & CStr(item) & _
                  "' and the " & TypeName(Sentinels.item(Key)) & " value '" & CStr(Sentinels.item(Key)) & _
                  "' have been specified. Please check the TrueStrings, FalseStrings and MissingStrings arguments"
16        End If

17        Exit Sub
ErrHandler:
18        ReThrow "AddKeyToDict", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseTextFile
' Purpose    : Convert a text file to a 2-dim array with one column, one line of file to one element of array, works for
'              files with any style of line endings - Windows, Mac, Unix, or a mixture of line endings.
' Parameters :
'  FileNameOrContents : FileName or CSV-style string.
'  isFile             : If True then fist argument is the name of a file, else it's a CSV-style string.
'  CharSet            : Used only if useADODB is True and isFile is True.
'  SkipToLine         : Return starts at this line of the file.
'  NumLinesToReturn   : This many lines are returned. Pass zero for all lines from SkipToLine.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseTextFile(FileNameOrContents As String, isFile As Boolean, _
          CharSet As String, SkipToLine As Long, NumLinesToReturn As Long, _
          CallingFromWorksheet As Boolean) As Variant

          Const Err_FileEmpty As String = "File is empty"
          Dim Buffer As String
          Dim BufferUpdatedTo As Long
          Dim FoundCR As Boolean
          Dim HaveReachedSkipToLine As Boolean
          Dim i As Long 'Index to read from Buffer
          Dim j As Long 'Index to write to Starts, Lengths
          Dim Err_StringTooLong As String
          Dim Lengths() As Long
          Dim MSLIA As Long
          Dim NumLinesFound As Long
          Dim PosCR As Long
          Dim PosLF As Long
          Dim ReturnArray() As String
          Dim SearchFor() As String
          Dim Starts() As Long
          Dim Stream As Object
          Dim Streaming As Boolean
          Dim Tmp As Long
          Dim Which As Long

1         On Error GoTo ErrHandler
          
2         If isFile Then

3             Set Stream = CreateObject("ADODB.Stream")
4             Stream.CharSet = CharSet
5             Stream.Open
6             Stream.LoadFromFile FileNameOrContents
7             If Stream.EOS Then Throw Err_FileEmpty
          
8             If NumLinesToReturn = 0 Then
9                 Buffer = ReadAllFromStream(Stream)
10                Streaming = False
11            Else
12                GetMoreFromStream Stream, vbNullString, vbNullString, Buffer, BufferUpdatedTo
13                Streaming = True
14            End If
15        Else
16            Buffer = FileNameOrContents
17            Streaming = False
18        End If
             
19        If Streaming Then
20            ReDim SearchFor(1 To 2)
21            SearchFor(1) = vbLf
22            SearchFor(2) = vbCr
23        End If

24        ReDim Starts(1 To 8): ReDim Lengths(1 To 8)
          
25        If Not Streaming Then
              'Ensure Buffer terminates with vbCrLf
26            If Right$(Buffer, 1) <> vbCr And Right$(Buffer, 1) <> vbLf Then
27                Buffer = Buffer & vbCrLf
28            ElseIf Right$(Buffer, 1) = vbCr Then
29                Buffer = Buffer & vbLf
30            End If
31            BufferUpdatedTo = Len(Buffer)
32        End If
          
33        NumLinesFound = 0
34        i = 0: j = 1
          
35        Starts(1) = i + 1
36        If SkipToLine = 1 Then HaveReachedSkipToLine = True

37        Do
38            If Not Streaming Then
39                If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
40                If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
41                If PosCR < PosLF Then
42                    FoundCR = True
43                    i = PosCR
44                Else
45                    FoundCR = False
46                    i = PosLF
47                End If
48            Else
49                i = SearchInBuffer(SearchFor, i + 1, Stream, vbNullString, _
                      vbNullString, Which, Buffer, BufferUpdatedTo)
50                FoundCR = (Which = 2)
51            End If

52            If i >= BufferUpdatedTo + 1 Then
53                Exit Do
54            End If

55            If j + 1 > UBound(Starts) Then
56                ReDim Preserve Starts(1 To UBound(Starts) * 2)
57                ReDim Preserve Lengths(1 To UBound(Lengths) * 2)
58            End If

59            Lengths(j) = i - Starts(j)
60            If FoundCR Then
61                If Mid$(Buffer, i + 1, 1) = vbLf Then
                      'Ending is Windows rather than Mac or Unix.
62                    i = i + 1
63                End If
64            End If
                          
65            Starts(j + 1) = i + 1
                          
66            j = j + 1
67            NumLinesFound = NumLinesFound + 1
68            If Not HaveReachedSkipToLine Then
69                If NumLinesFound = SkipToLine - 1 Then
70                    HaveReachedSkipToLine = True
71                    Tmp = Starts(j)
72                    ReDim Starts(1 To 8): ReDim Lengths(1 To 8)
73                    j = 1: NumLinesFound = 0
74                    Starts(1) = Tmp
75                End If
76            ElseIf NumLinesToReturn > 0 Then
77                If NumLinesFound = NumLinesToReturn Then
78                    Exit Do
79                End If
80            End If
81        Loop
         
82        If SkipToLine > NumLinesFound Then
83            If NumLinesToReturn = 0 Then 'Attempting to read from SkipToLine to the end of the file, but that would _
                                            be zero or a negative number of rows. So throw an error.
                                   
84                Throw "SkipToLine (" & CStr(SkipToLine) & ") exceeds the number of lines in the file (" & _
                      CStr(NumLinesFound) & ")"
85            Else
                  'Attempting to read a set number of rows, function will return an array of null strings
86                NumLinesFound = 0
87            End If
88        End If
89        If NumLinesToReturn = 0 Then NumLinesToReturn = NumLinesFound

90        ReDim ReturnArray(1 To NumLinesToReturn, 1 To 1)
91        MSLIA = MaxStringLengthInArray()
92        For i = 1 To MinLngs(NumLinesToReturn, NumLinesFound)
93            If CallingFromWorksheet Then
94                If Lengths(i) > MSLIA Then
95                    Err_StringTooLong = "Line " & Format$(i, "#,###") & " of the file is of length " & Format$(Lengths(i), "###,###")
96                    If MSLIA >= 32767 Then
97                        Err_StringTooLong = Err_StringTooLong & ". Excel cells cannot contain strings longer than " & Format$(MSLIA, "####,####")
98                    Else
99                        Err_StringTooLong = Err_StringTooLong & _
                              ". An array containing a string longer than " & Format$(MSLIA, "###,###") & _
                              " cannot be returned from VBA to an Excel worksheet"
100                   End If
101                   Throw Err_StringTooLong
102               End If
103           End If
104           ReturnArray(i, 1) = Mid$(Buffer, Starts(i), Lengths(i))
105       Next i

106       ParseTextFile = ReturnArray

107       Exit Function
ErrHandler:
108       ReThrow "ParseTextFile", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CSVWrite
' Purpose   : Creates a comma-separated file on disk containing Data. Any existing file of the same
'             name is overwritten. If successful, the function returns FileName, otherwise an "error
'             string" (starts with `#`, ends with `!`) describing what went wrong.
' Arguments
' Data      : An array of data, or an Excel range. Elements may be strings, numbers, dates, Booleans, empty,
'             Excel errors or null values. Data typically has two dimensions, but if Data has only one
'             dimension then the output file has a single column, one field per row.
' FileName  : The full name of the file, including the path. Alternatively, if FileName is omitted, then the
'             function returns Data converted CSV-style to a string.
' QuoteAllStrings: If `TRUE` (the default) then elements of Data that are strings are quoted before being
'             written to file, other elements (Numbers, Booleans, Errors) are not quoted. If `FALSE` then
'             the only elements of Data that are quoted are strings containing Delimiter, line feed,
'             carriage return or double quote. In both cases, double quotes are escaped by another double
'             quote. If "Raw" then no strings are quoted. Use this option with care, the file written may
'             not be in valid CSV format.
' DateFormat: A format string that determines how dates, including cells formatted as dates, appear in the
'             file. If omitted, defaults to `yyyy-mm-dd`.
' DateTimeFormat: Format for datetimes. Defaults to `ISO` which abbreviates `yyyy-mm-ddThh:mm:ss`. Use
'             `ISOZ` for ISO8601 format with time zone the same as the PC's clock. Use with care, daylight
'             saving may be inconsistent across the datetimes in data.
' Delimiter : The delimiter string, if omitted defaults to a comma. Delimiter may have more than one
'             character.
' Encoding  : Allowed entries are `ANSI` (the default), `UTF-8`, `UTF-16`, `UTF-8NOBOM` and `UTF-16NOBOM`.
'             An error will result if this argument is `ANSI` but Data contains characters with code point
'             above 127.
' EOL       : Sets the file's line endings. Enter `Windows`, `Unix` or `Mac`. Also supports the line-ending
'             characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13) or the strings `CRLF`, `LF`
'             or `CR`. The default is `Windows` if FileName is provided, or `Unix` if not. The last line of
'             the file is written with a line ending.
' TrueString: How the Boolean value True is to be represented in the file. Optional, defaulting to "True".
' FalseString: How the Boolean value False is to be represented in the file. Optional, defaulting to
'             "False".
'
' Notes     : See also companion function CSVRead.
'
'             For discussion of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
' -----------------------------------------------------------------------------------------------------------------------
Public Function CSVWrite(ByVal Data As Variant, Optional ByVal FileName As String, _
          Optional ByVal QuoteAllStrings As Variant = True, Optional ByVal DateFormat As String = "YYYY-MM-DD", _
          Optional ByVal DateTimeFormat As String = "ISO", Optional ByVal Delimiter As String = ",", _
          Optional ByVal Encoding As String = "ANSI", Optional ByVal EOL As String = vbNullString, _
          Optional TrueString As String = "True", Optional FalseString As String = "False") As String
Attribute CSVWrite.VB_Description = "Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an ""error string"" (starts with `#`, ends with `!`) describing what went wrong."
Attribute CSVWrite.VB_ProcData.VB_Invoke_Func = " \n14"

          Const Err_Delimiter1 = "Delimiter must have at least one character"
          Const Err_Delimiter2 As String = "Delimiter cannot start with a " & _
              "double quote, line feed or carriage return"
              
          Const Err_Dimensions As String = "Data has more than two dimensions, which is not supported"
          
          Dim Encoder As Scripting.Dictionary
          Dim EOLIsWindows As Boolean
          Dim FileContents As String
          Dim i As Long
          Dim j As Long
          Dim Lines() As String
          Dim OneLine() As String
          Dim WriteToFile As Boolean
          Dim QuoteSimpleStrings As Boolean
          Dim QuoteComplexStrings As Boolean

1         On Error GoTo ErrHandler
          
2         WriteToFile = Len(FileName) > 0
                    
3         If Len(Delimiter) = 0 Then
4             Throw Err_Delimiter1
5         End If
6         If Left$(Delimiter, 1) = DQ Or Left$(Delimiter, 1) = vbLf Or Left$(Delimiter, 1) = vbCr Then
7             Throw Err_Delimiter2
8         End If
          
9         ParseQuoteAllStrings QuoteAllStrings, QuoteSimpleStrings, QuoteComplexStrings
          
10        ValidateTrueAndFalseStrings TrueString, FalseString, Delimiter

11        WriteToFile = Len(FileName) > 0

12        If EOL = vbNullString Then
13            If WriteToFile Then
14                EOL = vbCrLf
15            Else
16                EOL = vbLf
17            End If
18        End If

19        EOL = OStoEOL(EOL, "EOL")
20        EOLIsWindows = EOL = vbCrLf
          
21        If DateFormat = "" Or UCase(DateFormat) = "ISO" Then
              'Avoid DateFormat being the null string as that would make CSVWrite's _
               behaviour depend on Windows locale (via calls to Format$ in function Encode).
22            DateFormat = "yyyy-mm-dd"
23        End If
          
24        Select Case UCase$(DateTimeFormat)
              Case "ISO", ""
25                DateTimeFormat = "yyyy-mm-ddThh:mm:ss"
26            Case "ISOZ"
27                DateTimeFormat = ISOZFormatString()
28        End Select

29        If TypeName(Data) = "Range" Then
              'Preserve elements of type Date by using .Value, not .Value2
30            Data = Data.value
31        End If
32        Select Case NumDimensions(Data)
              Case 0
                  Dim Tmp() As Variant
33                ReDim Tmp(1 To 1, 1 To 1)
34                Tmp(1, 1) = Data
35                Data = Tmp
36            Case 1
37                ReDim Tmp(LBound(Data) To UBound(Data), 1 To 1)
38                For i = LBound(Data) To UBound(Data)
39                    Tmp(i, 1) = Data(i)
40                Next i
41                Data = Tmp
42            Case Is > 2
43                Throw Err_Dimensions
44        End Select

45        Set Encoder = MakeEncoder(TrueString, FalseString)

46        ReDim OneLine(LBound(Data, 2) To UBound(Data, 2))
47        ReDim Lines(LBound(Data) To UBound(Data) + 1) 'add one to ensure that result has a terminating EOL
        
48        For i = LBound(Data) To UBound(Data)
49            For j = LBound(Data, 2) To UBound(Data, 2)
50                OneLine(j) = Encode(Data(i, j), QuoteSimpleStrings, QuoteComplexStrings, DateFormat, DateTimeFormat, Delimiter, Encoder)
51            Next j
52            Lines(i) = VBA.Join(OneLine, Delimiter)
53        Next i
54        FileContents = VBA.Join(Lines, EOL)
          
55        If WriteToFile Then
56            CSVWrite = SaveTextFile(FileName, FileContents, Encoding)
57        Else
58            If Len(FileContents) > MaxStringLengthInArray() Then
59                If TypeName(Application.Caller) = "Range" Then
60                    Throw "Cannot return string of length " & Format$(CStr(Len(FileContents)), "#,###") & _
                          " to a cell of an Excel worksheet"
61                End If
62            End If
63            CSVWrite = FileContents
64        End If
          
65        Exit Function
ErrHandler:
66        CSVWrite = ReThrow("CSVWrite", Err, m_ErrorStyle = es_ReturnString)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveTextFile
' Author     : Philip Swannell
' Date       : 28-Nov-2022
' Purpose    : Save a text file, overwriting any existing file of the same name. For ANSI encoding uses
'              Scripting.FileStream so that an error occurs if FileContents contains unsupported characters. Alternative
'              would be to use ADODB.Stream with CharSet = "windows-1252" but in that case unsupported characters get
'              written to file as question marks, resulting in loss of data.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SaveTextFile(FileName As String, FileContents As String, Encoding As String)

          Const adModeReadWrite = 3
          Const adSaveCreateOverWrite = 2
          Const adTypeBinary = 1
          Const adTypeText = 2
          Dim BOMLength As Long
          Dim CharSet As String
          Dim iStr As Object
          Dim oStr As Object
          Dim Stream As Object
                    
          Const Err_Encoding As String = "Encoding must be ""ANSI"" (the default), ""UTF-8"", ""UTF-16"", ""UTF-8NOBOM"" or ""UTF-16NOBOM"""

1         On Error GoTo ErrHandler
          'Ignore hyphens and spaces
2         Select Case UCase$(Replace$(Replace$(Encoding, "-", vbNullString), " ", vbNullString))
              Case "UTF8", "UTF16", "UTF8BOM", "UTF16BOM"
3                 Set Stream = CreateObject("ADODB.Stream")
4                 CharSet = IIf(InStr(Encoding, "8") > 0, "utf-8", "utf-16")
5                 Stream.Open
6                 Stream.Type = adTypeText
7                 Stream.CharSet = CharSet
8                 Stream.WriteText FileContents
9                 Stream.SaveToFile FileName, adSaveCreateOverWrite
10                Stream.Close: Set Stream = Nothing
11                SaveTextFile = FileName

12            Case "UTF8NOBOM", "UTF16NOBOM"
                  ' Adapted from https://stackoverflow.com/questions/52339439/how-to-create-utf-16-file-in-vbscript

13                If UCase(Encoding) = "UTF-8NOBOM" Then
14                    CharSet = "utf-8"
15                    BOMLength = 3
16                Else
17                    CharSet = "utf-16"
18                    BOMLength = 2
19                End If

20                Set iStr = CreateObject("ADODB.Stream")
21                Set oStr = CreateObject("ADODB.Stream")

                  ' one stream for converting the text to UTF bytes
22                iStr.Mode = adModeReadWrite
23                iStr.Type = adTypeText
24                iStr.CharSet = CharSet
25                iStr.Open
26                iStr.WriteText FileContents

                  ' one steam to write bytes to a file
27                oStr.Mode = adModeReadWrite
28                oStr.Type = adTypeBinary
29                oStr.Open

                  ' switch first stream to binary mode and skip BOM
30                iStr.Position = 0
31                iStr.Type = adTypeBinary
32                iStr.Position = BOMLength

                  ' write remaining bytes to file and clean up
33                oStr.Write iStr.Read
34                oStr.SaveToFile FileName, adSaveCreateOverWrite
35                oStr.Close
36                iStr.Close
37                SaveTextFile = FileName

38            Case "ANSI", ""
39                If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
                  Dim ED As String
                  Dim EN As Long
40                On Error Resume Next
41                Set Stream = m_FSO.CreateTextFile(FileName, True, False)
42                EN = Err.Number: ED = Err.Description
43                On Error GoTo ErrHandler
44                If EN <> 0 Then Throw "Error '" & ED & "' when attempting to create file '" + FileName + "'"
45                WriteWrap Stream, FileContents
46                Stream.Close: Set Stream = Nothing
47                SaveTextFile = FileName
48            Case Else
49                Throw Err_Encoding
50        End Select

51        Exit Function
ErrHandler:
52        If Not Stream Is Nothing Then
53            Stream.Close
54            Set Stream = Nothing
55        End If
56        ReThrow "SaveTextFile", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ValidateTrueAndFalseStrings
' Purpose    : Stop the user from making bad choices for either TrueString or FalseString, e.g: strings that would be
'              interpreted as (the wrong) Boolean, or as numbers, dates or empties, strings containing line feed
'              characters, containing the delimiter etc.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ValidateTrueAndFalseStrings(TrueString As String, FalseString As String, Delimiter As String)
             
1         On Error GoTo ErrHandler
2         If LCase$(TrueString) = "true" Then
3             If LCase$(FalseString) = "false" Then
4                 Exit Function
5             End If
6         End If
          
7         If LCase$(TrueString) = "false" Then Throw "TrueString cannot take the value '" & TrueString & "'"
8         If LCase$(FalseString) = "true" Then Throw "FalseString cannot take the value '" & FalseString & "'"

9         If TrueString = FalseString Then
10            Throw "Got '" & TrueString & "' for both TrueString and FalseString, but these cannot be equal to one another"
11        End If
          
12        ValidateBooleanRepresentation TrueString, "TrueString", Delimiter
13        ValidateBooleanRepresentation FalseString, "FalseString", Delimiter

14        Exit Function
ErrHandler:
15        ReThrow "ValidateTrueAndFalseStrings", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ValidateBooleanRepresentation
' Purpose    : Stop the user from making bad choices for either TrueString or FalseString, e.g: strings that would be
'              interpreted as (the wrong) Boolean, or as numbers, dates or empties, strings containing line feed
'              characters, containing the delimiter etc.
' Parameters :
'  strValue : The string chosen e.g. "TRUE" or "VRAI" or "Yes"
'  strName  : Either the string "TrueString" or the string "FalseString", used for error message generation
'  Delimiter: The delimiter character used in the file.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ValidateBooleanRepresentation(strValue As String, strName As String, Delimiter As String)
          
          Dim Converted As Boolean
          Dim DateSeparator As Variant
          Dim DQCount As Long
          Dim DtOut As Date
          Dim i As Long
          Dim SysDateSeparator As String
              
1         On Error GoTo ErrHandler
2         SysDateSeparator = Application.International(xlDateSeparator)

3         If strValue = "" Then Throw strName & " cannot be the zero-length string"

4         If InStr(strValue, vbLf) > 0 Then Throw strName & " contains a line feed character (ascii 10), which is not permitted"
5         If InStr(strValue, vbCr) > 0 Then Throw strName & " contains a carriage return character (ascii 13), which is not permitted"
6         If InStr(strValue, Delimiter) > 0 Then Throw strName & " contains Delimiter '" & Delimiter & "' which is not permitted"
7         If InStr(strValue, DQ) > 0 Then
8             DQCount = Len(strValue) - Len(Replace$(strValue, DQ, vbNullString))
9             If DQCount <> 2 Or Left$(strValue, 1) <> DQ Or Right$(strValue, 1) <> DQ Then
10                Throw "When " & strName & " contains any double quote characters they must be at the start, the end and nowhere else"
11            End If
12        End If
              
13        If IsNumeric(strValue) Then Throw "Got '" & strValue & "' as " & strName & " but that's not valid because it represents a number"
              
14        For i = 0 To 2
15            For Each DateSeparator In Array("/", "-", " ")
                  Converted = False
16                CastToDate strValue, DtOut, i, _
                      CStr(DateSeparator), SysDateSeparator, Converted
17                If Converted Then
18                    Throw "Got '" & strValue & "' as " & _
                          strName & " but that's not valid because it represents a date"
19                End If
20            Next
21        Next

22        Exit Function
ErrHandler:
23        ReThrow "ValidateBooleanRepresentation", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ValidateCSVField
' Purpose    : Throw an error if FieldValue could not be a field of a CSV file, e.g. contains a line feed character but
'              does not start and end with double quotes.
' Parameters :
'  FieldValue : The field value to be validated
'  FieldName  : The name of the field - used for constructing the error description
'  Delimiter: The delimiter in use in the CSV
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ValidateCSVField(FieldValue As String, FieldName As String, Delimiter As String)

          Dim DQsGood As Boolean
          Dim HasCR As Boolean
          Dim HasDelim As Boolean
          Dim HasDQ As Boolean
          Dim HasLF As Boolean
          Dim InnerPart As String

1         On Error GoTo ErrHandler
2         If Len(Delimiter) > 0 Then
3             HasDelim = InStr(FieldValue, Delimiter) > 0
4         End If
5         HasCR = InStr(FieldValue, vbCr) > 0
6         HasLF = InStr(FieldValue, vbLf) > 0
7         HasDQ = InStr(FieldValue, DQ) > 0

8         If Not (HasDelim Or HasCR Or HasLF Or HasDQ) Then
9             Exit Sub
10        End If

11        If HasDQ Then
12            DQsGood = True
13            If Left$(FieldValue, 1) <> DQ Then
14                DQsGood = False
15            ElseIf Right$(FieldValue, 1) <> DQ Then
16                DQsGood = False
17            Else
18                If Len(FieldValue) < 2 Then
19                    DQsGood = False
20                Else
21                    InnerPart = Mid$(FieldValue, 2, Len(FieldValue) - 2)
22                    If InStr(InnerPart, DQ) > 0 Then
23                        If Len(Replace$(InnerPart, DQ & DQ, "")) <> Len(Replace$(InnerPart, DQ, "")) Then
24                            DQsGood = False
25                        End If
26                    End If
27                End If
28            End If
29        End If

30        If HasCR Or HasLF Or HasDelim Or HasDQ Then
31            If Not DQsGood Then
32                Throw "Got '" & Replace$(Replace$(FieldValue, vbCr, "<CR>"), vbLf, "<LF>") & "' as " & _
                      FieldName & ", but that cannot be a field in a CSV file, since it is not correctly quoted"
33            End If
34        End If

35        Exit Sub
ErrHandler:
36        ReThrow "ValidateCSVField", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OStoEOL
' Purpose    : Convert text describing an operating system to the end-of-line marker employed. Note that "Mac" converts
'              to vbCr but Apple operating systems since OSX use vbLf, matching Unix.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OStoEOL(OS As String, ArgName As String) As String

          Const Err_Invalid As String = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associated end of line characters"

1         On Error GoTo ErrHandler
2         Select Case LCase$(OS)
              Case "windows", vbCrLf, "crlf"
3                 OStoEOL = vbCrLf
4             Case "unix", "linux", vbLf, "lf"
5                 OStoEOL = vbLf
6             Case "mac", vbCr, "cr"
7                 OStoEOL = vbCr
8             Case Else
9                 Throw ArgName & Err_Invalid
10        End Select

11        Exit Function
ErrHandler:
12        ReThrow "OStoEOL", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeEncoder
' Author     : Philip Swannell
' Date       : 26-Nov-2022
' Purpose    : Returns a Dictionary for use in method Encode and which contains (as keys) the sentinel values (Booleans,
'              null, errors) and (as values) the strings to which those sentinels are converted.
'              This arrangement means that Excel errors are converted to their English-language representation even when
'              the Excel display language is not Excel.
' -----------------------------------------------------------------------------------------------------------------------
Private Function MakeEncoder(TrueString As String, FalseString As String) As Scripting.Dictionary

          Dim d As New Scripting.Dictionary
1         On Error GoTo ErrHandler
2         d.Add True, TrueString
3         d.Add False, FalseString
4         d.Add Null, "NULL"
5         d.Add CVErr(2000), "#NULL!"
6         d.Add CVErr(2007), "#DIV/0!"
7         d.Add CVErr(2015), "#VALUE!"
8         d.Add CVErr(2023), "#REF!"
9         d.Add CVErr(2029), "#NAME?"
10        d.Add CVErr(2036), "#NUM!"
11        d.Add CVErr(2042), "#N/A"
12        d.Add CVErr(2043), "#GETTING_DATA!"
13        d.Add CVErr(2045), "#SPILL!"
14        d.Add CVErr(2046), "#CONNECT!"
15        d.Add CVErr(2047), "#BLOCKED!"
16        d.Add CVErr(2048), "#UNKNOWN!"
17        d.Add CVErr(2049), "#FIELD!"
18        d.Add CVErr(2050), "#CALC!"
19        Set MakeEncoder = d

20        Exit Function
ErrHandler:
21        ReThrow "MakeEncoder", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Encode
' Purpose    : Encode arbitrary value as a string, sub-routine of CSVWrite.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Encode(ByVal x As Variant, QuoteSimpleStrings As Boolean, QuoteComplexStrings As Boolean, DateFormat As String, _
          DateTimeFormat As String, Delim As String, Encoder As Scripting.Dictionary) As String
          
1         On Error GoTo ErrHandler
2         Select Case VarType(x)

              Case vbString
                'We do not handle case QuoteSimpleStrings = TRUE and QuoteComplexStrings = FALSE as that case is never encountered
3                 If Not QuoteComplexStrings Then
4                     Encode = x
5                 ElseIf InStr(x, DQ) > 0 Then
6                     Encode = DQ & Replace$(x, DQ, DQ2) & DQ
7                 ElseIf QuoteSimpleStrings Then
8                     Encode = DQ & x & DQ
9                 ElseIf InStr(x, vbCr) > 0 Then
10                    Encode = DQ & x & DQ
11                ElseIf InStr(x, vbLf) > 0 Then
12                    Encode = DQ & x & DQ
13                ElseIf InStr(x, Delim) > 0 Then
14                    Encode = DQ & x & DQ
15                Else
16                    Encode = x
17                End If
18            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbEmpty, 20 '20 = vbLongLong - not available on 32 bit.
19                Encode = CStr(x)
20            Case vbBoolean, vbError, vbNull
21                Encode = Encoder(x)
22            Case vbDate
23                If CLng(x) = x Then
24                    Encode = Format$(x, DateFormat)
25                Else
26                    Encode = Format$(x, DateTimeFormat)
27                End If
28            Case Else
29                Throw "Cannot convert variant of type " & TypeName(x) & " to String"
30        End Select
31        Exit Function
ErrHandler:
32        ReThrow "Encode", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WriteWrap
' Purpose    : Wrapper to TextStream.Write[Line] to give more informative error message than "invalid procedure call or
'              argument" if the error is caused by attempting to write illegal characters to a stream opened with
'              TriStateFalse.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub WriteWrap(T As TextStream, text As String)

          Dim ErrDesc As String
          Dim ErrNum As Long
          Dim i As Long

1         On Error GoTo ErrHandler
2         T.Write text
3         Exit Sub

ErrHandler:
4         ErrNum = Err.Number
5         If ErrNum = 5 Then
6             For i = 1 To Len(text)
7                 If Not CanWriteCharToAscii(Mid$(text, i, 1)) Then
8                     ErrDesc = "Data contains characters that cannot be written to an ascii file (first found is '" & _
                          Mid$(text, i, 1) & "' with unicode character code " & AscW(Mid$(text, i, 1)) & _
                          "). Try calling CSVWrite with argument Encoding as ""UTF-8"" or ""UTF-16"""
9                     Throw ErrDesc
10                    Exit For
11                End If
12            Next i
13        End If
14        ReThrow "WriteWrap", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CanWriteCharToAscii
' Purpose    : Not all characters for which AscW(c) < 255 can be written to an ascii file. If AscW(c) is in the following
'              list then they cannot:
'             128,130,131,132,133,134,135,136,137,138,139,140,142,145,146,147,148,149,150,151,152,153,154,155,156,158,159
' -----------------------------------------------------------------------------------------------------------------------
Private Function CanWriteCharToAscii(c As String) As Boolean
          Dim code As Long
1         code = AscW(c)
2         If code > 255 Or code < 0 Then
3             CanWriteCharToAscii = False
4         Else
5             CanWriteCharToAscii = Chr$(AscW(c)) = c
6         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MaxStringLengthInArray
' Purpose    : Different versions of Excel have different limits for the longest string that can be an element of an
'              array passed from a VBA UDF back to Excel. I believe the limit is 255 for Excel 2013 and earlier
'              and 32,767 later versions of Excel including Excel 365.
' -----------------------------------------------------------------------------------------------------------------------
Private Function MaxStringLengthInArray() As Long
          Static Res As Long
1         If Res = 0 Then
2             Select Case Val(Application.Version)
                  Case Is <= 15 'Excel 2013 and earlier
3                     Res = 255
4                 Case Else
5                     Res = 32767
6             End Select
7         End If
8         MaxStringLengthInArray = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Download
' Purpose   : Downloads bits from the Internet and saves them to a file.
'             See https://msdn.microsoft.com/en-us/library/ms775123(v=vs.85).aspx
' -----------------------------------------------------------------------------------------------------------------------
Private Function Download(URLAddress As String, ByVal FileName As String) As String
          Dim EN As Long
          Dim ErrString As String
          Dim Res As Long
          Dim TargetFolder As String

1         On Error GoTo ErrHandler
          
2         TargetFolder = FileFromPath(FileName, False)
3         CreatePath TargetFolder
4         If FileExists(FileName) Then
5             On Error Resume Next
6             FileDelete FileName
7             EN = Err.Number
8             On Error GoTo ErrHandler
9             If EN <> 0 Then
10                Throw "Cannot download from URL '" & URLAddress & "' because target file '" & FileName & _
                      "' already exists and cannot be deleted. Is the target file open in a program such as Excel?"
11            End If
12        End If
          
13        Res = URLDownloadToFile(0, URLAddress, FileName, 0, 0)
14        If Res <> 0 Then
15            ErrString = ParseDownloadError(CLng(Res))
16            Throw "Windows API function URLDownloadToFile returned error code " & CStr(Res) & _
                  " with description '" & ErrString & "'"
17        End If
18        If Not FileExists(FileName) Then Throw "Windows API function URLDownloadToFile did not report an error, " & _
              "but appears to have not successfuly downloaded a file from " & URLAddress & " to " & FileName
              
19        Download = FileName

20        Exit Function
ErrHandler:
21        ReThrow "Download", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDownloadError, sub of Download
'              https://www.vbforums.com/showthread.php?882757-URLDownloadToFile-error-codes
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseDownloadError(ErrNum As Long) As String
          Dim ErrString As String
1         Select Case ErrNum
              Case &H80004004
2                 ErrString = "Aborted"
3             Case &H800C0001
4                 ErrString = "Destination File Exists"
5             Case &H800C0002
6                 ErrString = "Invalid Url"
7             Case &H800C0003
8                 ErrString = "No Session"
9             Case &H800C0004
10                ErrString = "Cannot Connect"
11            Case &H800C0005
12                ErrString = "Resource Not Found"
13            Case &H800C0006
14                ErrString = "Object Not Found"
15            Case &H800C0007
16                ErrString = "Data Not Available"
17            Case &H800C0008
18                ErrString = "Download Failure"
19            Case &H800C0009
20                ErrString = "Authentication Required"
21            Case &H800C000A
22                ErrString = "No Valid Media"
23            Case &H800C000B
24                ErrString = "Connection Timeout"
25            Case &H800C000C
26                ErrString = "Invalid Request"
27            Case &H800C000D
28                ErrString = "Unknown Protocol"
29            Case &H800C000E
30                ErrString = "Security Problem"
31            Case &H800C000F
32                ErrString = "Cannot Load Data"
33            Case &H800C0010
34                ErrString = "Cannot Instantiate Object"
35            Case &H800C0014
36                ErrString = "Redirect Failed"
37            Case &H800C0015
38                ErrString = "Redirect To Dir"
39            Case &H800C0016
40                ErrString = "Cannot Lock Request"
41            Case Else
42                ErrString = "Unknown"
43        End Select
44        ParseDownloadError = ErrString
End Function

Private Function GetFileSize(FilePath As String)

1         On Error GoTo ErrHandler1
2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
3         GetFileSize = m_FSO.GetFile(FilePath).Size

4         On Error GoTo ErrHandler2
5         If GetFileSize >= (2 ^ 31) Then
6             Throw "File is too large. It is " & Format(GetFileSize, "###,###") & _
                  " bytes, which exceeds the maximum allowed size of " & Format((2 ^ 31) - 1, "###,###")
7         End If
8         Exit Function

ErrHandler1:
9         Throw "Could not find file '" & FilePath & "'"

ErrHandler2:
10        ReThrow "GetFileSize", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileExists
' Purpose    : Returns True if FileName exists on disk, False o.w.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileExists(FileName As String) As Boolean
          Dim F As Scripting.File
1         On Error GoTo ErrHandler
2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
3         Set F = m_FSO.GetFile(FileName)
4         FileExists = True
5         Exit Function
ErrHandler:
6         FileExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FolderExists
' Purpose   : Returns True or False. Does not matter if FolderPath has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FolderExists(FolderPath As String) As Boolean
          Dim F As Scripting.Folder
          
1         On Error GoTo ErrHandler
2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
          
3         Set F = m_FSO.GetFolder(FolderPath)
4         FolderExists = True
5         Exit Function
ErrHandler:
6         FolderExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileDelete
' Purpose    : Delete a file, returns True or error string.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub FileDelete(FileName As String)
          Dim F As Scripting.File
1         On Error GoTo ErrHandler

2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
3         Set F = m_FSO.GetFile(FileName)
4         F.Delete

5         Exit Sub
ErrHandler:
6         ReThrow "FileDelete", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreatePath
' Purpose   : Creates a folder on disk. FolderPath can be passed in as C:\This\That\TheOther even if the
'             folder C:\This does not yet exist. If successful returns the name of the
'             folder. If not successful throws an error.
' Arguments
' FolderPath: Path of the folder to be created. For example C:\temp\My_New_Folder. It does not matter if
'             this path has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CreatePath(ByVal FolderPath As String)

          Dim F As Scripting.Folder
          Dim i As Long
          Dim ParentFolderName As String
          Dim ThisFolderName As String

1         On Error GoTo ErrHandler

2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject

3         If Left$(FolderPath, 2) = "\\" Then
4         ElseIf Mid$(FolderPath, 2, 2) <> ":\" Or _
              Asc(UCase$(Left$(FolderPath, 1))) < 65 Or _
              Asc(UCase$(Left$(FolderPath, 1))) > 90 Then
5             Throw "First three characters of FolderPath must give drive letter followed by "":\"" or else be""\\"" for " & _
                  "UNC folder name"
6         End If

7         FolderPath = Replace$(FolderPath, "/", "\")

8         If Right$(FolderPath, 1) <> "\" Then
9             FolderPath = FolderPath & "\"
10        End If

11        If FolderExists(FolderPath) Then
12            GoTo EarlyExit
13        End If

          'Go back until we find parent folder that does exist
14        For i = Len(FolderPath) - 1 To 3 Step -1
15            If Mid$(FolderPath, i, 1) = "\" Then
16                If FolderExists(Left$(FolderPath, i)) Then
17                    Set F = m_FSO.GetFolder(Left$(FolderPath, i))
18                    ParentFolderName = Left$(FolderPath, i)
19                    Exit For
20                End If
21            End If
22        Next i

23        If F Is Nothing Then Throw "Cannot create folder " & Left$(FolderPath, 3)

          'now add folders one level at a time
24        For i = Len(ParentFolderName) + 1 To Len(FolderPath)
25            If Mid$(FolderPath, i, 1) = "\" Then
                  
26                ThisFolderName = Mid$(FolderPath, InStrRev(FolderPath, "\", i - 1) + 1, _
                      i - 1 - InStrRev(FolderPath, "\", i - 1))
27                F.SubFolders.Add ThisFolderName
28                Set F = m_FSO.GetFolder(Left$(FolderPath, i))
29            End If
30        Next i

EarlyExit:
31        Set F = m_FSO.GetFolder(FolderPath)
32        Set F = Nothing

33        Exit Sub
ErrHandler:
34        ReThrow "CreatePath", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileFromPath
' Purpose    : Split file-with-path to file name (if ReturnFileName is True) or path otherwise.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileFromPath(FullFileName As String, Optional ReturnFileName As Boolean = True) As Variant
          Dim SlashPos As Long
          Dim SlashPos2 As Long

1         On Error GoTo ErrHandler

2         SlashPos = InStrRev(FullFileName, "\")
3         SlashPos2 = InStrRev(FullFileName, "/")
4         If SlashPos2 > SlashPos Then SlashPos = SlashPos2
5         If SlashPos = 0 Then Throw "Neither '\' nor '/' found"

6         If ReturnFileName Then
7             FileFromPath = Mid$(FullFileName, SlashPos + 1)
8         Else
9             FileFromPath = Left$(FullFileName, SlashPos - 1)
10        End If

11        Exit Function
ErrHandler:
12        ReThrow "FileFromPath", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NumDimensions
' Purpose   : Returns the number of dimensions in an array variable, or 0 if the variable
'             is not an array.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NumDimensions(x As Variant) As Long
          Dim i As Long
          Dim y As Long
1         If Not IsArray(x) Then
2             NumDimensions = 0
3             Exit Function
4         Else
5             On Error GoTo ExitPoint
6             i = 1
7             Do While True
8                 y = LBound(x, i)
9                 i = i + 1
10            Loop
11        End If
ExitPoint:
12        NumDimensions = i - 1
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OneDArrayToTwoDArray
' Purpose    : Convert 1-d array of 2-element 1-d arrays into a 1-based, two-column, 2-d array.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OneDArrayToTwoDArray(x As Variant) As Variant
          Const Err_1DArray As String = "If ConvertTypes is given as a 1-dimensional array, each element must " & _
              "be a 1-dimensional array with two elements"

          Dim i As Long
          Dim k As Long
          Dim TwoDArray() As Variant
          
1         On Error GoTo ErrHandler
2         ReDim TwoDArray(1 To UBound(x) - LBound(x) + 1, 1 To 2)
3         For i = LBound(x) To UBound(x)
4             k = k + 1
5             If Not IsArray(x(i)) Then Throw Err_1DArray
6             If NumDimensions(x(i)) <> 1 Then Throw Err_1DArray
7             If UBound(x(i)) - LBound(x(i)) <> 1 Then Throw Err_1DArray
8             TwoDArray(k, 1) = x(i)(LBound(x(i)))
9             TwoDArray(k, 2) = x(i)(1 + LBound(x(i)))
10        Next i
11        OneDArrayToTwoDArray = TwoDArray
12        Exit Function
ErrHandler:
13        ReThrow "OneDArrayToTwoDArray", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetLocalOffsetToUTC
' Purpose    : Get the PC's offset to UTC.
' See "gogeek"'s post at
' https://stackoverflow.com/questions/1600875/how-to-get-the-current-datetime-in-utc-from-an-excel-vba-macro
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetLocalOffsetToUTC() As Double
          Dim dt As Object
          Dim TimeNow As Date
          Dim UTC As Date
1         On Error GoTo ErrHandler
2         TimeNow = Now()

3         Set dt = CreateObject("WbemScripting.SWbemDateTime")
4         dt.SetVarDate TimeNow
5         UTC = dt.GetVarDate(False)
6         GetLocalOffsetToUTC = (TimeNow - UTC)

7         Exit Function
ErrHandler:
8         ReThrow "GetLocalOffsetToUTC", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FunctionWizardActive
' Purpose    : Test if Excel's Function Wizard is active to allow early exit in slow functions.
' https://stackoverflow.com/questions/20866484/can-i-disable-a-vba-udf-calculation-when-the-insert-function-function-arguments
' -----------------------------------------------------------------------------------------------------------------------
Private Function FunctionWizardActive() As Boolean
          
1         On Error GoTo ErrHandler
2         If Not Application.CommandBars.item("Standard").Controls.item(1).Enabled Then
3             FunctionWizardActive = True
4         End If

5         Exit Function
ErrHandler:
6         ReThrow "FunctionWizardActive", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Throw
' Purpose    : Error handling - companion to ReThrow
' Parameters :
'  Description  : Description of what went wrong.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Throw(ByVal Description As String)
1         Err.Raise vbObjectError + 1, "Throw", Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReThrow
' Purpose    : Common error handling to be used in the error handler of all methods.
' Parameters :
'  FunctionName: The name of the function from which ReThrow is called, typically in the function's error handler.
'  ReturnString: Pass in True if the method is a "top level" method that's exposed to the user and we wish for the
'                function to return an error string (starts with #, ends with !).
'                Pass in False if we want to (re)throw an error, with annotated Description.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ReThrow(FunctionName As String, Error As ErrObject, Optional ReturnString As Boolean = False)

          Dim ErrorDescription As String
          Dim ErrorNumber As Long
          Dim LineDescription As String
          
1         ErrorDescription = Error.Description
2         ErrorNumber = Error.Number

          'Build up call stack, i.e. annotate error description by prepending #<FunctionName> and appending !
3         If m_ErrorStringsEmbedCallStack Then
4             If Erl = 0 Then
5                 LineDescription = " (line unknown): "
6             Else
7                 LineDescription = " (line " & CStr(Erl) & "): "
8             End If
9         Else
10            LineDescription = ": "
11        End If

12        If ReturnString Or m_ErrorStringsEmbedCallStack Then
13            ErrorDescription = "#" & FunctionName & LineDescription & ErrorDescription & "!"
14        End If

15        If ReturnString Then
16            ReThrow = ErrorDescription
17        Else
18            Err.Raise ErrorNumber, , ErrorDescription
19        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ThrowIfError
' Purpose   : In the event of an error, methods intended to be callable from spreadsheets
'             return an error string (starts with "#", ends with "!"). ThrowIfError allows such
'             methods to be used from VBA code while keeping error handling robust
'             MyVariable = ThrowIfError(MyFunctionThatReturnsAStringIfAnErrorHappens(...))
' -----------------------------------------------------------------------------------------------------------------------
Public Function ThrowIfError(Data As Variant) As Variant
1         ThrowIfError = Data
2         If VarType(Data) = vbString Then
3             If Left$(Data, 1) = "#" Then
4                 If Right$(Data, 1) = "!" Then
5                     Throw CStr(Data)
6                 End If
7             End If
8         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsNumber
' Purpose   : Is a singleton a number?
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsNumber(x As Variant) As Boolean
1         Select Case VarType(x)
              Case vbDouble, vbInteger, vbSingle, vbLong ', vbCurrency, vbDecimal
2                 IsNumber = True
3         End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NCols
' Purpose   : Number of columns in an array. Missing has zero rows, 1-dimensional arrays
'             have one row and the number of columns returned by this function.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NCols(Optional TheArray As Variant) As Long
1         On Error GoTo ErrHandler
2         If TypeName(TheArray) = "Range" Then
3             NCols = TheArray.Columns.count
4         ElseIf IsMissing(TheArray) Then
5             NCols = 0
6         ElseIf VarType(TheArray) < vbArray Then
7             NCols = 1
8         Else
9             Select Case NumDimensions(TheArray)
                  Case 1
10                    NCols = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
11                Case Else
12                    NCols = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
13            End Select
14        End If

15        Exit Function
ErrHandler:
16        ReThrow "NCols", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NRows
' Purpose   : Number of rows in an array. Missing has zero rows, 1-dimensional arrays have one row.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NRows(Optional TheArray As Variant) As Long
1         On Error GoTo ErrHandler
2         If TypeName(TheArray) = "Range" Then
3             NRows = TheArray.Rows.count
4         ElseIf IsMissing(TheArray) Then
5             NRows = 0
6         ElseIf VarType(TheArray) < vbArray Then
7             NRows = 1
8         Else
9             Select Case NumDimensions(TheArray)
                  Case 1
10                    NRows = 1
11                Case Else
12                    NRows = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
13            End Select
14        End If

15        Exit Function
ErrHandler:
16        ReThrow "NRows", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Transpose
' Purpose   : Returns the transpose of an array.
' Arguments
' TheArray  : An array of arbitrary values.
'             Return is always 1-based, even when input is zero-based.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Transpose(TheArray As Variant) As Variant
          Dim Co As Long
          Dim i As Long
          Dim j As Long
          Dim m As Long
          Dim N As Long
          Dim Result As Variant
          Dim Ro As Long
1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray, N, m
3         Ro = LBound(TheArray, 1) - 1
4         Co = LBound(TheArray, 2) - 1
5         ReDim Result(1 To m, 1 To N)
6         For i = 1 To N
7             For j = 1 To m
8                 Result(j, i) = TheArray(i + Ro, j + Co)
9             Next j
10        Next i
11        Transpose = Result
12        Exit Function
ErrHandler:
13        ReThrow "Transpose", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArray
' Purpose   : In-place amendment of singletons and one-dimensional arrays to two dimensions.
'             singletons and 1-d arrays are returned as 2-d 1-based arrays. Leaves two
'             two dimensional arrays untouched (i.e. a zero-based 2-d array will be left as zero-based).
'             See also Force2DArrayR that also handles Range objects.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Force2DArray(ByRef TheArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
          Dim TwoDArray As Variant

1         On Error GoTo ErrHandler

2         Select Case NumDimensions(TheArray)
              Case 0
3                 ReDim TwoDArray(1 To 1, 1 To 1)
4                 TwoDArray(1, 1) = TheArray
5                 TheArray = TwoDArray
6                 NR = 1: NC = 1
7             Case 1
                  Dim i As Long
                  Dim LB As Long
8                 LB = LBound(TheArray, 1)
9                 NR = 1: NC = UBound(TheArray, 1) - LB + 1
10                ReDim TwoDArray(1 To 1, 1 To NC)
11                For i = 1 To UBound(TheArray, 1) - LBound(TheArray) + 1
12                    TwoDArray(1, i) = TheArray(LB + i - 1)
13                Next i
14                TheArray = TwoDArray
15            Case 2
16                NR = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
17                NC = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
                  'Nothing to do
18            Case Else
19                Throw "Cannot convert array of dimension greater than two"
20        End Select

21        Exit Sub
ErrHandler:
22        ReThrow "Force2DArray", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArrayR
' Purpose   : When writing functions to be called from sheets, we often don't want to process
'             the inputs as Range objects, but instead as Arrays. This method converts the
'             input into a 2-dimensional 1-based array (even if it's a single cell or single row of cells)
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Force2DArrayR(ByRef RangeOrArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
1         If TypeName(RangeOrArray) = "Range" Then RangeOrArray = RangeOrArray.Value2
2         Force2DArray RangeOrArray, NR, NC
End Sub

Private Function MaxLngs(x As Long, y As Long) As Long
1         If x > y Then
2             MaxLngs = x
3         Else
4             MaxLngs = y
5         End If
End Function

Private Function MinLngs(x As Long, y As Long) As Long
1         If x > y Then
2             MinLngs = y
3         Else
4             MinLngs = x
5         End If
End Function

Private Function IsOneAndZero(a As Variant, b As Variant) As Boolean
1         If IsNumber(a) Then
2             If a = 1 Then
3                 If IsNumber(b) Then
4                     If b = 0 Then
5                         IsOneAndZero = True
6                     End If
7                 End If
8             End If
9         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AmendSkipToColAndNumCols
' Author     : Philip Swannell
' Date       : 06-Feb-2023
' Purpose    : Check arguments SkipToCol and NumCols and if necessary convert them from String to Long, by matching into
'              the header row of the file.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AmendSkipToColAndNumCols(ByVal FileName As String, ByRef SkipToCol As Variant, ByRef NumCols As Variant, _
          Optional ByVal Delimiter As Variant, Optional ByVal IgnoreRepeated As Boolean, Optional ByVal Comment As String, _
          Optional ByVal IgnoreEmptyLines As Boolean, Optional ByVal HeaderRowNum As Long, Optional ByVal Encoding As Variant)
                
          Const Err_BadInput = " must be a positive integer or a string matching a header in the file"
          Dim Headers As Variant
          Dim i As Long
          Dim origSkipToCol As String
          Dim ReadFile As Boolean
          Dim RealHeaderRowNum As Long
                               
1         On Error GoTo ErrHandler
          
2         If IsNumber(SkipToCol) Then
3             If SkipToCol <> CLng(SkipToCol) Or SkipToCol < 1 Then
4                 Throw "SkipToCol" & Err_BadInput
5             End If
6         ElseIf VarType(SkipToCol) = vbString Then
7             ReadFile = True
8         ElseIf VarType(SkipToCol) <> vbString Then
9             Throw "SkipToCol" & Err_BadInput
10        End If
          
11        If IsNumber(NumCols) Then
12            If NumCols <> CLng(NumCols) Or NumCols < 0 Then
13                Throw "NumCols" & Err_BadInput
14            End If
15        ElseIf VarType(NumCols) = vbString Then
16            ReadFile = True
17        ElseIf VarType(NumCols) <> vbString Then
18            Throw "NumCols" & Err_BadInput
19        End If
          
20        If Not ReadFile Then Exit Sub
          
21        Headers = ThrowIfError(CSVRead(FileName, False, Delimiter, IgnoreRepeated, , Comment, IgnoreEmptyLines, , _
              HeaderRowNum, 1, 1, 0, , , , , Encoding))
22        RealHeaderRowNum = IIf(HeaderRowNum = 0, 1, HeaderRowNum)
23        If VarType(SkipToCol) = vbString Then
24            origSkipToCol = SkipToCol
25            For i = 1 To NCols(Headers)
26                If Headers(1, i) = SkipToCol Then
27                    SkipToCol = i
28                    Exit For
29                End If
30            Next i
31            If VarType(SkipToCol) = vbString Then Throw "Argument SkipToCol was given as the string '" & SkipToCol & _
                  "', but that cannot be found in the header row (row " & RealHeaderRowNum & ") of the file."
32        End If

33        If VarType(NumCols) = vbString Then
34            For i = 1 To NCols(Headers)
35                If Headers(1, i) = NumCols Then
36                    If i >= SkipToCol Then
37                        NumCols = i - SkipToCol + 1
38                    Else
39                        NumCols = SkipToCol - i + 1
40                        SkipToCol = i
41                    End If
42                    Exit For
43                End If
44            Next i
45            If VarType(NumCols) = vbString Then Throw "Argument NumCols was given as the string '" & NumCols & _
                  "', but that cannot be found in the header row (row " & RealHeaderRowNum & ") of the file."
46        End If

47        Exit Sub
ErrHandler:
48        ReThrow "AmendSkipToColAndNumCols", Err
End Sub

Private Sub ParseQuoteAllStrings(QuoteAllStrings, ByRef QuoteSimpleStrings As Boolean, QuoteComplexStrings As Boolean)

          Const Err_QuoteAllStrings = "QuoteAllStrings must be TRUE (quote all strings in Data), FALSE (quote only strings containing delimiter, double quote, line feed or carriage return) or ""Raw"" (no strings are not quoted)"

1         If VarType(QuoteAllStrings) = vbBoolean Then
2             If QuoteAllStrings Then
3                 QuoteSimpleStrings = True
4                 QuoteComplexStrings = True
5             Else
6                 QuoteSimpleStrings = False
7                 QuoteComplexStrings = True
8             End If
9         ElseIf VarType(QuoteAllStrings) = vbString Then
10            If StrComp(QuoteAllStrings, "Raw", vbTextCompare) = 0 Then
11                QuoteSimpleStrings = False
12                QuoteComplexStrings = False
13            Else
14                Throw Err_QuoteAllStrings
15            End If
16        Else
17            Throw Err_QuoteAllStrings
18        End If
End Sub


' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead
' Purpose    : Register the function CSVRead with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub RegisterCSVRead()
    Const Description As String = "Returns the contents of a comma-separated file on disk as an array."
    Dim ArgDescs() As String

    On Error GoTo ErrHandler

    ReDim ArgDescs(1 To 19)
    ArgDescs(1) = "The full name of the file, including the path, or else a URL of a file, or else a string in CSV " & _
                  "format."
    ArgDescs(2) = "Type conversion: Boolean or string. Allowed letters NDBETQK. N = Numbers, D = Dates, B = " & _
                  "Booleans, E = Excel errors, T = trim leading & trailing spaces, Q = quoted fields also " & _
                  "converted, K = quotes kept. TRUE = NDB, FALSE = no conversion."
    ArgDescs(3) = "Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe found " & _
                  "outside quoted regions within the first 10,000 characters. Enter FALSE to  see the file's " & _
                  "contents as would be displayed in a text editor."
    ArgDescs(4) = "Whether delimiters which appear at the start of a line, the end of a line or immediately after " & _
                  "another delimiter should be ignored while parsing; useful for fixed-width files with delimiter " & _
                  "padding between fields."
    ArgDescs(5) = "The format of dates in the file such as `Y-M-D` (the default), `M-D-Y` or `Y/M/D`. Also `ISO` " & _
                  "for ISO8601 (e.g., 2021-08-26T09:11:30) or `ISOZ` (time zone given e.g. " & _
                  "2021-08-26T13:11:30+05:00), in which case dates-with-time are returned in UTC time."
    ArgDescs(6) = "Rows that start with this string will be skipped while parsing."
    ArgDescs(7) = "Whether empty rows/lines in the file should be skipped while parsing (if `FALSE`, each column " & _
                  "will be assigned ShowMissingsAs for that empty row)."
    ArgDescs(8) = "The row in the file containing headers. Optional and defaults to 0. Type conversion is not " & _
                  "applied to fields in the header row, though leading and trailing spaces are trimmed."
    ArgDescs(9) = "The first row in the file that's included in the return. Optional and defaults to one more than " & _
                  "HeaderRowNum."
    ArgDescs(10) = "The column in the file at which reading starts, as a number or a string matching one of the " & _
                   "file's headers. Optional and defaults to 1 to read from the first column."
    ArgDescs(11) = "The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the " & _
                   "end of the file are read."
    ArgDescs(12) = "If a number, sets the number of columns to read from the file. If a string matching one of the " & _
                   "file's headers, sets the last column to be read. If omitted (or zero), all columns from " & _
                   "SkipToCol are read."
    ArgDescs(13) = "Indicates how `TRUE` values are represented in the file. May be a string, an array of strings " & _
                   "or a range containing strings; by default, `TRUE`, `True` and `true` are recognised."
    ArgDescs(14) = "Indicates how `FALSE` values are represented in the file. May be a string, an array of strings " & _
                   "or a range containing strings; by default, `FALSE`, `False` and `false` are recognised."
    ArgDescs(15) = "Indicates how missing values are represented in the file. May be a string, an array of strings " & _
                   "or a range containing strings. By default, only an empty field (consecutive delimiters) is " & _
                   "considered missing."
    ArgDescs(16) = "Fields which are missing in the file (consecutive delimiters) or match one of the " & _
                   "MissingStrings are returned in the array as ShowMissingsAs. Defaults to Empty, but the null " & _
                   "string or `#N/A!` error value can be good alternatives."
    ArgDescs(17) = "Allowed entries are `ASCII`, `ANSI`, `UTF-8`, or `UTF-16`. For most files this argument can be " & _
                   "omitted and CSVRead will detect the file's encoding."
    ArgDescs(18) = "The character that represents a decimal point. If omitted, then the value from Windows " & _
                   "regional settings is used."
    ArgDescs(19) = "For use from VBA only."
    Application.MacroOptions "CSVRead", Description, , , , , , , , , ArgDescs
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function CSVRead failed with error: " & Err.Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVWrite
' Purpose    : Register the function CSVWrite with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub RegisterCSVWrite()
    Const Description As String = "Creates a comma-separated file on disk containing Data. Any existing file of " & _
                                  "the same name is overwritten. If successful, the function returns FileName, " & _
                                  "otherwise an ""error string"" (starts with `#`, ends with `!`) describing what " & _
                                  "went wrong."
    Dim ArgDescs() As String

    On Error GoTo ErrHandler

    ReDim ArgDescs(1 To 10)
    ArgDescs(1) = "An array of data, or an Excel range. Elements may be strings, numbers, dates, Booleans, empty, " & _
                  "Excel errors or null values. Data typically has two dimensions, but if Data has only one " & _
                  "dimension then the output file has a single column, one field per row."
    ArgDescs(2) = "The full name of the file, including the path. Alternatively, if FileName is omitted, then the " & _
                  "function returns Data converted CSV-style to a string."
    ArgDescs(3) = "If TRUE (the default) then all strings in Data are quoted before being written to file. If " & _
                  "FALSE only strings containing Delimiter, line feed, carriage return or quote are quoted. If " & _
                  """Raw"" no strings are quoted. The file may not be valid csv format."
    ArgDescs(4) = "A format string that determines how dates, including cells formatted as dates, appear in the " & _
                  "file. If omitted, defaults to `yyyy-mm-dd`."
    ArgDescs(5) = "Format for datetimes. Defaults to `ISO` which abbreviates `yyyy-mm-ddThh:mm:ss`. Use `ISOZ` for " & _
                  "ISO8601 format with time zone the same as the PC's clock. Use with care, daylight saving may be " & _
                  "inconsistent across the datetimes in data."
    ArgDescs(6) = "The delimiter string, if omitted defaults to a comma. Delimiter may have more than one " & _
                  "character."
    ArgDescs(7) = "Allowed entries are `ANSI` (the default), `UTF-8`, `UTF-16`, `UTF-8NOBOM` and `UTF-16NOBOM`. An " & _
                  "error will result if this argument is `ANSI` but Data contains characters with code point above " & _
                  "127."
    ArgDescs(8) = "Sets the file's line endings. Enter `Windows`, `Unix` or `Mac`. Also supports the line-ending " & _
                  "characters themselves or the strings `CRLF`, `LF` or `CR`. The default is `Windows` if FileName " & _
                  "is provided, or `Unix` if not."
    ArgDescs(9) = "How the Boolean value True is to be represented in the file. Optional, defaulting to ""True""."
    ArgDescs(10) = "How the Boolean value False is to be represented in the file. Optional, defaulting to " & _
                   """False""."
    Application.MacroOptions "CSVWrite", Description, , , , , , , , , ArgDescs
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function CSVWrite failed with error: " & Err.Description
End Sub

