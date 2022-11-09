Attribute VB_Name = "modCSVReadWrite"
' VBA-CSV
' Copyright (C) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme
' This version at: https://github.com/PGS62/VBA-CSV/releases/tag/v0.23

'Installation:
'1) Import this module into your project (Open VBA Editor, Alt + F11; File > Import File).

'2) Add two references (In VBA Editor Tools > References)
'   Microsoft Scripting Runtime
'   Microsoft VBScript Regular Expressions 5.5 (or a later version if available)

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
'      https://github.com/PGS62/VBA-CSV/releases/download/v0.23/VBA-CSV-Intellisense.xlsx
'      into the workbook that contains this VBA code.

Option Explicit

Private m_FSO As Scripting.FileSystemObject
Private Const DQ = """"
Private Const DQ2 = """"""

#If VBA7 And Win64 Then
    'for 64-bit Excel
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As Long
#Else
    'for 32-bit Excel
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Private Enum enmErrorStyle
    es_ReturnString = 0
    es_RaiseError = 1
End Enum

Private Const m_ErrorStyle As Long = es_ReturnString

Private Const m_LBound As Long = 1 'Sets the array lower bounds of the return from CSVRead.
'To return zero-based arrays, change the value of this constant to 0.

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
'             ConvertTypes should be a string of zero or more letters from allowed characters `NDBETQ`.
'
'             The most commonly useful letters are:
'             1) `N` number fields are returned as numbers (Doubles).
'             2) `D` date fields (that respect DateFormat) are returned as Dates.
'             3) `B` fields matching TrueStrings or FalseStrings are returned as Booleans.
'
'             ConvertTypes is optional and defaults to the null string for no type conversion. `TRUE` is
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
' Encoding  : Allowed entries are `ASCII`, `ANSI`, `UTF-8`, or `UTF-16`. For most files this argument can be
'             omitted and CSVRead will detect the file's encoding. If auto-detection does not work, then
'             it's possible that the file is encoded `UTF-8` or `UTF-16` but without a byte option mark to
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
' Notes     : See also companion function CSVRead.
'
'             For discussion of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
' -----------------------------------------------------------------------------------------------------------------------
Public Function CSVRead(ByVal FileName As String, Optional ByVal ConvertTypes As Variant = False, _
    Optional ByVal Delimiter As Variant, Optional ByVal IgnoreRepeated As Boolean, _
    Optional ByVal DateFormat As String = "Y-M-D", Optional ByVal Comment As String, _
    Optional ByVal IgnoreEmptyLines As Boolean, Optional ByVal HeaderRowNum As Long, _
    Optional ByVal SkipToRow As Long, Optional ByVal SkipToCol As Long = 1, _
    Optional ByVal NumRows As Long, Optional ByVal NumCols As Long, _
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
          Const Err_NumCols As String = "NumCols must be positive to read a given number of columns, or zero or omitted " & _
              "to read all columns from SkipToCol to the maximum column encountered"
          Const Err_NumRows As String = "NumRows must be positive to read a given number of rows, or zero or omitted to " & _
              "read all rows from SkipToRow to the end of the file"
          Const Err_Seps1 As String = "DecimalSeparator must be a single character"
          Const Err_Seps2 As String = "DecimalSeparator must not be equal to the first character of Delimiter or to a " & _
              "line-feed or carriage-return"
          Const Err_SkipToCol As String = "SkipToCol must be at least 1"
          Const Err_SkipToRow As String = "SkipToRow must be at least 1"
          Const Err_Comment As String = "Comment must not contain double-quote, line feed or carriage return"
          Const Err_HeaderRowNum As String = "HeaderRowNum must be greater than or equal to zero and less than or equal to SkipToRow"
          
          Dim AcceptWithoutTimeZone As Boolean
          Dim AcceptWithTimeZone As Boolean
          Dim Adj As Long
          Dim AnyConversion As Boolean
          Dim AnySentinels As Boolean
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
          Dim SysDateOrder As Long
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
              ConvertQuoted, TrimFields, ColByColFormatting, HeaderRowNum, CTDict

43        Set Sentinels = New Scripting.Dictionary
44        MakeSentinels Sentinels, ConvertQuoted, strDelimiter, MaxSentinelLength, AnySentinels, ShowBooleansAsBooleans, _
              ShowErrorsAsErrors, ShowMissingsAs, TrueStrings, FalseStrings, MissingStrings
          
45        If ShowDatesAsDates Then
46            ParseDateFormat DateFormat, DateOrder, DateSeparator, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone
47            SysDateOrder = Application.International(xlDateOrder)
48            SysDateSeparator = Application.International(xlDateSeparator)
49        End If

50        If HeaderRowNum < 0 Then Throw Err_HeaderRowNum
51        If SkipToRow = 0 Then SkipToRow = HeaderRowNum + 1
52        If HeaderRowNum > SkipToRow Then Throw Err_HeaderRowNum
53        If SkipToCol = 0 Then SkipToCol = 1
54        If SkipToRow < 1 Then Throw Err_SkipToRow
55        If SkipToCol < 1 Then Throw Err_SkipToCol
56        If NumRows < 0 Then Throw Err_NumRows
57        If NumCols < 0 Then Throw Err_NumCols

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
115                       UnquotedLength = Len(Unquote(Mid$(CSVContents, Starts(k), Lengths(k)), DQ, 4))
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
                          Lengths(k), TrimFields, DQ, QuoteCounts(k), ConvertQuoted, ShowNumbersAsNumbers, SepStandard, _
                          DecimalSeparator, SysDecimalSeparator, ShowDatesAsDates, ISO8601, AcceptWithoutTimeZone, _
                          AcceptWithTimeZone, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, AnySentinels, _
                          Sentinels, MaxSentinelLength, ShowMissingsAs)
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
187                   CT = False
188               End If
                  
189               ParseCTString CT, ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, _
                      ShowErrorsAsErrors, ConvertQuoted, TrimFields
                  
190               AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
                      ShowBooleansAsBooleans Or ShowErrorsAsErrors
                      
191               Set Sentinels = New Scripting.Dictionary
                  
192               MakeSentinels Sentinels, ConvertQuoted, strDelimiter, MaxSentinelLength, AnySentinels, ShowBooleansAsBooleans, _
                      ShowErrorsAsErrors, ShowMissingsAs, TrueStrings, FalseStrings, MissingStrings

193               For i = 1 To NR
194                   If Not IsEmpty(ReturnArray(i + Adj, j + Adj)) Then
195                       Field = CStr(ReturnArray(i + Adj, j + Adj))
196                       QC = CountQuotes(Field, DQ)
197                       ReturnArray(i + Adj, j + Adj) = ConvertField(Field, AnyConversion, _
                              Len(ReturnArray(i + Adj, j + Adj)), TrimFields, DQ, QC, ConvertQuoted, _
                              ShowNumbersAsNumbers, SepStandard, DecimalSeparator, SysDecimalSeparator, _
                              ShowDatesAsDates, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone, DateOrder, _
                              DateSeparator, SysDateOrder, SysDateSeparator, AnySentinels, Sentinels, _
                              MaxSentinelLength, ShowMissingsAs)
198                   End If
199               Next i
200           Next j
201       End If

202       CSVRead = ReturnArray

203       Exit Function

ErrHandler:
204       CSVRead = ReThrow("CSVRead", Err, m_ErrorStyle = es_ReturnString)
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
6             Select Case UCase$(Replace(Replace(Encoding, "-", vbNullString), " ", vbNullString))
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
                      
12                Case "UTF8"
13                    Encoding = "UTF-8"
14                    CharSet = "utf-8"
15                Case "UTF16"
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
6                 .Pattern = "^[NDBETQ]*$"
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
                  IIf(InStr(1, CT, "T", vbTextCompare), "T", vbNullString)
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
'  TrimFields            : Set only if ConvertTypes is not an array
'  ColByColFormatting    : Set to True if ConvertTypes is an array
'  HeaderRowNum          : As passed to CSVRead, used to throw an error if HeaderRowNum has not been specified when
'                          it needs to have been.
'  CTDict                : Set to a dictionary keyed on the elements of the left column (or top row) of ConvertTypes,
'                          each element containing the corresponding right (or bottom) element.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseConvertTypes(ByVal ConvertTypes As Variant, ByRef ShowNumbersAsNumbers As Boolean, _
    ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
    ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef TrimFields As Boolean, _
    ByRef ColByColFormatting As Boolean, HeaderRowNum As Long, ByRef CTDict As Scripting.Dictionary)
          
          Const Err_2D As String = "If ConvertTypes is given as a two dimensional array then the " & _
              " lower bounds in each dimension must be 1"
          Const Err_Ambiguous As String = "ConvertTypes is ambiguous, it can be interpreted as two rows, or as two columns"
          Const Err_BadColumnIdentifier As String = "Column identifiers in the left column (or top row) of " & _
              "ConvertTypes must be strings or non-negative whole numbers"
          Const Err_BadCT As String = "Type Conversion given in bottom row (or right column) of ConvertTypes must be " & _
              "Booleans or strings containing letters NDBETQ"
          Const Err_ConvertTypes As String = "ConvertTypes must be a Boolean, a string with allowed letters ""NDBETQ"" or an array"
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
2         If VarType(ConvertTypes) = vbString Or VarType(ConvertTypes) = vbBoolean Or IsEmpty(ConvertTypes) Then
3             ParseCTString CStr(ConvertTypes), ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, _
                  ShowErrorsAsErrors, ConvertQuoted, TrimFields
4             ColByColFormatting = False
5             Exit Sub
6         End If

7         If TypeName(ConvertTypes) = "Range" Then ConvertTypes = ConvertTypes.Value2
8         ND = NumDimensions(ConvertTypes)
9         If ND = 1 Then
10            ConvertTypes = OneDArrayToTwoDArray(ConvertTypes)
11        ElseIf ND = 2 Then
12            If LBound(ConvertTypes, 1) <> 1 Or LBound(ConvertTypes, 2) <> 1 Then
13                Throw Err_2D
14            End If
15        End If

16        NR = NRows(ConvertTypes)
17        NC = NCols(ConvertTypes)
18        If NR = 2 And NC = 2 Then
              'Tricky - have we been given two rows or two columns?
19            If Not IsCTValid(ConvertTypes(2, 2)) Then Throw Err_ConvertTypes
20            If IsCTValid(ConvertTypes(1, 2)) And IsCTValid(ConvertTypes(2, 1)) Then
21                If StandardiseCT(ConvertTypes(1, 2)) <> StandardiseCT(ConvertTypes(2, 1)) Then
22                    Throw Err_Ambiguous
23                End If
24            End If
25            If IsCTValid(ConvertTypes(2, 1)) Then
26                ConvertTypes = Transpose(ConvertTypes)
27                Transposed = True
28            End If
29        ElseIf NR = 2 Then
30            ConvertTypes = Transpose(ConvertTypes)
31            Transposed = True
32            NR = NC
33        ElseIf NC <> 2 Then
34            Throw Err_ConvertTypes
35        End If
36        LCN = LBound(ConvertTypes, 2)
37        RCN = LCN + 1
38        For i = LBound(ConvertTypes, 1) To UBound(ConvertTypes, 1)
39            ColIdentifier = ConvertTypes(i, LCN)
40            CT = ConvertTypes(i, RCN)
41            If IsNumber(ColIdentifier) Then
42                If ColIdentifier <> CLng(ColIdentifier) Then
43                    Throw Err_BadColumnIdentifier & _
                          " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & _
                          ") is " & CStr(ColIdentifier)
44                ElseIf ColIdentifier < 0 Then
45                    Throw Err_BadColumnIdentifier & " but ConvertTypes(" & _
                          IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & ") is " & CStr(ColIdentifier)
46                End If
47            ElseIf VarType(ColIdentifier) <> vbString Then
48                Throw Err_BadColumnIdentifier & " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & _
                      ") is of type " & TypeName(ColIdentifier)
49            End If
50            If Not IsCTValid(CT) Then
51                If VarType(CT) = vbString Then
52                    Throw Err_BadCT & " but ConvertTypes(" & IIf(Transposed, "2," & CStr(i), CStr(i) & ",2") & _
                          ") is string """ & CStr(CT) & """"
53                Else
54                    Throw Err_BadCT & " but ConvertTypes(" & IIf(Transposed, "2," & CStr(i), CStr(i) & ",2") & _
                          ") is of type " & TypeName(CT)
55                End If
56            End If

57            If CTDict.Exists(ColIdentifier) Then
58                If Not CTsEqual(CTDict.item(ColIdentifier), CT) Then
59                    Throw "ConvertTypes is contradictory. Column " & CStr(ColIdentifier) & _
                          " is specified to be converted using two different conversion rules: " & CStr(CT) & _
                          " and " & CStr(CTDict.item(ColIdentifier))
60                End If
61            Else
62                CT = StandardiseCT(CT)
                  'Need this line to ensure that we parse the DateFormat provided when doing Col-by-col type conversion
63                If InStr(CT, "D") > 0 Then ShowDatesAsDates = True
64                If VarType(ColIdentifier) = vbString Then
65                    If HeaderRowNum = 0 Then
66                        Throw Err_HeaderRowNum
67                    End If
68                End If
69                CTDict.Add ColIdentifier, CT
70            End If
71        Next i
72        ColByColFormatting = True
73        Exit Sub
ErrHandler:
74        ReThrow "ParseConvertTypes", Err
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
'  TrimFields            : Should leading and trailing spaces be trimmed from fields?
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseCTString(ByVal ConvertTypes As String, ByRef ShowNumbersAsNumbers As Boolean, _
    ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
    ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef TrimFields As Boolean)

          Const Err_ConvertTypes As String = "ConvertTypes must be Boolean or string with allowed letters NDBETQ. " & _
              """N"" show numbers as numbers, ""D"" show dates as dates, ""B"" show Booleans " & _
              "as Booleans, ""E"" show Excel errors as errors, ""T"" to trim leading and trailing " & _
              "spaces from fields, ""Q"" rules NDBE apply even to quoted fields, TRUE = ""NDB"" " & _
              "(convert unquoted numbers, dates and Booleans), FALSE = no conversion"
          Const Err_Quoted As String = "ConvertTypes is incorrect, ""Q"" indicates that conversion should apply even to " & _
              "quoted fields, but none of ""N"", ""D"", ""B"" or ""E"" are present to indicate which type conversion to apply"
          Dim i As Long

1         On Error GoTo ErrHandler

2         If ConvertTypes = "True" Or ConvertTypes = "False" Then
3             ConvertTypes = StandardiseCT(CBool(ConvertTypes))
4         End If

5         ShowNumbersAsNumbers = False
6         ShowDatesAsDates = False
7         ShowBooleansAsBooleans = False
8         ShowErrorsAsErrors = False
9         ConvertQuoted = False
10        For i = 1 To Len(ConvertTypes)
              'Adding another letter? Also change method IsCTValid!
11            Select Case UCase$(Mid$(ConvertTypes, i, 1))
                  Case "N"
12                    ShowNumbersAsNumbers = True
13                Case "D"
14                    ShowDatesAsDates = True
15                Case "B"
16                    ShowBooleansAsBooleans = True
17                Case "E"
18                    ShowErrorsAsErrors = True
19                Case "Q"
20                    ConvertQuoted = True
21                Case "T"
22                    TrimFields = True
23                Case Else
24                    Throw Err_ConvertTypes & " Found unrecognised character '" _
                          & Mid$(ConvertTypes, i, 1) & "'"
25            End Select
26        Next i
          
27        If ConvertQuoted And Not (ShowNumbersAsNumbers Or ShowDatesAsDates Or _
              ShowBooleansAsBooleans Or ShowErrorsAsErrors) Then
28            Throw Err_Quoted
29        End If

30        Exit Sub
ErrHandler:
31        ReThrow "ParseCTString", Err
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
' Purpose    : Attempt to detect the file's encoding by looking for a byte option mark
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
              'File is probably encoded UTF-16 LE BOM (little endian, with Byte Option Marker)
15            DetectEncoding = "UTF-16"
16        ElseIf (intAsc1Chr = 254) And (intAsc2Chr = 255) Then
              'File is probably encoded UTF-16 BE BOM (big endian, with Byte Option Marker)
17            DetectEncoding = "UTF-16"
18        Else
19            If T.AtEndOfStream Then
20                DetectEncoding = "ANSI"
21            End If
22            intAsc3Chr = Asc(T.Read(1))
23            If (intAsc1Chr = 239) And (intAsc2Chr = 187) And (intAsc3Chr = 191) Then
                  'File is probably encoded UTF-8 with BOM
24                DetectEncoding = "UTF-8"
25            Else
                  'We don't know, assume ANSI but that may be incorrect.
26                DetectEncoding = "ANSI"
27            End If
28        End If

EarlyExit:
29        T.Close: Set T = Nothing
30        Exit Function

ErrHandler:
31        If Not T Is Nothing Then T.Close
32        ReThrow "DetectEncoding", Err
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
                  'Will be exact if the file has a BOM (2 bytes) and contains only _
                   ansi characters (2 bytes each). When file contains non-ansi characters _
                   this will overestimate the character count.
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
          Dim SysDateOrder As Long
          Dim SysDateSeparator As String
          Dim TrialDelim As Variant
          
1         On Error GoTo ErrHandler
2         For Each TrialDelim In Array(",", vbTab, "|", ";", vbCr, vbLf)
3             DelimAt = InStr(FirstChunk, CStr(TrialDelim))
4             If DelimAt > 0 Then
5                 FirstField = Left$(FirstChunk, DelimAt - 1)
6                 If InStr(FirstField, "-") > 0 Or InStr(FirstField, "/") > 0 Or InStr(FirstField, " ") > 0 Then

7                     SysDateOrder = Application.International(xlDateOrder)
8                     SysDateSeparator = Application.International(xlDateSeparator)

9                     For Each DateSeparator In Array("/", "-", " ")
10                        For DateOrder = 0 To 2
11                            CastToDate FirstField, DtOut, DateOrder, CStr(DateSeparator), SysDateOrder, SysDateSeparator, Converted
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

          'No commonly-used delimiter found in the file outside quoted regions _
           and in the first MAX_CHUNKS * CHUNK_SIZE characters. Assume comma _
           unless that's the decimal separator.
          
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
          
13        Err_DateFormat = "DateFormat not valid should be one of 'ISO', 'ISOZ', 'M-D-Y', 'D-M-Y', 'Y-M-D', " & _
              "'M/D/Y', 'D/M/Y', 'Y/M/D', 'M D Y', 'D M Y' or 'Y M D'" & ". Omit to use the default date format of 'Y-M-D'"
              
          'Replace repeated D's with a single D, etc since CastToDate only needs _
           to know the order in which the three parts of the date appear.
14        If Len(DateFormat) > 5 Then
15            DateFormat = UCase$(DateFormat)
16            ReplaceRepeats DateFormat, "D"
17            ReplaceRepeats DateFormat, "M"
18            ReplaceRepeats DateFormat, "Y"
19        End If
             
20        If Len(DateFormat) = 0 Then 'use "Y-M-D"
21            DateOrder = 2
22            DateSeparator = "-"
23        ElseIf Len(DateFormat) <> 5 Then
24            Throw Err_DateFormat
25        ElseIf Mid$(DateFormat, 2, 1) <> Mid$(DateFormat, 4, 1) Then
26            Throw Err_DateFormat
27        Else
28            DateSeparator = Mid$(DateFormat, 2, 1)
29            If DateSeparator <> "/" And DateSeparator <> "-" And DateSeparator <> " " Then Throw Err_DateFormat
30            Select Case UCase$(Left$(DateFormat, 1) & Mid$(DateFormat, 3, 1) & Right$(DateFormat, 1))
                  Case "MDY"
31                    DateOrder = 0
32                Case "DMY"
33                    DateOrder = 1
34                Case "YMD"
35                    DateOrder = 2
36                Case Else
37                    Throw Err_DateFormat
38            End Select
39        End If

40        Exit Sub
ErrHandler:
41        ReThrow "ParseDateFormat", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReplaceRepeats
' Purpose    : Replace repeated instances of a character in a string with a single instance.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ReplaceRepeats(ByRef TheString As String, TheChar As String)
          Dim ChCh As String
1         ChCh = TheChar & TheChar
2         Do While InStr(TheString, ChCh) > 0
3             TheString = Replace(TheString, ChCh, TheChar)
4         Loop
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCSVContents
' Purpose    : Parse the contents of a CSV file. Returns a string Buffer together with arrays which assist splitting
'              Buffer into a two-dimensional array.
' Parameters :
'  ContentsOrStream: The contents of a CSV file as a string, or else a Scripting.TextStream.
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
                                      ColIndexes, QuoteCounts, j)
126                           End If
127                       Else
128                           If RowNum = HeaderRowNum Then
129                               HeaderRow = GetLastParsedRow(Buffer, Starts, Lengths, _
                                      ColIndexes, QuoteCounts, j)
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
                      'Malformed Buffer (not RFC4180 compliant). There should always be an even number of double quotes. _
                       If there are an odd number then all text after the last double quote in the file will be (part of) _
                       the last field in the last line.
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
180           If NumRows = 0 Then 'Attempting to read from SkipToRow to the end of the file, but that would be zero or _
                                   a negative number of rows. So throw an error.
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
    ColIndexes() As Long, QuoteCounts() As Long, j As Long) As Variant
          Dim NC As Long

          Dim Field As String
          Dim i As Long
          Dim Res() As String

1         On Error GoTo ErrHandler
2         NC = ColIndexes(j)

3         ReDim Res(1 To 1, 1 To NC)
4         For i = j To j - NC + 1 Step -1
5             Field = Mid$(Buffer, Starts(i), Lengths(i))
6             Res(1, NC + i - j) = Unquote(Trim$(Field), DQ, QuoteCounts(i))
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

          Const ChunkSize As Long = 5000  ' The number of characters to read from the stream on each call. _
                                            Set to a small number for testing logic and a bigger number for _
                                            performance, but not too high since a common use case is reading _
                                            just the first line of a file. Suggest 5000? Note that when reading _
                                            an entire file (NumRows argument to CSVRead is zero) function _
                                            GetMoreFromStream is not called.
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

          'Line below arranges that when calling Instr(Buffer,....) we don't pointlessly scan the space characters _
           we can be sure that there is space in the buffer to write the extra characters thanks to
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
'  SysDateOrder         : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
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
    TrimFields As Boolean, QuoteChar As String, quoteCount As Long, ConvertQuoted As Boolean, _
    ShowNumbersAsNumbers As Boolean, SepStandard As Boolean, DecimalSeparator As String, _
    SysDecimalSeparator As String, ShowDatesAsDates As Boolean, ISO8601 As Boolean, _
    AcceptWithoutTimeZone As Boolean, AcceptWithTimeZone As Boolean, DateOrder As Long, _
    DateSeparator As String, SysDateOrder As Long, SysDateSeparator As String, _
    AnySentinels As Boolean, Sentinels As Dictionary, MaxSentinelLength As Long, _
    ShowMissingsAs As Variant) As Variant

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
29            If Left$(Field, 1) = QuoteChar Then
30                If Right$(Field, 1) = QuoteChar Then
31                    Field = Mid$(Field, 2, FieldLength - 2)
32                    If quoteCount > 2 Then
33                        Field = Replace(Field, QuoteChar & QuoteChar, QuoteChar)
34                    End If
35                    If ConvertQuoted Then
36                        FieldLength = Len(Field)
37                    Else
38                        ConvertField = Field
39                        Exit Function
40                    End If
41                End If
42            End If
43        End If

44        If Not ConvertQuoted Then
45            If quoteCount > 0 Then
46                ConvertField = Field
47                Exit Function
48            End If
49        End If

50        If ShowNumbersAsNumbers Then
51            CastToDouble Field, dblResult, SepStandard, DecimalSeparator, SysDecimalSeparator, Converted
52            If Converted Then
53                ConvertField = dblResult
54                Exit Function
55            End If
56        End If

57        If ShowDatesAsDates Then
58            If ISO8601 Then
59                CastISO8601 Field, dtResult, Converted, AcceptWithoutTimeZone, AcceptWithTimeZone
60            Else
61                CastToDate Field, dtResult, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
62            End If
63            If Not Converted Then
64                If InStr(Field, ":") > 0 Then
65                    CastToTime Field, dtResult, Converted
66                    If Not Converted Then
67                        CastToTimeB Field, dtResult, Converted
68                    End If
69                End If
70            End If
71            If Converted Then
72                ConvertField = dtResult
73                Exit Function
74            End If
75        End If

76        ConvertField = Field
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Unquote
' Purpose    : Unquote a field.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Unquote(ByVal Field As String, QuoteChar As String, quoteCount As Long) As String

1         On Error GoTo ErrHandler
2         If quoteCount > 0 Then
3             If Left$(Field, 1) = QuoteChar Then
4                 If Right$(QuoteChar, 1) = QuoteChar Then
5                     Field = Mid$(Field, 2, Len(Field) - 2)
6                     If quoteCount > 2 Then
7                         Field = Replace(Field, QuoteChar & QuoteChar, QuoteChar)
8                     End If
9                 End If
10            End If
11        End If
12        Unquote = Field

13        Exit Function
ErrHandler:
14        ReThrow "Unquote", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble, sub-routine of ConvertField
' Purpose    : Casts strIn to double where strIn has specified decimals separator.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDouble(strIn As String, ByRef dblOut As Double, SepStandard As Boolean, _
    DecimalSeparator As String, SysDecimalSeparator As String, ByRef Converted As Boolean)
          
1         On Error GoTo ErrHandler
2         If SepStandard Then
3             dblOut = CDbl(strIn)
4         Else
5             dblOut = CDbl(Replace(strIn, DecimalSeparator, SysDecimalSeparator))
6         End If
7         Converted = True
ErrHandler:
          'Do nothing - strIn was not a string representing a number.
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
Private Sub CastToDate(strIn As String, ByRef DtOut As Date, DateOrder As Long, _
    DateSeparator As String, SysDateOrder As Long, SysDateSeparator As String, _
    ByRef Converted As Boolean)

          Dim D As String
          Dim m As String
          Dim pos1 As Long 'First date separator
          Dim pos2 As Long 'Second date separator
          Dim pos3 As Long 'Space to separate date from time
          Dim pos4 As Long 'decimal point for fractions of a second
          Dim Converted2 As Boolean
          Dim HasFractionalSecond As Boolean
          Dim HasTimePart As Boolean
          Dim TimePart As String
          Dim TimePartConverted As Date
          Dim y As String
          
1         On Error GoTo ErrHandler
          
2         pos1 = InStr(strIn, DateSeparator)
3         If pos1 = 0 Then Exit Sub
4         pos2 = InStr(pos1 + 1, strIn, DateSeparator)
5         If pos2 = 0 Then Exit Sub
6         pos3 = InStr(pos2 + 1, strIn, " ")
          
7         HasTimePart = pos3 > 0
          
8         If Not HasTimePart Then
9             If DateOrder = 2 Then 'Y-M-D is unambiguous as long as year given as 4 digits
10                If pos1 = 5 Then
11                    DtOut = CDate(strIn)
12                    Converted = True
13                    Exit Sub
14                End If
15            ElseIf DateOrder = SysDateOrder Then
16                DtOut = CDate(strIn)
17                Converted = True
18                Exit Sub
19            End If
20            If DateOrder = 0 Then 'M-D-Y
21                m = Left$(strIn, pos1 - 1)
22                D = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
23                y = Mid$(strIn, pos2 + 1)
24            ElseIf DateOrder = 1 Then 'D-M-Y
25                D = Left$(strIn, pos1 - 1)
26                m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
27                y = Mid$(strIn, pos2 + 1)
28            ElseIf DateOrder = 2 Then 'Y-M-D
29                y = Left$(strIn, pos1 - 1)
30                m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
31                D = Mid$(strIn, pos2 + 1)
32            Else
33                Throw "DateOrder must be 0, 1, or 2"
34            End If
35            If SysDateOrder = 0 Then
36                DtOut = CDate(m & SysDateSeparator & D & SysDateSeparator & y)
37                Converted = True
38            ElseIf SysDateOrder = 1 Then
39                DtOut = CDate(D & SysDateSeparator & m & SysDateSeparator & y)
40                Converted = True
41            ElseIf SysDateOrder = 2 Then
42                DtOut = CDate(y & SysDateSeparator & m & SysDateSeparator & D)
43                Converted = True
44            End If
45            Exit Sub
46        End If

47        pos4 = InStr(pos3 + 1, strIn, ".")
48        HasFractionalSecond = pos4 > 0

49        If DateOrder = 0 Then 'M-D-Y
50            m = Left$(strIn, pos1 - 1)
51            D = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
52            y = Mid$(strIn, pos2 + 1, pos3 - pos2 - 1)
53            TimePart = Mid$(strIn, pos3)
54        ElseIf DateOrder = 1 Then 'D-M-Y
55            D = Left$(strIn, pos1 - 1)
56            m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
57            y = Mid$(strIn, pos2 + 1, pos3 - pos2 - 1)
58            TimePart = Mid$(strIn, pos3)
59        ElseIf DateOrder = 2 Then 'Y-M-D
60            y = Left$(strIn, pos1 - 1)
61            m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
62            D = Mid$(strIn, pos2 + 1, pos3 - pos2 - 1)
63            TimePart = Mid$(strIn, pos3)
64        Else
65            Throw "DateOrder must be 0, 1, or 2"
66        End If
67        If Not HasFractionalSecond Then
68            If DateOrder = 2 Then 'Y-M-D is unambiguous as long as year given as 4 digits
69                If pos1 = 5 Then
70                    DtOut = CDate(strIn)
71                    Converted = True
72                    Exit Sub
73                End If
74            ElseIf DateOrder = SysDateOrder Then
75                DtOut = CDate(strIn)
76                Converted = True
77                Exit Sub
78            End If
          
79            If SysDateOrder = 0 Then
80                DtOut = CDate(m & SysDateSeparator & D & SysDateSeparator & y & TimePart)
81                Converted = True
82            ElseIf SysDateOrder = 1 Then
83                DtOut = CDate(D & SysDateSeparator & m & SysDateSeparator & y & TimePart)
84                Converted = True
85            ElseIf SysDateOrder = 2 Then
86                DtOut = CDate(y & SysDateSeparator & m & SysDateSeparator & D & TimePart)
87                Converted = True
88            End If
89        Else 'CDate does not cope with fractional seconds, so use CastToTimeB
90            CastToTimeB Mid$(TimePart, 2), TimePartConverted, Converted2
91            If Converted2 Then
92                If SysDateOrder = 0 Then
93                    DtOut = CDate(m & SysDateSeparator & D & SysDateSeparator & y) + TimePartConverted
94                    Converted = True
95                ElseIf SysDateOrder = 1 Then
96                    DtOut = CDate(D & SysDateSeparator & m & SysDateSeparator & y) + TimePartConverted
97                    Converted = True
98                ElseIf SysDateOrder = 2 Then
99                    DtOut = CDate(y & SysDateSeparator & m & SysDateSeparator & D) + TimePartConverted
100                   Converted = True
101               End If
102           End If
103       End If

104       Exit Sub
ErrHandler:
          'Do nothing - was not a string representing a date with the specified date order and date separator.
End Sub

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

          Dim L As Long
          Dim LocalTime As Double
          Dim MilliPart As Double
          Dim MinusPos As Long
          Dim PlusPos As Long
          Dim Sign As Long
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
                  'This works irrespective of Windows regional settings
32                DtOut = CDate(strIn)
33                Converted = True
34                Exit Sub
35            End If
36        ElseIf L < 10 Then
37            Converted = False
38            Exit Sub
39        ElseIf L > 40 Then
40            Converted = False
41            Exit Sub
42        End If

43        Converted = False
          
44        If AcceptWithoutTimeZone Then
45            If AcceptWithTimeZone Then
46                If Not rxYesYes.Test(strIn) Then Exit Sub
47            Else
48                If Not RxYesNo.Test(strIn) Then Exit Sub
49            End If
50        Else
51            If AcceptWithTimeZone Then
52                If Not RxNoYes.Test(strIn) Then Exit Sub
53            Else
54                If Not rxNoNo.Test(strIn) Then Exit Sub
55            End If
56        End If
          
          'Replace the "T" separator
57        Mid$(strIn, 11, 1) = " "
          
58        If L = 19 Then
59            DtOut = CDate(strIn)
60            Converted = True
61            Exit Sub
62        End If

63        If Right$(strIn, 1) = "Z" Then
64            Sign = 0
65            ZAtEnd = True
66        Else
67            PlusPos = InStr(20, strIn, "+")
68            If PlusPos > 0 Then
69                Sign = 1
70            Else
71                MinusPos = InStr(20, strIn, "-")
72                If MinusPos > 0 Then
73                    Sign = -1
74                End If
75            End If
76        End If

77        If Mid$(strIn, 20, 1) = "." Then 'Have fraction of a second
78            Select Case Sign
                  Case 0
                      'Example: "2021-08-23T08:47:20.920Z"
79                    MilliPart = CDbl(Mid$(strIn, 20, IIf(ZAtEnd, L - 20, L - 19)))
80                Case 1
                      'Example: "2021-08-23T08:47:20.920+05:00"
81                    MilliPart = CDbl(Mid$(strIn, 20, PlusPos - 20))
82                Case -1
                      'Example: "2021-08-23T08:47:20.920-05:00"
83                    MilliPart = CDbl(Mid$(strIn, 20, MinusPos - 20))
84            End Select
85        End If
          
86        LocalTime = CDate(Left$(strIn, 19)) + MilliPart / 86400

          Dim Adjust As Date
87        Select Case Sign
              Case 0
88                DtOut = LocalTime
89                Converted = True
90                Exit Sub
91            Case 1
92                If L <> PlusPos + 5 Then Exit Sub
93                Adjust = CDate(Right$(strIn, 5))
94                DtOut = LocalTime - Adjust
95                Converted = True
96            Case -1
97                If L <> MinusPos + 5 Then Exit Sub
98                Adjust = CDate(Right$(strIn, 5))
99                DtOut = LocalTime + Adjust
100               Converted = True
101       End Select

102       Exit Sub
ErrHandler:
          'Was not recognised as ISO8601 date
103   End Sub

      ' -----------------------------------------------------------------------------------------------------------------------
      ' Procedure  : GetLocalOffsetToUTC
      ' Purpose    : Get the PC's offset to UTC.
      'See "gogeek"'s post at _
 https://stackoverflow.com/questions/1600875/how-to-get-the-current-datetime-in-utc-from-an-excel-vba-macro

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
53                NewKey = DQ & Replace(Keys(i), DQ, DQ2) & DQ
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
'             carriage return or double quote. In all cases, double quotes are escaped by another double
'             quote.
' DateFormat: A format string that determines how dates, including cells formatted as dates, appear in the
'             file. If omitted, defaults to `yyyy-mm-dd`.
' DateTimeFormat: Format for datetimes. Defaults to `ISO` which abbreviates `yyyy-mm-ddThh:mm:ss`. Use
'             `ISOZ` for ISO8601 format with time zone the same as the PC's clock. Use with care, daylight
'             saving may be inconsistent across the datetimes in data.
' Delimiter : The delimiter string, if omitted defaults to a comma. Delimiter may have more than one
'             character.
' Encoding  : Allowed entries are `ANSI` (the default), `UTF-8` and `UTF-16`. An error will result if this
'             argument is `ANSI` but Data contains characters that cannot be written to an ANSI file.
'             `UTF-8` and `UTF-16` files are written with a byte option mark.
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
        Optional ByVal QuoteAllStrings As Boolean = True, Optional ByVal DateFormat As String = "YYYY-MM-DD", _
        Optional ByVal DateTimeFormat As String = "ISO", Optional ByVal Delimiter As String = ",", _
        Optional ByVal Encoding As String = "ANSI", Optional ByVal EOL As String = vbNullString, _
        Optional TrueString As String = "True", Optional FalseString As String = "False") As String
Attribute CSVWrite.VB_Description = "Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an ""error string"" (starts with `#`, ends with `!`) describing what went wrong."
Attribute CSVWrite.VB_ProcData.VB_Invoke_Func = " \n14"

          Const Err_Delimiter1 = "Delimiter must have at least one character"
          Const Err_Delimiter2 As String = "Delimiter cannot start with a " & _
              "double quote, line feed or carriage return"
              
          Const Err_Dimensions As String = "Data has more than two dimensions, which is not supported"
          Const Err_Encoding As String = "Encoding must be ""ANSI"" (the default) or ""UTF-8"" or ""UTF-16"""
          
          Dim EOLIsWindows As Boolean
          Dim i As Long
          Dim j As Long
          Dim Lines() As String
          Dim OneLine() As String
          Dim OneLineJoined As String
          Dim Stream As Object
          Dim Unicode As Boolean
          Dim WriteToFile As Boolean

1         On Error GoTo ErrHandler
          
2         WriteToFile = Len(FileName) > 0
          
3         If WriteToFile Then
4             Select Case UCase$(Encoding)
                  Case ""
5                     Encoding = "ANSI"
6                 Case "ANSI", "UTF-8", "UTF-16"
7                 Case Else
8                     Throw Err_Encoding
9             End Select
10        End If
          
11        If Len(Delimiter) = 0 Then
12            Throw Err_Delimiter1
13        End If
14        If Left$(Delimiter, 1) = DQ Or Left$(Delimiter, 1) = vbLf Or Left$(Delimiter, 1) = vbCr Then
15            Throw Err_Delimiter2
16        End If
          
17        ValidateTrueAndFalseStrings TrueString, FalseString, Delimiter

18        WriteToFile = Len(FileName) > 0

19        If EOL = vbNullString Then
20            If WriteToFile Then
21                EOL = vbCrLf
22            Else
23                EOL = vbLf
24            End If
25        End If

26        EOL = OStoEOL(EOL, "EOL")
27        EOLIsWindows = EOL = vbCrLf
          
28        If DateFormat = "" Or UCase(DateFormat) = "ISO" Then
              'Avoid DateFormat being the null string as that would make CSVWrite's _
               behaviour depend on Windows locale (via calls to Format$ in function Encode).
29            DateFormat = "yyyy-mm-dd"
30        End If
          
31        Select Case UCase$(DateTimeFormat)
              Case "ISO", ""
32                DateTimeFormat = "yyyy-mm-ddThh:mm:ss"
33            Case "ISOZ"
34                DateTimeFormat = ISOZFormatString()
35        End Select

36        If TypeName(Data) = "Range" Then
              'Preserve elements of type Date by using .Value, not .Value2
37            Data = Data.value
38        End If
39        Select Case NumDimensions(Data)
              Case 0
                  Dim Tmp() As Variant
40                ReDim Tmp(1 To 1, 1 To 1)
41                Tmp(1, 1) = Data
42                Data = Tmp
43            Case 1
44                ReDim Tmp(LBound(Data) To UBound(Data), 1 To 1)
45                For i = LBound(Data) To UBound(Data)
46                    Tmp(i, 1) = Data(i)
47                Next i
48                Data = Tmp
49            Case Is > 2
50                Throw Err_Dimensions
51        End Select
          
52        ReDim OneLine(LBound(Data, 2) To UBound(Data, 2))
          
53        If WriteToFile Then
54            If UCase$(Encoding) = "UTF-8" Then
55                Set Stream = CreateObject("ADODB.Stream")
56                Stream.Open
57                Stream.Type = 2 'Text
58                Stream.CharSet = "utf-8"
          
59                For i = LBound(Data) To UBound(Data)
60                    For j = LBound(Data, 2) To UBound(Data, 2)
61                        OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, Delimiter, TrueString, FalseString)
62                    Next j
63                    OneLineJoined = VBA.Join(OneLine, Delimiter) & EOL
64                    Stream.WriteText OneLineJoined
65                Next i
66                Stream.SaveToFile FileName, 2 'adSaveCreateOverWrite

67                CSVWrite = FileName
68            Else
69                Unicode = UCase$(Encoding) = "UTF-16"
70                If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
                  Dim EN As Long, ED As String
71                On Error Resume Next
72                Set Stream = m_FSO.CreateTextFile(FileName, True, Unicode)
73                EN = Err.Number: ED = Err.Description
74                On Error GoTo ErrHandler
75                If EN <> 0 Then Throw "Error '" & ED & "' when attempting to create file '" + FileName + "'"
        
76                For i = LBound(Data) To UBound(Data)
77                    For j = LBound(Data, 2) To UBound(Data, 2)
78                        OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, Delimiter, TrueString, FalseString)
79                    Next j
80                    OneLineJoined = VBA.Join(OneLine, Delimiter)
81                    WriteLineWrap Stream, OneLineJoined, EOLIsWindows, EOL, Unicode
82                Next i

83                Stream.Close: Set Stream = Nothing
84                CSVWrite = FileName
85            End If
86        Else

87            ReDim Lines(LBound(Data) To UBound(Data) + 1) 'add one to ensure that result has a terminating EOL
        
88            For i = LBound(Data) To UBound(Data)
89                For j = LBound(Data, 2) To UBound(Data, 2)
90                    OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, Delimiter, TrueString, FalseString)
91                Next j
92                Lines(i) = VBA.Join(OneLine, Delimiter)
93            Next i
94            CSVWrite = VBA.Join(Lines, EOL)
95            If Len(CSVWrite) > 32767 Then
96                If TypeName(Application.Caller) = "Range" Then
97                    Throw "Cannot return string of length " & Format$(CStr(Len(CSVWrite)), "#,###") & _
                          " to a cell of an Excel worksheet"
98                End If
99            End If
100       End If
          
101       Exit Function
ErrHandler:

102       If Not Stream Is Nothing Then
103           Stream.Close
104           Set Stream = Nothing
105       End If
          
106       CSVWrite = ReThrow("CSVWrite", Err, m_ErrorStyle = es_ReturnString)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ValidateTrueAndFalseStrings
' Purpose    : Stop the user from making bad choices for either TrueString or FalseString, e.g: strings that would be
'              interpreted as (the wrong) Boolean, or as numbers, dates or empties, strings containing line feed
'              characters, containing the delimiter etc.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ValidateTrueAndFalseStrings(TrueString As String, FalseString As String, Delimiter As String)
             
1         If LCase$(TrueString) = "true" Then
2             If LCase$(FalseString) = "false" Then
3                 Exit Function
4             End If
5         End If
          
6         If LCase$(TrueString) = "false" Then Throw "TrueString cannot take the value '" & TrueString & "'"
7         If LCase$(FalseString) = "true" Then Throw "FalseString cannot take the value '" & FalseString & "'"

8         If TrueString = FalseString Then
9             Throw "Got '" & TrueString & "' for both TrueString and FalseString, but these cannot be equal to one another"
10        End If
          
11        ValidateBooleanRepresentation TrueString, "TrueString", Delimiter
12        ValidateBooleanRepresentation FalseString, "FalseString", Delimiter
          
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
          Dim SysDateOrder As Long
          Dim SysDateSeparator As String
              
1         SysDateOrder = Application.International(xlDateOrder)
2         SysDateSeparator = Application.International(xlDateSeparator)

3         If strValue = "" Then Throw strName & " cannot be the zero-length string"

4         If InStr(strValue, vbLf) > 0 Then Throw strName & " contains a line feed character (ascii 10), which is not permitted"
5         If InStr(strValue, vbCr) > 0 Then Throw strName & " contains a carriage return character (ascii 13), which is not permitted"
6         If InStr(strValue, Delimiter) > 0 Then Throw strName & " contains Delimiter '" & Delimiter & "' which is not permitted"
7         If InStr(strValue, DQ) > 0 Then
8             DQCount = Len(strValue) - Len(Replace(strValue, DQ, vbNullString))
9             If DQCount <> 2 Or Left$(strValue, 1) <> DQ Or Right$(strValue, 1) <> DQ Then
10                Throw "When " & strName & " contains any double quote characters they must be at the start, the end and nowhere else"
11            End If
12        End If
              
13        If IsNumeric(strValue) Then Throw "Got '" & strValue & "' as " & strName & " but that's not valid because it represents a number"
              
14        For i = 1 To 3
15            For Each DateSeparator In Array("/", "-", " ")
16                CastToDate strValue, DtOut, i, _
                      CStr(DateSeparator), SysDateOrder, SysDateSeparator, Converted
17                If Converted Then
18                    Throw "Got '" & strValue & "' as " & _
                          strName & " but that's not valid because it represents a date"
19                End If
20            Next
21        Next

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
23                        If Len(Replace(InnerPart, DQ & DQ, "")) <> Len(Replace(InnerPart, DQ, "")) Then
24                            DQsGood = False
25                        End If
26                    End If
27                End If
28            End If
29        End If

30        If HasCR Or HasLF Or HasDelim Or HasDQ Then
31            If Not DQsGood Then
32                Throw "Got '" & Replace(Replace(FieldValue, vbCr, "<CR>"), vbLf, "<LF>") & "' as " & _
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
' Procedure  : Encode
' Purpose    : Encode arbitrary value as a string, sub-routine of CSVWrite.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Encode(ByVal x As Variant, ByVal QuoteAllStrings As Boolean, ByVal DateFormat As String, _
    ByVal DateTimeFormat As String, ByVal Delim As String, TrueString As String, FalseString As String) As String
          
1         On Error GoTo ErrHandler
2         Select Case VarType(x)

              Case vbString
3                 If InStr(x, DQ) > 0 Then
4                     Encode = DQ & Replace$(x, DQ, DQ2) & DQ
5                 ElseIf QuoteAllStrings Then
6                     Encode = DQ & x & DQ
7                 ElseIf InStr(x, vbCr) > 0 Then
8                     Encode = DQ & x & DQ
9                 ElseIf InStr(x, vbLf) > 0 Then
10                    Encode = DQ & x & DQ
11                ElseIf InStr(x, Delim) > 0 Then
12                    Encode = DQ & x & DQ
13                Else
14                    Encode = x
15                End If
16            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbEmpty  'vbLongLong - not available on 16 bit.
17                Encode = CStr(x)
18            Case vbBoolean
19                Encode = IIf(x, TrueString, FalseString)
20            Case vbDate
21                If CLng(x) = CDbl(x) Then
22                    Encode = Format$(x, DateFormat)
23                Else
24                    Encode = Format$(x, DateTimeFormat)
25                End If
26            Case vbNull
27                Encode = "NULL"
28            Case vbError
29                Select Case CStr(x) 'Editing this case statement? Edit also its inverse, see method MakeSentinels
                      Case "Error 2000"
30                        Encode = "#NULL!"
31                    Case "Error 2007"
32                        Encode = "#DIV/0!"
33                    Case "Error 2015"
34                        Encode = "#VALUE!"
35                    Case "Error 2023"
36                        Encode = "#REF!"
37                    Case "Error 2029"
38                        Encode = "#NAME?"
39                    Case "Error 2036"
40                        Encode = "#NUM!"
41                    Case "Error 2042"
42                        Encode = "#N/A"
43                    Case "Error 2043"
44                        Encode = "#GETTING_DATA!"
45                    Case "Error 2045"
46                        Encode = "#SPILL!"
47                    Case "Error 2046"
48                        Encode = "#CONNECT!"
49                    Case "Error 2047"
50                        Encode = "#BLOCKED!"
51                    Case "Error 2048"
52                        Encode = "#UNKNOWN!"
53                    Case "Error 2049"
54                        Encode = "#FIELD!"
55                    Case "Error 2050"
56                        Encode = "#CALC!"
57                    Case Else
58                        Encode = CStr(x)        'should never hit this line...
59                End Select
60            Case Else
61                Throw "Cannot convert variant of type " & TypeName(x) & " to String"
62        End Select
63        Exit Function
ErrHandler:
64        ReThrow "Encode", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WriteLineWrap
' Purpose    : Wrapper to TextStream.Write[Line] to give more informative error message than "invalid procedure call or
'              argument" if the error is caused by attempting to write illegal characters to a stream opened with
'              TriStateFalse.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub WriteLineWrap(T As TextStream, text As String, EOLIsWindows As Boolean, EOL As String, Unicode As Boolean)

          Dim ErrDesc As String
          Dim ErrNum As Long
          Dim i As Long

1         On Error GoTo ErrHandler
2         If EOLIsWindows Then
3             T.WriteLine text
4         Else
5             T.Write text
6             T.Write EOL
7         End If

8         Exit Sub

ErrHandler:
9         ErrNum = Err.Number
10        If Not Unicode Then
11            If ErrNum = 5 Then
12                For i = 1 To Len(text)
13                    If Not CanWriteCharToAscii(Mid$(text, i, 1)) Then
14                        ErrDesc = "Data contains characters that cannot be written to an ascii file (first found is '" & _
                              Mid$(text, i, 1) & "' with unicode character code " & AscW(Mid$(text, i, 1)) & _
                              "). Try calling CSVWrite with argument Encoding as ""UTF-8"" or ""UTF-16"""
15                        Throw ErrDesc
16                        Exit For
17                    End If
18                Next i
19            End If
20        End If
21        ReThrow "WriteLineWrap", Err
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
1         On Error GoTo ErrHandler
2         If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
3         GetFileSize = m_FSO.GetFile(FilePath).Size

4         Exit Function
ErrHandler:
5         Throw "Could not find file '" & FilePath & "'"
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

7         FolderPath = Replace(FolderPath, "/", "\")

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
' Procedure  : Throw
' Purpose    : Error handling - companion to ReThrow
' Parameters :
'  Description  : Description of what went wrong.
'  WithCallStack: Should subsequent calls to ReThrow append the names of the functions and line numbers in the call
'                 stack?
'                 For anticipated errors (which Throw will be responsible for) this should (usually) be False (the
'                 default)to avoid cluttering the error description.
'                 But for unanticipated errors (probably not generated by Throw) it's very useful to see an error
'                 description at the top of the call stack that includes information on the functions and line numbers,
'                 since that greatly speeds up debugging.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Throw(ByVal Description As String, Optional WithCallStack As Boolean = False)
          'ErrorNumber being vbObjectError + 100 suppresses annotation in ReThrow
1         Err.Raise vbObjectError + IIf(WithCallStack, 1, 100), , Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReThrow
' Purpose    : Common error handling to be used in the error handler of all methods.
' Parameters :
'  FunctionName: The name
'  Error       : The Err error object
'  TopLevel    : Pass in True if the method is a "top level" method that's exposed to the user and we wish for the function
'                to return an error string (starts with #, ends with !).
'                Pass in False if we want to (re)throw an error, anotated as long as ErrorNumber is not vbObjectError + 100
' -----------------------------------------------------------------------------------------------------------------------
Private Function ReThrow(FunctionName As String, Error As ErrObject, Optional TopLevel As Boolean = False)

          Dim ErrorDescription As String
          Dim ErrorNumber As Long
          Dim LineDescription As String
          Dim ShowCallStack As Boolean
          
1         ErrorDescription = Error.Description
2         ErrorNumber = Err.Number
3         ShowCallStack = ErrorNumber <> vbObjectError + 100
          
4         If ShowCallStack Or TopLevel Then
              'Build up call stack, i.e. annotate error description by prepending #<FunctionName> and appending !
5             If Erl <> 0 And ShowCallStack Then
                  'Code has line numbers, annotate with line number
6                 LineDescription = " (line " & CStr(Erl) & "): "
7             Else
8                 LineDescription = ": "
9             End If
10            ErrorDescription = "#" & FunctionName & LineDescription & ErrorDescription & "!"
11        End If

12        If TopLevel Then
13            ReThrow = ErrorDescription
14        Else
15            Err.Raise ErrorNumber, , ErrorDescription
16        End If
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead
' Purpose    : Register the function CSVRead with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub RegisterCSVRead()
          Const Description As String = "Returns the contents of a comma-separated file on disk as an array."
          Dim ArgDescs() As String

1         On Error GoTo ErrHandler

2         ReDim ArgDescs(1 To 19)
3         ArgDescs(1) = "The full name of the file, including the path, or else a URL of a file, or else a string in CSV " & _
              "format."
4         ArgDescs(2) = "Type conversion: Boolean or string. Allowed letters NDBETQ. N = convert Numbers, D = convert " & _
              "Dates, B = convert Booleans, E = convert Excel errors, T = trim leading & trailing spaces, Q = " & _
              "quoted fields also converted. TRUE = NDB, FALSE = no conversion."
5         ArgDescs(3) = "Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe found " & _
              "outside quoted regions within the first 10,000 characters. Enter FALSE to  see the file's " & _
              "contents as would be displayed in a text editor."
6         ArgDescs(4) = "Whether delimiters which appear at the start of a line, the end of a line or immediately after " & _
              "another delimiter should be ignored while parsing; useful for fixed-width files with delimiter " & _
              "padding between fields."
7         ArgDescs(5) = "The format of dates in the file such as `Y-M-D` (the default), `M-D-Y` or `Y/M/D`. Also `ISO` " & _
              "for ISO8601 (e.g., 2021-08-26T09:11:30) or `ISOZ` (time zone given e.g. " & _
              "2021-08-26T13:11:30+05:00), in which case dates-with-time are returned in UTC time."
8         ArgDescs(6) = "Rows that start with this string will be skipped while parsing."
9         ArgDescs(7) = "Whether empty rows/lines in the file should be skipped while parsing (if `FALSE`, each column " & _
              "will be assigned ShowMissingsAs for that empty row)."
10        ArgDescs(8) = "The row in the file containing headers. Optional and defaults to 0. Type conversion is not " & _
              "applied to fields in the header row, though leading and trailing spaces are trimmed."
11        ArgDescs(9) = "The first row in the file that's included in the return. Optional and defaults to one more than " & _
              "HeaderRowNum."
12        ArgDescs(10) = "The column in the file at which reading starts. Optional and defaults to 1 to read from the " & _
              "first column."
13        ArgDescs(11) = "The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the " & _
              "end of the file are read."
14        ArgDescs(12) = "The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol " & _
              "are read."
15        ArgDescs(13) = "Indicates how `TRUE` values are represented in the file. May be a string, an array of strings " & _
              "or a range containing strings; by default, `TRUE`, `True` and `true` are recognised."
16        ArgDescs(14) = "Indicates how `FALSE` values are represented in the file. May be a string, an array of strings " & _
              "or a range containing strings; by default, `FALSE`, `False` and `false` are recognised."
17        ArgDescs(15) = "Indicates how missing values are represented in the file. May be a string, an array of strings " & _
              "or a range containing strings. By default, only an empty field (consecutive delimiters) is " & _
              "considered missing."
18        ArgDescs(16) = "Fields which are missing in the file (consecutive delimiters) or match one of the " & _
              "MissingStrings are returned in the array as ShowMissingsAs. Defaults to Empty, but the null " & _
              "string or `#N/A!` error value can be good alternatives."
19        ArgDescs(17) = "Allowed entries are `ASCII`, `ANSI`, `UTF-8`, or `UTF-16`. For most files this argument can be " & _
              "omitted and CSVRead will detect the file's encoding."
20        ArgDescs(18) = "The character that represents a decimal point. If omitted, then the value from Windows " & _
              "regional settings is used."
21        ArgDescs(19) = "For use from VBA only."
22        Application.MacroOptions "CSVRead", Description, , , , , , , , , ArgDescs
23        Exit Sub

ErrHandler:
24        Debug.Print "Warning: Registration of function CSVRead failed with error: " & Err.Description
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

1         On Error GoTo ErrHandler

2         ReDim ArgDescs(1 To 10)
3         ArgDescs(1) = "An array of data, or an Excel range. Elements may be strings, numbers, dates, Booleans, empty, " & _
              "Excel errors or null values. Data typically has two dimensions, but if Data has only one " & _
              "dimension then the output file has a single column, one field per row."
4         ArgDescs(2) = "The full name of the file, including the path. Alternatively, if FileName is omitted, then the " & _
              "function returns Data converted CSV-style to a string."
5         ArgDescs(3) = "If TRUE (the default) then all strings in Data are quoted before being written to file. If " & _
              "FALSE only strings containing Delimiter, line feed, carriage return or double quote are quoted. " & _
              "Double quotes are always escaped by another double quote."
6         ArgDescs(4) = "A format string that determines how dates, including cells formatted as dates, appear in the " & _
              "file. If omitted, defaults to `yyyy-mm-dd`."
7         ArgDescs(5) = "Format for datetimes. Defaults to `ISO` which abbreviates `yyyy-mm-ddThh:mm:ss`. Use `ISOZ` for " & _
              "ISO8601 format with time zone the same as the PC's clock. Use with care, daylight saving may be " & _
              "inconsistent across the datetimes in data."
8         ArgDescs(6) = "The delimiter string, if omitted defaults to a comma. Delimiter may have more than one " & _
              "character."
9         ArgDescs(7) = "Allowed entries are `ANSI` (the default), `UTF-8` and `UTF-16`. An error will result if this " & _
              "argument is `ANSI` but Data contains characters that cannot be written to an ANSI file. `UTF-8` " & _
              "and `UTF-16` files are written with a byte option mark."
10        ArgDescs(8) = "Sets the file's line endings. Enter `Windows`, `Unix` or `Mac`. Also supports the line-ending " & _
              "characters themselves or the strings `CRLF`, `LF` or `CR`. The default is `Windows` if FileName " & _
              "is provided, or `Unix` if not."
11        ArgDescs(9) = "How the Boolean value True is to be represented in the file. Optional, defaulting to ""True""."
12        ArgDescs(10) = "How the Boolean value False is to be represented in the file. Optional, defaulting to " & _
              """False""."
13        Application.MacroOptions "CSVWrite", Description, , , , , , , , , ArgDescs
14        Exit Sub

ErrHandler:
15        Debug.Print "Warning: Registration of function CSVWrite failed with error: " + Err.Description
End Sub

