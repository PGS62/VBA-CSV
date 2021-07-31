Attribute VB_Name = "modCSV"
Option Explicit
Private Const DQ = """"
Private Const DQ2 = """"""
Private Const Err_EmptyFile = "File is empty"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead
' Author     : Philip Swannell
' Date       : 30-Jul-2021
' Purpose    : Register the function CSVRead with the Excel Function Wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Sub RegisterCSVRead()
    Const FnDesc = "Returns the contents of a comma-separated file on disk as an array."
    Dim ArgDescs() As String
    ReDim ArgDescs(1 To 13)
    ArgDescs(1) = "The full name of the file, including the path."
    ArgDescs(2) = "TRUE to convert Numbers, Dates, Logicals and Excel Errors into their typed values, or FALSE to leave as strings. For more control enter a string containing the letters N,D,L, and E eg ""NL"" to convert just numbers and logicals, not dates or errors."
    ArgDescs(3) = "Delimiter character. Defaults to the first instance of comma, tab, semi-colon, colon or vertical bar found in the file outside quoted regions. Enter FALSE to to see the file's raw contents as would be displayed in a text editor."
    ArgDescs(4) = "The format of dates in the file such as D-M-Y, M-D-Y or Y/M/D. If omitted a value is read from Windows regional settings. Repeated D's (or M's or Y's) are equivalent to single instances, so that d-m-y and DD-MMM-YYYY are equivalent."
    ArgDescs(5) = "The ""row"" in the file at which reading starts. Optional and defaults to 1 to read from the first row."
    ArgDescs(6) = "The ""column"" in the file at which reading starts. Optional and defaults to 1 to read from the first column."
    ArgDescs(7) = "The number of rows to read from the file. If omitted (or zero), all rows from StartRow to the end of the file are read."
    ArgDescs(8) = "The number of columns to read from the file. If omitted (or zero), all columns from StartColumn are read."
    ArgDescs(9) = """Windows"" (or ascii 13 + 10, or FALSE), ""Unix"" (or ascii 10, or TRUE) or ""Mac"" (or ascii 13)  to specify line-endings, or ""Mixed"" for inconsistent line endings. Omit to infer from the first line ending found outside quoted regions."
    ArgDescs(10) = "Enter TRUE if the file is unicode, FALSE if the file is ascii. Omit to guess (via function sFileIsUnicode)."
    ArgDescs(11) = "Value to represent empty fields (successive delimiters) in the file. May be a string or an Empty value. Optional and defaults to the zero-length string."
    ArgDescs(12) = "The character that represents a decimal point. If omitted, then the value from Windows regional settings is used."
    ArgDescs(13) = "For use from VBA, sets the lower bound for indexing into the returned array and must be 1 or 0. Optional, defaults to 1."
    Application.MacroOptions "CSVRead", FnDesc, , , , , , , , , ArgDescs
End Sub


' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVWrite
' Author     : Philip Swannell
' Date       : 31-Jul-2021
' Purpose    : Register the function CSVWrite with the Excel Function Wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Sub RegisterCSVWrite()
    Const FnDesc = "Creates a csv file on disk containing the data in the array Data. Any existing file of the same name is overwritten. If successful, the function returns the name of the file written, otherwise an error string. "
    Dim ArgDescs() As String
    ReDim ArgDescs(1 To 9)
    ArgDescs(1) = "The full name of the file, including the path."
    ArgDescs(2) = "An array of arbitrary data. Elements may be strings, numbers, dates, Booleans, empty, Excel errors or null values."
    ArgDescs(3) = "If TRUE (the default) then ALL strings are quoted before being written to file. Otherwise (FALSE) only strings containing characters comma, line feed, carriage return or double quote are quoted. Double quotes are always escaped by a second double quote."
    ArgDescs(4) = "A format string, such as 'yyyy-mm-dd' that determine how dates (e.g. cells formatted as dates) appear in the file."
    ArgDescs(5) = "A format string, such as 'yyyy-mm-dd hh:mm:ss' that determine how elements of dates with time appear in the file. Currently the companion function CVSRead is not capable of interpreting fields written in DateTime format."
    ArgDescs(6) = "If TRUE then the file written is encoded as unicode. Defaults to FALSE for an ascii file."
    ArgDescs(7) = "If FALSE (the default) the file written will be ascii. If TRUE the file written will be Unicode."
    ArgDescs(8) = "Enter the required line ending character as ""Windows"" (or ascii 13 plus ascii 10), or ""Unix"" (or ascii 10) or ""Mac"" (or ascii 13). If omitted defaults to ""Windows""."
    ArgDescs(9) = "This argument is for develop purposes only, it will soon be removed."
    Application.MacroOptions "CSVWrite", FnDesc, , , , , , , , , ArgDescs
End Sub




Private Function CreateCSVStream(FileName As String, Optional ByVal EOL As String, Optional ByVal FileIsUnicode As Variant) As clsCSVStream
          Dim CSVS As New clsCSVStream

1         On Error GoTo ErrHandler
2         If VarType(FileIsUnicode) <> vbBoolean Then
3             FileIsUnicode = ThrowIfError(IsUnicodeFile(FileName))
4         End If
5         If EOL = "" Then
6             EOL = InferEOL(FileName, CBool(FileIsUnicode))
7         End If

8         CSVS.Init FileName, EOL, CBool(FileIsUnicode)
9         Set CreateCSVStream = CSVS

10        Exit Function
ErrHandler:
11        Throw "#CreateCSVStream (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Sub Throw(ByVal ErrorString As String)
1         Err.Raise vbObjectError + 1, , ErrorString
End Sub

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVRead
' Author    : Philip Swannell
' Date      : 30-Jul-2021
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path.
' TypeConversion: TRUE to convert Numbers, Dates, Logicals and Excel Errors into their typed values, or
'             FALSE to leave as strings. For more control enter a string containing the
'             letters N,D,L, and E eg "NL" to convert just numbers and logicals, not dates
'             or errors.
' Delimiter : Delimiter character. Defaults to the first instance of comma, tab, semi-colon, colon or
'             vertical bar found in the file outside quoted regions. Enter FALSE to to see
'             the file's raw contents as would be displayed in a text editor.
' DateFormat: The format of dates in the file such as D-M-Y, M-D-Y or Y/M/D. If omitted a value is read
'             from Windows regional settings. Repeated D's (or M's or Y's) are equivalent
'             to single instances, so that d-m-y and DD-MMM-YYYY are equivalent.
' StartRow  : The "row" in the file at which reading starts. Optional and defaults to 1 to read from the
'             first row.
' StartCol  : The "column" in the file at which reading starts. Optional and defaults to 1 to read from
'             the first column.
' NumRows   : The number of rows to read from the file. If omitted (or zero), all rows from StartRow to
'             the end of the file are read.
' NumCols   : The number of columns to read from the file. If omitted (or zero), all columns from
'             StartColumn are read.
' LineEndings: "Windows" (or ascii 13 + 10, or FALSE), "Unix" (or ascii 10, or TRUE) or "Mac" (or ascii
'             13)  to specify line-endings, or "Mixed" for inconsistent line endings. Omit
'             to infer from the first line ending found outside quoted regions.
' Unicode   : Enter TRUE if the file is unicode, FALSE if the file is ascii. Omit to guess (via function
'             sFileIsUnicode).
' ShowMissingsAs: Value to represent empty fields (successive delimiters) in the file. May be a string or an
'             Empty value. Optional and defaults to the zero-length string.
' DecimalSeparator: The character that represents a decimal point. If omitted, then the value from Windows
'             regional settings is used.
' LowerBounds: For use from VBA, sets the lower bound for indexing into the returned array and must be 1
'             or 0. Optional, defaults to 1.
'
' Notes     : See also sFileSaveCSV for which this function is the inverse.
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
Function CSVRead(FileName As String, Optional TypeConversion As Variant = False, Optional ByVal Delimiter As Variant, _
          Optional DateFormat As String, Optional ByVal StartRow As Long = 1, Optional ByVal StartCol As Long = 1, _
          Optional ByVal NumRows As Long = 0, Optional ByVal NumCols As Long = 0, Optional ByVal LineEndings As Variant, _
          Optional ByVal Unicode As Variant, Optional ByVal ShowMissingsAs As Variant = "", _
          Optional DecimalSeparator As String = vbNullString, Optional LowerBounds As Variant = 1)
Attribute CSVRead.VB_Description = "Returns the contents of a comma-separated file on disk as an array."
Attribute CSVRead.VB_ProcData.VB_Invoke_Func = " \n14"

          Const Err_Delimiter = "Delimiter character must be passed as a string, FALSE for no delimiter, or else omitted to infer from the file's contents"
          Const Err_FileIsUniCode = "Unicode must be passed as TRUE or FALSE, or omitted to infer from the file's contents"
          Const Err_InFuncWiz = "#Disabled in Function Dialog!"
          Const Err_LineEndings = "LineEndings must be one of ""Windows"", ""Unix"" or ""Mac"", or omitted to infer from the file's contents"
          Const Err_LowerBounds = "LowerBounds must be 0 or 1. The argument determines whether the array returned uses zero-based or 1-based indexing"
          Const Err_NumCols = "NumCols must be positive to read a given number or columns, or zero or omitted to read all columns from StartCol to the maximum column encountered."
          Const Err_NumRows = "NumRows must be positive to read a given number or rows, or zero or omitted to read all rows from StartRow to the end of the file."
          Const Err_Seps = "DecimalSeparator must be different from Delimiter"
          Const Err_StartCol = "StartCol must be at least 1."
          Const Err_StartRow = "StartRow must be at least 1."
          
          Dim AltDelimiter As String
          Dim AnyConversion As Boolean
          Dim CSVS As clsCSVStream
          Dim DateOrder As Long
          Dim DateSeparator As String
          Dim EOL As String
          Dim F As Scripting.File
          Dim FSO As New Scripting.FileSystemObject
          Dim i As Long
          Dim j As Long
          Dim LB As Long
          Dim Lines() As String
          Dim MixedLineEndings As Boolean
          Dim NoQuotesInFile As Boolean
          Dim NotDelimited As Boolean
          Dim NumColsInReturn As Long
          Dim NumInRow As Long
          Dim NumRowsInReturn As Long
          Dim OneRow() As String
          Dim p As Long
          Dim q As Long
          Dim ReadAll As String
          Dim RemoveQuotes As Boolean
          Dim ReturnArray() As Variant
          Dim SepsStandard As Boolean
          Dim ShowDatesAsDates As Boolean
          Dim ShowErrorsAsErrors As Boolean
          Dim ShowLogicalsAsLogicals As Boolean
          Dim ShowMissingAsNullString As Boolean
          Dim ShowNumbersAsNumbers As Boolean
          Dim SplitLimit As Long
          Dim strDelimiter As String
          Dim SysDateOrder As Long
          Dim SysDateSeparator As String
          Dim SysDecimalSeparator As String
          Dim T As Scripting.TextStream
          
1         On Error GoTo ErrHandler

2         If FunctionWizardActive() Then
3             CSVRead = Err_InFuncWiz
4             Exit Function
5         End If

          'Parse and validate inputs...
6         If IsEmpty(Unicode) Or IsMissing(Unicode) Then
7             Unicode = ThrowIfError(IsUnicodeFile(FileName))
8         ElseIf VarType(Unicode) <> vbBoolean Then
9             Throw Err_FileIsUniCode
10        End If
          
11        If TypeName(LineEndings) = "Range" Then LineEndings = LineEndings.Value
12        If IsMissing(LineEndings) Or IsEmpty(LineEndings) Then
13            EOL = InferEOL(FileName, CBool(Unicode))
14        ElseIf VarType(LineEndings) = vbString Then
15            Select Case LCase(LineEndings)
                  Case "windows", vbCrLf
16                    EOL = vbCrLf
17                Case "unix", vbLf
18                    EOL = vbLf
19                Case "mac", vbCr
20                    EOL = vbCr
21                Case "mixed"
22                    MixedLineEndings = True
23                Case Else
24                    Throw Err_LineEndings
25            End Select
26        ElseIf VarType(LineEndings) = vbBoolean Then
              'For backward compatibility - LineEndings was f.k.a. FileIsUnix
27            If LineEndings Then
28                EOL = vbLf
29            Else
30                EOL = vbCrLf
31            End If
32        Else
33            Throw Err_LineEndings
34        End If

35        If VarType(Delimiter) = vbBoolean Then
36            If Not Delimiter Then
37                NotDelimited = True
38            Else
39                Throw Err_Delimiter
40            End If
41        ElseIf VarType(Delimiter) = vbString Then
42            strDelimiter = Delimiter
43        ElseIf IsEmpty(Delimiter) Or IsMissing(Delimiter) Then
44            strDelimiter = InferDelimiter(FileName, CBool(Unicode))
45        Else
46            Throw Err_Delimiter
47        End If

48        ParseTypeConversion TypeConversion, ShowNumbersAsNumbers, _
              ShowDatesAsDates, ShowLogicalsAsLogicals, ShowErrorsAsErrors, RemoveQuotes

49        If ShowNumbersAsNumbers Then
50            If ((DecimalSeparator = Application.DecimalSeparator) Or DecimalSeparator = vbNullString) Then
51                SepsStandard = True
52            ElseIf DecimalSeparator = strDelimiter Then
53                Throw Err_Seps
54            End If
55        End If

56        If ShowDatesAsDates Then
57            ParseDateFormat DateFormat, DateOrder, DateSeparator
58            SysDateOrder = Application.International(xlDateOrder)
59            SysDateSeparator = Application.International(xlDateSeparator)
60        End If

61        If StartRow < 1 Then Throw Err_StartRow
62        If StartCol < 1 Then Throw Err_StartCol
63        If NumRows < 0 Then Throw Err_NumRows
64        If NumCols < 0 Then Throw Err_NumCols

65        If TypeName(ShowMissingsAs) = "Range" Then
66            ShowMissingsAs = ShowMissingsAs.Value
67        End If
68        If Not (IsEmpty(ShowMissingsAs) Or VarType(ShowMissingsAs) = vbString) Then
69            Throw "ShowMissingsAs must be Empty or a string"
70        End If
71        If VarType(ShowMissingsAs) = vbString Then
72            ShowMissingAsNullString = ShowMissingsAs = ""
73        End If
          If TypeName(LowerBounds) = "Range" Then LowerBounds = LowerBounds.Value
74        If IsMissing(LowerBounds) Or IsEmpty(LowerBounds) Then
75            LB = 1
76        ElseIf IsNumeric(LowerBounds) Then
77            LB = CLng(LowerBounds)
78            If LB <> 0 And LB <> 1 Then Throw Err_LowerBounds
79        Else
80            Throw Err_LowerBounds
81        End If
          'End of input validation
                
82        If NotDelimited Then
83            CSVRead = ShowTextFile(FileName, StartRow, NumRows, MixedLineEndings, EOL, CBool(Unicode))
84            Exit Function
85        End If
                
          'In this case (reading the entire file) performance is better if we don't use _
           clsCSVStream but instead use method SplitContents on the entire file contents.
86        If StartRow = 1 And StartCol = 1 And NumRows = 0 And NumCols = 0 Then
87            Set F = FSO.GetFile(FileName)
88            Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))
89            If T.AtEndOfStream Then
90                T.Close: Set T = Nothing: Set F = Nothing
91                Throw Err_EmptyFile
92            End If

93            ReadAll = T.ReadAll
94            T.Close: Set T = Nothing: Set F = Nothing
95            Lines = SplitContents(ReadAll, EOL, AltDelimiter, NoQuotesInFile)

96            If NoQuotesInFile Then
97                RemoveQuotes = False
98            End If
99        Else
100           Set CSVS = CreateCSVStream(FileName, EOL, Unicode)
101           For i = 1 To StartRow - 1
102               CSVS.ReadLine
103           Next i
104           CSVS.StartRecording
105           If NumRows > 0 Then
106               For i = 1 To NumRows
107                   CSVS.ReadLine
108               Next
109           Else
110               While Not CSVS.AtEndOfStream
111                   CSVS.ReadLine
112               Wend
113           End If

114           Lines = CSVS.ReportAllLinesRead()
115           If Not CSVS.QuotesEncountered Then
116               RemoveQuotes = False
117           End If
118           Set CSVS = Nothing
119       End If
120       NumRowsInReturn = UBound(Lines) - LBound(Lines) + 1
          
121       If NumCols = 0 Then
122           NumColsInReturn = 1
123           SplitLimit = -1
124       Else
125           NumColsInReturn = NumCols
126           SplitLimit = StartCol - 1 + NumCols + 1
127       End If

128       AnyConversion = RemoveQuotes Or ShowNumbersAsNumbers Or ShowDatesAsDates Or _
              ShowLogicalsAsLogicals Or ShowErrorsAsErrors Or (Not ShowMissingAsNullString)

129       LB = LowerBounds
130       ReDim ReturnArray(LB To LB + NumRowsInReturn - 1, LB To LB + NumColsInReturn - 1)

131       For i = LBound(Lines) To UBound(Lines)
132           OneRow = SplitLine(Lines(i), strDelimiter, AltDelimiter, NoQuotesInFile, SplitLimit)
133           NumInRow = UBound(OneRow) - LBound(OneRow) + 1
134           If SplitLimit > 0 Then
135               If NumInRow = SplitLimit Then
136                   NumInRow = SplitLimit - 1
137               End If
138           End If

              'Ragged files: Current line has more elements than maximum length of prior lines. _
               First we need to append columns on the right of ReturnArray (Redim Preserve) then for _
               cells in rows < i, populate the columns just added with ShowMissingsAs..
139           If NumCols = 0 Then
140               If NumInRow - StartCol + 1 > NumColsInReturn Then
141                   ReDim Preserve ReturnArray(LB To LB + NumRowsInReturn - 1, LB To NumInRow - StartCol + LB)
142                   If Not IsEmpty(ShowMissingsAs) Then
143                       For p = 1 To i
144                           For q = NumColsInReturn + 1 To NumInRow
145                               ReturnArray(p + LB - 1, q + LB - 1) = ShowMissingsAs
146                           Next q
147                       Next p
148                   End If
149                   NumColsInReturn = NumInRow - StartCol + 1
150               End If
151           End If

152           If AnyConversion Then
153               For j = 1 To MinLngs(NumColsInReturn, NumInRow - StartCol + 1)
154                   ReturnArray(i + LB, j + LB - 1) = CastToVariant(OneRow(j + StartCol - 2), _
                          RemoveQuotes, ShowNumbersAsNumbers, SepsStandard, DecimalSeparator, SysDecimalSeparator, _
                          ShowDatesAsDates, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, _
                          ShowLogicalsAsLogicals, ShowErrorsAsErrors, ShowMissingAsNullString, ShowMissingsAs)
155               Next j
156           Else
157               For j = 1 To MinLngs(NumColsInReturn, NumInRow - StartCol + 1)
158                   ReturnArray(i + LB, j + LB - 1) = OneRow(j + StartCol - 2)
159               Next j
160           End If

              'Ragged files: Current line has fewer elements than maximum length of prior lines. _
               We need to pad the remainder of the current line with ShowMissingsAs.
161           If NumInRow - StartCol + 1 < NumColsInReturn Then
162               If Not IsEmpty(ShowMissingsAs) Then
163                   For j = NumInRow - StartCol + 2 To NumColsInReturn
164                       ReturnArray(i + LB, j + LB - 1) = ShowMissingsAs
165                   Next j
166               End If
167           End If
168       Next i

169       CSVRead = ReturnArray

170       Exit Function

ErrHandler:
171       CSVRead = "#CSVRead (line " & CStr(Erl) + "): " & Err.Description & "!"
172       If Not CSVS Is Nothing Then
173           Set CSVS = Nothing
174       End If
175       If Not T Is Nothing Then
176           T.Close
177           Set T = Nothing
178       End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseTypeConversion
' Author     : Philip Swannell
' Date       : 29-Jul-2021
' Purpose    : Parse the input TypeConversion to set five Boolean flags which are passed by reference
' Parameters :
'  TypeConversion        :
'  ShowNumbersAsNumbers  : Should fields in the file that look like numbers be returned as Numbers? (Doubles)
'  ShowDatesAsDates      : Should fields in the file that look like dates with the specified DateFormat be returned as Dates?
'  ShowLogicalsAsLogicals: Should fields in the file that are TRUE or FALSE (case insensitive) be returned as Booleans?
'  ShowErrorsAsErrors    : Should fields in the file that looklike Excel errors (#N/A #REF! etc) be returned as errors?
'  RemoveQuotes          : Should quoted fields be unquoted?
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseTypeConversion(ByVal TypeConversion As Variant, ByRef ShowNumbersAsNumbers As Boolean, _
          ByRef ShowDatesAsDates As Boolean, ByRef ShowLogicalsAsLogicals As Boolean, _
          ByRef ShowErrorsAsErrors As Boolean, ByRef RemoveQuotes As Boolean)

          Const Err_TypeConversion = "Enter N to show Numbers as Numbers, D to show Dates as Dates, L to show Logicals as Logicals, E to show Excel Errors as Errors, Q to show Quoted fields with Quotes."
          Dim i As Long

1         On Error GoTo ErrHandler
2         If TypeName(TypeConversion) = "Range" Then
3             TypeConversion = TypeConversion.Value
4         End If

5         If VarType(TypeConversion) = vbBoolean Then
6             If TypeConversion Then
7                 ShowNumbersAsNumbers = True
8                 ShowDatesAsDates = True
9                 ShowLogicalsAsLogicals = True
10                ShowErrorsAsErrors = True
11                RemoveQuotes = True
12            Else
13                ShowNumbersAsNumbers = False
14                ShowDatesAsDates = False
15                ShowLogicalsAsLogicals = False
16                ShowErrorsAsErrors = False
17                RemoveQuotes = True
18            End If
19        ElseIf VarType(TypeConversion) = vbString Then
20            ShowNumbersAsNumbers = False
21            ShowDatesAsDates = False
22            ShowLogicalsAsLogicals = False
23            ShowErrorsAsErrors = False
24            RemoveQuotes = True
25            For i = 1 To Len(TypeConversion)
26                Select Case UCase(Mid(TypeConversion, i, 1))
                      Case "N"
27                        ShowNumbersAsNumbers = True
28                    Case "D"
29                        ShowDatesAsDates = True
30                    Case "L"
31                        ShowLogicalsAsLogicals = True
32                    Case "E"
33                        ShowErrorsAsErrors = True
34                    Case "Q"
35                        RemoveQuotes = False
36                    Case Else
37                        Throw "Unrecognised character '" + Mid(TypeConversion, i, 1) + "' in TypeConversion."
38                End Select
39            Next i
40        Else
41            Throw "TypeConversion must be TRUE (convert all types), FALSE (no conversion) or a string, " + Err_TypeConversion
42        End If

43        Exit Function
ErrHandler:
44        Throw "#ParseTypeConversion (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MinLngs
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Returns the minimum of a & b.
' -----------------------------------------------------------------------------------------------------------------------
Private Function MinLngs(a As Long, b As Long)
1         If a < b Then
2             MinLngs = a
3         Else
4             MinLngs = b
5         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferDelimiter
' Author     : Philip Swannell
' Date       : 14-Dec-2017
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

1         On Error GoTo ErrHandler

2         Set FSO = New FileSystemObject
3         Set F = FSO.GetFile(FileName)
4         Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))

5         If T.AtEndOfStream Then
6             T.Close: Set T = Nothing: Set F = Nothing
7             Throw "File is empty"
8         End If

9         EvenQuotes = True
10        While Not T.AtEndOfStream
11            Contents = T.Read(CHUNK_SIZE)
12            For i = 1 To Len(Contents)
13                Select Case Mid$(Contents, i, 1)
                      Case DQ
14                        EvenQuotes = Not EvenQuotes
15                    Case ",", vbTab, "|", ";", ":"
16                        If EvenQuotes Then
17                            InferDelimiter = Mid$(Contents, i, 1)
18                            T.Close: Set T = Nothing: Set F = Nothing
19                            Exit Function
20                        Else
21                            FoundInEven = True
22                        End If
23                End Select
24            Next i
25        Wend

          'No commonly-used delimiter found in the file outside quoted regions. There are two possibilities: _
          either the file has only one column or some other character has been used, returning comma is _
              equivalent to assuming the former.

26        InferDelimiter = ","

27        Exit Function
ErrHandler:
28        CopyOfErr = "#InferDelimiter (line " & CStr(Erl) + "): " & Err.Description & "!"
29        If Not T Is Nothing Then
30            T.Close
31            Set T = Nothing: Set F = Nothing: Set FSO = Nothing
32        End If
33        Throw CopyOfErr
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferEOL
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Infer the EOL character from the file's contents, setting it to the first of vbCrLf, vbLf or vbCr found
'              outside quoted regions.
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferEOL(FileName As String, Unicode As Boolean) As String
             
          Const CHUNK_SIZE = 1000
          
          Dim EvenQuotes As Boolean
          Dim F As Scripting.File
          Dim FileContents As String
          Dim FSO As New FileSystemObject
          Dim i As Long
          Dim LenChunk As Long
          Dim T As Scripting.TextStream
          Dim CheckFirstCharOfNextChunk As Boolean
          
1         On Error GoTo ErrHandler

2         Set F = FSO.GetFile(FileName)
3         Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))
4         If T.AtEndOfStream Then
5             T.Close: Set T = Nothing: Set F = Nothing
6             Throw Err_EmptyFile
7         End If
          
8         EvenQuotes = True
9         While Not T.AtEndOfStream
10            FileContents = T.Read(CHUNK_SIZE)
11            LenChunk = Len(FileContents)
12            If CheckFirstCharOfNextChunk Then
13                If Left$(FileContents, 1) = vbLf Then
14                    InferEOL = vbCrLf
15                Else
16                    InferEOL = vbCr
17                End If
18                GoTo EarlyExit
19            End If

20            For i = 1 To LenChunk
21                Select Case Mid(FileContents, i, 1)
                      Case DQ
22                        EvenQuotes = Not EvenQuotes
23                    Case vbCr
24                        If EvenQuotes Then
25                            If i < LenChunk Then
26                                If Mid$(FileContents, i + 1, 1) = vbLf Then
27                                    InferEOL = vbCrLf
28                                Else
29                                    InferEOL = vbCr
30                                End If
31                                GoTo EarlyExit
32                            ElseIf T.AtEndOfStream Then 'Mac file with only one line
33                                InferEOL = vbCr
34                                GoTo EarlyExit
35                            Else
36                                CheckFirstCharOfNextChunk = True
37                            End If
38                        End If
39                    Case vbLf
40                        If EvenQuotes Then
41                            InferEOL = vbLf
42                            GoTo EarlyExit
43                        End If
44                End Select
45            Next
46        Wend

          'No end of line exists outside quoted regions, so the file is a single line without a _
           trailing EOL. The guess made for EOL is irrelevant.
47        InferEOL = vbCrLf

EarlyExit:
48        T.Close: Set T = Nothing: Set F = Nothing
49        Exit Function

ErrHandler:
50        Throw "#InferEOL (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SplitLine
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Split one line of a csv file, splitting on those delimiter characters that are preceded by an odd number
'              of double quote characters.
' Parameters :
'  Line          : The string to be split.
'  Delimiter     : The delimiter.
'  Limit         : Passed to VBA.split in the event that we are reading some but not all columns of a file.
'  AltDelimiter  : A 1-character string known not to be contained in Line.
'  NoQuotesInFile: Pass as true if there are no double quotes in the file and therefore necessarily no double quotes
'                  in Line.
'  SplitLimit    : For a speed up when we are reading fewer than all the columns in the file has to be passed as
'                  StartCol + NumCols, i.e. one more than the number of columns we are interested in since the
'                  remaining columns appear (unsplit) in the last element of the return. For all columns pass
'                  SplitLimit = -1.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SplitLine(ByVal Line As String, Delimiter As String, ByVal AltDelimiter As String, _
          NoQuotesInFile As Boolean, SplitLimit As Long)
          
          Dim EvenQuotes As Boolean
          Dim i As Long
          Dim RetForEmptyLine(0 To 0) As String
          Dim NoQuotesInLine As Boolean
          
1         On Error GoTo ErrHandler

2         If NoQuotesInFile Then
3             NoQuotesInLine = True
4         ElseIf InStr(Line, DQ) = 0 Then
5             NoQuotesInLine = True
6         End If

7         If NoQuotesInLine Then
8             If Len(Line) = 0 Then
                  'VBA.Split has undesirable behaviour when passed the null string.
                  'Return has zero elements (LBound = 0, UBound = -1).
                  'We need return to be a zero-based, one-element array of
                  'strings, whose single element is the empty string.
9                 SplitLine = RetForEmptyLine
10            Else
11                SplitLine = VBA.Split(Line, Delimiter, SplitLimit)
12            End If
13        Else
14            If Len(AltDelimiter) = 0 Then
15                AltDelimiter = CharNotInString(Line)
16            End If
17            EvenQuotes = True
18            For i = 1 To Len(Line)
19                Select Case Mid$(Line, i, 1)
                      Case DQ
20                        EvenQuotes = Not EvenQuotes
21                    Case Delimiter
22                        If EvenQuotes Then
23                            Mid$(Line, i, 1) = AltDelimiter
24                        End If
25                End Select
26            Next i
27            SplitLine = VBA.Split(Line, AltDelimiter, SplitLimit)
28        End If

29        Exit Function
ErrHandler:
30        Throw "#SplitLine (line " & CStr(Erl) + "): " & Err.Description & "!"
31    End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SplitContents
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Split the contents of a text file into an array, one element per line but only splitting on EOL characters
'              that appear after an even number of double quote characters.
' Parameters :
'  Contents         : The contents of a file as a string.
'  EOL              : The end of line character(s).
'  CharNotInContents: Passed by reference. If contents contains double quotes then this is set to a character known not
'                     to be in Contents - useful for subsequent calls to SplitLine.
'  NoQuotesInFile   : Passed by reference and set to indicate if Contents contains no double quotes.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SplitContents(ByVal Contents As String, EOL As String, ByRef CharNotInContents As String, _
          ByRef NoQuotesInFile As Boolean) As String()
          
          Dim AltEOL As String
          Dim EvenQuotes As Boolean
          Dim i As Long
          Dim Result() As String
          Dim RetForNullString(0 To 0) As String
          Dim TrailingEOL As Boolean
          
1         On Error GoTo ErrHandler

2         TrailingEOL = Right$(Contents, Len(EOL)) = EOL

3         If Len(Contents) = 0 Then
              'VBA.Split has strange behaviour when passed an empty string.
              'Return has zero elements (LBound = 0, UBound = -1).
              'We need return to be a zero-based, one-element array of
              'strings, whose single element is the empty string.
4             Result = RetForNullString
5         ElseIf InStr(Contents, DQ) = 0 Then
6             Result = VBA.Split(Contents, EOL)
7             NoQuotesInFile = True
8         ElseIf EOL = vbLf Or EOL = vbCr Then
9             AltEOL = CharNotInString(Contents)
10            CharNotInContents = AltEOL
11            EvenQuotes = True
12            For i = 1 To Len(Contents)
13                Select Case Mid$(Contents, i, 1)
                      Case DQ
14                        EvenQuotes = Not EvenQuotes
15                    Case EOL
16                        If EvenQuotes Then
17                            Mid$(Contents, i, 1) = AltEOL
18                        End If
19                End Select
20            Next i
21            Result = VBA.Split(Contents, AltEOL)
22        ElseIf EOL = vbCrLf Then
23            AltEOL = CharNotInString(Contents)
24            CharNotInContents = AltEOL
25            EvenQuotes = True
26            For i = 1 To Len(Contents)
27                Select Case Mid$(Contents, i, 1)
                      Case DQ
28                        EvenQuotes = Not EvenQuotes
29                    Case vbLf
30                        If Mid$(Contents, i - 1, 1) = vbCr Then
31                            If EvenQuotes Then
32                                Mid$(Contents, i, 1) = AltEOL
33                                Mid$(Contents, i - 1, 1) = AltEOL
34                            End If
35                        End If
36                End Select
37            Next i
38            Result = VBA.Split(Contents, AltEOL + AltEOL)
39        Else
40            Throw "Unexpected error, Line ending character must be ascii 10 (for Unix), ascii 13 (Mac) or ascii 10 plus ascii 13 (Windows)"
41        End If

42        If TrailingEOL Then
43            If Len(Result(UBound(Result))) = 0 Then
44                ReDim Preserve Result(LBound(Result) To UBound(Result) - 1)
45            End If
46        End If
        
47        SplitContents = Result

48        Exit Function
ErrHandler:
49        Throw "#SplitContents (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDateFormat
' Author     : Philip Swannell
' Date       : 27-Jul-2021
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

1         On Error GoTo ErrHandler

          'Replace repeated D's with a single D, etc since sParseDateCore only needs _
          to know the order in which the three parts of the date appear.
2         If Len(DateFormat) > 5 Then
3             DateFormat = UCase(DateFormat)
4             ReplaceRepeats DateFormat, "D"
5             ReplaceRepeats DateFormat, "M"
6             ReplaceRepeats DateFormat, "Y"
7         End If

8         If Len(DateFormat) = 0 Then
9             DateOrder = Application.International(xlDateOrder)
10            DateSeparator = Application.International(xlDateSeparator)
11        ElseIf Len(DateFormat) <> 5 Then
12            Throw Err_DateFormat + WindowsDefaultDateFormat
13        ElseIf Mid$(DateFormat, 2, 1) <> Mid$(DateFormat, 4, 1) Then
14            Throw Err_DateFormat + WindowsDefaultDateFormat
15        Else
16            DateSeparator = Mid$(DateFormat, 2, 1)
17            Select Case UCase$(Left$(DateFormat, 1) + Mid$(DateFormat, 3, 1) + Right$(DateFormat, 1))
                  Case "MDY"
18                    DateOrder = 0
19                Case "DMY"
20                    DateOrder = 1
21                Case "YMD"
22                    DateOrder = 2
23                Case Else
24                    Throw Err_DateFormat + WindowsDefaultDateFormat
25            End Select
26        End If

27        Exit Sub
ErrHandler:
28        Throw "#ParseDateFormat (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReplaceRepeats
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Replace repeated instances of a character in a string with a single instance.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ReplaceRepeats(ByRef TheString As String, TheChar As String)
          Dim ChCh As String
1         ChCh = TheChar & TheChar
2         While InStr(TheString, ChCh) > 0
3             TheString = Replace(TheString, ChCh, TheChar)
4         Wend
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WindowsDefaultDateFormat
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Returns a description of the system date format, used only for error message generation.
' -----------------------------------------------------------------------------------------------------------------------
Private Function WindowsDefaultDateFormat() As String
          Dim DS As String
1         On Error GoTo ErrHandler
2         DS = Application.International(xlDateSeparator)
3         Select Case Application.International(xlDateOrder)
              Case 0
4                 WindowsDefaultDateFormat = "M" + DS + "D" + DS + "Y"
5             Case 1
6                 WindowsDefaultDateFormat = "D" + DS + "M" + DS + "Y"
7             Case 2
8                 WindowsDefaultDateFormat = "Y" + DS + "M" + DS + "D"
9         End Select

10        Exit Function
ErrHandler:
11        WindowsDefaultDateFormat = "Cannot determine!"
End Function

Private Function SafeClng(x As String)
1         On Error GoTo ErrHandler
2         SafeClng = CLng(x)
3         Exit Function
ErrHandler:
4         SafeClng = 0
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CharNotInString
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Return a character not contained in specified string, or throw an error if not possible.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CharNotInString(Str As String) As String
          Dim i As Long
          Dim TheChar As String
          Const Err_NoAltDelimiter = "Cannot parse files which contain all characters from ascii 0 to ascii 255"
          
1         On Error GoTo ErrHandler

          'Character chr(0) is rare in text files so good idea to try that first. _
          See https://stackoverflow.com/questions/30825491/does-0-appear-naturally-in-text-files

2         For i = 0 To 7
3             TheChar = Chr(i)
4             If InStr(Str, TheChar) = 0 Then
5                 CharNotInString = TheChar
6                 Exit Function
7             End If
8         Next i
9         For i = 255 To 8 Step -1
10            TheChar = Chr(i)
11            If InStr(Str, TheChar) = 0 Then
12                CharNotInString = TheChar
13                Exit Function
14            End If
15        Next i

16        Throw Err_NoAltDelimiter

17        Exit Function
ErrHandler:
18        Throw "#CharNotInString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ShowTextFile
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Parse any text file to a 1-column two-dimensional array of strings. No splitting into columns and no
'              casting.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ShowTextFile(FileName, StartRow As Long, NumRows As Long, MixedLineEndings As Boolean, EOL As String, _
    FileIsUnicode As Boolean)

          Dim FSO As Scripting.FileSystemObject
          Dim F As Scripting.File
          Dim T As Scripting.TextStream
          Dim ReadAll As String
          Dim Contents1D() As String
          Dim Contents2D() As String
          Dim i As Long

1         On Error GoTo ErrHandler
2         Set FSO = New FileSystemObject
3         Set F = FSO.GetFile(FileName)

4         Set T = F.OpenAsTextStream(ForReading, IIf(FileIsUnicode, TristateTrue, TristateFalse))
5         For i = 1 To StartRow - 1
6             T.SkipLine
7         Next

8         If NumRows = 0 Then
9             ReadAll = T.ReadAll
10            T.Close: Set T = Nothing: Set F = Nothing: Set FSO = Nothing

11            If MixedLineEndings Then
12                ReadAll = Replace(ReadAll, vbCrLf, vbLf)
13                ReadAll = Replace(ReadAll, vbCr, vbLf)
14                EOL = vbLf
15            End If

              'Text files may or may not be terminated by EOL characters...
16            If Right$(ReadAll, Len(EOL)) = EOL Then
17                ReadAll = Left$(ReadAll, Len(ReadAll) - Len(EOL))
18            End If

19            If Len(ReadAll) = 0 Then
20                ReDim Contents1D(0 To 0)
21            Else
22                Contents1D = VBA.Split(ReadAll, EOL)
23            End If
24            ReDim Contents2D(1 To UBound(Contents1D) - LBound(Contents1D) + 1, 1 To 1)
25            For i = LBound(Contents1D) To UBound(Contents1D)
26                Contents2D(i + 1, 1) = Contents1D(i)
27            Next i
28        Else
29            ReDim Contents2D(1 To NumRows, 1 To 1)

30            For i = 1 To NumRows
31                Contents2D(i, 1) = T.ReadLine
32            Next i

33            T.Close: Set T = Nothing: Set F = Nothing: Set FSO = Nothing
34        End If

35        ShowTextFile = Contents2D

36        Exit Function
ErrHandler:
37        Throw "#ShowTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToVariant
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Convert a string to a variable of another type, or return the string unchanged if conversion not possible.
'              Always unquotes quoted strings.
' Parameters :
'  strIn                    : The input string.
'  RemoveQuote              : If TRUE then is strIn both starts and ends with double quotes then the return is strIn less
'                             those two characters and with adjacent pairs of double quotes replaced by single double quotes.
'Numbers
'  ShowNumbersAsNumbers     : Boolean - should conversion to Double be attempted?
'  SepsStandard             : Are the decimal separator and Thousands separator the same as the system defaults? If true
'                             then the next three arguments are ignored.
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
'  ShowErrorsAsErrors       : Should strings that match how errors are represented in Excel worksheets be converted to those errors?
'Missings
'  ShowMissingAsNullString  : Should Missings (consecutive delimiters in the file) be represented as zero-length strings.
'  ShowMissingsAs           : Ignored if ShowMissingAsNullString is true, otherwise a variant into which missings are converted.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CastToVariant(strIn As String, RemoveQuotes As Boolean, ShowNumbersAsNumbers As Boolean, SepsStandard As Boolean, _
          DecimalSeparator As String, SysDecimalSeparator As String, _
          ShowDatesAsDates As Boolean, DateOrder As Long, DateSeparator As String, SysDateOrder As Long, _
          SysDateSeparator As String, ShowLogicalsAsLogicals As Boolean, _
          ShowErrorsAsErrors As Boolean, ShowMissingAsNullString As Boolean, ShowMissingsAs As Variant)

          Dim Converted As Boolean
          Dim dblResult As Double
          Dim dtResult As Date
          Dim bResult As Boolean
          Dim eResult As Variant

1         If RemoveQuotes Then
2             StripQuotes strIn, Converted
3             If Converted Then
4                 CastToVariant = strIn
5                 Exit Function
6             End If
7         End If

8         If ShowNumbersAsNumbers Then
9             CastToDouble strIn, dblResult, SepsStandard, DecimalSeparator, SysDecimalSeparator, Converted
10            If Converted Then
11                CastToVariant = dblResult
12                Exit Function
13            End If
14        End If

15        If ShowDatesAsDates Then
16            CastToDate strIn, dtResult, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
17            If Converted Then
18                CastToVariant = dtResult
19                Exit Function
20            End If
21        End If

22        If ShowLogicalsAsLogicals Then
23            CastToBool strIn, bResult, Converted
24            If Converted Then
25                CastToVariant = bResult
26                Exit Function
27            End If
28        End If

29        If ShowErrorsAsErrors Then
30            CastToError strIn, eResult, Converted
31            If Converted Then
32                CastToVariant = eResult
33                Exit Function
34            End If
35        End If

36        If Not ShowMissingAsNullString Then
37            If Len(strIn) = 0 Then
38                CastToVariant = ShowMissingsAs
39                Exit Function
40            End If
41        End If

42        CastToVariant = strIn
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StripQuotes
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Undo how Strings are encoded when written to CSV files.
' Parameters :
'  Str         : String to be converted
'  Converted   : Boolean flag set to True if conversion happens.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub StripQuotes(ByRef Str As String, ByRef Converted As Boolean)
1         If Left$(Str, 1) = DQ Then
2             If Right$(Str, 1) = DQ Then
3                 Str = Mid$(Str, 2, Len(Str) - 2)
4                 Str = Replace(Str, DQ2, DQ)
5                 Converted = True
10            End If
11        End If
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Casts string to double where string has specified decimals separator.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDouble(strIn As String, ByRef dblOut As Double, SepsStandard As Boolean, DecimalSeparator As String, _
           SysDecimalSeparator As String, ByRef Converted As Boolean)
          
1         On Error GoTo ErrHandler
2         If SepsStandard Then
3             dblOut = CDbl(strIn)
4         Else
5             dblOut = CDbl(Replace(strIn, DecimalSeparator, SysDecimalSeparator))
6         End If
7         Converted = True
ErrHandler:
          'Do nothing - strIn was not a string representing a number.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDate
' Author     : Philip Swannell
' Date       : 27-Jul-2021
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
          
1         On Error GoTo ErrHandler
2         pos1 = InStr(strIn, DateSeparator)
3         If pos1 = 0 Then Exit Sub
4         pos2 = InStr(pos1 + 1, strIn, DateSeparator)
5         If pos2 = 0 Then Exit Sub

6         If DateOrder = 0 Then
7             m = Left$(strIn, pos1 - 1)
8             D = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
9             y = Mid$(strIn, pos2 + 1)
10        ElseIf DateOrder = 1 Then
11            D = Left$(strIn, pos1 - 1)
12            m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
13            y = Mid$(strIn, pos2 + 1)
14        ElseIf DateOrder = 2 Then
15            y = Left$(strIn, pos1 - 1)
16            m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
17            D = Mid$(strIn, pos2 + 1)
18        Else
19            Throw "DateOrder must be 0, 1, or 2"
20        End If
21        If SysDateOrder = 0 Then
22            dtOut = CDate(m + SysDateSeparator + D + SysDateSeparator + y)
24            Converted = True
25        ElseIf SysDateOrder = 1 Then
26            dtOut = CDate(D + SysDateSeparator + m + SysDateSeparator + y)
28            Converted = True
29        ElseIf SysDateOrder = 2 Then
30            dtOut = CDate(y + SysDateSeparator + m + SysDateSeparator + D)
32            Converted = True
33        End If

34        Exit Sub
ErrHandler:
          'Do nothing - was not a string representing a date with the specified date order and date separator.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToBool
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Convert string to Boolean, case insensitive.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CastToBool(strIn As String, ByRef bOut As Boolean, ByRef Converted)
          Dim l As Long
1         If VarType(strIn) = vbString Then
2             l = Len(strIn)
3             If l = 4 Then
4                 If StrComp(strIn, "true", vbTextCompare) = 0 Then
5                     bOut = True
6                     Converted = True
7                 End If
8             ElseIf l = 5 Then
9                 If StrComp(strIn, "false", vbTextCompare) = 0 Then
10                    bOut = False
11                    Converted = True
12                End If
13            End If
14        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToError
' Author     : Philip Swannell
' Date       : 27-Jul-2021
' Purpose    : Convert the string representation of Excel errors back to Excel errors.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToError(strIn As String, ByRef eOut As Variant, ByRef Converted As Boolean)
1         On Error GoTo ErrHandler
2         If Left(strIn, 1) = "#" Then
3             Converted = True
4             Select Case strIn 'Editing this function? Then its inverse function Encode!!!!
                  Case "#DIV/0!"
5                     eOut = CVErr(xlErrDiv0)
6                 Case "#NAME?"
7                     eOut = CVErr(xlErrName)
8                 Case "#REF!"
9                     eOut = CVErr(xlErrRef)
10                Case "#NUM!"
11                    eOut = CVErr(xlErrNum)
12                Case "#NULL!"
13                    eOut = CVErr(xlErrNull)
14                Case "#N/A"
15                    eOut = CVErr(xlErrNA)
16                Case "#VALUE!"
17                    eOut = CVErr(xlErrValue)
18                Case "#SPILL!"
19                    eOut = CVErr(2045)    'CVErr(xlErrNoSpill)'These constants introduced in Excel 2016
20                Case "#BLOCKED!"
21                    eOut = CVErr(2047)    'CVErr(xlErrBlocked)
22                Case "#CONNECT!"
23                    eOut = CVErr(2046)    'CVErr(xlErrConnect)
24                Case "#UNKNOWN!"
25                    eOut = CVErr(2048)    'CVErr(xlErrUnknown)
26                Case "#GETTING_DATA!"
27                    eOut = CVErr(2043)    'CVErr(xlErrGettingData)
28                Case "#FIELD!"
29                    eOut = CVErr(2049)    'CVErr(xlErrField)
30                Case "#CALC!"
31                    eOut = CVErr(2050)    'CVErr(xlErrField)
32                Case Else
33                    Converted = False
34            End Select
35        End If

36        Exit Sub
ErrHandler:
39        Throw "#CastToError (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsUnicodeFile
' Author     : Philip Swannell
' Date       : 28-Jul-2021
' Purpose    : Tests if a file is unicode by reading the byte-order-mark. Return is True, False or an error string, so
'              calls should usually be wrapped in ThrowIfError. Adapted from
'              https://stackoverflow.com/questions/36188224/vba-test-encoding-of-a-text-file
' -----------------------------------------------------------------------------------------------------------------------
Public Function IsUnicodeFile(FilePath As String)
          Static FSO As Scripting.FileSystemObject
          Dim T As Scripting.TextStream

          Dim intAsc1Chr As Long
          Dim intAsc2Chr As Long

1         On Error GoTo ErrHandler
2         If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
3         If (FSO.FileExists(FilePath) = False) Then
4             IsUnicodeFile = "#File not found!"
5             Exit Function
6         End If

          ' 1=Read-only, False=do not create if not exist, -1=Unicode 0=ASCII
7         Set T = FSO.OpenTextFile(FilePath, 1, False, 0)
8         If T.AtEndOfStream Then
9             T.Close: Set T = Nothing
10            IsUnicodeFile = False
11            Exit Function
12        End If
13        intAsc1Chr = Asc(T.Read(1))
14        If T.AtEndOfStream Then
15            T.Close: Set T = Nothing
16            IsUnicodeFile = False
17            Exit Function
18        End If
19        intAsc2Chr = Asc(T.Read(1))
20        T.Close
21        If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
22            IsUnicodeFile = True
23        Else
24            IsUnicodeFile = False
25        End If

26        Exit Function
ErrHandler:
27        IsUnicodeFile = "#IsUnicodeFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function OStoEOL(OS As String, ArgName As String) As String

          Const Err_Invalid = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associented end of line characters."

1         Select Case LCase(OS)
              Case "windows", vbCrLf
2                 OStoEOL = vbCrLf
3             Case "unix", vbLf
4                 OStoEOL = vbLf
5             Case "mac", vbCr
6                 OStoEOL = vbCr
7             Case Else
8                 Throw ArgName + Err_Invalid
9         End Select
End Function




'---------------------------------------------------------------------------------------------------------
' Procedure : CSVWrite
' Author    : Philip Swannell
' Date      : 31-Jul-2021
' Purpose   : Creates a csv file on disk containing the data in the array Data. Any existing file of the
'             same name is overwritten. If successful, the function returns the name of the
'             file written, otherwise an error string.
' Arguments
' FileName  : The full name of the file, including the path.
' Data      : An array of arbitrary data. Elements may be strings, numbers, dates, Booleans, empty,
'             Excel errors or null values.
' QuoteAllStrings: If TRUE (the default) then ALL strings are quoted before being written to file. Otherwise
'             (FALSE) only strings containing characters comma, line feed, carriage return
'             or double quote are quoted. Double quotes are always escaped by a second
'             double quote.
' DateFormat: A format string, such as 'yyyy-mm-dd' that determine how dates (e.g. cells formatted as
'             dates) appear in the file.
' DateTimeFormat: A format string, such as 'yyyy-mm-dd hh:mm:ss' that determine how elements of dates with
'             time appear in the file. Currently the companion function CVSRead is not
'             capable of interpreting fields written in DateTime format.
' Delimiter : If TRUE then the file written is encoded as unicode. Defaults to FALSE for an ascii file.
' Unicode   : If FALSE (the default) the file written will be ascii. If TRUE the file written will be
'             Unicode.
' EOL       : Enter the required line ending character as "Windows" (or ascii 13 plus ascii 10), or
'             "Unix" (or ascii 10) or "Mac" (or ascii 13). If omitted defaults to
'             "Windows".
' Ragged    : This argument is for development purposes only, it will soon be removed.
'
' Notes     : See also CSVRead which is the inverse of this function.
'
'             For definition of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
'---------------------------------------------------------------------------------------------------------

Function CSVWrite(FileName As String, ByVal Data As Variant, Optional QuoteAllStrings As Boolean = True, _
          Optional DateFormat As String = "yyyy-mm-dd", Optional DateTimeFormat As String = "yyyy-mm-dd hh:mm:ss", _
          Optional Delimiter As String = ",", Optional Unicode As Boolean, Optional ByVal EOL As String = vbCrLf, Optional Ragged As Boolean = False)
Attribute CSVWrite.VB_Description = "Creates a csv file on disk containing the data in the array Data. Any existing file of the same name is overwritten. If successful, the function returns the name of the file written, otherwise an error string. "
Attribute CSVWrite.VB_ProcData.VB_Invoke_Func = " \n14"

          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim j As Long
          Dim k As Long
          
          Dim OneLine() As String
          Dim OneLineJoined As String
          Dim T As TextStream
          Dim EOLIsWindows As Boolean
          
          Const Err_EOL = "EOL must be one of ""Windows"", ""Unix"" or ""Mac"" or the associated end of line characters"
          Const Err_Delimiter = "Delimiter must be one character, and cannot be double quote or line feed characters"

1         On Error GoTo ErrHandler

2         EOL = OStoEOL(EOL, "EOL")

3         If Len(Delimiter) <> 1 Or Delimiter = """" Or Delimiter = vbCr Or Delimiter = vbLf Then
4             Throw Err_Delimiter
5         End If

6         If TypeName(Data) = "Range" Then
              'Preserve elements of type Date by using .Value, not .Value2
7             Data = Data.Value
8         End If
9         Force2DArray Data 'Coerce 0-dim & 1-dim to 2-dims.

10        Set FSO = New FileSystemObject
11        Set T = FSO.CreateTextFile(FileName, True, Unicode)

12        ReDim OneLine(LBound(Data, 2) To UBound(Data, 2))

13        For i = LBound(Data) To UBound(Data)
14            For j = LBound(Data, 2) To UBound(Data, 2)
15                OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat)
16            Next j
17            OneLineJoined = VBA.Join(OneLine, Delimiter)

              'If writing in "Ragged" style, remove terminating delimiters
18            If Ragged Then
19                For k = Len(OneLineJoined) To 1 Step -1
20                    If Mid(OneLineJoined, k, 1) <> Delimiter Then Exit For
21                Next k
22                If k < Len(OneLineJoined) Then
23                    OneLineJoined = Left(OneLineJoined, k)
24                End If
25            End If
              
26            If EOLIsWindows Then
27                T.WriteLine OneLineJoined
28            Else
29                T.Write OneLineJoined
30                T.Write EOL
31            End If
32        Next i

          'Quote from https://tools.ietf.org/html/rfc4180#section-2
          'The last record in the file may or may not have an ending line break. _
           We follow Excel (File save as CSV) and *do* put a line break after the last line.

33        T.Close: Set T = Nothing: Set FSO = Nothing
34        CSVWrite = FileName
35        Exit Function
ErrHandler:
36        CSVWrite = "#CSVWrite (line " & CStr(Erl) + "): " & Err.Description & "!"
37        If Not T Is Nothing Then Set T = Nothing: Set FSO = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Encode
' Author     : Philip Swannell
' Date       : 23-Jul-2021
' Purpose    : Encode arbitrary value as a string, sub-routine of CSVWrite.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Encode(x As Variant, QuoteAllStrings As Boolean, DateFormat As String, DateTimeFormat As String) As String
          Const DQ = """"
          Const DQ2 = """"""

1         On Error GoTo ErrHandler
2         Select Case VarType(x)

              Case vbString
3                 If InStr(x, DQ) > 0 Then
4                     Encode = DQ + Replace(x, DQ, DQ2) + DQ
5                 ElseIf QuoteAllStrings Then
6                     Encode = DQ + x + DQ
7                 ElseIf InStr(x, vbCr) > 0 Then
8                     Encode = DQ + x + DQ
9                 ElseIf InStr(x, vbLf) > 0 Then
10                    Encode = DQ + x + DQ
11                ElseIf InStr(x, ",") > 0 Then
12                    Encode = DQ + x + DQ
13                Else
14                    Encode = x
15                End If
16            Case vbBoolean, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbLongLong, vbEmpty
17                Encode = CStr(x)
18            Case vbDate
19                If CLng(x) = CDbl(x) Then
20                    Encode = Format$(x, DateFormat)
21                Else
22                    Encode = Format$(x, DateTimeFormat)
23                End If
24            Case vbNull
25                Encode = "NULL"
26            Case vbError
27                Select Case CStr(x) 'Editing this case statement? Edit also its inverse - CastToError
                      Case "Error 2000"
28                        Encode = "#NULL!"
29                    Case "Error 2007"
30                        Encode = "#DIV/0!"
31                    Case "Error 2015"
32                        Encode = "#VALUE!"
33                    Case "Error 2023"
34                        Encode = "#REF!"
35                    Case "Error 2029"
36                        Encode = "#NAME?"
37                    Case "Error 2036"
38                        Encode = "#NUM!"
39                    Case "Error 2042"
40                        Encode = "#N/A"
41                    Case "Error 2043"
42                        Encode = "#GETTING_DATA!"
43                    Case "Error 2045"
44                        Encode = "#SPILL!"
45                    Case "Error 2046"
46                        Encode = "#CONNECT!"
47                    Case "Error 2047"
48                        Encode = "#BLOCKED!"
49                    Case "Error 2048"
50                        Encode = "#UNKNOWN!"
51                    Case "Error 2049"
52                        Encode = "#FIELD!"
53                    Case "Error 2050"
54                        Encode = "#CALC!"
55                    Case Else
56                        Encode = CStr(x)        'should never hit this line...
57                End Select
58            Case Else
59                Throw "Cannot convert variant of type " + TypeName(x) + " to String"
60        End Select
61        Exit Function
ErrHandler:
62        Throw "#Encode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FunctionWizardActive
' Author     : Philip Swannell
' Date       : 13-Jul-2021
' Purpose    : Test if the Function wizard is active to allow early exit in slow functions.
' -----------------------------------------------------------------------------------------------------------------------
Function FunctionWizardActive() As Boolean
          
1         On Error GoTo ErrHandler
2         If TypeName(Application.Caller) = "Range" Then
3             If Not Application.CommandBars("Standard").Controls(1).Enabled Then
4                 FunctionWizardActive = True
5             End If
6         End If

7         Exit Function
ErrHandler:
8         Throw "#FunctionWizardActive (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Force2DArray
' Author    : Philip Swannell
' Date      : 18-Jun-2013
' Purpose   : In-place amendment of singletons and one-dimensional arrays to two dimensions.
'             singletons and 1-d arrays are returned as 2-d 1-based arrays. Leaves two
'             two dimensional arrays untouched (i.e. a zero-based 2-d array will be left as zero-based).
'             See also Force2DArrayR that also handles Range objects.
'---------------------------------------------------------------------------------------
Sub Force2DArray(ByRef TheArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
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
22        Throw "#Force2DArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : NumDimensions
' Author    : Philip Swannell
' Date      : 16-Jun-2013
' Purpose   : Returns the number of dimensions in an array variable, or 0 if the variable
'             is not an array.
'---------------------------------------------------------------------------------------
Function NumDimensions(x As Variant) As Long
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

'---------------------------------------------------------------------------------------
' Procedure : ThrowIfError
' Author    : Philip Swannell
' Date      : 24-Jun-2013
' Purpose   : In the event of an error, methods intended to be callable from spreadsheets
'             return an error string (starts with "#", ends with "!"). ThrowIfError allows such
'             methods to be used from VBA code while keeping error handling robust
'             MyVariable = ThrowIfError(MyFunctionThatReturnsAStringIfAnErrorHappens(...))
'---------------------------------------------------------------------------------------
Function ThrowIfError(Data As Variant)
1         ThrowIfError = Data
2         If VarType(Data) = vbString Then
3             If Left$(Data, 1) = "#" Then
4                 If Right$(Data, 1) = "!" Then
5                     Throw CStr(Data)
6                 End If
7             End If
8         End If
End Function

