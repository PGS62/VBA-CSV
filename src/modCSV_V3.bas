Attribute VB_Name = "modCSV_V3"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead_V3
' Purpose    : Register the function CSVRead_V3 with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Sub RegisterCSVRead_V3()
          Const FnDesc = "Returns the contents of a comma-separated file on disk as an array."
          Dim ArgDescs() As String
1         ReDim ArgDescs(1 To 10)
2         ArgDescs(1) = "The full name of the file, including the path."
3         ArgDescs(2) = "TRUE to convert Numbers, Dates, Booleans and Excel Errors into their typed values, or FALSE to leave as strings. For more control enter a string containing the letters N, D, B, E eg ""NB"" to convert just numbers and Booleans, not dates or errors."
4         ArgDescs(3) = "Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe found outside quoted regions. Enter FALSE to to see the file's raw contents as would be displayed in a text editor. Delimiter may have more than one character."
5         ArgDescs(4) = "The format of dates in the file such as D-M-Y, M-D-Y or Y/M/D. If omitted a value is read from Windows regional settings. Repeated D's (or M's or Y's) are equivalent to single instances, so that d-m-y and DD-MMM-YYYY are equivalent."
6         ArgDescs(5) = "The row in the file at which reading starts. Optional and defaults to 1 to read from the first row."
7         ArgDescs(6) = "The column in the file at which reading starts. Optional and defaults to 1 to read from the first column."
8         ArgDescs(7) = "The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the end of the file are read."
9         ArgDescs(8) = "The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol are read."
10        ArgDescs(9) = "Enter TRUE if the file is unicode, FALSE if the file is ascii. Omit to infer from the file's contents."
11        ArgDescs(10) = "The character that represents a decimal point. If omitted, then the value from Windows regional settings is used."
12        Application.MacroOptions "CSVRead_V3", FnDesc, , , , , , , , , ArgDescs
End Sub

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVRead_V3
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path.
' ConvertTypes: TRUE to convert Numbers, Dates, Booleans and Excel Errors into their typed values, or
'             FALSE to leave as strings. For more control enter a string containing the
'             letters N, D, B, E eg "NB" to convert just numbers and Booleans, not dates or
'             errors.
' Delimiter : Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe
'             found outside quoted regions. Enter FALSE to to see the file's raw contents
'             as would be displayed in a text editor. Delimiter may have more than one
'             character.
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
'             "pads" with Empty values.
'             * Fields which start with a double quote but do not end with a double quote
'             are handled by being returned unchanged. Necessarily such fields have an even
'             number of double quotes, or otherwise the field will be treated as the last
'             field in the file.
'             * The standard states that csv files should have Windows-style line endings,
'             but the function supports files with Windows, Unix and (old) Mac line
'             endings. Files may also have mixed line endings.
'---------------------------------------------------------------------------------------------------------
Function CSVRead_V3(FileName As String, Optional ConvertTypes As Variant = False, Optional ByVal Delimiter As Variant, _
          Optional DateFormat As String, Optional ByVal SkipToRow As Long = 1, Optional ByVal SkipToCol As Long = 1, _
          Optional ByVal NumRows As Long = 0, Optional ByVal NumCols As Long = 0, _
          Optional ByVal Unicode As Variant, Optional DecimalSeparator As String = vbNullString)

          Const DQ = """"
          Const DQ2 = """"""
          Const Err_Delimiter = "Delimiter character must be passed as a string, FALSE for no delimiter. Omit to guess from file contents"
          Const Err_FileEmpty = "File is empty"
          Const Err_FileIsUniCode = "Unicode must be passed as TRUE or FALSE. Omit to infer from file contents"
          Const Err_InFuncWiz = "#Disabled in Function Dialog!"
          Const Err_NumCols = "NumCols must be positive to read a given number of columns, or zero or omitted to read all columns from SkipToCol to the maximum column encountered."
          Const Err_NumRows = "NumRows must be positive to read a given number of rows, or zero or omitted to read all rows from SkipToRow to the end of the file."
          Const Err_Seps = "DecimalSeparator must be different from Delimiter"
          Const Err_SkipToCol = "SkipToCol must be at least 1."
          Const Err_SkipToRow = "SkipToRow must be at least 1."
          Dim AnyConversion As Boolean
          Dim ColIndexes() As Long
          Dim CSVContents As String
          Dim DateOrder As Long
          Dim DateSeparator As String
          Dim F As Scripting.File
          Dim FSO As New Scripting.FileSystemObject
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim Lengths() As Long
          Dim NotDelimited As Boolean
          Dim NumColsFound As Long
          Dim NumColsInReturn As Long
          Dim NumFields As Long
          Dim NumRowsFound As Long
          Dim NumRowsInReturn As Long
          Dim QuoteCounts() As Long
          Dim RemoveQuotes As Boolean
          Dim ReturnArray() As Variant
          Dim RowIndexes() As Long
          Dim SepsStandard As Boolean
          Dim ShowDatesAsDates As Boolean
          Dim ShowErrorsAsErrors As Boolean
          Dim ShowLogicalsAsLogicals As Boolean
          Dim ShowMissingAsNullString As Boolean
          Dim ShowNumbersAsNumbers As Boolean
          Dim Starts() As Long
          Dim strDelimiter As String
          Dim SysDateOrder As Long
          Dim SysDateSeparator As String
          Dim SysDecimalSeparator As String
          Dim T As Scripting.TextStream
          Dim ThisField As String
          
1         On Error GoTo ErrHandler

2         If FunctionWizardActive() Then
3             CSVRead_V3 = Err_InFuncWiz
4             Exit Function
5         End If

          'Parse and validate inputs...
6         If IsEmpty(Unicode) Or IsMissing(Unicode) Then
7             Unicode = IsUnicodeFile(FileName)
8         ElseIf VarType(Unicode) <> vbBoolean Then
9             Throw Err_FileIsUniCode
10        End If

11        If VarType(Delimiter) = vbBoolean Then
12            If Not Delimiter Then
13                NotDelimited = True
14            Else
15                Throw Err_Delimiter
16            End If
17        ElseIf VarType(Delimiter) = vbString Then
18            strDelimiter = Delimiter
19        ElseIf IsEmpty(Delimiter) Or IsMissing(Delimiter) Then
20            strDelimiter = InferDelimiter(FileName, CBool(Unicode))
21        Else
22            Throw Err_Delimiter
23        End If

24        ParseConvertTypes ConvertTypes, ShowNumbersAsNumbers, _
              ShowDatesAsDates, ShowLogicalsAsLogicals, ShowErrorsAsErrors, RemoveQuotes

25        If ShowNumbersAsNumbers Then
26            If ((DecimalSeparator = Application.DecimalSeparator) Or DecimalSeparator = vbNullString) Then
27                SepsStandard = True
28            ElseIf DecimalSeparator = strDelimiter Then
29                Throw Err_Seps
30            End If
31        End If

32        If ShowDatesAsDates Then
33            ParseDateFormat DateFormat, DateOrder, DateSeparator
34            SysDateOrder = Application.International(xlDateOrder)
35            SysDateSeparator = Application.International(xlDateSeparator)
36        End If

37        If SkipToRow < 1 Then Throw Err_SkipToRow
38        If SkipToCol < 1 Then Throw Err_SkipToCol
39        If NumRows < 0 Then Throw Err_NumRows
40        If NumCols < 0 Then Throw Err_NumCols
          'End of input validation
                
41        If NotDelimited Then
42            CSVRead_V3 = ShowTextFile(FileName, SkipToRow, NumRows, CBool(Unicode))
43            Exit Function
44        End If
                
45        Set F = FSO.GetFile(FileName)
46        Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))
47        If T.AtEndOfStream Then
48            T.Close: Set T = Nothing: Set F = Nothing
49            Throw Err_FileEmpty
50        End If
                
51        If SkipToRow = 1 And NumRows = 0 Then
52            CSVContents = T.ReadAll
53            T.Close: Set T = Nothing: Set F = Nothing
54            Call ParseCSVContents(CSVContents, DQ, strDelimiter, SkipToRow, NumRows, NumRowsFound, NumColsFound, NumFields, Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts)
55        Else
56            CSVContents = ParseCSVContents(T, DQ, strDelimiter, SkipToRow, NumRows, NumRowsFound, NumColsFound, NumFields, Starts, Lengths, RowIndexes, ColIndexes, QuoteCounts)
57            T.Close
58        End If
          
59        AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
              ShowLogicalsAsLogicals Or ShowErrorsAsErrors Or (Not ShowMissingAsNullString)
              
60        If NumCols = 0 Then
61            NumColsInReturn = NumColsFound - SkipToCol + 1
62        Else
63            NumColsInReturn = NumCols
64        End If
65        If NumRows = 0 Then
66            NumRowsInReturn = NumRowsFound
67        Else
68            NumRowsInReturn = NumRows
69        End If
              
70        ReDim ReturnArray(1 To NumRowsInReturn, 1 To NumColsInReturn)
              
71        For k = 1 To NumFields
72            i = RowIndexes(k)
73            j = ColIndexes(k) - SkipToCol + 1
74            If j >= 1 And j <= NumColsInReturn Then
75                If QuoteCounts(k) = 0 Or Not RemoveQuotes Then
76                    ThisField = Mid(CSVContents, Starts(k), Lengths(k))
77                ElseIf Mid(CSVContents, Starts(k), 1) = DQ And Mid(CSVContents, Starts(k) + Lengths(k) - 1, 1) = DQ Then
78                    ThisField = Mid(CSVContents, Starts(k) + 1, Lengths(k) - 2)
79                    If QuoteCounts(k) > 2 Then
80                        ThisField = Replace(ThisField, DQ2, DQ)
81                    End If
82                Else 'Field which does not start with quote but contains quotes, so not RFC4180 compliant. We do not replace DQ2 by DQ in this case.
83                    ThisField = Mid(CSVContents, Starts(k), Lengths(k))
84                End If
              
85                If AnyConversion And QuoteCounts(k) = 0 Then
86                    ReturnArray(i, j) = CastToVariant(ThisField, _
                          ShowNumbersAsNumbers, SepsStandard, DecimalSeparator, SysDecimalSeparator, _
                          ShowDatesAsDates, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, _
                          ShowLogicalsAsLogicals, ShowErrorsAsErrors)
87                Else
88                    ReturnArray(i, j) = ThisField
89                End If
90            End If
91        Next k

92        CSVRead_V3 = ReturnArray

93        Exit Function

ErrHandler:
94        CSVRead_V3 = "#CSVRead_V3 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FunctionWizardActive
' Purpose    : Test if the Function wizard is active to allow early exit in slow functions.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FunctionWizardActive() As Boolean
          
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseConvertTypes
' Purpose    : Parse the input ConvertTypes to set five Boolean flags which are passed by reference
' Parameters :
'  ConvertTypes        :
'  ShowNumbersAsNumbers  : Should fields in the file that look like numbers be returned as Numbers? (Doubles)
'  ShowDatesAsDates      : Should fields in the file that look like dates with the specified DateFormat be returned as Dates?
'  ShowLogicalsAsLogicals: Should fields in the file that are TRUE or FALSE (case insensitive) be returned as Booleans?
'  ShowErrorsAsErrors    : Should fields in the file that look like Excel errors (#N/A #REF! etc) be returned as errors?
'  RemoveQuotes          : Should quoted fields be unquoted?
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseConvertTypes(ByVal ConvertTypes As Variant, ByRef ShowNumbersAsNumbers As Boolean, _
          ByRef ShowDatesAsDates As Boolean, ByRef ShowLogicalsAsLogicals As Boolean, _
          ByRef ShowErrorsAsErrors As Boolean, ByRef RemoveQuotes As Boolean)

          Const Err_ConvertTypes = "ConvertTypes must be TRUE (convert all types), FALSE (no conversion) or a string of letter: " & _
              "'N' to show numbers as numbers, 'D' to show dates as dates, 'L' to show logicals as logicals, `E` to show Excel " & _
              "errors as errors, Q to show quoted fields with their quotes."
          Dim i As Long

1         On Error GoTo ErrHandler
2         If TypeName(ConvertTypes) = "Range" Then
3             ConvertTypes = ConvertTypes.value
4         End If

5         If VarType(ConvertTypes) = vbBoolean Then
6             If ConvertTypes Then
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
19        ElseIf VarType(ConvertTypes) = vbString Then
20            ShowNumbersAsNumbers = False
21            ShowDatesAsDates = False
22            ShowLogicalsAsLogicals = False
23            ShowErrorsAsErrors = False
24            RemoveQuotes = True
25            For i = 1 To Len(ConvertTypes)
26                Select Case UCase(Mid(ConvertTypes, i, 1))
                      Case "N"
27                        ShowNumbersAsNumbers = True
28                    Case "D"
29                        ShowDatesAsDates = True
30                    Case "L", "B" 'Logicals aka Booleans
31                        ShowLogicalsAsLogicals = True
32                    Case "E"
33                        ShowErrorsAsErrors = True
34                    Case "Q"
35                        RemoveQuotes = False
36                    Case Else
37                        Throw "Unrecognised character '" + Mid(ConvertTypes, i, 1) + "' in ConvertTypes."
38                End Select
39            Next i
40        Else
41            Throw Err_ConvertTypes
42        End If

43        Exit Sub
ErrHandler:
44        Throw "#ParseConvertTypes (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Min4
' Purpose    : Returns the minimum of four inputs and an indicator of which of the four was the minimum
' -----------------------------------------------------------------------------------------------------------------------
Private Function Min4(N1 As Long, N2 As Long, N3 As Long, N4 As Long, ByRef Which As Long) As Long

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
27        Throw "#IsUnicodeFile (line " & CStr(Erl) + "): " & Err.Description & "!"
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
                      Case QuoteChar
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

11            ReadAll = Replace(ReadAll, vbCrLf, vbLf)
12            ReadAll = Replace(ReadAll, vbCr, vbLf)

              'Text files may or may not be terminated by EOL characters...
13            If Right$(ReadAll, 1) = vbLf Then
14                ReadAll = Left$(ReadAll, Len(ReadAll) - 1)
15            End If

16            If Len(ReadAll) = 0 Then
17                ReDim Contents1D(0 To 0)
18            Else
19                Contents1D = VBA.Split(ReadAll, vbLf)
20            End If
21            ReDim Contents2D(1 To UBound(Contents1D) - LBound(Contents1D) + 1, 1 To 1)
22            For i = LBound(Contents1D) To UBound(Contents1D)
23                Contents2D(i + 1, 1) = Contents1D(i)
24            Next i
25        Else
26            ReDim Contents2D(1 To NumRows, 1 To 1)

27            For i = 1 To NumRows 'BUG, won't work for Mac files. TODO Fix this?
28                Contents2D(i, 1) = T.ReadLine
29            Next i

30            T.Close: Set T = Nothing: Set F = Nothing: Set FSO = Nothing
31        End If

32        ShowTextFile = Contents2D

33        Exit Function
ErrHandler:
34        Throw "#ShowTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCSVContents
' Purpose    : Parse the contents of a CSV file.
' Parameters :
'  ContentsOrStream: The contents of a CSV file as a string, or else a Scripting.TextStream.
'  QuoteChar       : The quote character, usually ascii 34 ("), which allow fields to contain characters that would otherwise
'                    be significant to parsing, such as delimiters or new line characters.
'  Delimiter       : The string that separates fields within each line. Typically a single character, but needn't be.
'  SkipToRow       : Rows in the file prior to SkipToRow are ignored.
'  NumRows         : The number of rows to parse. 0 for all rows from SkipToRow to the end of the file.
'  NumRowsFound    : Set to the number of rows in the file.
'  NumColsFound    : Set to the number of columns in the file, i.e. the maximum number of fields in any single line.
'  NumFields       : Set to the number of fields in the file.  May be less than NumRowsFound times NumColsFound if not all
'                    lines have the same number of fields.
'  Starts          : Set to an array of size at least NumFields. Element k gives the point in CSVContents at which the kth
'                    field starts.
'  Lengths         : Set to an array of size at least NumFields. Element k gives the length of the kth field.
'  IsLasts         : Set to an array of size at least NumFields. Element k indicates whether the kth field is the last field
'                    in its line.
'  QuoteCounts     : Set to an array of size at least NumFields. Element k gives the number of QuoteChars that appear in the
'                    kth field.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseCSVContents(ContentsOrStream As Variant, QuoteChar As String, Delimiter As String, SkipToRow As Long, NumRows As Long, ByRef NumRowsFound As Long, ByRef NumColsFound As Long, _
    ByRef NumFields As Long, ByRef Starts() As Long, ByRef Lengths() As Long, RowIndexes() As Long, ColIndexes() As Long, QuoteCounts() As Long) As String

          Const Err_ContentsOrStream = "ContentsOrStream must either be a string or a TextStream"
          Dim Buffer As String
          Dim BufferUpdatedTo As Long
          Dim ColNum As Long
          Dim EvenQuotes As Boolean
          Dim HaveReachedSkipToRow As Boolean
          Dim i As Long 'Index to read from Buffer
          Dim j As Long 'Index to write to Starts, Lengths, RowIndexes and ColIndexes
          Dim LDlm As Long
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

1         On Error GoTo ErrHandler
          
2         If VarType(ContentsOrStream) = vbString Then
              'TODO Remove this bodge, useful for testing
3             If LCase(Left(ContentsOrStream, 3)) = "c:\" Then
                  Dim FSO As New FileSystemObject, F As Scripting.File
4                 Set F = FSO.GetFile(ContentsOrStream)
5                 Set T = F.OpenAsTextStream()
6                 Streaming = True
7             Else
8                 Buffer = ContentsOrStream
9                 Streaming = False
10            End If
11        ElseIf TypeName(ContentsOrStream) = "TextStream" Then
12            Set T = ContentsOrStream
13            If NumRows = 0 Then
14                Buffer = T.ReadAll
15                T.Close
16                Streaming = False
17            Else
18                Call GetMoreFromStream(T, Delimiter, QuoteChar, Buffer, BufferUpdatedTo)
19                Streaming = True
20            End If
21        Else
22            Throw Err_ContentsOrStream
23        End If
             
24        If Streaming Then
25            ReDim SearchFor(1 To 4)
26            SearchFor(1) = Delimiter
27            SearchFor(2) = vbLf
28            SearchFor(3) = vbCr
29            SearchFor(4) = QuoteChar
30            ReDim QuoteArray(1 To 1)
31            QuoteArray(1) = QuoteChar
32        End If

33        ReDim Starts(1 To 8)
34        ReDim Lengths(1 To 8)
35        ReDim RowIndexes(1 To 8)
36        ReDim ColIndexes(1 To 8)
37        ReDim QuoteCounts(1 To 8)
          
38        LDlm = Len(Delimiter)
39        OrigLen = Len(Buffer)
40        If Not Streaming Then
              'Ensure Buffer terminates with vbCrLf
41            If Right(Buffer, 1) <> vbCr And Right(Buffer, 1) <> vbLf Then
42                Buffer = Buffer + vbCrLf
43            ElseIf Right(Buffer, 1) = vbCr Then
44                Buffer = Buffer + vbLf
45            End If
46            BufferUpdatedTo = Len(Buffer)
47        End If
          
48        j = 1
49        ColNum = 1: RowNum = 1
50        EvenQuotes = True
51        Starts(1) = 1
52        If SkipToRow = 1 Then HaveReachedSkipToRow = True

53        Do
54            If EvenQuotes Then
55                If Not Streaming Then
56                    If PosDL <= i Then PosDL = InStr(i + 1, Buffer, Delimiter): If PosDL = 0 Then PosDL = BufferUpdatedTo + 1
57                    If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
58                    If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
59                    If PosQC <= i Then PosQC = InStr(i + 1, Buffer, QuoteChar): If PosQC = 0 Then PosQC = BufferUpdatedTo + 1
60                    i = Min4(PosDL, PosLF, PosCR, PosQC, Which)
61                Else
62                    i = SearchInBuffer(SearchFor, i + 1, T, Delimiter, QuoteChar, Which, Buffer, BufferUpdatedTo)
63                End If

64                If i = BufferUpdatedTo + 1 Then
65                    Exit Do
66                End If

67                If j + 1 > UBound(Starts) Then
68                    ReDim Preserve Starts(1 To UBound(Starts) * 2)
69                    ReDim Preserve Lengths(1 To UBound(Lengths) * 2)
70                    ReDim Preserve RowIndexes(1 To UBound(RowIndexes) * 2)
71                    ReDim Preserve ColIndexes(1 To UBound(ColIndexes) * 2)
72                    ReDim Preserve QuoteCounts(1 To UBound(QuoteCounts) * 2)
73                End If

74                Select Case Which
                      Case 1
                          'Found Delimiter
75                        Lengths(j) = i - Starts(j)
76                        Starts(j + 1) = i + LDlm
77                        ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
78                        ColNum = ColNum + 1
79                        QuoteCounts(j) = QuoteCount: QuoteCount = 0
80                        j = j + 1
81                        NumFields = NumFields + 1
82                        i = i + LDlm - 1
83                    Case 2, 3
84                        Lengths(j) = i - Starts(j)
85                        If Which = 2 Then
                              'Unix line ending
86                            Starts(j + 1) = i + 1
87                        ElseIf Mid(Buffer, i + 1, 1) = vbLf Then
                              'Windows line ending. - it is safe to look one character ahead since Buffer terminates with vbCrLf
88                            Starts(j + 1) = i + 2
89                            i = i + 1
90                        Else
                              'Mac line ending (Mac pre OSX)
91                            Starts(j + 1) = i + 1
92                        End If

93                        If ColNum > NumColsFound Then NumColsFound = ColNum
94                        ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
95                        ColNum = 1: RowNum = RowNum + 1
                          
96                        QuoteCounts(j) = QuoteCount: QuoteCount = 0
97                        j = j + 1
98                        NumFields = NumFields + 1
                          
99                        If HaveReachedSkipToRow Then
100                           If RowNum = NumRows + 1 Then
101                               Exit Do
102                           End If
103                       Else
104                           If RowNum = SkipToRow Then
105                               HaveReachedSkipToRow = True
106                               tmp = Starts(j)
107                               ReDim Starts(1 To 8): ReDim Lengths(1 To 8): ReDim RowIndexes(1 To 8): ReDim ColIndexes(1 To 8): ReDim QuoteCounts(1 To 8)
108                               RowNum = 1: j = 1
109                               Starts(1) = tmp
110                           End If
111                       End If
112                   Case 4
                          'Found QuoteChar
113                       EvenQuotes = False
114                       QuoteCount = QuoteCount + 1
115               End Select
116           Else
117               If Not Streaming Then
118                   PosQC = InStr(i + 1, Buffer, QuoteChar)
119               Else
120                   If PosQC <= i Then PosQC = SearchInBuffer(QuoteArray(), i + 1, T, Delimiter, QuoteChar, 0, Buffer, BufferUpdatedTo)
121               End If
                  
122               If PosQC = 0 Then
                      'Malformed Buffer (not RFC4180 compliant). There should always be an even number of double quotes. _
                       If there are an odd number then all text after the last double quote in the file will be (part of) _
                       the last field in the last line.
123                   Lengths(j) = OrigLen - Starts(j) + 1
124                   ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                      
125                   RowNum = RowNum + 1
126                   If ColNum > NumColsFound Then NumColsFound = ColNum
127                   NumFields = NumFields + 1
128                   Exit Do
129               Else
130                   i = PosQC
131                   EvenQuotes = True
132                   QuoteCount = QuoteCount + 1
133               End If
134           End If
135       Loop
          
136       NumRowsFound = RowNum - 1

137       ParseCSVContents = Buffer

138       Exit Function
ErrHandler:
139       Throw "#ParseCSVContents (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SearchInBuffer
' Purpose    : Returns the location in the buffer of the first-encountered string amongst the elements of SearchFor,
'              starting the search at point SearchFrom and finishing the search at point BufferUpdatedTo. If none found in
'              that region returns BufferUpdatedTo + 1. Otherwise returns the location of the first found and sets the
'              by-reference argument Which to indicate which element of SearchFor was the first to be found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SearchInBuffer(SearchFor() As String, StartingAt As Long, T As Scripting.TextStream, Delimiter As String, QuoteChar As String, ByRef Which As Long, ByRef Buffer As String, ByRef BufferUpdatedTo As Long)

          Dim InstrRes As Long
          Dim PrevBufferUpdatedTo As Long

1         On Error GoTo ErrHandler

2         InstrRes = InStrMulti(SearchFor, Buffer, StartingAt, BufferUpdatedTo, Which)
3         If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
4             SearchInBuffer = InstrRes
5             Exit Function
6         ElseIf T.AtEndOfStream Then
7             SearchInBuffer = BufferUpdatedTo + 1
8             Exit Function
9         End If

10        Do
11            PrevBufferUpdatedTo = BufferUpdatedTo
12            GetMoreFromStream T, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
13            InstrRes = InStrMulti(SearchFor, Buffer, PrevBufferUpdatedTo + 1, BufferUpdatedTo, Which)
14            If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
15                SearchInBuffer = InstrRes
16                Exit Function
17            ElseIf T.AtEndOfStream Then
18                SearchInBuffer = BufferUpdatedTo + 1
19                Exit Function
20            End If
21        Loop
22        Exit Function
ErrHandler:
23        Throw "#SearchInBuffer (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InStrMulti
' Purpose    : Returns the first point in SearchWithin at which one of the elements of SearchFor is found, search is
'              restricted to region [StartingAt, EndingAt] and Which is updated with the index into SearchFor of the first string found
' -----------------------------------------------------------------------------------------------------------------------
Function InStrMulti(SearchFor() As String, SearchWithin As String, StartingAt As Long, EndingAt As Long, ByRef Which As Long)

          Dim InstrRes() As Long
          Dim i As Long
          Dim LB As Long, UB As Long
          Const Inf = 2147483647
          Dim Result As Long

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
18        Throw "#InStrMulti (line " & CStr(Erl) + "): " & Err.Description & "!"
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
Private Sub GetMoreFromStream(T As Scripting.TextStream, Delimiter As String, QuoteChar As String, ByRef Buffer As String, ByRef BufferUpdatedTo As Long)
          Const CHUNKSIZE = 5000  ' The number of characters to read from the stream on each call. _
                                    Set to a small number for testing logic and a bigger number for _
                                    performance, but not too high since a common use case is reading _
                                    just the first line of a file. Suggest 5000?
          Dim NewChars
          Dim OKToExit
          Dim i As Long
          Dim NCharsToWriteToBuffer As Long
          Dim ExpandBufferBy As Long
          Dim FirstPass As Boolean

1         On Error GoTo ErrHandler
2         FirstPass = True
3         Do
4             NewChars = T.Read(IIf(FirstPass, CHUNKSIZE, 1))
5             FirstPass = False
6             If T.AtEndOfStream Then
                  'Ensure NewChars terminates with vbCrLf
7                 If Right(NewChars, 1) <> vbCr And Right(NewChars, 1) <> vbLf Then
8                     NewChars = NewChars + vbCrLf
9                 ElseIf Right(NewChars, 1) = vbCr Then
10                    NewChars = NewChars + vbLf
11                End If
12            End If

13            NCharsToWriteToBuffer = Len(NewChars) + Len(Delimiter) + 3

14            If BufferUpdatedTo + NCharsToWriteToBuffer > Len(Buffer) Then
15                ExpandBufferBy = MaxLngs(Len(Buffer), NCharsToWriteToBuffer)
16                Buffer = Buffer & String(ExpandBufferBy, "?")
17            End If
              
18            Mid(Buffer, BufferUpdatedTo + 1, Len(NewChars)) = NewChars
19            BufferUpdatedTo = BufferUpdatedTo + Len(NewChars)

20            OKToExit = True
              'Ensure we don't leave the buffer updated to part way through a two-character end of line marker.
21            If Right(NewChars, 1) = vbCr Then
22                OKToExit = False
23            End If
              'Ensure we don't leave the buffer updated to a point part-way through a multi-character delimiter
24            If Len(Delimiter) > 1 Then
25                For i = 1 To Len(Delimiter) - 1
26                    If Mid$(Buffer, BufferUpdatedTo - i + 1, i) = Left(Delimiter, i) Then
27                        OKToExit = False
28                        Exit For
29                    End If
30                Next i
31                If Mid(Buffer, BufferUpdatedTo - Len(Delimiter) + 1, Len(Delimiter)) = Delimiter Then
32                    OKToExit = True
33                End If
34            End If
35            If OKToExit Then Exit Do
36        Loop

          'Line below arranges that when calling Instr(Buffer,....) we don't pointlessly scan the space characters _
           we can be sure that there is space in the buffer to write the extra characters thanks to
37        Mid(Buffer, BufferUpdatedTo + 1, Len(Delimiter) + 3) = vbCrLf & QuoteChar & Delimiter

38        Exit Sub
ErrHandler:
39        Throw "#GetMoreFromStream (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function MaxLngs(x As Long, y As Long) As Long
1         If x > y Then
2             MaxLngs = x
3         Else
4             MaxLngs = y
5         End If
End Function

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
' -----------------------------------------------------------------------------------------------------------------------
Private Function CastToVariant(strIn As String, ShowNumbersAsNumbers As Boolean, SepsStandard As Boolean, _
    DecimalSeparator As String, SysDecimalSeparator As String, _
    ShowDatesAsDates As Boolean, DateOrder As Long, DateSeparator As String, SysDateOrder As Long, _
    SysDateSeparator As String, ShowLogicalsAsLogicals As Boolean, _
    ShowErrorsAsErrors As Boolean)

          Dim Converted As Boolean
          Dim dblResult As Double
          Dim dtResult As Date
          Dim bResult As Boolean
          Dim eResult As Variant

1         If ShowNumbersAsNumbers Then
2             CastToDouble strIn, dblResult, SepsStandard, DecimalSeparator, SysDecimalSeparator, Converted
3             If Converted Then
4                 CastToVariant = dblResult
5                 Exit Function
6             End If
7         End If

8         If ShowDatesAsDates Then
9             CastToDate strIn, dtResult, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
10            If Converted Then
11                CastToVariant = dtResult
12                Exit Function
13            End If
14        End If

15        If ShowLogicalsAsLogicals Then
16            CastToBool strIn, bResult, Converted
17            If Converted Then
18                CastToVariant = bResult
19                Exit Function
20            End If
21        End If

22        If ShowErrorsAsErrors Then
23            CastToError strIn, eResult, Converted
24            If Converted Then
25                CastToVariant = eResult
26                Exit Function
27            End If
28        End If

29        CastToVariant = strIn
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble
' Purpose    : Casts strIn to double where strIn has specified decimals separator.
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
          
1         On Error GoTo ErrHandler
2         pos1 = InStr(strIn, DateSeparator)
3         If pos1 = 0 Then Exit Sub
4         pos2 = InStr(pos1 + 1, strIn, DateSeparator)
5         If pos2 = 0 Then Exit Sub

6         If DateOrder = 0 Then
7             m = Left$(strIn, pos1 - 1)
8             d = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
9             y = Mid$(strIn, pos2 + 1)
10        ElseIf DateOrder = 1 Then
11            d = Left$(strIn, pos1 - 1)
12            m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
13            y = Mid$(strIn, pos2 + 1)
14        ElseIf DateOrder = 2 Then
15            y = Left$(strIn, pos1 - 1)
16            m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
17            d = Mid$(strIn, pos2 + 1)
18        Else
19            Throw "DateOrder must be 0, 1, or 2"
20        End If
21        If SysDateOrder = 0 Then
22            dtOut = CDate(m + SysDateSeparator + d + SysDateSeparator + y)
23            Converted = True
24        ElseIf SysDateOrder = 1 Then
25            dtOut = CDate(d + SysDateSeparator + m + SysDateSeparator + y)
26            Converted = True
27        ElseIf SysDateOrder = 2 Then
28            dtOut = CDate(y + SysDateSeparator + m + SysDateSeparator + d)
29            Converted = True
30        End If

31        Exit Sub
ErrHandler:
          'Do nothing - was not a string representing a date with the specified date order and date separator.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToBool
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
37        Throw "#CastToError (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function OStoEOL(OS As String, ArgName As String) As String

          Const Err_Invalid = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associented end of line characters."

1         On Error GoTo ErrHandler
2         Select Case LCase(OS)
              Case "windows", vbCrLf
3                 OStoEOL = vbCrLf
4             Case "unix", vbLf
5                 OStoEOL = vbLf
6             Case "mac", vbCr
7                 OStoEOL = vbCr
8             Case Else
9                 Throw ArgName + Err_Invalid
10        End Select

11        Exit Function
ErrHandler:
12        Throw "#OStoEOL (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVWrite
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
' Notes     : See also CSVRead_V1 which is the inverse of this function.
'
'             For definition of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
'---------------------------------------------------------------------------------------------------------

Function CSVWrite(FileName As String, ByVal data As Variant, Optional QuoteAllStrings As Boolean = True, _
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
          Const DQ = """"
          
          'Const Err_Delimiter = "Delimiter must be one character, and cannot be double quote or line feed characters"
          Const Err_Delimiter = "Delimiter must not contain double quote or line feed characters"

1         On Error GoTo ErrHandler

2         EOL = OStoEOL(EOL, "EOL")
3         EOLIsWindows = EOL = vbCrLf

          'If Len(Delimiter) <> 1 Or Delimiter = """" Or Delimiter = vbCr Or Delimiter = vbLf Then
4         If InStr(Delimiter, DQ) > 0 Or InStr(Delimiter, vbLf) > 0 Or InStr(Delimiter, vbCr) > 0 Then
5             Throw Err_Delimiter
6         End If

7         If TypeName(data) = "Range" Then
              'Preserve elements of type Date by using .Value, not .Value2
8             data = data.value
9         End If
10        Force2DArray data 'Coerce 0-dim & 1-dim to 2-dims.

11        Set FSO = New FileSystemObject
12        Set T = FSO.CreateTextFile(FileName, True, Unicode)

13        ReDim OneLine(LBound(data, 2) To UBound(data, 2))

14        For i = LBound(data) To UBound(data)
15            For j = LBound(data, 2) To UBound(data, 2)
16                OneLine(j) = Encode(data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat)
17            Next j
18            OneLineJoined = VBA.Join(OneLine, Delimiter)

              'If writing in "Ragged" style, remove terminating delimiters
19            If Ragged Then
20                For k = Len(OneLineJoined) To 1 Step -1
21                    If Mid(OneLineJoined, k, 1) <> Delimiter Then Exit For
22                Next k
23                If k < Len(OneLineJoined) Then
24                    OneLineJoined = Left(OneLineJoined, k)
25                End If
26            End If
              
27            WriteLineWrap T, OneLineJoined, EOLIsWindows, EOL, Unicode
28        Next i

          'Quote from https://tools.ietf.org/html/rfc4180#section-2 : _
          "The last record in the file may or may not have an ending line break." _
           We follow Excel (File save as CSV) and *do* put a line break after the last line.

29        T.Close: Set T = Nothing: Set FSO = Nothing
30        CSVWrite = FileName
31        Exit Function
ErrHandler:
32        CSVWrite = "#CSVWrite (line " & CStr(Erl) + "): " & Err.Description & "!"
33        If Not T Is Nothing Then Set T = Nothing: Set FSO = Nothing

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WriteLineWrap
' Purpose    : Wrapper to TextStream.Write(Line) to give more informative error message than "invalid procedure call or
'              argument" if the error is caused by attempting to write Unicode characters to ascii file
' -----------------------------------------------------------------------------------------------------------------------
Private Sub WriteLineWrap(T As TextStream, text As String, EOLIsWindows As Boolean, EOL As String, Unicode As Boolean)
          Dim ErrDesc As String
          Dim ErrNum As Long
          Dim i As Long
          Dim ErrLN As String

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
10        ErrDesc = Err.Description
11        ErrLN = CStr(Erl)
12        If Not Unicode Then
13            If ErrNum = 5 Then
14                For i = 1 To Len(text)
15                    If AscW(Mid(text, i, 1)) > 255 Then
16                        ErrDesc = "Data contains unicode characters (first found has code " & CStr(AscW(Mid(text, i, 1))) & ") which cannot be written to an ascii file. Set argument Unicode to True"
17                        Exit For
18                    End If
19                Next i
20            End If
21        End If
22        Throw "#WriteLineWrap (line " & ErrLN + "): " & ErrDesc & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Encode
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

