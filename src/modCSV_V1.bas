Attribute VB_Name = "modCSV_V1"
Option Explicit
Private Const DQ = """"
Private Const DQ2 = """"""
Private Const Err_EmptyFile = "File is empty"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead_V1
' Purpose    : Register the function CSVRead_V1 with the Excel Function Wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Sub RegisterCSVRead_V1()
    Const FnDesc = "Returns the contents of a comma-separated file on disk as an array."
    Dim ArgDescs() As String
    ReDim ArgDescs(1 To 12)
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
    Application.MacroOptions "CSVRead_V1", FnDesc, , , , , , , , , ArgDescs
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVWrite
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

    On Error GoTo ErrHandler
    If VarType(FileIsUnicode) <> vbBoolean Then
        FileIsUnicode = IsUnicodeFile(FileName)
    End If
    If EOL = "" Then
        EOL = InferEOL(FileName, CBool(FileIsUnicode))
    End If

    CSVS.Init FileName, EOL, CBool(FileIsUnicode)
    Set CreateCSVStream = CSVS

    Exit Function
ErrHandler:
    Throw "#CreateCSVStream: " & Err.Description & "!"
End Function

Sub Throw(ByVal ErrorString As String)
    Err.Raise vbObjectError + 1, , ErrorString
End Sub

'---------------------------------------------------------------------------------------------------------
' Procedure : CSVRead_V1
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path.
' ConvertTypes: TRUE to convert Numbers, Dates, Logicals and Excel Errors into their typed values, or
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
Function CSVRead_V1(FileName As String, Optional ConvertTypes As Variant = False, Optional ByVal Delimiter As Variant, _
        Optional DateFormat As String, Optional ByVal StartRow As Long = 1, Optional ByVal StartCol As Long = 1, _
        Optional ByVal NumRows As Long = 0, Optional ByVal NumCols As Long = 0, Optional ByVal LineEndings As Variant, _
        Optional ByVal Unicode As Variant, Optional ByVal ShowMissingsAs As Variant = "", _
        Optional DecimalSeparator As String = vbNullString)
Attribute CSVRead_V1.VB_Description = "Returns the contents of a comma-separated file on disk as an array."
Attribute CSVRead_V1.VB_ProcData.VB_Invoke_Func = " \n14"

    Const Err_Delimiter = "Delimiter character must be passed as a string, FALSE for no delimiter, or else omitted to infer from the file's contents"
    Const Err_FileIsUniCode = "Unicode must be passed as TRUE or FALSE, or omitted to infer from the file's contents"
    Const Err_InFuncWiz = "#Disabled in Function Dialog!"
    Const Err_LineEndings = "LineEndings must be one of ""Windows"", ""Unix"" or ""Mac"", or omitted to infer from the file's contents"
    Const Err_NumCols = "NumCols must be positive to read a given number or columns, or zero or omitted to read all columns from StartCol to the maximum column encountered."
    Const Err_NumRows = "NumRows must be positive to read a given number or rows, or zero or omitted to read all rows from StartRow to the end of the file."
    Const Err_Seps = "DecimalSeparator must be different from Delimiter"
    Const Err_StartCol = "StartCol must be at least 1."
    Const Err_StartRow = "StartRow must be at least 1."
    
    'The lower bounds for the array returned by this function is set by the constant LB rather than _
     by any "Option Base 1" or "Option Base 0" that might appear at the start of the module.
    Const LB = 1
    
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
    Dim Lines() As String
    Dim MixedLineEndings As Boolean
    Dim NotDelimited As Boolean
    Dim NumColsInReturn As Long
    Dim NumInRow As Long
    Dim NumRowsInReturn As Long
    Dim OneRow() As String
    Dim p As Long
    Dim q As Long
    Dim ReadAll As String
    Dim RemoveQuotes As Boolean
    Dim QuotesEncountered As Boolean
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
    Dim TrimLastLine As Boolean
    
    On Error GoTo ErrHandler

    If FunctionWizardActive() Then
        CSVRead_V1 = Err_InFuncWiz
        Exit Function
    End If

    'Parse and validate inputs...
    If IsEmpty(Unicode) Or IsMissing(Unicode) Then
        Unicode = IsUnicodeFile(FileName)
    ElseIf VarType(Unicode) <> vbBoolean Then
        Throw Err_FileIsUniCode
    End If
    
    If TypeName(LineEndings) = "Range" Then LineEndings = LineEndings.value
    If IsMissing(LineEndings) Or IsEmpty(LineEndings) Then
        EOL = InferEOL(FileName, CBool(Unicode))
    ElseIf VarType(LineEndings) = vbString Then
        Select Case LCase(LineEndings)
            Case "windows", vbCrLf
                EOL = vbCrLf
            Case "unix", vbLf
                EOL = vbLf
            Case "mac", vbCr
                EOL = vbCr
            Case "mixed"
                MixedLineEndings = True
            Case Else
                Throw Err_LineEndings
        End Select
    ElseIf VarType(LineEndings) = vbBoolean Then
        'For backward compatibility - LineEndings was f.k.a. FileIsUnix
        If LineEndings Then
            EOL = vbLf
        Else
            EOL = vbCrLf
        End If
    Else
        Throw Err_LineEndings
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

    ParseConvertTypes ConvertTypes, ShowNumbersAsNumbers, _
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

    If StartRow < 1 Then Throw Err_StartRow
    If StartCol < 1 Then Throw Err_StartCol
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
        CSVRead_V1 = ShowTextFile(FileName, StartRow, NumRows, MixedLineEndings, EOL, CBool(Unicode))
        Exit Function
    End If
          
    'In this case (reading the entire file) performance is better if we don't use _
     clsCSVStream but instead use method SplitContents on the entire file contents.
    If StartRow = 1 And StartCol = 1 And NumRows = 0 And NumCols = 0 Then
        Set F = FSO.GetFile(FileName)
        Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))
        If T.AtEndOfStream Then
            T.Close: Set T = Nothing: Set F = Nothing
            Throw Err_EmptyFile
        End If

        ReadAll = T.ReadAll
        T.Close: Set T = Nothing: Set F = Nothing
        TrimLastLine = Right(ReadAll, Len(EOL)) = EOL
        
        Dim UnusedChar As String
        If Not Unicode Then
            UnusedChar = ChrW(9986)
        Else
            UnusedChar = CharNotInString(ReadAll)
        End If
        
        AltDelimiter = String(Len(EOL), UnusedChar)
        
        Lines = SplitNew(ReadAll, EOL, """", , AltDelimiter, QuotesEncountered)
        If Not QuotesEncountered Then
            RemoveQuotes = False
        End If
        If TrimLastLine Then
            If Len(Lines(UBound(Lines))) = 0 Then 'This will only be the case if the trailing end of line appeared after an even number of quotes
                ReDim Preserve Lines(LBound(Lines) To UBound(Lines) - 1)
            End If
        End If

    Else
        Set CSVS = CreateCSVStream(FileName, EOL, Unicode)
        For i = 1 To StartRow - 1
            CSVS.ReadLine
        Next i
        CSVS.StartRecording
        If NumRows > 0 Then
            For i = 1 To NumRows
                CSVS.ReadLine
            Next
        Else
            While Not CSVS.AtEndOfStream
                CSVS.ReadLine
            Wend
        End If

        Lines = CSVS.ReportAllLinesRead()
        If Not CSVS.QuotesEncountered Then
            RemoveQuotes = False
        End If
        Set CSVS = Nothing
    End If
    NumRowsInReturn = UBound(Lines) - LBound(Lines) + 1
    
    If NumCols = 0 Then
        NumColsInReturn = 1
        SplitLimit = -1
    Else
        NumColsInReturn = NumCols
        SplitLimit = StartCol - 1 + NumCols + 1
    End If

    AnyConversion = RemoveQuotes Or ShowNumbersAsNumbers Or ShowDatesAsDates Or _
        ShowLogicalsAsLogicals Or ShowErrorsAsErrors Or (Not ShowMissingAsNullString)

    ReDim ReturnArray(LB To LB + NumRowsInReturn - 1, LB To LB + NumColsInReturn - 1)

    For i = LBound(Lines) To UBound(Lines)
        OneRow = SplitNew(Lines(i), strDelimiter, , SplitLimit, UnusedChar)

        NumInRow = UBound(OneRow) - LBound(OneRow) + 1
        If SplitLimit > 0 Then
            If NumInRow = SplitLimit Then
                NumInRow = SplitLimit - 1
            End If
        End If

        'Ragged files: Current line has more elements than maximum length of prior lines. _
         First we need to append columns on the right of ReturnArray (Redim Preserve) then for _
         cells in rows < i, populate the columns just added with ShowMissingsAs..
        If NumCols = 0 Then
            If NumInRow - StartCol + 1 > NumColsInReturn Then
                ReDim Preserve ReturnArray(LB To LB + NumRowsInReturn - 1, LB To NumInRow - StartCol + LB)
                If Not IsEmpty(ShowMissingsAs) Then
                    For p = 1 To i
                        For q = NumColsInReturn + 1 To NumInRow
                            ReturnArray(p + LB - 1, q + LB - 1) = ShowMissingsAs
                        Next q
                    Next p
                End If
                NumColsInReturn = NumInRow - StartCol + 1
            End If
        End If

        If AnyConversion Then
            For j = 1 To MinLngs(NumColsInReturn, NumInRow - StartCol + 1)
                ReturnArray(i + LB, j + LB - 1) = CastToVariant(OneRow(j + StartCol - 2), _
                    RemoveQuotes, ShowNumbersAsNumbers, SepsStandard, DecimalSeparator, SysDecimalSeparator, _
                    ShowDatesAsDates, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, _
                    ShowLogicalsAsLogicals, ShowErrorsAsErrors, ShowMissingAsNullString, ShowMissingsAs)
            Next j
        Else
            For j = 1 To MinLngs(NumColsInReturn, NumInRow - StartCol + 1)
                ReturnArray(i + LB, j + LB - 1) = OneRow(j + StartCol - 2)
            Next j
        End If

        'Ragged files: Current line has fewer elements than maximum length of prior lines. _
         We need to pad the remainder of the current line with ShowMissingsAs.
        If NumInRow - StartCol + 1 < NumColsInReturn Then
            If Not IsEmpty(ShowMissingsAs) Then
                For j = NumInRow - StartCol + 2 To NumColsInReturn
                    ReturnArray(i + LB, j + LB - 1) = ShowMissingsAs
                Next j
            End If
        End If
    Next i

    CSVRead_V1 = ReturnArray

    Exit Function

ErrHandler:
    CSVRead_V1 = "#CSVRead_V1 (line " & CStr(Erl) + "): " & Err.Description & "!"
    If Not CSVS Is Nothing Then
        Set CSVS = Nothing
    End If
    If Not T Is Nothing Then
        T.Close
        Set T = Nothing
    End If

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
    Throw "#ParseConvertTypes: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MinLngs
' Purpose    : Returns the minimum of a & b.
' -----------------------------------------------------------------------------------------------------------------------
Private Function MinLngs(a As Long, b As Long)
    If a < b Then
        MinLngs = a
    Else
        MinLngs = b
    End If
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

    On Error GoTo ErrHandler

    Set FSO = New FileSystemObject
    Set F = FSO.GetFile(FileName)
    Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))

    If T.AtEndOfStream Then
        T.Close: Set T = Nothing: Set F = Nothing
        Throw "File is empty"
    End If

    EvenQuotes = True
    While Not T.AtEndOfStream
        Contents = T.Read(CHUNK_SIZE)
        For i = 1 To Len(Contents)
            Select Case Mid$(Contents, i, 1)
                Case DQ
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
' Procedure  : InferEOL
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
    
    On Error GoTo ErrHandler

    Set F = FSO.GetFile(FileName)
    Set T = F.OpenAsTextStream(ForReading, IIf(Unicode, TristateTrue, TristateFalse))
    If T.AtEndOfStream Then
        T.Close: Set T = Nothing: Set F = Nothing
        Throw Err_EmptyFile
    End If
    
    EvenQuotes = True
    While Not T.AtEndOfStream
        FileContents = T.Read(CHUNK_SIZE)
        LenChunk = Len(FileContents)
        If CheckFirstCharOfNextChunk Then
            If Left$(FileContents, 1) = vbLf Then
                InferEOL = vbCrLf
            Else
                InferEOL = vbCr
            End If
            GoTo EarlyExit
        End If

        For i = 1 To LenChunk
            Select Case Mid(FileContents, i, 1)
                Case DQ
                    EvenQuotes = Not EvenQuotes
                Case vbCr
                    If EvenQuotes Then
                        If i < LenChunk Then
                            If Mid$(FileContents, i + 1, 1) = vbLf Then
                                InferEOL = vbCrLf
                            Else
                                InferEOL = vbCr
                            End If
                            GoTo EarlyExit
                        ElseIf T.AtEndOfStream Then 'Mac file with only one line
                            InferEOL = vbCr
                            GoTo EarlyExit
                        Else
                            CheckFirstCharOfNextChunk = True
                        End If
                    End If
                Case vbLf
                    If EvenQuotes Then
                        InferEOL = vbLf
                        GoTo EarlyExit
                    End If
            End Select
        Next
    Wend

    'No end of line exists outside quoted regions, so the file is a single line without a _
     trailing EOL. The guess made for EOL is irrelevant.
    InferEOL = vbCrLf

EarlyExit:
    T.Close: Set T = Nothing: Set F = Nothing
    Exit Function

ErrHandler:
    Throw "#InferEOL: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SplitNew
' Purpose    : Drop-in replacement for VBA.Split, but rather than splitting Expression at all occurrences of Delimiter,
'              splits occur at only those instances that are preceded by an even number of quote characters (DQ).
'Notes
'            1) DQ must have only one character and defaults to the double-quote character.
'            2) DQ must not be contained in Delimiter, which may have more than one character.
'            3) AltDelim, should be either omitted or passed as a string which is:
'               a) the same length as Delimiter; and
'               b) not a sub-string of Expression.
'               If AltDelim is omitted and if both DQ and Delimiter are sub-strings of Expression then AltDelim is set to
'               a string satisfying those two conditions.
'            4) QuotesEncountered (optional and passed by reference) is set to indicate if string DQ is found in Expression.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SplitNew(ByVal Expression As String, Optional Delimiter As String = vbCrLf, Optional DQ As String = """", _
          Optional Limit As Long = -1, Optional ByRef AltDelim As String, Optional ByRef QuotesEncountered As Boolean)

    Dim DelimPos As Long
    Dim DQPos As Long
    Dim EvenDQs As Boolean
    Dim LDelim As Long
    Dim NDelims As Long
    Dim Ret() As String

    On Error GoTo ErrHandler
    
    If Len(Expression) = 0 Then
        QuotesEncountered = False
        ReDim Ret(0 To 0)
        SplitNew = Ret
        Exit Function
    End If

    DQPos = 0
    LDelim = Len(Delimiter)
    DelimPos = 1 - LDelim
    
    DQPos = InStr(DQPos + 1, Expression, DQ)
    DelimPos = InStr(DelimPos + LDelim, Expression, Delimiter)
    If DelimPos = 0 Then
        QuotesEncountered = DQPos > 0
        ReDim Ret(0 To 0)
        Ret(0) = Expression
        SplitNew = Ret
        Exit Function
    End If
    
    If DQPos = 0 Then
        QuotesEncountered = False
        SplitNew = VBA.Split(Expression, Delimiter, Limit)
        Exit Function
    Else
        QuotesEncountered = True
    End If
    
    EvenDQs = True
    If Len(AltDelim) <> Len(Delimiter) Then
        AltDelim = String(Len(Delimiter), CharNotInString(Expression))
    End If

    While DelimPos > 0 And Limit = -1 Or NDelims < Limit - 1
        While (DQPos < DelimPos) And DQPos > 0
            EvenDQs = Not EvenDQs
            DQPos = InStr(DQPos + 1, Expression, DQ)
        Wend
        If EvenDQs Then
            Mid$(Expression, DelimPos, LDelim) = AltDelim
            NDelims = NDelims + 1
        End If
        DelimPos = InStr(DelimPos + LDelim, Expression, Delimiter)
    Wend

    If NDelims = 0 Then
        ReDim Ret(0 To 0)
        Ret(0) = Expression
        SplitNew = Ret
        Exit Function
    End If
    
    SplitNew = VBA.Split(Expression, AltDelim, Limit)

    Exit Function
ErrHandler:
    Throw "#SplitNew: " & Err.Description & "!"
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
' Procedure  : CharNotInString
' Purpose    : Return a character not contained in specified string, or throw an error if attempt fails.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CharNotInString(Str As String) As String
    Dim i As Long
    Dim TheChar As String
    Const Err_NoAltDelimiter = "Unexpected error in method CharNotInString"
    
    On Error GoTo ErrHandler

    'Try 100 unicode characters spaced by 100, starting with a scissors emoji, scissors seem appropriate for markers at which we split a string.
    For i = 9986 To 19886 Step 100
        TheChar = ChrW(i)
        If InStr(Str, TheChar) = 0 Then
            CharNotInString = TheChar
            Exit Function
        End If
    Next i

    Throw Err_NoAltDelimiter

    Exit Function
ErrHandler:
    Throw "#CharNotInString: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ShowTextFile
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

        If MixedLineEndings Then
            ReadAll = Replace(ReadAll, vbCrLf, vbLf)
            ReadAll = Replace(ReadAll, vbCr, vbLf)
            EOL = vbLf
        End If

        'Text files may or may not be terminated by EOL characters...
        If Right$(ReadAll, Len(EOL)) = EOL Then
            ReadAll = Left$(ReadAll, Len(ReadAll) - Len(EOL))
        End If

        If Len(ReadAll) = 0 Then
            ReDim Contents1D(0 To 0)
        Else
            Contents1D = VBA.Split(ReadAll, EOL)
        End If
        ReDim Contents2D(1 To UBound(Contents1D) - LBound(Contents1D) + 1, 1 To 1)
        For i = LBound(Contents1D) To UBound(Contents1D)
            Contents2D(i + 1, 1) = Contents1D(i)
        Next i
    Else
        ReDim Contents2D(1 To NumRows, 1 To 1)

        For i = 1 To NumRows
            Contents2D(i, 1) = T.ReadLine
        Next i

        T.Close: Set T = Nothing: Set F = Nothing: Set FSO = Nothing
    End If

    ShowTextFile = Contents2D

    Exit Function
ErrHandler:
    Throw "#ShowTextFile: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToVariant
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

    If RemoveQuotes Then
        StripQuotes strIn, Converted
        If Converted Then
            CastToVariant = strIn
            Exit Function
        End If
    End If

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
' Procedure  : StripQuotes
' Purpose    : Undo how Strings are encoded when written to CSV files.
' Parameters :
'  Str         : String to be converted
'  Converted   : Boolean flag set to True if conversion happens.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub StripQuotes(ByRef Str As String, ByRef Converted As Boolean)
    If Left$(Str, 1) = DQ Then
        If Right$(Str, 1) = DQ Then
            Str = Mid$(Str, 2, Len(Str) - 2)
            Str = Replace(Str, DQ2, DQ)
            Converted = True
        End If
    End If
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble
' Purpose    : Casts string to double where string has specified decimals separator.
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
    If T.AtEndOfStream Then
        T.Close: Set T = Nothing
        IsUnicodeFile = False
        Exit Function
    End If
    intAsc1Chr = Asc(T.Read(1))
    If T.AtEndOfStream Then
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

Function OStoEOL(OS As String, ArgName As String) As String

    Const Err_Invalid = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associented end of line characters."

    Select Case LCase(OS)
        Case "windows", vbCrLf
            OStoEOL = vbCrLf
        Case "unix", vbLf
            OStoEOL = vbLf
        Case "mac", vbCr
            OStoEOL = vbCr
        Case Else
            Throw ArgName + Err_Invalid
    End Select
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
' Author     : Philip Swannell
' Date       : 08-Aug-2021
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

'---------------------------------------------------------------------------------------
' Procedure : Force2DArray
' Purpose   : In-place amendment of singletons and one-dimensional arrays to two dimensions.
'             singletons and 1-d arrays are returned as 2-d 1-based arrays. Leaves two
'             two dimensional arrays untouched (i.e. a zero-based 2-d array will be left as zero-based).
'             See also Force2DArrayR that also handles Range objects.
'---------------------------------------------------------------------------------------
Sub Force2DArray(ByRef TheArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
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

'---------------------------------------------------------------------------------------
' Procedure : ThrowIfError
' Purpose   : In the event of an error, methods intended to be callable from spreadsheets
'             return an error string (starts with "#", ends with "!"). ThrowIfError allows such
'             methods to be used from VBA code while keeping error handling robust
'             MyVariable = ThrowIfError(MyFunctionThatReturnsAStringIfAnErrorHappens(...))
'---------------------------------------------------------------------------------------
Function ThrowIfError(data As Variant)
    ThrowIfError = data
    If VarType(data) = vbString Then
        If Left$(data, 1) = "#" Then
            If Right$(data, 1) = "!" Then
                Throw CStr(data)
            End If
        End If
    End If
End Function

