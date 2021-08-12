# VBA-CSV
CSV reading and writing for VBA and Excel spreadsheets, via two functions `CSVRead` and `CSVWrite`.

# Installation
1. Download the latest release.
2. Import `modCSVReadWrite.bas` into your project (Open VBA Editor, `Alt + F11`; File > Import File).
3. Include a reference to "Microsoft Scripting Runtime" (In VBA Editor Tools > References).
4. If you plan to call the functions from spreadsheet formulas then you might like to tell Excel's Function Wizard about them by adding calls to `RegisterCSVRead` and `RegisterCSVWrite` to the project's `Workbook_Open` event. Example:
```vba
Private Sub Workbook_Open()
    RegisterCSVWrite
    RegisterCSVRead
End Sub
```

# Documentation
#### _CSVRead_
Returns the contents of a comma-separated file on disk as an array.
```vba
Public Function CSVRead(FileName As String, Optional ConvertTypes As Variant = False, _
    Optional ByVal Delimiter As Variant, Optional DateFormat As String, _
    Optional ByVal SkipToRow As Long = 1, Optional ByVal SkipToCol As Long = 1, _
    Optional ByVal NumRows As Long = 0, Optional ByVal NumCols As Long = 0, _
    Optional ByVal ShowMissingsAs As Variant = """", Optional ByVal UTF16 As Variant, _
    Optional DecimalSeparator As String = vbNullString)
```

|Argument|Description|
|:-------|:----------|
|`FileName`|The full name of the file, including the path.|
|`ConvertTypes`|TRUE to convert Numbers, Dates, Booleans and Errors into their typed values, or FALSE to leave as strings. Or a string of characters N, D, B, E, Q. E.g. "BN" to convert just Booleans and numbers. Q indicates quoted strings should retain their quotes.|
|`Delimiter`|Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe found outside quoted regions. Enter FALSE to  see the file's raw contents as would be displayed in a text editor. `Delimiter` may have more than one character.|
|`DateFormat`|The format of dates in the file such as "D-M-Y", "M-D-Y" or "Y/M/D". If omitted a value is read from Windows regional settings. Repeated D's (or M's or Y's) are equivalent to single instances, so that "d-m-y" and "DD-MMM-YYYY" are equivalent.|
|`SkipToRow`|The row in the file at which reading starts. Optional and defaults to 1 to read from the first row.|
|`SkipToCol`|The column in the file at which reading starts. Optional and defaults to 1 to read from the first column.|
|`NumRows`|The number of rows to read from the file. If omitted (or zero), all rows from `SkipToRow` to the end of the file are read.|
|`NumCols`|The number of columns to read from the file. If omitted (or zero), all columns from `SkipToCol` are read.|
|`ShowMissingsAs`|Fields which are missing in the file (i.e. consecutive delimiters) are represented by `ShowMissingsAs`. Defaults to the null string, but can be any string or Empty.|
|`UTF16`|Enter TRUE if the file is UTF-16 encoded, FALSE otherwise. Omit to guess from the file's contents.|
|`DecimalSeparator`|The character that represents a decimal point. If omitted, then the value from Windows regional settings is used.|

#### _CSVWrite_
Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an "error string" (starts with #, ends with !) describing what went wrong.
```vba
Public Function CSVWrite(FileName As String, ByVal Data As Variant, _
    Optional QuoteAllStrings As Boolean = True, Optional DateFormat As String = ""yyyy-mm-dd"", _
    Optional DateTimeFormat As String = ""yyyy-mm-dd hh:mm:ss"", _
    Optional Delimiter As String = "","", Optional UTF16 As Boolean, _
    Optional ByVal EOL As String = vbCrLf)
```

|Argument|Description|
|:-------|:----------|
|`FileName`|The full name of the file, including the path.|
|`Data`|An array of `Data`. Elements may be strings, numbers, dates, Booleans, empty, Excel errors or null values.|
|`QuoteAllStrings`|If TRUE (the default) then all strings in `Data` are quoted before being written to file. If FALSE only strings containing `Delimiter`, line feed, carriage return or double quote are quoted. Double quotes are always escaped by another double quote.|
|`DateFormat`|A format string that determine how dates, including cells formatted as dates, appear in the file. If omitted, defaults to "yyyy-mm-dd".|
|`DateTimeFormat`|A format string that determines how dates with non-zero time part appear in the file. If omitted defaults to "yyyy-mm-dd hh:mm:ss".The companion function CVSRead is not currently capable of interpreting fields written in DateTime format.|
|`Delimiter`|The `Delimiter` string, if omitted defaults to a comma. `Delimiter` may have more than one character.|
|`UTF16`|If FALSE (the default) the file written will be UTF-8. If TRUE the file written will be UTF-16 LE BOM. An error will result if this argument is FALSE but `Data` contains characters with code points above 255.|
|`EOL`|Enter the line ending as "Windows" (or CRLF), or "Unix" (or LF) or "Mac" (or CR). If omitted defaults to "Windows".|
