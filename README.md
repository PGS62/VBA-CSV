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
# Acknowledgements
I re-wrote the parsing code of `CSVRead` after examining "sdkn104"'s code available [here](https://github.com/sdkn104/VBA-CSV); my approach is now similar to the one employed there.

The documentation borrows freely from that of Julia's [CSV.jl](https://csv.juliadata.org/stable/), though sadly VBA is not capable of Julia's extremely high performance. More on performance here.


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
|`ConvertTypes`|`ConvertTypes` provides control over whether fields in the file are converted to typed values in the returned array or remain as strings.<br/><br/>In all cases, only "unquoted" fields are converted. A field is "quoted" if its first and last characters are double-quotes, otherwise it is unquoted.<br/><br/>`ConvertTypes` may take values FALSE (the default), TRUE, or a string made up of the letters "N", "D", "B", "E" or "Q".<br/><br/>Four possible type conversions are available:<br/>1) If `ConvertTypes` includes the letter "N" then unquoted fields that represent numbers are converted to numbers (of type Double).<br/>2) If `ConvertTypes` includes the letter "D" then unquoted fields that represent dates (respecting `DateFormat`) are converted to Dates.<br/>3) If `ConvertTypes` includes the letter "B" then unquoted fields that read "true" or "false" are converted to Booleans. The match is not case sensitive so "TRUE", "FALSE", "True" and "False" are also converted.<br/>4) If `ConvertTypes` includes the letter "E" then unquoted fields that match Excel"s representation of error values are converted to error values. There are fourteen such strings, including "#N/A", "#NAME?", "#VALUE!" and "#DIV/0!".<br/><br/>For convenience, `ConvertTypes` can also take the value TRUE - all four conversions take place, or FALSE - no type conversion takes place. If `ConvertTypes` is omitted no type conversion takes place.<br/><br/>Quoted fields are returned with their leading and trailing double-quote characters removed and consecutive pairs of double-quotes replaced by single double-quotes. But if `ConvertTypes` contains the letter "Q" then this behaviour is changed so that quoted fields are returned exactly as they appear in the file, with all double-quotes retained.|
|`Delimiter`|By default, `CSVRead` will try to detect a file's delimiter as the first instance of comma, tab, semi-colon, colon or pipe found outside quoted regions. If it can't auto-detect the delimiter, it will assume comma. If your file includes a different character or string delimiter you should pass that as the `Delimiter` argument.|
|`DateFormat`|The format of dates in the file such as "D-M-Y", "M-D-Y" or "Y/M/D". If omitted a value is read from Windows regional settings. Repeated D's (or M's or Y's) are equivalent to single instances, so that "d-m-y" and "DD-MMM-YYYY" are equivalent.|
|`SkipToRow`|The row in the file at which reading starts. Optional and defaults to 1 to read from the first row.|
|`SkipToCol`|The column in the file at which reading starts. Optional and defaults to 1 to read from the first column.|
|`NumRows`|The number of rows to read from the file. If omitted (or zero), all rows from `SkipToRow` to the end of the file are read.|
|`NumCols`|The number of columns to read from the file. If omitted (or zero), all columns from `SkipToCol` are read.|
|`ShowMissingsAs`|Fields which are missing in the file (consecutive delimiters) are represented by `ShowMissingsAs`. Defaults to the null string, but can be any string or Empty. If `NumRows` is greater than the number of rows in the file then the return is "padded" with the value of `ShowMissingsAs`. Likewise if `NumCols` is greater than the number of columns in the file.|
|`Unicode`|In most cases, this argument can be omitted, in which case `CSVRead` will examine the file's byte order mark to guess how the file should be opened - i.e. the correct "format" argument to pass to VBA's method OpenAsTextStream, which requires an argument indicating whether the file is `Unicode` or ASCII. Alternatively, enter TRUE for a `Unicode` file or FALSE for an ASCII file.|
|`DecimalSeparator`|In many places in the world, floating point number decimals are separated with a comma instead of a period (3,14 vs. 3.14). `CSVRead` can correctly parse these numbers by passing in the `DecimalSeparator` as a comma. Note that you probably need to explicitly pass `Delimiter` in this case, since the parser will probably think that it detected comma as the delimiter.|

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
|`Data`|An array of data. Elements may be strings, numbers, dates, Booleans, empty, Excel errors or null values.|
|`QuoteAllStrings`|If TRUE (the default) then all strings in `Data` are quoted before being written to file. If FALSE only strings containing `Delimiter`, line feed, carriage return or double quote are quoted. Double quotes are always escaped by another double quote.|
|`DateFormat`|A format string that determine how dates, including cells formatted as dates, appear in the file. If omitted, defaults to "yyyy-mm-dd".|
|`DateTimeFormat`|A format string that determines how dates with non-zero time part appear in the file. If omitted defaults to "yyyy-mm-dd hh:mm:ss".The companion function `CSVRead` is not capable of converting fields written in DateTime format back from strings into Dates.|
|`Delimiter`|The delimiter string, if omitted defaults to a comma. `Delimiter` may have more than one character.|
|`Unicode`|If FALSE (the default) the file written will be encoded UTF-8. If TRUE the file written will be encoded UTF-16 LE BOM. An error will result if this argument is FALSE but `Data` contains strings with characters whose code points exceed 255.|
|`EOL`|Controls the line endings of the file written. Enter "Windows" (the default), "Unix" or "Mac". Also supports the line-ending characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13) or the strings "CRLF", "LF" or "CR".|
