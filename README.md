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

The documentation borrows freely from that of Julia's [CSV.jl](https://csv.juliadata.org/stable/), though sadly VBA is not capable of Julia's extremely high performance. More on performance here. For testing `CSVRead`, I also make use of the suite of test files that the authors of CSV.jl have created [here](https://github.com/JuliaData/CSV.jl/tree/main/test/testfiles).


# Documentation
#### _CSVRead_
Returns the contents of a comma-separated file on disk as an array.
```vba
Public Function CSVRead(FileName As String, Optional ConvertTypes As Variant = False, _
    Optional ByVal Delimiter As Variant, Optional IgnoreRepeated As Boolean, _
    Optional DateFormat As String, Optional Comment As String, Optional ByVal SkipToRow As Long = 1, _
    Optional ByVal SkipToCol As Long = 1, Optional ByVal NumRows As Long = 0, _
    Optional ByVal NumCols As Long = 0, Optional ByVal ShowMissingsAs As Variant = "", _
    Optional ByVal Unicode As Variant, Optional DecimalSeparator As String = vbNullString)
```

|Argument|Description|
|:-------|:----------|
|`FileName`|The full name of the file, including the path.|
|`ConvertTypes`|`ConvertTypes` provides control over whether fields in the file are converted to typed values in the return or remain as strings, and also sets the treatment of "quoted fields" and space characters.<br/><br/>`ConvertTypes` may take values FALSE (the default), TRUE, or a string of zero or more letters from "NDBETQR".<br/><br/>If `ConvertTypes` is:<br/>* FALSE then no conversion takes place other than quoted fields being unquoted.<br/>* TRUE then unquoted numbers, dates, Booleans and errors are converted, equivalent to "NDBE".<br/><br/>If `ConvertTypes` is a string including:<br/>1) "N" then fields that represent numbers are converted to numbers (Doubles).<br/>2) "D" then fields that represent dates (respecting `DateFormat`) are converted to Dates.<br/>3) "B" then fields that read true or false are converted to Booleans. The match is not case sensitive so TRUE, FALSE, True and False are also converted.<br/>4) "E" then fields that match Excel"s representation of error values are converted to error values. There are fourteen such strings, including #N/A, #NAME?, #VALUE! and #DIV/0!.<br/>5) "T" then leading and trailing spaces are trimmed from fields. In the case of quoted fields, this will not remove spaces between the quotes.<br/>6) "Q" then conversion happens for both quoted and unquoted fields; otherwise only unquoted fields are converted.<br/>7) "R" then quoted fields retain their quotes, otherwise they are "unquoted" i.e. have their leading and trailing characters removed and consecutive pairs of double-quotes replaced by a single double quote.|
|`Delimiter`|By default, `CSVRead` will try to detect a file's delimiter as the first instance of comma, tab, semi-colon, colon or pipe found outside quoted regions in the first 10,000 characters of the file. If it can't auto-detect the delimiter, it will assume comma. If your file includes a different character or string delimiter you should pass that as the `Delimiter` argument.<br/><br/>Alternatively, enter FALSE as the delimiter to treat the file as "not a delimited file". In this case the return will mimic how the file would appear in a text editor such as NotePad. The file will by split into lines at all line breaks (irrespective of double-quotes) and each element of the return will be a line of the file.|
|`IgnoreRepeated`|Whether delimiters which appear at the start of a line or immediately after another delimiter or at the end of a line, should be ignored while parsing; useful-for fixed-width files with delimiter padding between fields.|
|`DateFormat`|The format of dates in the file such as "Y-M-D", "M-D-Y" or "Y/M/D". If omitted, "Y-M-D" is assumed, to match [ISO8601](https://en.wikipedia.org/wiki/ISO_8601). Repeated D's (or M's or Y's) are equivalent to single instances, so that "Y-M-D" and "YYYY-MMM-DD" are equivalent.|
|`Comment`|Rows that start with this string will be skipped while parsing.|
|`SkipToRow`|The row in the file at which reading starts. Optional and defaults to 1 to read from the first row.|
|`SkipToCol`|The column in the file at which reading starts. Optional and defaults to 1 to read from the first column.|
|`NumRows`|The number of rows to read from the file. If omitted (or zero), all rows from `SkipToRow` to the end of the file are read.|
|`NumCols`|The number of columns to read from the file. If omitted (or zero), all columns from `SkipToCol` are read.|
|`ShowMissingsAs`|Fields which are missing in the file (consecutive delimiters) are represented by `ShowMissingsAs`. Defaults to the null string, but can be any string or Empty. If `NumRows` is greater than the number of rows in the file then the return is "padded" with the value of `ShowMissingsAs`. Likewise if `NumCols` is greater than the number of columns in the file.|
|`Unicode`|In most cases, this argument can be omitted, in which case `CSVRead` will examine the file's byte order mark to guess how the file should be opened - i.e. the correct "format" argument to pass to VBA's method OpenAsTextStream, which requires an argument indicating whether the file is `Unicode` or ASCII. Alternatively, enter TRUE for a `Unicode` file or FALSE for an ASCII file.|
|`DecimalSeparator`|In many places in the world, floating point number decimals are separated with a comma instead of a period (3,14 vs. 3.14). `CSVRead` can correctly parse these numbers by passing in the `DecimalSeparator` as a comma, in which case comma ceases to be a candidate if the parser needs to guess the `Delimiter`.|

#### _CSVWrite_
Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns `FileName`, otherwise an "error string" (starts with `#`, ends with `!`) describing what went wrong.
```vba
Public Function CSVWrite(FileName As String, ByVal data As Variant, _
    Optional QuoteAllStrings As Boolean = True, Optional DateFormat As String = "yyyy-mm-dd", _
    Optional DateTimeFormat As String = "yyyy-mm-dd hh:mm:ss", _
    Optional Delimiter As String = ",", Optional Unicode As Boolean, _
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
|`EOL`|Controls the line endings of the file written. Enter "Windows" (the default), "Unix" or "Mac". Also supports the line-ending characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13) or the strings "CRLF", "LF" or "CR". The last line of the file is written with a line ending.|
