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
        Optional DateFormat As String, Optional Comment As String, Optional IgnoreEmptyLines As Boolean = True, _
        Optional ByVal SkipToRow As Long = 1, Optional ByVal SkipToCol As Long = 1, _
        Optional ByVal NumRows As Long = 0, Optional ByVal NumCols As Long = 0, _
        Optional TrueStrings As Variant, Optional FalseStrings As Variant, _
        Optional MissingStrings As Variant, Optional ByVal ShowMissingsAs As Variant, _
        Optional ByVal Encoding As Variant, Optional DecimalSeparator As String = vbNullString)
```

|Argument|Description|
|:-------|:----------|
|`FileName`|Either:<br/>* The name of a file, with path.<br/>* A URL of a file.<br/>* A string in CSV format.|
|`ConvertTypes`|`ConvertTypes` provides control over whether fields in the file are converted to typed values in the return or remain as strings, and also sets the treatment of "quoted fields" and space characters.<br/><br/>`ConvertTypes` may take values FALSE (the default), TRUE, or a string of zero or more letters from "NDBETQR".<br/><br/>If `ConvertTypes` is:<br/>* FALSE then no conversion takes place other than quoted fields being unquoted.<br/>* TRUE then unquoted numbers, dates, Booleans and errors are converted, equivalent to "NDBE".<br/><br/>If `ConvertTypes` is a string including:<br/>1) "N" then fields that represent numbers are converted to numbers (Doubles).<br/>2) "D" then fields that represent dates (respecting `DateFormat`) are converted to Dates.<br/>3) "B" then fields that read true or false are converted to Booleans. The match is not case sensitive so TRUE, FALSE, True and False are also converted.<br/>4) "E" then fields that match Excel"s representation of error values are converted to error values. There are fourteen such strings, including #N/A, #NAME?, #VALUE! and #DIV/0!.<br/>5) "T" then leading and trailing spaces are trimmed from fields. In the case of quoted fields, this will not remove spaces between the quotes.<br/>6) "Q" then conversion happens for both quoted and unquoted fields; otherwise only unquoted fields are converted.<br/>7) "R" then quoted fields retain their quotes, otherwise they are "unquoted" i.e. have their leading and trailing characters removed and consecutive pairs of double-quotes replaced by a single double quote.|
|`Delimiter`|By default, `CSVRead` will try to detect a file's delimiter as the first instance of comma, tab, semi-colon, colon or pipe found outside quoted regions in the first 10,000 characters of the file. If it can't auto-detect the delimiter, it will assume comma. If your file includes a different character or string delimiter you should pass that as the `Delimiter` argument.<br/><br/>Alternatively, enter FALSE as the delimiter to treat the file as "not a delimited file". In this case the return will mimic how the file would appear in a text editor such as NotePad. The file will by split into lines at all line breaks (irrespective of double-quotes) and each element of the return will be a line of the file.|
|`IgnoreRepeated`|Whether delimiters which appear at the start of a line or immediately after another delimiter or at the end of a line, should be ignored while parsing; useful-for fixed-width files with delimiter padding between fields.|
|`DateFormat`|The format of dates in the file such as "Y-M-D", "M-D-Y" or "Y/M/D". If omitted, "Y-M-D" is assumed, to match [ISO8601](https://en.wikipedia.org/wiki/ISO_8601). Repeated D's (or M's or Y's) are equivalent to single instances, so that "Y-M-D" and "YYYY-MMM-DD" are equivalent.|
|`Comment`|Rows that start with this string will be skipped while parsing.|
|`IgnoreEmptyLines`|Whether empty rows/lines in the file should be skipped while parsing (if FALSE, each column will be assigned `ShowMissingsAs` for that empty row).|
|`SkipToRow`|The row in the file at which reading starts. Optional and defaults to 1 to read from the first row.|
|`SkipToCol`|The column in the file at which reading starts. Optional and defaults to 1 to read from the first column.|
|`NumRows`|The number of rows to read from the file. If omitted (or zero), all rows from `SkipToRow` to the end of the file are read.|
|`NumCols`|The number of columns to read from the file. If omitted (or zero), all columns from `SkipToCol` are read.|
|`TrueStrings`|Indicates how `True` values are represented in the file. May be a string, an array of strings or a range containing strings; by default "TRUE", "True" and "true" are recognised.|
|`FalseStrings`|Indicates how `False` values are represented in the file. May be a string, an array of strings or a range containing strings; by default "FALSE", "False" and "false" are recognised.|
|`MissingStrings`|Indicates how missing values are represented in the file. May be a string, an array of strings or a range containing strings. By default only an empty field (consecutive delimiters) is considered missing.|
|`ShowMissingsAs`|Fields which are missing in the file (consecutive delimiters) or match one of the `MissingStrings` are returned in the array as `ShowMissingsAs`. Defaults to Empty, but the null string or #N/A! error value can be good alternatives.<br/><br/>If `NumRows` is greater than the number of rows in the file then the return is "padded" with the value of `ShowMissingsAs`. Likewise if `NumCols` is greater than the number of columns in the file.|
|`Encoding`|Allowed entries are "UTF-16", "UTF-8", "UTF-8-BOM", and "ANSI", but for most files this argument can be omitted and `CSVRead` will detect the file's encoding. If auto-detection does not work then it's possible that the file is encoded UTF-16 but without a byte option mark, so try entering `Encoding` as "UTF-16".|
|`DecimalSeparator`|In many places in the world, floating point number decimals are separated with a comma instead of a period (3,14 vs. 3.14). `CSVRead` can correctly parse these numbers by passing in the `DecimalSeparator` as a comma, in which case comma ceases to be a candidate if the parser needs to guess the `Delimiter`.|

#### _CSVWrite_
Creates a comma-separated file on disk containing Data. Any existing file of the same name is overwritten. If successful, the function returns FileName, otherwise an "error string" (starts with #, ends with !) describing what went wrong.
```vba
Public Function CSVWrite(ByVal Data As Variant, Optional FileName As String, _
    Optional QuoteAllStrings As Boolean = True, Optional DateFormat As String = "yyyy-mm-dd", _
    Optional DateTimeFormat As String = "yyyy-mm-ddThh:mm:ss", _
    Optional Delimiter As String = ",", Optional Unicode As Boolean, _
    Optional ByVal EOL As String = "")
```

|Argument|Description|
|:-------|:----------|
|`Data`|An array of data, or an Excel range. Elements may be strings, numbers, dates, Booleans, empty, Excel errors or null values.|
|`FileName`|The full name of the file, including the path. Alternatively, if `FileName` is omitted, then the function returns a string of `Data` converted to CSV format.|
|`QuoteAllStrings`|If TRUE (the default) then elements of `Data` that are strings are quoted before being written to file, other elements (Numbers, Booleans, Errors) are not quoted. If FALSE then the only elements of `Data` that are quoted are strings containing `Delimiter`, line feed, carriage return or double quote. In all cases, double quotes are escaped by another double quote.|
|`DateFormat`|A format string that determine how dates, including cells formatted as dates, appear in the file. If omitted, defaults to "yyyy-mm-dd".|
|`DateTimeFormat`|A format string that determines how dates with non-zero time part appear in the file. If omitted defaults to "yyyy-mm-dd hh:mm:ss".The companion function `CSVRead` is not capable of converting fields written in DateTime format back from strings into Dates.|
|`Delimiter`|The delimiter string, if omitted defaults to a comma. `Delimiter` may have more than one character.|
|`Unicode`|If FALSE (the default) the file written will be encoded UTF-8. If TRUE the file written will be encoded UTF-16 LE BOM. An error will result if this argument is FALSE but `Data` contains strings with characters whose code points exceed 255.|
|`EOL`|Controls the line endings of the file written. Enter "Windows" (the default), "Unix" or "Mac". Also supports the line-ending characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13) or the strings "CRLF", "LF" or "CR". The last line of the file is written with a line ending.|
