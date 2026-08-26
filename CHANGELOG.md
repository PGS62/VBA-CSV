# Changelog

## [v0.33] - 2026-08-26

Performance improvements to `CSVRead`, worth around 6% on a large, mostly-numeric file (tested on a 113Mb file, 1,462 rows x 5,742 columns, `ConvertTypes:=TRUE`):

- Reduced per-field overhead in type conversion.
- Better initial sizing of internal buffers, particularly beneficial when reading only a small number of rows from a large file (e.g. via `NumRows`), which no longer allocates space as if reading the whole file.

## [v0.32] - 2025-09-10

This release fixes some edge-cases:

- If `NumRows` is larger than the number of rows in the file then the additional rows of the return from `CSVRead` are padded with `ShowMissingsAs` rather than `Empty` as before.
- Ditto for large values of `NumCols`.
- Supplying `CSVRead`'s arguments `SkipToCol` and `NumCols` as `Empty` or a blank cell is now equivalent to omitting those arguments.

## [v0.31] - 2024-06-26

Previously VBA-CSV only worked if the host application was Excel. It now works for Word and PowerPoint too.

## [v0.30] - 2023-12-02

`CSVWrite` now accepts an `Encoding` argument of `UTF-8NOBOM` and `UTF-16NOBOM` in which case the file created does not contain a [byte order mark](https://en.wikipedia.org/wiki/Byte_order_mark).

## [v0.29] - 2023-10-26

This release introduces new possibilities for handling quoted fields in which `CSVRead` never removes leading and trailing quotes from fields and `CSVWrite` never appends them. See the new `K` option for the `ConvertTypes` argument to `CSVWrite` and the new `Raw` option for the `QuoteAllStrings` argument to `CSVRead`.

## [v0.28] - 2023-03-02

In this version, `CSVRead` throws a more informative error message than previously when the file is larger than 2GB.

## [v0.27] - 2023-02-24

- Two constants now control the parsing of dates with two-digit years, as described [here](https://github.com/PGS62/VBA-CSV#two-digit-years).
- Fixes a bug in `CSVRead` where a string such as "1-May-2023" (date order `D-M-Y`) would be read as a date even if `DateFormat` was set to `M-D-Y`.

## [v0.26] - 2023-02-06

- `CSVRead`'s argument `SkipToCol`, which determines the first column read from the file, can now be passed as a string matching an element of the file's header row.
- `CSVRead`'s argument `NumCols` can now be specified as a string matching an element of the file's header row, and this determines the rightmost column read from the file.

## [v0.25] - 2022-11-28

Improved performance of `CSVWrite`.

For test data (from [this file](https://github.com/vincentarelbundock/Rdatasets/blob/master/csv/openintro/military.csv)) with 1.4 million rows and 7 columns, time to execute `CSVWrite` falls from 24 seconds to 11 seconds.

## [v0.24] - 2022-11-24

In this release `CSVWrite` has changed behaviour on PCs for which the VBA display language is not English. Error values are written to file using their English language representation, such as `#NAME?`, `#DIV/0!`, `#SPILL!` etc. This change ensures that "round-tripping" between `CSVWrite` and `CSVRead` is preserved.

## [v0.23] - 2022-11-18

Earlier versions did not work correctly if the [Office Display language](https://support.microsoft.com/en-us/office/change-the-language-office-uses-in-its-menus-and-proofing-tools-f5c54ff9-a6fa-4348-a43c-760e7ef148f8) was not English. This version fixes that.

## [v0.22] - 2022-11-09

This release has improved performance for large files.

`CSVRead` execution time is reduced by approximately 35% for large files in this release. This results from switching from using `Scripting.TextStream` to using `ADODB.Stream` for the low-level task of reading data from disk into a string variable (and working around the [performance shortcoming](https://mskb.pkisolutions.com/kb/280067) of the latter). The relevant method is [`ReadAllFromStream`](https://github.com/PGS62/VBA-CSV/blob/d890c91b18a9a66723bf9fa1c4592dbc934608c3/vba/VBA-CSV.xlsm/modCSVReadWrite.bas#L565).

The largest file in [RDatasets](https://github.com/vincentarelbundock/Rdatasets) is 76Mb in size and has 1.4 million rows and seven columns. Time to execute `CSVRead("C:\Projects\RDatasets\csv\openintro\military.csv")` falls from 22 seconds using v0.4 to 14 seconds using v0.22.

## [v0.20] - 2022-11-04

Improved code to "guess" the delimiter in a csv file when the `Delimiter` argument is not provided. Specifically, the code now handles better the case of the first field in the first row of the file being a date-with-time such as `2022-11-04T14:41:01`. Previously the code would have guessed the delimiter to be `:`, which would likely have been incorrect.

## [v0.19] - 2022-11-03

This release fixes some "edge cases", with appropriate tests added.

- Fixed bug in `CSVWrite` when argument `QuoteAllStrings` was `FALSE` and argument `Delimiter` was not a comma.
- Fixed bug in `CSVRead` when handling "ragged" files with varying number of fields per line and argument `ShowMissingAs` is provided.
- `CSVWrite` now recognises `"ISO"` as its `DateFormat` argument for ISO8601 date format (yyyy-mm-dd).

## [v0.18] - 2022-10-31

Comments at the top of the module modCSVReadWrite.bas mention how to provide intellisense help using https://github.com/Excel-DNA/IntelliSense.

## [v0.17] - 2022-10-24

- `CSVRead`: Improved handling of files in which Boolean values are represented by quoted strings. For example, [this file](https://vincentarelbundock.github.io/Rdatasets/csv/carData/TitanicSurvival.csv) represents Booleans with quoted strings `"yes"` and `"no"`. In this version v0.17 the following formula now correctly converts those strings to Booleans.
```
=CSVRead("https://vincentarelbundock.github.io/Rdatasets/csv/carData/TitanicSurvival.csv",TRUE,,,,,,,,,,,"""yes""","""no""")
```

## [v0.16] - 2022-10-21

- Bug fix: Function `CSVRead` errored (type-mismatch) if arguments `SkipToRow` and `HeaderRow` were equal and greater than 1 and argument `ConvertTypes` was not `False`.
- `CSVRead` now supports `DateFormat` such as "D M Y" i.e. with space character separator.
- `CSVWrite` has additional arguments `TrueString` and `FalseString`.

## [v0.15] - 2021-10-14

- Improved performance when `ConvertTypes` is not `FALSE`.
- Bug fix to handling of non RFC4180-compliant files.
- `CSVWrite` now accepts `Data` argument with zero or one dimension(s). In the case of one dimension, the file created has a single column, one field per row.

## [v0.14] - 2021-09-30

Performance improvement for ragged files.

## [v0.13] - 2021-09-28

Now provide modCSVReadWrite.bas as if it were a binary file. This makes the GitHub API for download-counting work correctly.

## [v0.12] - 2021-09-27

Argument `DateFormat` to `CSVRead` now defaults to `Y-M-D` if omitted, rather than being read from Windows Regional Settings.

## [v0.11] - 2021-09-24

Added tests relating to 32K field size limit.

## [v0.10] - 2021-09-24

Handle field length >32K.

## [v0.9] - 2021-09-23

Late bind to ADODB.

## [v0.8] - 2021-09-23

`CSVWrite` can now write `UTF-8` files.

## [v0.7] - 2021-09-22

Refactoring and more tests.

## [v0.6] - 2021-09-21

Improved performance when Delimited = False.

## [v0.5] - 2021-09-20

Corrected handling of files that use UTF-8 encoding, by using ADODB.Stream in place of Scripting.TextStream for such files.

## [v0.4] - 2021-09-09

Improved performance of type conversion.

## [v0.3] - 2021-09-05

Candidate for first public release.

## [v0.2] - 2021-09-01

Now compiles and passes tests on 32 bit Excel 365.

## [v0.1] - 2021-08-30

First release, tested on 64-bit Excel 365. Not yet tested on other versions of Excel.
