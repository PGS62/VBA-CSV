Attribute VB_Name = "modAdHoc"
Option Explicit

Public Function TestStringLengthLimit(N As Long) As Variant
    Dim Result(1 To 2, 1 To 1) As Variant

    Result(1, 1) = String(N, "x")
    Result(2, 1) = N 'Result needs to hold both Longs and Strings

    TestStringLengthLimit = Result

End Function

'Tested impact of encoding on reading speed. Minimal impact for large (>100Mb) files.
'====================================================================================================
'NumRows 0
'ArraysIdentical(Res1, Res2) True
'ArraysIdentical(Res1, Res3) True
'ansi 62.24525919999
'UTF-8          57.4321774999844
'UTF-16         63.1366654000012

Private Sub SpeedTestEncoding()

    Const File1 As String = "c:\Temp\25millionchars_ansi.csv"   '100 Mb
    Const File2 As String = "c:\Temp\25millionchars_utf-8.csv"  '114 Mb
    Const File3 As String = "c:\Temp\25millionchars_utf-16.csv" '199 Mb

    Const NumRows As Long = 0 'ALL 100,000 rows
    Dim res1 As Variant
    Dim res2 As Variant
    Dim Res3 As Variant
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim t4 As Double

    t1 = ElapsedTime()
    res1 = CSVRead(File1, , , , , , , , , , NumRows)
    t2 = ElapsedTime()
    res2 = CSVRead(File2, , , , , , , , , , NumRows)
    t3 = ElapsedTime()
    Res3 = CSVRead(File3, , , , , , , , , , NumRows)
    t4 = ElapsedTime()

    Debug.Print String(100, "=")
    Debug.Print "NumRows", NumRows
    Debug.Print "ArraysIdentical(Res1, Res2)", ArraysIdentical(res1, res2)
    Debug.Print "ArraysIdentical(Res1, Res3)", ArraysIdentical(res1, Res3)
    Debug.Print "ANSI", t2 - t1
    Debug.Print "UTF-8", t3 - t2
    Debug.Print "UTF-16", t4 - t3
End Sub

