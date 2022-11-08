Attribute VB_Name = "Module1"
Option Explicit

Function TestSpeed()
    Dim t1 As Double, t2 As Double, N As Long, Res As Variant, i As Long

    Const FileName = "C:\Projects\VBA-CSV\testfiles\single_character.csv"

    N = 10000

    t1 = ElapsedTime()

    For i = 1 To N
        '~3.8E-04
        Res = CSVRead(FileName, False, ",", , , , , , , , , , , , , , "ANSI")
        
        '~6.5E-04
        'res = CSVRead(FileName, False)
    Next i
    t2 = ElapsedTime()

    TestSpeed = (t2 - t1) / N
Debug.Print TestSpeed


End Function
