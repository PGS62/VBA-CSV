Attribute VB_Name = "modExperiments"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestInStrIdea
' Author     : Philip Swannell
' Date       : 04-Aug-2021
' Purpose    : Comments in sdkn's function StrCount implied that it's slow to pass long string to InStr, but I can find no
'              evidence that that is true.
' Parameters :
' -----------------------------------------------------------------------------------------------------------------------
Sub TestInStrIdea()

    Dim LongStr As String, t1 As Double, t2 As Double, Res As Long
    Dim ShortStr As String
    Dim NumTests As Long, i As Long

    NumTests = 100000

    LongStr = String(100000000, "x")
    Debug.Print Len(LongStr)
    Mid(LongStr, 5, 1) = "y"

    ShortStr = String(10, "x")
    Mid(ShortStr, 5, 1) = "y"
    
    t1 = sElapsedTime()
    For i = 1 To NumTests
        Res = InStr(LongStr, "y")
    Next i
    t2 = sElapsedTime()
    Debug.Print t2 - t1

    t1 = sElapsedTime()
    For i = 1 To NumTests
        Res = InStr(ShortStr, "y")
    Next i
    t2 = sElapsedTime()
    Debug.Print t2 - t1

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestSplitVsConstruct
' Author     : Philip Swannell
' Date       : 05-Aug-2021
' Purpose    : Test strategy for building a 2-d array of strings from sub-strings of one longstring, the code simplifies the problem
'              by making each sub-string identical. In reality in both cases there needs to be a preparatory stage at which the breakpoints are identified
' Parameters :
' -----------------------------------------------------------------------------------------------------------------------
Sub TestSplitVsConstruct()

    Dim NumRows As Long
    Dim NumCols As Long
    Dim EOL As String
    Dim CellContents As String
    Dim Result1() As String
    Dim Result2() As String

    NumRows = 10000
    NumCols = 100
    CellContents = "Hello"
    EOL = vbCrLf

    Dim CC2 As String

    Dim LongString
    Dim t1 As Double, t2 As Double, t3 As Double, t4 As Double
    Dim i As Long, j As Long, k As Long
    Dim ChunkLength As Long
    Dim LCellContents As Long
    Dim SplitRes() As String

    CC2 = CellContents + EOL

    ChunkLength = Len(CellContents) + Len(EOL)
    LCellContents = Len(CellContents)

    LongString = String(NumRows * NumCols * ChunkLength, " ")

    For i = 1 To NumRows * NumCols
        Mid(LongString, 1 + (i - 1) * ChunkLength, ChunkLength) = CC2
    Next

    t1 = sElapsedTime()
    k = 0
    ReDim Result1(1 To NumRows, 1 To NumCols)
    For i = 1 To NumRows
        For j = 1 To NumCols
            k = k + 1
            Result1(i, j) = Mid(LongString, 1 + (k - 1) * ChunkLength, LCellContents)
        Next j
    Next i
    t2 = sElapsedTime()

    k = 0
    ReDim Result2(1 To NumRows, 1 To NumCols)
    SplitRes = VBA.Split(LongString, EOL)
    For i = 1 To NumRows
        For j = 1 To NumCols
            Result2(i, j) = SplitRes(k)
            k = k + 1
        Next j
    Next i
    t3 = sElapsedTime()
    
    Debug.Print String(50, "=")
    Debug.Print "NumRows = " & NumRows & ", NumCols = " & NumCols & ", CellLength = " & LCellContents
    Debug.Print "Not using Split", t2 - t1
    Debug.Print "Using Split    ", t3 - t2
    Debug.Print "Ratio          ", (t3 - t2) / (t2 - t1)
    Debug.Print "Results Identical", Application.Run("sArraysIdentical", Result1, Result2)

End Sub

