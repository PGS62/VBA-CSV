Attribute VB_Name = "modCSVTestRoundTrip"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RoundTripTest
' Purpose    : Tests multiple times that CSVRead correctly round-trips data previously saved to disk by CSVWrite.
'              Tests include:
'           *  Embedded line feeds in quoted strings.
'           *  Files with Windows, Unix or (old) Mac line endings.
'           *  Both unicode and ascii files.
'           *  That the delimiter is automatically detected by CSVRead (reliable only if files have all strings quoted).
'           *  That unicode vs ascii is automatically detected.
'           *  That line endings are automatically detected.
'Results are printed to the VBA immediate window, if a difference is detected
' -----------------------------------------------------------------------------------------------------------------------
Sub RoundTripTest()

    Dim AllowLineFeed As Variant
    Dim data As Variant
    Dim DateFormat As Variant
    Dim Delimiter As Variant
    Dim EOL As String
    Dim ExtraInfo As String
    Dim Folder As String
    Dim k As Long
    Dim NCols As Variant
    Dim NRows As Variant
    Dim NumTests As Long
    Dim OS As Variant
    Dim Unicode As Variant
    Dim WhatDiffers As String
    
    On Error GoTo ErrHandler
    
    Folder = Environ("Temp") & "\VBA-CSV\RoundTripTests"

    ThrowIfError CreatePath(Folder)

    For Each OS In Array("Windows", "Unix", "Mac")
        EOL = IIf(OS = "Windows", vbCrLf, IIf(OS = "Unix", vbLf, vbCr))
    
        For Each Unicode In Array(True, False)
            For Each Delimiter In Array(",", "::::")
                For Each NRows In Array(1, 2, 20)
                    For Each NCols In Array(1, 2, 10)
              
                        'For Variants we need to vary AllowLineFeed and DateFormat
                        For Each AllowLineFeed In Array(True, False)
                            For Each DateFormat In Array("mmm-dd-yyyy", "dd-mmm-yyyy", "yyyy-mm-dd")
                                data = RandomVariants(CLng(NRows), CLng(NCols), CBool(AllowLineFeed), CBool(Unicode), EOL)
                                NumTests = NumTests + 1
                                ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomVariants" & IIf(AllowLineFeed, "WithLineFeed", "")
                                RoundTripTestCore Folder, CStr(OS), data, CStr(DateFormat), CBool(Unicode), CStr(OS), CStr(Delimiter), ExtraInfo, WhatDiffers
                                    
                            Next DateFormat
                        Next AllowLineFeed

                        'For Dates, we need to vary DateFormat
                        For Each DateFormat In Array("mmm-dd-yyyy", "dd-mmm-yyyy", "yyyy-mm-dd")
                            data = RandomDates(CLng(NRows), CLng(NCols))
                            NumTests = NumTests + 1
                            ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomDates"
                            RoundTripTestCore Folder, CStr(OS), data, CStr(DateFormat), CBool(Unicode), CStr(OS), CStr(Delimiter), ExtraInfo, WhatDiffers
                                
                        Next DateFormat

                        'For Strings, we need to vary AllowLineFeed
                        For Each AllowLineFeed In Array(True, False)
                            data = RandomStrings(CLng(NRows), CLng(NCols), CBool(Unicode), CBool(AllowLineFeed), EOL)
                            NumTests = NumTests + 1
                            ExtraInfo = "Test " & CStr(NumTests) & " " & IIf(AllowLineFeed, "RandomStringsWithLineFeeds", "RandomStrings")
                            RoundTripTestCore Folder, CStr(OS), data, CStr(DateFormat), CBool(Unicode), CStr(OS), CStr(Delimiter), ExtraInfo, WhatDiffers
                        Next AllowLineFeed

                        For k = 1 To 4
                            NumTests = NumTests + 1
                            If k = 1 Then
                                data = RandomBooleans(CLng(NRows), CLng(NCols))
                                ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomBooleans"
                            ElseIf k = 2 Then
                                data = RandomDoubles(CLng(NRows), CLng(NCols))
                                ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomDoubles"
                            ElseIf k = 3 Then
                                data = RandomErrorValues(CLng(NRows), CLng(NCols))
                                ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomErrorValues"
                            ElseIf k = 4 Then
                                data = RandomLongs(CLng(NRows), CLng(NCols))
                                ExtraInfo = "Test " & CStr(NumTests) & " " & "RandomLongs"
                            End If
                            RoundTripTestCore Folder, CStr(OS), data, CStr(DateFormat), CBool(Unicode), CStr(OS), CStr(Delimiter), ExtraInfo, WhatDiffers
                        Next k
                        'Print a heartbeat...
                        If NumTests Mod 10 = 0 Then Debug.Print NumTests
                        DoEvents 'Kick Immediate window back to life?
                    Next NCols
                Next NRows
            Next Delimiter
        Next Unicode
    Next OS
    Debug.Print "Finished"

    Exit Sub
ErrHandler:
    MsgBox "#RoundTripTest: " & Err.Description & "!", vbCritical
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RoundTripTestCore
' Purpose    : Test for "round trip" between functions CSVRead and CSVWrite. We write data to a file, then read it back
'              and test that the read-back data is identical to the starting data. For round-tripping to work we must write
'              with QuoteAllStrings being TRUE and read back with ShowMissingsAs being Empty (to be able to distinguish
'              Empty and null string. Also method RandomDoubles only generates doubles that have exact representation as
'              strings (avoid errors of order 10E-15).
' -----------------------------------------------------------------------------------------------------------------------
Function RoundTripTestCore(Folder As String, OS As String, ByVal data As Variant, DateFormat As String, Unicode As Boolean, EOL As String, Delimiter As String, ExtraInfo As String, ByRef WhatDiffers As String)

          Dim DataReadBack

1         On Error GoTo ErrHandler
          Dim FileName As String
          Dim NR As Long
          Dim NC As Long
2         WhatDiffers = ""

3         NR = sNRows(data)
4         NC = sNCols(data)

5         FileName = NameThatFile(Folder, OS, NR, NC, ExtraInfo, CBool(Unicode), False)

6         ThrowIfError CSVWrite(FileName, data, True, DateFormat, , Delimiter, Unicode, EOL)

          'The Call to CSVRead has to infer both Encoding and EOL
7         DataReadBack = CSVRead(FileName, True, Delimiter, DateFormat:=DateFormat, ShowMissingsAs:=Empty)

8         If Not sArraysIdentical(data, DataReadBack, True, False, WhatDiffers) Then
9             Debug.Print FileName
10            Debug.Print WhatDiffers
11        End If

12        Exit Function
ErrHandler:
13        Throw "#RoundTripTestCore: " & Err.Description & "!"
End Function

