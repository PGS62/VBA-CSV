VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCSVStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Class module for reading CSV files, wraps Scripting.TextStream.
'Advantages over using Scripting.TextStream:
'1) Knows to ignore end of line characters that are within quoted regions,
'   or more strictly speaking, appear in the file after an odd number of double quotes.
'2) Handles (old, pre-OSX) Mac-style line endings of vbCr, as well as Windows (vbCrLf) and Unix (vbLf).

'Usage:
'1) Initialise by calling Init method, or call method CreateCSVStream
'2) Call ReadLine as many times as desired. In contrast to Scripting.TextStream.ReadLine,
'   the method is a Sub without a return value.
'3) Call ReportAllLinesRead. Return is 0-based 1-dimensional array of strings of all lines read so far.
'4) Can skip lines at the start by calling StartRecording after calling ReadLine one or more times.
'5) To read all lines make use of AtEndOfStream property within a While loop.
'6) There is no Close method, but the underlying TextStream is closed in the class Terminate event.

Private Const DQ = """"
Private m_AltEOL As String
Private m_AtEndOfStream As Boolean
Private m_Buffer As String
Private m_BufferReadFrom As Long    'The index into m_Buffer of the first character of the first line read, or the first line read after StartRecording is called, i.e. the character succeeding the EOL
Private m_BufferReadTo As Long      'The index into m_Buffer of the last character of the last line read, i.e. the character preceding the EOL
Private m_BufferUpdatedTo As Long   'The index of the last character in the buffer that has been updated from file, but not necessarily yet been read by calls to ReadLine
Private m_CHUNK_SIZE  As Long
Private m_EOL As String
Private m_EOL1stChar As String
Private m_EOL2ndChar As String
Private m_LEOL As Long
Private m_QuotesEncountered As Boolean
Private m_T As Scripting.TextStream

Public Property Get atEndOfStream() As Boolean
    atEndOfStream = m_AtEndOfStream
End Property

Public Property Get QuotesEncountered() As Boolean
    QuotesEncountered = m_QuotesEncountered
End Property

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StartRecording
' Purpose    : Set flags so that subsequent call to ReportAllLinesRead returns lines read starting from the point
'              StartRecording is called.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub StartRecording()
    m_BufferReadFrom = m_BufferReadTo + m_LEOL + 1
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Init
' Purpose    : Initializer, wrapped by function CreateCSVStream.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Init(FileName As String, EOL As String, FileIsUnicode As Boolean)
    Dim FSO As New Scripting.FileSystemObject
    Dim F As Scripting.File

    Const Err_EOL = "EOL not recognized. Allowed values are ""Windows"" (or ascii characters 13 and 10), ""Unix"" (or ascii character 10) or ""Mac"" (or ascii character 13)"

    On Error GoTo ErrHandler
    Select Case LCase(EOL)
        Case "windows", vbCrLf
            m_EOL = vbCrLf
            m_EOL2ndChar = vbLf
            m_AltEOL = Chr(0) + Chr(0)
        Case "unix", vbLf
            m_EOL = vbLf
            m_AltEOL = Chr(0)
        Case "mac", vbCr
            m_EOL = vbCr
            m_AltEOL = Chr(0)
        Case Else
            Throw Err_EOL
    End Select
    m_EOL1stChar = Left(m_EOL, 1)

    m_LEOL = Len(m_EOL)
    m_BufferReadTo = -m_LEOL
    m_BufferReadFrom = 1

    Set F = FSO.GetFile(FileName)
    Set m_T = F.OpenAsTextStream(ForReading, IIf(FileIsUnicode, TristateTrue, TristateFalse))
    m_CHUNK_SIZE = MaxLngs(1000, F.Size / 100)

    Exit Sub
ErrHandler:
    Throw "#clsCSVStream.init: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetMoreFromFile
' Purpose    : Read data from the file into the buffer. If necessary the buffer is doubled in size.
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetMoreFromFile()
    Dim NewChars As String

    On Error GoTo ErrHandler
    Dim i As Long
    Dim ReadOneMore As Boolean
    Const Err_Ascii0 = "File contains ascii character zero. In the current implementation, such files " + _
        "can only br read if arguments StartRow, StartCol, NumRows and NumCols are omitted to read the whole file. Sorry"

    For i = 1 To 2
        NewChars = m_T.Read(IIf(i = 1, m_CHUNK_SIZE, 1))
        If InStr(NewChars, Chr(0)) > 0 Then
            Throw Err_Ascii0
        End If

        If (Len(m_Buffer) - m_BufferUpdatedTo) < Len(NewChars) Then
            m_Buffer = m_Buffer & String(MaxLngs(Len(NewChars), Len(m_Buffer)), " ")
        End If
        Mid$(m_Buffer, m_BufferUpdatedTo + 1, Len(NewChars)) = NewChars
        m_BufferUpdatedTo = m_BufferUpdatedTo + Len(NewChars)

        If m_T.atEndOfStream Then
            Exit Function
        End If

        If i = 1 Then
            'Have to be careful that we don't split a two-character EOL marker in half. _
             If the last character read was vbCr then do a second pass through the _
             loop, reading only one further character from the file.
            ReadOneMore = (m_EOL = vbCrLf) And (Right(NewChars, 1) = vbCr)
            If Not ReadOneMore Then Exit For
        End If

    Next i

    Exit Function
ErrHandler:
    Throw "#clsCSVStream.GetMoreFromFile: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReportAllLinesRead
' Purpose    : Return 1-dim array of all lines read since the last call to StartRecording.
' -----------------------------------------------------------------------------------------------------------------------
Public Function ReportAllLinesRead()
    On Error GoTo ErrHandler
    Dim ZeroReturn() As String

    If m_BufferReadTo = m_BufferReadFrom - 1 Then
        ReDim ZeroReturn(0 To 0)
        ReportAllLinesRead = ZeroReturn
    Else
        ReportAllLinesRead = VBA.Split(GetChars(m_Buffer, m_BufferReadFrom, m_BufferReadTo), m_AltEOL)
    End If

    Exit Function
ErrHandler:
    Throw "#ReportAllLinesRead: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReadLine
' Purpose    : Read a line. Unlike Scripting.TextStream.ReadLine this method has no return. Instead it prepares the buffer
'              for subsequent parsing by ReportAllLinesRead by a) reading more data from file (if necessary) and b) replacing
'              instances of the the end of line characters (m_EOL) with the alternative (m_AltEOL).
' -----------------------------------------------------------------------------------------------------------------------
Public Sub ReadLine()
    Dim i As Long
    
    Dim EvenQuotes As Boolean
    
    On Error GoTo ErrHandler
    EvenQuotes = True
    i = m_BufferReadTo + m_LEOL + 1
    While True
        If i > m_BufferUpdatedTo Then 'what if zero chars in file?
            If Not m_T.atEndOfStream Then
                GetMoreFromFile
            Else
                Throw "attempt to read past end of file"
            End If
        End If
        If Mid$(m_Buffer, i, 1) = DQ Then
            m_QuotesEncountered = True
            EvenQuotes = Not EvenQuotes
        End If
        If EvenQuotes Then
            If Mid$(m_Buffer, i, 1) = m_EOL1stChar Then
                If m_LEOL = 1 Or (i < m_BufferUpdatedTo And Mid$(m_Buffer, i + 1, 1) = m_EOL2ndChar) Then
                    m_BufferReadTo = i - 1
                    SetChars m_Buffer, m_BufferReadTo + 1, m_AltEOL
                    If i = m_BufferUpdatedTo - m_LEOL + 1 Then
                        If m_T.atEndOfStream Then
                            m_AtEndOfStream = True
                        End If
                    End If
                    Exit Sub
                End If
            End If
        End If
        If i = m_BufferUpdatedTo Then
            If m_T.atEndOfStream Then 'Reading the last line of the file, which was not terminated with EOL
                m_BufferReadTo = i
                m_AtEndOfStream = True
                Exit Sub
            End If
        End If
        i = i + 1
    Wend
    Exit Sub
ErrHandler:
    Throw "#clsCSVStream.ReadLine: " & Err.Description & "!"
End Sub

Private Function GetChars(x As String, First As Long, Last As Long)
    GetChars = Mid$(x, First, Last - First + 1)
End Function

Private Function SetChars(x As String, First As Long, InsertThis As String)
    Mid$(x, First, Len(InsertThis)) = InsertThis
End Function

Private Function MaxLngs(x As Long, y As Long) As Long
    If x > y Then
        MaxLngs = x
    Else
        MaxLngs = y
    End If
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    m_T.Close
    Set m_T = Nothing
    m_Buffer = ""
End Sub

'Copy of Throw, to make this module  self-contained.
Private Sub Throw(ByVal ErrorString As String)
    Throw ErrorString
End Sub

