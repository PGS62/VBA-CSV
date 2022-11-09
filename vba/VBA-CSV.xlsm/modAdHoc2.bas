Attribute VB_Name = "modAdHoc2"
Option Explicit

Function TestSpeed()
          Dim t1 As Double, t2 As Double, N As Long, Res As Variant, i As Long

          Const FileName = "C:\Projects\VBA-CSV\testfiles\single_character.csv"

1         N = 10000

2         t1 = ElapsedTime()

3         For i = 1 To N
              '~3.8E-04
4             Res = CSVRead(FileName, False, ",", , , , , , , , , , , , , , "ANSI")
              
              '~6.5E-04
              'res = CSVRead(FileName, False)
5         Next i
6         t2 = ElapsedTime()

7         TestSpeed = (t2 - t1) / N
8         Debug.Print TestSpeed

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReThrow
' Purpose    : Common error handling. For use in the error handler of all methods
' Parameters :
'  FunctionName: The name
'  Error       : The Err error object
'  TopLevel    : Set to True if the method is a "top level" method that's exposed to the user and we wish for the function
'                to return an error string (starts with #, ends with !) rather than raise an error.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ReThrow(FunctionName As String, Error As ErrObject, Optional TopLevel As Boolean = False)
          Dim ErrorDescription As String
          Dim LineDescription As String
          Dim ErrorNumber As Long
          
1         ErrorDescription = Error.Description
2         ErrorNumber = Err.Number
          
3         If ErrorNumber <> vbObjectError + 100 Or TopLevel Then
              'Build up call stack, i.e. annotate error description by prepending with #FunctionName and appending !
4             If Erl <> 0 Then
                  'Code has line numbers, annotate with line number
5                 LineDescription = " (line " & CStr(Erl) & "): "
6             Else
7                 LineDescription = ": "
8             End If
9             ErrorDescription = "#" & FunctionName & LineDescription & ErrorDescription & "!"
10        End If

11        If TopLevel Then
12            ReThrow = ErrorDescription
13        Else
14            Err.Raise ErrorNumber, , ErrorDescription
15        End If

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Throw
' Purpose    : Simple error handling.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Throw(ByVal ErrorString As String, Optional WithCallStack As Boolean = False)
1         Err.Raise vbObjectError + IIf(WithCallStack, 1, 100), , ErrorString
End Sub

Private Function Foo()
1         On Error GoTo ErrHandler
2         Foo = Bar()
3         Exit Function
ErrHandler:
4         Foo = ReThrow("Foo", Err, True)
End Function

Private Function Bar()
1         On Error GoTo ErrHandler
2         Bar = Baz()
3         Exit Function
ErrHandler:
4         ReThrow "Bar", Err
End Function

Private Function Baz()
1         On Error GoTo ErrHandler
2         Baz = Grumpy()
3         Exit Function
ErrHandler:
4         ReThrow "Baz", Err
End Function

Private Function Grumpy()
1         On Error GoTo ErrHandler
2         Throw "WTF"
3         Exit Function
ErrHandler:
4         ReThrow "Grumpy", Err
End Function



