Function MyFunction(param)
  'Some code here that does not handle the error properly 
  On Error Resume Next
  'More code that might generate an error
  If Err.Number <> 0 Then
    'Handle the error incorrectly
    MsgBox "An error occurred."
  End If
End Function