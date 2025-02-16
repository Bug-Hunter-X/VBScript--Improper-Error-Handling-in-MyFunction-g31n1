Function MyFunction(param)
  On Error GoTo ErrorHandler
  'Some code here
  'More code that might generate an error
  Exit Function
  
ErrorHandler:
  Select Case Err.Number
    Case 13: 'Type mismatch
      MsgBox "Type mismatch error: Please check the input parameter.", vbCritical
    Case 91: 'Object variable or With block variable not set
      MsgBox "Object variable not set: Please check if all objects are properly initialized.", vbCritical
    Case Else
      MsgBox "An unexpected error occurred: " & Err.Number & " - " & Err.Description, vbCritical
      'Log the error: You may write the error information to a log file here
  End Select
  Err.Clear
End Function