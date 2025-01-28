Function MyFunction(param1, param2)
  On Error Resume Next
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 5, , "Parameters cannot be empty"
  End If
  On Error GoTo ErrHandler
  ' ... rest of the function
  Exit Function
ErrHandler:
  MsgBox "An error occurred: " & Err.Description, vbCritical
  Err.Clear
End Function