Function MyFunction(param1, param2)
  On Error GoTo ErrorHandler
  'Some code here
  If param1 = "" Then
    Err.Raise 13, , "Parameter 1 cannot be empty.  Additional details: " & param2
  End If
  'More code here
  Exit Function

ErrorHandler:
  MsgBox "An error occurred: " & Err.Number & " - " & Err.Description, vbCritical
  'Additional logging or handling if needed
End Function