Function MyFunction(param1)
  If IsEmpty(param1) Then
    Err.Raise 13, , "Type mismatch: Expected parameter"
  End If
  ' ... rest of the function ...
End Function