Function MyFunction(param1)
  If IsEmpty(param1) Or IsNull(param1) Then
    ' Handle the empty/null parameter gracefully
    param1 = ""
  End If
  ' ... rest of the function ...
End Function