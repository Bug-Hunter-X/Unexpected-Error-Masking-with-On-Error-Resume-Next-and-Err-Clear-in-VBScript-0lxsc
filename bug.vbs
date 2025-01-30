Function GetValue(key)
  On Error Resume Next
  SetValue key, "someValue"
  If Err.Number <> 0 Then
    Err.Clear
    result = someMap.Item(key)
  Else
    result = "defaultValue"
  End If
  GetValue = result
End Function