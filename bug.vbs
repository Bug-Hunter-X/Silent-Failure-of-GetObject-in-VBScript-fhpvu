Function GetObject(progID)
  On Error Resume Next
  Set GetObject = GetObject(progID)
  If Err.Number <> 0 Then
    Err.Clear
    Set GetObject = Nothing
  End If
End Function

Sub TestGetObject()
  Dim objExcel As Object

  Set objExcel = GetObject(, "Excel.Application")

  If objExcel Is Nothing Then
    MsgBox "Excel is not running."
  Else
    MsgBox "Excel is running."
    objExcel.Quit
  End If

  Set objExcel = Nothing
End Sub