Function GetObjectSafe(progID)
  On Error Resume Next
  Set GetObjectSafe = GetObject(progID)
  If Err.Number <> 0 Then
    Err.Clear
    Set GetObjectSafe = Nothing
  End If
End Function

Sub TestGetObjectSafe()
  Dim objExcel As Object

  Set objExcel = GetObjectSafe(, "Excel.Application")

  If objExcel Is Nothing Then
    MsgBox "Excel is not running or ProgID is incorrect.", vbExclamation
  Else
    MsgBox "Excel is running."
    objExcel.Quit
  End If

  Set objExcel = Nothing
End Sub