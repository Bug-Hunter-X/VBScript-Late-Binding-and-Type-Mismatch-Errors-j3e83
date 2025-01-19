Option Explicit

Dim objFSO As Object

' Early binding: Declare the object variable with its specific type.
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO Is Nothing Then
  Err.Raise vbError, , "Failed to create FileSystemObject"
Else
  ' Type checking: Ensure the object has the expected properties before accessing them.
  If TypeName(objFSO) = "Scripting.FileSystemObject" Then
    If objFSO.FileExists("test.txt") Then
      WScript.Echo "File exists!"
    Else
      WScript.Echo "File does not exist."
    End If
  Else
    Err.Raise vbError, , "Unexpected object type"
  End If
End If

Set objFSO = Nothing