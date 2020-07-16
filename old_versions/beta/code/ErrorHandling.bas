Attribute VB_Name = "ErrorHandling"
Public Const appending = 8
Public fso
Public file
Public errorData As String

Public Sub initError()
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set file = fso.OpenTextFile(App.Path & "\log.txt", appending, True)
End Sub

Public Sub writeError(data As String)

   file.WriteLine (Date & " - " & Err.Description & " : " & Err.Number & " in " & data)
   file.Close

End Sub
