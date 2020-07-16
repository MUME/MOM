Attribute VB_Name = "ErrorHandling"
Option Explicit
Public Const appending = 8
Public fso As New Scripting.FileSystemObject
Public file
Public errorData As String
Public errorModule As String

Public Sub writeError(data As String)
On Error Resume Next
   LOST = True
   SyncError = True
   frmMap.Caption = mapTitle & " - Error"
   Set file = fso.OpenTextFile(App.Path & "\log.txt", appending, True)
   file.WriteLine ("Logged: " & Date & " - " & data)
   file.Close
   Set file = Nothing
   Call informClient(data)
On Error GoTo 0
End Sub

