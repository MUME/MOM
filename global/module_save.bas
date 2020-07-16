Attribute VB_Name = "save"
Option Explicit
Private row As Long
Private col As Long

Public oldRow As Integer
Public oldCol As Integer

Private warp As Boolean

Public Sub saveWorld()
errorData = errorData & "saveWorld -> "
If DEBUGMODE = False Then On Error GoTo errorhandler
   Dim cursor As Integer
   
   On Error Resume Next
   Call fso.DeleteFile(filePath & ".bak", True)
   Call fso.MoveFile(filePath, filePath & ".bak")
   On Error GoTo errorhandler
   
   Open filePath For Output As #1
   For cursor = 1 To theCount
      If LenB(aData(cursor, cDATA)) <> 0 Then
         Print #1, aData(cursor, cENCRYPTED)
      End If
   Next
   Close #1
Exit Sub

errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "save SaveWorld"
   writeError (errorModule)
End Sub

'      If aData(cursor, cROW) > 35 And aData(cursor, cROW) < 48 And aData(cursor, cCOL) > 431 And aData(cursor, cCOL) < 443 Then
'         If LenB(aData(cursor, cDATA)) <> 0 Then Print #1, aData(cursor, cENCRYPTED)
'      End If
