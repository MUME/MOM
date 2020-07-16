Attribute VB_Name = "locate"
Option Explicit
Public locatorCount As Integer
Public locateRetry As Integer

Public Sub WorldLocate(room As String, data As String)
If DEBUGMODE = False Then On Error GoTo errorhandler
   errorData = errorData & "WorldLocate -> "
   If WorldLoaded = False Then Exit Sub
   roomcount = 0
   locatorCount = 0
   LOST = True
   MappingMode = False
   MappingData = False
   frmTools.Hide
   frmMap.mnuEdit.Enabled = True
   frmMap.mnuEdit.Visible = True
   GetDescription = True
   frmMap.tcpPlayer.SendData ("EXAMINE") & vbCrLf
   If wasMapMode Then
      wasMapMode = False
      frmMap.tcpPlayer.SendData ("BRIEF ON") & vbCrLf
      If MUDname = "MUME" Then frmMap.tcpPlayer.SendData ("SPAM OFF") & vbCrLf
   End If
Exit Sub

errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "locate WorldLocate"
   writeError (errorModule)
End Sub
