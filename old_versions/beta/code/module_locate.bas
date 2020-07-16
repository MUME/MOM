Attribute VB_Name = "locate"
Option Explicit
Public Sub flipMappingMode()
   compactMode = Not (compactMode)
   If compactMode = True Then
      Call setMapModeOFF
      frmTools.Visible = False
   Else
      If WorldLoaded = False Then Exit Sub
      Call setMapModeON
      frmTools.Visible = True
   End If
End Sub

Public Sub AreaLocate()
   If WorldLoaded = True Then
      If MappingMode = True Then
         compactMode = False
         MappingMode = False
         Call flipMappingMode
      End If
      fleeRadius = 0
      Call caseFleeHandler(currentRoomName, currentExits, 20, True)
   End If
End Sub

Public Sub WorldLocate(ByRef room As String, ByRef data As String)
On Error GoTo errorhandler

   If WorldLoaded = False Then Exit Sub
   If MappingMode = True Then
      MappingMode = False
      compactMode = False
      Call flipMappingMode
   End If
   Out_Of_Sync = True
   GetDescription = False
   fleeMatch = 0
   Call setNewExits(data)  'data represents the Exits:... line
   Dim n As Long, m As Long, row As Long, col As Long
   For n = arrMinRow To arrMaxRow
      For m = arrMinCol To arrMaxCol
         If compareFleeExit(room, n, m) = True Then
            fleeMatch = fleeMatch + 1
            arrTmpFleeStack(fleeMatch, 1) = n
            arrTmpFleeStack(fleeMatch, 2) = m
         Else
            If arr(n, m) < 0 Then
               row = Abs(Fix(arr(n, m) / 1000))
               col = Abs(arr(n, m)) - (row * 1000)
               n = row
               m = col - 1
            End If
         End If
      Next
   Next
   If fleeMatch = 0 Then
      Call SYNC_FALSE
      Exit Sub
   End If
   If fleeMatch = 1 Then
      virtualRow = arrTmpFleeStack(fleeMatch, 1)
      virtualCol = arrTmpFleeStack(fleeMatch, 2)
      Call SYNC_TRUE
      Exit Sub
   End If
   If fleeMatch > 1 Then
      GetDescription = True
      frmMap.tcpClient.SendData "EXAMINE" & vbLf
   End If

Exit Sub
errorhandler:
   errorData = "locate WorldLocate"
   writeError (errorData)
End Sub

