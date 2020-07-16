Attribute VB_Name = "flee"
Option Explicit
Option Compare Binary
Public GetDescription As Boolean
Public arrRoomStack(1 To 30, 1 To 2) As Long
Public arrMoveStack(1 To 30) As String
Public arrTmpRoomStack(1 To 30, 1 To 2) As Long
Public arrTmpMoveStack(1 To 30) As String
Public Get_In_Sync As Boolean
Public Out_Of_Sync As Boolean
Public newNorth As Boolean
Public newEast As Boolean
Public newSouth As Boolean
Public newWest As Boolean
Public newUp As Boolean
Public newDown As Boolean
Public arrFleeStack(1 To 500, 1 To 2) As Long      ' Depends of FleeMaxRadius
Public arrTmpFleeStack(1 To 500, 1 To 2) As Long   ' Depends of FleeMaxRadius
Public fleeRadius As Long
Public fleeMaxRadius
Public fleeStackCount As Long
Public fleeMatch As Long
Public currentRoomName As String
Public tmpDesc
Public tmpDescMap As String * 16
Public tmpDescRoom As String * 16
Public currentExits As String
Public currentString As String
Public currentDesc As String
Public fleeSpecialCase As Boolean

Public Sub caseFleeHandler(ByRef room As String, ByRef data As String, ByVal radius, checkDesc As Boolean)
On Error GoTo errorhandler

   GetDescription = False
   fleeMatch = 0
   Call resetBuffer
Retry:
   If fleeRadius >= radius Then Exit Sub
   fleeRadius = fleeRadius + 1
   Call createFleeStack(fleeRadius)
   If getSync(room, data) = True Then
      Call SYNC_TRUE
   Else
      If fleeRadius = radius Then
         Call SYNC_FALSE
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
            If checkDesc = True Then
               GetDescription = True
               frmMap.tcpClient.SendData "EXAMINE" & vbLf
            Else
               virtualRow = arrTmpFleeStack(fleeMatch, 1)
               virtualCol = arrTmpFleeStack(fleeMatch, 2)
               Call SYNC_TRUE
            End If
         End If
         Exit Sub
      End If
      GoTo Retry
   End If
Exit Sub
errorhandler:
   errorData = "flee caseFleeHandler"
   writeError (errorData)
End Sub

Public Sub cmpFleeDesc(ByRef Description As String)
On Error GoTo errorhandler
   Dim n As Long
   tmpDescRoom = EncryptDesc(Description)
   For n = 1 To fleeMatch
      tmpDescMap = arrDescription(arrTmpFleeStack(n, 1), arrTmpFleeStack(n, 2))
      If tmpDescMap = tmpDescRoom Then
         virtualRow = arrTmpFleeStack(n, 1)
         virtualCol = arrTmpFleeStack(n, 2)
         Call SYNC_TRUE
         Exit Sub
      End If
   Next
   Call SYNC_FALSE
   frmTools.status = "Unable to locate character!"

Exit Sub
errorhandler:
   errorData = "flee cmpFleeDesc"
   writeError (errorData)
End Sub

Public Sub createFleeStack(ByVal wRadius As Long)
   fleeStackCount = 0   'theExitUp = False And theExitDown = False Then
   If Out_Of_Sync = False And _
      fleeRadius = 1 And _
      thePortalNorth = False And theDoorPortalNorth = False And _
      thePortalEast = False And theDoorPortalEast = False And _
      thePortalSouth = False And theDoorPortalSouth = False And _
      thePortalWest = False And theDoorPortalWest = False And _
      thePortalUp = False And theDoorPortalUp = False And _
      thePortalDown = False And theDoorPortalDown = False Then

      fleeSpecialCase = True
      fleeStackCount = fleeStackCount + 1
      arrFleeStack(fleeStackCount, 1) = theRow
      arrFleeStack(fleeStackCount, 2) = theCol
      
      If theExitNorth Then
         fleeStackCount = fleeStackCount + 1
         arrFleeStack(fleeStackCount, 1) = theRow - 1
         arrFleeStack(fleeStackCount, 2) = theCol
      End If
      If theExitEast Then
         fleeStackCount = fleeStackCount + 1
         arrFleeStack(fleeStackCount, 1) = theRow
         arrFleeStack(fleeStackCount, 2) = theCol + 1
      End If
      If theExitSouth Then
         fleeStackCount = fleeStackCount + 1
         arrFleeStack(fleeStackCount, 1) = theRow + 1
         arrFleeStack(fleeStackCount, 2) = theCol
      End If
      If theExitWest Then
         fleeStackCount = fleeStackCount + 1
         arrFleeStack(fleeStackCount, 1) = theRow
         arrFleeStack(fleeStackCount, 2) = theCol - 1
      End If
   Else
      fleeSpecialCase = False
      Dim row As Long, col As Long, withStep As Long
      For row = theRow - wRadius To theRow + wRadius
         If (fleeRadius = 1) Or (row = theRow - wRadius) Or (row = theRow + wRadius) Then
            withStep = 1
         Else
            withStep = wRadius + wRadius
         End If
         For col = theCol - wRadius To theCol + wRadius Step withStep
            If checkArrayLimit(row, col) = True Then
               If arr(row, col) > 0 Then
                  fleeStackCount = fleeStackCount + 1
                  arrFleeStack(fleeStackCount, 1) = row
                  arrFleeStack(fleeStackCount, 2) = col
               End If
            End If
         Next
      Next
   End If
'   Dim n
  ' For n = 1 To fleeStackCount
  '    If DEBUG_MODE = True Then Debug.Print "FLEE STACK: (" & arrFleeStack(n, 1) & "," & arrFleeStack(n, 2) & ")"
 '  Next
End Sub

Public Function getSync(ByRef room As String, ByRef exitsData As String)
   getSync = False
   Call setNewExits(exitsData)   'data represents the Exits:... line
   Dim n As Long
   For n = 1 To fleeStackCount
      If compareFleeExit(room, arrFleeStack(n, 1), arrFleeStack(n, 2)) = True Then
         fleeMatch = fleeMatch + 1
         arrTmpFleeStack(fleeMatch, 1) = arrFleeStack(n, 1)
         arrTmpFleeStack(fleeMatch, 2) = arrFleeStack(n, 2)
      End If
   Next
   If fleeSpecialCase = True And fleeMatch = 1 Then
      virtualRow = arrTmpFleeStack(fleeMatch, 1)
      virtualCol = arrTmpFleeStack(fleeMatch, 2)
      getSync = True
      Exit Function
   End If
End Function

Public Function compareFleeExit(ByRef room As String, ByRef row, ByRef col)
   compareFleeExit = False
   
   currentData = arr(row, col)
 '  If DEBUG_MODE = True Then Debug.Print "NewExits: " & " North=" & newNorth & " East=" & newEast & " South=" & newSouth & " West=" & newWest & " Up=" & newUp & " Down=" & newDown
   If currentData <= 0 Then Exit Function
   If (newNorth = True) Then                          'exit north exists
      If (currentData And N_MAP) > 0 Then    'Room has exit NORTH And MAP has NORTH, then Ok!
  '       If DEBUG_MODE = True Then Debug.Print "      : North = True"
      Else
   '      If DEBUG_MODE = True Then Debug.Print "      : North is different"
         Exit Function
      End If
   End If
   If (newEast = True) Then
      If (currentData And E_MAP) > 0 Then
   '      If DEBUG_MODE = True Then Debug.Print "      : East = True"
      Else
   '      If DEBUG_MODE = True Then Debug.Print "      : East is different"
         Exit Function
      End If
   End If
   If (newSouth = True) Then
      If (currentData And S_MAP) > 0 Then
    '     If DEBUG_MODE = True Then Debug.Print "      : South = True"
      Else
   '      If DEBUG_MODE = True Then Debug.Print "      : South is different"
         Exit Function
      End If
   End If
   If (newWest = True) Then
      If (currentData And W_MAP) > 0 Then
   '      If DEBUG_MODE = True Then Debug.Print "      : West = True"
      Else
    '     If DEBUG_MODE = True Then Debug.Print "      : West is different"
         Exit Function
      End If
   End If
   If (newUp = True) Then
      If (currentData And U_MAP) > 0 Then
    '     If DEBUG_MODE = True Then Debug.Print "      : Up = True"
      Else
   '      If DEBUG_MODE = True Then Debug.Print "      : Up is different"
         Exit Function
      End If
   End If
   If (newDown = True) Then
      If (currentData And D_MAP) > 0 Then
    '     If DEBUG_MODE = True Then Debug.Print "      : Down = True"
      Else
    '     If DEBUG_MODE = True Then Debug.Print "      : Down is different"
         Exit Function
      End If
   End If
   
   If False Then
      'Public Const N_MAP = 96
      'Public Const N_noexit = 0
      'Public Const N_exit = 32
      'Public Const N_door = 64
      'Public Const N_portal = 96
      If (newNorth = False) Then
         If (currentData And N_MAP) = 0 Or _
            (currentData And N_door) > 0 Or _
            (currentData And N_portal) > 0 Then
         Else
            Exit Function
         End If
      End If
      If (newEast = False) Then
         If (currentData And E_MAP) = 0 Or _
            (currentData And E_door) > 0 Or _
            (currentData And E_portal) > 0 Then
         Else
            Exit Function
         End If
      End If
      If (newSouth = False) Then
         If (currentData And S_MAP) = 0 Or _
            (currentData And S_door) > 0 Or _
            (currentData And S_portal) > 0 Then
         Else
            Exit Function
         End If
      End If
      If (newWest = False) Then
         If (currentData And W_MAP) = 0 Or _
            (currentData And W_door) > 0 Or _
            (currentData And W_portal) > 0 Then
         Else
            Exit Function
         End If
      End If
      If (newUp = False) Then
         If (currentData And U_MAP) = 0 Or _
            (currentData And U_door) > 0 Or _
            (currentData And U_portal) > 0 Then
         Else
            Exit Function
         End If
      End If
      If (newDown = False) Then
         If (currentData And D_MAP) = 0 Or _
            (currentData And D_door) > 0 Or _
            (currentData And D_portal) > 0 Then
         Else
            Exit Function
         End If
      End If
   End If

   currentRoom = arrRoomname(row, col)
'   If DEBUG_MODE = True Then Debug.Print "LOADING: " & arrDesc(row, col) & "    from arrDesc(" & row & "," & col & ")"
 '  If DEBUG_MODE = True Then Debug.Print "COMPARE:             " & room & "  ===  " & currentRoom(0) & "   "
   If Len(room) = Len(currentRoom) And room = currentRoom Then
 '     If DEBUG_MODE = True Then Debug.Print "COMPARE: MATCH!    >" & room & " =" & currentRoom(0) & "<"
   Else
  '    If DEBUG_MODE = True Then Debug.Print "COMPARE: FAILURE!  >" & room & "<>" & currentRoom(0) & "<"
      Exit Function
   End If
   
   compareFleeExit = True
End Function

Public Sub setNewExits(ByRef data As String)
   If checkString(data, "north") = True Then
      newNorth = True
   Else
      newNorth = False
   End If
   If checkString(data, "east") = True Then
      newEast = True
   Else
      newEast = False
   End If
   If checkString(data, "south") = True Then
      newSouth = True
   Else
      newSouth = False
   End If
   If checkString(data, "west") = True Then
      newWest = True
   Else
      newWest = False
   End If
   If checkString(data, "up") = True Then
      newUp = True
   Else
      newUp = False
   End If
   If checkString(data, "down") = True Then
      newDown = True
   Else
      newDown = False
   End If
End Sub

Public Sub SYNC_TRUE()
   fleeRadius = 0
   Out_Of_Sync = False
   theRow = virtualRow
   theCol = virtualCol
   frmMap.Caption = "Synchronization is successful!"
   Call loadRoom(theRow, theCol)
   Call DrawMap
End Sub
Public Sub SYNC_FALSE()
   frmMap.Circle (150, 150), 60, QBColor(12)
   frmMap.Circle (150, 150), 59, QBColor(12)
   Out_Of_Sync = True
   frmMap.Caption = "Unable to locate character!"
End Sub

Public Function checkArrayLimit(ByRef wRow As Long, ByRef wCol As Long)
   checkArrayLimit = True
   If wRow < arrMinRow Then checkArrayLimit = False
   If wCol < arrMinCol Then checkArrayLimit = False
   If wRow > arrMaxRow Then checkArrayLimit = False
   If wCol > arrMaxCol Then checkArrayLimit = False
End Function
