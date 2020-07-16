Attribute VB_Name = "movement"
Option Explicit
Option Compare Binary
Public Const limit = 2
Public theRow  As Long
Public theCol As Long
Public tmpRow As Long
Public tmpCol As Long
Public virtualRow As Long
Public virtualCol As Long
Public roomCount As Long
Public newRoomCount As Long
Public currentDir As String
Public tmpMove As String
Public currentData As Long
Public currentRoom
Public theMove As String
Public AlasCount As Long

Public Function checkString(ByVal data, ByVal search)
   If InStr(1, data, search) > 0 Then
      checkString = True
   Else
      checkString = False
   End If
End Function

Public Function checkTheMap(ByRef theCount As Long, _
                              ByRef theMap As Long, _
                              ByVal row As Long, ByVal col As Long, _
                              ByRef data As String)
On Error GoTo errorhandler

'If DEBUG_MODE = True Then Debug.Print "______________________Remove_Room_Move"
'Dim n
'For n = 1 To theCount
'   If DEBUG_MODE = True Then Debug.Print n & ". |" & arrRoomStack(n, 1) & "," & arrRoomStack(n, 2) & "     |" & arrMoveStack(n) & "|"
'Next

   If (arr(virtualRow, virtualCol) And theMap) > 0 And (theCount < arrMaxRoom) Then    'there is an Exit
      checkTheMap = False
      currentData = (arr(virtualRow, virtualCol) And theMap)
      
      Select Case currentData
      Case N_exit, E_exit, S_exit, W_exit
         checkTheMap = True
         If addRoom(theCount, row, col, data) = False Then checkTheMap = False
      
      Case N_door, E_door, S_door, W_door
         checkTheMap = True
         If addRoom(theCount, row, col, data) = False Then checkTheMap = False
       
      Case U_exit, U_door, U_portal, U_doorportal
         currentRoom = Split(arrDesc(virtualRow, virtualCol), ";")
         tmpRow = currentRoom(14)
         tmpCol = currentRoom(15)
         checkTheMap = True
         If addRoom(theCount, tmpRow, tmpCol, data) = False Then checkTheMap = False
       
      Case D_exit, D_door, D_portal, D_doorportal
         currentRoom = Split(arrDesc(virtualRow, virtualCol), ";")
         tmpRow = currentRoom(17)
         tmpCol = currentRoom(18)
         checkTheMap = True
         If addRoom(theCount, tmpRow, tmpCol, data) = False Then checkTheMap = False
      
      Case N_portal, N_doorportal
         currentRoom = Split(arrDesc(virtualRow, virtualCol), ";")
         tmpRow = currentRoom(2)
         tmpCol = currentRoom(3)
         checkTheMap = True
         If addRoom(theCount, tmpRow, tmpCol, data) = False Then checkTheMap = False
       
      Case E_portal, E_doorportal
         currentRoom = Split(arrDesc(virtualRow, virtualCol), ";")
         tmpRow = currentRoom(5)
         tmpCol = currentRoom(6)
         checkTheMap = True
         If addRoom(theCount, tmpRow, tmpCol, data) = False Then checkTheMap = False
       
      Case S_portal, S_doorportal
         currentRoom = Split(arrDesc(virtualRow, virtualCol), ";")
         tmpRow = currentRoom(8)
         tmpCol = currentRoom(9)
         checkTheMap = True
         If addRoom(theCount, tmpRow, tmpCol, data) = False Then checkTheMap = False
       
      Case W_portal, W_doorportal
         currentRoom = Split(arrDesc(virtualRow, virtualCol), ";")
         tmpRow = currentRoom(11)
         tmpCol = currentRoom(12)
         checkTheMap = True
         If addRoom(theCount, tmpRow, tmpCol, data) = False Then checkTheMap = False
      End Select
   End If

Exit Function
errorhandler:
   errorData = "movement checkTheMap"
   writeError (errorData)
   checkTheMap = True
   Out_Of_Sync = True
End Function

Public Sub updateTheRoom()
On Error GoTo errorhandler
   If arrRoomStack(1, 1) > 0 And arrRoomStack(1, 2) > 0 Then
      If roomCount >= limit Then frmMap.tcpClient.SendData arrMoveStack(limit) & vbLf  '¤¤¤¤¤¤¤
      theRow = arrRoomStack(1, 1)
      theCol = arrRoomStack(1, 2)
      Call removeRoom
      Call loadRoom(theRow, theCol)
      If SyncError = True Then
         Call SYNC_FALSE
         Exit Sub
      End If
      Call DrawMap
      'Call DrawVirtualMoves
   End If

Exit Sub
errorhandler:
   errorData = "movement updateTheRoom"
   writeError (errorData)
End Sub

Public Sub removeRoom()
   If roomCount > 0 Then
      Dim n
      roomCount = roomCount - 1
      For n = arrMinRoom To roomCount
         arrRoomStack(n, 1) = arrRoomStack(n + 1, 1)
         arrRoomStack(n, 2) = arrRoomStack(n + 1, 2)
         arrMoveStack(n) = arrMoveStack(n + 1)
      Next
      arrRoomStack(roomCount + 1, 1) = 0
      arrRoomStack(roomCount + 1, 2) = 0
      arrMoveStack(roomCount + 1) = ""

'If DEBUG_MODE = True Then Debug.Print "______________________Remove_Room_Move"
'For n = 1 To roomCount
'   If DEBUG_MODE = True Then Debug.Print n & ". |" & arrRoomStack(n, 1) & "," & arrRoomStack(n, 2) & "     |" & arrMoveStack(n) & "|"
'Next
   
   End If
End Sub

Public Function addRoom(ByRef theCount As Long, ByVal row As Long, ByVal col As Long, ByRef data As String)
   If theCount < arrMaxRoom Then
      theCount = theCount + 1
      arrRoomStack(theCount, 1) = row
      arrRoomStack(theCount, 2) = col
      arrMoveStack(theCount) = data    '& vbLf
      virtualRow = row
      virtualCol = col
      addRoom = True

'Dim n
'If DEBUG_MODE = True Then Debug.Print "_______________________Add_Room_Move"
'For n = 1 To theCount
'   If DEBUG_MODE = True Then Debug.Print n & ". |" & arrRoomStack(n, 1) & "," & arrRoomStack(n, 2) & "     |" & arrMoveStack(n) & "|"
'Next

   Else
      addRoom = False
   End If
End Function

Public Sub Collision()
   Dim n As Long, m As Long, k As Long
On Error GoTo errorhandler
   
'If DEBUG_MODE = True Then Debug.Print "RoomCount=" & roomCount & "_________________Collision_Start_______________________"
'For n = 1 To roomCount
'   If DEBUG_MODE = True Then Debug.Print n & ". " & arrMoveStack(n) & "|" & arrRoomStack(n, 1) & "," & arrRoomStack(n, 2) & "|"
'Next
'If DEBUG_MODE = True Then Debug.Print "Current Row=" & theRow & "___ col=" & theCol

   If limit = 2 And roomCount Then
      roomCount = 0
      virtualRow = theRow
      virtualCol = theCol
      Exit Sub
   End If
   If roomCount > limit Then roomCount = limit  '¤¤¤¤¤¤¤¤¤¤¤¤¤
   tmpMove = arrMoveStack(arrMinRoom)
   Call removeRoom
   virtualRow = theRow
   virtualCol = theCol
   newRoomCount = 0
   AlasCount = 0
   
   For k = arrMinRoom To roomCount
      If tmpMove <> arrMoveStack(k) Then Exit For
   Next
   
   For n = k To roomCount
      If checkTheMove(newRoomCount, arrMoveStack(n)) = True Then
      Else
         If n = k Then AlasCount = 1
         If n < roomCount Then
            For m = n To roomCount - 1
               If arrMoveStack(m) <> arrMoveStack(m + 1) Then
                  n = m
                  Exit For
               End If
               If n = k Then AlasCount = AlasCount + 1
            Next
         End If
      End If
   Next
   
   If newRoomCount > 0 Then
      If roomCount >= limit Then frmMap.tcpClient.SendData arrMoveStack(limit) & vbLf  '¤¤¤¤¤¤¤¤¤¤¤¤¤
      theRow = arrRoomStack(arrMinRoom, 1)
      theCol = arrRoomStack(arrMinRoom, 2)
      virtualRow = arrRoomStack(roomCount, 1)
      virtualCol = arrRoomStack(roomCount, 2)
   End If
   roomCount = newRoomCount

'If DEBUG_MODE = True Then Debug.Print "AlasCount=" & AlasCount
'If DEBUG_MODE = True Then Debug.Print "RoomCount=" & roomCount & "_________________Collision_End_______________________"
'   For n = 1 To roomCount
'      If DEBUG_MODE = True Then Debug.Print n & ". " & arrMoveStack(n) & "|" & arrRoomStack(n, 1) & "," & arrRoomStack(n, 2) & "|"
'   Next
'If DEBUG_MODE = True Then Debug.Print "Current Row=" & theRow & "___ col=" & theCol

Exit Sub
errorhandler:
   errorData = "movement Collision"
   writeError (errorData)
End Sub

Public Function checkTheMove(ByRef theCount As Long, ByRef command As String)
   checkTheMove = False
   Select Case command
   Case "n"
      If checkTheMap(theCount, N_MAP, virtualRow - 1, virtualCol, "n") = True Then
         checkTheMove = True
      End If
   Case "e"
      If checkTheMap(theCount, E_MAP, virtualRow, virtualCol + 1, "e") = True Then
         checkTheMove = True
      End If
   Case "s"
      If checkTheMap(theCount, S_MAP, virtualRow + 1, virtualCol, "s") = True Then
         checkTheMove = True
      End If
   Case "w"
      If checkTheMap(theCount, W_MAP, virtualRow, virtualCol - 1, "w") = True Then
         checkTheMove = True
      End If
   Case "u"
      If checkTheMap(theCount, U_MAP, virtualRow, virtualCol, "u") = True Then
         checkTheMove = True
      End If
   Case "d"
      If checkTheMap(theCount, D_MAP, virtualRow, virtualCol, "d") = True Then
         checkTheMove = True
      End If
   End Select
End Function

Public Sub resetBuffer()
   roomCount = 0
   Erase arrMoveStack
   Erase arrRoomStack
   virtualRow = theRow
   virtualCol = theCol
End Sub

Public Sub cancelBuffer()
   If roomCount > 0 Then
      roomCount = limit
      virtualRow = arrRoomStack(limit, 1)
      virtualCol = arrRoomStack(limit, 2)
   End If
End Sub

