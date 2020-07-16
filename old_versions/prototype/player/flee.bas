Attribute VB_Name = "flee"
Option Explicit
Option Compare Binary
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
Public arrFleeStack(1 To 441, 1 To 2) As Long
Public fleeRadius As Long
Public Const fleeMaxRadius = 10
Public fleeStackCount As Long
Public fleeMatch As Long
Public currentRoomName As String
Public currentExits As String
Public currentString As String

Public Sub caseFleeHandler(ByRef room As String, ByRef data As String)
   Call resetBuffer
   Call createFleeStack(fleeRadius)
   If getSync(room, data) = True Then
      Call SYNC_TRUE
      Exit Sub
   Else
      Call SYNC_FALSE
      Exit Sub
   End If
End Sub

Public Sub createFleeStack(ByVal wRadius As Long)
   fleeStackCount = 0
   If fleeRadius = 1 And theRoomUp = False And theRoomDown = False Then
      fleeStackCount = fleeStackCount + 1
      arrFleeStack(fleeStackCount, 1) = theRow
      arrFleeStack(fleeStackCount, 2) = theCol
      If theRoomNorth = True Then
         fleeStackCount = fleeStackCount + 1
         arrFleeStack(fleeStackCount, 1) = theRow - 1
         arrFleeStack(fleeStackCount, 2) = theCol
      End If
      If theRoomEast = True Then
         fleeStackCount = fleeStackCount + 1
         arrFleeStack(fleeStackCount, 1) = theRow
         arrFleeStack(fleeStackCount, 2) = theCol + 1
      End If
      If theRoomSouth = True Then
         fleeStackCount = fleeStackCount + 1
         arrFleeStack(fleeStackCount, 1) = theRow + 1
         arrFleeStack(fleeStackCount, 2) = theCol
      End If
      If theRoomWest = True Then
         fleeStackCount = fleeStackCount + 1
         arrFleeStack(fleeStackCount, 1) = theRow
         arrFleeStack(fleeStackCount, 2) = theCol - 1
      End If
   Else
      Dim row As Long, col As Long
      If wRadius > fleeMaxRadius Then wRadius = fleeMaxRadius
      For row = theRow - wRadius To theRow + wRadius
         For col = theCol - wRadius To theCol + wRadius
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
   Dim n
   For n = 1 To fleeStackCount
      If debug_mode = True Then Debug.Print "FLEE STACK: (" & arrFleeStack(n, 1) & "," & arrFleeStack(n, 2) & ")"
   Next
End Sub

Public Function getSync(ByRef room As String, ByRef data As String)
   Call setNewExits(data)  'data represents the Exits:... line
   fleeMatch = 0
   Dim n As Long
   For n = 1 To fleeStackCount
      If debug_mode = True Then Debug.Print
      If debug_mode = True Then Debug.Print n & ". Case __________________________________________________"
      If compareFleeExit(room, arrFleeStack(n, 1), arrFleeStack(n, 2)) = True Then
         If fleeMatch > 1 Then
            If debug_mode = True Then Debug.Print "COMPARE: Multiple matches, FALSE"
            getSync = False
            Exit Function
         Else
            If debug_mode = True Then Debug.Print "COMPARE: exits is TRUE"
            fleeMatch = fleeMatch + 1
            virtualRow = arrFleeStack(n, 1)
            virtualCol = arrFleeStack(n, 2)
         End If
      Else
         If debug_mode = True Then Debug.Print "COMPARE: exits is FALSE"
      End If
   Next
   If fleeMatch = 0 Then
      If debug_mode = True Then Debug.Print "Bad! Out of sync!"
      getSync = False
   Else
      If debug_mode = True Then Debug.Print "Ok! The character and map are in sync!"
      getSync = True
   End If
End Function

Public Function compareFleeExit(ByRef room As String, ByRef row, ByRef col)
   compareFleeExit = False
   currentData = arr(row, col)
   If debug_mode = True Then Debug.Print "LOADING: " & currentData & "    from arr(" & row & "," & col & ")"
   If currentData = 0 Then
      If debug_mode = True Then Debug.Print "------------ ROOM DOES NOT EXIST, skipping -------------"
      Exit Function
   End If
   currentRoom = Split(arrDesc(row, col), ";")
   If debug_mode = True Then Debug.Print "LOADING: " & arrDesc(row, col) & "    from arrDesc(" & row & "," & col & ")"
   If debug_mode = True Then Debug.Print "COMPARE:                  >" & room & "=" & currentRoom(0) & "<"
   If room = currentRoom(0) Then
      If debug_mode = True Then Debug.Print "COMPARE: MATCH!    >" & room & " =" & currentRoom(0) & "<"
   Else
      If debug_mode = True Then Debug.Print "COMPARE: FAILURE!  >" & room & "<>" & currentRoom(0) & "<"
      Exit Function
   End If
   
   If debug_mode = True Then Debug.Print "COMPARE: Exits..."
   If debug_mode = True Then Debug.Print "NewExits: " & " North=" & newNorth & " East=" & newEast & " South=" & newSouth & " West=" & newWest & " Up=" & newUp & " Down=" & newDown
   If currentData <= 0 Then
      If debug_mode = True Then Debug.Print "COMPARE: Room does not exist!"
      Exit Function
   End If
   If (newNorth = True) Then                    'exit north exists
      If ((currentData And N_map) > 0) = True Then 'and on map there is exit north, then Ok!
         If debug_mode = True Then Debug.Print "      : North = True"
      Else
         If debug_mode = True Then Debug.Print "      : North is different"
         Exit Function
      End If
   Else
      If ((currentData And N_map) = 0) = True Then 'exit north didn't exist, and neither on map, then Ok!
         If debug_mode = True Then Debug.Print "      : North = False"
      Else
         If debug_mode = True Then Debug.Print "      : North is different"
         Exit Function
      End If
   End If
   If (newEast = True) Then
      If ((currentData And E_map) > 0) = True Then
         If debug_mode = True Then Debug.Print "      : East = True"
      Else
         If debug_mode = True Then Debug.Print "      : East is different"
         Exit Function
      End If
   Else
      If ((currentData And E_map) = 0) = True Then
         If debug_mode = True Then Debug.Print "      : East = False"
      Else
         If debug_mode = True Then Debug.Print "      : East is different"
         Exit Function
      End If
   End If
   If (newSouth = True) Then
      If ((currentData And S_map) > 0) = True Then
         If debug_mode = True Then Debug.Print "      : South = True"
      Else
         If debug_mode = True Then Debug.Print "      : South is different"
         Exit Function
      End If
   Else
      If ((currentData And S_map) = 0) = True Then
         If debug_mode = True Then Debug.Print "      : South = False"
      Else
         If debug_mode = True Then Debug.Print "      : South is different"
         Exit Function
      End If
   End If
   If (newWest = True) Then
      If ((currentData And W_map) > 0) = True Then
         If debug_mode = True Then Debug.Print "      : West = True"
      Else
         If debug_mode = True Then Debug.Print "      : West is different"
         Exit Function
      End If
   Else
      If ((currentData And W_map) = 0) = True Then
         If debug_mode = True Then Debug.Print "      : South = False"
      Else
         If debug_mode = True Then Debug.Print "      : West is different"
         Exit Function
      End If
   End If
   If (newUp = True) Then
      If ((currentData And U_map) > 0) = True Then
         If debug_mode = True Then Debug.Print "      : Up = True"
      Else
         If debug_mode = True Then Debug.Print "      : Up is different"
         Exit Function
      End If
   Else
      If ((currentData And U_map) = 0) = True Then
         If debug_mode = True Then Debug.Print "      : Up = False"
      Else
         If debug_mode = True Then Debug.Print "      : Up is different"
         Exit Function
      End If
   End If
   If (newDown = True) Then
      If ((currentData And D_map) > 0) = True Then
         If debug_mode = True Then Debug.Print "      : Down = True"
      Else
         If debug_mode = True Then Debug.Print "      : Down is different"
         Exit Function
      End If
   Else
      If ((currentData And D_map) = 0) = True Then
         If debug_mode = True Then Debug.Print "      : Down = False"
      Else
         If debug_mode = True Then Debug.Print "      : Down is different"
         Exit Function
      End If
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
   Out_Of_Sync = False
   fleeRadius = 1
   theRow = virtualRow
   theCol = virtualCol
   BestEST.status.ForeColor = &HC0FFC0
   BestEST.status.Caption = "Ok."
   Call LoadRoom(theRow, theCol)
   Call DrawMap
End Sub
Public Sub SYNC_FALSE()
   Out_Of_Sync = True
   fleeRadius = fleeRadius + 1
   BestEST.status.ForeColor = &HFF&
   BestEST.status.Caption = "Out of Sync!"
End Sub

Public Function checkArrayLimit(ByRef wRow As Long, ByRef wCol As Long)
   If wRow < arrMinRow Then
      checkArrayLimit = False
      Exit Function
   End If
   If wCol < arrMinCol Then
      checkArrayLimit = False
      Exit Function
   End If
   If wRow > arrMaxRow Then
      checkArrayLimit = False
      Exit Function
   End If
   If wCol > arrMaxCol Then
      checkArrayLimit = False
      Exit Function
   End If
   checkArrayLimit = True
End Function
