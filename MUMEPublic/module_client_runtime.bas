Attribute VB_Name = "CLIENT_RUNTIME"
Option Explicit
Private a As Integer
Private b As String
Private c As String
Public stackOUT As Integer
Public stackIN As Integer
Public coll As Integer
Public atCrLf As Integer
Public atLf As Integer
Public specialLen As Integer
Public surfing As Boolean
Public lookfor As String
Public targetFound As Boolean

Public Function handleSpecial(strData As String)
errorData = errorData & "handleSpecial -> "
   handleSpecial = False
   If Len(strData) = specialLen Then
      frmMap.tcpPlayer.SendData strData
      handleSpecial = True
      Exit Function
   End If
End Function

Public Function handleRuntimeCommand(ByVal strData As String)
If DEBUGMODE = False Then On Error GoTo errorhandler
errorData = errorData & "handleRuntimeCommand -> "
   handleRuntimeCommand = False
   a = Len(strData)
   b = Mid(strData, 1, 1)
   If b = "_" Then
      c = Mid(strData, 1, 5)
      Select Case c
'      Case "_walk"
'         Call frmMap.mnuWalk_Click
'      Case "_here"
'         Call frmMap.mnuHere_Click
'      Case "_lead"
'         frmMap.mnuFollow.Checked = True
'         If Len(strData) > 8 Then leader = Mid(strData, 7, Len(strData) - 7) Else leader = ""
'         Call informClient("Leader: >" & leader & "<")
'         handleRuntimeCommand = True: Exit Function
'      Case "_undo"
'         If canUndo Then
'            Dim data As Long
'            data = arrData(arrWorld(virtualRow, virtualCol), cDATA)
'
'            If (data And N_exit) Then  'NORTH
'               If (data And N_portal) = N_portal Or (data And N_doorportal) = N_doorportal Or (data And (N_hiddendoor Or N_portal)) = (N_hiddendoor Or N_portal) Then
'                  If (arrData(arrWorld(virtualRow, virtualCol), cNPORTALR)) = undoRow And _
'                     (arrData(arrWorld(virtualRow, virtualCol), cNPORTALC)) = undoCol Then
'                        b = "n": a = 2: strData = "n" & vbCrLf
'                  End If
'               Else
'                  If (arrData(arrWorld(virtualRow - 1, virtualCol), cDATA)) > 0 Then
'                     If (virtualRow - 1) = undoRow Then b = "n": a = 2: strData = "n" & vbCrLf
'                  End If
'               End If
'            End If
'
'            If (data And S_exit) Then  'SOUTH
'               If (data And S_portal) = S_portal Or (data And S_doorportal) = S_doorportal Or (data And (S_hiddendoor Or S_portal)) = (S_hiddendoor Or S_portal) Then
'                  If (arrData(arrWorld(virtualRow, virtualCol), cSPORTALR)) = undoRow And _
'                     (arrData(arrWorld(virtualRow, virtualCol), cSPORTALC)) = undoCol Then
'                        b = "s": a = 2: strData = "s" & vbCrLf
'                  End If
'               Else
'                  If (arrData(arrWorld(virtualRow + 1, virtualCol), cDATA)) > 0 Then
'                     If (virtualRow + 1) = undoRow Then b = "s": a = 2: strData = "s" & vbCrLf
'                  End If
'               End If
'            End If
'
'            If (data And E_exit) Then  'EAST
'               If (data And E_portal) = E_portal Or (data And E_doorportal) = E_doorportal Or (data And (E_hiddendoor Or E_portal)) = (E_hiddendoor Or E_portal) Then
'                  If (arrData(arrWorld(virtualRow, virtualCol), cEPORTALR)) = undoRow And _
'                     (arrData(arrWorld(virtualRow, virtualCol), cEPORTALC)) = undoCol Then
'                        b = "e": a = 2: strData = "e" & vbCrLf
'                  End If
'               Else
'                  If (arrData(arrWorld(virtualRow, virtualCol + 1), cDATA)) > 0 Then
'                     If (virtualCol + 1) = undoCol Then b = "e": a = 2: strData = "e" & vbCrLf
'                  End If
'               End If
'            End If
'
'            If (data And W_exit) Then  'WEST
'               If (data And W_portal) = W_portal Or (data And W_doorportal) = W_doorportal Or (data And (W_hiddendoor Or W_portal)) = (W_hiddendoor Or W_portal) Then
'                  If (arrData(arrWorld(virtualRow, virtualCol), cWPORTALR)) = undoRow And _
'                     (arrData(arrWorld(virtualRow, virtualCol), cWPORTALC)) = undoCol Then
'                        b = "w": a = 2: strData = "w" & vbCrLf
'                  End If
'               Else
'                  If (arrData(arrWorld(virtualRow, virtualCol - 1), cDATA)) > 0 Then
'                     If (virtualCol - 1) = undoCol Then b = "w": a = 2: strData = "w" & vbCrLf
'                  End If
'               End If
'            End If
'
'            If (data And W_exit) Then  'WEST
'               If (data And W_portal) = W_portal Or (data And W_doorportal) = W_doorportal Or (data And (W_hiddendoor Or W_portal)) = (W_hiddendoor Or W_portal) Then
'                  If (arrData(arrWorld(virtualRow, virtualCol), cWPORTALR)) = undoRow And _
'                     (arrData(arrWorld(virtualRow, virtualCol), cWPORTALC)) = undoCol Then
'                        b = "w": a = 2: strData = "w" & vbCrLf
'                  End If
'               Else
'                  If (arrData(arrWorld(virtualRow, virtualCol - 1), cDATA)) > 0 Then
'                     If (virtualCol - 1) = undoCol Then b = "w": a = 2: strData = "w" & vbCrLf
'                  End If
'               End If
'            End If
'            If (data And U_exit) Then  'UP
'               If (arrData(arrWorld(virtualRow, virtualCol), cUPORTALR)) = undoRow And _
'                  (arrData(arrWorld(virtualRow, virtualCol), cUPORTALC)) = undoCol Then
'                     b = "u": a = 2: strData = "u" & vbCrLf
'               End If
'            End If
'            If (data And D_exit) Then  'DOWN
'               If (arrData(arrWorld(virtualRow, virtualCol), cDPORTALR)) = undoRow And _
'                  (arrData(arrWorld(virtualRow, virtualCol), cDPORTALC)) = undoCol Then
'                     b = "d": a = 2: strData = "d" & vbCrLf
'               End If
'            End If
'         Else
'            strData = ""
'         End If
'      Case "_obey"
'         strData = Replace(strData, vbLf, "", , , vbTextCompare)
'         strData = Replace(strData, vbCr, "", , , vbTextCompare)
'         strData = Mid(strData, 7)
'         If InStr(1, strData, "[n]", vbTextCompare) > 0 Then
'            If Len(arrData(arrWorld(virtualRow, virtualCol), cNDOOR)) > 0 Then
'               strData = Replace(strData, "[n]", arrData(arrWorld(virtualRow, virtualCol), cNDOOR), , , vbTextCompare)
'            Else
'               strData = Replace(strData, "[n]", "exit north", , , vbTextCompare)
'            End If
'         End If
'         If InStr(1, strData, "[e]", vbTextCompare) > 0 Then
'            If Len(arrData(arrWorld(virtualRow, virtualCol), cEDOOR)) > 0 Then
'               strData = Replace(strData, "[e]", arrData(arrWorld(virtualRow, virtualCol), cEDOOR), , , vbTextCompare)
'            Else
'               strData = Replace(strData, "[e]", "exit east", , , vbTextCompare)
'            End If
'         End If
'         If InStr(1, strData, "[s]", vbTextCompare) > 0 Then
'            If Len(arrData(arrWorld(virtualRow, virtualCol), cSDOOR)) > 0 Then
'               strData = Replace(strData, "[s]", arrData(arrWorld(virtualRow, virtualCol), cSDOOR), , , vbTextCompare)
'            Else
'               strData = Replace(strData, "[s]", "exit south", , , vbTextCompare)
'            End If
'         End If
'         If InStr(1, strData, "[w]", vbTextCompare) > 0 Then
'            If Len(arrData(arrWorld(virtualRow, virtualCol), cWDOOR)) > 0 Then
'               strData = Replace(strData, "[w]", arrData(arrWorld(virtualRow, virtualCol), cWDOOR), , , vbTextCompare)
'            Else
'               strData = Replace(strData, "[w]", "exit west", , , vbTextCompare)
'            End If
'         End If
'         If InStr(1, strData, "[u]", vbTextCompare) > 0 Then
'            If Len(arrData(arrWorld(virtualRow, virtualCol), cUDOOR)) > 0 Then
'               strData = Replace(strData, "[u]", arrData(arrWorld(virtualRow, virtualCol), cUDOOR), , , vbTextCompare)
'            Else
'               strData = Replace(strData, "[u]", "exit up", , , vbTextCompare)
'            End If
'         End If
'         If InStr(1, strData, "[d]", vbTextCompare) > 0 Then
'            If Len(arrData(arrWorld(virtualRow, virtualCol), cDDOOR)) > 0 Then
'               strData = Replace(strData, "[d]", arrData(arrWorld(virtualRow, virtualCol), cDDOOR), , , vbTextCompare)
'            Else
'               strData = Replace(strData, "[d]", "exit down", , , vbTextCompare)
'            End If
'         End If
'         frmMap.tcpPlayer.SendData Replace(strData, "@", "", , , vbTextCompare) & vbCrLf
'         Call informClient(lookHeader & WHITE & ";" & BOLD & lookFooter & strData & colourEndCode, True)
'         'lookheader & WHITE & ";" & BOLD & lookfooter
'         handleRuntimeCommand = True: Exit Function
'GROUP FEATURE
'      Case "_grpt"
'         If frmMap.mnuGroup.Checked Then Call groupTell(strData)
'         handleRuntimeCommand = True: Exit Function
      Case "_hide"
         If frmMap.WindowState = vbMinimized Then
            frmMap.WindowState = vbNormal
            Call DrawMap
         Else
            frmMap.WindowState = vbMinimized
         End If
         handleRuntimeCommand = True: Exit Function
      Case "_free"
         theExitNorth = True
         theExitEast = True
         theExitSouth = True
         theExitWest = True
         theExitUp = True
         theExitDown = True
         handleRuntimeCommand = True: Exit Function
      Case "_canc"
         Call cancelBuffer
         handleRuntimeCommand = True: Exit Function
      Case "_show"
         frmMap.Hide
         handleRuntimeCommand = True: Exit Function
      Case "_sync"
         Call frmMap.mnuLocate_Click
         handleRuntimeCommand = True: Exit Function
      Case "_tool"
         Call frmMap.mnuTools_Click
         handleRuntimeCommand = True: Exit Function
      Case "_map1"
         Call frmMap.mnuMap1_Click
         handleRuntimeCommand = True: Exit Function
      Case "_map2"
         Call frmMap.mnuMap2_Click
         handleRuntimeCommand = True: Exit Function
      Case "_map3"
         Call frmMap.mnumap3_Click
         handleRuntimeCommand = True: Exit Function
      Case "_door"
         Call frmMap.mnuDoornames_Click
        handleRuntimeCommand = True: Exit Function
      Case "_find"
         lookfor = Trim(LCase(Mid(strData, 7, a - 7)))
         Call DrawMap
         handleRuntimeCommand = True: Exit Function
      Case "_help"
         Call informClient(vbLf & "Row:= " & theRow & ", Col:=" & theCol)
         Call informClient("--------------------------------------------------------")
         Call informClient("_help                - display this help.")
         Call informClient("_hide                - hides/shows the map.")
         Call informClient("_get                 - read room data.")
         Call informClient("_update              - update room data.")
         Call informClient("_mapnorth            - maps the room north.")
         Call informClient("_mapeast, _mapsouth, _mapwest, _mapup, _mapdown")
         Call informClient("_n                   - create exit north. _e, _s, _w, _u, _d")
         Call informClient("_t [road|plain|forest|swamp|hill|Underground|water|mountain]")
         Call informClient("_sun                 - toggle sun on/off")
         Call informClient("_ride                - toggle rideable on/off")
         Call informClient("_nd [doorname]       - create door. _ed, _sd, _wd, _ud, _dd")
         Call informClient("_np [row],[column]   - create portal. _ep, _sp, _wp, up, _dp")
         Call informClient("_go [row],[column]   - jump to specified row/col on map")
         Call informClient("_movemapnorth        - move map. _movemapeast, _movemapsouth, _movemapwest")
         Call informClient("_free                - free movement, when dislocated.")
         Call informClient("--------------------------------------------------------")
         Call informClient(" ")
         frmMap.tcpPlayer.SendData vbCrLf
         handleRuntimeCommand = True: Exit Function
      End Select
      If checkString(strData, "_go ") = True Then
         Call gotoArea(Mid(strData, 5))
         handleRuntimeCommand = True: Exit Function
      End If
   End If

   If LOST = True Then
      frmMap.tcpPlayer.SendData strData
      handleRuntimeCommand = True: Exit Function
   End If
   
Dim theCommand
If a = (1 + specialLen) Then
   Select Case b
   Case "n"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, newNorth, N_MAP, virtualRow - 1, virtualCol)
   Case "e"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, newEast, E_MAP, virtualRow, virtualCol + 1)
   Case "s"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, newSouth, S_MAP, virtualRow + 1, virtualCol)
   Case "w"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, newWest, W_MAP, virtualRow, virtualCol - 1)
   Case "u"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, newUp, U_MAP, virtualRow, virtualCol)
   Case "d"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, newDown, D_MAP, virtualRow, virtualCol)
   Case Else
      frmMap.tcpPlayer.SendData strData
   End Select
Else
   Dim i
   If specialLen = 1 Then theCommand = Split(strData, vbLf)     'jmc  - Lf
   If specialLen = 2 Then theCommand = Split(strData, vbCrLf)   'zmud, telnet - CrLf
   For i = LBound(theCommand) To UBound(theCommand) - 1
      If Len(theCommand(i)) = 1 Then
         Select Case theCommand(i)
         Case "n"
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, newNorth, N_MAP, virtualRow - 1, virtualCol)
         Case "e"
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, newEast, E_MAP, virtualRow, virtualCol + 1)
         Case "s"
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, newSouth, S_MAP, virtualRow + 1, virtualCol)
         Case "w"
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, newWest, W_MAP, virtualRow, virtualCol - 1)
         Case "u"
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, newUp, U_MAP, virtualRow, virtualCol)
         Case "d"
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, newDown, D_MAP, virtualRow, virtualCol)
         Case Else
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            'frmMap.tcpPlayer.SendData strData
         End Select
      Else
         frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
      End If
   Next
End If
Exit Function
errorhandler:
   errorModule = Err.Description & "(" & Err.Number & ") -> " & "Client_Runtime"
   writeError (errorModule)
End Function

Public Function travel(data, roomcount, newExit, theMap, row, col)
   travel = True
   If LOST = True Then Exit Function
   If freedom Then freedom = False: Exit Function
   If (theExitNorth Or theExitEast Or theExitSouth Or theExitWest Or theExitUp Or theExitDown) = False Then
      Call SYNC_FALSE("There are no exits!")
      travel = True: Exit Function  'was just EXIT FUNCTION
   End If
   If checkArrayLimit(virtualRow, virtualCol) = False Then
      Call SYNC_FALSE("Invalid virtual coordinates!")
      travel = True: Exit Function
   End If
   If checkArrayLimit(row, col) = False Then
      Call SYNC_FALSE("Invalid reference coordinates!")
      travel = True: Exit Function
   End If
   travel = False
   If surfing = True Then
      surfing = False
      theRow = virtualRow
      theCol = virtualCol
   End If
   If (arrData(arrWorld(virtualRow, virtualCol), cDATA) And theMap) > 0 Then  'there is an Exit
      currentData = (arrData(arrWorld(virtualRow, virtualCol), cDATA) And theMap)
      Select Case currentData
      Case N_exit, E_exit, S_exit, W_exit
         If stackRoom(roomcount, row, col, data) Then travel = True
      Case N_door, E_door, S_door, W_door, N_hiddendoor, E_hiddendoor, S_hiddendoor, W_hiddendoor
         If stackRoom(roomcount, row, col, data) Then travel = True
      Case U_exit, U_door, U_portal, U_doorportal, U_hiddendoor, (U_hiddendoor Or U_portal)
         If stackRoom(roomcount, arrData(arrWorld(virtualRow, virtualCol), cUPORTALR), arrData(arrWorld(virtualRow, virtualCol), cUPORTALC), data) Then travel = True
      Case D_exit, D_door, D_portal, D_doorportal, D_hiddendoor, (D_hiddendoor Or D_portal)
         If stackRoom(roomcount, arrData(arrWorld(virtualRow, virtualCol), cDPORTALR), arrData(arrWorld(virtualRow, virtualCol), cDPORTALC), data) Then travel = True
      Case N_portal, N_doorportal, (N_hiddendoor Or N_portal)
         If stackRoom(roomcount, arrData(arrWorld(virtualRow, virtualCol), cNPORTALR), arrData(arrWorld(virtualRow, virtualCol), cNPORTALC), data) Then travel = True
      Case E_portal, E_doorportal, (E_hiddendoor Or E_portal)
         If stackRoom(roomcount, arrData(arrWorld(virtualRow, virtualCol), cEPORTALR), arrData(arrWorld(virtualRow, virtualCol), cEPORTALC), data) Then travel = True
      Case S_portal, S_doorportal, (S_hiddendoor Or S_portal)
         If stackRoom(roomcount, arrData(arrWorld(virtualRow, virtualCol), cSPORTALR), arrData(arrWorld(virtualRow, virtualCol), cSPORTALC), data) Then travel = True
      Case W_portal, W_doorportal, (W_hiddendoor Or W_portal)
         If stackRoom(roomcount, arrData(arrWorld(virtualRow, virtualCol), cWPORTALR), arrData(arrWorld(virtualRow, virtualCol), cWPORTALC), data) Then travel = True
      End Select
   Else
'      If stackOUT = stackIN And virtualRow = theRow And virtualCol = theCol Then
'         If stackRoom(roomcount, virtualRow, virtualCol, data) Then travel = True
'         Call informClient("Alas! --- THIS IS IT! --- You cannot go that way.")
'         If newExit Then
'            travel = True
'            Call SYNC_FALSE("New exit discovered!")
'         End If
'         travel = True
'      Else
         If stackRoom(roomcount, virtualRow, virtualCol, data) Then travel = True
         Call informClient("...... no exit ......")
'      End If
   End If
End Function

Public Function stackRoom(roomcount, row, col, data)
'   If row = 300 And col = 600 Then Call informClient(vbLf & ("DO YOU HAVE A DEATHWISH?")): stackRoom = False
   stackRoom = False
   If checkArrayLimit(row, col) = False Then
      stackRoom = True
      Call SYNC_FALSE("Invalid stack reference coordinates!")
      Exit Function
   End If
   If roomcount < arrMaxRoom Then 'ubound(roomstack)
      If roomcount < 1 Then   'roomcount < 1, resetting
         roomcount = 0
         stackIN = 0
         stackOUT = 1
      End If
      If stackIN = 0 And stackOUT = 0 Then   'invalid stack, stackin=0 and stackout=0
         roomcount = 0
         stackIN = 0
         stackOUT = 1
         stackRoom = True
         Call SYNC_FALSE("Invalid stack, in = 0, out = 0!")
         Exit Function
      End If
      If stackIN > 0 And stackOUT > stackIN Then   'stackout > stackin+1 " & stackIN + 1
         roomcount = 0
         stackIN = 0
         stackOUT = 1
         stackRoom = True
         Call SYNC_FALSE("Invalid stack, in > 0, out > stackin!")
         Exit Function
      End If
      If stackIN = arrMaxRoom Then
         If stackOUT > 1 Then 'stackIN & "=stackin >= arrmaxroom=" & arrMaxRoom
            Dim n
            For n = 0 To (stackIN - stackOUT)
               arrRoomstack(n + 1, 1) = arrRoomstack(stackOUT + n, 1)
               arrRoomstack(n + 1, 2) = arrRoomstack(stackOUT + n, 2)
               arrMovestack(n + 1) = arrMovestack(stackOUT + n)
            Next
            stackIN = roomcount
            stackOUT = 1
         Else
            stackRoom = True
            Call SYNC_FALSE("Invalid stack, out = " & stackOUT)
            Exit Function
         End If
      End If
      roomcount = roomcount + 1
      stackIN = stackIN + 1
      arrRoomstack(stackIN, 1) = row
      arrRoomstack(stackIN, 2) = col
      arrMovestack(stackIN) = data    '& vbCrLf
      virtualRow = row
      virtualCol = col
      stackRoom = True
   Else
      roomcount = 0
      stackIN = 0
      stackOUT = 1
      stackRoom = True
      Call SYNC_FALSE("Invalid stack, roomcount >= arrMaxRoom, roomcount = " & roomcount & ", in = " & stackIN & ", out = " & stackOUT)
   End If
End Function
Public Sub newUpdateTheRoom()
errorData = errorData & "newUpdateTheRoom -> "
   If stackIN = 0 Then
      Call SYNC_FALSE("Invalid room update stack, in = 0!")
   Else
      If arrRoomstack(stackOUT, 1) > 0 And arrRoomstack(stackOUT, 2) > 0 Then
         theRow = arrRoomstack(stackOUT, 1)
         theCol = arrRoomstack(stackOUT, 2)
         'next room
         If roomcount > 0 Then
            roomcount = roomcount - 1
            stackOUT = stackOUT + 1
            If roomcount = 0 Then stackOUT = 1: stackIN = 0
         Else
         End If
         Call loadRoom(theRow, theCol)
         If SyncError Then
            LOST = True
            If Autosync And MappingMode = False Then
               Call caseFleeHandler(currentRoomName, currentExits, 6, False, False): Exit Sub
            Else
'loadroom error
               Call SYNC_FALSE("Invalid room update, cannot load room!"): Exit Sub
            End If
         End If
         Call DrawMap
      End If
   End If
End Sub
Public Sub newCollision()
errorData = errorData & "newCollision -> "
   roomcount = roomcount - 1
   stackOUT = stackOUT + 1
   virtualRow = theRow
   virtualCol = theCol
   If roomcount = 0 Then stackOUT = 1: stackIN = 0: Erase arrMovestack: Exit Sub
   For coll = stackOUT To stackIN
      b = arrMovestack(coll)
      Select Case b
      Case "n"
         Call collisionTravel(N_MAP, virtualRow - 1, virtualCol)
      Case "e"
         Call collisionTravel(E_MAP, virtualRow, virtualCol + 1)
      Case "s"
         Call collisionTravel(S_MAP, virtualRow + 1, virtualCol)
      Case "w"
         Call collisionTravel(W_MAP, virtualRow, virtualCol - 1)
      Case "u"
         Call collisionTravel(U_MAP, virtualRow, virtualCol)
      Case "d"
         Call collisionTravel(D_MAP, virtualRow, virtualCol)
      End Select
   Next
End Sub

Public Function collisionTravel(theMap, row, col)
   If (theExitNorth Or theExitEast Or theExitSouth Or theExitWest Or theExitUp Or theExitDown) = False Then
      Call SYNC_FALSE("Collision travel, there are no exist!"): Exit Function
   End If
   If checkArrayLimit(virtualRow, virtualCol) = False Then
      Call SYNC_FALSE("Collision travel, virtual coordinates are invalid!"): Exit Function
   End If
   If checkArrayLimit(row, col) = False Then
      Call SYNC_FALSE("Collision travel, reference coordinates are invalid!"): Exit Function
   End If
   If (arrData(arrWorld(virtualRow, virtualCol), cDATA) And theMap) > 0 Then  'there is an Exit
      currentData = (arrData(arrWorld(virtualRow, virtualCol), cDATA) And theMap)
      Select Case currentData
      Case N_exit, E_exit, S_exit, W_exit
         Call collisionStackRoom(row, col)
      Case N_door, E_door, S_door, W_door, N_hiddendoor, E_hiddendoor, S_hiddendoor, W_hiddendoor
         Call collisionStackRoom(row, col)
      Case U_exit, U_door, U_portal, U_doorportal, U_hiddendoor, (U_hiddendoor Or U_portal)
         Call collisionStackRoom(arrData(arrWorld(virtualRow, virtualCol), cUPORTALR), arrData(arrWorld(virtualRow, virtualCol), cUPORTALC))
      Case D_exit, D_door, D_portal, D_doorportal, D_hiddendoor, (D_hiddendoor Or D_portal)
         Call collisionStackRoom(arrData(arrWorld(virtualRow, virtualCol), cDPORTALR), arrData(arrWorld(virtualRow, virtualCol), cDPORTALC))
      Case N_portal, N_doorportal, (N_hiddendoor Or N_portal)
         Call collisionStackRoom(arrData(arrWorld(virtualRow, virtualCol), cNPORTALR), arrData(arrWorld(virtualRow, virtualCol), cNPORTALC))
      Case E_portal, E_doorportal, (E_hiddendoor Or E_portal)
         Call collisionStackRoom(arrData(arrWorld(virtualRow, virtualCol), cEPORTALR), arrData(arrWorld(virtualRow, virtualCol), cEPORTALC))
      Case S_portal, S_doorportal, (S_hiddendoor Or S_portal)
         Call collisionStackRoom(arrData(arrWorld(virtualRow, virtualCol), cSPORTALR), arrData(arrWorld(virtualRow, virtualCol), cSPORTALC))
      Case W_portal, W_doorportal, (W_hiddendoor Or W_portal)
         Call collisionStackRoom(arrData(arrWorld(virtualRow, virtualCol), cWPORTALR), arrData(arrWorld(virtualRow, virtualCol), cWPORTALC))
      End Select
   Else
      Call collisionStackRoom(virtualRow, virtualCol)
   End If
End Function
Public Function collisionStackRoom(row, col)
   If checkArrayLimit(row, col) = False Then
      Call SYNC_FALSE("Collision stack room, reference coordinates are invalid!")
      Exit Function
   End If
   arrRoomstack(coll, 1) = row
   arrRoomstack(coll, 2) = col
   virtualRow = row
   virtualCol = col
End Function
