Attribute VB_Name = "CLIENT_RUNTIME"
Option Explicit
Private a As Integer
Private b As String
Private c As String
Private n As Integer
Public stackOUT As Integer
Public stackIN As Integer
Public coll As Integer
Public atCrLf As Integer
Public atLf As Integer
Public specialLen As Integer
Public surfing As Boolean
Public lookfor As String
Public targetFound As Boolean
Public arrTimer(0 To 20, 0 To 1) As String
Public lastMoveDirection As String
Public Function noCRLF(ByVal s As String) As String
         s = Replace(s, vbCr, "", , , vbBinaryCompare)
         s = Replace(s, vbLf, "", , , vbBinaryCompare)
         noCRLF = s
End Function
Public Function handleSpecial(strData As String) As Boolean
errorData = errorData & "handleSpecial -> "
   handleSpecial = False
   If LenB(strData) = specialLen Then
      frmMap.tcpPlayer.SendData strData
      handleSpecial = True
      If frmMap.mnuPortals.Checked Then
         Call DrawMap
      End If
      Exit Function
   End If
End Function
Public Function getMyTime(myTime As String) As String
   If LenB(myTime) <> 0 Then
      Dim mins, secs As Long
      secs = Fix(DateTime.DateDiff("s", myTime, Now(), vbMonday))
      mins = Fix(secs / 60)
      secs = secs - (mins * 60)
      If secs < 0 Then secs = 0
      getMyTime = " (" & Right("0" & CStr(mins), 2) & ":" & Right("0" & CStr(secs), 2) & ")"
   Else
      getMyTime = vbNullString
   End If
End Function


Public Function handleRuntimeCommand(ByVal strData As String) As Boolean
If DEBUGMODE = False Then On Error GoTo errorhandler
errorData = errorData & "handleRuntimeCommand -> "
   handleRuntimeCommand = False

If GODMODE And LOST = False Then
    Dim mys As String
    mys = LCase(strData)
    
    'If LOST = False And _
    (InStrB(1, mys, "open ", vbBinaryCompare) > 0 Or InStrB(1, mys, "close ", vbBinaryCompare) > 0 _
    Or InStrB(1, mys, "bash ", vbBinaryCompare) > 0 _
    Or InStrB(1, mys, "lock ", vbBinaryCompare) > 0 _
    Or InStrB(1, mys, "unlock ", vbBinaryCompare) > 0 _
    Or InStrB(1, mys, "pick ", vbBinaryCompare) > 0) Then
    If InStrB(1, mys, " exit ", vbBinaryCompare) > 0 Then
      Select Case True
      Case InStrB(1, mys, " exit n", vbBinaryCompare) > 0
         If LenB(aData(getIndex(virtualRow, virtualCol), cNDOOR)) > 0 Then mys = Replace(mys, "exit n", aData(getIndex(virtualRow, virtualCol), cNDOOR) & " n", , 1, vbBinaryCompare)
      Case InStrB(1, mys, " exit e", vbBinaryCompare) > 0
         If LenB(aData(getIndex(virtualRow, virtualCol), cEDOOR)) > 0 Then mys = Replace(mys, "exit e", aData(getIndex(virtualRow, virtualCol), cEDOOR) & " e", , 1, vbBinaryCompare)
      Case InStrB(1, mys, " exit s", vbBinaryCompare) > 0
         If LenB(aData(getIndex(virtualRow, virtualCol), cSDOOR)) > 0 Then mys = Replace(mys, "exit s", aData(getIndex(virtualRow, virtualCol), cSDOOR) & " s", , 1, vbBinaryCompare)
      Case InStrB(1, mys, " exit w", vbBinaryCompare) > 0
         If LenB(aData(getIndex(virtualRow, virtualCol), cWDOOR)) > 0 Then mys = Replace(mys, "exit w", aData(getIndex(virtualRow, virtualCol), cWDOOR) & " w", , 1, vbBinaryCompare)
      Case InStrB(1, mys, " exit u", vbBinaryCompare) > 0
         If LenB(aData(getIndex(virtualRow, virtualCol), cUDOOR)) > 0 Then mys = Replace(mys, "exit u", aData(getIndex(virtualRow, virtualCol), cUDOOR) & " u", , 1, vbBinaryCompare)
      Case InStrB(1, mys, " exit d", vbBinaryCompare) > 0
         If LenB(aData(getIndex(virtualRow, virtualCol), cDDOOR)) > 0 Then mys = Replace(mys, "exit d", aData(getIndex(virtualRow, virtualCol), cDDOOR) & " d", , 1, vbBinaryCompare)
      End Select
      If InStrB(1, mys, "@", vbBinaryCompare) > 0 Then mys = Replace(mys, "@", "", , 1, vbBinaryCompare)
      frmMap.tcpPlayer.SendData mys ' & vbCrLf
      Call informClient(lookHeader & WHITE & ";" & BOLD & lookFooter & "" & noCRLF(mys) & "" & colourEndCode, False)
      handleRuntimeCommand = True: Exit Function
   End If
End If


   Dim timeNumber, elapsed As String
   a = LenB(strData)
   b = LCase(MidB$(strData, 1, LenB("1")))
   If b = "_" Then
      c = LCase$(MidB$(strData, 1, LenB("12345")))
      Select Case c
      Case "_test"
         Call frmMap.mnuTest_Click
         handleRuntimeCommand = True: Exit Function
      Case "_lost"
         LOST = True
         handleRuntimeCommand = True: Exit Function
      Case "_time"
         handleRuntimeCommand = True
         strData = Replace(strData, vbCr, "", , , vbBinaryCompare)
         strData = Replace(strData, vbLf, "", , , vbBinaryCompare)
         timeNumber = Trim(Mid(strData, 6, 1))
         If LenB(timeNumber) = 0 Then
            For n = 0 To 9
               If LenB(arrTimer(n, 1)) <> 0 Then
                  informClient CStr(n) & ") " & arrTimer(n, 0) & Space(50 - Len(arrTimer(n, 0))) & getMyTime(arrTimer(n, 1)), True
               End If
            Next
         Else
            If Not IsNumeric(timeNumber) Then Exit Function
            If timeNumber < LBound(arrTimer) Or timeNumber > UBound(arrTimer) Then Exit Function
            timeNumber = CInt(timeNumber)
            Dim timeName As String
            timeName = Trim(Mid(strData, 8))
            ' kui sisestati "_time1", siis nullime
            If LenB(timeName) = 0 Then
               informClient " - " & arrTimer(timeNumber, 0) & " lasted " & getMyTime(arrTimer(timeNumber, 1)), True
               arrTimer(timeNumber, 0) = vbNullString
               arrTimer(timeNumber, 1) = vbNullString
            Else
               arrTimer(timeNumber, 0) = timeName
               arrTimer(timeNumber, 1) = Now()
            End If
         End If
         Exit Function
'      Case "_hist"   'history of enemies
'         Call showEnemy
'         handleRuntimeCommand = True: Exit Function
'      Case "_null"   'history of enemies
'         Erase arrEnemies
'         indexEnemies = 0
'         handleRuntimeCommand = True: Exit Function
      Case "_walk"
         If GODMODE Then Call frmMap.mnuWalk_Click
      Case "_here"
         If GODMODE Then Call frmMap.mnuHere_Click
      Case "_lead"
         'If GODMODE Then
            frmMap.mnuFollow.Checked = True
            If LenB(strData) > LenB("12345678") Then
               leader = Trim(MidB$(strData, 13, LenB(strData) - LenB("1234567")))
               leader = Replace(leader, vbCr, "", , , vbBinaryCompare)
               leader = Replace(leader, vbLf, "", , , vbBinaryCompare)
               leader = UCase(Mid(leader, 1, 1)) & LCase(Mid(leader, 2))
            Else
               leader = vbNullString
            End If
            Call informClient("Leader: >" & leader & "<")
            handleRuntimeCommand = True: Exit Function
         'End If
      Case "_undo"
      If GODMODE Then
         If canUndo Then
            If LenB(aData(getIndex(virtualRow, virtualCol), cDATA)) = 0 Then Exit Function
            Dim data As Long
            data = aData(getIndex(virtualRow, virtualCol), cDATA)
            If (data And N_exit) Then  'NORTH
               If (data And N_portal) = N_portal Or (data And N_doorportal) = N_doorportal Or (data And (N_hiddendoor Or N_portal)) = (N_hiddendoor Or N_portal) Then
                  If (aData(getIndex(virtualRow, virtualCol), cNPORTALR)) = undoRow And _
                     (aData(getIndex(virtualRow, virtualCol), cNPORTALC)) = undoCol Then
                        b = "n": a = LenB("12"): strData = "n" & vbCrLf
                  End If
               Else
                  If LenB(aData(getIndex(virtualRow - 1, virtualCol), cDATA)) <> 0 Then
                     If (virtualRow - 1) = undoRow Then b = "n": a = LenB("12"): strData = "n" & vbCrLf
                  End If
               End If
            End If
            
            If (data And S_exit) Then  'SOUTH
               If (data And S_portal) = S_portal Or (data And S_doorportal) = S_doorportal Or (data And (S_hiddendoor Or S_portal)) = (S_hiddendoor Or S_portal) Then
                  If (aData(getIndex(virtualRow, virtualCol), cSPORTALR)) = undoRow And _
                     (aData(getIndex(virtualRow, virtualCol), cSPORTALC)) = undoCol Then
                        b = "s": a = LenB("12"): strData = "s" & vbCrLf
                  End If
               Else
                  If LenB(aData(getIndex(virtualRow + 1, virtualCol), cDATA)) <> 0 Then
                     If (virtualRow + 1) = undoRow Then b = "s": a = LenB("12"): strData = "s" & vbCrLf
                  End If
               End If
            End If
            
            If (data And E_exit) Then  'EAST
               If (data And E_portal) = E_portal Or (data And E_doorportal) = E_doorportal Or (data And (E_hiddendoor Or E_portal)) = (E_hiddendoor Or E_portal) Then
                  If (aData(getIndex(virtualRow, virtualCol), cEPORTALR)) = undoRow And _
                     (aData(getIndex(virtualRow, virtualCol), cEPORTALC)) = undoCol Then
                        b = "e": a = LenB("12"): strData = "e" & vbCrLf
                  End If
               Else
                  If LenB(aData(getIndex(virtualRow, virtualCol + 1), cDATA)) <> 0 Then
                     If (virtualCol + 1) = undoCol Then b = "e": a = LenB("12"): strData = "e" & vbCrLf
                  End If
               End If
            End If
            
            If (data And W_exit) Then  'WEST
               If (data And W_portal) = W_portal Or (data And W_doorportal) = W_doorportal Or (data And (W_hiddendoor Or W_portal)) = (W_hiddendoor Or W_portal) Then
                  If (aData(getIndex(virtualRow, virtualCol), cWPORTALR)) = undoRow And _
                     (aData(getIndex(virtualRow, virtualCol), cWPORTALC)) = undoCol Then
                        b = "w": a = LenB("12"): strData = "w" & vbCrLf
                  End If
               Else
                  If LenB(aData(getIndex(virtualRow, virtualCol - 1), cDATA)) <> 0 Then
                     If (virtualCol - 1) = undoCol Then b = "w": a = LenB("12"): strData = "w" & vbCrLf
                  End If
               End If
            End If

            If (data And U_exit) Then  'UP
               If (aData(getIndex(virtualRow, virtualCol), cUPORTALR)) = undoRow And _
                  (aData(getIndex(virtualRow, virtualCol), cUPORTALC)) = undoCol Then
                     b = "u": a = LenB("12"): strData = "u" & vbCrLf
               End If
            End If
            If (data And D_exit) Then  'DOWN
               If (aData(getIndex(virtualRow, virtualCol), cDPORTALR)) = undoRow And _
                  (aData(getIndex(virtualRow, virtualCol), cDPORTALC)) = undoCol Then
                     b = "d": a = LenB("12"): strData = "d" & vbCrLf
               End If
            End If
         Else
            strData = vbNullString
         End If
      End If
      Case "_hide", "_show"
         If frmMap.WindowState = vbMinimized Then
            frmMap.WindowState = vbNormal
            Call DrawMap
         Else
            frmMap.WindowState = vbMinimized
         End If
         handleRuntimeCommand = True: Exit Function
      Case "_canc"
         Call cancelBuffer
         handleRuntimeCommand = True: Exit Function
      'Case "_show"
      '   frmMap.Hide
      '   handleRuntimeCommand = True: Exit Function
      Case "_sync", "_gogo"
         Call frmMap.mnuLocate_Click
         handleRuntimeCommand = True: Exit Function
      Case "_tool", "_mapp"
         Call frmMap.mnuTools_Click
         handleRuntimeCommand = True: Exit Function
      Case "_desc"
         frmMap.mnuMapDescription_Click: Exit Function
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
         If a > 14 Then
            lookfor = LCase(MidB$(strData, 13, a - 14))
            Call DrawMap
         End If
         handleRuntimeCommand = True: Exit Function
     Case "_goto"
         If a > 14 Then
            Call gotoArea(MidB$(strData, 13, a - 14))
         End If
         handleRuntimeCommand = True: Exit Function
      Case "_help"
         Call informClient(vbLf & "Row:= " & theROW & ", Col:=" & theCOL)
         Call informClient("========================================================")
         Call informClient("_sync, _gogo         - finds your location and stops mapping")
         Call informClient("_tool, _mapp         - starts mapping mode")
         Call informClient("_desc                - switches description show when mapping")
         Call informClient("---------------- mapping -------------------------------")
         Call informClient("_mapnorth, _mapDIR.. - maps the room north.")
         Call informClient("_n                   - create exit north. _e, _s, _w, _u, _d")
         Call informClient("_t [road|plain|forest|swamp|hill|underground|water|mountain]")
         Call informClient("_sun                 - toggle sun on/off")
         Call informClient("_ride                - toggle rideable on/off")
         Call informClient("_nd [doorname]       - create door. _ed, _sd, _wd, _ud, _dd")
         Call informClient("_np [row],[column]   - create portal. _ep, _sp, _wp, up, _dp")
         Call informClient("_movemapnorth        - move map. _movemapeast, _movemapsouth, _movemapwest")
         Call informClient("--------------------------------------------------------")
         Call informClient("_find                - searches roomnames and shows on map")
         Call informClient("_hide or _show       - minimizes or maximizes the map.")
         Call informClient("--------------------------------------------------------")
         Call informClient("_time                - shows active timers")
         Call informClient("_time1               - resets 1st timer")
         Call informClient(" example: #action {You feel strong.} {_time2 - strength}")
         Call informClient(" example: #action {You feel weak.} {_time2}")
         Call informClient("--------------------------------------------------------")
         frmMap.tcpPlayer.SendData vbCrLf
         handleRuntimeCommand = True: Exit Function
      End Select
   End If

   If LOST Then handleRuntimeCommand = True
      'frmMap.tcpPlayer.SendData strData
      'Exit Function
      'End If
   
Dim theCommand
If a = (2 + specialLen) Then
   Select Case b
   Case "n"
      lastMoveDirection = "n"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, N_MAP, virtualRow - 1, virtualCol)
   Case "e"
      lastMoveDirection = "e"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, E_MAP, virtualRow, virtualCol + 1)
   Case "s"
      lastMoveDirection = "s"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, S_MAP, virtualRow + 1, virtualCol)
   Case "w"
      lastMoveDirection = "w"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, W_MAP, virtualRow, virtualCol - 1)
   Case "u"
      lastMoveDirection = "u"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, U_MAP, virtualRow, virtualCol)
   Case "d"
      lastMoveDirection = "d"
      frmMap.tcpPlayer.SendData strData
      Call travel(b, roomcount, D_MAP, virtualRow, virtualCol)
   Case Else
      frmMap.tcpPlayer.SendData strData
   End Select
Else
   Dim i
   If specialLen = 2 Then theCommand = Split(strData, vbLf, , vbBinaryCompare)   'jmc  - Lf
   If specialLen = 4 Then theCommand = Split(strData, vbCrLf, , vbBinaryCompare) 'zmud, telnet - CrLf
   For i = LBound(theCommand) To UBound(theCommand) - 1
      If LenB(theCommand(i)) = 2 Then
         Select Case theCommand(i)
         Case "n"
            lastMoveDirection = vbNullString
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, N_MAP, virtualRow - 1, virtualCol)
         Case "e"
            lastMoveDirection = vbNullString
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, E_MAP, virtualRow, virtualCol + 1)
         Case "s"
            lastMoveDirection = vbNullString
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, S_MAP, virtualRow + 1, virtualCol)
         Case "w"
            lastMoveDirection = vbNullString
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, W_MAP, virtualRow, virtualCol - 1)
         Case "u"
            lastMoveDirection = vbNullString
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, U_MAP, virtualRow, virtualCol)
         Case "d"
            lastMoveDirection = vbNullString
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
            Call travel(theCommand(i), roomcount, D_MAP, virtualRow, virtualCol)
         Case Else
            frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
         End Select
      Else
         frmMap.tcpPlayer.SendData theCommand(i) & vbCrLf
      End If
   Next
End If
Exit Function

errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "Client_Runtime"
   writeError (errorModule)
End Function


Public Function travel(data, roomcount, theMap, row, col, Optional leading As Boolean)
   If LOST Then Exit Function
   travel = True
   If isValid(virtualRow, virtualCol) = False Then Call SYNC_FALSE("Invalid virtual coordinates!"): travel = True: Exit Function
   If isValid(row, col) = False Then Call SYNC_FALSE("Invalid reference coordinates!"): travel = True: Exit Function
   If surfing Then
      surfing = False
      theROW = virtualRow
      theCOL = virtualCol
   End If
   travel = False
   If LenB(aData(getIndex(virtualRow, virtualCol), cDATA)) <> 0 Then
      If (aData(getIndex(virtualRow, virtualCol), cDATA) And theMap) > 0 Then 'there is an Exit
        
         Select Case (aData(getIndex(virtualRow, virtualCol), cDATA) And theMap)  'currentData
         Case N_exit, N_door, N_hiddendoor, N_portal, N_doorportal, (N_hiddendoor Or N_portal)
            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cNPORTALR), aData(getIndex(virtualRow, virtualCol), cNPORTALC), data) Then travel = True
         Case E_exit, E_door, E_hiddendoor, E_portal, E_doorportal, (E_hiddendoor Or E_portal)
            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cEPORTALR), aData(getIndex(virtualRow, virtualCol), cEPORTALC), data) Then travel = True
         Case S_exit, S_door, S_hiddendoor, S_portal, S_doorportal, (S_hiddendoor Or S_portal)
            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cSPORTALR), aData(getIndex(virtualRow, virtualCol), cSPORTALC), data) Then travel = True
         Case W_exit, W_door, W_hiddendoor, W_portal, W_doorportal, (W_hiddendoor Or W_portal)
            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cWPORTALR), aData(getIndex(virtualRow, virtualCol), cWPORTALC), data) Then travel = True
         Case U_exit, U_door, U_portal, U_doorportal, U_hiddendoor, (U_hiddendoor Or U_portal)
            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cUPORTALR), aData(getIndex(virtualRow, virtualCol), cUPORTALC), data) Then travel = True
         Case D_exit, D_door, D_portal, D_doorportal, D_hiddendoor, (D_hiddendoor Or D_portal)
            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cDPORTALR), aData(getIndex(virtualRow, virtualCol), cDPORTALC), data) Then travel = True
         End Select

'before level structure
'         Select Case (aData(getIndex(virtualRow, virtualCol), cDATA) And theMap)  'currentData
'         Case N_exit, E_exit, S_exit, W_exit
'            If stackRoom(roomcount, row, col, data) Then travel = True
'         Case N_door, E_door, S_door, W_door, N_hiddendoor, E_hiddendoor, S_hiddendoor, W_hiddendoor
'            If stackRoom(roomcount, row, col, data) Then travel = True
'         Case U_exit, U_door, U_portal, U_doorportal, U_hiddendoor, (U_hiddendoor Or U_portal)
'            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cUPORTALR), aData(getIndex(virtualRow, virtualCol), cUPORTALC), data) Then travel = True
'         Case D_exit, D_door, D_portal, D_doorportal, D_hiddendoor, (D_hiddendoor Or D_portal)
'            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cDPORTALR), aData(getIndex(virtualRow, virtualCol), cDPORTALC), data) Then travel = True
'         Case N_portal, N_doorportal, (N_hiddendoor Or N_portal)
'            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cNPORTALR), aData(getIndex(virtualRow, virtualCol), cNPORTALC), data) Then travel = True
'         Case E_portal, E_doorportal, (E_hiddendoor Or E_portal)
'            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cEPORTALR), aData(getIndex(virtualRow, virtualCol), cEPORTALC), data) Then travel = True
'         Case S_portal, S_doorportal, (S_hiddendoor Or S_portal)
'            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cSPORTALR), aData(getIndex(virtualRow, virtualCol), cSPORTALC), data) Then travel = True
'         Case W_portal, W_doorportal, (W_hiddendoor Or W_portal)
'            If stackRoom(roomcount, aData(getIndex(virtualRow, virtualCol), cWPORTALR), aData(getIndex(virtualRow, virtualCol), cWPORTALC), data) Then travel = True
'         End Select
      Else
         If theMap = False Then
            If stackRoom(roomcount, row, col, data) Then travel = True '### enter/leave case ###
         Else
            If stackRoom(roomcount, virtualRow, virtualCol, data) Then travel = True
         End If
      End If
   Else
      If stackRoom(roomcount, virtualRow, virtualCol, data) Then travel = True
   End If
End Function

Public Function stackRoom(roomcount, row, col, data)
   stackRoom = False
   If isValid(row, col) = False Then
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
      
      
'------- NORMAL SITUATION -------
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
'========== END ===========
      
      
      'If DEBUGMODE Then Debug.Print "S T A C K   + " & UCase(data) & "(" & roomcount & ")"
      'If DEBUGMODE Then Debug.Print "STACKING room " & UCase(data) & " arr(#" & stackIN & ", r=" & row & ", c=" & col & ", " & UCase(data) & ")    roomcount = " & roomcount
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
         theROW = arrRoomstack(stackOUT, 1)
         theCOL = arrRoomstack(stackOUT, 2)
         'next room
         If roomcount > 0 Then
            roomcount = roomcount - 1
'If DEBUGMODE Then Debug.Print "S T A C K   - " & UCase(arrMovestack(stackOUT)) & "(" & roomcount & ")"
            stackOUT = stackOUT + 1
            If roomcount = 0 Then
               stackOUT = 1
               stackIN = 0
            End If
         End If
         Call loadRoom(theROW, theCOL)
         If SyncError Then
            LOST = True
            If Autosync And MappingMode = False Then
               Call caseFleeHandler(currentRoomname, currentExits, 6, False, False): Exit Sub
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
   virtualRow = theROW
   virtualCol = theCOL
   If roomcount <= 0 Then
      stackOUT = 1
      stackIN = 0
      Erase arrMovestack
      Exit Sub
   End If
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

Public Function collisionTravel(theMap, row As Integer, col As Integer)
   If (theExitNorth Or theExitEast Or theExitSouth Or theExitWest Or theExitUp Or theExitDown) = False Then
      Call SYNC_FALSE("Collision travel, there are no exits!"): Exit Function
   End If
   If isValid(virtualRow, virtualCol) = False Then
      Call SYNC_FALSE("Collision travel, virtual coordinates are invalid!"): Exit Function
   End If
   If isValid(row, col) = False Then
      Call SYNC_FALSE("Collision travel, reference coordinates are invalid!"): Exit Function
   End If
   If LenB(aData(getIndex(virtualRow, virtualCol), cDATA)) <> 0 Then
      If (aData(getIndex(virtualRow, virtualCol), cDATA) And theMap) > 0 Then  'there is an Exit
         currentData = (aData(getIndex(virtualRow, virtualCol), cDATA) And theMap)
         Select Case currentData
         Case N_exit, E_exit, S_exit, W_exit
            Call collisionStackRoom(row, col)
         Case N_door, E_door, S_door, W_door, N_hiddendoor, E_hiddendoor, S_hiddendoor, W_hiddendoor
            Call collisionStackRoom(row, col)
         Case U_exit, U_door, U_portal, U_doorportal, U_hiddendoor, (U_hiddendoor Or U_portal)
            Call collisionStackRoom(aData(getIndex(virtualRow, virtualCol), cUPORTALR), aData(getIndex(virtualRow, virtualCol), cUPORTALC))
         Case D_exit, D_door, D_portal, D_doorportal, D_hiddendoor, (D_hiddendoor Or D_portal)
            Call collisionStackRoom(aData(getIndex(virtualRow, virtualCol), cDPORTALR), aData(getIndex(virtualRow, virtualCol), cDPORTALC))
         Case N_portal, N_doorportal, (N_hiddendoor Or N_portal)
            Call collisionStackRoom(aData(getIndex(virtualRow, virtualCol), cNPORTALR), aData(getIndex(virtualRow, virtualCol), cNPORTALC))
         Case E_portal, E_doorportal, (E_hiddendoor Or E_portal)
            Call collisionStackRoom(aData(getIndex(virtualRow, virtualCol), cEPORTALR), aData(getIndex(virtualRow, virtualCol), cEPORTALC))
         Case S_portal, S_doorportal, (S_hiddendoor Or S_portal)
            Call collisionStackRoom(aData(getIndex(virtualRow, virtualCol), cSPORTALR), aData(getIndex(virtualRow, virtualCol), cSPORTALC))
         Case W_portal, W_doorportal, (W_hiddendoor Or W_portal)
            Call collisionStackRoom(aData(getIndex(virtualRow, virtualCol), cWPORTALR), aData(getIndex(virtualRow, virtualCol), cWPORTALC))
         End Select
      Else
         If theMap = False Then
            Call collisionStackRoom(row, col) '### enter/leave case ###
         Else
            Call collisionStackRoom(virtualRow, virtualCol)
         End If
      End If
   Else
      Call collisionStackRoom(virtualRow, virtualCol)
   End If
End Function

Public Function collisionStackRoom(row, col)
   If isValid(row, col) = False Then Call SYNC_FALSE("Collision stack room, reference coordinates are invalid!"): Exit Function
   arrRoomstack(coll, 1) = row
   arrRoomstack(coll, 2) = col
   virtualRow = row
   virtualCol = col
End Function
