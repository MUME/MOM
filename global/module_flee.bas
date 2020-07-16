Attribute VB_Name = "flee"
Option Explicit
Public tmpMatch As Integer
Public tmpMatchIndex As Integer
Public GetDescription As Boolean
Public arrRoomstack(1 To 100, 1 To 2) As Integer
Public arrMovestack(1 To 100) As String
Public Get_In_Sync As Boolean
Public LOST As Boolean
Public newNorth As Boolean
Public newEast As Boolean
Public newSouth As Boolean
Public newWest As Boolean
Public newUp As Boolean
Public newDown As Boolean
Public arrTmpFleeStack(1 To 100, 1 To 2) As Integer   ' Depends of FleeMaxRadius
Public fleeMaxRadius
Public fleeStackCount As Long
Public fleeMatch As Long
Public currentRoomname As String
Public tmpMapDesc As String '* 16
Public tmpRoomDesc As String '* 16
Public encryptedDescription As String
Public currentExits As String
Public currentString As String
Public currentDesc As String
Public oldCurrentDesc As String
Public fleeSpecialCase As Boolean
Public fleeRetry As Long
Public arrCheckStack(1 To 10) As String
Public cursor As Integer
Public syncIndex As Integer
Public crawlRadius As Integer
Public stack(1 To 1000, 1 To 2) As Integer
Public isFleeing As Boolean

Public Sub caseFleeHandler(room As String, data As String, radius, checkDesc As Boolean, isFleeing As Boolean)
errorData = errorData & "caseFleeHandler -> "
If DEBUGMODE = False Then On Error GoTo errorhandler Else On Error GoTo 0
Dim n As Integer
   If (radius + 1) * 2 >= UBound(stack) Then
      Call SYNC_FALSE("caseFleeHandler, crawler stack full")
      Exit Sub
   End If
   
   cursor = 0
   crawlRadius = radius
   Dim dirs As Long
   dirs = (N_MAP Or E_MAP Or S_MAP Or W_MAP Or U_MAP Or D_MAP)
   Call Crawler(dirs, stack, 0, theROW, theCOL)

   GetDescription = checkDesc
   roomcount = 0
   Erase arrMovestack
   Erase arrRoomstack
   virtualRow = theROW
   virtualCol = theCOL

   Call setNewExits(data)   'data represents the Exits:... line
   fleeMatch = 0
   For n = 1 To cursor
      If compareFleeExit(room, stack(n, 1), stack(n, 2)) Then
         fleeMatch = fleeMatch + 1
         If fleeMatch > 1 Then Exit For ' no point checking further
         arrTmpFleeStack(fleeMatch, 1) = stack(n, 1)
         arrTmpFleeStack(fleeMatch, 2) = stack(n, 2)
      End If
   Next
   
   Select Case fleeMatch
   Case 0
      Call SYNC_FALSE("room not found!")
   Case 1
      virtualRow = arrTmpFleeStack(fleeMatch, 1)
      virtualCol = arrTmpFleeStack(fleeMatch, 2)
      Call SYNC_TRUE
   Case Is > 1
      If GetDescription Then
         frmMap.tcpPlayer.SendData "EXAMINE" & vbCrLf
         Exit Sub
      Else
         Call SYNC_FALSE("multiple matches!")
      End If
   End Select
Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "flee caseFleeHandler"
   writeError (errorModule)
End Sub

Public Function compareFleeExit(room As String, row As Integer, col As Integer) As Boolean
   compareFleeExit = False
   If LenB(aData(getIndex(row, col), cDATA)) = 0 Then Exit Function
   If LenB(room) <> LenB(aData(getIndex(row, col), cROOMNAME)) Then Exit Function

   currentData = aData(getIndex(row, col), cDATA)
   If newNorth Then If (currentData And N_MAP) = 0 Then Exit Function
   If newEast Then If (currentData And E_MAP) = 0 Then Exit Function
   If newSouth Then If (currentData And S_MAP) = 0 Then Exit Function
   If newWest Then If (currentData And W_MAP) = 0 Then Exit Function
   If newUp Then If (currentData And U_MAP) = 0 Then Exit Function
   If newDown Then If (currentData And D_MAP) = 0 Then Exit Function

   If (newNorth = False) Then If (currentData And N_MAP) = 0 Or (currentData And N_hiddendoor) = N_hiddendoor Then Else Exit Function
   If (newEast = False) Then If (currentData And E_MAP) = 0 Or (currentData And E_hiddendoor) = E_hiddendoor Then Else Exit Function
   If (newSouth = False) Then If (currentData And S_MAP) = 0 Or (currentData And S_hiddendoor) = S_hiddendoor Then Else Exit Function
   If (newWest = False) Then If (currentData And W_MAP) = 0 Or (currentData And W_hiddendoor) = W_hiddendoor Then Else Exit Function
   If (newUp = False) Then If (currentData And U_MAP) = 0 Or (currentData And U_hiddendoor) = U_hiddendoor Then Else Exit Function
   If (newDown = False) Then If (currentData And D_MAP) = 0 Or (currentData And D_hiddendoor) = D_hiddendoor Then Else Exit Function
   
   If room = aData(getIndex(row, col), cROOMNAME) Then compareFleeExit = True
End Function

Public Sub setNewExits(data As String)
errorData = errorData & "setNewExits -> "
   If checkStringCS(data, "north") Then newNorth = True Else newNorth = False
   If checkStringCS(data, "east") Then newEast = True Else newEast = False
   If checkStringCS(data, "south") Then newSouth = True Else newSouth = False
   If checkStringCS(data, "west") Then newWest = True Else newWest = False
   If checkStringCS(data, "up") Then newUp = True Else newUp = False
   If checkStringCS(data, "down") Then newDown = True Else newDown = False
End Sub

Public Sub SYNC_TRUE(Optional message As String, Optional index As Integer)
   fleeRetry = 0
   LOST = False
   theROW = virtualRow
   theCOL = virtualCol
   theLEVEL = getLng(aData(index, cLEVEL))
   Call loadRoom(theROW, theCOL)
   Call DrawMap
   If LOST = False Then
      Call informClient("Ok. " & message)
      frmMap.Caption = mapTitle & " - " & "Ok"
   End If
End Sub

Public Sub SYNC_FALSE(message)
   LOST = True
   Erase actual
   Erase potential
   If MappingMode = False Then
      Dim s As String
      s = "Lost! "
      If LenB(message) <> 0 Then
         s = s & "(" & message & ")"
         Call informClient(s)
      End If
   End If
   If LOST = True Then
      frmMap.Caption = mapTitle & " - " & "Lost"
   End If
End Sub

Public Function isValid(row, col) As Boolean
   isValid = False
   If row < arrMinRow Then Exit Function
   If col < arrMinCol Then Exit Function
   If row > arrMaxRow Then Exit Function
   If col > arrMaxCol Then Exit Function
   isValid = True
End Function

Public Function Crawler(dirs, ByRef arrName() As Integer, ByVal reach As Integer, ByRef row As Integer, ByRef col As Integer)
' was by value
   If reach = crawlRadius Then Exit Function
   If (dirs And N_MAP) = N_MAP Then _
      If Scout(arrName, row, col, N_MAP, -1, 0) Then Call Crawler(dirs, arrName, reach + 1, arrName(cursor, 1), arrName(cursor, 2))
   If (dirs And E_MAP) = E_MAP Then _
      If Scout(arrName, row, col, E_MAP, 0, 1) Then Call Crawler(dirs, arrName, reach + 1, arrName(cursor, 1), arrName(cursor, 2))
   If (dirs And S_MAP) = S_MAP Then _
      If Scout(arrName, row, col, S_MAP, 1, 0) Then Call Crawler(dirs, arrName, reach + 1, arrName(cursor, 1), arrName(cursor, 2))
   If (dirs And W_MAP) = W_MAP Then _
      If Scout(arrName, row, col, W_MAP, 0, -1) Then Call Crawler(dirs, arrName, reach + 1, arrName(cursor, 1), arrName(cursor, 2))
   If (dirs And U_MAP) = U_MAP Then _
      If Scout(arrName, row, col, U_MAP, 0, 0) Then Call Crawler(dirs, arrName, reach + 1, arrName(cursor, 1), arrName(cursor, 2))
   If (dirs And D_MAP) = D_MAP Then _
      If Scout(arrName, row, col, D_MAP, 0, 0) Then Call Crawler(dirs, arrName, reach + 1, arrName(cursor, 1), arrName(cursor, 2))
End Function

Public Function Scout(ByRef arrName, row As Integer, col As Integer, map, rowOffset As Integer, colOffset As Integer) As Boolean
   Scout = False
   If LenB(aData(getIndex(row, col), cDATA)) = 0 Then Exit Function
   If (aData(getIndex(row, col), cDATA) And map) = 0 Then Exit Function
   Select Case (aData(getIndex(row, col), cDATA) And map)
   Case N_exit, E_exit, S_exit, W_exit, N_door, E_door, S_door, W_door, N_hiddendoor, E_hiddendoor, S_hiddendoor, W_hiddendoor
      If mismatch(arrName, row + rowOffset, col + colOffset) Then Exit Function
      cursor = cursor + 1
      arrName(cursor, 1) = row + rowOffset
      arrName(cursor, 2) = col + colOffset
   Case U_exit, U_door, U_portal, U_doorportal, U_hiddendoor, (U_hiddendoor Or U_portal)
      If mismatch(arrName, aData(getIndex(row, col), cUPORTALR), aData(getIndex(row, col), cUPORTALC)) Then Exit Function
      cursor = cursor + 1
      arrName(cursor, 1) = aData(getIndex(row, col), cUPORTALR)
      arrName(cursor, 2) = aData(getIndex(row, col), cUPORTALC)
   Case D_exit, D_door, D_portal, D_doorportal, D_hiddendoor, (D_hiddendoor Or D_portal)
      If mismatch(arrName, aData(getIndex(row, col), cDPORTALR), aData(getIndex(row, col), cDPORTALC)) Then Exit Function
      cursor = cursor + 1
      arrName(cursor, 1) = aData(getIndex(row, col), cDPORTALR)
      arrName(cursor, 2) = aData(getIndex(row, col), cDPORTALC)
   Case N_portal, N_doorportal, (N_hiddendoor Or N_portal)
      If mismatch(arrName, aData(getIndex(row, col), cNPORTALR), aData(getIndex(row, col), cNPORTALC)) Then Exit Function
      cursor = cursor + 1
      arrName(cursor, 1) = aData(getIndex(row, col), cNPORTALR)
      arrName(cursor, 2) = aData(getIndex(row, col), cNPORTALC)
   Case E_portal, E_doorportal, (E_hiddendoor Or E_portal)
      If mismatch(arrName, aData(getIndex(row, col), cEPORTALR), aData(getIndex(row, col), cEPORTALC)) Then Exit Function
      cursor = cursor + 1
      arrName(cursor, 1) = aData(getIndex(row, col), cEPORTALR)
      arrName(cursor, 2) = aData(getIndex(row, col), cEPORTALC)
   Case S_portal, S_doorportal, (S_hiddendoor Or S_portal)
      If mismatch(arrName, aData(getIndex(row, col), cSPORTALR), aData(getIndex(row, col), cSPORTALC)) Then Exit Function
      cursor = cursor + 1
      arrName(cursor, 1) = aData(getIndex(row, col), cSPORTALR)
      arrName(cursor, 2) = aData(getIndex(row, col), cSPORTALC)
   Case W_portal, W_doorportal, (W_hiddendoor Or W_portal)
      If mismatch(arrName, aData(getIndex(row, col), cWPORTALR), aData(getIndex(row, col), cWPORTALC)) Then Exit Function
      cursor = cursor + 1
      arrName(cursor, 1) = aData(getIndex(row, col), cWPORTALR)
      arrName(cursor, 2) = aData(getIndex(row, col), cWPORTALC)
   Case Else
      Scout = False
   End Select
   Scout = True
End Function

Public Function mismatch(ByRef arrName, row, col) As Boolean
Dim n As Integer
   mismatch = True
' check boundaries
   If isValid(row, col) = False Then Exit Function
' disgard entering coordinates twice
   For n = 1 To cursor
      If arrName(n, 1) = row Then
         If arrName(n, 2) = col Then Exit Function
      End If
   Next
   mismatch = False
End Function

Public Sub cmpWorldDesc(ByRef description As String)
'If DEBUGMODE = False Then On Error GoTo errorhandler
'errorData = errorData & "cmpWorldDesc -> "
   tmpMatch = 0
   tmpRoomDesc = encryptDesc(description)
   For cursor = 1 To theCount
      If LenB(aData(cursor, cDESCRIPTION)) = LenB(tmpRoomDesc) Then
         If aData(cursor, cDESCRIPTION) = tmpRoomDesc Then ' matches
            tmpMatch = tmpMatch + 1
            If tmpMatch > 1 Then Exit For ' no point checking further
            virtualRow = aData(cursor, cROW)
            virtualCol = aData(cursor, cCOL)
         End If
      End If
   Next
   
   Select Case tmpMatch
   Case 0
      Call SYNC_FALSE("")
   Case 1
      Call SYNC_TRUE("Updated.", cursor)
   Case Is > 1
      Call SYNC_FALSE("multiple matches.")
   End Select

Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "flee cmpWorldDesc"
   writeError (errorModule)
End Sub

Public Sub getSynced(ByRef description As String, Optional saveEncrypted As Boolean)
'If DEBUGMODE = False Then On Error GoTo errorhandler
'errorData = errorData & "cmpWorldDesc -> "
   tmpMatch = 0
   tmpRoomDesc = CRC32(description)
   If saveEncrypted Then encryptedDescription = tmpRoomDesc
   
   For cursor = 1 To theCount
      If aData(cursor, cDESCRIPTION) = tmpRoomDesc Then ' matches
         tmpMatch = tmpMatch + 1
         If tmpMatch > 1 Then Exit For ' no point checking further
         virtualRow = aData(cursor, cROW)
         virtualCol = aData(cursor, cCOL)
      End If
   Next
   Select Case tmpMatch
   Case 0
      Call SYNC_FALSE("room not found.")
   Case 1
      Call SYNC_TRUE("", cursor)
   Case Is > 1
      Call SYNC_FALSE("multiple matches.")
   End Select

Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "flee cmpWorldDesc"
   writeError (errorModule)
End Sub

