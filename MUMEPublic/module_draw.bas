Attribute VB_Name = "drawing"
Option Explicit
Option Compare Binary
Public mapRadius As Long
Public roomsize As Long
Public absX As Long
Public absY As Long
Public theCenter As Long
Public theMaximum As Long
Public boxX As Long
Public boxY As Long
Public half As Integer

Public Sub DrawMap()
Dim row As Long, col As Long
errorData = errorData & "DrawMap -> "
On Error GoTo 0 'errorhandler
   If Not (Initialized) Or frmMap.WindowState = vbMinimized Then Exit Sub 'is hidden or minimized
   BitBlt frmMap.hdc, 0, 0, frmMap.ScaleWidth, frmMap.ScaleHeight, 0, 0, 0, vbBlackness
   BitBlt frmMapBuffer.hdc, 0, 0, frmMapBuffer.ScaleWidth, frmMapBuffer.ScaleHeight, 0, 0, 0, vbBlackness
  
   mapRadius = 3
   If LOST = True And MappingMode = True Then
      mapRadius = Int(((frmMap.ScaleWidth + frmMap.ScaleHeight) / 4) / roomsize)
   End If
   
   If targetFound Then targetRow = theRow: targetCol = theCol
   For row = theRow - mapRadius To theRow + mapRadius
      For col = theCol - mapRadius To theCol + mapRadius
         If checkArrayLimit(row, col) Then
            absX = ((1 + col - (theCol - mapRadius)) * roomsize) - roomsize
            absY = ((1 + row - (theRow - mapRadius)) * roomsize) - roomsize
            If arrWorld(row, col) > 0 Then
               If drawTerrain(arrData(arrWorld(row, col), cDATA), absX, absY, row, col) = True Then
                  Call drawExit(arrData(arrWorld(row, col), cDATA), absX, absY)
                  Call drawFlag(arrData(arrWorld(row, col), cDATA), absX + 2, absY + 2)
                  If viewNotes Then Call drawNote(arrData(arrWorld(row, col), cDATA), absX, absY, row, col)
'#SE                  If viewTarget And row = targetRow And col = targetCol Then BitBlt frmMapBuffer.hdc, absX, absY, roomsize, roomsize, DCflagQuestion, 0, 0, vbSrcCopy
               End If
            Else
               If frmMap.mnuGridXY.Checked = True Then Call drawGridXY(absX, absY, row, col)
            End If
         End If
      Next
   Next
   
'#SE   If frmMap.mnuGroup.Checked = True Then Call drawGroup
   If frmMap.mnuGrid.Checked = True Then Call drawGrid
   If frmMap.mnuPortals.Checked = True Then Call checkPortal
   half = roomsize / 2
   theCenter = Int((mapRadius + 1) * roomsize - (half))
   theMaximum = (2 * mapRadius + 1) * roomsize
'#SE   If frmMap.mnuMovement.Checked = True Then Call drawMovement
'#SE   If viewPlayers Then Call drawPlayers
   If Len(lookfor) > 0 Then Call drawFind
   BitBlt frmMapBuffer.hdc, theCenter - (half), theCenter - (half), roomsize, roomsize, DCPlayerMask, 0, 0, vbSrcAnd
   BitBlt frmMapBuffer.hdc, theCenter - (half), theCenter - (half), roomsize, roomsize, DCPlayer, 0, 0, vbSrcPaint
   If viewDoornames Then Call drawDoornames
   BitBlt frmMap.hdc, _
         0, 0, _
         frmMap.ScaleWidth, frmMap.ScaleHeight, _
         frmMapBuffer.hdc, _
         ((2 * roomsize * mapRadius + roomsize) - frmMap.ScaleWidth) / 2, _
         ((2 * roomsize * mapRadius + roomsize) - frmMap.ScaleHeight) / 2, _
         vbSrcCopy
   frmMap.Refresh

Exit Sub
errorhandler:
   errorModule = Err.Description & "(" & Err.Number & ") -> " & "drawing DrawMap"
   writeError (errorModule)
End Sub

Public Function drawGridXY(x As Long, y As Long, row As Long, col As Long)
   Dim s As String
   s = row & vbCrLf & col
   Call myText(frmMapBuffer, s, x + roomsize / 2, y + roomsize / 3, False, 6)
End Function

Public Sub drawGrid()
   Dim jj As Integer
   For jj = theCenter - (1 + mapRadius * roomsize) - half To theCenter + (1 + mapRadius * roomsize) - half Step roomsize
      frmMapBuffer.Line (jj, theCenter - (mapRadius * roomsize))-(jj, theCenter + (mapRadius * roomsize)), RGB(200, 200, 200)
      frmMapBuffer.Line (theCenter - (mapRadius * roomsize), jj)-(theCenter + (mapRadius * roomsize), jj), RGB(200, 200, 200)
   Next jj
End Sub

Public Function drawNote(celldata, x As Long, y As Long, row As Long, col As Long)
   If Len(arrData(arrWorld(row, col), cNOTE)) > 0 Then
      Call myText(frmMapBuffer, arrData(arrWorld(row, col), cNOTE), x + roomsize / 2, y - 2 - roomsize / 3)
   End If
End Function

Public Sub checkPortal()
Dim row As Long, col As Long
errorData = errorData & "checkPortal -> "
   For row = theRow - mapRadius To theRow + mapRadius
      For col = theCol - mapRadius To theCol + mapRadius
         If checkArrayLimit(row, col) = True Then
            absX = ((1 + col - (theCol - mapRadius)) * roomsize) - roomsize
            absY = ((1 + row - (theRow - mapRadius)) * roomsize) - roomsize
            Dim celldata
            celldata = arrData(arrWorld(row, col), cDATA)
            If celldata > 0 Then
               If (celldata And N_MAP) >= N_portal Or _
                  (celldata And E_MAP) >= E_portal Or _
                  (celldata And S_MAP) >= S_portal Or _
                  (celldata And W_MAP) >= W_portal Or _
                  (celldata And U_MAP) >= U_portal Or _
                  (celldata And D_MAP) >= D_portal Then
                     If (celldata And N_MAP) >= N_portal Then Call drawPortal(row, col, N_MAP)
                     If (celldata And E_MAP) >= E_portal Then Call drawPortal(row, col, E_MAP)
                     If (celldata And S_MAP) >= S_portal Then Call drawPortal(row, col, S_MAP)
                     If (celldata And W_MAP) >= W_portal Then Call drawPortal(row, col, W_MAP)
                     If (celldata And U_MAP) >= U_portal Then Call drawPortal(row, col, U_MAP)
                     If (celldata And D_MAP) >= D_portal Then Call drawPortal(row, col, D_MAP)
               End If
            End If
         End If
      Next
   Next
End Sub

Public Sub drawPortal(row, col, theMap)
Dim tmpData
Dim tmpRoom
Dim tmpRow As Long
Dim tmpCol As Long
Dim startX, startY, middleX, middleY, targetX, targetY
   tmpData = arrData(arrWorld(row, col), cDATA)
   Select Case (tmpData And theMap)
   Case U_exit, U_door, U_portal, U_doorportal, U_hiddendoor, (U_hiddendoor Or U_portal)
     tmpRow = arrData(arrWorld(row, col), cUPORTALR)
     tmpCol = arrData(arrWorld(row, col), cUPORTALC)
     startX = ((1 + col - (theCol - mapRadius)) * roomsize) - roomsize + half
     startY = ((1 + row - (theRow - mapRadius)) * roomsize) - roomsize + half
     middleX = (1 + col - (theCol - mapRadius)) * roomsize ' - (roomsize)
     middleY = (1 + row - (theRow - mapRadius)) * roomsize - (roomsize)
     targetX = ((1 + tmpCol - (theCol - mapRadius)) * roomsize) - roomsize + half
     targetY = ((1 + tmpRow - (theRow - mapRadius)) * roomsize) - roomsize + half
   Case D_exit, D_door, D_portal, D_doorportal, D_hiddendoor, (D_hiddendoor Or D_portal)
     tmpRow = arrData(arrWorld(row, col), cDPORTALR)
     tmpCol = arrData(arrWorld(row, col), cDPORTALC)
     startX = ((1 + col - (theCol - mapRadius)) * roomsize) - roomsize + half
     startY = ((1 + row - (theRow - mapRadius)) * roomsize) - roomsize + half
     middleX = (1 + col - (theCol - mapRadius)) * roomsize - (roomsize)
     middleY = (1 + row - (theRow - mapRadius)) * roomsize
     targetX = ((1 + tmpCol - (theCol - mapRadius)) * roomsize) - roomsize + half
     targetY = ((1 + tmpRow - (theRow - mapRadius)) * roomsize) - roomsize + half
   Case N_exit, N_doorportal, N_portal, (N_hiddendoor Or N_portal)
     tmpRow = arrData(arrWorld(row, col), cNPORTALR)
     tmpCol = arrData(arrWorld(row, col), cNPORTALC)
     startX = ((1 + col - (theCol - mapRadius)) * roomsize) - roomsize + half
     startY = ((1 + row - (theRow - mapRadius)) * roomsize) - roomsize + half
     middleX = (1 + col - (theCol - mapRadius)) * roomsize - (half)
     middleY = (1 + row - (theRow - mapRadius)) * roomsize - (roomsize)
     targetX = ((1 + tmpCol - (theCol - mapRadius)) * roomsize) - roomsize + half
     targetY = ((1 + tmpRow - (theRow - mapRadius)) * roomsize) - roomsize + half
   Case E_exit, E_doorportal, E_portal, (E_hiddendoor Or E_portal)
     tmpRow = arrData(arrWorld(row, col), cEPORTALR)
     tmpCol = arrData(arrWorld(row, col), cEPORTALC)
     startX = ((1 + col - (theCol - mapRadius)) * roomsize) - roomsize + half
     startY = ((1 + row - (theRow - mapRadius)) * roomsize) - roomsize + half
     middleX = (1 + col - (theCol - mapRadius)) * roomsize
     middleY = (1 + row - (theRow - mapRadius)) * roomsize - (half)
     targetX = ((1 + tmpCol - (theCol - mapRadius)) * roomsize) - roomsize + half
     targetY = ((1 + tmpRow - (theRow - mapRadius)) * roomsize) - roomsize + half
   Case S_exit, S_doorportal, S_portal, (S_hiddendoor Or S_portal)
     tmpRow = arrData(arrWorld(row, col), cSPORTALR)
     tmpCol = arrData(arrWorld(row, col), cSPORTALC)
     startX = ((1 + col - (theCol - mapRadius)) * roomsize) - roomsize + half
     startY = ((1 + row - (theRow - mapRadius)) * roomsize) - roomsize + half
     middleX = (1 + col - (theCol - mapRadius)) * roomsize - (half)
     middleY = (1 + row - (theRow - mapRadius)) * roomsize
     targetX = ((1 + tmpCol - (theCol - mapRadius)) * roomsize) - roomsize + half
     targetY = ((1 + tmpRow - (theRow - mapRadius)) * roomsize) - roomsize + half
   Case W_exit, W_doorportal, W_portal, (W_hiddendoor Or W_portal)
     tmpRow = arrData(arrWorld(row, col), cWPORTALR)
     tmpCol = arrData(arrWorld(row, col), cWPORTALC)
     startX = ((1 + col - (theCol - mapRadius)) * roomsize) - roomsize + half
     startY = ((1 + row - (theRow - mapRadius)) * roomsize) - roomsize + half
     middleX = (1 + col - (theCol - mapRadius)) * roomsize - (roomsize)
     middleY = (1 + row - (theRow - mapRadius)) * roomsize - (half)
     targetX = ((1 + tmpCol - (theCol - mapRadius)) * roomsize) - roomsize + half
     targetY = ((1 + tmpRow - (theRow - mapRadius)) * roomsize) - roomsize + half
   End Select
   
   If tmpRow = 300 And tmpCol = 600 Then Exit Sub  'do not draw deathtrap portal
   frmMapBuffer.ForeColor = QBColor(15)
   Call myLine(frmMapBuffer.hdc, startX, startY, middleX, middleY)
   frmMapBuffer.ForeColor = QBColor(7)
   Call myLine(frmMapBuffer.hdc, middleX, middleY, targetX, targetY)
   'Call myLine(frmMapBuffer.hdc, startX, startY, targetX - 2, targetY + 2)
   frmMapBuffer.Circle (startX, startY), 2, QBColor(15)
   '( myLine(frmMapBuffer.hdc, startX, startY, targetX, targetY)
   'Call myLine(frmMapBuffer.hdc, startX, startY, middleX, middleY)
End Sub

Public Function drawMovement()
Dim n As Integer
Dim x As Integer
Dim y As Integer
   theCenter = (mapRadius + 1) * roomsize - (half)
   If roomcount > 0 Then
      For n = stackOUT To stackIN
   Select Case roomsize
   Case 14, 22
      x = theCenter + ((arrRoomstack(n, 2) - theCol) * roomsize) - Int(roomsize / 2)
      y = theCenter + ((arrRoomstack(n, 1) - theRow) * roomsize) - Int(roomsize / 2)
   Case 32
      x = theCenter + ((arrRoomstack(n, 2) - theCol) * roomsize) - Int(roomsize / 3)
      y = theCenter + ((arrRoomstack(n, 1) - theRow) * roomsize) - Int(roomsize / 3)
   End Select
         'x = theCenter + ((arrRoomstack(n, 2) - theCol) * roomsize)
         'y = theCenter + ((arrRoomstack(n, 1) - theRow) * roomsize)
         If x > theMaximum Or x < 1 Or y > theMaximum Or y < 1 Then
         Else
            BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCMoveMask, 0, 0, vbSrcAnd
            BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCMove, 0, 0, vbSrcPaint
            If n = stackIN Then
               BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCMoveMask, 0, 0, vbSrcAnd
               BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCMoveEnd, 0, 0, vbSrcPaint
            End If
         End If
      Next
   End If
End Function

Public Function drawTerrain(celldata, x As Long, y As Long, Optional ByVal row As Integer, Optional ByVal col As Integer)
Dim width As Long
Dim height As Long
   width = roomsize
   height = roomsize
   If celldata <= 0 Then
      drawTerrain = False
      Exit Function
   End If
   drawTerrain = True
   
   Select Case (celldata And SPECIAL_MAP)
   Case shop
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCShop, 0, 0, vbSrcCopy
      Exit Function
   Case guild
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCGuild, 0, 0, vbSrcCopy
      Exit Function
   Case inn
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCInn, 0, 0, vbSrcCopy
      Exit Function
   Case bridge
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCBridge, 0, 0, vbSrcCopy
      Exit Function
   Case city
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCCity, 0, 0, vbSrcCopy
      Exit Function
   Case dungeon
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCDungeon, 0, 0, vbSrcCopy
      Exit Function
   End Select
   
   Select Case (celldata And TERRAIN_MAP)
   Case road
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCRoad, 0, 0, vbSrcCopy
      If frmMap.mnuMap2.Checked Then
         If (celldata And N_MAP) > N_noexit And arrData(arrWorld(row - 1, col), cDATA) > 0 And ((arrData(arrWorld(row - 1, col), cDATA) And TERRAIN_MAP) = road Or (arrData(arrWorld(row - 1, col), cDATA) And TERRAIN_MAP) = underground) Then _
            BitBlt frmMapBuffer.hdc, x + 7, y, 8, 10, DCRoadV, 0, 0, vbSrcCopy
         If (celldata And E_MAP) > E_noexit And arrData(arrWorld(row, col + 1), cDATA) > 0 And ((arrData(arrWorld(row, col + 1), cDATA) And TERRAIN_MAP) = road Or (arrData(arrWorld(row, col + 1), cDATA) And TERRAIN_MAP) = underground) Then _
            BitBlt frmMapBuffer.hdc, x + 10, y + 7, 12, 8, DCRoadH, 0, 0, vbSrcCopy
         If (celldata And S_MAP) > S_noexit And arrData(arrWorld(row + 1, col), cDATA) > 0 And ((arrData(arrWorld(row + 1, col), cDATA) And TERRAIN_MAP) = road Or (arrData(arrWorld(row + 1, col), cDATA) And TERRAIN_MAP) = underground) Then _
            BitBlt frmMapBuffer.hdc, x + 7, y + 10, 8, 12, DCRoadV, 0, 0, vbSrcCopy
         If (celldata And W_MAP) > W_noexit And arrData(arrWorld(row, col - 1), cDATA) > 0 And ((arrData(arrWorld(row, col - 1), cDATA) And TERRAIN_MAP) = road Or (arrData(arrWorld(row, col - 1), cDATA) And TERRAIN_MAP) = underground) Then _
            BitBlt frmMapBuffer.hdc, x, y + 7, 8, 12, DCRoadH, 0, 0, vbSrcCopy
      End If
      If frmMap.mnuMap1.Checked Then
         If (celldata And N_MAP) > N_noexit And arrData(arrWorld(row - 1, col), cDATA) > 0 And ((arrData(arrWorld(row - 1, col), cDATA) And TERRAIN_MAP) = road Or (arrData(arrWorld(row - 1, col), cDATA) And TERRAIN_MAP) = underground) Then _
            BitBlt frmMapBuffer.hdc, x + 8, y, 16, 8, DCRoad, 8, 16, vbSrcCopy
         If (celldata And E_MAP) > E_noexit And arrData(arrWorld(row, col + 1), cDATA) > 0 And ((arrData(arrWorld(row, col + 1), cDATA) And TERRAIN_MAP) = road Or (arrData(arrWorld(row, col + 1), cDATA) And TERRAIN_MAP) = underground) Then _
            BitBlt frmMapBuffer.hdc, x + 24, y + 8, 8, 16, DCRoad, 8, 8, vbSrcCopy
         If (celldata And S_MAP) > S_noexit And arrData(arrWorld(row + 1, col), cDATA) > 0 And ((arrData(arrWorld(row + 1, col), cDATA) And TERRAIN_MAP) = road Or (arrData(arrWorld(row + 1, col), cDATA) And TERRAIN_MAP) = underground) Then _
            BitBlt frmMapBuffer.hdc, x + 8, y + 24, 16, 8, DCRoad, 8, 8, vbSrcCopy
         If (celldata And W_MAP) > W_noexit And arrData(arrWorld(row, col - 1), cDATA) > 0 And ((arrData(arrWorld(row, col - 1), cDATA) And TERRAIN_MAP) = road Or (arrData(arrWorld(row, col - 1), cDATA) And TERRAIN_MAP) = underground) Then _
            BitBlt frmMapBuffer.hdc, x, y + 8, 8, 16, DCRoad, 16, 8, vbSrcCopy
      End If
   Case plain
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCPlain, 0, 0, vbSrcCopy
   Case forest
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCForest, 0, 0, vbSrcCopy
      If False And frmMap.mnuMap1.Checked Then
         If (celldata And N_MAP) > N_noexit And arrData(arrWorld(row - 1, col), cDATA) > 0 And (arrData(arrWorld(row - 1, col), cDATA) And TERRAIN_MAP) = forest Then _
            BitBlt frmMapBuffer.hdc, x + 8, y, 16, 8, DCForest, 8, 16, vbSrcCopy
         If (celldata And E_MAP) > E_noexit And arrData(arrWorld(row, col + 1), cDATA) > 0 And (arrData(arrWorld(row, col + 1), cDATA) And TERRAIN_MAP) = forest Then _
            BitBlt frmMapBuffer.hdc, x + 24, y + 8, 8, 16, DCForest, 8, 8, vbSrcCopy
         If (celldata And S_MAP) > S_noexit And arrData(arrWorld(row + 1, col), cDATA) > 0 And (arrData(arrWorld(row + 1, col), cDATA) And TERRAIN_MAP) = forest Then _
            BitBlt frmMapBuffer.hdc, x + 8, y + 24, 16, 8, DCForest, 8, 8, vbSrcCopy
         If (celldata And W_MAP) > W_noexit And arrData(arrWorld(row, col - 1), cDATA) > 0 And (arrData(arrWorld(row, col - 1), cDATA) And TERRAIN_MAP) = forest Then _
            BitBlt frmMapBuffer.hdc, x, y + 8, 8, 16, DCForest, 16, 8, vbSrcCopy
      End If
   Case swamp
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCSwamp, 0, 0, vbSrcCopy
      If frmMap.mnuMap1.Checked Then
         If (celldata And N_MAP) > N_noexit And arrData(arrWorld(row - 1, col), cDATA) > 0 And (arrData(arrWorld(row - 1, col), cDATA) And TERRAIN_MAP) = swamp Then _
            BitBlt frmMapBuffer.hdc, x + 8, y, 16, 8, DCSwamp, 8, 16, vbSrcCopy
         If (celldata And E_MAP) > E_noexit And arrData(arrWorld(row, col + 1), cDATA) > 0 And (arrData(arrWorld(row, col + 1), cDATA) And TERRAIN_MAP) = swamp Then _
            BitBlt frmMapBuffer.hdc, x + 24, y + 8, 8, 16, DCSwamp, 8, 8, vbSrcCopy
         If (celldata And S_MAP) > S_noexit And arrData(arrWorld(row + 1, col), cDATA) > 0 And (arrData(arrWorld(row + 1, col), cDATA) And TERRAIN_MAP) = swamp Then _
            BitBlt frmMapBuffer.hdc, x + 8, y + 24, 16, 8, DCSwamp, 8, 8, vbSrcCopy
         If (celldata And W_MAP) > W_noexit And arrData(arrWorld(row, col - 1), cDATA) > 0 And (arrData(arrWorld(row, col - 1), cDATA) And TERRAIN_MAP) = swamp Then _
            BitBlt frmMapBuffer.hdc, x, y + 8, 8, 16, DCSwamp, 16, 8, vbSrcCopy
         If (celldata And N_MAP) > N_noexit And _
            (celldata And W_MAP) > W_noexit And _
            arrData(arrWorld(row - 1, col - 1), cDATA) > 0 And _
            (arrData(arrWorld(row - 1, col - 1), cDATA) And TERRAIN_MAP) = swamp _
         Then BitBlt frmMapBuffer.hdc, x, y, 16, 16, DCSwamp, 8, 8, vbSrcCopy
         
         If (celldata And N_MAP) > N_noexit And _
            (celldata And E_MAP) > E_noexit And _
            arrData(arrWorld(row - 1, col + 1), cDATA) > 0 And _
            (arrData(arrWorld(row - 1, col + 1), cDATA) And TERRAIN_MAP) = swamp _
         Then BitBlt frmMapBuffer.hdc, x + 16, y, 16, 16, DCSwamp, 8, 8, vbSrcCopy
         
         If (celldata And S_MAP) > S_noexit And _
            (celldata And W_MAP) > W_noexit And _
            arrData(arrWorld(row + 1, col - 1), cDATA) > 0 And _
            (arrData(arrWorld(row + 1, col - 1), cDATA) And TERRAIN_MAP) = swamp _
         Then BitBlt frmMapBuffer.hdc, x, y + 16, 16, 16, DCSwamp, 8, 8, vbSrcCopy
         
         If (celldata And S_MAP) > S_noexit And _
            (celldata And E_MAP) > E_noexit And _
            arrData(arrWorld(row + 1, col + 1), cDATA) > 0 And _
            (arrData(arrWorld(row + 1, col + 1), cDATA) And TERRAIN_MAP) = swamp _
         Then BitBlt frmMapBuffer.hdc, x + 16, y + 16, 16, 16, DCSwamp, 8, 8, vbSrcCopy
            
      End If
   Case hill
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCHill, 0, 0, vbSrcCopy
      If False And frmMap.mnuMap1.Checked Then
         If (celldata And N_MAP) > N_noexit And arrData(arrWorld(row - 1, col), cDATA) > 0 And (arrData(arrWorld(row - 1, col), cDATA) And TERRAIN_MAP) = hill Then _
            BitBlt frmMapBuffer.hdc, x + 8, y, 16, 8, DCHill, 8, 16, vbSrcCopy
         If (celldata And E_MAP) > E_noexit And arrData(arrWorld(row, col + 1), cDATA) > 0 And (arrData(arrWorld(row, col + 1), cDATA) And TERRAIN_MAP) = hill Then _
            BitBlt frmMapBuffer.hdc, x + 24, y + 8, 8, 16, DCHill, 8, 8, vbSrcCopy
         If (celldata And S_MAP) > S_noexit And arrData(arrWorld(row + 1, col), cDATA) > 0 And (arrData(arrWorld(row + 1, col), cDATA) And TERRAIN_MAP) = hill Then _
            BitBlt frmMapBuffer.hdc, x + 8, y + 24, 16, 8, DCHill, 8, 8, vbSrcCopy
         If (celldata And W_MAP) > W_noexit And arrData(arrWorld(row, col - 1), cDATA) > 0 And (arrData(arrWorld(row, col - 1), cDATA) And TERRAIN_MAP) = hill Then _
            BitBlt frmMapBuffer.hdc, x, y + 8, 8, 16, DCHill, 16, 8, vbSrcCopy
      End If
   Case underground
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCUnderground, 0, 0, vbSrcCopy
      If False And frmMap.mnuMap1.Checked Then
         If (celldata And N_MAP) > N_noexit And arrData(arrWorld(row - 1, col), cDATA) > 0 And (arrData(arrWorld(row - 1, col), cDATA) And TERRAIN_MAP) = underground Then _
            BitBlt frmMapBuffer.hdc, x + 8, y, 16, 8, DCUnderground, 8, 16, vbSrcCopy
         If (celldata And E_MAP) > E_noexit And arrData(arrWorld(row, col + 1), cDATA) > 0 And (arrData(arrWorld(row, col + 1), cDATA) And TERRAIN_MAP) = underground Then _
            BitBlt frmMapBuffer.hdc, x + 24, y + 8, 8, 16, DCUnderground, 8, 8, vbSrcCopy
         If (celldata And S_MAP) > S_noexit And arrData(arrWorld(row + 1, col), cDATA) > 0 And (arrData(arrWorld(row + 1, col), cDATA) And TERRAIN_MAP) = underground Then _
            BitBlt frmMapBuffer.hdc, x + 8, y + 24, 16, 8, DCUnderground, 8, 8, vbSrcCopy
         If (celldata And W_MAP) > W_noexit And arrData(arrWorld(row, col - 1), cDATA) > 0 And (arrData(arrWorld(row, col - 1), cDATA) And TERRAIN_MAP) = underground Then _
            BitBlt frmMapBuffer.hdc, x, y + 8, 8, 16, DCUnderground, 16, 8, vbSrcCopy
      End If
   Case water
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCWater, 0, 0, vbSrcCopy
      If frmMap.mnuMap1.Checked Then
         If (celldata And N_MAP) > N_noexit And arrData(arrWorld(row - 1, col), cDATA) > 0 And (arrData(arrWorld(row - 1, col), cDATA) And TERRAIN_MAP) = water Then _
            BitBlt frmMapBuffer.hdc, x + 8, y, 16, 8, DCWater, 8, 16, vbSrcCopy
         If (celldata And E_MAP) > E_noexit And arrData(arrWorld(row, col + 1), cDATA) > 0 And (arrData(arrWorld(row, col + 1), cDATA) And TERRAIN_MAP) = water Then _
            BitBlt frmMapBuffer.hdc, x + 24, y + 8, 8, 16, DCWater, 8, 8, vbSrcCopy
         If (celldata And S_MAP) > S_noexit And arrData(arrWorld(row + 1, col), cDATA) > 0 And (arrData(arrWorld(row + 1, col), cDATA) And TERRAIN_MAP) = water Then _
            BitBlt frmMapBuffer.hdc, x + 8, y + 24, 16, 8, DCWater, 8, 8, vbSrcCopy
         If (celldata And W_MAP) > W_noexit And arrData(arrWorld(row, col - 1), cDATA) > 0 And (arrData(arrWorld(row, col - 1), cDATA) And TERRAIN_MAP) = water Then _
            BitBlt frmMapBuffer.hdc, x, y + 8, 8, 16, DCWater, 16, 8, vbSrcCopy
         
         If (celldata And N_MAP) > N_noexit And _
            (celldata And W_MAP) > W_noexit And _
            arrData(arrWorld(row - 1, col - 1), cDATA) > 0 And _
            (arrData(arrWorld(row - 1, col - 1), cDATA) And TERRAIN_MAP) = water _
         Then BitBlt frmMapBuffer.hdc, x, y, 16, 16, DCWater, 8, 8, vbSrcCopy
         
         If (celldata And N_MAP) > N_noexit And _
            (celldata And E_MAP) > E_noexit And _
            arrData(arrWorld(row - 1, col + 1), cDATA) > 0 And _
            (arrData(arrWorld(row - 1, col + 1), cDATA) And TERRAIN_MAP) = water _
         Then BitBlt frmMapBuffer.hdc, x + 16, y, 16, 16, DCWater, 8, 8, vbSrcCopy
         
         If (celldata And S_MAP) > S_noexit And _
            (celldata And W_MAP) > W_noexit And _
            arrData(arrWorld(row + 1, col - 1), cDATA) > 0 And _
            (arrData(arrWorld(row + 1, col - 1), cDATA) And TERRAIN_MAP) = water _
         Then BitBlt frmMapBuffer.hdc, x, y + 16, 16, 16, DCWater, 8, 8, vbSrcCopy
         
         If (celldata And S_MAP) > S_noexit And _
            (celldata And E_MAP) > E_noexit And _
            arrData(arrWorld(row + 1, col + 1), cDATA) > 0 And _
            (arrData(arrWorld(row + 1, col + 1), cDATA) And TERRAIN_MAP) = water _
         Then BitBlt frmMapBuffer.hdc, x + 16, y + 16, 16, 16, DCWater, 8, 8, vbSrcCopy
            
      End If
   Case mountain
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCMountain, 0, 0, vbSrcCopy
      If False And frmMap.mnuMap1.Checked Then
         If (celldata And N_MAP) > N_noexit And arrData(arrWorld(row - 1, col), cDATA) > 0 And (arrData(arrWorld(row - 1, col), cDATA) And TERRAIN_MAP) = mountain Then _
            BitBlt frmMapBuffer.hdc, x + 8, y, 16, 8, DCMountain, 8, 16, vbSrcCopy
         If (celldata And E_MAP) > E_noexit And arrData(arrWorld(row, col + 1), cDATA) > 0 And (arrData(arrWorld(row, col + 1), cDATA) And TERRAIN_MAP) = mountain Then _
            BitBlt frmMapBuffer.hdc, x + 24, y + 8, 8, 16, DCMountain, 8, 8, vbSrcCopy
         If (celldata And S_MAP) > S_noexit And arrData(arrWorld(row + 1, col), cDATA) > 0 And (arrData(arrWorld(row + 1, col), cDATA) And TERRAIN_MAP) = mountain Then _
            BitBlt frmMapBuffer.hdc, x + 8, y + 24, 16, 8, DCMountain, 8, 8, vbSrcCopy
         If (celldata And W_MAP) > W_noexit And arrData(arrWorld(row, col - 1), cDATA) > 0 And (arrData(arrWorld(row, col - 1), cDATA) And TERRAIN_MAP) = road Then _
            BitBlt frmMapBuffer.hdc, x, y + 8, 8, 16, DCMountain, 16, 8, vbSrcCopy
      End If
   End Select
End Function

Public Function drawFlag(celldata, x As Long, y As Long)
   Select Case (celldata And FLAG_MAP)
   Case FLAG_WATER
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagWater, 0, 0, vbSrcCopy
   Case FLAG_ITEM
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagItem, 0, 0, vbSrcCopy
   Case FLAG_HERB
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagHerb, 0, 0, vbSrcCopy
   Case FLAG_TREASURY
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagTreasury, 0, 0, vbSrcCopy
   Case FLAG_KEY
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagKey, 0, 0, vbSrcCopy
   Case FLAG_MAGIC
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagMagic, 0, 0, vbSrcCopy
   Case FLAG_MUDLLE
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagMudlle, 0, 0, vbSrcCopy
   Case FLAG_QUEST
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagQuest, 0, 0, vbSrcCopy
   Case FLAG_QUESTION
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagQuestion, 0, 0, vbSrcCopy
   End Select
End Function

Public Sub drawWall(row As Long, col As Long, x, y)
   If checkArrayLimit(row - 1, col) Then _
      If arrData(arrWorld(row - 1, col), cDATA) > 0 Then _
         BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
   
   If checkArrayLimit(row, col + 1) = True Then _
      If arrData(arrWorld(row, col + 1), cDATA) > 0 Then _
         BitBlt frmMapBuffer.hdc, x + roomsize - 1, y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
   
   If checkArrayLimit(row + 1, col) = True Then _
      If arrData(arrWorld(row + 1, col), cDATA) > 0 Then _
         BitBlt frmMapBuffer.hdc, x, y + roomsize - 1, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
   
   If checkArrayLimit(row, col - 1) = True Then _
      If arrData(arrWorld(row, col - 1), cDATA) > 0 Then _
         BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
End Sub

Public Sub drawExit(celldata, x As Long, y As Long)
   If (celldata And ROOM_MAP) = noRide_Dark Or (celldata And ROOM_MAP) = noRide_Sun Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCnoRideMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCnoRide, 0, 0, vbSrcPaint
   End If

   If (celldata And MONSTER_MAP) = MONSTER_MAP Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCMonsterMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCMonster, 0, 0, vbSrcPaint
   End If

   If (celldata And ROOM_MAP) = noRide_Dark Or (celldata And ROOM_MAP) = Ride_Dark Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCDarkMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCDark, 0, 0, vbSrcPaint
   End If
   
   If (celldata And N_MAP) = N_noexit Then
      BitBlt frmMapBuffer.hdc, x, y + 0, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + 1, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_noexit Then
      BitBlt frmMapBuffer.hdc, x + roomsize - 1, y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + roomsize - 2, y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
   End If
   If (celldata And S_MAP) = S_noexit Then
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 1, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 2, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_noexit Then
      BitBlt frmMapBuffer.hdc, x + 0, y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + 1, y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
   End If
'PORTAL
   If (celldata And N_MAP) = N_portal Then
      BitBlt frmMapBuffer.hdc, x, y + 0, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
'      BitBlt frmMapBuffer.hdc, x, y + 1, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_portal Then
      BitBlt frmMapBuffer.hdc, x + roomsize - 1, y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
'      BitBlt frmMapBuffer.hdc, x + roomsize - 2, y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
   End If
   If (celldata And S_MAP) = S_portal Then
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 1, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
'      BitBlt frmMapBuffer.hdc, x, y + roomsize - 2, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_portal Then
      BitBlt frmMapBuffer.hdc, x + 0, y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
'      BitBlt frmMapBuffer.hdc, x + 1, y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
   End If
'DOOR
   If (celldata And N_MAP) = N_door Then
      BitBlt frmMapBuffer.hdc, x, y + 0, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + 1, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_door Then
      BitBlt frmMapBuffer.hdc, x + roomsize - 1, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + roomsize - 2, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And S_MAP) = S_door Then
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 1, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 2, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_door Then
      BitBlt frmMapBuffer.hdc, x + 0, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + 1, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
   End If
'HIDDENDOOR
   If (celldata And S_MAP) = S_hiddendoor Then
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 1, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 2, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 3, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_hiddendoor Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + 1, y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + 2, y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy
   End If
   If (celldata And N_MAP) = N_hiddendoor Then
      BitBlt frmMapBuffer.hdc, x, y + 0, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + 1, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + 2, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_hiddendoor Then
      BitBlt frmMapBuffer.hdc, x + roomsize - 1, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + roomsize - 2, y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + roomsize - 3, y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy
   End If
'DOORPORTAL
   If (celldata And S_MAP) = S_doorportal Then
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 1, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 2, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 3, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_doorportal Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + 1, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + 2, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And N_MAP) = N_doorportal Then
      BitBlt frmMapBuffer.hdc, x, y + 0, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + 1, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + 2, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_doorportal Then
      BitBlt frmMapBuffer.hdc, x + roomsize - 1, y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + roomsize - 2, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + roomsize - 3, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
   End If
'HIDDENDOORPORTAL
   If (celldata And S_MAP) = (S_hiddendoor Or S_portal) Then
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 1, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 2, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + roomsize - 3, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = (W_hiddendoor Or W_portal) Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + 1, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + 2, y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy
   End If
   If (celldata And N_MAP) = (N_hiddendoor Or N_portal) Then
      BitBlt frmMapBuffer.hdc, x, y + 0, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + 1, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x, y + 2, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = (E_hiddendoor Or E_portal) Then
      BitBlt frmMapBuffer.hdc, x + roomsize - 1, y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + roomsize - 2, y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, x + roomsize - 3, y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy
   End If

   If (celldata And U_MAP) = U_exit Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCUpMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCUp, 0, 0, vbSrcPaint
   End If
   If (celldata And U_MAP) = U_portal Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCUpMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCUpPortal, 0, 0, vbSrcPaint
   End If
   If (celldata And U_MAP) = U_door Or (celldata And U_MAP) = U_doorportal Or (celldata And U_MAP) = U_hiddendoor Or (celldata And U_MAP) = U_MAP Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCUpMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCUpDoor, 0, 0, vbSrcPaint
   End If

   If (celldata And D_MAP) = D_exit Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCDownMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCDown, 0, 0, vbSrcPaint
   End If
   If (celldata And D_MAP) = D_portal Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCDownMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCDownPortal, 0, 0, vbSrcPaint
   End If
   If (celldata And D_MAP) = D_door Or (celldata And D_MAP) = D_doorportal Or (celldata And D_MAP) = D_hiddendoor Or (celldata And D_MAP) = D_MAP Then
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCDownMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCDownDoor, 0, 0, vbSrcPaint
   End If
Exit Sub
errorhandler:
  theRow = 15
  theCol = 15
  Resume Next
End Sub

Public Function drawDoornames()
Dim n, m
errorData = errorData & "drawDoornames -> "
   If roomsize > 30 Then
      n = 2
      m = 1
   Else
      n = 3
      m = 2
   End If
   If Len(theDoornameUp) > 0 Then Call myText(frmMapBuffer, theDoornameUp, theCenter + 4, theCenter - (n * roomsize), True, 8)
   If Len(theDoornameNorth) > 0 Then Call myText(frmMapBuffer, theDoornameNorth, theCenter + 4, theCenter - (m * roomsize), True, 8)
   If Len(theDoornameEast) > 0 Then Call myText(frmMapBuffer, theDoornameEast, theCenter + (n * roomsize), theCenter - 2, True, 8)
   If Len(theDoornameWest) > 0 Then Call myText(frmMapBuffer, theDoornameWest, theCenter - (n * roomsize), theCenter - 2, True, 8)
   If Len(theDoornameSouth) > 0 Then Call myText(frmMapBuffer, theDoornameSouth, theCenter + 4, theCenter + (m * roomsize), True, 8)
   If Len(theDoornameDown) > 0 Then Call myText(frmMapBuffer, theDoornameDown, theCenter + 4, theCenter + (n * roomsize), True, 8)
End Function

Public Sub drawPlayers()
Dim x As Integer
Dim y As Integer
Dim cursor As Integer
Dim n As Integer
Dim s As String
   
   If viewPlayers = False Then Exit Sub
   For n = LBound(arrPlayers, 1) To arrPlayersIndex - 1
      For cursor = 1 To theCount
         If StrComp(arrData(cursor, cROOMNAME), arrPlayers(n), vbTextCompare) = 0 Then
            x = theCenter + ((arrData(cursor, cCOL) - theCol) * roomsize) '- half
            y = theCenter + ((arrData(cursor, cROW) - theRow) * roomsize) '- half
            If x >= theMaximum Or x <= 1 Or y >= theMaximum Or y <= 1 Then
               'out of area, skip drawing
            Else
'               BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCMoveMask, 0, 0, vbSrcAnd
'               BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCMove, 0, 0, vbSrcPaint
               Call myText(frmMapBuffer, arrPlayersNames(n), x, y)
            End If
         End If
      Next
   Next
   viewPlayers = False
   
'   informClient ("")
'   For n = LBound(arrPlayers, 1) To arrPlayersIndex - 1
'      informClient (CStr(n + 1) & ") " & arrPlayers(n))
'   Next
'   informClient ("")
End Sub

Public Sub drawFind()
Dim cursor As Integer, s As String
Dim x As Integer, y As Integer
   For cursor = 1 To theCount
      If InStr(1, LCase(arrData(cursor, cROOMNAME)), lookfor, vbTextCompare) Then
         s = "(" & arrData(cursor, cROW) & "," & arrData(cursor, cCOL) & ")"
         s = s & Space(10 - Len(s))
         Call informClient(s & arrData(cursor, cROOMNAME) & " - " & _
            "N[" & arrData(cursor, cNDOOR) & "], " & _
            "E[" & arrData(cursor, cEDOOR) & "], " & _
            "S[" & arrData(cursor, cSDOOR) & "], " & _
            "W[" & arrData(cursor, cWDOOR) & "], " & _
            "U[" & arrData(cursor, cUDOOR) & "], " & _
            "D[" & arrData(cursor, cDDOOR) & "]", True)

'         x = theCenter + ((arrData(cursor, cCOL) - theCol) * roomsize) '- half
'         y = theCenter + ((arrData(cursor, cROW) - theRow) * roomsize) '- half
'         If x > theMaximum Or x < 1 Or y > theMaximum Or y < 1 Then
'            'out of area, skip drawing
'         Else
'            Call myText(frmMapBuffer, "@", x, y, , 8)
'         End If
      End If
   Next
   s = "[" & arrData(arrWorld(theRow, theCol), cROW) & "," & arrData(arrWorld(theRow, theCol), cCOL) & "]"
   s = s & Space(10 - Len(s))
   Call informClient(s & "Current coordinates.", True)
   lookfor = ""
End Sub
