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
Public timer As New cHiResTimer

Public Sub DrawMap()
Dim row As Integer, col As Integer
Dim cell As Long
errorData = errorData & "DrawMap -> "
   If frmMap.WindowState = vbMinimized Then Exit Sub
   If Not (Initialized) Then Exit Sub
   If frmMap.mnuPortals.Checked Then
      frmMapBuffer.Cls
      frmMap.Cls
   Else
      frmMapBuffer.Cls
   End If
   mapRadius = Int(((frmMap.ScaleWidth + frmMap.ScaleHeight) / 4) / roomsize)
   half = roomsize / 2
   '#target    If targetFound Then targetRow = theROW: targetCol = theCOL
   
   For row = theROW - mapRadius To theROW + mapRadius
      For col = theCOL - mapRadius To theCOL + mapRadius
         If isValid(row, col) Then
            absX = ((1 + col - (theCOL - mapRadius)) * roomsize) - roomsize
            absY = ((1 + row - (theROW - mapRadius)) * roomsize) - roomsize
            cell = getLng(aData(getInt(aWorld(row, col, theLEVEL)), cDATA))
            If cell <> 0 Then 'kui samas maailmas on ruum
               Call drawTerrain(cell, absX, absY, row, col)
               Call drawExit(cell, absX, absY)
               Call drawFlag(cell, absX + 2, absY + 2)
               If viewNotes Then Call drawNote(cell, absX, absY, row, col)
               If row = flagRow And col = flagCol Then BitBlt frmMapBuffer.hdc, absX + half, absY + half, roomsize, roomsize, DCflagQuestion, 0, 0, vbSrcCopy
               'If viewTarget And row = targetRow And col = targetCol Then BitBlt frmMapBuffer.hdc, absX, absY, roomsize, roomsize, DCflagQuestion, 0, 0, vbSrcCopy
            Else
               If frmMap.mnuGridXY.Checked = True Then Call drawGridXY(absX, absY, row, col)
            End If
         End If
      Next
   Next
   
'   If frmMap.mnuGroup.Checked = True Then Call drawGroup
   If frmMap.mnuGrid.Checked = True Then Call drawGrid
   
   theCenter = Int((mapRadius + 1) * roomsize - (half))
   theMaximum = (2 * mapRadius + 1) * roomsize
   If frmMap.mnuMovement.Checked Then Call drawMovement
   
   If frmMap.mnuPlayers.Checked Then
      Call drawPlayers
   End If
   
   If frmMap.mnuNotes.Checked And LenB(lookfor) <> 0 Then Call drawFind
   If frmMap.mnuDoornames.Checked Then Call drawDoornames ' viewDoornames
   
   BitBlt frmMapBuffer.hdc, theCenter - (half), theCenter - (half), roomsize, roomsize, DCPlayerMask, 0, 0, vbSrcAnd
   BitBlt frmMapBuffer.hdc, theCenter - (half), theCenter - (half), roomsize, roomsize, DCPlayer, 0, 0, vbSrcPaint
   
   'checkportali sees kontrollitakse joonistamist
   Call checkPortal
   
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
   errorModule = Err.description & "(" & Err.Number & ") -> " & "drawing DrawMap"
   writeError (errorModule)
End Sub

Public Function drawGridXY(X As Long, Y As Long, row As Integer, col As Integer)
   Dim s As String
   s = row & vbCrLf & col
   Call myText(frmMapBuffer, s, X + roomsize / 2, Y + roomsize / 3, False, 6)
End Function

Public Sub drawGrid()
   Dim jj As Integer
   For jj = theCenter - (1 + mapRadius * roomsize) - half To theCenter + (1 + mapRadius * roomsize) - half Step roomsize
      frmMapBuffer.Line (jj, theCenter - (mapRadius * roomsize))-(jj, theCenter + (mapRadius * roomsize)), RGB(50, 50, 50)
      frmMapBuffer.Line (theCenter - (mapRadius * roomsize), jj)-(theCenter + (mapRadius * roomsize), jj), RGB(50, 50, 50)
   Next jj
End Sub

Public Function drawNote(ByRef celldata As Long, X As Long, Y As Long, row As Integer, col As Integer)
   If LenB(aData(getIndex(row, col), cNOTE)) <> 0 Then
      Dim note
      note = Split(aData(getIndex(row, col), cNOTE), "|", , vbBinaryCompare)
      Call myText(frmMapBuffer, note(0), X + roomsize / 2, Y - 2 - roomsize / 3)
   End If
End Function

Public Sub checkPortal()
Dim row As Integer, col As Integer
errorData = errorData & "checkPortal -> "
   staticLevel = True
   For row = theROW - mapRadius To theROW + mapRadius
      For col = theCOL - mapRadius To theCOL + mapRadius
         If isValid(row, col) = True Then
            If getIndex(row, col) <> 0 Then
               Call drawPortal("n", row, col, getInt(aData(getIndex(row, col), cNPORTALR)), getInt(aData(getIndex(row, col), cNPORTALC)), -half, -roomsize)
               Call drawPortal("e", row, col, getInt(aData(getIndex(row, col), cEPORTALR)), getInt(aData(getIndex(row, col), cEPORTALC)), 0, -half)
               Call drawPortal("s", row, col, getInt(aData(getIndex(row, col), cSPORTALR)), getInt(aData(getIndex(row, col), cSPORTALC)), -half, 0)
               Call drawPortal("w", row, col, getInt(aData(getIndex(row, col), cWPORTALR)), getInt(aData(getIndex(row, col), cWPORTALC)), -roomsize, -half)
               Call drawPortal("u", row, col, getInt(aData(getIndex(row, col), cUPORTALR)), getInt(aData(getIndex(row, col), cUPORTALC)), 0, -roomsize)
               Call drawPortal("d", row, col, getInt(aData(getIndex(row, col), cDPORTALR)), getInt(aData(getIndex(row, col), cDPORTALC)), -roomsize, 0)
'               If (getData(getIndex(row, col)) And N_MAP) >= N_portal Then Call drawPortal(row, col, N_MAP)
'               If (getData(getIndex(row, col)) And E_MAP) >= E_portal Then Call drawPortal(row, col, E_MAP)
'               If (getData(getIndex(row, col)) And S_MAP) >= S_portal Then Call drawPortal(row, col, S_MAP)
'               If (getData(getIndex(row, col)) And W_MAP) >= W_portal Then Call drawPortal(row, col, W_MAP)
'               If (getData(getIndex(row, col)) And U_MAP) >= U_portal Then Call drawPortal(row, col, U_MAP)
'               If (getData(getIndex(row, col)) And D_MAP) >= D_portal Then Call drawPortal(row, col, D_MAP)
            End If
         End If
      Next
   Next
   staticLevel = False
End Sub

'Public Sub drawPortal(row As Integer, col As Integer, theMap As Long)
Public Sub drawPortal(direction As String, row As Integer, col As Integer, tmpRow As Integer, tmpCol As Integer, oX As Integer, oY As Integer)
Dim tmpData As Long
'Dim tmpRow As Integer: Dim tmpCol As Integer:
'Dim r As Integer: Dim c As Integer
'Dim oX As Integer: Dim oY As Integer
Dim startX As Integer: Dim startY As Integer: Dim middleX As Integer: Dim middleY As Integer: Dim targetX As Integer: Dim targetY As Integer
Dim portalWest As Boolean
portalWest = False
   
   If tmpRow = 0 Or tmpCol = 0 Then Exit Sub 'kui portal puudub
      If direction = "u" Or direction = "d" Then
         'always draw portal up and down
      Else
         If (Abs(row - tmpRow) = 1 And Abs(col - tmpCol) = 0) Or (Abs(row - tmpRow) = 0 And Abs(col - tmpCol) = 1) Then Exit Sub 'kui ruumid on kõrvuti
      End If
   
   'portaali joonistamist alustatakse PAREMALT ALT NURGAST VALITUD RUUMI NURGAST
'   Select Case (aData(getIndex(row, col), cDATA) And theMap)
'   Case U_exit, U_door, U_portal, U_doorportal, U_hiddendoor, (U_hiddendoor Or U_portal):
'      r = cUPORTALR: c = cUPORTALC: oX = 0: oY = -roomsize
'   Case D_exit, D_door, D_portal, D_doorportal, D_hiddendoor, (D_hiddendoor Or D_portal):
'      r = cDPORTALR: c = cDPORTALC: oX = -roomsize: oY = 0
'   Case N_exit, N_doorportal, N_portal, (N_hiddendoor Or N_portal):
'      r = cNPORTALR: c = cNPORTALC: oX = -half: oY = -roomsize
'   Case E_exit, E_doorportal, E_portal, (E_hiddendoor Or E_portal):
'      r = cEPORTALR: c = cEPORTALC: oX = 0: oY = -half
'   Case S_exit, S_doorportal, S_portal, (S_hiddendoor Or S_portal):
'      r = cSPORTALR: c = cSPORTALC: oX = -half: oY = 0
'   Case W_exit, W_doorportal, W_portal, (W_hiddendoor Or W_portal): portalWest = True:
'      r = cWPORTALR: c = cWPORTALC: oX = -roomsize: oY = -half
'   End Select
'   tmpRow = aData(getIndex(row, col), r)
'   tmpCol = aData(getIndex(row, col), c)
   
   startX = ((1 + col - (theCOL - mapRadius)) * roomsize) - roomsize + half
   startY = ((1 + row - (theROW - mapRadius)) * roomsize) - roomsize + half
   middleX = ((1 + col - (theCOL - mapRadius)) * roomsize) + oX
   middleY = ((1 + row - (theROW - mapRadius)) * roomsize) + oY
   targetX = ((1 + tmpCol - (theCOL - mapRadius)) * roomsize) - roomsize + half
   targetY = ((1 + tmpRow - (theROW - mapRadius)) * roomsize) - roomsize + half
   
   If (theROW = row And theCOL = col) Then
      frmMapBuffer.ForeColor = QBColor(15)
      Call myLine(frmMapBuffer.hdc, startX, startY, middleX, middleY)
      Call myLine(frmMapBuffer.hdc, middleX, middleY, targetX, targetY)
      Call myLine(frmMapBuffer.hdc, middleX - 1, middleY, targetX - 1, targetY)
   Else
      If (portalWest And tmpRow = theROW And tmpCol = theCOL) Then
         frmMapBuffer.ForeColor = QBColor(15)
      Else
         frmMapBuffer.ForeColor = QBColor(0)
      End If
      If frmMap.mnuPortals.Checked Then
         If frmMap.mnuFancy.Checked Then Call myLine(frmMapBuffer.hdc, startX, startY, middleX, middleY)
         Call myLine(frmMapBuffer.hdc, startX, startY, middleX, middleY)
         Call myLine(frmMapBuffer.hdc, middleX, middleY, targetX, targetY)
      End If
   End If
   If frmMap.mnuPortals.Checked Then
      If frmMap.mnuFancy.Checked Then frmMapBuffer.Circle (startX, startY), 2, QBColor(15)
   End If
End Sub

Public Function drawMovement()
Dim n As Integer
Dim X As Integer
Dim Y As Integer
   theCenter = (mapRadius + 1) * roomsize - (half)
   If roomcount > 0 Then
      For n = stackOUT To stackIN
         Select Case roomsize
         Case 14, 22, 32
            X = theCenter + ((arrRoomstack(n, 2) - theCOL) * roomsize) - Int(roomsize / 2)
            Y = theCenter + ((arrRoomstack(n, 1) - theROW) * roomsize) - Int(roomsize / 2)
         'Case 32
            'x = theCenter + ((arrRoomstack(n, 2) - theCOL) * roomsize) - Int(roomsize / 3)
            'y = theCenter + ((arrRoomstack(n, 1) - theROW) * roomsize) - Int(roomsize / 3)
         End Select
         'x = theCenter + ((arrRoomstack(n, 2) - theCOL) * roomsize)
         'y = theCenter + ((arrRoomstack(n, 1) - theROW) * roomsize)
'#         If x > theMaximum Or x < 1 Or y > theMaximum Or y < 1 Then
'         Else
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMoveMask, 0, 0, vbSrcAnd
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMove, 0, 0, vbSrcPaint
            If n = stackIN Then
               BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMoveMask, 0, 0, vbSrcAnd
               BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMoveEnd, 0, 0, vbSrcPaint
            End If
 '        End If
      Next
   End If
End Function

Public Function drawTerrain(ByRef celldata As Long, X As Long, Y As Long, Optional ByRef row As Integer, Optional ByRef col As Integer)
   Dim exploration As Boolean
   exploration = False
   drawTerrain = True
   
   If MappingMode Then
      'If MUDname = "WARP" And frmMap.mnuWalkthrough.Checked Then
      If frmMap.mnuWalkthrough.Checked Then
         If LenB(aData(getIndex(row, col), cDESCRIPTION)) = 0 Or LenB(aData(getIndex(row, col), cDESCRIPTION)) > 20 Then exploration = True
      End If
   End If
   
   Dim is_dark_room As Boolean
   If (celldata And ROOM_MAP) = noRide_Dark Or (celldata And ROOM_MAP) = Ride_Dark Then is_dark_room = True
   
   If frmMap.mnuMap1.Checked Then
      Select Case (celldata And TERRAIN_MAP)
      Case road, plain:
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCPlain, 0, 0, vbSrcCopy
      Case swamp:
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCSwamp, 0, 0, vbSrcCopy
      Case water:
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCWater, 0, 0, vbSrcCopy
         If LenB(aData(getIndex(row - 1, col), cDATA)) <> 0 Then If (celldata And N_MAP) > N_noexit And (aData(getIndex(row - 1, col), cDATA) And TERRAIN_MAP) = water Then BitBlt frmMapBuffer.hdc, X + 8, Y, 16, 8, DCWater, 8, 16, vbSrcCopy
         If LenB(aData(getIndex(row, col + 1), cDATA)) <> 0 Then If (celldata And E_MAP) > E_noexit And (aData(getIndex(row, col + 1), cDATA) And TERRAIN_MAP) = water Then BitBlt frmMapBuffer.hdc, X + 24, Y + 8, 8, 16, DCWater, 8, 8, vbSrcCopy
         If LenB(aData(getIndex(row + 1, col), cDATA)) <> 0 Then If (celldata And S_MAP) > S_noexit And (aData(getIndex(row + 1, col), cDATA) And TERRAIN_MAP) = water Then BitBlt frmMapBuffer.hdc, X + 8, Y + 24, 16, 8, DCWater, 8, 8, vbSrcCopy
         If LenB(aData(getIndex(row, col - 1), cDATA)) <> 0 Then If (celldata And W_MAP) > W_noexit And (aData(getIndex(row, col - 1), cDATA) And TERRAIN_MAP) = water Then BitBlt frmMapBuffer.hdc, X, Y + 8, 8, 16, DCWater, 16, 8, vbSrcCopy
         If LenB(aData(getIndex(row - 1, col - 1), cDATA)) <> 0 Then If (celldata And N_MAP) > N_noexit And (celldata And W_MAP) > W_noexit And (aData(getIndex(row - 1, col - 1), cDATA) And TERRAIN_MAP) = water Then BitBlt frmMapBuffer.hdc, X, Y, 16, 16, DCWater, 8, 8, vbSrcCopy
         If LenB(aData(getIndex(row - 1, col + 1), cDATA)) <> 0 Then If (celldata And N_MAP) > N_noexit And (celldata And E_MAP) > E_noexit And (aData(getIndex(row - 1, col + 1), cDATA) And TERRAIN_MAP) = water Then BitBlt frmMapBuffer.hdc, X + 16, Y, 16, 16, DCWater, 8, 8, vbSrcCopy
         If LenB(aData(getIndex(row + 1, col - 1), cDATA)) <> 0 Then If (celldata And S_MAP) > S_noexit And (celldata And W_MAP) > W_noexit And (aData(getIndex(row + 1, col - 1), cDATA) And TERRAIN_MAP) = water Then BitBlt frmMapBuffer.hdc, X, Y + 16, 16, 16, DCWater, 8, 8, vbSrcCopy
         If LenB(aData(getIndex(row + 1, col + 1), cDATA)) <> 0 Then If (celldata And S_MAP) > S_noexit And (celldata And E_MAP) > E_noexit And (aData(getIndex(row + 1, col + 1), cDATA) And TERRAIN_MAP) = water Then BitBlt frmMapBuffer.hdc, X + 16, Y + 16, 16, 16, DCWater, 8, 8, vbSrcCopy
      Case forest:      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCForest, 0, 0, vbSrcCopy
      Case hill:        BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCHill, 0, 0, vbSrcCopy
      Case mountain:    BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMountain, 0, 0, vbSrcCopy
      Case underground: BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUnderground, 0, 0, vbSrcCopy
      End Select
   End If
   
   If frmMap.mnuMap2.Checked Then
      Select Case (celldata And TERRAIN_MAP)
      Case road, plain:
         If is_dark_room Then
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUnderground, 0, 0, vbSrcCopy
         Else
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCPlain, 0, 0, vbSrcCopy
         End If
      Case forest:
         If is_dark_room Then
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCForestDark, 0, 0, vbSrcCopy
         Else
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCForest, 0, 0, vbSrcCopy
         End If
      Case swamp:
         If is_dark_room Then
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCSwampDark, 0, 0, vbSrcCopy
         Else
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCSwamp, 0, 0, vbSrcCopy
         End If
      Case hill:
         If is_dark_room Then
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCHillDark, 0, 0, vbSrcCopy
         Else
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCHill, 0, 0, vbSrcCopy
         End If
      Case water
         If is_dark_room Then
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCWaterDark, 0, 0, vbSrcCopy
         Else
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCWater, 0, 0, vbSrcCopy
         End If
      Case mountain:
         If is_dark_room Then
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMountainDark, 0, 0, vbSrcCopy
         Else
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMountain, 0, 0, vbSrcCopy
         End If
      Case water
         If is_dark_room Then
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMountainDark, 0, 0, vbSrcCopy
         Else
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMountain, 0, 0, vbSrcCopy
         End If
      Case underground:
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUnderground, 0, 0, vbSrcCopy
      End Select
   End If

   If frmMap.mnuMap3.Checked Then
      Select Case (celldata And TERRAIN_MAP)
      Case road:        BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCRoad, 0, 0, vbSrcCopy
      Case plain:       BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCPlain, 0, 0, vbSrcCopy
      Case forest:      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCForest, 0, 0, vbSrcCopy
      Case swamp:       BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCSwamp, 0, 0, vbSrcCopy
      Case hill:        BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCHill, 0, 0, vbSrcCopy
      Case water:       BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCWater, 0, 0, vbSrcCopy
      Case mountain:    BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMountain, 0, 0, vbSrcCopy
      Case underground: BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUnderground, 0, 0, vbSrcCopy
      End Select
   End If


'DRAWING ROAD OVER TERRAIN
   If (celldata And ISROAD) = ISROAD Then
      Dim c As Integer
      If frmMap.mnuMap1.Checked Then
         BitBlt frmMapBuffer.hdc, X + 11, Y + 11, 10, 10, DCRoadMask, 11, 11, vbSrcAnd
         BitBlt frmMapBuffer.hdc, X + 11, Y + 11, 10, 10, DCRoad, 11, 11, vbSrcPaint
         c = getIndex(row, col, "n")
         If ISROAD = (getData(c) And ISROAD) Or (getData(c) And TERRAIN_MAP) = underground Then
            BitBlt frmMapBuffer.hdc, X + 11, Y, 10, 11, DCRoadMask, 11, 0, vbSrcAnd
            BitBlt frmMapBuffer.hdc, X + 11, Y, 10, 11, DCRoad, 11, 0, vbSrcPaint
         End If
         c = getIndex(row, col, "s")
         If ISROAD = (getData(c) And ISROAD) Or (getData(c) And TERRAIN_MAP) = underground Then
            BitBlt frmMapBuffer.hdc, X + 11, Y + 21, 10, 11, DCRoadMask, 11, 21, vbSrcAnd
            BitBlt frmMapBuffer.hdc, X + 11, Y + 21, 10, 11, DCRoad, 11, 21, vbSrcPaint
         End If
         
         c = getIndex(row, col, "w")
         If ISROAD = (getData(c) And ISROAD) Or (getData(c) And TERRAIN_MAP) = underground Then
            BitBlt frmMapBuffer.hdc, X, Y + 11, 11, 10, DCRoadMask, 0, 11, vbSrcAnd
            BitBlt frmMapBuffer.hdc, X, Y + 11, 11, 10, DCRoad, 0, 11, vbSrcPaint
         End If
         c = getIndex(row, col, "e")
         If ISROAD = (getData(c) And ISROAD) Or (getData(c) And TERRAIN_MAP) = underground Then
            BitBlt frmMapBuffer.hdc, X + 21, Y + 11, 11, 10, DCRoadMask, 21, 11, vbSrcAnd
            BitBlt frmMapBuffer.hdc, X + 21, Y + 11, 11, 10, DCRoad, 21, 11, vbSrcPaint
         End If
      Else

         BitBlt frmMapBuffer.hdc, X + 7, Y + 7, 8, 8, DCRoadMask, 7, 7, vbSrcAnd
         BitBlt frmMapBuffer.hdc, X + 7, Y + 7, 8, 8, DCRoad, 7, 7, vbSrcPaint
         
         c = getIndex(row, col, "n")
         If ISROAD = (getData(c) And ISROAD) Or (getData(c) And TERRAIN_MAP) = underground Then
            BitBlt frmMapBuffer.hdc, X + 7, Y, 8, 7, DCRoadMask, 7, 0, vbSrcAnd
            BitBlt frmMapBuffer.hdc, X + 7, Y, 8, 7, DCRoad, 7, 0, vbSrcPaint
         End If
         c = getIndex(row, col, "s")
         If ISROAD = (getData(c) And ISROAD) Or (getData(c) And TERRAIN_MAP) = underground Then
            BitBlt frmMapBuffer.hdc, X + 7, Y + 15, 8, 7, DCRoadMask, 7, 15, vbSrcAnd
            BitBlt frmMapBuffer.hdc, X + 7, Y + 15, 8, 7, DCRoad, 7, 15, vbSrcPaint
         End If
         
         c = getIndex(row, col, "w")
         If ISROAD = (getData(c) And ISROAD) Or (getData(c) And TERRAIN_MAP) = underground Then
            BitBlt frmMapBuffer.hdc, X, Y + 7, 7, 8, DCRoadMask, 0, 7, vbSrcAnd
            BitBlt frmMapBuffer.hdc, X, Y + 7, 7, 8, DCRoad, 0, 7, vbSrcPaint
         End If
         c = getIndex(row, col, "e")
         If ISROAD = (getData(c) And ISROAD) Or (getData(c) And TERRAIN_MAP) = underground Then
            BitBlt frmMapBuffer.hdc, X + 15, Y + 7, 7, 8, DCRoadMask, 15, 7, vbSrcAnd
            BitBlt frmMapBuffer.hdc, X + 15, Y + 7, 7, 8, DCRoad, 15, 7, vbSrcPaint
         End If

      End If
   End If

'DRAWING A BRIDGE
   If InStrB(1, getRoom(getIndex(row, col)), "bridge", vbBinaryCompare) > 0 Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCBridge, 0, 0, vbSrcCopy
   End If
   
'DRAWING SPECIAL ENTRANCE
   Select Case (celldata And SPECIAL_MAP)
   Case shop:    BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCShop, 0, 0, vbSrcCopy
   Case guild:   BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCGuild, 0, 0, vbSrcCopy
   Case inn:     BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCInn, 0, 0, vbSrcCopy
   Case city:    BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCCity, 0, 0, vbSrcCopy
   Case dungeon:
      If is_dark_room Then
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDungeonDark, 0, 0, vbSrcCopy
      Else
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDungeon, 0, 0, vbSrcCopy
      End If
   End Select

   If exploration Then
      Dim m As Integer
      m = 2
      BitBlt frmMapBuffer.hdc, X + 3, Y + 3, roomsize - 3, roomsize - 3, DCMove, 0, 0, vbSrcCopy
   End If
   
End Function

Public Function drawFlag(ByRef celldata As Long, X As Long, Y As Long)
   Select Case (celldata And FLAG_MAP)
   Case FLAG_WATER
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCflagWater, 0, 0, vbSrcCopy
   Case FLAG_ITEM
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCflagItem, 0, 0, vbSrcCopy
   Case FLAG_HERB
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCflagHerb, 0, 0, vbSrcCopy
   Case FLAG_TREASURY
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCflagTreasury, 0, 0, vbSrcCopy
   Case FLAG_KEY
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCflagKey, 0, 0, vbSrcCopy
   Case FLAG_MAGIC
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCflagMagic, 0, 0, vbSrcCopy
'   Case FLAG_MUDLLE
'      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagMudlle, 0, 0, vbSrcCopy
   Case FLAG_QUEST
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCflagQuest, 0, 0, vbSrcCopy
'   Case FLAG_QUESTION
'      BitBlt frmMapBuffer.hdc, x, y, roomsize, roomsize, DCflagQuestion, 0, 0, vbSrcCopy
   End Select
End Function

Public Sub drawWall(row As Integer, col As Integer, X, Y)
   If isValid(row - 1, col) Then _
      If LenB(aData(getIndex(row - 1, col), cDATA)) <> 0 Then _
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
   
   If isValid(row, col + 1) = True Then _
      If LenB(aData(getIndex(row, col + 1), cDATA)) <> 0 Then _
         BitBlt frmMapBuffer.hdc, X + roomsize - 1, Y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
   
   If isValid(row + 1, col) = True Then _
      If LenB(aData(getIndex(row + 1, col), cDATA)) <> 0 Then _
         BitBlt frmMapBuffer.hdc, X, Y + roomsize - 1, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
   
   If isValid(row, col - 1) = True Then _
      If LenB(aData(getIndex(row, col - 1), cDATA)) <> 0 Then _
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
End Sub

Public Sub drawExit(ByRef celldata As Long, X As Long, Y As Long)
   If (celldata And ROOM_MAP) = noRide_Dark Or (celldata And ROOM_MAP) = noRide_Sun Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCnoRideMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCnoRide, 0, 0, vbSrcPaint
   End If

   
   If (celldata And MONSTER_MAP) = MONSTER_EASY Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMonsterMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMonster1, 0, 0, vbSrcPaint
   End If
   If (celldata And MONSTER_MAP) = MONSTER_HARD Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMonsterMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMonster2, 0, 0, vbSrcPaint
   End If
   If (celldata And MONSTER_MAP) = MONSTER_GROUP Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMonsterMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCMonster3, 0, 0, vbSrcPaint
   End If
   
   
   
   If frmMap.mnuMap1.Checked Then
      If (celldata And ROOM_MAP) = noRide_Dark Or (celldata And ROOM_MAP) = Ride_Dark Then
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDarkMask, 0, 0, vbSrcAnd
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDark, 0, 0, vbSrcPaint
      End If
   End If
   
   If (celldata And N_MAP) = N_noexit Then
      BitBlt frmMapBuffer.hdc, X, Y + 0, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 1, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_noexit Then
      BitBlt frmMapBuffer.hdc, X + roomsize - 1, Y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + roomsize - 2, Y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
   End If
   If (celldata And S_MAP) = S_noexit Then
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 1, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 2, roomsize, roomsize, DCHWall, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_noexit Then
      BitBlt frmMapBuffer.hdc, X + 0, Y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + 1, Y, roomsize, roomsize, DCVWall, 0, 0, vbSrcCopy
   End If
'PORTAL
   If (celldata And N_MAP) = N_portal Then
      BitBlt frmMapBuffer.hdc, X, Y + 0, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
'      BitBlt frmMapBuffer.hdc, x, y + 1, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_portal Then
      BitBlt frmMapBuffer.hdc, X + roomsize - 1, Y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
'      BitBlt frmMapBuffer.hdc, x + roomsize - 2, y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
   End If
   If (celldata And S_MAP) = S_portal Then
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 1, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
'      BitBlt frmMapBuffer.hdc, x, y + roomsize - 2, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_portal Then
      BitBlt frmMapBuffer.hdc, X + 0, Y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
'      BitBlt frmMapBuffer.hdc, x + 1, y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
   End If
'DOOR
   If (celldata And N_MAP) = N_door Then
      BitBlt frmMapBuffer.hdc, X, Y + 0, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 1, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_door Then
      BitBlt frmMapBuffer.hdc, X + roomsize - 1, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + roomsize - 2, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And S_MAP) = S_door Then
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 1, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 2, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_door Then
      BitBlt frmMapBuffer.hdc, X + 0, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + 1, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
   End If
'HIDDENDOOR
   If (celldata And S_MAP) = S_hiddendoor Then
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 1, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 2, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 3, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_hiddendoor Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy 'DCHDoor
      BitBlt frmMapBuffer.hdc, X + 1, Y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy 'DCVBlack
      BitBlt frmMapBuffer.hdc, X + 2, Y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy 'DCVBlack
   End If
   If (celldata And N_MAP) = N_hiddendoor Then
      BitBlt frmMapBuffer.hdc, X, Y + 0, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 1, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 2, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_hiddendoor Then
      BitBlt frmMapBuffer.hdc, X + roomsize - 1, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + roomsize - 2, Y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + roomsize - 3, Y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy
   End If
   
   
   
'DOORPORTAL
   If (celldata And S_MAP) = S_doorportal Then
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 1, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 2, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 3, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_doorportal Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + 1, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + 2, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And N_MAP) = N_doorportal Then
      BitBlt frmMapBuffer.hdc, X, Y + 0, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 1, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 2, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_doorportal Then
      BitBlt frmMapBuffer.hdc, X + roomsize - 1, Y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + roomsize - 2, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + roomsize - 3, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
   End If
'HIDDENDOORPORTAL
   If (celldata And S_MAP) = (S_hiddendoor Or S_portal) Then
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 1, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy 'DCHPortal
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 2, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy 'DCHDoor
      BitBlt frmMapBuffer.hdc, X, Y + roomsize - 3, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy 'DCHBlack
   End If
   If (celldata And W_MAP) = (W_hiddendoor Or W_portal) Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + 1, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + 2, Y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy
   End If
   If (celldata And N_MAP) = (N_hiddendoor Or N_portal) Then
      BitBlt frmMapBuffer.hdc, X, Y + 0, roomsize, roomsize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 1, roomsize, roomsize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 2, roomsize, roomsize, DCHBlack, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = (E_hiddendoor Or E_portal) Then
      BitBlt frmMapBuffer.hdc, X + roomsize - 1, Y, roomsize, roomsize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + roomsize - 2, Y, roomsize, roomsize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + roomsize - 3, Y, roomsize, roomsize, DCVBlack, 0, 0, vbSrcCopy
   End If
'UP
   If (celldata And U_MAP) = U_exit Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUpMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUp, 0, 0, vbSrcPaint
   End If
   If (celldata And U_MAP) = U_portal Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUpMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUpPortal, 0, 0, vbSrcPaint
   End If
   If (celldata And U_door) = U_door Or (celldata And U_doorportal) = U_doorportal Or (celldata And U_hiddendoor) = U_hiddendoor Or (celldata And U_MAP) = U_MAP Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUpMask, 0, 0, vbSrcAnd
      If (celldata And U_hiddendoor) = U_hiddendoor Then
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUpHiddenDoor, 0, 0, vbSrcPaint
      Else
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCUpDoor, 0, 0, vbSrcPaint
      End If
   End If
   If (celldata And D_door) = D_door Or (celldata And D_doorportal) = D_doorportal Or (celldata And D_hiddendoor) = D_hiddendoor Or (celldata And D_MAP) = D_MAP Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDownMask, 0, 0, vbSrcAnd
      If (celldata And D_hiddendoor) = D_hiddendoor Then
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDownHiddenDoor, 0, 0, vbSrcPaint
      Else
         BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDownDoor, 0, 0, vbSrcPaint
      End If
   End If
'DOWN
   If (celldata And D_MAP) = D_exit Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDownMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDown, 0, 0, vbSrcPaint
   End If
   If (celldata And D_MAP) = D_portal Then
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDownMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCDownPortal, 0, 0, vbSrcPaint
   End If
Exit Sub
errorhandler:
  theROW = 15
  theCOL = 15
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
   If LenB(theDoornameUp) <> 0 Then Call myText(frmMapBuffer, theDoornameUp, theCenter + 4, theCenter - (n * roomsize), True, 8)
   If LenB(theDoornameNorth) <> 0 Then Call myText(frmMapBuffer, theDoornameNorth, theCenter + 4, theCenter - (m * roomsize), True, 8)
   If LenB(theDoornameEast) <> 0 Then Call myText(frmMapBuffer, theDoornameEast, theCenter + (n * roomsize), theCenter - 2, True, 8)
   If LenB(theDoornameWest) <> 0 Then Call myText(frmMapBuffer, theDoornameWest, theCenter - (n * roomsize), theCenter - 2, True, 8)
   If LenB(theDoornameSouth) <> 0 Then Call myText(frmMapBuffer, theDoornameSouth, theCenter + 4, theCenter + (m * roomsize), True, 8)
   If LenB(theDoornameDown) <> 0 Then Call myText(frmMapBuffer, theDoornameDown, theCenter + 4, theCenter + (n * roomsize), True, 8)
End Function

Public Sub drawPlayers()
Dim X As Integer
Dim Y As Integer
Dim cursor As Integer
Dim n As Integer
Dim s As String
   If viewPlayers Then
      For n = LBound(arrPlayers, 1) To arrPlayersIndex - 1
         For cursor = 1 To theCount
            If aData(cursor, cROOMNAME) = arrPlayers(n) Then
               X = theCenter + ((aData(cursor, cCOL) - theCOL) * roomsize) '- half
               Y = theCenter + ((aData(cursor, cROW) - theROW) * roomsize) '- half
               If X >= theMaximum Or X <= 1 Or Y >= theMaximum Or Y <= 1 Then
                  'out of area, skip drawing
               Else
                  Call myText(frmMapBuffer, arrPlayersNames(n), X, Y)
               End If
            End If
         Next
      Next
   End If
   viewPlayers = False
   
   
   If GODMODE And frmMap.mnuReceiver.Checked Then
      Dim myrow As Integer, mycol As Integer
      Dim ss As Variant, sss As Variant
      Dim x2 As Integer, y2 As Integer
      On Error Resume Next
      Open App.Path & "\_all.txt" For Input As #1 ' .. open the file
      Line Input #1, s ' Read a line into the variable
      Close #1 ' All done: close the file
      If Err.Number <> 0 Then
         'missing file
      Else
         ss = Split(s, ";", , vbBinaryCompare)
         For n = LBound(ss) To UBound(ss)
            sss = Split(ss(n), ",", , vbBinaryCompare)
            myrow = sss(3)
            mycol = sss(4)
            X = theCenter + ((sss(2) - theCOL) * roomsize)
            x2 = theCenter + ((sss(4) - theCOL) * roomsize)
            Y = theCenter + ((sss(1) - theROW) * roomsize)
            y2 = theCenter + ((sss(3) - theROW) * roomsize)
            'If x >= theMaximum Or x <= 1 Or y >= theMaximum Or y <= 1 Then
               'out of area, skip drawing
            'Else
               If frmMap.mnuPortals.Checked Then
                  frmMapBuffer.ForeColor = QBColor(11)
                  Call myLine(frmMapBuffer.hdc, X, Y, x2, y2)
                  Call myLine(frmMapBuffer.hdc, x2 - half + (n * 1), y2 - CInt(roomsize / 1.5), x2 - half + (n * 1), y2)
                  BitBlt frmMapBuffer.hdc, X - half, Y - half, roomsize, roomsize, DCPlayerMask, 0, 0, vbSrcAnd
                  'BitBlt frmMapBuffer.hdc, x2 - half, y2 - half, roomsize, roomsize, DCPlayerMask, 0, 0, vbSrcAnd
               End If
               If myrow <> theROW Or mycol <> theCOL Then
                  Call myText(frmMapBuffer, Left(sss(0), 3), x2, y2, False, 8, QBColor(11))
               End If
            'End If
         Next
      End If
   End If
   
End Sub

Public Sub drawFind()
Dim cursor As Integer, s As String
Dim X As Integer, Y As Integer
   For cursor = 1 To theCount
      If InStrB(1, LCase(aData(cursor, cROOMNAME)), LCase(lookfor), vbBinaryCompare) Then
         s = "(" & aData(cursor, cROW) & "," & aData(cursor, cCOL) & ")"
         Call informClient(s & Space(10 - Len(s)) & aData(cursor, cROOMNAME) & " - " & _
            "N[" & aData(cursor, cNDOOR) & "], " & _
            "E[" & aData(cursor, cEDOOR) & "], " & _
            "S[" & aData(cursor, cSDOOR) & "], " & _
            "W[" & aData(cursor, cWDOOR) & "], " & _
            "U[" & aData(cursor, cUDOOR) & "], " & _
            "D[" & aData(cursor, cDDOOR) & "]", True)

         X = theCenter + ((aData(cursor, cCOL) - theCOL) * roomsize) '- half
         Y = theCenter + ((aData(cursor, cROW) - theROW) * roomsize) '- half
         If X > theMaximum Or X < 1 Or Y > theMaximum Or Y < 1 Then
            'out of area, skip drawing
         Else
            BitBlt frmMapBuffer.hdc, X, Y, roomsize, roomsize, DCflagQuestion, 0, 0, vbSrcCopy
            Call myText(frmMapBuffer, "X", X, Y, , 8)
         End If
      End If
   Next
   s = "[" & aData(getIndex(theROW, theCOL), cROW) & "," & aData(getIndex(theROW, theCOL), cCOL) & "]"
   Call informClient("---------", True)
   Call informClient(s & Space(10 - Len(s)) & "Your current coordinates.", True)
   lookfor = vbNullString
End Sub
