Attribute VB_Name = "drawing"
Option Explicit
Option Compare Binary
Public mapRadius As Long
Public RoomSize As Long
Public absX As Long
Public absY As Long
Public theCenter As Long
Public theMaximum As Long

Public Sub DrawMap()
Dim row As Long, col As Long
On Error GoTo errorhandler
   
   If frmMap.mnuSmall.Checked = True Then
      mapRadius = 3
      RoomSize = 40
   End If
   If frmMap.mnuNormal.Checked = True Then
      mapRadius = 6
      RoomSize = 22
   End If
   If frmMap.mnuLarge.Checked = True Then
      mapRadius = 10
      RoomSize = 14
   End If

   BitBlt frmMapBuffer.hdc, 0, 0, frmMapBuffer.ScaleWidth, frmMapBuffer.ScaleHeight, 0, 0, 0, vbBlackness
   For row = theRow - mapRadius To theRow + mapRadius
      For col = theCol - mapRadius To theCol + mapRadius
         If checkArrayLimit(row, col) = True Then
            absX = ((1 + col - (theCol - mapRadius)) * RoomSize) - RoomSize
            absY = ((1 + row - (theRow - mapRadius)) * RoomSize) - RoomSize
            If drawTerrain(arr(row, col), absX, absY) = True Then
               Call drawExit(arr(row, col), absX, absY)
            End If
         End If
      Next
   Next
   
   If frmMap.mnuPortals.Checked = True Then Call checkPortal
   
   theCenter = CInt((mapRadius + 1) * RoomSize - (RoomSize / 2))
   theMaximum = (2 * mapRadius + 1) * RoomSize

   BitBlt frmMapBuffer.hdc, theCenter - (RoomSize / 2), theCenter - (RoomSize / 2), RoomSize, RoomSize, DCPlayerMask, 0, 0, vbSrcAnd
   BitBlt frmMapBuffer.hdc, theCenter - (RoomSize / 2), theCenter - (RoomSize / 2), RoomSize, RoomSize, DCPlayer, 0, 0, vbSrcPaint
   
   If frmMap.mnuMovement.Checked = True Then Call drawMovement
   If frmMap.mnuDoornamesHide.Checked = False Then Call drawDoornames
   BitBlt frmMap.hdc, 10, 10, 300, 300, frmMapBuffer.hdc, 0, 0, vbSrcCopy
   
   frmMap.Refresh
   
Exit Sub
errorhandler:
   errorData = "drawing DrawMap"
   writeError (errorData)

End Sub


Public Sub checkPortal()
Dim row As Long, col As Long
   
   For row = theRow - mapRadius To theRow + mapRadius
      For col = theCol - mapRadius To theCol + mapRadius
         If checkArrayLimit(row, col) = True Then
            absX = ((1 + col - (theCol - mapRadius)) * RoomSize) - RoomSize
            absY = ((1 + row - (theRow - mapRadius)) * RoomSize) - RoomSize
            Dim celldata
            celldata = arr(row, col)
            If celldata > 0 Then
               If (celldata And N_MAP) = N_portal Or _
                  (celldata And E_MAP) = E_portal Or _
                  (celldata And S_MAP) = S_portal Or _
                  (celldata And W_MAP) = W_portal Or _
                  (celldata And U_MAP) = U_portal Or _
                  (celldata And D_MAP) = D_portal Then
                     If (celldata And N_MAP) = N_portal Then Call drawPortal(row, col, N_MAP)
                     If (celldata And E_MAP) = E_portal Then Call drawPortal(row, col, E_MAP)
                     If (celldata And S_MAP) = S_portal Then Call drawPortal(row, col, S_MAP)
                     If (celldata And W_MAP) = W_portal Then Call drawPortal(row, col, W_MAP)
                     If (celldata And U_MAP) = U_portal Then Call drawPortal(row, col, U_MAP)
                     If (celldata And D_MAP) = D_portal Then Call drawPortal(row, col, D_MAP)
               End If
            End If
         End If
      Next
   Next

End Sub

Public Sub drawPortal(ByRef row, ByRef col, ByRef theMap)
Dim tmpData
Dim tmpRoom
Dim tmpRow As Long
Dim tmpCol As Long
Dim startX, startY, endX, endY

   tmpData = arr(row, col)
   Select Case (tmpData And theMap)
   Case U_exit, U_door, U_portal
     tmpRoom = Split(arrDesc(row, col), ";")
     tmpRow = tmpRoom(14)
     tmpCol = tmpRoom(15)
     startX = ((1 + col - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 4)
     startY = ((1 + row - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 4)
     endX = ((1 + tmpCol - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
     endY = ((1 + tmpRow - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
   Case D_exit, D_door, D_portal
     tmpRoom = Split(arrDesc(row, col), ";")
     tmpRow = tmpRoom(17)
     tmpCol = tmpRoom(18)
     startX = ((1 + col - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 1.3)
     startY = ((1 + row - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 1.3)
     endX = ((1 + tmpCol - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
     endY = ((1 + tmpRow - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
   Case N_portal
     tmpRoom = Split(arrDesc(row, col), ";")
     tmpRow = tmpRoom(2)
     tmpCol = tmpRoom(3)
     startX = ((1 + col - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
     startY = ((1 + row - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 4)
     endX = ((1 + tmpCol - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
     endY = ((1 + tmpRow - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
   Case E_portal
     tmpRoom = Split(arrDesc(row, col), ";")
     tmpRow = tmpRoom(5)
     tmpCol = tmpRoom(6)
     startX = ((1 + col - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 1.3)
     startY = ((1 + row - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
     endX = ((1 + tmpCol - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
     endY = ((1 + tmpRow - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
   Case S_portal
     tmpRoom = Split(arrDesc(row, col), ";")
     tmpRow = tmpRoom(8)
     tmpCol = tmpRoom(9)
     startX = ((1 + col - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
     startY = ((1 + row - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 1.3)
     endX = ((1 + tmpCol - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
     endY = ((1 + tmpRow - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
   Case W_portal
     tmpRoom = Split(arrDesc(row, col), ";")
     tmpRow = tmpRoom(11)
     tmpCol = tmpRoom(12)
     startX = ((1 + col - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 4)
     startY = ((1 + row - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
     endX = ((1 + tmpCol - (theCol - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
     endY = ((1 + tmpRow - (theRow - mapRadius)) * RoomSize) - RoomSize + (RoomSize / 2)
   End Select

   'Dim hpen As Long
   'hpen = CreatePen(PS_SOLID, 1, RGB(230, 100, 50))
   frmMapBuffer.ForeColor = QBColor(0)
   Call myLine(frmMapBuffer.hdc, startX, startY, endX, endY)
   'DeleteObject hpen

End Sub

Public Function drawMovement()

Dim n As Integer
Dim X As Integer
Dim Y As Integer
   
   theCenter = (mapRadius + 1) * RoomSize - (RoomSize / 2)
   If roomCount > 0 Then
      For n = 1 To roomCount
         X = theCenter + ((arrRoomStack(n, 2) - theCol) * RoomSize)
         Y = theCenter + ((arrRoomStack(n, 1) - theRow) * RoomSize)
         If X > theMaximum Or X < 1 Or Y > theMaximum Or Y < 1 Then
         Else
            BitBlt frmMapBuffer.hdc, X - (RoomSize / 2), Y - (RoomSize / 2), RoomSize, RoomSize, DCMoveMask, 0, 0, vbSrcAnd
            BitBlt frmMapBuffer.hdc, X - (RoomSize / 2), Y - (RoomSize / 2), RoomSize, RoomSize, DCMove, 0, 0, vbSrcPaint
            If n = roomCount Then
               BitBlt frmMapBuffer.hdc, X - (RoomSize / 2), Y - (RoomSize / 2), RoomSize, RoomSize, DCMoveMask, 0, 0, vbSrcAnd
               BitBlt frmMapBuffer.hdc, X - (RoomSize / 2), Y - (RoomSize / 2), RoomSize, RoomSize, DCMoveEnd, 0, 0, vbSrcPaint
            End If
         End If
      Next
   End If

End Function

Public Function drawTerrain(ByRef celldata, ByRef X As Long, ByRef Y As Long)
Dim width As Long
Dim height As Long
   
   width = RoomSize
   height = RoomSize
   If celldata <= 0 Then
      drawTerrain = False
      Exit Function
   End If
   
   Select Case (celldata And TERRAIN_MAP)
   Case road
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCRoad, 0, 0, vbSrcCopy
   Case plain
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCPlain, 0, 0, vbSrcCopy
   Case forest
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCForest, 0, 0, vbSrcCopy
   Case swamp
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCSwamp, 0, 0, vbSrcCopy
   Case hill
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCHill, 0, 0, vbSrcCopy
   Case mountain
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCMountain, 0, 0, vbSrcCopy
   Case water
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCWater, 0, 0, vbSrcCopy
   Case special
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCSpecial, 0, 0, vbSrcCopy
   End Select
   drawTerrain = True
End Function

Public Sub drawExit(ByRef celldata, ByRef X As Long, ByRef Y As Long)
'On Error GoTo ErrorHandler
   
   If (celldata And ROOM_MAP) = noRide_Dark Or (celldata And ROOM_MAP) = Ride_Dark Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCDarkMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCDark, 0, 0, vbSrcPaint
   End If
   If (celldata And ROOM_MAP) = noRide_Dark Or (celldata And ROOM_MAP) = noRide_Sun Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCnoRideMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCnoRide, 0, 0, vbSrcPaint
   End If

   If (celldata And N_MAP) = N_noexit Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCHWall, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_noexit Then
      BitBlt frmMapBuffer.hdc, X + RoomSize - 1, Y, RoomSize, RoomSize, DCVWall, 0, 0, vbSrcCopy
   End If
   If (celldata And S_MAP) = S_noexit Then
      BitBlt frmMapBuffer.hdc, X, Y + RoomSize - 1, RoomSize, RoomSize, DCHWall, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_noexit Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCVWall, 0, 0, vbSrcCopy
   End If
   
   If (celldata And S_MAP) = S_portal Then
      BitBlt frmMapBuffer.hdc, X, Y + RoomSize - 1, RoomSize, RoomSize, DCHPortal, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_portal Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCVPortal, 0, 0, vbSrcCopy
   End If
   If (celldata And N_MAP) = N_portal Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCHPortal, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_portal Then
      BitBlt frmMapBuffer.hdc, X + RoomSize - 1, Y, RoomSize, RoomSize, DCVPortal, 0, 0, vbSrcCopy
   End If
   
   If (celldata And S_MAP) = S_door Then
      BitBlt frmMapBuffer.hdc, X, Y + RoomSize - 1, RoomSize, RoomSize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_door Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCVDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And N_MAP) = N_door Then
      BitBlt frmMapBuffer.hdc, X, Y + 1, RoomSize, RoomSize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_door Then
      BitBlt frmMapBuffer.hdc, X + RoomSize - 1, Y, RoomSize, RoomSize, DCVDoor, 0, 0, vbSrcCopy
   End If

   If (celldata And S_MAP) = S_hiddendoor Then
      BitBlt frmMapBuffer.hdc, X, Y + RoomSize - 1, RoomSize, RoomSize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + RoomSize - 2, RoomSize, RoomSize, DCHWall, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_hiddendoor Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + 1, Y, RoomSize, RoomSize, DCVWall, 0, 0, vbSrcCopy
   End If
   If (celldata And N_MAP) = N_hiddendoor Then
      BitBlt frmMapBuffer.hdc, X, Y + 1, RoomSize, RoomSize, DCHDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 2, RoomSize, RoomSize, DCHWall, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_hiddendoor Then
      BitBlt frmMapBuffer.hdc, X + RoomSize - 1, Y, RoomSize, RoomSize, DCVDoor, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + RoomSize - 2, Y, RoomSize, RoomSize, DCVWall, 0, 0, vbSrcCopy
   End If
'DOOR AND PORTAL
   If (celldata And S_MAP) = S_doorportal Then
      BitBlt frmMapBuffer.hdc, X, Y + RoomSize - 1, RoomSize, RoomSize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + RoomSize - 2, RoomSize, RoomSize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_doorportal Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + 1, Y, RoomSize, RoomSize, DCVDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And N_MAP) = N_doorportal Then
      BitBlt frmMapBuffer.hdc, X, Y + 1, RoomSize, RoomSize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 2, RoomSize, RoomSize, DCHDoor, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_doorportal Then
      BitBlt frmMapBuffer.hdc, X + RoomSize - 1, Y, RoomSize, RoomSize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + RoomSize - 2, Y, RoomSize, RoomSize, DCVDoor, 0, 0, vbSrcCopy
   End If
'HIDDEN DOOR AND PORTAL
   If (celldata And S_MAP) = S_MAP Then
      BitBlt frmMapBuffer.hdc, X, Y + RoomSize - 1, RoomSize, RoomSize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + RoomSize - 2, RoomSize, RoomSize, DCHWall, 0, 0, vbSrcCopy
   End If
   If (celldata And W_MAP) = W_MAP Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + 1, Y, RoomSize, RoomSize, DCVWall, 0, 0, vbSrcCopy
   End If
   If (celldata And N_MAP) = N_MAP Then
      BitBlt frmMapBuffer.hdc, X, Y + 1, RoomSize, RoomSize, DCHPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X, Y + 2, RoomSize, RoomSize, DCHWall, 0, 0, vbSrcCopy
   End If
   If (celldata And E_MAP) = E_MAP Then
      BitBlt frmMapBuffer.hdc, X + RoomSize - 1, Y, RoomSize, RoomSize, DCVPortal, 0, 0, vbSrcCopy
      BitBlt frmMapBuffer.hdc, X + RoomSize - 2, Y, RoomSize, RoomSize, DCVWall, 0, 0, vbSrcCopy
   End If

   If (celldata And U_MAP) = U_exit Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCUpMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCUp, 0, 0, vbSrcPaint
   End If
   If (celldata And U_MAP) = U_portal Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCUpMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCUpPortal, 0, 0, vbSrcPaint
   End If
   If (celldata And U_MAP) = U_door Or (celldata And U_MAP) = U_doorportal Or (celldata And U_MAP) = U_hiddendoor Or (celldata And U_MAP) = U_MAP Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCUpMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCUpDoor, 0, 0, vbSrcPaint
   End If

   If (celldata And D_MAP) = D_exit Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCDownMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCDown, 0, 0, vbSrcPaint
   End If
   If (celldata And D_MAP) = D_portal Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCDownMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCDownPortal, 0, 0, vbSrcPaint
   End If
   If (celldata And D_MAP) = D_door Or (celldata And D_MAP) = D_doorportal Or (celldata And D_MAP) = D_hiddendoor Or (celldata And D_MAP) = D_MAP Then
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCDownMask, 0, 0, vbSrcAnd
      BitBlt frmMapBuffer.hdc, X, Y, RoomSize, RoomSize, DCDownDoor, 0, 0, vbSrcPaint
   End If
Exit Sub

errorhandler:
  theRow = 15
  theCol = 15
  Resume Next
End Sub

Public Function drawDoornames()
Dim n, m
   If RoomSize > 30 Then
      n = 2
      m = 1
   Else
      n = 3
      m = 2
   End If
   
   Call myText(frmMapBuffer, theDoornameUp, theCenter - (n * RoomSize), theCenter - (n * RoomSize))
   Call myText(frmMapBuffer, theDoornameNorth, theCenter, theCenter - (m * RoomSize))
   Call myText(frmMapBuffer, theDoornameEast, theCenter + (n * RoomSize), theCenter)
   Call myText(frmMapBuffer, theDoornameWest, theCenter - (n * RoomSize), theCenter)
   Call myText(frmMapBuffer, theDoornameSouth, theCenter, theCenter + (m * RoomSize))
   Call myText(frmMapBuffer, theDoornameDown, theCenter + (n * RoomSize), theCenter + (n * RoomSize))

End Function
