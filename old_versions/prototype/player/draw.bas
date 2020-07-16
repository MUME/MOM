Attribute VB_Name = "drawing"
Option Explicit
Option Compare Binary
Public Const mapRadius = 5
Public absX As Long
Public absY As Long
Public Const RoomSize = 26
Public theCenter As Long
Public theMaximum As Long
Public H As Long
Public V As Long

Public Sub DrawVirtualMoves()
   target.ScaleMode = 3
   Dim n As Integer, row As Long, col As Long
   If roomCount > 0 Then
      For n = 1 To roomCount
         row = theCenter + ((arrRoomStack(n, 2) - theCol) * RoomSize)
         col = theCenter + ((arrRoomStack(n, 1) - theRow) * RoomSize)
         If row > theMaximum Or row < 1 Or col > theMaximum Or col < 1 Then
         Else
            target.Circle (row, col), 1, QBColor(12)
            target.Circle (row, col), 2, QBColor(0)
            If n = roomCount Then target.Circle (row, col), 4, QBColor(15)
         End If
      Next
   End If
End Sub

Public Sub DrawMap()
   Dim row As Long, col As Long
   target.ScaleMode = 3
   For row = theRow - mapRadius To theRow + mapRadius
      For col = theCol - mapRadius To theCol + mapRadius
         If checkArrayLimit(row, col) = True Then
            Call Draw(arr(row, col), 1 + row - (theRow - mapRadius), 1 + col - (theCol - mapRadius))
         End If
      Next
   Next
   theCenter = (mapRadius + 1) * RoomSize - (RoomSize / 2)
   theMaximum = (2 * mapRadius + 1) * RoomSize
   target.Circle (theCenter, theCenter), 1, QBColor(5)
   target.Circle (theCenter, theCenter), 2, QBColor(5)
   target.Circle (theCenter, theCenter), 3, QBColor(0)
   target.Circle (theCenter, theCenter), 4, QBColor(15)
   target.Circle (theCenter, theCenter), 5, QBColor(0)
End Sub
Public Sub caseArea(ByVal t As Long, ByRef x As Long, ByRef y As Long, ByRef pos As Long)
   Dim width As Long
   Dim height As Long
   If pos = V Then
      width = 2
      height = RoomSize
   Else
      width = RoomSize
      height = 2
   End If
   Select Case t
   Case road
      target.PaintPicture pRoad, x, y, width, height
   Case plain
      target.PaintPicture pField, x, y, width, height
   Case forest
      target.PaintPicture pForest, x, y, width, height
   Case swamp
      target.PaintPicture pSwamp, x, y, width, height
   Case hill
      target.PaintPicture pHill, x, y, width, height
   Case mountain
      target.PaintPicture pMountain, x, y, width, height
   Case water
      target.PaintPicture pWater, x, y, width, height
   Case special
      target.PaintPicture pSpecial, x, y, width, height
   End Select
End Sub

Public Sub Draw(ByRef celldata, ByVal row As Long, ByVal col As Long)
'On Error GoTo ErrorHandler
   target.ScaleMode = 3
   absX = (col * RoomSize) - RoomSize
   absY = (row * RoomSize) - RoomSize
   If celldata <= 0 Then
      target.PaintPicture pNone, absX, absY, RoomSize, RoomSize
      Exit Sub
   End If
'   If celldata < 0 Then
'      target.PaintPicture pNone, absX, absY, RoomSize, RoomSize
'      target.Circle (absX + (RoomSize / 2), absY + (RoomSize / 2)), 3, QBColor(15)
'      Exit Sub
'   End If
'   If (celldata And room_map) = noRide_Dark Then theRoom_text = theRoom_text + vbCrLf + "noride, dark"
'   If (celldata And room_map) = noRide_Sun Then theRoom_text = theRoom_text + vbCrLf + "noride, sun"
'   If (celldata And room_map) = Ride_Dark Then theRoom_text = theRoom_text + vbCrLf + "ride, dark"
'   If (celldata And room_map) = Ride_Sun Then theRoom_text = theRoom_text + vbCrLf + "ride, sun"
   
   If (celldata And terrain_map) = road Then
      target.PaintPicture pRoad, absX, absY, RoomSize, RoomSize
   End If
   If (celldata And terrain_map) = plain Then
      target.PaintPicture pField, absX, absY, RoomSize, RoomSize
   End If
   If (celldata And terrain_map) = forest Then
      target.PaintPicture pForest, absX, absY, RoomSize, RoomSize
   End If
   If (celldata And terrain_map) = swamp Then
      target.PaintPicture pSwamp, absX, absY, RoomSize, RoomSize
   End If
   If (celldata And terrain_map) = hill Then
      target.PaintPicture pHill, absX, absY, RoomSize, RoomSize
   End If
   If (celldata And terrain_map) = mountain Then
      target.PaintPicture pMountain, absX, absY, RoomSize, RoomSize
   End If
   If (celldata And terrain_map) = water Then
      target.PaintPicture pWater, absX, absY, RoomSize, RoomSize
   End If
   If (celldata And terrain_map) = special Then
      target.PaintPicture pSpecial, absX, absY, RoomSize, RoomSize
   End If
'------------------- DIRECTIONS -------------------------
'   If (celldata And N_map) = N_exit Then
'      Call caseArea((celldata And terrain_map), absX, absY, H)
'      target.Line (absX, absY)-(absX + RoomSize, absY), QBColor(15)
'   End If
'   If (celldata And E_map) = E_exit Then
'      Call caseArea((celldata And terrain_map), absX + RoomSize, absY, V)
'      'target.Line (absX + RoomSize, absY)-(absX + RoomSize, absY + RoomSize), QBColor(15)
'   End If
'   If (celldata And S_map) = S_exit Then
'      Call caseArea((celldata And terrain_map), absX, absY + RoomSize, H)
'      target.Line (absX, absY + RoomSize)-(absX + RoomSize, absY + RoomSize), QBColor(15)
'   End If
'   If (celldata And W_map) = W_exit Then
'      Call caseArea((celldata And terrain_map), absX, absY, V)
'      target.Line (absX, absY)-(absX, absY + RoomSize), QBColor(15)
'   End If

   If (celldata And S_map) = S_door Then
      target.Line (absX, absY + RoomSize - 1)-(absX + RoomSize, absY + RoomSize - 1), QBColor(12)
   End If
   If (celldata And W_map) = W_door Then
      target.Line (absX + 1, absY)-(absX + 1, absY + RoomSize), QBColor(12)
   End If
   If (celldata And N_map) = N_door Then
      target.Line (absX, absY + 1)-(absX + RoomSize, absY + 1), QBColor(12)
'      target.Line (absX, absY + 1)-(absX + RoomSize, absY + 1), QBColor(12)
   End If
   If (celldata And E_map) = E_door Then
'      target.Line (absX + RoomSize, absY)-(absX + RoomSize, absY + RoomSize), QBColor(12)
      target.Line (absX + RoomSize - 1, absY)-(absX + RoomSize - 1, absY + RoomSize), QBColor(12)
   End If

   If (celldata And S_map) = S_special Then
      target.Line (absX, absY + RoomSize - 1)-(absX + RoomSize, absY + RoomSize - 1), QBColor(13)
   End If
   If (celldata And W_map) = W_special Then
      target.Line (absX + 1, absY)-(absX + 1, absY + RoomSize), QBColor(13)
   End If
   If (celldata And N_map) = N_special Then
      target.Line (absX, absY + 1)-(absX + RoomSize, absY + 1), QBColor(13)
   End If
   If (celldata And E_map) = E_special Then
      target.Line (absX + RoomSize - 1, absY)-(absX + RoomSize - 1, absY + RoomSize), QBColor(13)
   End If

   If (celldata And N_map) = N_noexit Then
      target.PaintPicture phNone, absX, absY
'      target.Line (absX, absY)-(absX + RoomSize, absY), QBColor(0)
   End If
   If (celldata And E_map) = E_noexit Then
      target.PaintPicture pvNone, absX + RoomSize - 1, absY
'      target.Line (absX + RoomSize, absY)-(absX + RoomSize, absY + RoomSize), QBColor(0)
   End If
   If (celldata And S_map) = S_noexit Then
      target.PaintPicture phNone, absX, absY + RoomSize - 1
      'target.Line (absX, absY + RoomSize)-(absX + RoomSize, absY + RoomSize), QBColor(0)
   End If
   If (celldata And W_map) = W_noexit Then
      target.PaintPicture pvNone, absX, absY
      'target.Line (absX, absY)-(absX, absY + RoomSize), QBColor(0)
   End If
   
   
   If (celldata And U_map) = U_noexit Then
   End If
   If (celldata And U_map) = U_exit Then
       target.Circle (absX + 6, absY + 5), 1, QBColor(15)
       target.Circle (absX + 6, absY + 5), 2, QBColor(15)
       target.Circle (absX + 6, absY + 5), 3, QBColor(0)
   End If
   If (celldata And U_map) = U_door Then
      target.Circle (absX + 6, absY + 5), 1, QBColor(12)
      target.Circle (absX + 6, absY + 5), 2, QBColor(12)
      target.Circle (absX + 6, absY + 5), 3, QBColor(0)
   End If
   If (celldata And U_map) = U_special Then
      target.Circle (absX + 6, absY + 5), 1, QBColor(11)
      target.Circle (absX + 6, absY + 5), 2, QBColor(11)
      target.Circle (absX + 6, absY + 5), 3, QBColor(0)
   End If
   If (celldata And D_map) = D_noexit Then
   End If
   If (celldata And D_map) = D_exit Then
      target.Circle (absX + RoomSize - 6, absY + RoomSize - 5), 1, QBColor(15)
      target.Circle (absX + RoomSize - 6, absY + RoomSize - 5), 2, QBColor(15)
      target.Circle (absX + RoomSize - 6, absY + RoomSize - 5), 3, QBColor(0)
   End If
   If (celldata And D_map) = D_door Then
      target.Circle (absX + RoomSize - 6, absY + RoomSize - 5), 1, QBColor(12)
      target.Circle (absX + RoomSize - 6, absY + RoomSize - 5), 2, QBColor(12)
      target.Circle (absX + RoomSize - 6, absY + RoomSize - 5), 3, QBColor(0)
   End If
   If (celldata And D_map) = D_special Then
      target.Circle (absX + RoomSize - 6, absY + RoomSize - 5), 1, QBColor(11)
      target.Circle (absX + RoomSize - 6, absY + RoomSize - 5), 2, QBColor(11)
      target.Circle (absX + RoomSize - 6, absY + RoomSize - 5), 3, QBColor(0)
   End If
   '   MsgBox theRoom_text
Exit Sub

ErrorHandler:
  theRow = 15
  theCol = 15
  Resume Next
End Sub
