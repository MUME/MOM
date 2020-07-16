VERSION 5.00
Begin VB.Form frmWorld 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Map of Arda"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   Icon            =   "frmWorld.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   287
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   575
End
Attribute VB_Name = "frmWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub drawWorld(ByRef startRow As Integer, ByRef startCol As Integer, ByRef roomsize As Integer)
Dim row As Integer, col As Integer
   Call loadGraphics(App.Path & mapWorldPath)
   BitBlt frmWorld.hdc, 0, 0, frmWorld.ScaleWidth, frmWorld.ScaleHeight, 0, 0, 0, vbBlackness
   For row = startRow To Me.ScaleHeight / roomsize + startRow
      For col = startCol To Me.ScaleWidth / roomsize + startCol
         If isValid(row, col) = True Then
            If LenB(aData(getIndex(row, col), cDATA)) <> 0 Then
               'Dim width As Long
               'Dim height As Long
               'width = roomsize
               'height = roomsize
               X = (1 + (col - startCol) * roomsize)
               Y = (1 + (row - startRow) * roomsize)
               Select Case (aData(getIndex(row, col), cDATA) And TERRAIN_MAP)
               Case road
                  BitBlt frmWorld.hdc, X, Y, roomsize, roomsize, DCRoad, 0, 0, vbSrcCopy
               Case plain
                  BitBlt frmWorld.hdc, X, Y, roomsize, roomsize, DCPlain, 0, 0, vbSrcCopy
               Case forest
                  BitBlt frmWorld.hdc, X, Y, roomsize, roomsize, DCForest, 0, 0, vbSrcCopy
               Case swamp
                  BitBlt frmWorld.hdc, X, Y, roomsize, roomsize, DCSwamp, 0, 0, vbSrcCopy
               Case hill
                  BitBlt frmWorld.hdc, X, Y, roomsize, roomsize, DCHill, 0, 0, vbSrcCopy
               Case underground
                  BitBlt frmWorld.hdc, X, Y, roomsize, roomsize, DCUnderground, 0, 0, vbSrcCopy
               Case water
                  BitBlt frmWorld.hdc, X, Y, roomsize, roomsize, DCWater, 0, 0, vbSrcCopy
               Case mountain
                  BitBlt frmWorld.hdc, X, Y, roomsize, roomsize, DCMountain, 0, 0, vbSrcCopy
               End Select
            End If
         End If
      Next
   Next
   frmWorld.Circle ((theCOL - startCol) * roomsize, (theROW - startRow) * roomsize), mapRadius * roomsize, QBColor(15)
   If frmMap.mnuMap1.Checked = True Then Call frmMap.mnuMap1_Click
   If frmMap.mnuMap2.Checked = True Then Call frmMap.mnuMap2_Click
   If frmMap.mnuMap3.Checked = True Then Call frmMap.mnumap3_Click
   WinTopMost.MakeTopMost frmWorld.hWnd
   Me.Show
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   theCOL = X
   theROW = Y
   If getIndex(theROW, theCOL) > 0 Then
      theDoornameNorth = aData(getIndex(theROW, theCOL), cNDOOR)
      theDoornameEast = aData(getIndex(theROW, theCOL), cEDOOR)
      theDoornameSouth = aData(getIndex(theROW, theCOL), cSDOOR)
      theDoornameWest = aData(getIndex(theROW, theCOL), cWDOOR)
      theDoornameUp = aData(getIndex(theROW, theCOL), cUDOOR)
      theDoornameDown = aData(getIndex(theROW, theCOL), cDDOOR)
   End If
   Call DrawMap
   'frmWorld.WindowState = vbMinimized
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub
