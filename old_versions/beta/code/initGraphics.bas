Attribute VB_Name = "initGraphics"
Option Explicit
Public pRoad As Image
Public pField As Image
Public pForest As Image
Public pSwamp As Image
Public pHill As Image
Public pMountain As Image
Public pWater As Image
Public pSpecial As Image

Public DCRoad As Long
Public DCPlain As Long
Public DCForest As Long
Public DCSwamp As Long
Public DCHill As Long
Public DCMountain As Long
Public DCWater As Long
Public DCSpecial As Long
Public DCNone As Long
Public DCHDoor As Long
Public DCHPortal As Long
Public DCHWall As Long
Public DCVDoor As Long
Public DCVPortal As Long
Public DCVWall As Long
Public DCDark As Long
Public DCDarkMask As Long
Public DCnoRideMask As Long
Public DCnoRide As Long
Public DCUpMask As Long
Public DCDownMask As Long
Public DCUp As Long
Public DCDown As Long
Public DCDownPortal As Long
Public DCDownDoor As Long
Public DCUpPortal As Long
Public DCUpDoor As Long
Public DCPlayer As Long
Public DCPlayerMask As Long
Public DCMoveMask As Long
Public DCMove As Long
Public DCMoveEnd As Long
Public Const mapSmallPath = "\bitmap40\"
Public Const mapNormalPath = "\bitmap22\"
Public Const mapLargePath = "\bitmap14\"

Public Sub loadGraphics(bitmapPath As String)

   Call clearGraphics

   Set pRoad = frmTools.road
   Set pField = frmTools.plain
   Set pForest = frmTools.forest
   Set pSwamp = frmTools.swamp
   Set pHill = frmTools.hill
   Set pMountain = frmTools.mountain
   Set pWater = frmTools.water
   Set pSpecial = frmTools.special

   DCRoad = GenerateBitmapDC(bitmapPath & "road.bmp")
   DCPlain = GenerateBitmapDC(bitmapPath & "plain.bmp")
   DCForest = GenerateBitmapDC(bitmapPath & "forest.bmp")
   DCSwamp = GenerateBitmapDC(bitmapPath & "swamp.bmp")
   DCHill = GenerateBitmapDC(bitmapPath & "hill.bmp")
   DCMountain = GenerateBitmapDC(bitmapPath & "mountain.bmp")
   DCWater = GenerateBitmapDC(bitmapPath & "water.bmp")
   DCSpecial = GenerateBitmapDC(bitmapPath & "special.bmp")
   DCHDoor = GenerateBitmapDC(bitmapPath & "HDoor.bmp")
   DCHPortal = GenerateBitmapDC(bitmapPath & "HPortal.bmp")
   DCHWall = GenerateBitmapDC(bitmapPath & "HWall.bmp")
   DCVDoor = GenerateBitmapDC(bitmapPath & "VDoor.bmp")
   DCVPortal = GenerateBitmapDC(bitmapPath & "VPortal.bmp")
   DCVWall = GenerateBitmapDC(bitmapPath & "VWall.bmp")
   DCDarkMask = GenerateBitmapDC(bitmapPath & "darkmask.bmp")
   DCDark = GenerateBitmapDC(bitmapPath & "dark.bmp")
   DCnoRideMask = GenerateBitmapDC(bitmapPath & "noridemask.bmp")
   DCnoRide = GenerateBitmapDC(bitmapPath & "noride.bmp")
   
   DCUpMask = GenerateBitmapDC(bitmapPath & "upmask.bmp")
   DCUp = GenerateBitmapDC(bitmapPath & "up.bmp")
   DCUpDoor = GenerateBitmapDC(bitmapPath & "updoor.bmp")
   DCUpPortal = GenerateBitmapDC(bitmapPath & "upportal.bmp")
   
   DCDownMask = GenerateBitmapDC(bitmapPath & "downmask.bmp")
   DCDown = GenerateBitmapDC(bitmapPath & "down.bmp")
   DCDownDoor = GenerateBitmapDC(bitmapPath & "downdoor.bmp")
   DCDownPortal = GenerateBitmapDC(bitmapPath & "downportal.bmp")
   
   DCPlayerMask = GenerateBitmapDC(bitmapPath & "playermask.bmp")
   DCPlayer = GenerateBitmapDC(bitmapPath & "player.bmp")

   DCMoveMask = GenerateBitmapDC(bitmapPath & "movemask.bmp")
   DCMove = GenerateBitmapDC(bitmapPath & "move.bmp")
   DCMoveEnd = GenerateBitmapDC(bitmapPath & "DCMoveEnd.bmp")

End Sub
