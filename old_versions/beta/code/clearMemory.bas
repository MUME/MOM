Attribute VB_Name = "clearMemory"
Public Sub ZIP()

   Call clearGraphics
   Unload frmMap
   Unload frmTools
   End

End Sub
   
Public Sub clearGraphics()
   
   DeleteDC (DCRoad)
   DeleteDC (DCPlain)
   DeleteDC (DCForest)
   DeleteDC (DCSwamp)
   DeleteDC (DCHill)
   DeleteDC (DCMountain)
   DeleteDC (DCWater)
   DeleteDC (DCSpecial)
   DeleteDC (DCNone)
   DeleteDC (DCHDoor)
   DeleteDC (DCHPortal)
   DeleteDC (DCHWall)
   DeleteDC (DCVDoor)
   DeleteDC (DCVPortal)
   DeleteDC (DCVWall)
   DeleteDC (DCDark)
   DeleteDC (DCDarkMask)
   DeleteDC (DCnoRideMask)
   DeleteDC (DCnoRide)
   DeleteDC (DCUpMask)
   DeleteDC (DCDownMask)
   DeleteDC (DCUp)
   DeleteDC (DCDown)
   DeleteDC (DCDownPortal)
   DeleteDC (DCDownDoor)
   DeleteDC (DCUpPortal)
   DeleteDC (DCUpDoor)
   DeleteDC (DCPlayer)
   DeleteDC (DCPlayerMask)
   DeleteDC (DCMoveMask)
   DeleteDC (DCMove)
   DeleteDC (DCMoveEnd)
   Set pRoad = Nothing
   Set pField = Nothing
   Set pForest = Nothing
   Set pSwamp = Nothing
   Set pHill = Nothing
   Set pMountain = Nothing
   Set pWater = Nothing
   Set pSpecial = Nothing

End Sub
