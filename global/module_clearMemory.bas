Attribute VB_Name = "clearMemory"
Option Explicit
Public Declare Sub SleepAPI Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)
    
Public Sub ZIP()
   On Error Resume Next
   'frmMap.objInformer.Quit
   'SleepAPI 300
   'Set frmMap.objInformer = Nothing
   Set cast128 = Nothing
   Set md5 = Nothing
   Call clearGraphics
'   call clearGroupGrahics
   End
End Sub
   
Public Sub clearGraphics()
   DeleteDC (DCNone)
   DeleteDC (DCRoad)
   DeleteDC (DCRoadMask)
   
   'map2
   DeleteDC (DCRoadH)
   DeleteDC (DCRoadV)
   DeleteDC (DCPlain)
   DeleteDC (DCForest)
   DeleteDC (DCSwamp)
   DeleteDC (DCHill)
   DeleteDC (DCWater)
   DeleteDC (DCMountain)
   DeleteDC (DCUnderground)
   
   DeleteDC (DCForestDark)
   DeleteDC (DCSwampDark)
   DeleteDC (DCHillDark)
   DeleteDC (DCWaterDark)
   DeleteDC (DCMountainDark)

   DeleteDC (DCHDoor)
   DeleteDC (DCHPortal)
   DeleteDC (DCHWall)
   DeleteDC (DCVDoor)
   DeleteDC (DCVPortal)
   DeleteDC (DCVWall)
   
   DeleteDC (DCVBlack)
   DeleteDC (DCHBlack)
   DeleteDC (DCVWhite)
   DeleteDC (DCHWhite)
   
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
   'DeleteDC (DCMonster)
   DeleteDC (DCMonster1)
   DeleteDC (DCMonster2)
   DeleteDC (DCMonster3)
   DeleteDC (DCMonsterMask)
   
   DeleteDC (DCShop)
   DeleteDC (DCGuild)
   DeleteDC (DCInn)
   DeleteDC (DCBridge)
   DeleteDC (DCCity)
   DeleteDC (DCDungeon)
   DeleteDC (DCDungeonDark)

   DeleteDC (DCflagWater)
   DeleteDC (DCflagItem)
   DeleteDC (DCflagHerb)
   DeleteDC (DCflagTreasury)
   DeleteDC (DCflagKey)
   DeleteDC (DCflagMagic)
   DeleteDC (DCflagMudlle)
   DeleteDC (DCflagQuest)
   DeleteDC (DCflagQuestion)
   
   DeleteDC (DCflagWaterMask)
   DeleteDC (DCflagItemMask)
   DeleteDC (DCflagHerbMask)
   DeleteDC (DCflagTreasuryMask)
   DeleteDC (DCflagKeyMask)
   DeleteDC (DCflagMagicMask)
   DeleteDC (DCflagMudlleMask)
   DeleteDC (DCflagQuestMask)
   DeleteDC (DCflagQuestionMask)
End Sub

'Public Sub clearGroupGraphics()
'   DeleteDC (DCpuppetOne)
'   DeleteDC (DCpuppetTwo)
'   DeleteDC (DCpuppetThree)
'   DeleteDC (DCpuppetFour)
'   DeleteDC (DCpuppetOneBar)
'   DeleteDC (DCpuppetTwoBar)
'   DeleteDC (DCpuppetThreeBar)
'   DeleteDC (DCpuppetFourBar)
'   DeleteDC (DCpuppetBar)
'End Sub
