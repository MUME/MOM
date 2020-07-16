Attribute VB_Name = "initGlobals"
Option Explicit
' RULES:
' byteIndex => (charIndex * 2) - 1
' charIndex => (byteIndex + 1) / 2
'MidB => MidB$(string, INDEX(MUST BE converted to BYTE position i.e. 1 + LenB("kala")))

Public MUDname As String

Public theLEVEL As Long
Public staticLevel As Boolean

Public Const lookHeader = "["
Public Const lookFooter = "m"
Public Const colourEndCode = "[0m"
Public systemRoot As String

Public tmpOutput As String
Public compactMode As Boolean
Public tempdata As Long
Public viewPortals As Boolean
Public viewMovement As Boolean
Public viewNotes As Boolean
Public viewDoornames As Boolean

Public followMode As Boolean
Public theClientName As String
Public alwaysOnTop As Boolean
Public canUndo As Boolean
Public undoRow As Integer
Public undoCol As Integer

Public Const BOLD = "1"
Public Const UNDERLINE = "4"
Public Const BLACK = "30"
Public Const RED = "31"
Public Const GREEN = "32"
Public Const YELLOW = "33"
Public Const BLUE = "34"
Public Const MAGENTA = "35"
Public Const CYAN = "36"
Public Const WHITE = "37"

Public MappingMode As Boolean
Public MappingData As Boolean
Public MappingCase As Long
Public theDesc As String

Public Const ROOM_MAP = 3
Public Const noRide_Dark = 0
Public Const noRide_Sun = 1
Public Const Ride_Dark = 2
Public Const Ride_Sun = 3
'____________________________
Public theTerrain As Long
Public Const TERRAIN_MAP = 28
Public Const road = 0
Public Const plain = 4
Public Const forest = 8
Public Const swamp = 12
Public Const hill = 16
Public Const underground = 20
Public Const water = 24
Public Const mountain = 28
'_____________________________
Public Const N_MAP = 224
Public Const N_noexit = 0
Public Const N_exit = 32
Public Const N_door = 64
Public Const N_hiddendoor = 96
Public Const N_portal = 160
Public Const N_doorportal = 192
'_____________________________
Public Const E_MAP = 1792
Public Const E_noexit = 0
Public Const E_exit = 256
Public Const E_portal = 1280
Public Const E_door = 512
Public Const E_doorportal = 1536
Public Const E_hiddendoor = 768
'_____________________________
Public Const S_MAP = 14336
Public Const S_noexit = 0
Public Const S_exit = 2048
Public Const S_portal = 10240
Public Const S_door = 4096
Public Const S_doorportal = 12288
Public Const S_hiddendoor = 6144
'_____________________________
Public Const W_MAP = 114688
Public Const W_noexit = 0
Public Const W_exit = 16384
Public Const W_portal = 81920
Public Const W_door = 32768
Public Const W_doorportal = 98304
Public Const W_hiddendoor = 49152
'_____________________________
Public Const U_MAP = 917504
Public Const U_noexit = 0
Public Const U_exit = 131072
Public Const U_portal = 655360
Public Const U_door = 262144
Public Const U_doorportal = 786432
Public Const U_hiddendoor = 393216
'_____________________________
Public Const D_MAP = 7340032
Public Const D_noexit = 0
Public Const D_exit = 1048576
Public Const D_portal = 5242880
Public Const D_door = 2097152
Public Const D_doorportal = 6291456
Public Const D_hiddendoor = 3145728
'_____________________________
' 0 = no monster
' 8388608 => easy monster (~lev 10)
' 1073741824 => hard monster (~lev 15)
' 1082130432 => hard monster group(~lev 20)
'Public Const MONSTER_MAP = 8388608 'oldmap
Public Const MONSTER_EASY = 8388608
Public Const MONSTER_HARD = 1073741824
Public Const MONSTER_GROUP = 1082130432
Public Const MONSTER_MAP = 1082130432 'new map is 8388608 + 1073741824

'_____________________________

        Public Const shop = 16777216
       Public Const guild = 33554432
         Public Const inn = 50331648
      'Public Const bridge = 67108864
        Public Const city = 83886080
    Public Const dungeon = 100663296
Public Const SPECIAL_MAP = 117440512
'_____________________________________
'FREE SPECIAL TERRAIN
Public theFlag As Long
Public theRoad As Long
             
             Public Const FLAG_NONE = 0
    Public Const FLAG_WATER = 134217728
     Public Const FLAG_ITEM = 268435456
     Public Const FLAG_HERB = 402653184
 Public Const FLAG_TREASURY = 536870912
      Public Const FLAG_KEY = 671088640
    Public Const FLAG_MAGIC = 805306368
    Public Const FLAG_QUEST = 939524096
      Public Const FLAG_MAP = 939524096
    'Public Const FLAG_MAP = 2013265920
 'Public Const FLAG_CREATURE = 939524096 'flag_mapiks
  ''''''Public Const FLAG_MUDLLE = 1073741824
''''''Public Const FLAG_TIMEBOMB = 1207959552
'   Public Const FLAG_QUEST = 1342177280
''''''Public Const FLAG_QUESTION = 1476395008

       Public Const ISROAD = 67108864

'_____________________________
Public arrMaxData As Integer
Public arrMinRoom As Integer
Public arrMaxRoom As Integer
Public arrMinMove As Integer
Public arrMaxMove As Integer
Public arrMinRow As Integer
Public arrMinCol As Integer
Public arrMaxRow As Integer
Public arrMaxCol As Integer
Public theSun As Boolean
Public theRide As Boolean
Public theMonster As Boolean
Public theRoomString
Public theRoom As Long
Public theRoomname As String
Public theRoomdesc As String
Public theRoomStringOk As Boolean
Public theExitNorth As Boolean
Public theExitEast As Boolean
Public theExitSouth As Boolean
Public theExitWest As Boolean
Public theExitUp As Boolean
Public theExitDown As Boolean
Public theDoorNorth As Boolean
Public theDoorEast As Boolean
Public theDoorSouth As Boolean
Public theDoorWest As Boolean
Public theDoorUp As Boolean
Public theDoorDown As Boolean

Public thePortalNorth As Boolean
Public thePortalEast  As Boolean
Public thePortalSouth As Boolean
Public thePortalWest  As Boolean
Public thePortalUp    As Boolean
Public thePortalDown  As Boolean

Public theDoorPortalNorth As Boolean
Public theDoorPortalEast As Boolean
Public theDoorPortalSouth As Boolean
Public theDoorPortalWest As Boolean
Public theDoorPortalUp As Boolean
Public theDoorPortalDown As Boolean

Public theDoornameNorth As String
Public theDoornameEast As String
Public theDoornameSouth As String
Public theDoornameWest As String
Public theDoornameUp As String
Public theDoornameDown As String

Public theHiddendoorNorth As Boolean
Public theHiddendoorEast As Boolean
Public theHiddendoorSouth As Boolean
Public theHiddendoorWest As Boolean
Public theHiddendoorUp As Boolean
Public theHiddendoorDown As Boolean

Public theROWNorth As Long
Public theROWEast As Long
Public theROWSouth As Long
Public theROWWest As Long
Public theROWUp As Long
Public theROWDown As Long
Public theCOLNorth As Long
Public theCOLEast As Long
Public theCOLSouth As Long
Public theCOLWest As Long
Public theCOLUp As Long
Public theCOLDown As Long

'graphics
'======================================================
Public DCUpHiddenDoor As Long
Public DCDownHiddenDoor As Long

Public DCNone As Long
Public DCRoad As Long
Public DCRoadMask As Long

Public DCRoadH As Long
Public DCRoadV As Long
Public DCPlain As Long
Public DCUnderground As Long
Public DCForest As Long
Public DCSwamp As Long
Public DCHill As Long
Public DCWater As Long
Public DCMountain As Long

Public DCForestDark As Long
Public DCSwampDark As Long
Public DCHillDark As Long
Public DCWaterDark As Long
Public DCMountainDark As Long

Public DCHDoor As Long
Public DCHPortal As Long
Public DCHWall As Long
Public DCVDoor As Long
Public DCVPortal As Long
Public DCVWall As Long

Public DCVBlack As Long
Public DCHBlack As Long
Public DCVWhite As Long
Public DCHWhite As Long

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
Public DCPlayer2 As Long
Public DCPlayerMask As Long
Public DCMoveMask As Long
Public DCMove As Long
Public DCMoveEnd As Long
Public DCMonster As Long
Public DCMonster1 As Long
Public DCMonster2 As Long
Public DCMonster3 As Long
Public DCMonsterMask As Long

Public DCShop As Long
Public DCGuild As Long
Public DCInn As Long
Public DCBridge As Long
Public DCCity As Long
Public DCDungeon As Long
Public DCDungeonDark As Long

Public DCflagWater As Long
Public DCflagItem As Long
Public DCflagHerb As Long
Public DCflagTreasury As Long
Public DCflagKey As Long
Public DCflagMagic As Long
Public DCflagMudlle As Long
Public DCflagQuest As Long
Public DCflagQuestion As Long

Public DCflagWaterMask As Long
Public DCflagItemMask As Long
Public DCflagHerbMask As Long
Public DCflagTreasuryMask As Long
Public DCflagKeyMask As Long
Public DCflagMagicMask As Long
Public DCflagMudlleMask As Long
Public DCflagQuestMask As Long
Public DCflagQuestionMask As Long

Public Const mapSmallPath = "\map32x32\"
Public Const mapNormalPath = "\map22x22\"
Public Const mapLargePath = "\map14x14\"
Public Const mapWorldPath = "\map1x1\"

Public arrCollision(1 To 50) As String
Public nn As Integer

Public Sub loadMOMini()
   Dim line As String
   Dim arrLine As Variant
   Dim file
   Dim ss As TextStream
   Dim OpenFileForReading
   OpenFileForReading = 1
   If fso.FileExists(App.Path & "\MOM.ini") = False Then Call saveMOMini
   Set file = fso.OpenTextFile(App.Path & "\MOM.ini", ForReading, True)
   Set file = fso.GetFile(App.Path & "\MOM.ini")
   Set ss = file.OpenAsTextStream(OpenFileForReading)
   
   Do While Not ss.AtEndOfStream
      line = ss.readLine
      If (MidB$(line, 1, LenB("1")) <> "#") Then
         arrLine = Split(line, "=", , vbBinaryCompare)
         Call setVariable(arrLine(0), arrLine(1))
      End If
   Loop
   ss.Close
   Set file = Nothing

   If Len(lookColour) = 0 Then
      lookColour = "[32m"
      Call saveMOMini
      MsgBox "MOM look colour is 'green'." & vbCrLf & _
      "Please change MUME roomname to green: 'change colour look green'", _
      "NB! If you want to use other colour, then change the look colour from MOM menu." & vbCrLf, _
      vbOKOnly, "Hello!"
      Call saveMOMini
   End If
   If Len(roomdescriptionColour) = 0 Then
      roomdescriptionColour = "[37m"
      Call saveMOMini
      MsgBox "The OnlineMap description colour is 'white'." & vbCrLf & _
      "Please change MUME description to white: 'change colour roomdescription white'" & vbCrLf & _
      "NB! If you want to use other colour, then change the roomdescription value in 'MOM.ini' file." & vbCrLf, _
      vbOKOnly, "Hello!"
      Call saveMOMini
   End If
   If LenB(arrCollision(LBound(arrCollision))) = 0 Then
      arrCollision(1) = "Alas, you cannot go that way..."
      arrCollision(2) = "doesn't want you riding"
      arrCollision(3) = "The descent is too steep, you need to climb to go there."
      arrCollision(4) = "The ascent is too steep, you need to climb to go there."
      arrCollision(5) = "Maybe you should get on your feet first?"
      arrCollision(6) = "In your dreams, or what?"
      arrCollision(7) = "Your mount refuses to follow your orders!"
      arrCollision(8) = "No way! You are fighting for your life!"
      arrCollision(9) = "Oops! You cannot go there riding!"
      arrCollision(10) = "You can't go into deep water!"
      arrCollision(11) = "You failed swimming there."
      arrCollision(12) = "You need to swim to go there."
      arrCollision(13) = "Nah... You feel too relaxed to do that.."
      arrCollision(14) = "You failed to climb there and fall down, hurting yourself."
      arrCollision(15) = " seems to be closed."
      arrCollision(16) = " too exhausted"
      Call saveMOMini
   End If
End Sub

Public Sub saveMOMini()
   Set file = fso.OpenTextFile(App.Path & "\MOM.ini", ForWriting, True)
   file.WriteLine "client=" & theClientName
   file.WriteLine "alwaysontop=" & IIF(frmMap.mnuAlwaysOnTop.Checked, 1, 0)
   file.WriteLine "systemroot=" & IIF(LenB(systemRoot) = 0, "C:\\WINNT", systemRoot)
   file.WriteLine "remotehost=" & IIF(LenB(frmMap.tcpPlayer.RemoteHost) = 0, "mume.pvv.org", frmMap.tcpPlayer.RemoteHost)
   file.WriteLine "remoteport=" & IIF(frmMap.tcpPlayer.RemotePort = 0, "23", frmMap.tcpPlayer.RemotePort)
   file.WriteLine "localhost=" & "localhost"
   file.WriteLine "localport=" & IIF(frmMap.tcpMUD.LocalPort = 0, "1001", frmMap.tcpMUD.LocalPort)
   file.WriteLine "lookcolour=" & lookColour
   file.WriteLine "roomdescriptioncolour=" & roomdescriptionColour
   file.WriteLine "autosync=" & IIF(frmMap.mnuAutosync.Checked, 1, 0)
   Select Case True
      Case frmMap.mnuMap1.Checked: file.WriteLine "mapsize=1"
      Case frmMap.mnuMap2.Checked: file.WriteLine "mapsize=2"
      Case frmMap.mnuMap3.Checked: file.WriteLine "mapsize=3"
   End Select
   file.WriteLine "notes=" & IIF(frmMap.mnuNotes.Checked, 1, 0)
   file.WriteLine "doornames=" & IIF(frmMap.mnuDoornames.Checked, 1, 0)
   file.WriteLine "portals=" & IIF(frmMap.mnuPortals.Checked, 1, 0)
   file.WriteLine "spam=" & IIF(frmMap.mnuSpam.Checked, 1, 0)
   file.WriteLine "brief=" & IIF(frmMap.mnuBrief.Checked, 1, 0)
   file.WriteLine "grid=" & IIF(frmMap.mnuGrid.Checked, 1, 0)
   file.WriteLine "gridxy=" & IIF(frmMap.mnuGridXY.Checked, 1, 0)
   file.WriteLine "locateretry=" & IIF(locateRetry < 1, 3, locateRetry)
   file.WriteLine "feedback=" & IIF(frmMap.mnuFeedback.Checked, 1, 0)
   file.WriteLine "mapdescription=" & IIF(frmMap.mnuMapDescription.Checked, 1, 0)
   file.WriteLine "players=" & IIF(frmMap.mnuPlayers.Checked, 1, 0)
   file.WriteLine "movement=" & IIF(frmMap.mnuMovement.Checked, 1, 0)
   file.WriteLine "name=mudmaniac"
   file.WriteLine "productname=" & "windows"
   file.WriteLine "setuppath=" & App.Path
   Dim i As Integer
   For i = 1 To UBound(arrCollision)
      If LenB(arrCollision(i)) <> 0 Then
         file.WriteLine "collision=" & arrCollision(i)
      End If
   Next
   file.Close
   Set file = Nothing
End Sub


Public Function setVariable(ByVal name As String, ByVal val As String)
   Dim value As String
   value = LCase(Trim$(val))
   name = LCase(Trim$(name))
   Select Case name
      Case "collision"
         nn = nn + 1
         arrCollision(nn) = val
      Case "name"
      Case "systemroot"
         systemRoot = value
      Case "productname"
      Case "setuppath"
      Case "remotehost"
         frmMap.tcpPlayer.RemoteHost = value
      Case "remoteport"
         frmMap.tcpPlayer.RemotePort = value
      Case "localhost"
      Case "localport"
         frmMap.tcpMUD.LocalPort = value
      Case "lookcolour"
         Call frmMap.changeRoomColour(True, value)
      Case "roomdescriptioncolour"
         roomdescriptionColour = value
      Case "autosync"
         frmMap.mnuAutosync.Checked = Not CBool(value)
         Call frmMap.mnuAutosync_Click
      Case "mapsize"
         Select Case value
            Case 1: Call frmMap.mnuMap1_Click
            Case 2: Call frmMap.mnuMap2_Click
            Case 3: Call frmMap.mnumap3_Click
            Case Else: Call frmMap.mnuMap1_Click
         End Select
      Case "locateretry"
         If LenB(value) = 0 Then locateRetry = 3 Else locateRetry = CInt(value)
      Case "client"
         If LenB(value) = 0 Then theClientName = "jaba mud client" Else theClientName = value
'----------
      Case "movement"
         frmMap.mnuMovement.Checked = Not CBool(value)
         Call frmMap.mnuMovement_Click
      Case "notes"
         frmMap.mnuNotes.Checked = Not CBool(value)
         Call frmMap.mnuNotes_Click
      Case "doornames"
         frmMap.mnuDoornames.Checked = Not CBool(value)
         Call frmMap.mnuDoornames_Click
      Case "players"
         frmMap.mnuPlayers.Checked = Not CBool(value)
         Call frmMap.mnuPlayers_Click
      Case "portals"
         frmMap.mnuPortals.Checked = Not CBool(value)
         Call frmMap.mnuPortals_Click
      Case "spam"
         frmMap.mnuSpam.Checked = Not CBool(value)
         Call frmMap.mnuSpam_Click
      Case "brief"
         frmMap.mnuBrief.Checked = Not CBool(value)
         Call frmMap.mnuBrief_Click
      Case "grid"
         frmMap.mnuGrid.Checked = Not CBool(value)
         Call frmMap.mnuGrid_Click
      Case "gridxy"
         frmMap.mnuGridXY.Checked = Not CBool(value)
         Call frmMap.mnuGridXY_Click
      Case "alwaysontop"
         frmMap.mnuAlwaysOnTop.Checked = Not CBool(value)
         Call frmMap.mnuAlwaysOnTop_Click
      Case "feedback"
         frmMap.mnuFeedback.Checked = Not CBool(value)
         Call frmMap.mnuFeedback_Click
      Case "mapdescription"
         frmMap.mnuMapDescription.Checked = Not CBool(value)
         Call frmMap.mnuMapDescription_Click
      Case Else
         errorModule = "File <MOM.ini> contains undefined or invalid values! name=" & name & ", value=" & value
         writeError (errorModule)
   End Select
End Function

Public Sub loadGraphics(bitmapPath As String)
   Call clearGraphics
   DCCity = GenerateBitmapDC(bitmapPath & "terrain_city.bmp")
   
   DCRoad = GenerateBitmapDC(bitmapPath & "terrain_road.bmp")
   DCRoadMask = GenerateBitmapDC(bitmapPath & "terrain_road_mask.bmp")
   
   'map2
   DCRoadH = GenerateBitmapDC(bitmapPath & "terrain_roadH.bmp")
   DCRoadV = GenerateBitmapDC(bitmapPath & "terrain_roadV.bmp")
   
   DCPlain = GenerateBitmapDC(bitmapPath & "terrain_plain.bmp")
   DCForest = GenerateBitmapDC(bitmapPath & "terrain_forest.bmp")
   DCSwamp = GenerateBitmapDC(bitmapPath & "terrain_swamp.bmp")
   DCHill = GenerateBitmapDC(bitmapPath & "terrain_hill.bmp")
   DCWater = GenerateBitmapDC(bitmapPath & "terrain_water.bmp")
   DCMountain = GenerateBitmapDC(bitmapPath & "terrain_mountain.bmp")
   
   DCForestDark = GenerateBitmapDC(bitmapPath & "terrain_forest_dark.bmp")
   DCSwampDark = GenerateBitmapDC(bitmapPath & "terrain_swamp_dark.bmp")
   DCHillDark = GenerateBitmapDC(bitmapPath & "terrain_hill_dark.bmp")
   DCWaterDark = GenerateBitmapDC(bitmapPath & "terrain_water_dark.bmp")
   DCMountainDark = GenerateBitmapDC(bitmapPath & "terrain_mountain_dark.bmp")
   
   DCUnderground = GenerateBitmapDC(bitmapPath & "terrain_underground.bmp")
   DCShop = GenerateBitmapDC(bitmapPath & "special_shop.bmp")
   DCGuild = GenerateBitmapDC(bitmapPath & "special_guild.bmp")
   DCInn = GenerateBitmapDC(bitmapPath & "special_inn.bmp")
   DCBridge = GenerateBitmapDC(bitmapPath & "special_bridge.bmp")
   DCCity = GenerateBitmapDC(bitmapPath & "special_city.bmp")
   DCDungeon = GenerateBitmapDC(bitmapPath & "special_dungeon.bmp")
   
   DCDungeonDark = GenerateBitmapDC(bitmapPath & "special_dungeon_dark.bmp")

   DCDarkMask = GenerateBitmapDC(bitmapPath & "type_darkmask.bmp")
   DCDark = GenerateBitmapDC(bitmapPath & "type_dark.bmp")
   DCnoRideMask = GenerateBitmapDC(bitmapPath & "type_noridemask.bmp")
   DCnoRide = GenerateBitmapDC(bitmapPath & "type_noride.bmp")

   DCHBlack = GenerateBitmapDC(bitmapPath & "exit_Hblack.bmp")
   DCVBlack = GenerateBitmapDC(bitmapPath & "exit_Vblack.bmp")
   DCHWhite = GenerateBitmapDC(bitmapPath & "exit_Hwhite.bmp")
   DCVWhite = GenerateBitmapDC(bitmapPath & "exit_Vwhite.bmp")
   
   DCHDoor = GenerateBitmapDC(bitmapPath & "exit_Hdoor.bmp")
   DCHPortal = GenerateBitmapDC(bitmapPath & "exit_Hportal.bmp")
   
   DCHWall = GenerateBitmapDC(bitmapPath & "exit_Hwall.bmp")
   DCVWall = GenerateBitmapDC(bitmapPath & "exit_Vwall.bmp")
   
   DCHWall = GenerateBitmapDC(bitmapPath & "exit_Hblack.bmp")
   DCVDoor = GenerateBitmapDC(bitmapPath & "exit_Vdoor.bmp")
   DCVPortal = GenerateBitmapDC(bitmapPath & "exit_Vportal.bmp")

   DCUpHiddenDoor = GenerateBitmapDC(bitmapPath & "exit_uphiddendoor.bmp")
   DCDownHiddenDoor = GenerateBitmapDC(bitmapPath & "exit_downhiddendoor.bmp")

   DCVWall = GenerateBitmapDC(bitmapPath & "exit_Vblack.bmp")
   DCUpMask = GenerateBitmapDC(bitmapPath & "exit_upmask.bmp")
   DCUp = GenerateBitmapDC(bitmapPath & "exit_up.bmp")
   DCUpDoor = GenerateBitmapDC(bitmapPath & "exit_updoor.bmp")
   DCUpPortal = GenerateBitmapDC(bitmapPath & "exit_upportal.bmp")
   DCDownMask = GenerateBitmapDC(bitmapPath & "exit_downmask.bmp")
   DCDown = GenerateBitmapDC(bitmapPath & "exit_down.bmp")
   DCDownDoor = GenerateBitmapDC(bitmapPath & "exit_downdoor.bmp")
   DCDownPortal = GenerateBitmapDC(bitmapPath & "exit_downportal.bmp")
   
   DCPlayerMask = GenerateBitmapDC(bitmapPath & "playermask.bmp")
   DCPlayer = GenerateBitmapDC(bitmapPath & "player.bmp")
   DCPlayer2 = GenerateBitmapDC(bitmapPath & "player2.bmp")
   
   
   DCMonsterMask = GenerateBitmapDC(bitmapPath & "monstermask.bmp")
   'DCMonster = GenerateBitmapDC(bitmapPath & "monster.bmp")
   DCMonster1 = GenerateBitmapDC(bitmapPath & "monster1.bmp")
   DCMonster2 = GenerateBitmapDC(bitmapPath & "monster2.bmp")
   DCMonster3 = GenerateBitmapDC(bitmapPath & "monster3.bmp")
   
   DCMoveMask = GenerateBitmapDC(bitmapPath & "movemask.bmp")
   DCMove = GenerateBitmapDC(bitmapPath & "move.bmp")
   DCMoveEnd = GenerateBitmapDC(bitmapPath & "moveEnd.bmp")
  
   DCflagWater = GenerateBitmapDC(bitmapPath & "flag_Water.bmp")
   DCflagWaterMask = GenerateBitmapDC(bitmapPath & "flag_WaterMask.bmp")
   DCflagItem = GenerateBitmapDC(bitmapPath & "flag_Item.bmp")
   DCflagItemMask = GenerateBitmapDC(bitmapPath & "flag_ItemMask.bmp")
   DCflagHerb = GenerateBitmapDC(bitmapPath & "flag_Herb.bmp")
   DCflagHerbMask = GenerateBitmapDC(bitmapPath & "flag_HerbMask.bmp")
   DCflagTreasury = GenerateBitmapDC(bitmapPath & "flag_Treasury.bmp")
   DCflagTreasuryMask = GenerateBitmapDC(bitmapPath & "flag_TreasuryMask.bmp")
   DCflagKey = GenerateBitmapDC(bitmapPath & "flag_Key.bmp")
   DCflagKeyMask = GenerateBitmapDC(bitmapPath & "flag_KeyMask.bmp")
   DCflagMagic = GenerateBitmapDC(bitmapPath & "flag_Magic.bmp")
   DCflagMagicMask = GenerateBitmapDC(bitmapPath & "flag_MagicMask.bmp")
   DCflagMudlle = GenerateBitmapDC(bitmapPath & "flag_Mudlle.bmp")
   DCflagMudlleMask = GenerateBitmapDC(bitmapPath & "flag_MudlleMask.bmp")
   DCflagQuest = GenerateBitmapDC(bitmapPath & "flag_Quest.bmp")
   DCflagQuestMask = GenerateBitmapDC(bitmapPath & "flag_QuestMask.bmp")
   DCflagQuestion = GenerateBitmapDC(bitmapPath & "flag_Question.bmp")
   DCflagQuestionMask = GenerateBitmapDC(bitmapPath & "flag_QuestionMask.bmp")

'GROUP FEATURE
'   Call initGroupGrahics
End Sub
