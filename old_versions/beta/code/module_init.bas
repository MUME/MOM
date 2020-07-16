Attribute VB_Name = "initVariables"
Option Explicit
Option Compare Binary

Public Const DEBUG_MODE = False
Public arr(1 To 300, 1 To 600) As Long
Public arrRoomname(1 To 300, 1 To 600) As String
Public arrDescription(1 To 300, 1 To 600) As String
Public arrDesc(1 To 300, 1 To 600) As String
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
Public Const mountain = 20
Public Const water = 24
Public Const special = 28
'_____________________________
Public Const N_MAP = 224
Public Const N_noexit = 0
Public Const N_exit = 32
Public Const N_portal = 160
Public Const N_door = 64
Public Const N_doorportal = 192
Public Const N_hiddendoor = 96
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
Public Const MONSTER_MAP = 8388608
'_____________________________

Public arrMinRoom As Long
Public arrMaxRoom As Long
Public arrMinMove As Long
Public arrMaxMove As Long
Public arrMinRow As Long
Public arrMinCol As Long
Public arrMaxRow As Long
Public arrMaxCol As Long
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

Public theRowNorth As Long
Public theRowEast As Long
Public theRowSouth As Long
Public theRowWest As Long
Public theRowUp As Long
Public theRowDown As Long
Public theColNorth As Long
Public theColEast As Long
Public theColSouth As Long
Public theColWest As Long
Public theColUp As Long
Public theColDown As Long

Public compactMode As Boolean
Public tempData As Long
Public viewPortals As Boolean
Public viewMovement As Boolean

Public Sub loadVariables()

   WorldLoaded = False
   MappingData = False
   MappingGetUpdate = False
   dataFromMUME = False
   Erase arr
   Erase arrDesc
   Erase arrRoomStack
   Erase arrMoveStack
   arrMinRow = LBound(arr, 1)
   arrMinCol = LBound(arr, 2)
   arrMaxRow = UBound(arr, 1)
   arrMaxCol = UBound(arr, 2)
   arrMinRoom = LBound(arrRoomStack)
   arrMaxRoom = UBound(arrRoomStack)
   arrMinMove = LBound(arrMoveStack)
   arrMaxMove = UBound(arrMoveStack)

End Sub

