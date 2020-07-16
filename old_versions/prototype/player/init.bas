Attribute VB_Name = "init"
Option Explicit
Option Compare Binary
Public Const pretty = &HA56E3A
Public MAP_MODE As Boolean
Public MAP_THE_DATA As Boolean
Public MAP_THE_CASE As Long
Public READ_NORTH As Boolean
Public theDesc As String
Public Const debug_mode = False
Public Const room_map = 3
Public Const noRide_Dark = 0
Public Const noRide_Sun = 1
Public Const Ride_Dark = 2
Public Const Ride_Sun = 3
'____________________________
Public Const terrain_map = 28
Public Const road = 0   '8
Public Const plain = 4   '16
Public Const forest = 8 '32
Public Const swamp = 12  '64
Public Const hill = 16   '128
Public Const mountain = 20  '256
Public Const water = 24  '512
Public Const special = 28  '512
'_____________________________
Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4
Public Const UP = 5
Public Const DOWN = 6
Public Const N_map = 96
Public Const N_noexit = 0
Public Const N_exit = 32
Public Const N_door = 64
Public Const N_special = 96
Public Const E_map = 384
Public Const E_noexit = 0
Public Const E_exit = 128
Public Const E_door = 256
Public Const E_special = 384
Public Const S_map = 1536
Public Const S_noexit = 0
Public Const S_exit = 512
Public Const S_door = 1024
Public Const S_special = 1536
Public Const W_map = 6144
Public Const W_noexit = 0
Public Const W_exit = 2048
Public Const W_door = 4096
Public Const W_special = 6144
Public Const U_map = 24576
Public Const U_noexit = 0
Public Const U_exit = 8192
Public Const U_door = 16384
Public Const U_special = 24576
Public Const D_map = 98304
Public Const D_noexit = 0
Public Const D_exit = 32768
Public Const D_door = 65536
Public Const D_special = 98304
'Public Const rowDirectionEast = 131072
'Public Const rowDirectionWest = 393216
'Public Const rowDirectionNorth = 0
'Public Const rowDirectionSouth = 131072   '262144
'Public Const rowDirectionMap = 131072  '393216
'Public Const rowOffsetMap = 7864320
'Public Const rowDivision = 524288
'Public Const colDirectionNorth = 0
'Public Const colDirectionSouth = 16777216
'Public Const colDirectionEast = 0   '8388608
'Public Const colDirectionWest = 8388608   '393216
'Public Const colDirectionMap = 8388608 '25165824
'Public Const colOffsetMap = 1006632960
'Public Const colDivision = 33554432
Public target As PictureBox
Public pNone As Image
Public phNone As Image
Public pvNone As Image
Public pRoad As Image
Public pField As Image
Public pForest As Image
Public pSwamp As Image
Public pHill As Image
Public pMountain As Image
Public pWater As Image
Public pSpecial As Image
Public arr(1 To 600, 1 To 300) As Long
Public arrDesc(1 To 600, 1 To 300) As String
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
Public theRoomName As String
Public theRoomDesc As String
Public theRoomStringOk As Boolean
Public theRoomNorth As Boolean
Public theRoomEast As Boolean
Public theRoomSouth As Boolean
Public theRoomWest As Boolean
Public theRoomUp As Boolean
Public theRoomDown As Boolean
Public theDoorNorth As Boolean
Public theDoorEast As Boolean
Public theDoorSouth As Boolean
Public theDoorWest As Boolean
Public theDoorUp As Boolean
Public theDoorDown As Boolean
Public theSpecialNorth As Boolean
Public theSpecialEast As Boolean
Public theSpecialSouth As Boolean
Public theSpecialWest As Boolean
Public theSpecialUp As Boolean
Public theSpecialDown As Boolean
Public theDoorNameNorth As String
Public theDoorNameEast As String
Public theDoorNameSouth As String
Public theDoorNameWest As String
Public theDoorNameUp As String
Public theDoorNameDown As String
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
Public tempData

Public Sub Initialize()
   WorldLoaded = False
   MAP_THE_DATA = False
   Set target = BestEST.map
   Set pNone = BestEST.none
   Set phNone = BestEST.Hnone
   Set pvNone = BestEST.Vnone
   Set pRoad = BestEST.road
   Set pField = BestEST.field
   Set pForest = BestEST.forest
   Set pSwamp = BestEST.swamp
   Set pHill = BestEST.hill
   Set pMountain = BestEST.mountain
   Set pWater = BestEST.water
   Set pSpecial = BestEST.special
   BestEST.width = 9600
   BestEST.height = 6800
   BestEST.Left = 5000
   compactMode = True
End Sub

Public Sub InvalidData()
   BestEST.status.Caption = "Invalid data!"
End Sub

