Attribute VB_Name = "mapping"
Option Explicit
Public mapTerrain As Long
Public mapRide As Boolean
Public mapSun As Boolean
Public mapMonster As Boolean
Public mapDoornameNorth As String
Public mapDoornameEast As String
Public mapDoornameSouth As String
Public mapDoornameWest As String
Public mapDoornameUp As String
Public mapDoornameDown As String

Public mapHiddendoorNorth As Boolean
Public mapHiddendoorEast As Boolean
Public mapHiddendoorSouth As Boolean
Public mapHiddendoorWest As Boolean
Public mapHiddendoorUp As Boolean
Public mapHiddendoorDown As Boolean

Public mapRowNorth As Long
Public mapRowEast As Long
Public mapRowSouth As Long
Public mapRowWest As Long
Public mapRowUp As Long
Public mapRowDown As Long
Public mapColNorth As Long
Public mapColEast As Long
Public mapColSouth As Long
Public mapColWest As Long
Public mapColUp As Long
Public mapColDown As Long

Public mapExitNorth As Boolean
Public mapExitEast As Boolean
Public mapExitSouth As Boolean
Public mapExitWest As Boolean
Public mapExitUp As Boolean
Public mapExitDown As Boolean

Public mapDoorNorth As Boolean
Public mapDoorEast As Boolean
Public mapDoorSouth As Boolean
Public mapDoorWest As Boolean
Public mapDoorUp As Boolean
Public mapDoorDown As Boolean

Public mapValue As Long
Public mapDesc As String
Public mapRoomName As String
Public mapExits As String
Public mapDescription As String
Public mapCase As Long
Public MappingGetUpdate As Boolean

Public dataFromMUME As Boolean

Public Sub mapUpdate()
   If checkArrayLimit(theRow, theCol) = False Then Exit Sub
   mapValue = 0
   mapDesc = ""
   If mapSun = True Then mapValue = mapValue + 1
   If mapRide = True Then mapValue = mapValue + 2
   If mapMonster = True Then mapValue = mapValue + MONSTER_MAP
   mapValue = mapValue + mapTerrain
   mapDesc = mapDesc & ";"

   Call createData(mapValue, mapDesc, _
         mapRowNorth, mapColNorth, _
         mapExitNorth, mapDoornameNorth, mapHiddendoorNorth, _
         N_noexit, N_exit, N_door, N_hiddendoor, N_portal, N_doorportal)

   Call createData(mapValue, mapDesc, _
         mapRowEast, mapColEast, _
         mapExitEast, mapDoornameEast, mapHiddendoorEast, _
         E_noexit, E_exit, E_door, E_hiddendoor, E_portal, E_doorportal)

   Call createData(mapValue, mapDesc, _
         mapRowSouth, mapColSouth, _
         mapExitSouth, mapDoornameSouth, mapHiddendoorSouth, _
         S_noexit, S_exit, S_door, S_hiddendoor, S_portal, S_doorportal)

   Call createData(mapValue, mapDesc, _
         mapRowWest, mapColWest, _
         mapExitWest, mapDoornameWest, mapHiddendoorWest, _
         W_noexit, W_exit, W_door, W_hiddendoor, W_portal, W_doorportal)

   Call createData(mapValue, mapDesc, _
         mapRowUp, mapColUp, _
         mapExitUp, mapDoornameUp, mapHiddendoorUp, _
         U_noexit, U_exit, U_door, U_hiddendoor, U_portal, U_doorportal)
   
   Call createData(mapValue, mapDesc, _
         mapRowDown, mapColDown, _
         mapExitDown, mapDoornameDown, mapHiddendoorDown, _
         D_noexit, D_exit, D_door, D_hiddendoor, D_portal, D_doorportal)
   
   mapDesc = mapDesc & ";"
   
   If mapValue > 0 And Len(mapDesc) > 2 And Len(mapRoomName) > 0 And Len(mapDescription) > 0 Then
      arr(theRow, theCol) = mapValue
      arrDesc(theRow, theCol) = mapDesc
      If dataFromMUME Then
         arrRoomname(theRow, theCol) = mapRoomName
         arrDescription(theRow, theCol) = EncryptDesc(mapDescription)
      Else
         arrRoomname(theRow, theCol) = mapRoomName
         arrDescription(theRow, theCol) = mapDescription
      End If
      Call loadRoom(theRow, theCol)
      Call DrawMap
   Else
      MsgBox ("Invalid data. Update cancelled!")
   End If

End Sub

Public Sub setMapTerrain(ByVal what)
   With frmTools
      Select Case what
      Case "road", 0
         mapTerrain = road
         Set .pictureTerrain.Picture = pRoad
      Case "plain", 4
         mapTerrain = plain
         Set .pictureTerrain.Picture = pField
      Case "forest", 8
         mapTerrain = forest
         Set .pictureTerrain.Picture = pForest
      Case "swamp", 12
         mapTerrain = swamp
         Set .pictureTerrain.Picture = pSwamp
      Case "hill", 16
         mapTerrain = hill
         Set .pictureTerrain.Picture = pHill
      Case "mountain", 20
         mapTerrain = mountain
         Set .pictureTerrain.Picture = pMountain
      Case "water", 24
         mapTerrain = water
         Set .pictureTerrain.Picture = pWater
      Case "special", 28
         mapTerrain = special
         Set .pictureTerrain.Picture = pSpecial
      Case Else
         mapTerrain = plain
         Set .pictureTerrain.Picture = pField
      End Select
   End With
End Sub

Public Sub zeroMap()
   mapRoomName = ""
   mapDescription = ""

   mapDoornameNorth = ""
   mapDoornameEast = ""
   mapDoornameSouth = ""
   mapDoornameWest = ""
   mapDoornameUp = ""
   mapDoornameDown = ""

   mapHiddendoorNorth = False
   mapHiddendoorEast = False
   mapHiddendoorSouth = False
   mapHiddendoorWest = False
   mapHiddendoorUp = False
   mapHiddendoorDown = False

   mapRowNorth = 0
   mapRowEast = 0
   mapRowSouth = 0
   mapRowWest = 0
   mapRowUp = 0
   mapRowDown = 0
   
   mapColNorth = 0
   mapColEast = 0
   mapColSouth = 0
   mapColWest = 0
   mapColUp = 0
   mapColDown = 0
   
   mapExitNorth = False
   mapExitEast = False
   mapExitSouth = False
   mapExitWest = False
   mapExitUp = False
   mapExitDown = False
   
   mapDoorNorth = False
   mapDoorEast = False
   mapDoorSouth = False
   mapDoorWest = False
   mapDoorUp = False
   mapDoorDown = False
   
   With frmTools
      .GoToNumber.Text = ""
      .Roomname.Text = ""
      .Description.Text = ""
      .Monster.Value = 0
      .nExit.Value = 0
      .nDoor.Text = ""
      .nHidden.Value = 0
      .nPortal.Text = ""
      .eExit.Value = 0
      .eDoor.Text = ""
      .eHidden.Value = 0
      .ePortal.Text = ""
      .sExit.Value = 0
      .sDoor.Text = ""
      .sHidden.Value = 0
      .sPortal.Text = ""
      .wExit.Value = 0
      .wDoor.Text = ""
      .wHidden.Value = 0
      .wPortal.Text = ""
      .uExit.Value = 0
      .uDoor.Text = ""
      .uHidden.Value = 0
      .uPortal.Text = ""
      .dExit.Value = 0
      .dDoor.Text = ""
      .dHidden.Value = 0
      .dPortal.Text = ""
   End With
End Sub

Public Sub setMapModeON()
   Set frmTools.pictureTerrain.Picture = pRoad
   mapTerrain = 0
   Call zeroMap
   MappingMode = True
   MappingData = False
   Out_Of_Sync = True
   roomCount = 0
End Sub

Public Sub setMapModeOFF()
   MappingMode = False
   MappingData = False
   Out_Of_Sync = True
   roomCount = 0
End Sub

Public Sub gotoArea(ByRef data)
On Error GoTo ErrorHandler
Dim currentRoom
   If Len(data) < 3 Then Err.Raise 555
      currentRoom = Split(data, ",")
      If UBound(currentRoom) > 0 Then
         theRow = currentRoom(0)
         theCol = currentRoom(1)
         Call loadRoom(theRow, theCol)
         Call DrawMap
      End If
Exit Sub
ErrorHandler:
   Call InvalidData
End Sub

Public Sub getRoomData()
On Error GoTo ErrorHandler
   Call zeroMap
   If checkArrayLimit(theRow, theCol) = True Then
      dataFromMUME = False
      MappingData = True
      MappingCase = 1
      frmTools.status.Caption = "Getting room data ..."
      frmMap.tcpClient.SendData "EXAMINE" & vbLf
   End If
Exit Sub
ErrorHandler:
   frmTools.status.Caption = "Proxy server down..."
End Sub

Public Sub GetMapData()
   dataFromMUME = False
   With frmTools
      mapRoomName = theRoomname
      .Roomname.Text = mapRoomName
      mapDescription = theRoomdesc
      .Description.Text = mapDescription
      setMapTerrain (theTerrain)
      mapSun = theSun
      If mapSun Then
         .Sun.Value = 1
      Else
         .Sun.Value = 0
      End If
      mapRide = theRide
      If mapRide Then
         .Ridable.Value = 1
      Else
         .Ridable.Value = 0
      End If
      If theMonster Then
         mapMonster = True
         .Monster = 1
      End If
      If mapMonster Then
         .Monster.Value = 1
      Else
         .Monster.Value = 0
      End If
      If theExitNorth Then
         mapExitNorth = True
         .nExit.Value = 1
      Else
         mapExitNorth = False
         .nExit.Value = 0
      End If
      If theExitEast Then
         mapExitEast = True
         .eExit.Value = 1
      Else
         mapExitEast = False
         .eExit.Value = 0
      End If
      If theExitSouth Then
         mapExitSouth = True
         .sExit.Value = 1
      Else
         mapExitSouth = False
         .sExit.Value = 0
      End If
      If theExitWest Then
         mapExitWest = True
         .wExit.Value = 1
      Else
         mapExitWest = False
         .wExit.Value = 0
      End If
      If theExitUp Then
         mapExitUp = True
         .uExit.Value = 1
      Else
         mapExitUp = False
         .uExit.Value = 0
      End If
      If theExitDown Then
         mapExitDown = True
         .dExit.Value = 1
      Else
         mapExitDown = False
         .dExit.Value = 0
      End If

'#########################
      
      If theDoorNorth Or theHiddendoorNorth Then
         mapDoornameNorth = theDoornameNorth
         .nDoor.Text = theDoornameNorth
         .nHidden.Visible = True
         If theHiddendoorNorth Then
            .nHidden.Value = 1
            mapHiddendoorNorth = True
         Else
            .nHidden.Value = 0
            mapHiddendoorNorth = False
         End If
      Else
         .nHidden.Value = 0
         .nHidden.Visible = False
         mapHiddendoorNorth = False
         mapDoornameNorth = ""
         .nDoor.Text = ""
      End If
      
      If theDoorEast Or theHiddendoorEast Then
         mapDoornameEast = theDoornameEast
         .eDoor.Text = theDoornameEast
         .eHidden.Visible = True
         If theHiddendoorEast Then
            .eHidden.Value = 1
            mapHiddendoorEast = True
         Else
            .eHidden.Value = 0
            mapHiddendoorEast = False
         End If
      Else
         .eHidden.Value = 0
         .eHidden.Visible = False
         mapHiddendoorEast = False
         mapDoornameEast = ""
         .eDoor.Text = ""
      End If
      
      If theDoorSouth Or theHiddendoorSouth Then
         mapDoornameSouth = theDoornameSouth
         .sDoor.Text = theDoornameSouth
         .sHidden.Visible = True
         If theHiddendoorSouth Then
            .sHidden.Value = 1
            mapHiddendoorSouth = True
         Else
            .sHidden.Value = 0
            mapHiddendoorSouth = False
         End If
      Else
         .sHidden.Value = 0
         .sHidden.Visible = False
         mapHiddendoorSouth = False
         mapDoornameSouth = ""
         .sDoor.Text = ""
      End If
      
      If theDoorWest Or theHiddendoorWest Then
         mapDoornameWest = theDoornameWest
         .wDoor.Text = theDoornameWest
         .wHidden.Visible = True
         If theHiddendoorWest Then
            .wHidden.Value = 1
            mapHiddendoorWest = True
         Else
            .wHidden.Value = 0
            mapHiddendoorWest = False
         End If
      Else
         .wHidden.Value = 0
         .wHidden.Visible = False
         mapHiddendoorWest = False
         mapDoornameWest = ""
         .wDoor.Text = ""
      End If
      
      If theDoorUp Or theHiddendoorUp Then
         mapDoornameUp = theDoornameUp
         .uDoor.Text = theDoornameUp
         .uHidden.Visible = True
         If theHiddendoorUp Then
            .uHidden.Value = 1
            mapHiddendoorUp = True
         Else
            .uHidden.Value = 0
            mapHiddendoorUp = False
         End If
      Else
         .uHidden.Value = 0
         .uHidden.Visible = False
         mapHiddendoorUp = False
         mapDoornameUp = ""
         .uDoor.Text = ""
      End If

      If theDoorDown Or theHiddendoorDown Then
         mapDoornameDown = theDoornameDown
         .dDoor.Text = theDoornameDown
         .dHidden.Visible = True
         If theHiddendoorDown Then
            .dHidden.Value = 1
            mapHiddendoorDown = True
         Else
            .dHidden.Value = 0
            mapHiddendoorDown = False
         End If
      Else
         .dHidden.Value = 0
         .dHidden.Visible = False
         mapHiddendoorDown = False
         mapDoornameDown = ""
         .dDoor.Text = ""
      End If

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
      If thePortalNorth Or theDoorPortalNorth Then
         mapRowNorth = theRowNorth
         mapColNorth = theColNorth
         .nPortal.Text = theRowNorth & "," & theColNorth
      Else
         mapRowNorth = 0
         mapColNorth = 0
         .nPortal.Text = ""
      End If
      If thePortalEast Or theDoorPortalEast Then
         mapRowEast = theRowEast
         mapColEast = theColEast
         .ePortal.Text = theRowEast & "," & theColEast
      Else
         mapRowEast = 0
         mapColEast = 0
         .ePortal.Text = ""
      End If
      If thePortalSouth Or theDoorPortalSouth Then
         mapRowSouth = theRowSouth
         mapColSouth = theColSouth
         .sPortal.Text = theRowSouth & "," & theColSouth
      Else
         mapRowSouth = 0
         mapColSouth = 0
         .sPortal.Text = ""
      End If
      If thePortalWest Or theDoorPortalWest Then
         mapRowWest = theRowWest
         mapColWest = theColWest
         .wPortal.Text = theRowWest & "," & theColWest
      Else
         mapRowWest = 0
         mapColWest = 0
         .wPortal.Text = ""
      End If
      If thePortalUp Or theDoorPortalUp Then
         mapRowUp = theRowUp
         mapColUp = theColUp
         .uPortal.Text = theRowUp & "," & theColUp
      Else
         mapRowUp = 0
         mapColUp = 0
         .uPortal.Text = ""
      End If
      If thePortalDown Or theDoorPortalDown Then
         mapRowDown = theRowDown
         mapColDown = theColDown
         .dPortal.Text = theRowDown & "," & theColDown
      Else
         mapRowDown = 0
         mapColDown = 0
         .dPortal.Text = ""
      End If
   End With
End Sub

Public Sub InvalidData()
   frmTools.status.Caption = "Invalid data!"
End Sub

