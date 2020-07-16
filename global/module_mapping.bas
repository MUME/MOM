Attribute VB_Name = "mapping"
Option Explicit
Public mapTerrain As Long
Public mapRoad As Long

Public mapFlag As Long
Public mapRide As Boolean
Public mapSun As Boolean
Public mapMonster As Boolean
Public mapMonsterValue As Long
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
Public mapCase As Integer

Public MappingGetUpdate As Boolean
Public dataFromMUD As Boolean
Public wasMapMode As Boolean
Public isBriefMode As Boolean
Public isSpamMode As Boolean

Public Sub mapUpdate()
errorData = errorData & "mapUpdate -> "
If isValid(theROW, theCOL) = False Then Exit Sub
If getIndex(theROW, theCOL) = 0 Then
   If canIncreaseTheCount Then theCount = theCount + 1
   aWorld(theROW, theCOL, theLEVEL) = theCount
End If

'with warp, when the new room is not empty and when walkthrough mode is on
'If MUDname = "WARP" And LenB(aData(getIndex(theROW, theCOL), cDATA)) <> 0 And frmMap.mnuWalkthrough.Checked Then

'NEW ROOM OR UPDATE
If LenB(aData(getIndex(theROW, theCOL), cDATA)) = 0 Or frmMap.mnuWalkthrough.Checked = False Then
   mapValue = 0
   If mapSun = True Then mapValue = mapValue Or 1
   If mapRide = True Then mapValue = mapValue Or 2
   mapValue = mapValue Or mapMonsterValue
   mapValue = mapValue Or mapTerrain
   mapValue = mapValue Or mapFlag
   
   mapValue = mapValue Or mapRoad
   If checkRoad(mapRoomName) = True Then mapValue = mapValue Or mapRoad
   
   Call createData(mapValue, _
         mapRowNorth, mapColNorth, _
         mapExitNorth, mapDoornameNorth, mapHiddendoorNorth, frmTools.nVisible, _
         N_noexit, N_exit, N_door, N_hiddendoor, N_portal, N_doorportal)

   Call createData(mapValue, _
         mapRowEast, mapColEast, _
         mapExitEast, mapDoornameEast, mapHiddendoorEast, frmTools.eVisible, _
         E_noexit, E_exit, E_door, E_hiddendoor, E_portal, E_doorportal)

   Call createData(mapValue, _
         mapRowSouth, mapColSouth, _
         mapExitSouth, mapDoornameSouth, mapHiddendoorSouth, frmTools.sVisible, _
         S_noexit, S_exit, S_door, S_hiddendoor, S_portal, S_doorportal)

   Call createData(mapValue, _
         mapRowWest, mapColWest, _
         mapExitWest, mapDoornameWest, mapHiddendoorWest, frmTools.wVisible, _
         W_noexit, W_exit, W_door, W_hiddendoor, W_portal, W_doorportal)

   Call createData(mapValue, _
         mapRowUp, mapColUp, _
         mapExitUp, mapDoornameUp, mapHiddendoorUp, frmTools.uVisible, _
         U_noexit, U_exit, U_door, U_hiddendoor, U_portal, U_doorportal)

   Call createData(mapValue, _
         mapRowDown, mapColDown, _
         mapExitDown, mapDoornameDown, mapHiddendoorDown, frmTools.dVisible, _
         D_noexit, D_exit, D_door, D_hiddendoor, D_portal, D_doorportal)

   If mapValue > 0 And LenB(mapRoomName) <> 0 Then
      aData(getIndex(theROW, theCOL), cDATA) = mapValue
      aData(getIndex(theROW, theCOL), cROW) = theROW
      aData(getIndex(theROW, theCOL), cCOL) = theCOL
      aData(getIndex(theROW, theCOL), cROOMNAME) = mapRoomName
      If dataFromMUD Then
         aData(getIndex(theROW, theCOL), cDESCRIPTION) = CRC32(mapDescription)
      Else
         aData(getIndex(theROW, theCOL), cDESCRIPTION) = mapDescription
      End If
      
      If LenB(aData(getIndex(theROW, theCOL), cDESCRIPTION)) = 0 Then
         Call informClient("Invalid description! Mapping cancelled, turn BRIEF MODE OFF.")
         Exit Sub
      End If
      
      aData(getIndex(theROW, theCOL), cNDOOR) = Trim$(mapDoornameNorth)
      aData(getIndex(theROW, theCOL), cEDOOR) = Trim$(mapDoornameEast)
      aData(getIndex(theROW, theCOL), cSDOOR) = Trim$(mapDoornameSouth)
      aData(getIndex(theROW, theCOL), cWDOOR) = Trim$(mapDoornameWest)
      aData(getIndex(theROW, theCOL), cUDOOR) = Trim$(mapDoornameUp)
      aData(getIndex(theROW, theCOL), cDDOOR) = Trim$(mapDoornameDown)
      aData(getIndex(theROW, theCOL), cNOTE) = aData(getIndex(theROW, theCOL), cNOTE)
      
      'määrame ruumile mäppimise hetkel oleva taseme
      aData(getIndex(theROW, theCOL), cLEVEL) = theLEVEL
      
      If dataFromMUD Then
         'Debug.Print "OLD: '" & mappingFromDir & "', " & oldRow & ", " & oldCol & "; Level(" & oldLevel & ")"
         'create portal between old and new room
         Call createportalMAPPING(getIndex(theROW, theCOL), mappingFromDir)
      Else
         
'         If (getData(cursor) And N_MAP) > 0 Then
'            aData(cursor, cNPORTALR) = aData(cursor, cROW) - 1: aData(cursor, cNPORTALC) = aData(cursor, cCOL): aData(cursor, cNLEVEL) = theLEVEL
'         End If
'         If (getData(cursor) And E_MAP) > 0 Then
'            aData(cursor, cEPORTALR) = aData(cursor, cROW): aData(cursor, cEPORTALC) = aData(cursor, cCOL) + 1: aData(cursor, cELEVEL) = theLEVEL
'         End If
'         If (getData(cursor) And S_MAP) > 0 Then
'            aData(cursor, cSPORTALR) = aData(cursor, cROW) + 1: aData(cursor, cSPORTALC) = aData(cursor, cCOL): aData(cursor, cSLEVEL) = theLEVEL
'         End If
'         If (getData(cursor) And W_MAP) > 0 Then
'            aData(cursor, cWPORTALR) = aData(cursor, cROW): aData(cursor, cWPORTALC) = aData(cursor, cCOL) - 1: aData(cursor, cWLEVEL) = theLEVEL
'         End If
         
         'set current room portal coordinates
         cursor = getIndex(theROW, theCOL)
         If mapRowNorth <> 0 Then aData(cursor, cNPORTALR) = mapRowNorth
         If mapColNorth <> 0 Then aData(cursor, cNPORTALC) = mapColNorth
         If mapRowEast <> 0 Then aData(cursor, cEPORTALR) = mapRowEast
         If mapColEast <> 0 Then aData(cursor, cEPORTALC) = mapColEast
         If mapRowSouth <> 0 Then aData(cursor, cSPORTALR) = mapRowSouth
         If mapColSouth <> 0 Then aData(cursor, cSPORTALC) = mapColSouth
         If mapRowWest <> 0 Then aData(cursor, cWPORTALR) = mapRowWest
         If mapColWest <> 0 Then aData(cursor, cWPORTALC) = mapColWest
         If mapRowUp <> 0 Then aData(cursor, cUPORTALR) = mapRowUp
         If mapColUp <> 0 Then aData(cursor, cUPORTALC) = mapColUp
         If mapRowDown <> 0 Then aData(cursor, cDPORTALR) = mapRowDown
         If mapColDown <> 0 Then aData(cursor, cDPORTALC) = mapColDown
      End If
      
      Call updateThis(getIndex(theROW, theCOL))
      Call loadRoom(theROW, theCOL)
      Call GetMapData
      Call DrawMap
   Else
      Call informClient("Invalid data. Map update cancelled!")
   End If

'SPECIAL UPDATE ROOM
Else
   mapValue = aData(getIndex(theROW, theCOL), cDATA)
   mapValue = (mapValue Or TERRAIN_MAP)
   mapValue = (mapValue Xor TERRAIN_MAP)
   mapValue = (mapValue Or mapTerrain)
   If mapValue > 0 And LenB(mapRoomName) <> 0 Then
      aData(getIndex(theROW, theCOL), cDATA) = mapValue
      aData(getIndex(theROW, theCOL), cROW) = theROW
      aData(getIndex(theROW, theCOL), cCOL) = theCOL
      aData(getIndex(theROW, theCOL), cROOMNAME) = mapRoomName
      
      If dataFromMUD Then
         aData(getIndex(theROW, theCOL), cDESCRIPTION) = CRC32(mapDescription)
      Else
         aData(getIndex(theROW, theCOL), cDESCRIPTION) = mapDescription
      End If
      
      If LenB(aData(getIndex(theROW, theCOL), cDESCRIPTION)) = 0 Then Call informClient("Invalid description! Mapping cancelled (WARP => turn brief mode off!")
      Call updateThis(getIndex(theROW, theCOL))
      Call loadRoom(theROW, theCOL)
      Call GetMapData
      Call DrawMap
   Else
      Call informClient("Invalid data. Map update cancelled!")
   End If
End If

End Sub

Public Sub setMapTerrain(ByRef s As String)
   errorData = errorData & "setMapTerrain -> "

If MUDname = "MUME" Then
   Select Case s
   Case vbNullString                 'terrain remains the previous selected
      mapTerrain = mapTerrain
   Case "+" 'road, trail
      mapTerrain = plain
      mapRoad = ISROAD
   Case "+", ".", ":", "road", "plain", 0, 4 'road, plain, brush
      mapTerrain = plain
   Case "forest", 8, "f"   'forest
      mapTerrain = forest
   Case "swamp", 12, "%"   'swamp, shallow water
      mapTerrain = swamp
   Case "hill", 16, "("    'hills
      mapTerrain = hill
   Case "underground", 20, "[", "=", "O", "#" 'indoors, city, tunnel, cavern
      mapTerrain = underground
   Case "water", 24, "~", "W", "U" 'water, rapids, underwater
      mapTerrain = water
   Case "mountain", 28, "<"     'mountain
      mapTerrain = mountain
   Case "city", city ' gates icon
      mapTerrain = city
   'Case "bridge", bridge
   '   mapTerrain = bridge
   Case "shop", shop
      mapTerrain = shop
   Case "guild", guild
      mapTerrain = guild
   Case "inn", inn
      mapTerrain = inn
   Case "dungeon", dungeon
      mapTerrain = dungeon
   Case Else
      mapTerrain = plain
   End Select
   
   
   
   
   
Else 'warp
   Select Case s
   Case vbNullString                 'terrain remains the previous selected
      mapTerrain = mapTerrain
   Case "road", 0, "=", "-", "+"   'road, trail
      mapTerrain = road
   Case "plain", 4, ":", "b" 'field, brush
      mapTerrain = plain
   Case "forest", 8, "%"   'forest
      mapTerrain = forest
   Case "swamp", 12, "*"   'swamp, shallow water
      mapTerrain = swamp
   Case "hill", 16, "("    'hills
      mapTerrain = hill
   Case "underground", 20, "#", "[", "$", "O" 'indoors, city, tunnel, cavern
      mapTerrain = underground
   Case "water", 24, "W", "U", "~" 'water, rapids, underwater
      mapTerrain = water
   Case "mountain", 28, "^"     'mountain
      mapTerrain = mountain
   Case "city", city
      mapTerrain = city
   Case "shop", shop
      mapTerrain = shop
   Case "guild", guild
      mapTerrain = guild
   Case "inn", inn
      mapTerrain = inn
   Case "dungeon", dungeon
      mapTerrain = dungeon
   Case Else
      mapTerrain = plain
   End Select
End If
End Sub

Public Sub setMapFlag(ByVal what)
   errorData = errorData & "setRoomFlag-> "
   Select Case what
   Case "water", FLAG_WATER
      mapFlag = FLAG_WATER
   Case "item", FLAG_ITEM
      mapFlag = FLAG_ITEM
   Case "herb", FLAG_HERB
      mapFlag = FLAG_HERB
   Case "treasury", FLAG_TREASURY
      mapFlag = FLAG_TREASURY
   Case "key", FLAG_KEY
      mapFlag = FLAG_KEY
   Case "magic", FLAG_MAGIC
      mapFlag = FLAG_MAGIC
   Case "quest", FLAG_QUEST
      mapFlag = FLAG_QUEST
   Case Else
      mapFlag = FLAG_NONE
   End Select
End Sub

Public Sub zeroMap()
   mapRoad = 0
   mapFlag = 0
   mapRoomName = vbNullString
   mapDescription = vbNullString
   mapMonster = False

   mapDoornameNorth = vbNullString
   mapDoornameEast = vbNullString
   mapDoornameSouth = vbNullString
   mapDoornameWest = vbNullString
   mapDoornameUp = vbNullString
   mapDoornameDown = vbNullString

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
      .Roomname.Caption = vbNullString
      'portal visibility
      .nVisible.value = 0
      .eVisible.value = 0
      .sVisible.value = 0
      .wVisible.value = 0
      .uVisible.value = 0
      .dVisible.value = 0
      'exits
      .nExit.value = 0
      .eExit.value = 0
      .sExit.value = 0
      .wExit.value = 0
      .uExit.value = 0
      .dExit.value = 0
      'hiddendoors
      .nHidden.value = 0
      .eHidden.value = 0
      .sHidden.value = 0
      .wHidden.value = 0
      .uHidden.value = 0
      .dHidden.value = 0
      'doornames
      .nDoor.text = vbNullString
      .eDoor.text = vbNullString
      .sDoor.text = vbNullString
      .wDoor.text = vbNullString
      .uDoor.text = vbNullString
      .dDoor.text = vbNullString
      'portal coordinates
      .nPortal.text = vbNullString
      .ePortal.text = vbNullString
      .sPortal.text = vbNullString
      .wPortal.text = vbNullString
      .uPortal.text = vbNullString
      .dPortal.text = vbNullString
   End With
End Sub

Public Sub setMapModeON()
On Error Resume Next
   frmMap.mnuEdit.Enabled = True
   frmMap.mnuEdit.Visible = True
   frmMap.tcpPlayer.SendData ("PROMPT ALL") & vbCrLf
   frmMap.tcpPlayer.SendData ("BRIEF OFF") & vbCrLf
   If MUDname = "MUME" Then frmMap.tcpPlayer.SendData ("SPAM ON") & vbCrLf
   wasMapMode = True

On Error GoTo 0
   mapTerrain = 0
   mapFlag = 0
   mapRoad = 0
   Call zeroMap
   MappingMode = True
   MappingData = False
   Call SYNC_FALSE("Mapping mode!")
   roomcount = 0
End Sub

Public Sub setMapModeOFF()
   frmMap.mnuEdit.Enabled = True
   frmMap.mnuEdit.Visible = True
   MappingMode = False
   MappingData = False
   roomcount = 0
End Sub

Public Sub gotoArea(data)
errorData = errorData & "gotoArea -> "
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(data) > 5 And InStrB(1, data, ",", vbBinaryCompare) > 0 Then
      Dim currentRoom
      currentRoom = Split(data, ",", 2, vbBinaryCompare)
      If UBound(currentRoom) = 1 Then
         theROW = currentRoom(0)
         theCOL = currentRoom(1)
         Call loadRoom(theROW, theCOL)
         Call DrawMap
      End If
   End If
Exit Sub
errorhandler:
   Call InvalidData
End Sub

Public Sub getRoomData()
errorData = errorData & "getRoomData -> "
   Call zeroMap
   If isValid(theROW, theCOL) = True Then
      dataFromMUD = True
      MappingData = True
      MappingCase = 1
      frmMap.tcpPlayer.SendData "EXAMINE" & vbCrLf
      retry = 0
   End If
End Sub

Public Sub GetMapData()
   errorData = errorData & "getMapData -> "
   Call zeroMap
   dataFromMUD = False
   
   With frmTools
      mapRoomName = theRoomname
      .Roomname.Caption = mapRoomName
      mapDescription = theRoomdesc
      .coordinates.Caption = "(" & theROW & "," & theCOL & ")"
      setMapTerrain (theTerrain)
      setMapFlag (theFlag)
      If (getData(getIndex(theROW, theCOL)) And ISROAD) = ISROAD Then
         mapRoad = theRoad
      Else
         mapRoad = 0
      End If
      mapSun = theSun
      If mapSun Then
         .Sun.value = 1
      Else
         .Sun.value = 0
      End If
      mapRide = theRide
      If mapRide Then
         .Ridable.value = 1
      Else
         .Ridable.value = 0
      End If
      
      
      If theExitNorth Then
         mapExitNorth = True
         .nExit.value = 1
         mapRowNorth = theROWNorth
         mapColNorth = theCOLNorth
         .nPortal.text = theROWNorth & "," & theCOLNorth
         If thePortalNorth Or theDoorPortalNorth Then .nVisible.value = 1
      End If
      
      If theExitEast Then
         mapExitEast = True
         .eExit.value = 1
         mapRowEast = theROWEast
         mapColEast = theCOLEast
         .ePortal.text = theROWEast & "," & theCOLEast
         If thePortalEast Or theDoorPortalEast Then .eVisible.value = 1
      End If
      
      If theExitSouth Then
         mapExitSouth = True
         .sExit.value = 1
         mapRowSouth = theROWSouth
         mapColSouth = theCOLSouth
         .sPortal.text = theROWSouth & "," & theCOLSouth
         If thePortalSouth Or theDoorPortalSouth Then .sVisible.value = 1
      End If
      
      If theExitWest Then
         mapExitWest = True
         .wExit.value = 1
         mapRowWest = theROWWest
         mapColWest = theCOLWest
         .wPortal.text = theROWWest & "," & theCOLWest
         If thePortalWest Or theDoorPortalWest Then .wVisible.value = 1
      End If
      
      If theExitUp Then
         mapExitUp = True
         .uExit.value = 1
         mapRowUp = theROWUp
         mapColUp = theCOLUp
         .uPortal.text = theROWUp & "," & theCOLUp
         If thePortalUp Or theDoorPortalUp Then .uVisible.value = 1
      End If
      
      If theExitDown Then
         mapExitDown = True
         .dExit.value = 1
         mapRowDown = theROWDown
         mapColDown = theCOLDown
         .dPortal.text = theROWDown & "," & theCOLDown
         If thePortalDown Or theDoorPortalDown Then .dVisible.value = 1
      End If
      
      
      
      If theDoorNorth Or theHiddendoorNorth Then
         mapDoornameNorth = theDoornameNorth
         .nDoor.text = theDoornameNorth
         .nHidden.Visible = True
         If theHiddendoorNorth Then
            .nHidden.value = 1
            mapHiddendoorNorth = True
         Else
            .nHidden.value = 0
            mapHiddendoorNorth = False
         End If
      Else
         .nHidden.value = 0
         .nHidden.Visible = False
         mapHiddendoorNorth = False
         mapDoornameNorth = vbNullString
         .nDoor.text = vbNullString
      End If
      
      If theDoorEast Or theHiddendoorEast Then
         mapDoornameEast = theDoornameEast
         .eDoor.text = theDoornameEast
         .eHidden.Visible = True
         If theHiddendoorEast Then
            .eHidden.value = 1
            mapHiddendoorEast = True
         Else
            .eHidden.value = 0
            mapHiddendoorEast = False
         End If
      Else
         .eHidden.value = 0
         .eHidden.Visible = False
         mapHiddendoorEast = False
         mapDoornameEast = vbNullString
         .eDoor.text = vbNullString
      End If
      
      If theDoorSouth Or theHiddendoorSouth Then
         mapDoornameSouth = theDoornameSouth
         .sDoor.text = theDoornameSouth
         .sHidden.Visible = True
         If theHiddendoorSouth Then
            .sHidden.value = 1
            mapHiddendoorSouth = True
         Else
            .sHidden.value = 0
            mapHiddendoorSouth = False
         End If
      Else
         .sHidden.value = 0
         .sHidden.Visible = False
         mapHiddendoorSouth = False
         mapDoornameSouth = vbNullString
         .sDoor.text = vbNullString
      End If
      
      If theDoorWest Or theHiddendoorWest Then
         mapDoornameWest = theDoornameWest
         .wDoor.text = theDoornameWest
         .wHidden.Visible = True
         If theHiddendoorWest Then
            .wHidden.value = 1
            mapHiddendoorWest = True
         Else
            .wHidden.value = 0
            mapHiddendoorWest = False
         End If
      Else
         .wHidden.value = 0
         .wHidden.Visible = False
         mapHiddendoorWest = False
         mapDoornameWest = vbNullString
         .wDoor.text = vbNullString
      End If
      
      If theDoorUp Or theHiddendoorUp Then
         mapDoornameUp = theDoornameUp
         .uDoor.text = theDoornameUp
         .uHidden.Visible = True
         If theHiddendoorUp Then
            .uHidden.value = 1
            mapHiddendoorUp = True
         Else
            .uHidden.value = 0
            mapHiddendoorUp = False
         End If
      Else
         .uHidden.value = 0
         .uHidden.Visible = False
         mapHiddendoorUp = False
         mapDoornameUp = vbNullString
         .uDoor.text = vbNullString
      End If

      If theDoorDown Or theHiddendoorDown Then
         mapDoornameDown = theDoornameDown
         .dDoor.text = theDoornameDown
         .dHidden.Visible = True
         If theHiddendoorDown Then
            .dHidden.value = 1
            mapHiddendoorDown = True
         Else
            .dHidden.value = 0
            mapHiddendoorDown = False
         End If
      Else
         .dHidden.value = 0
         .dHidden.Visible = False
         mapHiddendoorDown = False
         mapDoornameDown = vbNullString
         .dDoor.text = vbNullString
      End If






''''      If thePortalNorth Or theDoorPortalNorth Then
''''         mapRowNorth = theROWNorth
''''         mapColNorth = theCOLNorth
''''         .nPortal.text = theROWNorth & "," & theCOLNorth
''''      Else
''''         mapRowNorth = 0
''''         mapColNorth = 0
''''         .nPortal.text = vbNullString
''''      End If
''''      If thePortalEast Or theDoorPortalEast Then
''''         mapRowEast = theROWEast
''''         mapColEast = theCOLEast
''''         .ePortal.text = theROWEast & "," & theCOLEast
''''      Else
''''         mapRowEast = 0
''''         mapColEast = 0
''''         .ePortal.text = vbNullString
''''      End If
''''      If thePortalSouth Or theDoorPortalSouth Then
''''         mapRowSouth = theROWSouth
''''         mapColSouth = theCOLSouth
''''         .sPortal.text = theROWSouth & "," & theCOLSouth
''''      Else
''''         mapRowSouth = 0
''''         mapColSouth = 0
''''         .sPortal.text = vbNullString
''''      End If
''''      If thePortalWest Or theDoorPortalWest Then
''''         mapRowWest = theROWWest
''''         mapColWest = theCOLWest
''''         .wPortal.text = theROWWest & "," & theCOLWest
''''      Else
''''         mapRowWest = 0
''''         mapColWest = 0
''''         .wPortal.text = vbNullString
''''      End If
''''      If thePortalUp Or theDoorPortalUp Then
''''         mapRowUp = theROWUp
''''         mapColUp = theCOLUp
''''         .uPortal.text = theROWUp & "," & theCOLUp
''''      Else
''''         mapRowUp = 0
''''         mapColUp = 0
''''         .uPortal.text = vbNullString
''''      End If
''''      If thePortalDown Or theDoorPortalDown Then
''''         mapRowDown = theROWDown
''''         mapColDown = theCOLDown
''''         .dPortal.text = theROWDown & "," & theCOLDown
''''      Else
''''         mapRowDown = 0
''''         mapColDown = 0
''''         .dPortal.text = vbNullString
''''      End If
   End With
End Sub

Public Sub InvalidData()
   frmTools.Caption = "Tools - " & "Invalid data!"
End Sub

Public Function canIncreaseTheCount() As Boolean
   canIncreaseTheCount = False
   Dim cursor As Integer
   Dim freeSlot As Integer
   Dim fld As Integer
   Dim w As Integer
   freeSlot = 0
   If theCount >= arrMaxData Then
      informClient ("Revising stack, please wait!")
      For cursor = 1 To UBound(aData)
         If LenB(aData(cursor, cDATA)) = 0 Then
            freeSlot = cursor
            Exit For
         End If
      Next
      If freeSlot = 0 Then
         frmMap.Caption = "Arda is full, - sail West?"
         canIncreaseTheCount = False
         Exit Function
      Else
         theCount = freeSlot - 1
         For cursor = freeSlot + 1 To UBound(aData)
            If LenB(aData(cursor, cDATA)) <> 0 Then
               aWorld(aData(cursor, cROW), aData(cursor, cCOL), theLEVEL) = freeSlot
               For fld = LBound(aData, 2) To UBound(aData, 2)
                  aData(freeSlot, fld) = aData(cursor, fld)
               Next
               freeSlot = freeSlot + 1
            End If
         Next
         For cursor = freeSlot To UBound(aData)
            For fld = LBound(aData, 2) To UBound(aData, 2)
               aData(cursor, fld) = vbNullString
            Next
         Next
         theCount = freeSlot - 1
         For w = theCount + 1 To UBound(aData)
            aData(w, cDATA) = vbNullString
         Next
      End If
   Else
      canIncreaseTheCount = True
   End If
End Function

Public Sub clearArraySlot(row As Integer, col As Integer)
   aData(aWorld(row, col, selLevel), cENCRYPTED) = vbNullString
   aData(aWorld(row, col, selLevel), cROW) = 0
   aData(aWorld(row, col, selLevel), cCOL) = 0
   aData(aWorld(row, col, selLevel), cDATA) = vbNullString
   aData(aWorld(row, col, selLevel), cROOMNAME) = vbNullString
   aData(aWorld(row, col, selLevel), cDESCRIPTION) = vbNullString
   aData(aWorld(row, col, selLevel), cNPORTALR) = 0
   aData(aWorld(row, col, selLevel), cEPORTALR) = 0
   aData(aWorld(row, col, selLevel), cSPORTALR) = 0
   aData(aWorld(row, col, selLevel), cWPORTALR) = 0
   aData(aWorld(row, col, selLevel), cUPORTALR) = 0
   aData(aWorld(row, col, selLevel), cDPORTALR) = 0
   aData(aWorld(row, col, selLevel), cNPORTALC) = 0
   aData(aWorld(row, col, selLevel), cEPORTALC) = 0
   aData(aWorld(row, col, selLevel), cSPORTALC) = 0
   aData(aWorld(row, col, selLevel), cWPORTALC) = 0
   aData(aWorld(row, col, selLevel), cUPORTALC) = 0
   aData(aWorld(row, col, selLevel), cDPORTALC) = 0
   aData(aWorld(row, col, selLevel), cNDOOR) = vbNullString
   aData(aWorld(row, col, selLevel), cEDOOR) = vbNullString
   aData(aWorld(row, col, selLevel), cSDOOR) = vbNullString
   aData(aWorld(row, col, selLevel), cWDOOR) = vbNullString
   aData(aWorld(row, col, selLevel), cUDOOR) = vbNullString
   aData(aWorld(row, col, selLevel), cDDOOR) = vbNullString
   aData(aWorld(row, col, selLevel), cNOTE) = vbNullString
   aData(aWorld(row, col, selLevel), cLEVEL) = 0
   aData(aWorld(row, col, selLevel), cNLEVEL) = 0
   aData(aWorld(row, col, selLevel), cELEVEL) = 0
   aData(aWorld(row, col, selLevel), cSLEVEL) = 0
   aData(aWorld(row, col, selLevel), cWLEVEL) = 0
   aData(aWorld(row, col, selLevel), cULEVEL) = 0
   aData(aWorld(row, col, selLevel), cDLEVEL) = 0

   aWorld(row, col, selLevel) = 0
End Sub

Public Sub updateThis(ByRef index As Integer)
'.......................................................
Dim key As Variant
'convert case also
key = thekeymaker

Dim original As Variant
Dim encrypted As Variant

'set old values to link the portals
oldLevel = theLEVEL
oldRow = theROW
oldCol = theCOL

encrypted = cast128.cast128encode(key, _
   aData(index, cDATA) & ";" & _
   aData(index, cDESCRIPTION) & ";" & _
   aData(index, cROW) & ";" & _
   aData(index, cCOL))

aData(index, cENCRYPTED) = encrypted & ";" & aData(index, cROOMNAME) & ";" & _
   aData(index, cNDOOR) & ";" & aData(index, cEDOOR) & ";" & aData(index, cSDOOR) & ";" & aData(index, cWDOOR) & ";" & aData(index, cUDOOR) & ";" & aData(index, cDDOOR) & ";" & _
   aData(index, cNPORTALR) & ";" & aData(index, cNPORTALC) & ";" & _
   aData(index, cEPORTALR) & ";" & aData(index, cEPORTALC) & ";" & _
   aData(index, cSPORTALR) & ";" & aData(index, cSPORTALC) & ";" & _
   aData(index, cWPORTALR) & ";" & aData(index, cWPORTALC) & ";" & _
   aData(index, cUPORTALR) & ";" & aData(index, cUPORTALC) & ";" & _
   aData(index, cDPORTALR) & ";" & aData(index, cDPORTALC) & ";" & _
   aData(index, cNOTE) & ";" & _
   aData(index, cLEVEL) & ";" & _
   aData(index, cNLEVEL) & ";" & aData(index, cELEVEL) & ";" & aData(index, cSLEVEL) & ";" & aData(index, cWLEVEL) & ";" & aData(index, cULEVEL) & ";" & aData(index, cDLEVEL)
   
End Sub

Public Function checkRoad(ByVal Roomname As String) As Boolean
   checkRoad = False
   Dim start As Long
   Dim start2 As Long
   Roomname = LCase(Roomname)
   start = 0
   start2 = 0
   If start = 0 Then start = InStr(1, Roomname, "bridge", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, "crossing", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, "greenway", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, "old east road", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, "path ", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, " path", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, " paths", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, "trail", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, "trail ", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, " trail", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, " trails", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, "road", vbBinaryCompare)
   If start = 0 Then start = InStr(1, Roomname, " roads", vbBinaryCompare)
   If start > 0 Then
      If start2 = 0 Then start2 = InStrRev(Roomname, "edge of ", start, vbBinaryCompare)
      If start2 = 0 Then start2 = InStrRev(Roomname, "along ", start, vbBinaryCompare)
      If start2 = 0 Then start2 = InStrRev(Roomname, "beside ", start, vbBinaryCompare)
      If start2 = 0 Then start2 = InStrRev(Roomname, "besides ", start, vbBinaryCompare)
      If start2 = 0 Then start2 = InStrRev(Roomname, " by ", start, vbBinaryCompare)
      If start2 = 0 Then start2 = InStrRev(Roomname, "near ", start, vbBinaryCompare)
      If start2 = 0 Then start2 = InStrRev(Roomname, "off ", start, vbBinaryCompare)
      If start2 = 0 Then start2 = InStrRev(Roomname, " over ", start, vbBinaryCompare)
      If start2 = 0 Or (start < start2) Then
         checkRoad = True
      End If
   End If
End Function

Public Function dump(ByVal row As Integer, ByVal col As Integer, ByVal level As Integer)
   Dim cursor As Integer
   cursor = getInt(aWorld(row, row, level))
   'Debug.Print "   N (" & aData(cursor, cNPORTALR) & "," & aData(cursor, cNPORTALC) & ") - " & aData(cursor, cNLEVEL)
   'Debug.Print "   E (" & aData(cursor, cEPORTALR) & "," & aData(cursor, cEPORTALC) & ") - " & aData(cursor, cELEVEL)
   'Debug.Print "   S (" & aData(cursor, cSPORTALR) & "," & aData(cursor, cSPORTALC) & ") - " & aData(cursor, cSLEVEL)
   'Debug.Print "   W (" & aData(cursor, cWPORTALR) & "," & aData(cursor, cWPORTALC) & ") - " & aData(cursor, cWLEVEL)
   'Debug.Print " =========== "
End Function
