Attribute VB_Name = "mapping"
Public mapTerrain As Long
Public mapRide As Boolean
Public mapSun As Boolean
Public mapMonster As Boolean
Public mapDoorNameNorth As String
Public mapDoorNameEast As String
Public mapDoorNameSouth As String
Public mapDoorNameWest As String
Public mapDoorNameUp As String
Public mapDoorNameDown As String
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
Public mapRoomNorth As Boolean
Public mapRoomEast As Boolean
Public mapRoomSouth As Boolean
Public mapRoomWest As Boolean
Public mapRoomUp As Boolean
Public mapRoomDown As Boolean
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
Public mapSpecialNorth As Boolean
Public mapSpecialEast As Boolean
Public mapSpecialSouth As Boolean
Public mapSpecialWest As Boolean
Public mapSpecialUp As Boolean
Public mapSpecialDown As Boolean
Public mapcase As Long

Public Sub mapUpdate()
   If checkArrayLimit(theRow, theCol) = False Then Exit Sub
   mapValue = 0
   mapDesc = ""
   If mapSun = True Then mapValue = mapValue + 1
   If mapRide = True Then mapValue = mapValue + 2
   mapValue = mapValue + mapTerrain
   mapDesc = mapDesc & mapRoomName & ";"
   Call createData(mapRoomNorth, mapValue, mapDesc, N_noexit, N_exit, N_door, N_special, mapDoorNameNorth, mapRowNorth, mapColNorth)
   Call createData(mapRoomEast, mapValue, mapDesc, E_noexit, E_exit, E_door, E_special, mapDoorNameEast, mapRowEast, mapColEast)
   Call createData(mapRoomSouth, mapValue, mapDesc, S_noexit, S_exit, S_door, S_special, mapDoorNameSouth, mapRowSouth, mapColSouth)
   Call createData(mapRoomWest, mapValue, mapDesc, W_noexit, W_exit, W_door, W_special, mapDoorNameWest, mapRowWest, mapColWest)
   Call createData(mapRoomUp, mapValue, mapDesc, U_noexit, U_exit, U_door, U_special, mapDoorNameUp, mapRowUp, mapColUp)
   Call createData(mapRoomDown, mapValue, mapDesc, D_noexit, D_exit, D_door, D_special, mapDoorNameDown, mapRowDown, mapColDown)
   mapDesc = mapDesc & mapDescription & ";"
   If mapValue > 0 Then
      arr(theRow, theCol) = mapValue
      arrDesc(theRow, theCol) = mapDesc
      Call LoadRoom(theRow, theCol)
      Call DrawMap
      Exit Sub
   End If
End Sub

Public Sub setMapTerrain(ByVal what)
      Select Case what
      Case "road"
         mapTerrain = road
      Case "plain"
         mapTerrain = plain
      Case "field"
         mapTerrain = plain
      Case "forest"
         mapTerrain = forest
      Case "swamp"
         mapTerrain = swamp
      Case "hill"
         mapTerrain = hill
      Case "mountain"
         mapTerrain = mountain
      Case "water"
         mapTerrain = water
      Case "special"
         mapTerrain = special
      Case Else
         mapTerrain = road
      End Select
End Sub

Public Sub checkMapCommand(ByRef strData As String)
If MAP_MODE = False Then Exit Sub
With BestEST
   If checkString(strData, "MAP_UPDATE") = True Then
      Call mapUpdate
  '    Call UpdateReport
   End If
   If checkString(LCase(strData), "map_report") = True Then
    '  Call MapReport
      Exit Sub
   End If
   If checkString(strData, "t=") = True Then
      Call setMapTerrain(LCase(Mid(strData, 3, Len(strData) - 3)))
   '   Call MapReport
      Exit Sub
   End If
   If checkString(strData, "mapride") = True Then
      mapRide = Not (mapRide)
   '   Call MapReport
      Exit Sub
   End If
   If checkString(strData, "mapsun") = True Then
      mapSun = Not (mapSun)
     ' Call MapReport
      Exit Sub
   End If
   If checkString(strData, "dark") = True Then
      mapSun = 1
     ' Call MapReport
      Exit Sub
   End If
   If checkString(strData, "nexit") = True Then
      mapRoomNorth = Not (mapRoomNorth)
   '   Call MapReport
      Exit Sub
   End If
   If checkString(strData, "eexit") = True Then
      mapRoomEast = Not (mapRoomEast)
      'Call MapReport
      Exit Sub
   End If
   If checkString(strData, "sexit") = True Then
      mapRoomSouth = Not (mapRoomSouth)
   '   Call MapReport
      Exit Sub
   End If
   If checkString(strData, "wexit") = True Then
      mapRoomWest = Not (mapRoomWest)
   '   Call MapReport
      Exit Sub
   End If
   If checkString(strData, "uexit") = True Then
      mapRoomUp = Not (mapRoomUp)
     ' Call MapReport
      Exit Sub
   End If
   If checkString(strData, "dexit") = True Then
      mapRoomDown = Not (mapRoomDown)
   '   Call MapReport
      Exit Sub
   End If
   If checkString(strData, "ndoor=") = True Then
      mapDoorNameNorth = Mid(strData, 7, Len(strData) - 7)
   '   Call MapReport
      Exit Sub
   End If
   If checkString(strData, "edoor=") = True Then
      mapDoorNameEast = Mid(strData, 7, Len(strData) - 7)
    '  Call MapReport
      Exit Sub
   End If
   If checkString(strData, "sdoor=") = True Then
      mapDoorNameSouth = Mid(strData, 7, Len(strData) - 7)
   '   Call MapReport
      Exit Sub
   End If
   If checkString(strData, "wdoor=") = True Then
      mapDoorNameWest = Mid(strData, 7, Len(strData) - 7)
   '   Call MapReport
      Exit Sub
   End If
   If checkString(strData, "udoor=") = True Then
      mapDoorNameUp = Mid(strData, 7, Len(strData) - 7)
  '    Call MapReport
      Exit Sub
   End If
   If checkString(strData, "ddoor=") = True Then
      mapDoorNameDown = Mid(strData, 7, Len(strData) - 7)
     ' Call MapReport
      Exit Sub
   End If
   If checkString(strData, "nroom=") = True Then
      tempData = Split(Mid(strData, 7), ",")
      mapRowNorth = tempData(0)
      mapColNorth = tempData(1)
    '  Call MapReport
      Exit Sub
   End If
   If checkString(strData, "eroom=") = True Then
      tempData = Split(Mid(strData, 7), ",")
      mapRowEast = tempData(0)
      mapColEast = tempData(1)
   '   Call MapReport
      Exit Sub
   End If
   If checkString(strData, "sroom=") = True Then
      tempData = Split(Mid(strData, 7), ",")
      mapRowSouth = tempData(0)
      mapColSouth = tempData(1)
  '    Call MapReport
      Exit Sub
   End If
   If checkString(strData, "wroom=") = True Then
      tempData = Split(Mid(strData, 7), ",")
      mapRowWest = tempData(0)
      mapColWest = tempData(1)
      Exit Sub
   End If
   If checkString(strData, "uroom=") = True Then
      tempData = Split(Mid(strData, 7), ",")
      mapRowUp = tempData(0)
      mapColUp = tempData(1)
'     Call MapReport
      Exit Sub
   End If
   If checkString(strData, "droom=") = True Then
      tempData = Split(Mid(strData, 7), ",")
      mapRowDown = tempData(0)
      mapColDown = tempData(1)
 '    Call MapReport
      Exit Sub
   End If
   If checkString(strData, "goto=") = True Then
      Call gotoArea(strData)
      Exit Sub
   End If
   If checkString(strData, "MAP_MOVE_NORTH") = True Then
      theRow = theRow - 1
      Call LoadRoom(theRow, theCol)
      Call DrawMap
      Exit Sub
   End If
   If checkString(strData, "MAP_MOVE_EAST") = True Then
      theCol = theCol + 1
      Call LoadRoom(theRow, theCol)
      Call DrawMap
      Exit Sub
   End If
   If checkString(strData, "MAP_MOVE_WEST") = True Then
      theCol = theCol - 1
      Call LoadRoom(theRow, theCol)
      Call DrawMap
      Exit Sub
   End If
   If checkString(strData, "MAP_MOVE_SOUTH") = True Then
      theRow = theRow + 1
      Call LoadRoom(theRow, theCol)
      Call DrawMap
      Exit Sub
   End If
   If checkString(strData, "MAP_GET") = True Then
      Call mapGet
      Exit Sub
   End If
   theCommand = Split(strData, vbLf)
   For n = LBound(theCommand) To UBound(theCommand) - 1
      If Len(theCommand(n)) = 1 Then
         Select Case theCommand(n)
         Case "n"
            .tcpClient.SendData theCommand(n) & vbLf
            theRow = theRow - 1
            Call LoadRoom(theRow, theCol)
            Call DrawMap
         Case "e"
            .tcpClient.SendData theCommand(n) & vbLf
            theCol = theCol + 1
            Call LoadRoom(theRow, theCol)
            Call DrawMap
         Case "s"
            .tcpClient.SendData theCommand(n) & vbLf
            theRow = theRow + 1
            Call LoadRoom(theRow, theCol)
            Call DrawMap
         Case "w"
            .tcpClient.SendData theCommand(n) & vbLf
            theCol = theCol - 1
            Call LoadRoom(theRow, theCol)
            Call DrawMap
         Case Else
            .tcpClient.SendData strData
         End Select
      Else
         .tcpClient.SendData strData
      End If
   Next
End With
End Sub
Public Sub UpdateReport()
On Error GoTo ErrorHandler
With BestEST
   .tcpServer.SendData vbLf & "_____MAP UPDATE SUCCESSFUL___row=" & theRow & "__col=" & theCol & "__________________"
   .tcpServer.SendData vbLf & "Roomname  " & theRoomName
   .tcpServer.SendData vbLf & "Desc(50)  " & mapDescription
   .tcpServer.SendData vbLf & "Terrain   " & mapTerrain
   .tcpServer.SendData vbLf & "Ridable   " & mapRide
   .tcpServer.SendData vbLf & "Sun       " & mapSun
   .tcpServer.SendData vbLf & "_____________EXIT___DOOR____SPEC___ROW__COL______DOORNAME_________________"
   .tcpServer.SendData vbLf & "North       " & theRoomNorth & "   " & theDoorNorth & "   " & theSpecialNorth & "   " & theRowNorth & "   " & theColNorth & "          " & theDoorNameNorth
   .tcpServer.SendData vbLf & "East        " & theRoomEast & "   " & theDoorEast & "   " & theSpecialEast & "   " & theRowEast & "   " & theColEast & "          " & theDoorNameEast
   .tcpServer.SendData vbLf & "South       " & theRoomSouth & "   " & theDoorSouth & "   " & theSpecialSouth & "   " & theRowSouth & "   " & theColSouth & "          " & theDoorNameSouth
   .tcpServer.SendData vbLf & "West        " & theRoomWest & "   " & theDoorWest & "   " & theSpecialWest & "   " & theRowWest & "   " & theColWest & "          " & theDoorNameWest
   .tcpServer.SendData vbLf & "Up          " & theRoomUp & "   " & theDoorUp & "   " & theSpecialUp & "   " & theRowUp & "   " & theColUp & "          " & theDoorNameUp
   .tcpServer.SendData vbLf & "Down        " & theRoomDown & "   " & theDoorDown & "   " & theSpecialDown & "   " & theRowDown & "   " & theColDown & "          " & theDoorNameDown
End With
Exit Sub
ErrorHandler:
   BestEST.status = "Proxy server down..."
End Sub

Public Sub MapReport()
On Error GoTo ErrorHandler
With BestEST
   .tcpServer.SendData vbLf & "_____MAP REPORT _______row=" & theRow & "__col=" & theCol & "__________________"
   .tcpServer.SendData vbLf & "Roomname  " & mapRoomName
   .tcpServer.SendData vbLf & "Desc(50)  " & mapDescription
   .tcpServer.SendData vbLf & "Terrain   " & mapTerrain
   .tcpServer.SendData vbLf & "Ridable   " & mapRide
   .tcpServer.SendData vbLf & "Sun       " & mapSun
   .tcpServer.SendData vbLf & "_____________EXIT___DOOR____SPEC___ROW__COL______DOORNAME_________________"
   .tcpServer.SendData vbLf & "North       " & mapRoomNorth & "   " & mapDoorNorth & "   " & mapSpecialNorth & "   " & mapRowNorth & "   " & mapColNorth & "          " & mapDoorNameNorth
   .tcpServer.SendData vbLf & "East        " & mapRoomEast & "   " & mapDoorEast & "   " & mapSpecialEast & "   " & mapRowEast & "   " & mapColEast & "          " & mapDoorNameEast
   .tcpServer.SendData vbLf & "South       " & mapRoomSouth & "   " & mapDoorSouth & "   " & mapSpecialSouth & "   " & mapRowSouth & "   " & mapColSouth & "          " & mapDoorNameSouth
   .tcpServer.SendData vbLf & "West        " & mapRoomWest & "   " & mapDoorWest & "   " & mapSpecialWest & "   " & mapRowWest & "   " & mapColWest & "          " & mapDoorNameWest
   .tcpServer.SendData vbLf & "Up          " & mapRoomUp & "   " & mapDoorUp & "   " & mapSpecialUp & "   " & mapRowUp & "   " & mapColUp & "          " & mapDoorNameUp
   .tcpServer.SendData vbLf & "Down        " & mapRoomDown & "   " & mapDoorDown & "   " & mapSpecialDown & "   " & mapRowDown & "   " & mapColDown & "          " & mapDoorNameDown
End With
Exit Sub
ErrorHandler:
   BestEST.status = "Proxy server down..."
End Sub

Public Sub zeroMap()
   Set BestEST.pictureTerrain.Picture = pRoad
   mapRoomName = ""
   mapDescription = ""
   mapRide = True
   mapSun = True
   mapTerrain = 0
   mapDoorNameNorth = ""
   mapDoorNameEast = ""
   mapDoorNameSouth = ""
   mapDoorNameWest = ""
   mapDoorNameUp = ""
   mapDoorNameDown = ""
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
   mapSpecialNorth = False
   mapSpecialEast = False
   mapSpecialSouth = False
   mapSpecialWest = False
   mapSpecialUp = False
   mapSpecialDown = False
   mapRoomNorth = False
   mapRoomEast = False
   mapRoomSouth = False
   mapRoomWest = False
   mapRoomUp = False
   mapRoomDown = False
   mapDoorNorth = False
   mapDoorEast = False
   mapDoorSouth = False
   mapDoorWest = False
   mapDoorUp = False
   mapDoorDown = False
   With BestEST
      .GoToNumber.Text = ""
      .Roomname.Text = ""
      .Description.Text = ""
      .Sun.Value = 1
      .Ridable.Value = 1
      .Monster.Value = 0
      .nExit.Value = 0
      .nDoor.Text = ""
      .nNumber.Text = ""
      .eExit.Value = 0
      .eDoor.Text = ""
      .eNumber.Text = ""
      .sExit.Value = 0
      .sDoor.Text = ""
      .sNumber.Text = ""
      .wExit.Value = 0
      .wDoor.Text = ""
      .wNumber.Text = ""
      .uExit.Value = 0
      .uDoor.Text = ""
      .uNumber.Text = ""
      .dExit.Value = 0
      .dDoor.Text = ""
      .dNumber.Text = ""
   End With
End Sub

Public Sub setMapModeON()
   Call zeroMap
   MAP_MODE = True
   MAP_THE_DATA = False
   Out_Of_Sync = True
   roomCount = 0
End Sub

Public Sub setMapModeOFF()
   MAP_MODE = False
   MAP_THE_DATA = False
   Out_Of_Sync = True
   roomCount = 0
End Sub

Public Sub gotoArea(ByRef data)
'On Error GoTo ErrorHandler
   If Len(data) < 3 Then Err.Raise 555
      currentRoom = Split(data, ",")
      If UBound(currentRoom) > 0 Then
         theRow = currentRoom(0)
         theCol = currentRoom(1)
         Call LoadRoom(theRow, theCol)
         Call DrawMap
      End If
Exit Sub
ErrorHandler:
   Call InvalidData
End Sub

Public Sub mapGet()
On Error GoTo ErrorHandler
   Call zeroMap
   If checkArrayLimit(theRow, theCol) = True Then
      MAP_THE_DATA = True
      MAP_THE_CASE = 1
      BestEST.tcpClient.SendData "examine" & vbLf
   End If
Exit Sub
ErrorHandler:
   BestEST.status.Caption = "Proxy server down..."
End Sub

Public Sub GetMapData()
   With BestEST
      .Roomname.Text = ""
      .Description.Text = ""
      .Sun.Value = 0
      .Ridable.Value = 0
      .Monster.Value = 0
      .nExit.Value = 0
      .nDoor.Text = ""
      .nNumber.Text = ""
      .eExit.Value = 0
      .eDoor.Text = ""
      .eNumber.Text = ""
      .sExit.Value = 0
      .sDoor.Text = ""
      .sNumber.Text = ""
      .wExit.Value = 0
      .wDoor.Text = ""
      .wNumber.Text = ""
      .uExit.Value = 0
      .uDoor.Text = ""
      .uNumber.Text = ""
      .dExit.Value = 0
      .dDoor.Text = ""
      .dNumber.Text = ""
      
      mapRoomName = theRoomName
      .Roomname.Text = mapRoomName
      mapDescription = theRoomDesc
      .Description.Text = mapDescription
      mapSun = theSun
      If mapSun = True Then
         .Sun.Value = 1
      Else
         .Sun.Value = 0
      End If
      mapRide = theRide
      If mapRide = True Then
         .Ridable.Value = 1
      Else
         .Ridable.Value = 0
      End If
      mapMonster = theMonster
      If mapMonster = True Then
         .Monster.Value = 1
      Else
         .Monster.Value = 0
      End If
      If theRoomNorth = True Then
         mapRoomNorth = True
         .nExit.Value = 1
      Else
         mapRoomNorth = False
         .nExit.Value = 0
      End If
      If theRoomEast = True Then
         mapRoomEast = True
         .eExit.Value = 1
      Else
         mapRoomEast = False
         .eExit.Value = 0
      End If
      If theRoomSouth = True Then
         mapRoomSouth = True
         .sExit.Value = 1
      Else
         mapRoomSouth = False
         .sExit.Value = 0
      End If
      If theRoomWest = True Then
         mapRoomWest = True
         .wExit.Value = 1
      Else
         mapRoomWest = False
         .wExit.Value = 0
      End If
      If theRoomUp = True Then
         mapRoomUp = True
         .uExit.Value = 1
      Else
         mapRoomUp = False
         .uExit.Value = 0
      End If
      If theRoomDown = True Then
         mapRoomDown = True
         .dExit.Value = 1
      Else
         mapRoomDown = False
         .dExit.Value = 0
      End If
      If theDoorNorth = True Then
         mapDoorNorth = True
         .nDoor.Text = theDoorNameNorth
      Else
         mapDoorNorth = False
         .nDoor.Text = ""
      End If
      If theDoorEast = True Then
         mapDoorEast = True
         .eDoor.Text = theDoorNameEast
      Else
         mapDoorEast = False
         .eDoor.Text = ""
      End If
      If theDoorSouth = True Then
         mapDoorSouth = True
         .sDoor.Text = theDoorNameSouth
      Else
         mapDoorSouth = False
         .sDoor.Text = ""
      End If
      If theDoorWest = True Then
         mapDoorWest = True
         .wDoor.Text = theDoorNameWest
      Else
         mapDoorWest = False
         .wDoor.Text = ""
      End If
      If theDoorUp = True Then
         mapDoorUp = True
         .uDoor.Text = theDoorNameUp
      Else
         mapDoorUp = False
         .uDoor.Text = ""
      End If
      If theDoorDown = True Then
         mapDoorDown = True
         .dDoor.Text = theDoorNameDown
      Else
         mapDoorDown = False
         .dDoor.Text = ""
      End If
      If theSpecialNorth = True Then
         mapRowNorth = theRowNorth
         mapColNorth = theColNorth
         .nNumber.Text = theRowNorth & "," & theColNorth
      Else
         mapRowNorth = 0
         mapColNorth = 0
         .nNumber.Text = ""
      End If
      If theSpecialEast = True Then
         mapRowEast = theRowEast
         mapColEast = theColEast
         .nNumber.Text = theRowEast & "," & theColEast
      Else
         mapRowEast = 0
         mapColEast = 0
         .nNumber.Text = ""
      End If
      If theSpecialSouth = True Then
         mapRowSouth = theRowSouth
         mapColSouth = theColSouth
         .nNumber.Text = theRowSouth & "," & theColSouth
      Else
         mapRowSouth = 0
         mapColSouth = 0
         .nNumber.Text = ""
      End If
      If theSpecialWest = True Then
         mapRowWest = theRowWest
         mapColWest = theColWest
         .nNumber.Text = theRowWest & "," & theColWest
      Else
         mapRowWest = 0
         mapColWest = 0
         .nNumber.Text = ""
      End If
      If theSpecialUp = True Then
         mapRowUp = theRowUp
         mapColUp = theColUp
         .nNumber.Text = theRowUp & "," & theColUp
      Else
         mapRowUp = 0
         mapColUp = 0
         .nNumber.Text = ""
      End If
      If theSpecialDown = True Then
         mapRowDown = theRowDown
         mapColDown = theColDown
         .nNumber.Text = theRowDown & "," & theColDown
      Else
         mapRowDown = 0
         mapColDown = 0
         .nNumber.Text = ""
      End If
   End With
End Sub

