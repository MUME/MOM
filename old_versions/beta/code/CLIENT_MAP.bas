Attribute VB_Name = "CLIENT_MAPPING"
Public Function handleMappingCommand(ByRef strData As String)
On Error GoTo errorhandler
Dim tempData
Dim theCommand
Dim a As Long
Dim b As String
Dim n As Long
   handleMappingCommand = False
   a = Len(strData)
   b = Mid(strData, 1, 1)
   
   If MappingMode = True And b <> "_" Then
      With frmMap
         theCommand = Split(strData, vbLf)
         For n = LBound(theCommand) To UBound(theCommand) - 1
            If Len(theCommand(n)) = 1 Then
               Select Case theCommand(n)
               Case "n"
                  .tcpClient.SendData theCommand(n) & vbLf
                  theRow = theRow - 1
                  Call loadRoom(theRow, theCol)
                  Call DrawMap
               Case "e"
                  .tcpClient.SendData theCommand(n) & vbLf
                  theCol = theCol + 1
                  Call loadRoom(theRow, theCol)
                  Call DrawMap
               Case "s"
                  .tcpClient.SendData theCommand(n) & vbLf
                  theRow = theRow + 1
                  Call loadRoom(theRow, theCol)
                  Call DrawMap
               Case "w"
                  .tcpClient.SendData theCommand(n) & vbLf
                  theCol = theCol - 1
                  Call loadRoom(theRow, theCol)
                  Call DrawMap
               Case Else
                  .tcpClient.SendData strData
               End Select
            Else
               .tcpClient.SendData strData
            End If
         Next
      End With
      handleMappingCommand = True: Exit Function
   End If

   If a > 8 And b = "_" Then
      Select Case strData
      Case "_movenorth"
         Out_Of_Sync = True
         theRow = theRow - 1
         Call loadRoom(theRow, theCol)
         Call DrawMap
         handleMappingCommand = True: Exit Function
      Case "_moveeast"
         Out_Of_Sync = True
         theCol = theCol + 1
         Call loadRoom(theRow, theCol)
         Call DrawMap
         handleMappingCommand = True: Exit Function
      Case "_movewest"
         Out_Of_Sync = True
         theCol = theCol - 1
         Call loadRoom(theRow, theCol)
         Call DrawMap
         handleMappingCommand = True: Exit Function
      Case "_movesouth"
         Out_Of_Sync = True
         theRow = theRow + 1
         Call loadRoom(theRow, theCol)
         Call DrawMap
         handleMappingCommand = True: Exit Function
      End Select
   End If

   If MappingMode = True And a > 1 And b = "_" Then
      Select Case strData
      Case "_get"
         Call getRoomData
         handleMappingCommand = True: Exit Function
      Case "_update"
         Call mapUpdate
         handleMappingCommand = True: Exit Function
      Case "_map" 'get + update
         MappingGetUpdate = True
         Call getRoomData
         handleMappingCommand = True: Exit Function
      Case "_ride"
         mapRide = Not (mapRide)
         handleMappingCommand = True: Exit Function
      Case "_sun"
         mapSun = Not (mapSun)
         handleMappingCommand = True: Exit Function
      Case "_n"
         mapExitNorth = Not (mapExitNorth)
         handleMappingCommand = True: Exit Function
      Case "_e"
         mapExitEast = Not (mapExitEast)
         handleMappingCommand = True: Exit Function
      Case "_s"
         mapExitSouth = Not (mapExitSouth)
         handleMappingCommand = True: Exit Function
      Case "_w"
         mapExitWest = Not (mapExitWest)
         handleMappingCommand = True: Exit Function
      Case "_u"
         mapExitUp = Not (mapExitUp)
         handleMappingCommand = True: Exit Function
      Case "_d"
         mapExitDown = Not (mapExitDown)
         handleMappingCommand = True: Exit Function
      End Select
      
      If checkString(strData, "_t ") = True Then
         Call setMapTerrain(Mid(strData, 4, Len(strData) - 4))
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_nd ") = True Then
         mapDoornameNorth = Mid(strData, 5, Len(strData) - 5)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_ed ") = True Then
         mapDoornameEast = Mid(strData, 5, Len(strData) - 5)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_sd ") = True Then
         mapDoornameSouth = Mid(strData, 5, Len(strData) - 5)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_wd ") = True Then
         mapDoornameWest = Mid(strData, 5, Len(strData) - 5)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_ud ") = True Then
         mapDoornameUp = Mid(strData, 5, Len(strData) - 5)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_dd ") = True Then
         mapDoornameDown = Mid(strData, 5, Len(strData) - 5)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_np ") = True Then
         tempData = Split(Mid(strData, 5), ",")
         mapRowNorth = tempData(0)
         mapColNorth = tempData(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_ep ") = True Then
         tempData = Split(Mid(strData, 5), ",")
         mapRowEast = tempData(0)
         mapColEast = tempData(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_sp ") = True Then
         tempData = Split(Mid(strData, 5), ",")
         mapRowSouth = tempData(0)
         mapColSouth = tempData(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_wp ") = True Then
         tempData = Split(Mid(strData, 5), ",")
         mapRowWest = tempData(0)
         mapColWest = tempData(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_up ") = True Then
         tempData = Split(Mid(strData, 5), ",")
         mapRowUp = tempData(0)
         mapColUp = tempData(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_dp ") = True Then
         tempData = Split(Mid(strData, 5), ",")
         mapRowDown = tempData(0)
         mapColDown = tempData(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkString(strData, "_go ") = True Then
         Call gotoArea(Mid(strData, 5))
         handleMappingCommand = True: Exit Function
      End If
   End If

Exit Function
errorhandler:
   errorData = "Client_Mapping"
   writeError (errorData)
End Function
