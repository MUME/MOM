Attribute VB_Name = "CLIENT_MAPPING"
Option Explicit
Private a As Integer
Private b As String
Private n As Integer
Public mappingFromRow As Integer
Public mappingFromCol As Integer
Public mappingFromDir As String

Public Function handleMappingCommand(strData As String) As Boolean
errorData = errorData & "handleMappingCommand -> "
If DEBUGMODE = False Then On Error GoTo errorhandler
   handleMappingCommand = False
   If MappingMode = False Then Exit Function
   Dim rr As Integer
   Dim cc As Integer
   b = MidB$(strData, 1, 2)
   If b <> "_" Then
        If GODMODE Then
            Dim mys As String
            mys = LCase(strData)
            If InStrB(1, mys, "open ", vbBinaryCompare) > 0 _
            Or InStrB(1, mys, "close ", vbBinaryCompare) > 0 _
            Or InStrB(1, mys, "bash ", vbBinaryCompare) > 0 _
            Or InStrB(1, mys, "lock ", vbBinaryCompare) > 0 _
            Or InStrB(1, mys, "unlock ", vbBinaryCompare) > 0 _
            Or InStrB(1, mys, "pick ", vbBinaryCompare) > 0 Then
                If InStrB(1, mys, "exit n", vbBinaryCompare) > 0 Then
                    If LenB(aData(getIndex(virtualRow, virtualCol), cNDOOR)) > 0 Then mys = Replace(mys, "exit n", aData(getIndex(virtualRow, virtualCol), cNDOOR) & " n", , 1, vbBinaryCompare)
                End If
                If InStrB(1, mys, "exit e", vbBinaryCompare) > 0 Then
                    If LenB(aData(getIndex(virtualRow, virtualCol), cEDOOR)) > 0 Then mys = Replace(mys, "exit e", aData(getIndex(virtualRow, virtualCol), cEDOOR) & " e", , 1, vbBinaryCompare)
                End If
                If InStrB(1, mys, "exit s", vbBinaryCompare) > 0 Then
                    If LenB(aData(getIndex(virtualRow, virtualCol), cSDOOR)) > 0 Then mys = Replace(mys, "exit s", aData(getIndex(virtualRow, virtualCol), cSDOOR) & " s", , 1, vbBinaryCompare)
                End If
                If InStrB(1, mys, "exit w", vbBinaryCompare) > 0 Then
                    If LenB(aData(getIndex(virtualRow, virtualCol), cWDOOR)) > 0 Then mys = Replace(mys, "exit w", aData(getIndex(virtualRow, virtualCol), cWDOOR) & " w", , 1, vbBinaryCompare)
                End If
                If InStrB(1, mys, "exit u", vbBinaryCompare) > 0 Then
                    If LenB(aData(getIndex(virtualRow, virtualCol), cUDOOR)) > 0 Then mys = Replace(mys, "exit u", aData(getIndex(virtualRow, virtualCol), cUDOOR) & " u", , 1, vbBinaryCompare)
                End If
                If InStrB(1, mys, "exit d", vbBinaryCompare) > 0 Then
                    If LenB(aData(getIndex(virtualRow, virtualCol), cDDOOR)) > 0 Then mys = Replace(mys, "exit d", aData(getIndex(virtualRow, virtualCol), cDDOOR) & " d", , 1, vbBinaryCompare)
                End If
                If InStrB(1, mys, "@", vbBinaryCompare) > 0 Then mys = Replace(mys, "@", "", , 1, vbBinaryCompare)
                frmMap.tcpPlayer.SendData mys ' & vbCrLf
                'Call informClient(lookHeader & WHITE & ";" & BOLD & lookFooter & mys & colourEndCode & vbCrLf, True)
                handleMappingCommand = True: Exit Function
            End If
        End If
        
      Dim theCommand
      If specialLen = 2 Then theCommand = Split(strData, vbLf, , vbBinaryCompare)
      If specialLen = 4 Then theCommand = Split(strData, vbCrLf, , vbBinaryCompare)
      For n = LBound(theCommand) To UBound(theCommand) - 1
         If LenB(theCommand(n)) = 2 Then
            rr = theROW
            cc = theCOL
            Select Case theCommand(n)
            Case "u"
               frmMap.tcpPlayer.SendData theCommand(n) & vbCrLf
               theROW = IIF(LenB(aData(getIndex(rr, cc), cUPORTALR)) > 2, aData(getIndex(rr, cc), cUPORTALR), theROW)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cUPORTALC)) > 2, aData(getIndex(rr, cc), cUPORTALC), theCOL)
               Call loadRoom(theROW, theCOL)
               Call DrawMap
            Case "d"
               frmMap.tcpPlayer.SendData theCommand(n) & vbCrLf
               theROW = IIF(LenB(aData(getIndex(rr, cc), cDPORTALR)) > 2, aData(getIndex(rr, cc), cDPORTALR), theROW)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cDPORTALC)) > 2, aData(getIndex(rr, cc), cDPORTALC), theCOL)
               Call loadRoom(theROW, theCOL)
               Call DrawMap
            Case "n"
               frmMap.tcpPlayer.SendData theCommand(n) & vbCrLf
               'theROW = theROW - 1
               theROW = IIF(LenB(aData(getIndex(rr, cc), cNPORTALR)) > 2, aData(getIndex(rr, cc), cNPORTALR), theROW - 1)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cNPORTALC)) > 2, aData(getIndex(rr, cc), cNPORTALC), theCOL)
               Call loadRoom(theROW, theCOL)
               Call DrawMap
            Case "e"
               frmMap.tcpPlayer.SendData theCommand(n) & vbCrLf
               'theCOL = theCOL + 1
               theROW = IIF(LenB(aData(getIndex(rr, cc), cEPORTALR)) > 2, aData(getIndex(rr, cc), cEPORTALR), theROW)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cEPORTALC)) > 2, aData(getIndex(rr, cc), cEPORTALC), theCOL + 1)
               Call loadRoom(theROW, theCOL)
               Call DrawMap
            Case "s"
               frmMap.tcpPlayer.SendData theCommand(n) & vbCrLf
               'theROW = theROW + 1
               theROW = IIF(LenB(aData(getIndex(rr, cc), cSPORTALR)) > 2, aData(getIndex(rr, cc), cSPORTALR), theROW + 1)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cSPORTALC)) > 2, aData(getIndex(rr, cc), cSPORTALC), theCOL)
               Call loadRoom(theROW, theCOL)
               Call DrawMap
            Case "w"
               frmMap.tcpPlayer.SendData theCommand(n) & vbCrLf
               'theCOL = theCOL - 1
               theROW = IIF(LenB(aData(getIndex(rr, cc), cWPORTALR)) > 2, aData(getIndex(rr, cc), cWPORTALR), theROW)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cWPORTALC)) > 2, aData(getIndex(rr, cc), cWPORTALC), theCOL - 1)
               Call loadRoom(theROW, theCOL)
               Call DrawMap
            Case Else
               frmMap.tcpPlayer.SendData theCommand(n) & vbCrLf
            End Select
         Else
            frmMap.tcpPlayer.SendData theCommand(n) & vbCrLf
         End If
      Next
      handleMappingCommand = True: Exit Function
   Else
      Dim myCase As String
      myCase = MidB$(strData, 1, LenB(strData) - specialLen)
     
      Select Case myCase
      Case "_movemapnorth", "_movemapeast", "_movemapsouth", "_movemapwest"
         Select Case myCase
            Case "_movemapnorth"
               theROW = theROW - 1
            Case "_movemapeast"
               theCOL = theCOL + 1
            Case "_movemapsouth"
               theROW = theROW + 1
            Case "_movemapwest"
               theCOL = theCOL - 1
         End Select
         If isValid(theROW, theCOL) = True Then
            Call loadRoom(theROW, theCOL)
            Call DrawMap
         Else
            Call informClient("Out of map area!")
         End If
         handleMappingCommand = True: Exit Function
      Case "_mapup", "_mapdown", "_mapnorth", "_mapeast", "_mapsouth", "_mapwest"
      '--------------------------------------------------------------------
         rr = theROW
         cc = theCOL
         mappingFromRow = theROW
         mappingFromCol = theCOL
         Dim dir As String
         Call zeroMap
         Select Case myCase
         Case "_mapup"
            dir = "u": mappingFromDir = "d"
            If frmMap.mnuWalkthrough.Checked Then
               theROW = IIF(LenB(aData(getIndex(rr, cc), cUPORTALR)) > 2, aData(getIndex(rr, cc), cUPORTALR), theROW)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cUPORTALC)) > 2, aData(getIndex(rr, cc), cUPORTALC), theCOL)
            End If
         Case "_mapdown"
            dir = "d": mappingFromDir = "u"
            If frmMap.mnuWalkthrough.Checked Then
               theROW = IIF(LenB(aData(getIndex(rr, cc), cDPORTALR)) > 2, aData(getIndex(rr, cc), cDPORTALR), theROW)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cDPORTALC)) > 2, aData(getIndex(rr, cc), cDPORTALC), theCOL)
            End If
         Case "_mapnorth"
            dir = "n": mappingFromDir = "s"
            If frmMap.mnuWalkthrough.Checked Then
               theROW = IIF(LenB(aData(getIndex(rr, cc), cNPORTALR)) > 2, aData(getIndex(rr, cc), cNPORTALR), theROW - 1)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cNPORTALC)) > 2, aData(getIndex(rr, cc), cNPORTALC), theCOL)
            Else
               theROW = theROW - 1
            End If
         Case "_mapeast"
            dir = "e": mappingFromDir = "w"
            If frmMap.mnuWalkthrough.Checked Then
               theROW = IIF(LenB(aData(getIndex(rr, cc), cEPORTALR)) > 2, aData(getIndex(rr, cc), cEPORTALR), theROW)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cEPORTALC)) > 2, aData(getIndex(rr, cc), cEPORTALC), theCOL + 1)
            Else
               theCOL = theCOL + 1
            End If
         Case "_mapsouth"
            dir = "s": mappingFromDir = "n"
            If frmMap.mnuWalkthrough.Checked Then
               theROW = IIF(LenB(aData(getIndex(rr, cc), cSPORTALR)) > 2, aData(getIndex(rr, cc), cSPORTALR), theROW + 1)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cSPORTALC)) > 2, aData(getIndex(rr, cc), cSPORTALC), theCOL)
            Else
               theROW = theROW + 1
            End If
         Case "_mapwest"
            dir = "w": mappingFromDir = "e"
            If frmMap.mnuWalkthrough.Checked Then
               theROW = IIF(LenB(aData(getIndex(rr, cc), cWPORTALR)) > 2, aData(getIndex(rr, cc), cWPORTALR), theROW)
               theCOL = IIF(LenB(aData(getIndex(rr, cc), cWPORTALC)) > 2, aData(getIndex(rr, cc), cWPORTALC), theCOL - 1)
            Else
               theCOL = theCOL - 1
            End If
         End Select
         If isValid(theROW, theCOL) = True Then
            If LenB(dir) > 0 Then
               MappingGetUpdate = True
               dataFromMUD = True
               MappingData = True
               MappingCase = 1
               retry = 0
               frmMap.tcpPlayer.SendData dir & vbCrLf
            End If
         Else
            Call informClient("Out of map area!")
         End If
         handleMappingCommand = True: Exit Function
      '--------------------------------------------------------------------
      Case "_get"
         Call zeroMap
         Call getRoomData
         handleMappingCommand = True: Exit Function
      Case "_update"
         Call mapUpdate
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
      
      If checkStringCI(strData, "_t ") = True Then
         Call setMapTerrain(MidB$(strData, 7, LenB(strData) - (6 + specialLen)))
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_nd ") = True Then
         mapDoornameNorth = MidB$(strData, 9, LenB(strData) - (8 + specialLen))
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_ed ") = True Then
         mapDoornameEast = MidB$(strData, 9, LenB(strData) - (8 + specialLen))
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_sd ") = True Then
         mapDoornameSouth = MidB$(strData, 9, LenB(strData) - (8 + specialLen))
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_wd ") = True Then
         mapDoornameWest = MidB$(strData, 9, LenB(strData) - (8 + specialLen))
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_ud ") = True Then
         mapDoornameUp = MidB$(strData, 9, LenB(strData) - (8 + specialLen))
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_dd ") = True Then
         mapDoornameDown = MidB$(strData, 9, LenB(strData) - (8 + specialLen))
         handleMappingCommand = True: Exit Function
      End If
'manually mapping portals
      Dim tempdata
      If checkStringCI(strData, "_np ") = True Then
         tempdata = Split(MidB$(strData, 9), ",", , vbBinaryCompare)
         mapRowNorth = tempdata(0)
         mapColNorth = tempdata(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_ep ") = True Then
         tempdata = Split(MidB$(strData, 9), ",", , vbBinaryCompare)
         mapRowEast = tempdata(0)
         mapColEast = tempdata(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_sp ") = True Then
         tempdata = Split(MidB$(strData, 9), ",", , vbBinaryCompare)
         mapRowSouth = tempdata(0)
         mapColSouth = tempdata(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_wp ") = True Then
         tempdata = Split(MidB$(strData, 9), ",", , vbBinaryCompare)
         mapRowWest = tempdata(0)
         mapColWest = tempdata(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_up ") = True Then
         tempdata = Split(MidB$(strData, 9), ",", , vbBinaryCompare)
         mapRowUp = tempdata(0)
         mapColUp = tempdata(1)
         handleMappingCommand = True: Exit Function
      End If
      If checkStringCI(strData, "_dp ") = True Then
         tempdata = Split(MidB$(strData, 9), ",", , vbBinaryCompare)
         mapRowDown = tempdata(0)
         mapColDown = tempdata(1)
         handleMappingCommand = True: Exit Function
      End If
'      If checkStringCI(strData, "_go ") = True Then
'         Call gotoArea(MidB$(strData, 9))
'         handleMappingCommand = True: Exit Function
'      End If
   End If
Exit Function
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "Client_Mapping"
   writeError (errorModule)
End Function
