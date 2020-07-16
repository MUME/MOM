Attribute VB_Name = "load"
Option Explicit
Public registryPath As String
Public filePath As String
Public thekeymaker As Variant
Public theData As Long
Public WorldLoaded As Boolean
Public Initialized As Boolean
Public Roomsync As Boolean
Public Autosync As Boolean
Public SyncError As Boolean
Public theCommand
Public freedom As Boolean
Public arrWorld(1 To 300, 1 To 600) As Integer
Public arrData(0 To 20000, 0 To 24) As String
Public theCount As Integer
Public Const cENCRYPTED = 0
Public Const cROW = 1
Public Const cCOL = 2
Public Const cDATA = 3
Public Const cROOMNAME = 4
Public Const cDESCRIPTION = 5
Public Const cNDOOR = 6
Public Const cEDOOR = 7
Public Const cSDOOR = 8
Public Const cWDOOR = 9
Public Const cUDOOR = 10
Public Const cDDOOR = 11
Public Const cNPORTALR = 12
Public Const cNPORTALC = 13
Public Const cEPORTALR = 14
Public Const cEPORTALC = 15
Public Const cSPORTALR = 16
Public Const cSPORTALC = 17
Public Const cWPORTALR = 18
Public Const cWPORTALC = 19
Public Const cUPORTALR = 20
Public Const cUPORTALC = 21
Public Const cDPORTALR = 22
Public Const cDPORTALC = 23
Public Const cNOTE = 24

Public Sub loadWorld()
If DEBUGMODE = False Then On Error GoTo errorhandler
   errorData = "loadWorld -> "
   WorldLoaded = False
   
   Dim key As Variant
   key = getPassword()
   thekeymaker = key
   Dim v As Variant
   Dim encrypted As Variant
   Dim original As Variant
   theCount = 0
   
   Dim w As Integer
   For w = cENCRYPTED To cNOTE
      arrData(0, w) = 0
   Next
   
   Open filePath For Input As #1
   Dim failure As Boolean
   failure = False
   Do While Not EOF(1)
      Line Input #1, encrypted
'----- map convert sequence --------------------------------------------------
'      If True Then
'         original = cast128.cast128decode(key, encrypted)      'oldkey
'         encrypted = cast128.cast128encode("979048413", original)     'newkey
'      Else
         original = cast128.cast128decode(key, encrypted)
'      End If
'----- end -------------------------------------------------------------------
      v = Split(original, ";", , vbBinaryCompare)
      If theCount >= arrMaxData Then
         failure = True
         Exit Do
      Else
         theCount = theCount + 1
      End If
      arrData(theCount, cENCRYPTED) = encrypted
      For w = cROW - 1 To cNOTE - 1
         arrData(theCount, w + 1) = v(w)
      Next
      arrWorld(arrData(theCount, cROW), arrData(theCount, cCOL)) = theCount
   Loop
   For w = theCount + 1 To UBound(arrData)
      arrData(w, cDATA) = 0
   Next
   Close #1
   
   If failure Then MsgBox ("The world database is full!")
   WorldLoaded = True

Exit Sub
errorhandler:
   WorldLoaded = False
   MsgBox "Invalid database or corrupted installation!" & vbCrLf & vbCrLf & Err.description & "(" & Err.Number & ")"
   errorModule = Err.description & "(" & Err.Number & ") -> " & "load loadWorld"
   writeError (errorModule)
End Sub

Public Sub loadRoom(row As Long, col As Long)
errorData = errorData & "loadRoom -> "
If DEBUGMODE = False Then On Error GoTo errorhandler
Dim kala As Boolean
   If checkArrayLimit(row, col) = True Then
      theData = arrData(arrWorld(row, col), cDATA)
      theRoomStringOk = False
      theRoomname = vbNullString
      theRoomdesc = vbNullString
      SyncError = True
      If LOST = False And followMode = False And noexitsfound = False Then Call setNewExits(currentExits)
      
'      kala = False
'      If Roomsync And LOST = False Then
'         If (newNorth And (theData And N_MAP) = 0) Then kala = True
'         If (newEast And (theData And E_MAP) = 0) Then kala = True
'         If (newSouth And (theData And S_MAP) = 0) Then kala = True
'         If (newWest And (theData And W_MAP) = 0) Then kala = True
'         If (newUp And (theData And U_MAP) = 0) Then kala = True
'         If (newDown And (theData And D_MAP) = 0) Then kala = True
'         If kala Then Exit Sub
'         theRoomname = arrData(arrWorld(row, col), cROOMNAME)
'         theRoomdesc = arrData(arrWorld(row, col), cDESCRIPTION)
'         theRoomStringOk = True
'         If FollowMode = False Then
'            If theRoomname <> currentRoomName Then
'               Exit Sub
'            End If
'         End If
'      End If

      SyncError = False
      theTerrain = 0
      theFlag = 0
      theRide = False
      theSun = False
      theMonster = False
      theDoornameNorth = vbNullString
      theDoornameEast = vbNullString
      theDoornameSouth = vbNullString
      theDoornameWest = vbNullString
      theDoornameUp = vbNullString
      theDoornameDown = vbNullString
      theRowNorth = 0
      theRowEast = 0
      theRowSouth = 0
      theRowWest = 0
      theRowUp = 0
      theRowDown = 0
      theColNorth = 0
      theColEast = 0
      theColSouth = 0
      theColWest = 0
      theColUp = 0
      theColDown = 0
      
      theExitNorth = False
      theExitEast = False
      theExitSouth = False
      theExitWest = False
      theExitUp = False
      theExitDown = False
      theDoorNorth = False
      theDoorEast = False
      theDoorSouth = False
      theDoorWest = False
      theDoorUp = False
      theDoorDown = False
      theHiddendoorNorth = False
      theHiddendoorEast = False
      theHiddendoorSouth = False
      theHiddendoorWest = False
      theHiddendoorUp = False
      theHiddendoorDown = False
      
      thePortalNorth = False
      thePortalEast = False
      thePortalSouth = False
      thePortalWest = False
      thePortalUp = False
      thePortalDown = False
      
      theDoorPortalNorth = False
      theDoorPortalEast = False
      theDoorPortalSouth = False
      theDoorPortalWest = False
      theDoorPortalUp = False
      theDoorPortalDown = False
      
      If theData > 0 Then
      
         If MappingMode = True And theRoomStringOk = False Then
            theRoomname = arrData(arrWorld(row, col), cROOMNAME)
            theRoomdesc = arrData(arrWorld(row, col), cDESCRIPTION)
            theRoomStringOk = True
         End If
         
         theTerrain = (theData And SPECIAL_MAP)
         If theTerrain = 0 Then theTerrain = (theData And TERRAIN_MAP)
         theFlag = (theData And FLAG_MAP)
         If (theData And 1) = 1 Then theSun = True
         If (theData And 2) = 2 Then theRide = True
         If (theData And MONSTER_MAP) = MONSTER_MAP Then theMonster = True
         
         Call readDirection(row, col, theData, theExitNorth, _
            N_MAP, N_noexit, N_exit, N_hiddendoor, _
            thePortalNorth, theHiddendoorNorth, theDoorPortalNorth, theDoorNorth, _
            arrData(arrWorld(row, col), cNDOOR), theDoornameNorth, arrData(arrWorld(row, col), cNPORTALR), theRowNorth, arrData(arrWorld(row, col), cNPORTALC), theColNorth)
            
         Call readDirection(row, col, theData, theExitEast, _
            E_MAP, E_noexit, E_exit, E_hiddendoor, _
            thePortalEast, theHiddendoorEast, theDoorPortalEast, theDoorEast, _
            arrData(arrWorld(row, col), cEDOOR), theDoornameEast, arrData(arrWorld(row, col), cEPORTALR), theRowEast, arrData(arrWorld(row, col), cEPORTALC), theColEast)
            
         Call readDirection(row, col, theData, theExitSouth, _
            S_MAP, S_noexit, S_exit, S_hiddendoor, _
            thePortalSouth, theHiddendoorSouth, theDoorPortalSouth, theDoorSouth, _
            arrData(arrWorld(row, col), cSDOOR), theDoornameSouth, arrData(arrWorld(row, col), cSPORTALR), theRowSouth, arrData(arrWorld(row, col), cSPORTALC), theColSouth)
            
         Call readDirection(row, col, theData, theExitWest, _
            W_MAP, W_noexit, W_exit, W_hiddendoor, _
            thePortalWest, theHiddendoorWest, theDoorPortalWest, theDoorWest, _
            arrData(arrWorld(row, col), cWDOOR), theDoornameWest, arrData(arrWorld(row, col), cWPORTALR), theRowWest, arrData(arrWorld(row, col), cWPORTALC), theColWest)
            
         Call readDirection(row, col, theData, theExitUp, _
            U_MAP, U_noexit, U_exit, U_hiddendoor, _
            thePortalUp, theHiddendoorUp, theDoorPortalUp, theDoorUp, _
            arrData(arrWorld(row, col), cUDOOR), theDoornameUp, arrData(arrWorld(row, col), cUPORTALR), theRowUp, arrData(arrWorld(row, col), cUPORTALC), theColUp)
            
         Call readDirection(row, col, theData, theExitDown, _
            D_MAP, D_noexit, D_exit, D_hiddendoor, _
            thePortalDown, theHiddendoorDown, theDoorPortalDown, theDoorDown, _
            arrData(arrWorld(row, col), cDDOOR), theDoornameDown, arrData(arrWorld(row, col), cDPORTALR), theRowDown, arrData(arrWorld(row, col), cDPORTALC), theColDown)
      End If
   End If
   
Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "load loadRoom"
   writeError (errorModule)
End Sub

Public Sub readDirection( _
   row, col, data, roomIs, map, _
   noExit, yesExit, Hidden, Portal, hiddenDoor, doorPortal, doorExit, _
   arrDoor, Doorname, _
   arrRow, rowValue, _
   arrCol, colValue)

   If (data And map) = noExit Then
      ' exit does not exist
   Else
      roomIs = True
      If (data And map) = yesExit Then
         ' there is an exit !
      Else
         If theRoomStringOk = False Then
            theRoomname = arrData(arrWorld(row, col), cROOMNAME)
            theRoomStringOk = True
         End If
         If LenB(arrDoor) <> 0 Then
            doorExit = True
            Doorname = arrDoor
            If (data And Hidden) = Hidden Then hiddenDoor = True
            If arrRow > 0 And arrCol > 0 Then
               doorPortal = True
               rowValue = arrRow
               colValue = arrCol
            End If
         Else
            If arrRow > 0 And arrCol > 0 Then
               Portal = True
               rowValue = arrRow
               colValue = arrCol
            End If
         End If
      End If
   End If
End Sub

Public Sub createData(data, _
                     specialRow, specialCol, _
                     whatExit, Doorname, Hidden, _
                     noExit, yesExit, _
                     doorExit, hiddenDoor, Portal, doorPortal)
   If whatExit = False Then
      data = (data Or noExit)
   Else
      If specialRow > 0 And specialCol > 0 Then
         If LenB(Doorname) <> 0 Then
            If Hidden Then data = (data Or hiddenDoor)
            data = (data Or doorPortal)
         Else
            data = (data Or Portal)
         End If
      Else
         If LenB(Doorname) <> 0 Then
            If Hidden = True Then data = (data Or hiddenDoor)
            data = (data Or doorExit)
         Else
            data = (data Or yesExit)
         End If
      End If
   End If
End Sub

