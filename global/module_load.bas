Attribute VB_Name = "load"
Option Explicit
Public registryPath As String
Public filePath As String
Public thekeymaker As String
Public theData As Long
Public WorldLoaded As Boolean
Public Initialized As Boolean
Public Roomsync As Boolean
Public Autosync As Boolean
Public SyncError As Boolean
Public theCommand
Public freedom As Boolean
Public aWorld(1 To 300, 1 To 600, 0 To 0) As String
Public aData(0 To 30000, 0 To 31) As String
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
Public Const cLEVEL = 25
Public Const cNLEVEL = 26
Public Const cELEVEL = 27
Public Const cSLEVEL = 28
Public Const cWLEVEL = 29
Public Const cULEVEL = 30
Public Const cDLEVEL = 31



Public Sub loadWorld()
Dim yes As Boolean
Dim nn As Integer
Dim ss As String
If DEBUGMODE = False Then On Error GoTo errorhandler
   Dim key As Variant
   Dim readLine As String
   Dim vv As Variant
   Dim encrypted As Variant
   Dim original As Variant
   Dim failure As Boolean
   errorData = "loadWarpWorld -> "
   WorldLoaded = False
   key = getPassword()
   thekeymaker = key
   theCount = 0
   failure = False

   frmLogo.Show
   frmLogo.SetFocus
   frmLogo.Refresh
   'If Not fso.FileExists(filePath) Then
   '   Dim reply
   '   reply = MsgBox("You need to convert the map file." & vbCrLf & "This can take about a minute, please wait." & vbCrLf & "Please click OK to continue.", vbOKCancel, "Welcome to MUME Online Map!")
   '   If reply = vbOK Then
   '      frmLogo.SetFocus
   '      frmLogo.Refresh
   '      Call OLD_loadWorld
   '      MsgBox ("Map conversion successful. " & vbCrLf & "New map is named 'map51.txt'. Don't forget to backup!" & vbCrLf & "Please restart the program.")
   '   End If
   '   End
   'End If
   
   Dim errnum As Long
   
'#########  CONVERT PATH AND OLD KEY (set path, old and new key, run and save) ######
'filePath = "C:\map51.txt" 'change the filepath on 2nd run, also the data is saved into this filename
'key = cast128.cast128decode("780117demsi", "UZQe80IQeobJIG92ZQ/2kg==")
'##############################################

   Dim legacy As Boolean
   legacy = False 'piisab, kui yx on puudulik, siis on juba legaciga tegemist
   
   Open filePath For Input As #1
   Do While Not EOF(1)
      Line Input #1, readLine
      If LenB(readLine) <> 0 Then
         Dim v As Variant
         v = Split(readLine, ";", , vbBinaryCompare)
         encrypted = v(0) 'cENCRYPTED
         original = cast128.cast128decode(key, encrypted)
         If LenB(original) <> 0 Then

'MAP REPLACE CASE
'Dim dd, drow, dcol,
'drow_min = 130: drow_max = 70
'dcol_min = 200: dcol_max = 250
'dd = Split(original, ";", , vbBinaryCompare)
'drow = dd(2)
'dcol = dd(3)
'If (dcol >= dcol_min And (drow < drow_max Or drow > drow_min)) Or (dcol <= dcol_max And (drow < drow_max Or drow > drow_min)) Or (drow >= drow_max And (dcol < dcol_min Or dcol > dcol_max)) Or (drow <= drow_min And (dcol < dcol_min Or dcol > dcol_max)) Or drow < drow_max Or drow > drow_min Or dcol < dcol_min Or dcol > dcol_max Then
'do nothing
'Else 'to clean, run without else and 2nd time with new map, run with else

            If theCount < arrMaxData Then theCount = theCount + 1 Else failure = True: Exit Do
            'On Error Resume Next
            On Error GoTo 0
            vv = Split(original, ";", , vbBinaryCompare)
            aData(theCount, cDATA) = vv(0)
            aData(theCount, cDESCRIPTION) = vv(1)
            aData(theCount, cROW) = vv(2)
            aData(theCount, cCOL) = vv(3)
            aData(theCount, cROOMNAME) = v(1)
            aData(theCount, cNDOOR) = v(2)
            aData(theCount, cEDOOR) = v(3)
            aData(theCount, cSDOOR) = v(4)
            aData(theCount, cWDOOR) = v(5)
            aData(theCount, cUDOOR) = v(6)
            aData(theCount, cDDOOR) = v(7)
            aData(theCount, cNPORTALR) = v(8)
            aData(theCount, cNPORTALC) = v(9)
            aData(theCount, cEPORTALR) = v(10)
            aData(theCount, cEPORTALC) = v(11)
            aData(theCount, cSPORTALR) = v(12)
            aData(theCount, cSPORTALC) = v(13)
            aData(theCount, cWPORTALR) = v(14)
            aData(theCount, cWPORTALC) = v(15)
            aData(theCount, cUPORTALR) = v(16)
            aData(theCount, cUPORTALC) = v(17)
            aData(theCount, cDPORTALR) = v(18)
            aData(theCount, cDPORTALC) = v(19)
            aData(theCount, cNOTE) = v(20)
            If UBound(v) > 20 Then
               aData(theCount, cLEVEL) = v(21)
               aData(theCount, cNLEVEL) = v(22)
               aData(theCount, cELEVEL) = v(23)
               aData(theCount, cSLEVEL) = v(24)
               aData(theCount, cWLEVEL) = v(25)
               aData(theCount, cULEVEL) = v(26)
               aData(theCount, cDLEVEL) = v(27)
            End If
'legacy check
            If UBound(v) = 20 Then
               aData(theCount, cLEVEL) = 0
               aData(theCount, cNLEVEL) = 0
               aData(theCount, cELEVEL) = 0
               aData(theCount, cSLEVEL) = 0
               aData(theCount, cWLEVEL) = 0
               aData(theCount, cULEVEL) = 0
               aData(theCount, cDLEVEL) = 0
               legacy = True
            End If
            If (aData(theCount, cDATA) And TERRAIN_MAP) = 0 Then
               aData(theCount, cDATA) = (aData(theCount, cDATA) Or plain)
               aData(theCount, cDATA) = (aData(theCount, cDATA) Or ISROAD)
               legacy = True
            End If

'#########  CONVERT NEW KEY  #########
'thekeymaker = cast128.cast128decode("780117demsi", "UZQe80IQeobJIG92ZQ/2kg==")
'##############################################

            errnum = Err.Number
            If DEBUGMODE = False Then On Error GoTo errorhandler Else On Error GoTo 0
            If errnum <> 0 Then
               theCount = theCount - 1
            Else
               Call updateThis(theCount)
               aWorld(aData(theCount, cROW), aData(theCount, cCOL), aData(theCount, cLEVEL)) = theCount
            End If
            
   
'MAP REPLACE
'End If
           
            
         End If
      End If
   Loop
   Close #1
   
   If legacy Then
      Call makeportals
      Call makeroads
   End If
   
'GOTO MAP REPLACE COMMENT #1
   If failure Then MsgBox ("The world database is full!")
   If theCount = 0 Then MsgBox "Map is empty. If this was unintentional, then the map file is corrupted!"
   
   WorldLoaded = True
Exit Sub
errorhandler:
   WorldLoaded = False
   MsgBox "Invalid database or corrupted installation!" & vbCrLf & vbCrLf & Err.description & "(" & Err.Number & ")"
   errorModule = Err.description & "(" & Err.Number & ") -> " & "load loadWarpWorld"
   writeError (errorModule)
End Sub


Public Sub loadRoom(row As Integer, col As Integer)
errorData = errorData & "loadRoom -> "
If DEBUGMODE = False Then On Error GoTo errorhandler
Dim kala As Boolean
   If isValid(row, col) = True Then
      
      'theLEVEL = aData(getIndex(row, col), cLEVEL)
      
      theRoomStringOk = False
      theRoomname = vbNullString
      theRoomdesc = vbNullString
      SyncError = True
      If LOST = False And followMode = False And noexitsfound = False Then Call setNewExits(currentExits)
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
      theROWNorth = 0
      theROWEast = 0
      theROWSouth = 0
      theROWWest = 0
      theROWUp = 0
      theROWDown = 0
      theCOLNorth = 0
      theCOLEast = 0
      theCOLSouth = 0
      theCOLWest = 0
      theCOLUp = 0
      theCOLDown = 0
      
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
      
      
      If LenB(aData(getIndex(row, col), cDATA)) <> 0 Then
         theData = aData(getIndex(row, col), cDATA)
         
         'set old values to link the portals
         oldLevel = theLEVEL
         oldRow = theROW
         oldCol = theCOL
         
         If MappingMode = True And theRoomStringOk = False Then
            theRoomname = aData(getIndex(row, col), cROOMNAME)
            theRoomdesc = aData(getIndex(row, col), cDESCRIPTION)
            theRoomStringOk = True
         End If
         
         'theTerrain = (theData And SPECIAL_MAP)
         'If theTerrain = 0 Then
         theTerrain = (theData And TERRAIN_MAP)
         If (theData And ISROAD) = ISROAD Then theRoad = ISROAD
         theFlag = (theData And FLAG_MAP)
         If (theData And 1) = 1 Then theSun = True
         If (theData And 2) = 2 Then theRide = True
         If (theData And MONSTER_MAP) = MONSTER_MAP Then theMonster = True
         
         Call readDirection(row, col, theData, theExitNorth, _
            N_MAP, N_exit, N_hiddendoor, N_portal, N_doorportal, _
            thePortalNorth, theHiddendoorNorth, theDoorPortalNorth, theDoorNorth, _
            aData(getIndex(row, col), cNDOOR), theDoornameNorth, aData(getIndex(row, col), cNPORTALR), theROWNorth, aData(getIndex(row, col), cNPORTALC), theCOLNorth)
            
         Call readDirection(row, col, theData, theExitEast, _
            E_MAP, E_exit, E_hiddendoor, E_portal, E_doorportal, _
            thePortalEast, theHiddendoorEast, theDoorPortalEast, theDoorEast, _
            aData(getIndex(row, col), cEDOOR), theDoornameEast, aData(getIndex(row, col), cEPORTALR), theROWEast, aData(getIndex(row, col), cEPORTALC), theCOLEast)
            
         Call readDirection(row, col, theData, theExitSouth, _
            S_MAP, S_exit, S_hiddendoor, S_portal, S_doorportal, _
            thePortalSouth, theHiddendoorSouth, theDoorPortalSouth, theDoorSouth, _
            aData(getIndex(row, col), cSDOOR), theDoornameSouth, aData(getIndex(row, col), cSPORTALR), theROWSouth, aData(getIndex(row, col), cSPORTALC), theCOLSouth)
            
         Call readDirection(row, col, theData, theExitWest, _
            W_MAP, W_exit, W_hiddendoor, W_portal, W_doorportal, _
            thePortalWest, theHiddendoorWest, theDoorPortalWest, theDoorWest, _
            aData(getIndex(row, col), cWDOOR), theDoornameWest, aData(getIndex(row, col), cWPORTALR), theROWWest, aData(getIndex(row, col), cWPORTALC), theCOLWest)
            
         Call readDirection(row, col, theData, theExitUp, _
            U_MAP, U_exit, U_hiddendoor, U_portal, U_doorportal, _
            thePortalUp, theHiddendoorUp, theDoorPortalUp, theDoorUp, _
            aData(getIndex(row, col), cUDOOR), theDoornameUp, aData(getIndex(row, col), cUPORTALR), theROWUp, aData(getIndex(row, col), cUPORTALC), theCOLUp)
            
         Call readDirection(row, col, theData, theExitDown, _
            D_MAP, D_exit, D_hiddendoor, D_portal, D_doorportal, _
            thePortalDown, theHiddendoorDown, theDoorPortalDown, theDoorDown, _
            aData(getIndex(row, col), cDDOOR), theDoornameDown, aData(getIndex(row, col), cDPORTALR), theROWDown, aData(getIndex(row, col), cDPORTALC), theCOLDown)
      End If
   End If
   
Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "load loadRoom"
   writeError (errorModule)
End Sub

Public Sub readDirection(row As Integer, col As Integer, iData, roomIs, _
   map, iExit, iHiddenDoor, iPortal, iDoorPortal, _
   Portal, hiddenDoor, doorPortal, doorExit, _
   arrDoor, Doorname, _
   arrRow, rowValue, _
   arrCol, colValue)

   If (iData And map) = 0 Then
      ' exit does not exist
   Else
      roomIs = True
      
      If theRoomStringOk = False Then
         theRoomname = aData(getIndex(row, col), cROOMNAME)
         theRoomStringOk = True
      End If
      
      If LenB(arrDoor) <> 0 Then
         doorExit = True
         Doorname = arrDoor
      End If
      
      If (iHiddenDoor And iData) = iHiddenDoor Then hiddenDoor = True
      If (iPortal And iData) = iPortal Then Portal = True
      If (iDoorPortal And iData) = iDoorPortal Then doorPortal = True
      
      If arrRow > 0 And arrCol > 0 Then
         rowValue = getInt(CStr(arrRow))
         colValue = getInt(CStr(arrCol))
      End If
   End If
End Sub

Public Sub createData(data, _
                     specialRow, specialCol, _
                     whatExit, Doorname, Hidden, portalVisible, _
                     noExit, yesExit, _
                     doorExit, hiddenDoor, Portal, doorPortal)
                     
   If whatExit = False Then
      data = (data Or noExit)
   Else
      If LenB(Doorname) <> 0 Then
         If Hidden Then data = (data Or hiddenDoor)
         If portalVisible Then data = (data Or doorPortal) Else data = (data Or doorExit)
      Else
         If portalVisible Then data = (data Or Portal) Else data = (data Or yesExit)
      End If

'''      If specialRow > 0 And specialCol > 0 Then
'''         If LenB(Doorname) <> 0 Then
'''            If Hidden Then data = (data Or hiddenDoor)
'''            data = (data Or doorPortal)
'''         Else
'''            data = (data Or Portal)
'''         End If
'''      Else
'''         If LenB(Doorname) <> 0 Then
'''            If Hidden = True Then data = (data Or hiddenDoor)
'''            data = (data Or doorExit)
'''         Else
'''            data = (data Or yesExit)
'''         End If
'''      End If
   End If
End Sub

