Attribute VB_Name = "load"
Option Explicit
Option Compare Binary
Public theData As Long
Public WorldLoaded As Boolean
Public Room_Sync As Boolean
Public Auto_Sync As Boolean
Public SyncError As Boolean
Public theCommand

Public Sub loadWorld()
On Error GoTo errorhandler
   frmTools.status.ForeColor = &HC0FFC0
   frmTools.status.Caption = "Loading Arda! Please wait..."
   frmTools.Refresh
   errorData = "step -5"
   Dim cn As ADODB.Connection
   Set cn = New ADODB.Connection
   errorData = "step -4"
   Dim rstWorld As New ADODB.Recordset
   Set rstWorld = New ADODB.Recordset
   rstWorld.LockType = adLockReadOnly
   rstWorld.CursorType = adOpenStatic
   rstWorld.CursorLocation = adUseClient
   errorData = "step -3"
   Dim rstPortal As New ADODB.Recordset
   Set rstPortal = New ADODB.Recordset
   rstPortal.LockType = adLockReadOnly
   rstPortal.CursorType = adOpenStatic
   rstPortal.CursorLocation = adUseClient
   errorData = "step -2"
   Dim FileName As String
   Dim AccessConnStr As String
   FileName = App.Path & "\world.mdb"
   AccessConnStr = "Data Source=" & FileName & ";Provider=Microsoft.Jet.OLEDB.4.0;Mode=ReadWrite;Persist Security Info=False"
   errorData = "step 1"
   cn.Open AccessConnStr
   errorData = "step 2"
   rstWorld.Open "world", cn
   Set rstWorld.ActiveConnection = Nothing
   errorData = "step 3"
   rstPortal.Open "portal", cn
   errorData = "step 4"
   Set rstPortal.ActiveConnection = Nothing
   Dim tmpRow As Long
   Dim tmpCol As Long
   Dim n As Integer
   n = 0
   errorData = "step 5"
   Do While Not rstWorld.EOF
      n = n + 1
      tmpRow = rstWorld("row")
      tmpCol = rstWorld("col")
      arr(tmpRow, tmpCol) = rstWorld("arr")
      arrDesc(tmpRow, tmpCol) = rstWorld("arrdesc")
      arrRoomname(tmpRow, tmpCol) = rstWorld("roomname")
      arrDescription(tmpRow, tmpCol) = rstWorld("description")
      rstWorld.MoveNext
   Loop
   'DEBUG
   If n > 2000 Then
      MsgBox "Help!"
      End
   End If
   errorData = "step 6"
   Do While Not rstPortal.EOF
      arr(rstPortal("row"), rstPortal("col")) = rstPortal("portal")
      rstPortal.MoveNext
   Loop
   rstWorld.Close
   Set rstWorld = Nothing
   rstPortal.Close
   Set rstPortal = Nothing
   errorData = "step 7"
   frmTools.status.ForeColor = &HC0FFC0
   frmTools.status = "Load successful! Happy hunting."
   WorldLoaded = True

Exit Sub
errorhandler:
   errorData = errorData + "load loadWorld"
   writeError (errorData)
   WorldLoaded = False
   frmTools.status.ForeColor = &HFF&
   frmTools.status = "Load failed! Arda is in ruins!"
End Sub

Public Sub loadRoom(row As Long, col As Long)
On Error GoTo errorhandler
   If checkArrayLimit(row, col) = True Then
      theRoomStringOk = False
      theRoomname = ""
      theRoomdesc = ""
      theData = arr(row, col)

      SyncError = True
      If Room_Sync = True And Out_Of_Sync = False Then
         Call setNewExits(currentExits)
         If (newNorth And (theData And N_MAP) = 0) Then Exit Sub
         If (newEast And (theData And E_MAP) = 0) Then Exit Sub
         If (newSouth And (theData And S_MAP) = 0) Then Exit Sub
         If (newWest And (theData And W_MAP) = 0) Then Exit Sub
         If (newUp And (theData And U_MAP) = 0) Then Exit Sub
         If (newDown And (theData And D_MAP) = 0) Then Exit Sub
         
         theRoomString = Split(arrDesc(row, col), ";")
         theRoomname = arrRoomname(row, col)    'theRoomString(0)
         theRoomdesc = arrDescription(row, col) 'theRoomString(19)
         theRoomStringOk = True
         If theRoomname <> currentRoomName Then Exit Sub
      End If
      SyncError = False
      
      theTerrain = 0
      theRide = False
      theSun = False
      theMonster = False
      theDoornameNorth = ""
      theDoornameEast = ""
      theDoornameSouth = ""
      theDoornameWest = ""
      theDoornameUp = ""
      theDoornameDown = ""
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
      
      With frmTools
      If theData > 0 Then
      
         If MappingMode = True And theRoomStringOk = False Then
            theRoomString = Split(arrDesc(row, col), ";")
            theRoomname = arrRoomname(row, col)    'theRoomString(0)
            theRoomdesc = arrDescription(row, col) 'theRoomString(19)
            theRoomStringOk = True
         End If
         
         theTerrain = (theData And TERRAIN_MAP)
         If (theData And 1) = 1 Then theSun = True
         If (theData And 2) = 2 Then theRide = True
         If (theData And MONSTER_MAP) = MONSTER_MAP Then theMonster = True
         
         Call readDirection(row, col, theData, .nHidden, theExitNorth, _
            N_MAP, N_noexit, N_exit, N_hiddendoor, _
            thePortalNorth, theHiddendoorNorth, theDoorPortalNorth, theDoorNorth, _
            1, theDoornameNorth, 2, theRowNorth, 3, theColNorth)
            
         Call readDirection(row, col, theData, .eHidden, theExitEast, _
            E_MAP, E_noexit, E_exit, E_hiddendoor, _
            thePortalEast, theHiddendoorEast, theDoorPortalEast, theDoorEast, _
            4, theDoornameEast, 5, theRowEast, 6, theColEast)
            
         Call readDirection(row, col, theData, .sHidden, theExitSouth, _
            S_MAP, S_noexit, S_exit, S_hiddendoor, _
            thePortalSouth, theHiddendoorSouth, theDoorPortalSouth, theDoorSouth, _
            7, theDoornameSouth, 8, theRowSouth, 9, theColSouth)
            
         Call readDirection(row, col, theData, .wHidden, theExitWest, _
            W_MAP, W_noexit, W_exit, W_hiddendoor, _
            thePortalWest, theHiddendoorWest, theDoorPortalWest, theDoorWest, _
            10, theDoornameWest, 11, theRowWest, 12, theColWest)
            
         Call readDirection(row, col, theData, .uHidden, theExitUp, _
            U_MAP, U_noexit, U_exit, U_hiddendoor, _
            thePortalUp, theHiddendoorUp, theDoorPortalUp, theDoorUp, _
            13, theDoornameUp, 14, theRowUp, 15, theColUp)
            
         Call readDirection(row, col, theData, .dHidden, theExitDown, _
            D_MAP, D_noexit, D_exit, D_hiddendoor, _
            thePortalDown, theHiddendoorDown, theDoorPortalDown, theDoorDown, _
            16, theDoornameDown, 17, theRowDown, 18, theColDown)

      End If
      .row.Caption = theRow
      .col.Caption = theCol
      End With
   End If
Exit Sub

errorhandler:
   errorData = "load loadRoom"
   writeError (errorData)
   frmTools.status = "Invalid room data. Load cancelled!"
End Sub

Public Sub readDirection( _
   ByRef row, ByRef col, ByRef data, ByRef control, ByRef roomIs, ByRef map, _
   ByRef NoExit, ByRef YesExit, ByRef Hidden, ByRef Portal, ByRef HiddenDoor, ByRef DoorPortal, ByRef DoorExit, _
   ByVal arrDoor, ByRef Doorname, _
   ByVal arrRow, ByRef rowValue, _
   ByVal arrCol, ByRef colValue)

'   control.Value = 0
'   control.Visible = False
   If (data And map) = NoExit Then
      ' exit does not exist
   Else
      roomIs = True
      If (data And map) = YesExit Then
         ' there is an exit !
      Else
         If theRoomStringOk = False Then
            theRoomname = arrRoomname(row, col)
            theRoomString = Split(arrDesc(row, col), ";")
            theRoomStringOk = True
         End If
         If Len(theRoomString(arrDoor)) > 0 Then
 '           control.Visible = True
            DoorExit = True
            Doorname = theRoomString(arrDoor)
            If (data And Hidden) = Hidden Then
 '              control.Value = 1
               HiddenDoor = True
            End If
            If theRoomString(arrRow) > 0 And theRoomString(arrCol) > 0 Then
               DoorPortal = True
               rowValue = theRoomString(arrRow)
               colValue = theRoomString(arrCol)
            End If
         Else
            If theRoomString(arrRow) > 0 And theRoomString(arrCol) > 0 Then
               Portal = True
               rowValue = theRoomString(arrRow)
               colValue = theRoomString(arrCol)
            End If
         End If
      End If
   End If
End Sub

Public Sub createData(ByRef data, ByRef desc, _
                     ByRef specialRow, ByRef specialCol, _
                     ByRef whatExit, ByRef Doorname, ByRef Hidden, _
                     ByRef NoExit, ByRef YesExit, _
                     ByRef DoorExit, ByRef HiddenDoor, ByRef Portal, ByRef DoorPortal)

   If whatExit = False Then
      data = (data Or NoExit)
      desc = desc & ";0;0;"
   Else
      If specialRow > 0 And specialCol > 0 Then
         
         If Len(Doorname) > 0 Then
            If Hidden Then data = (data Or HiddenDoor)
            data = (data Or DoorPortal)
            desc = desc & Doorname & ";"
         Else
            data = (data Or Portal)
            desc = desc & ";"
         End If
         desc = desc & specialRow & ";"
         desc = desc & specialCol & ";"
      
      Else
         
         If Len(Doorname) > 0 Then
            If Hidden = True Then data = (data Or HiddenDoor)
            data = (data Or DoorExit)
            desc = desc & Doorname & ";0;0;"
         Else
            data = (data Or YesExit)
            desc = desc & ";0;0;"
         End If
      
      End If
   End If
End Sub
