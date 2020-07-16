Attribute VB_Name = "load"
Option Explicit
Option Compare Binary
Public theData As Long
Public WorldLoaded As Boolean

Public Sub LoadWorld()
   BestEST.status.ForeColor = &HC0FFC0
   BestEST.status.Caption = "Loading Arda! Please wait..."
   BestEST.Refresh
   Dim cn As ADODB.Connection
   Set cn = New ADODB.Connection
   Dim rstWorld As New Recordset
   Set rstWorld = New ADODB.Recordset
   rstWorld.LockType = adLockReadOnly
   rstWorld.CursorType = adOpenStatic
   rstWorld.CursorLocation = adUseClient
   Dim rstPortal As New Recordset
   Set rstPortal = New ADODB.Recordset
   rstPortal.LockType = adLockReadOnly
   rstPortal.CursorType = adOpenStatic
   rstPortal.CursorLocation = adUseClient
   Dim filename As String
   Dim AccessConnStr As String
   filename = "C:\mume\world.mdb"
   AccessConnStr = "Data Source=" & filename & ";Provider=Microsoft.Jet.OLEDB.4.0;Mode=ReadWrite;Persist Security Info=False"
   cn.Open AccessConnStr
   rstWorld.Open "world", cn
   Set rstWorld.ActiveConnection = Nothing
   rstPortal.Open "portal", cn
   Set rstPortal.ActiveConnection = Nothing
 
   Dim e
   Do While Not rstWorld.EOF
      arr(rstWorld("row"), rstWorld("col")) = rstWorld("arr")
      arrDesc(rstWorld("row"), rstWorld("col")) = rstWorld("arrdesc")
      rstWorld.MoveNext
   Loop
   Do While Not rstPortal.EOF
      arr(rstPortal("row"), rstPortal("col")) = rstPortal("portal")
      rstPortal.MoveNext
   Loop
   rstWorld.Close
   Set rstWorld = Nothing
   rstPortal.Close
   Set rstPortal = Nothing
   BestEST.status.ForeColor = &HC0FFC0
   BestEST.status = "Arda has been loaded!"
   WorldLoaded = True
Exit Sub
ErrorHandler:
   WorldLoaded = False
   BestEST.status.ForeColor = &HFF&
   BestEST.status = "ARDA IS IN RUINS!"
End Sub

Public Sub LoadRoom(row As Long, col As Long)
   If checkArrayLimit(row, col) = True Then
      theRoomStringOk = False
      theSpecialNorth = False
      theSpecialEast = False
      theSpecialSouth = False
      theSpecialWest = False
      theSpecialUp = False
      theSpecialDown = False
      theRide = False
      theSun = False
      theRoomName = ""
      theRoomDesc = ""
      theDoorNameNorth = ""
      theDoorNameEast = ""
      theDoorNameSouth = ""
      theDoorNameWest = ""
      theDoorNameUp = ""
      theDoorNameDown = ""
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
      theRoomNorth = False
      theRoomEast = False
      theRoomSouth = False
      theRoomWest = False
      theRoomUp = False
      theRoomDown = False
      theDoorNorth = False
      theDoorEast = False
      theDoorSouth = False
      theDoorWest = False
      theDoorUp = False
      theDoorDown = False
      theData = arr(row, col)
      With BestEST
      If theData > 0 Then
         If MAP_MODE = True Then
            theRoomString = Split(arrDesc(row, col), ";")
            theRoomName = theRoomString(0)
            theRoomDesc = theRoomString(19)
            theRoomStringOk = True
         End If
         If (theData And 1) = 1 Then theSun = True
         If (theData And 2) = 2 Then theRide = True
         Call readDirection(row, col, theData, .n_doorname, theRoomNorth, N_map, _
            N_noexit, N_exit, theDoorNorth, theSpecialNorth, _
            1, theDoorNameNorth, 2, theRowNorth, 3, theColNorth)
         Call readDirection(row, col, theData, .e_doorname, theRoomEast, E_map, _
            E_noexit, E_exit, theDoorEast, theSpecialEast, _
            4, theDoorNameEast, 5, theRowEast, 6, theColEast)
         Call readDirection(row, col, theData, .s_doorname, theRoomSouth, S_map, _
            S_noexit, S_exit, theDoorSouth, theSpecialSouth, _
            7, theDoorNameSouth, 8, theRowSouth, 9, theColSouth)
         Call readDirection(row, col, theData, .w_doorname, theRoomWest, W_map, _
            W_noexit, W_exit, theDoorWest, theSpecialWest, _
            10, theDoorNameWest, 11, theRowWest, 12, theColWest)
         Call readDirection(row, col, theData, .u_doorname, theRoomUp, U_map, _
            U_noexit, U_exit, theDoorUp, theSpecialUp, _
            13, theDoorNameUp, 14, theRowUp, 15, theColUp)
         Call readDirection(row, col, theData, .d_doorname, theRoomDown, D_map, _
            D_noexit, D_exit, theDoorDown, theSpecialDown, _
            16, theDoorNameDown, 17, theRowDown, 18, theColDown)
      End If
      .row.Caption = theRow
      .col.Caption = theCol
      End With
   End If
End Sub

Public Sub readDirection( _
   ByRef row, ByRef col, ByRef data, ByRef control, ByRef roomIs, ByRef map, _
   ByRef Noexit, ByRef Yesexit, ByRef Doorexit, ByRef Specialexit, _
   ByVal arrDoor, ByRef Doorname, _
   ByVal arrRow, ByRef rowValue, _
   ByVal arrCol, ByRef colValue)
  control.Visible = False                                 'theRoomNorth = False
  If (data And map) = Noexit Then
  Else
    roomIs = True                                           'theRoomNorth = True
    If (data And map) = Yesexit Then
      control.Caption = ""
    Else
      If theRoomStringOk = False Then
        theRoomString = Split(arrDesc(row, col), ";")
        theRoomName = theRoomString(0)
        theRoomStringOk = True
      End If
      If Len(theRoomString(arrDoor)) > 0 Then
        Doorexit = True
        Doorname = theRoomString(arrDoor)
        control.Visible = True
        control.Caption = Doorname
      Else
        control.Caption = ""
      End If
      If theRoomString(arrRow) > 0 And theRoomString(arrCol) > 0 Then
        Specialexit = True
        rowValue = theRoomString(arrRow)
        colValue = theRoomString(arrCol)
      End If
    End If
  End If
End Sub

Public Sub LoadArea(ByVal filename As String, toRow As Long, toCol As Long)
  Dim cn As ADODB.Connection
  Set cn = New ADODB.Connection
  Dim rst As New Recordset
  Set rst = New ADODB.Recordset
  rst.LockType = adLockReadOnly
  rst.CursorType = adOpenStatic
  rst.CursorLocation = adUseClient
  With cn
      .Provider = "Microsoft.Jet.OLEDB.4.0"
      .ConnectionString = "Data Source=" & filename & ";" & _
      "Extended Properties=Excel 8.0;"
      .Open
  End With
  Dim sql As String
  sql = "SELECT * FROM [Data]"
  rst.Open sql, cn
  Set rst.ActiveConnection = Nothing
  Dim row As Long
  Dim col As Long
  Dim theValue As Long

Do While Not rst.EOF
   theValue = 0
   theDesc = ""
   'row, column
   row = toRow + rst.Fields(0)
   col = toCol + rst.Fields(1)
   'terrain
   theValue = theValue + rst.Fields(3).Value
   'ride/noride, sun/dark
   If rst.Fields(4).Value = False And rst.Fields(5).Value = False Then theValue = theValue + noRide_Dark
   If rst.Fields(4).Value = False And rst.Fields(5).Value = True Then theValue = theValue + noRide_Sun
   If rst.Fields(4).Value = True And rst.Fields(5).Value = False Then theValue = theValue + Ride_Dark
   If rst.Fields(4).Value = True And rst.Fields(5).Value = True Then theValue = theValue + Ride_Sun
   'roomname
   If Len(rst.Fields(2)) > 0 Then theDesc = theDesc & rst.Fields(2) & ";"
   'DIRECTIONS       boolean          -                                            doorname           row             col
   Call createData(rst.Fields(6), theValue, theDesc, N_noexit, N_exit, N_door, N_special, rst.Fields(12), rst.Fields(18), rst.Fields(19))
   Call createData(rst.Fields(7), theValue, theDesc, E_noexit, E_exit, E_door, E_special, rst.Fields(13), rst.Fields(20), rst.Fields(21))
   Call createData(rst.Fields(8), theValue, theDesc, S_noexit, S_exit, S_door, S_special, rst.Fields(14), rst.Fields(22), rst.Fields(23))
   Call createData(rst.Fields(9), theValue, theDesc, W_noexit, W_exit, W_door, W_special, rst.Fields(15), rst.Fields(24), rst.Fields(25))
   Call createData(rst.Fields(10), theValue, theDesc, U_noexit, U_exit, U_door, U_special, rst.Fields(16), rst.Fields(26), rst.Fields(27))
   Call createData(rst.Fields(11), theValue, theDesc, D_noexit, D_exit, D_door, D_special, rst.Fields(17), rst.Fields(28), rst.Fields(29))
   'validity check
   If theValue > 0 Then
      arr(row, col) = theValue
      arrDesc(row, col) = theDesc
   End If
   rst.MoveNext
Loop
End Sub

Public Sub createData(ByRef whatExit, ByRef data, ByRef desc, ByRef Noexit, ByRef Yesexit, ByRef Doorexit, ByRef Specialexit, ByRef Doorname, ByRef specialRow, ByRef specialCol)
   If whatExit = False Then
      data = data + Noexit
      desc = desc & ";0;0;"
   Else
      If specialRow > 0 And specialCol > 0 Then
         data = data + Specialexit
         If Len(Doorname) > 0 Then
           desc = desc & Doorname & ";"
         Else
            desc = desc & ";"
         End If
         desc = desc & specialRow & ";"
         desc = desc & specialCol & ";"
      Else
         If Len(Doorname) > 0 Then
            data = data + Doorexit
            desc = desc & Doorname & ";0;0;"
         Else
            data = data + Yesexit
            desc = desc & ";0;0;"
         End If
      End If
   End If
End Sub
