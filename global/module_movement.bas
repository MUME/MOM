Attribute VB_Name = "movement"
Option Explicit
Public limit As Integer
Public theROW  As Integer
Public theCOL As Integer
Public tmpRow As Integer
Public tmpCol As Integer
Public virtualRow As Integer
Public virtualCol As Integer
Public roomcount As Long
Public newRoomCount As Long
Public currentDir As String
Public tmpMove As String
Public currentData As Long
Public currentRoom
Public theMove As String
Public alasCount As Long

Public Function checkStringCI(data, search) As Boolean
   If InStrB(1, LCase(data), LCase(search), vbBinaryCompare) <> 0 Then checkStringCI = True Else checkStringCI = False
End Function

Public Function checkStringCS(data, search) As Boolean
   If InStrB(1, data, search, vbBinaryCompare) <> 0 Then checkStringCS = True Else checkStringCS = False
End Function

Public Sub resetBuffer()
errorData = errorData & "resetBuffer -> "
   roomcount = 0
   Erase arrMovestack
   Erase arrRoomstack
   virtualRow = theROW
   virtualCol = theCOL
End Sub
Public Sub cancelBuffer()
errorData = errorData & "cancelBuffer -> "
   If roomcount > 0 Then
      roomcount = limit
      virtualRow = arrRoomstack(limit, 1)
      virtualCol = arrRoomstack(limit, 2)
   End If
End Sub
Public Function chkFleeMove(direction As String) As Boolean
   chkFleeMove = False
   Select Case direction
   Case "n", "north"
      If setNewCoordinates(N_MAP, theROW - 1, theCOL, "n") Then chkFleeMove = True
   Case "e", "east"
      If setNewCoordinates(E_MAP, theROW, theCOL + 1, "e") Then chkFleeMove = True
   Case "s", "south"
      If setNewCoordinates(S_MAP, theROW + 1, theCOL, "s") Then chkFleeMove = True
   Case "w", "west"
      If setNewCoordinates(W_MAP, theROW, theCOL - 1, "w") Then chkFleeMove = True
   Case "u", "up"
      If setNewCoordinates(U_MAP, theROW, theCOL, "u") Then chkFleeMove = True
   Case "d", "down"
      If setNewCoordinates(D_MAP, theROW, theCOL, "d") Then chkFleeMove = True
   End Select
End Function

Public Function setNewCoordinates(theMap As Long, row, col, data As String) As Boolean
   setNewCoordinates = False
   Dim tmpRow As Long
   Dim tmpCol As Long
   'säilitame orignaalväärtused
   tmpRow = 0
   tmpCol = 0
   'On Error Resume Next
   If LenB(aData(getIndex(theROW, theCOL), cDATA)) = 0 Then Exit Function
   If (aData(getIndex(theROW, theCOL), cDATA) And theMap) > 0 Then ' room exists
      Select Case (aData(getIndex(theROW, theCOL), cDATA) And theMap)
      Case U_exit, U_door, U_portal, U_doorportal, U_hiddendoor, (U_hiddendoor Or U_portal)
         tmpRow = aData(getIndex(theROW, theCOL), cUPORTALR)
         tmpCol = aData(getIndex(theROW, theCOL), cUPORTALC)
         setNewCoordinates = True
      Case D_exit, D_door, D_portal, D_doorportal, D_hiddendoor, (D_hiddendoor Or D_portal)
         tmpRow = aData(getIndex(theROW, theCOL), cDPORTALR)
         tmpCol = aData(getIndex(theROW, theCOL), cDPORTALC)
         setNewCoordinates = True
      Case N_exit, N_door, N_hiddendoor, N_portal, N_doorportal, (N_hiddendoor Or N_portal)
         tmpRow = aData(getIndex(theROW, theCOL), cNPORTALR)
         tmpCol = aData(getIndex(theROW, theCOL), cNPORTALC)
         setNewCoordinates = True
      Case E_exit, E_door, E_hiddendoor, E_portal, E_doorportal, (E_hiddendoor Or E_portal)
         tmpRow = aData(getIndex(theROW, theCOL), cEPORTALR)
         tmpCol = aData(getIndex(theROW, theCOL), cEPORTALC)
         setNewCoordinates = True
      Case S_exit, S_door, S_hiddendoor, S_portal, S_doorportal, (S_hiddendoor Or S_portal)
         tmpRow = aData(getIndex(theROW, theCOL), cSPORTALR)
         tmpCol = aData(getIndex(theROW, theCOL), cSPORTALC)
         setNewCoordinates = True
      Case W_exit, W_door, W_hiddendoor, W_portal, W_doorportal, (W_hiddendoor Or W_portal)
         tmpRow = aData(getIndex(theROW, theCOL), cWPORTALR)
         tmpCol = aData(getIndex(theROW, theCOL), cWPORTALC)
         setNewCoordinates = True
      End Select
   End If
   theROW = tmpRow
   theCOL = tmpCol
   
Exit Function
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "movement setNewCoordinates"
   writeError (errorModule)
   setNewCoordinates = True
End Function
