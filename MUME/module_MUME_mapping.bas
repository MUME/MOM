Attribute VB_Name = "MUME_Mapping"
Option Explicit
Dim n0 As Integer, n1 As Integer, n2 As Integer, n3 As Integer, n4 As Integer
Dim a As Integer, b As Integer, c As Integer
Public retry As Integer
Public debugoutput As String

Public Function handleMapping(ByRef strData As String)
If DEBUGMODE = False Then On Error GoTo errorhandler Else On Error GoTo 0
   errorData = errorData & "handleMapping -> "
   handleMapping = False
   
'# 'If MappingData = True Then
   If MappingData = True Then
      Select Case MappingCase
      Case 2
         n0 = InStrB(1, strData, lookColour, vbBinaryCompare)
         If n0 > 0 Then
            n1 = InStrB(1, MidB(strData, 1, n0), "You flee head over heels.", vbBinaryCompare)
            If n1 > 0 Then
               MappingCase = 0
               MappingData = False
               MappingGetUpdate = False
               Call informClient("Mapping cancelled!")
               theRow = mappingFromRow
               theCol = mappingFromCol
               handleMapping = True: Exit Function
            End If
         Else
            tmpCheck = True
Dim i As Integer
For i = LBound(arrCollision) To UBound(arrCollision)
   If LenB(arrCollision(i)) = 0 Then Exit For
   If tmpCheck = False Then Exit For
   If checkStringCS(strData, arrCollision(i)) Then tmpCheck = False
Next
            
            If tmpCheck Then If checkStringCS(strData, "It is pitch black...") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "You just see a dense fog around you...") Then tmpCheck = False
            If tmpCheck = False Then
               MappingCase = 0
               MappingData = False
               MappingGetUpdate = False
               Call informClient("Mapping cancelled!")
               theRow = mappingFromRow
               theCol = mappingFromCol
               handleMapping = True: Exit Function
            Else
               handleMapping = True: Exit Function
            End If
         End If
         
         Call zeroMap
'read
         mapRoomName = getRoomname(strData)
         frmTools.Roomname = mapRoomName
         mapDescription = getRoomDescription(strData)
'read EXITS
         Dim e1, e2
         e1 = 0: e2 = 0
         'e1 = InStrB(c, strData, "Exits: ", vbBinaryCompare)
         e1 = InStrB(1, strData, "Exits: ", vbBinaryCompare)
         If e1 > 0 Then
            e2 = InStrB(e1 + LenB("Exits: "), strData, vbCr, vbBinaryCompare)
            Dim sExits As String
            sExits = MidB(strData, e1 + LenB("Exits: "), e2 - (e1 + LenB("Exits: ")))
            If InStrB(1, sExits, "north", vbBinaryCompare) > 0 Then
               mapExitNorth = True
               frmTools.nExit = 1
               If InStrB(1, sExits, "(north)", vbBinaryCompare) > 0 Or InStrB(1, sExits, "[north]", vbBinaryCompare) > 0 Then mapDoornameNorth = "exit north": frmTools.nDoor = "exit north"
            End If
            If InStrB(1, sExits, "east", vbBinaryCompare) > 0 Then
               mapExitEast = True
               frmTools.eExit = 1
               If InStrB(1, sExits, "(east)", vbBinaryCompare) > 0 Or InStrB(1, sExits, "[east]", vbBinaryCompare) > 0 Then mapDoornameEast = "exit east": frmTools.eDoor = "exit east"
            End If
            If InStrB(1, sExits, "south", vbBinaryCompare) > 0 Then
               mapExitSouth = True
               frmTools.sExit = 1
               If InStrB(1, sExits, "(south)", vbBinaryCompare) > 0 Or InStrB(1, sExits, "[south]", vbBinaryCompare) > 0 Then mapDoornameSouth = "exit south": frmTools.sDoor = "exit south"
            End If
            If InStrB(1, sExits, "west", vbBinaryCompare) > 0 Then
               mapExitWest = True
               frmTools.wExit = 1
               If InStrB(1, sExits, "(west)", vbBinaryCompare) > 0 Or InStrB(1, sExits, "[west]", vbBinaryCompare) > 0 Then mapDoornameWest = "exit west": frmTools.wDoor = "exit west"
            End If
            If InStrB(1, sExits, "up", vbBinaryCompare) > 0 Then
               mapExitUp = True
               frmTools.uExit = 1
               If InStrB(1, sExits, "(up)", vbBinaryCompare) > 0 Or InStrB(1, sExits, "[up]", vbBinaryCompare) > 0 Then mapDoornameUp = "exit up": frmTools.uDoor = "exit up"
            End If
            If InStrB(1, sExits, "down", vbBinaryCompare) > 0 Then
               mapExitDown = True
               frmTools.dExit = 1
               If InStrB(1, sExits, "(down)", vbBinaryCompare) > 0 Or InStrB(1, sExits, "[down]", vbBinaryCompare) > 0 Then mapDoornameDown = "exit down": frmTools.dDoor = "exit down"
            End If
'read TERRAIN
            Dim y1 As Integer, y2 As Integer, s As String
            
            y1 = InStrRev(strData, vbCrLf, , vbBinaryCompare)
            s = Mid(strData, y1 + Len(vbCrLf) + 1, 1)
            setMapTerrain (s)
         End If
         MappingData = False
         If MappingGetUpdate = True Then
            MappingGetUpdate = False
            dataFromMUD = True
            Call mapUpdate
         End If
         MappingCase = 0
         handleMapping = True: Exit Function
      End Select
   End If
Exit Function

errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "Mume_Mapping handleMapping"
   writeError (errorModule)
End Function

