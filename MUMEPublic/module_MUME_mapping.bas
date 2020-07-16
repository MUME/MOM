Attribute VB_Name = "MUME_Mapping"
Option Explicit
Dim n0 As Integer, n1 As Integer, n2 As Integer, n3 As Integer, n4 As Integer
Dim a As Integer, b As Integer, c As Integer
Public retry As Integer
Public debugoutput As String

Public Function handleMapping(strData As String)
If DEBUGMODE = False Then On Error GoTo errorhandler
   errorData = errorData & "handleMapping -> "
   handleMapping = False
   
'# 'If MappingData = True Then
   If MappingData = True Then
      Select Case MappingCase
      Case 2
         n0 = InStr(1, strData, lookColour, vbBinaryCompare)
         If n0 > 0 Then
            n1 = InStr(1, Mid(strData, 1, n0), "You flee head over heels.", vbBinaryCompare)
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
            If tmpCheck Then If checkStringCS(strData, "No way! You are fighting for your life!") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "Alas, you cannot go that way...") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "Oops! You cannot go there riding!") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "You can't go into deep water!") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "Your mount refuses to follow your orders!") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "doesn't want you riding") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, " seems to be closed.") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "You failed swimming there.") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "You failed to climb there and fall down, hurting yourself.") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, " too exhausted.") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "Maybe you should get on your feet first?") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "Nah... You feel too relaxed to do that..") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "In your dreams, or what?") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "It is pitch black...") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "You just see a dense fog around you...") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "The descent is too steep, you need to climb to go there.") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "The ascent is too steep, you need to climb to go there.") Then tmpCheck = False
            If tmpCheck Then If checkStringCS(strData, "You need to swim to go there.") Then tmpCheck = False
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
'read ROOMNAME
         n1 = InStr(strData, lookColour)
         If n1 > 0 Then
            n2 = InStr(n1 + 5, strData, colourEndCode)
            If n2 > 0 Then
               mapRoomName = Mid(strData, n1 + 5, n2 - (n1 + 5))
'read DESCRIPTION
               a = n2 + 6
               b = InStr(a, strData, vbCrLf)
               c = InStrRev(strData, vbLf, b)
               If c > a Then
                  mapDescription = Mid(strData, a, c - a)
                  frmTools.Roomname = mapRoomName
               End If
            End If
         End If
'read EXITS
         Dim e1, e2
         e1 = 0: e2 = 0
         e1 = InStr(c, strData, "Exits: ", vbBinaryCompare)
         If e1 > 0 Then
            e2 = InStr(e1 + 7, strData, vbCr, vbBinaryCompare)
            Dim sExits As String
            sExits = Mid(strData, e1 + 7, e2 - (e1 + 7))
            If InStr(1, sExits, "north") > 0 Then
               mapExitNorth = True
               frmTools.nExit = 1
               If InStr(1, sExits, "(north)") > 0 Or InStr(1, sExits, "[north]") > 0 Then mapDoornameNorth = "exit north" ': frmTools.nDoor = "exit north"
            End If
            If InStr(1, sExits, "east") > 0 Then
               mapExitEast = True
               frmTools.eExit = 1
               If InStr(1, sExits, "(east)") > 0 Or InStr(1, sExits, "[east]") > 0 Then mapDoornameEast = "exit east" ': frmTools.eDoor = "exit east"
            End If
            If InStr(1, sExits, "south") > 0 Then
               mapExitSouth = True
               frmTools.sExit = 1
               If InStr(1, sExits, "(south)") > 0 Or InStr(1, sExits, "[south]") > 0 Then mapDoornameSouth = "exit south" ': frmTools.sDoor = "exit south"
            End If
            If InStr(1, sExits, "west") > 0 Then
               mapExitWest = True
               frmTools.wExit = 1
               If InStr(1, sExits, "(west)") > 0 Or InStr(1, sExits, "[west]") > 0 Then mapDoornameWest = "exit west" ': frmTools.wDoor = "exit west"
            End If
            If InStr(1, sExits, "up") > 0 Then
               mapExitUp = True
               frmTools.uExit = 1
               If InStr(1, sExits, "(up)") > 0 Or InStr(1, sExits, "[up]") > 0 Then mapDoornameUp = "exit up" ': frmTools.uDoor = "exit up"
            End If
            If InStr(1, sExits, "down") > 0 Then
               mapExitDown = True
               frmTools.dExit = 1
               If InStr(1, sExits, "(down)") > 0 Or InStr(1, sExits, "[down]") > 0 Then mapDoornameDown = "exit down" ': frmTools.dDoor = "exit down"
            End If
'read TERRAIN
            Dim y1 As Integer, y2 As Integer, s As String
            y1 = InStrRev(strData, ">", , vbBinaryCompare)
            If y1 > 0 Then
               y2 = InStrRev(strData, vbCr, y1, vbBinaryCompare)
               s = Mid(strData, y2 + 1, 1)
               If s = vbLf Then
                  s = Mid(strData, y2 + 3, 1)
               Else
                  s = Mid(strData, y2 + 2, 1)
               End If
               setMapTerrain (s)
            End If
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
   errorModule = Err.Description & "(" & Err.Number & ") -> " & "Mume_Mapping handleMapping"
   writeError (errorModule)
End Function

