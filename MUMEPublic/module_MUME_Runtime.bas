Attribute VB_Name = "MUME_Runtime"
Option Explicit
Dim n As Integer
Dim n0 As Integer, n1 As Integer, n2 As Integer, n3 As Integer, n4 As Integer
Dim a As Integer, b As Integer, c As Integer
Public lookColour As String
Public fgColour As String
Public bgColour As String
Public fgBold As String
Public fgUnderline As String
Public tmpCheck As Boolean
Public noexitsfound As Boolean
Public leader As String

Public arrPlayers(10) As String
Public arrPlayersNames(10) As String
Public arrPlayersIndex As Integer
Public viewPlayers As Boolean
Public viewTarget As Boolean
Public target As String
Public targetRow As Integer
Public targetCol As Integer
Public isYou1 As Boolean
Public isYou2 As Boolean

Public Function handleRunMode(ByVal strData As String)
If DEBUGMODE = False Then On Error GoTo errorhandler
errorData = errorData & "handleRunMode -> "
handleRunMode = False
   n0 = InStr(1, strData, lookColour, vbBinaryCompare)
   If n0 < 1 Then Exit Function
   n2 = InStr(n0, strData, colourEndCode, vbBinaryCompare)
   If n2 < 1 Then Exit Function
'FLEEING
   If LOST = False Then
      canUndo = False
      undoRow = virtualRow
      undoCol = virtualCol
       n1 = InStr(1, Mid(strData, 1, n0), "You flee head over heels.", vbBinaryCompare)
      If n1 > 0 Then
         errorData = errorData & "fleeing -> "
         n1 = n0 + Len(lookColour)  '5
         If n2 > 0 Then
            currentRoomName = Mid(strData, n1, n2 - n1)
            n3 = InStr(n2, strData, "Exits: ", vbBinaryCompare)
            If n3 > 0 Then
'YOU FLED
               n4 = InStr(n3 + 7, strData, ".", vbBinaryCompare)
               currentExits = Mid(strData, n3 + Len(colourEndCode), n4 - (n3 + Len(colourEndCode)))
               Call caseFleeHandler(currentRoomName, currentExits, 1, False, True)
               If LOST = False Then
                  If roomcount > 0 Then
                     For n = stackOUT To stackIN
                        If chkFleeMove(arrMovestack(n)) Then
                           virtualRow = theRow
                           virtualCol = theCol
                        End If
                     Next
                     currentRoomName = arrData(arrWorld(theRow, theCol), cROOMNAME)
                     Call loadRoom(theRow, theCol)
                     Call DrawMap
                  End If
                  'new _back feature
                  canUndo = True
               Else
                  fleeRetry = 0
               End If
               handleRunMode = True: Exit Function
            Else
               Call SYNC_FALSE("HandleRunMode, flee message without colour endcode!"): handleRunMode = True: Exit Function
            End If
         End If
      End If
   End If

   followMode = False
'#SE FOLLOW
'   If frmMap.mnuFollow.Checked Then
'      errorData = errorData & "follow -> "
'      If LOST = False Then
'         If n2 > 0 Then
'            currentRoomName = Mid(strData, n0 + Len(lookColour), n2 - (n0 + Len(lookColour)))
'            currentString = strData 'Mid(strData, 1, n0)
'            If Len(currentString) > Len("leaves up") Then
'               If InStr(1, currentString, leader & " leaves north", vbBinaryCompare) Then _
'                  followMode = True: Call travel("n", roomcount, newNorth, N_MAP, virtualRow - 1, virtualCol)
'               If InStr(1, currentString, leader & " leaves east", vbBinaryCompare) Then _
'                  followMode = True: Call travel("e", roomcount, newEast, E_MAP, virtualRow, virtualCol + 1)
'               If InStr(1, currentString, leader & " leaves south", vbBinaryCompare) Then _
'                  followMode = True: Call travel("s", roomcount, newSouth, S_MAP, virtualRow + 1, virtualCol)
'               If InStr(1, currentString, leader & " leaves west", vbBinaryCompare) Then _
'                  followMode = True: Call travel("w", roomcount, newWest, W_MAP, virtualRow, virtualCol - 1)
'               If InStr(1, currentString, leader & " leaves up", vbBinaryCompare) Then _
'                  followMode = True: Call travel("u", roomcount, newUp, U_MAP, virtualRow, virtualCol)
'               If InStr(1, currentString, leader & " leaves down", vbBinaryCompare) Then _
'                  followMode = True: Call travel("d", roomcount, newDown, D_MAP, virtualRow, virtualCol)
'               If followMode = True And roomcount > 0 Then Call newUpdateTheRoom: handleRunMode = True: Exit Function
'            End If
'         End If
'      End If
'   End If

'WALKING
   If n2 > 0 Then
      errorData = errorData & "walking -> "
      currentRoomName = Mid(strData, n0 + 5, n2 - (n0 + 5))
      n3 = InStr(n2, strData, "Exits: ", vbBinaryCompare)
      If n3 > 0 Then
'YOU MOVED
         noexitsfound = False
         n4 = InStr(n3 + 7, strData, ".", vbBinaryCompare)
         currentExits = Mid(strData, n3 + 6, n4 - (n3 + 6))
         If Autosync = True And LOST = True Then
            If fleeRetry < 5 Then
               fleeRetry = fleeRetry + 1
               Call caseFleeHandler(currentRoomName, currentExits, fleeRetry + 1, False, False)
               If LOST = False Then handleRunMode = True: Exit Function
            End If
         End If
      Else
         theExitNorth = True
         theExitEast = True
         theExitSouth = True
         theExitWest = True
         theExitUp = True
         theExitDown = True
      End If

      If LOST = False Then
         If roomcount > 0 Then
            Call newUpdateTheRoom
            handleRunMode = True: Exit Function
         Else
            theExitNorth = True
            theExitEast = True
            theExitSouth = True
            theExitWest = True
            theExitUp = True
            theExitDown = True
         End If
      End If
   End If
Exit Function
errorhandler:
   errorModule = Err.Description & "(" & Err.Number & ") -> " & "MUME_Runtime handleRunMode"
   writeError (errorModule)
End Function

Public Function handleCollision(strData As String)
errorData = errorData & "handleCollision -> "
   handleCollision = False
   If LOST = True Then Exit Function
   tmpCheck = True
   
   If isYou2 Then
      If tmpCheck Then If checkStringCS(strData, "Alas, you cannot go that way...") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "doesn't want you riding") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "The descent is too steep, you need to climb to go there.") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "The ascent is too steep, you need to climb to go there.") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "Maybe you should get on your feet first?") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "In your dreams, or what?") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "Your mount refuses to follow your orders!") Then tmpCheck = False
   End If
   If isYou1 Then
      If tmpCheck Then If checkStringCS(strData, "No way! You are fighting for your life!") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "Oops! You cannot go there riding!") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "You can't go into deep water!") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "You failed swimming there.") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "You need to swim to go there.") Then tmpCheck = False
      If tmpCheck Then If checkStringCS(strData, "Nah... You feel too relaxed to do that..") Then tmpCheck = False
      If tmpCheck Then If checkString(strData, "You failed to climb there and fall down, hurting yourself.") Then tmpCheck = False
   End If
   If tmpCheck Then If checkStringCS(strData, " seems to be closed.") Then tmpCheck = False
   If tmpCheck Then If checkStringCS(strData, " too exhausted") Then tmpCheck = False

   If tmpCheck = False Then
      Call newCollision
      handleCollision = True: Exit Function
   End If
   If checkStringCS(strData, "It is pitch black...") Then
      Call SYNC_FALSE("room is dark!"): handleCollision = True: Exit Function
   End If
   If checkStringCS(strData, "You just see a dense fog around you...") Then
      Call SYNC_FALSE("room is covered in fog!"): handleCollision = True: Exit Function
   End If
End Function

Public Function handleDescription(strData As String)
errorData = errorData & "handleDescription -> "
   handleDescription = False
   noexitsfound = True
   If GetDescription = True Then
      If locatorCount >= locateRetry Then GetDescription = False: handleDescription = False: Exit Function
      locatorCount = locatorCount + 1
      n1 = InStr(strData, lookColour)
      If n1 > 0 Then
         n2 = InStr(n1 + 5, strData, colourEndCode)
         If n2 > 0 Then
            currentRoomName = Mid(strData, n1 + 5, n2 - (n1 + 5))
            a = n2 + 6
            b = InStr(a, strData, vbCrLf)
            If b > 0 Then
               c = InStrRev(strData, vbLf, b)
               If c > a Then
                  currentDesc = Mid(strData, a, c - a)
                  GetDescription = False
                  Call cmpWorldDesc(currentDesc)
                  handleDescription = True: Exit Function
               Else
                  GetDescription = True
                  Call informClient(strData, True)
                  Call SYNC_FALSE("cannot read description..")
                  handleDescription = True: Exit Function
               End If
            Else
               GetDescription = False
               Call SYNC_FALSE(" -= Unknown room =- ")
            End If
         Else
            GetDescription = True
            Call informClient(strData, True)
            Call SYNC_FALSE("cannot read roomname endcode..")
         End If
      Else
         GetDescription = True
         Call informClient(strData, True)
         Call SYNC_FALSE("cannot read roomname startcode..")
      End If
   End If
End Function

Public Function handleWhere(strData As String)
   handleWhere = False
   viewPlayers = False
   Dim mystring As String
   If checkStringCS(strData, "Players in your zone") Then
      mystring = strData
      Erase arrPlayers
      Dim s As String
      Dim p As String
      Dim isNew As Boolean
      arrPlayersIndex = LBound(arrPlayers, 1)
      a = 1
      Do While InStr(a, mystring, " - ")
         a = InStr(a, mystring, " - ")
         b = InStr(a, mystring, vbLf)
         
         s = Mid(mystring, a + 3, b - (a + 3) - 1): isNew = True
         n0 = InStrRev(mystring, vbLf, a)
         p = Left(Trim(Mid(mystring, n0 + 1, 18)), 2)

         For n = LBound(arrPlayers) To UBound(arrPlayers)
            If StrComp(arrPlayers(n), s, vbTextCompare) = 0 Then
               arrPlayersNames(n) = arrPlayersNames(n) & "," & p
               isNew = False
               Exit For
            Else
'               mystring = Replace(mystring, "  - " & s, "[" & CStr(arrPlayersIndex + 1) & "] " & s)
               isNew = True
            End If
         Next

         If isNew And arrPlayersIndex < UBound(arrPlayers, 1) Then
            arrPlayers(arrPlayersIndex) = s
            arrPlayersNames(arrPlayersIndex) = p
            arrPlayersIndex = arrPlayersIndex + 1
         End If
         a = b
      Loop
      
 '     informClient (mystring)
      viewPlayers = True
      Call DrawMap
      handleWhere = True
   End If
End Function

