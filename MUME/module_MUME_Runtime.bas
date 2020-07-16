Attribute VB_Name = "MUME_Runtime"
Option Explicit
Dim n As Integer
Dim n0 As Integer, n1 As Integer, n2 As Integer, n3 As Integer, n4 As Integer, n5 As Integer
Dim a As Integer, b As Integer, c As Integer
Public lookColour As String
Public roomdescriptionColour As String
Public fgColour As String
Public bgColour As String
Public fgBold As String
Public fgUnderline As String
Public tmpCheck As Boolean
Public noexitsfound As Boolean
Public leader As String
Public arrPlayers(10) As String
Public arrPlayersNames(10) As String
Public arrEnemies(10) As String
Public indexEnemies As Integer
Public arrPlayersIndex As Integer
Public viewPlayers As Boolean
Public viewTarget As Boolean
Public target As String
Public targetRow As Integer
Public targetCol As Integer
Public isYou1 As Boolean
Public isYou2 As Boolean
Public debug_length As Integer
Public MUD_output_length As Long
Public actual(1 To 1000, 1 To 2) As Long
Public potential(1 To 1000, 1 To 2) As Long
Public resultCase As Boolean

Public myStopper As cHiResTimer



Public Function handleRunMode(ByRef strData As String) As Boolean
If DEBUGMODE = False Then On Error GoTo errorhandler Else On Error GoTo 0
errorData = errorData & "handleRunMode -> "
handleRunMode = False
   
   n0 = InStrB(1, strData, lookColour, vbBinaryCompare)
   If n0 < 1 Then Exit Function
   n2 = InStrB(n0, strData, colourEndCode & vbCrLf, vbBinaryCompare)
   If n2 < 1 Then Exit Function

'FLEEING
   n3 = InStrB(1, MidB(strData, 1, n0), "You flee head over heels.", vbBinaryCompare)
'   If n1 > 0 Then
'      Erase actual
'      Erase potential
'   End If
   If LOST = False Then
      canUndo = False
      undoRow = virtualRow
      undoCol = virtualCol
'kontrollime kas toimus flee
      n1 = InStrB(1, strData, "You flee head over heels.", vbBinaryCompare)
      If n1 > 0 Then
         errorData = errorData & "fleeing -> "
'n1 = n0 + LenB(lookColour)  '5
         If n2 > 0 Then
            'leiame fleemise rea alguse
            
            n4 = InStrB(n3 + LenB("You flee head over heels." & vbCrLf), strData, "You flee ", vbBinaryCompare)
            
            If n4 > 0 Then
               n5 = InStrB(n4, strData, "." & vbCrLf, vbBinaryCompare)
               Dim direction As String
               direction = MidB(strData, n4 + LenB("You flee "), n5 - (n4 + LenB("You flee ")))
'YOU FLED
               If chkFleeMove(direction) Then
                  virtualRow = theROW
                  virtualCol = theCOL
                  'ruum on ikkagi veel võibolla tulemas ja
                  roomcount = roomcount + 1
                  stackOUT = stackOUT - 1
                  Call newCollision
               End If
               Call loadRoom(theROW, theCOL)
               Call DrawMap
               canUndo = True
               handleRunMode = True: Exit Function
               
            Else
               Call SYNC_FALSE("HandleRunMode, flee message without colour endcode!"):
               handleRunMode = True: Exit Function
            End If
         End If
      End If
   End If

'FOLLOW
   followMode = False
   If frmMap.mnuFollow.Checked Then
      errorData = errorData & "follow -> "
      If LOST = False Then
         If n2 > 0 Then
            currentRoomname = MidB(strData, n0 + LenB(lookColour), n2 - (n0 + LenB(lookColour)))
            If followMode = False Then If checkStringCS(strData, leader & " leaves north") Then followMode = True: Call travel("n", roomcount, N_MAP, virtualRow - 1, virtualCol, checkStringCS(strData, "You were not able to keep your concentration while moving."))
            If followMode = False Then If checkStringCS(strData, leader & " leaves east") Then followMode = True: Call travel("e", roomcount, E_MAP, virtualRow, virtualCol + 1, checkStringCS(strData, "You were not able to keep your concentration while moving."))
            If followMode = False Then If checkStringCS(strData, leader & " leaves south") Then followMode = True: Call travel("s", roomcount, S_MAP, virtualRow + 1, virtualCol, checkStringCS(strData, "You were not able to keep your concentration while moving."))
            If followMode = False Then If checkStringCS(strData, leader & " leaves west") Then followMode = True: Call travel("w", roomcount, W_MAP, virtualRow, virtualCol - 1, checkStringCS(strData, "You were not able to keep your concentration while moving."))
            If followMode = False Then If checkStringCS(strData, leader & " leaves up") Then followMode = True: Call travel("u", roomcount, U_MAP, virtualRow, virtualCol, checkStringCS(strData, "You were not able to keep your concentration while moving."))
            If followMode = False Then If checkStringCS(strData, leader & " leaves down") Then followMode = True: Call travel("d", roomcount, D_MAP, virtualRow, virtualCol, checkStringCS(strData, "You were not able to keep your concentration while moving."))
            If followMode = True Then
               'kui liider flees ruumist. ta ei saa fleeda backridimisega
               If checkStringCS(strData, leader & " panics, ") Then followMode = False
            Else
               'backriding leader follow
               If followMode = False Then If checkStringCS(strData, leader & " and ") Then If checkStringCS(strData, " leave north") Then followMode = True: Call travel("n", roomcount, N_MAP, virtualRow - 1, virtualCol, checkStringCS(strData, "You were not able to keep your concentration while moving."))
               If followMode = False Then If checkStringCS(strData, leader & " and ") Then If checkStringCS(strData, " leave east") Then followMode = True: Call travel("e", roomcount, E_MAP, virtualRow, virtualCol + 1, checkStringCS(strData, "You were not able to keep your concentration while moving."))
               If followMode = False Then If checkStringCS(strData, leader & " and ") Then If checkStringCS(strData, " leave south") Then followMode = True: Call travel("s", roomcount, S_MAP, virtualRow + 1, virtualCol, checkStringCS(strData, "You were not able to keep your concentration while moving."))
               If followMode = False Then If checkStringCS(strData, leader & " and ") Then If checkStringCS(strData, " leave west") Then followMode = True: Call travel("w", roomcount, W_MAP, virtualRow, virtualCol - 1, checkStringCS(strData, "You were not able to keep your concentration while moving."))
               If followMode = False Then If checkStringCS(strData, leader & " and ") Then If checkStringCS(strData, " leave up") Then followMode = True: Call travel("u", roomcount, U_MAP, virtualRow, virtualCol, checkStringCS(strData, "You were not able to keep your concentration while moving."))
               If followMode = False Then If checkStringCS(strData, leader & " and ") Then If checkStringCS(strData, " leave down") Then followMode = True: Call travel("d", roomcount, D_MAP, virtualRow, virtualCol, checkStringCS(strData, "You were not able to keep your concentration while moving."))
            End If
            If followMode = True Then
               If roomcount > 0 Then
                  Call newUpdateTheRoom:
                  handleRunMode = True: Exit Function
               End If
            End If
         End If
      End If
   End If

'WALKING
   If n2 > 0 Then
      errorData = errorData & "walking -> "
      currentRoomname = MidB(strData, n0 + LenB(lookColour), n2 - (n0 + LenB(lookColour))) 'was 5
      n3 = InStrB(n2, strData, "Exits: ", vbBinaryCompare)
      If n3 > 0 Then
'YOU MOVED
         noexitsfound = False
         n4 = InStrB(n3, strData, ".", vbBinaryCompare) 'was 7
         currentExits = MidB(strData, n3 + LenB("Exits: "), n4 - (n3 + LenB("Exits: ")))
         
''''''''''         If Autosync And (Not MappingMode) And LOST Then
''''''''''            tmpMatch = 0
''''''''''            If actual(1, 1) > 0 Then
''''''''''               'kõndisid järgmisse ruumi, tuleb leida võimalikud.. vastavalt suunale
''''''''''               Dim dirs As Long
''''''''''               Select Case lastMoveDirection
''''''''''               Case "n": dirs = N_MAP
''''''''''               Case "e": dirs = E_MAP
''''''''''               Case "s": dirs = S_MAP
''''''''''               Case "w": dirs = W_MAP
''''''''''               Case "u": dirs = U_MAP
''''''''''               Case "d": dirs = D_MAP
''''''''''               Case Else: dirs = (N_MAP Or E_MAP Or S_MAP Or W_MAP Or U_MAP Or D_MAP)
''''''''''               End Select
''''''''''
''''''''''               crawlRadius = 1 'ainult ühe ruumi kaugusele vaatan..
''''''''''               cursor = 0 ' cursor crawleri jaoks
''''''''''               Dim i As Integer
'''''''''''Dim j As Integer, oldcursor As Integer
'''''''''''Debug.Print "i=" & i & ", j=" & j & "," & "oldcursor = " & oldcursor
''''''''''               For i = 1 To UBound(actual, 1) 'ühest nendest ruumidest tulid sina
''''''''''                  'lisame võimaliku ruumi, vastavalt kõnnitud suunale potentsiaalsete ruumide hulka
''''''''''                  If actual(i, 1) = 0 Then Exit For
'''''''''''oldcursor = cursor
''''''''''                  Call Crawler(dirs, potential, 0, actual(i, 1), actual(i, 2))
''''''''''               Next
'''''''''''Debug.Print "potential count of exits = " & cursor
''''''''''
'''''''''''For j = 1 To cursor
'''''''''''   Debug.Print "potential => " & aData(getIndex(potential(j, 1), potential(j, 2)), cROOMNAME)
'''''''''''Next
''''''''''
''''''''''               If potential(1, 1) > 0 Then
''''''''''                  'vaatame nüüd läbi need ruumid, kus me hetkel peaksime olema
''''''''''                  For cursor = 1 To UBound(potential, 1)
''''''''''                     If potential(cursor, 1) = 0 Then Exit For
''''''''''                     If LenB(aData(getIndex(potential(cursor, 1), potential(cursor, 2)), cROOMNAME)) = LenB(currentRoomname) Then
''''''''''                        If aData(getIndex(potential(cursor, 1), potential(cursor, 2)), cROOMNAME) = currentRoomname Then ' matches
''''''''''                           tmpMatch = tmpMatch + 1
''''''''''                           tmpMatchIndex = cursor
''''''''''                        End If
''''''''''                     End If
''''''''''                  Next
''''''''''                  Select Case tmpMatch
''''''''''                  Case 0
''''''''''                     Call SYNC_FALSE("room not found!") ' arvatavasti mappimata ruum
''''''''''                  Case 1
''''''''''                     virtualRow = potential(tmpMatchIndex, 1)
''''''''''                     virtualCol = potential(tmpMatchIndex, 2)
''''''''''                     Call SYNC_TRUE
''''''''''                  Case Else
''''''''''                     'ruumides kus me NÜÜD oleme paneme actual massiivi
''''''''''                     Erase actual
''''''''''                     Erase potential
''''''''''                  End Select
''''''''''               End If
''''''''''            End If

''''''''''            'vaatame kus me võiksime olla
''''''''''            tmpMatch = 0
''''''''''            If actual(1, 1) = 0 Then
'''''''''''Debug.Print "-------------- new world lookup----------------"
''''''''''               For cursor = 1 To theCount
''''''''''                  If LenB(aData(cursor, cROOMNAME)) = LenB(currentRoomname) Then
''''''''''                     If aData(cursor, cROOMNAME) = currentRoomname Then ' matches
''''''''''                        tmpMatch = tmpMatch + 1
''''''''''                        ' üks ruumidest, kus sa olla võid
''''''''''                        If tmpMatch > 100 Then
''''''''''                           Erase actual
''''''''''                           Exit For
''''''''''                        End If
''''''''''                        actual(tmpMatch, 1) = CLng(aData(cursor, cROW))
''''''''''                        actual(tmpMatch, 2) = CLng(aData(cursor, cCOL))
'''''''''''Debug.Print "actual(" & actual(tmpMatch, 1) & "," & actual(tmpMatch, 2) & ")"
''''''''''                     End If
''''''''''                  End If
''''''''''               Next
'''''''''''Debug.Print "world matched count = " & tmpMatch
''''''''''               Select Case tmpMatch
''''''''''               Case 0
''''''''''                  Call SYNC_FALSE("room not found!")
''''''''''               Case 1
''''''''''                  virtualRow = actual(tmpMatch, 1)
''''''''''                  virtualCol = actual(tmpMatch, 2)
''''''''''                  Call SYNC_TRUE
''''''''''               End Select
''''''''''            End If
''''''''''        End If
                  If LOST = False Then
                     If roomcount > 0 Then
                        Call newUpdateTheRoom
                        handleRunMode = True: Exit Function
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
            'If theROW = 0 Or theCOL = 0 Then Stop
            'Call newUpdateTheRoom
            'handleRunMode = True: Exit Function
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
   errorModule = Err.description & "(" & Err.Number & ") -> " & "MUME_Runtime handleRunMode"
   writeError (errorModule)
End Function

Public Function isSameRoom(ByRef Roomname As String)
   'kontrollime kas on sama ruuminimi, kui ei ole siis oleme eksinud
   If LenB(aData(getIndex(theROW, theCOL), cROOMNAME)) <> LenB(currentRoomname) Then
      LOST = True
   Else
     If aData(getIndex(theROW, theCOL), cROOMNAME) <> currentRoomname Then
         LOST = True
      End If
   End If
End Function
Public Function handleCollision(ByRef strData As String) As Boolean
errorData = errorData & "handleCollision -> "
handleCollision = False
If LOST = True Then Exit Function
tmpCheck = True

If checkStringCS(strData, "You flee head over heels.") Then Exit Function

Dim i As Integer
For i = LBound(arrCollision) To UBound(arrCollision)
   If LenB(arrCollision(i)) = 0 Then Exit For
   If tmpCheck = False Then Exit For
   If checkStringCS(strData, arrCollision(i)) Then tmpCheck = False
Next
If tmpCheck Then
   If checkStringCS(strData, "It is pitch black...") Then
      LOST = True
      MappingMode = True
      MappingData = False
      theROW = virtualRow
      theCOL = virtualCol
      Call DrawMap
   End If
End If
If tmpCheck Then
   If checkStringCS(strData, "You just see a dense fog around you...") Then
      LOST = True
      MappingMode = True
      MappingData = False
       theROW = virtualRow
      theCOL = virtualCol
      Call DrawMap
   End If
End If
   
   If tmpCheck = False Then
      Call newCollision
      handleCollision = True: Exit Function
   End If
'   If checkStringCS(strData, "It is pitch black...") Then
'    If GODMODE Then
'    Else
'      Call SYNC_FALSE("room is dark!"): handleCollision = True: Exit Function
'    End If
'   End If
'   If checkStringCS(strData, "You just see a dense fog around you...") Then
'       If GODMODE Then
'       Else
'            Call SYNC_FALSE("room is covered in fog!"): handleCollision = True: Exit Function
'       End If
'   End If
End Function

Public Function handleWhere(ByRef strData As String) As Boolean
   handleWhere = False
   viewPlayers = False

   If checkStringCS(strData, "Players in your zone") Then
      Erase arrPlayers
      Dim s As String
      Dim p As String
      Dim isNew As Boolean
      arrPlayersIndex = LBound(arrPlayers, 1)
      a = 1
      Do While InStr(a, strData, " - ", vbBinaryCompare)
         a = InStr(a, strData, " - ", vbBinaryCompare)
         b = InStr(a, strData, vbLf, vbBinaryCompare)
         
         s = Mid(strData, a + 3, b - (a + 3) - 1): isNew = True
         n0 = InStrRev(strData, vbLf, a, vbBinaryCompare)
         p = Left(Trim(Mid(strData, n0 + 1, 18)), 2)

         For n = LBound(arrPlayers) To UBound(arrPlayers)
            If arrPlayers(n) = s Then
               arrPlayersNames(n) = arrPlayersNames(n) & "," & p
               isNew = False
               Exit For
            Else
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
      viewPlayers = True
      Call DrawMap
      handleWhere = True
   End If
End Function
'
'Public Function handleEnemy(s As String, start As Integer)
'   Dim a, b, c As Integer
'   Dim mystring As String
'   a = 0: c = 0: b = 0
'   b = InStr(start, s, "* leaves ", vbBinaryCompare)
'   If b > 0 Then a = InStrRev(s, "*", b - 1, vbBinaryCompare)
'   If a > 0 Then c = InStr(b, s, ".")
'   If c > 0 Then
'      mystring = Mid(s, a, c - a)
'      arrEnemies(indexEnemies) = mystring
'      indexEnemies = indexEnemies + 1
'      If indexEnemies > UBound(arrEnemies) - 1 Then indexEnemies = 0
'      Call handleEnemy(Mid(s, c + 1), 1)
'   End If
'End Function

Public Function showEnemy()
      Dim i As Integer
      Call informClient(vbCrLf & "================== E N E M I E S ==================", True)
      For i = indexEnemies To UBound(arrEnemies) - 1
         Call informClient(arrEnemies(i), True)
      Next
      For i = LBound(arrEnemies) To indexEnemies - 1
         Call informClient(arrEnemies(i), True)
      Next
      Call informClient("___________________________________________________" & vbCrLf, True)
End Function

Public Function handleDescription(ByRef strData As String) As Boolean
errorData = errorData & "handleDescription -> "
   handleDescription = False
   noexitsfound = True
   If GetDescription = True Then
      If locatorCount >= locateRetry Then GetDescription = False: handleDescription = False: Exit Function
      locatorCount = locatorCount + 1
      currentRoomname = getRoomname(strData)
      If LenB(currentRoomname) <> 0 Then
          currentDesc = getRoomDescription(strData) ' uus kogu ruumikirjeldus
          If LenB(currentDesc) <> 0 Then
             GetDescription = False
             Call getSynced(currentDesc, True)
             If LOST = True And tmpMatch = 0 Then ' uue crc32 ja kirjelduse järgi ei leitud
                Call cmpWorldDesc(currentDesc)
                If LOST = False Then ' vana kirjelduse järgi leiti ja therow, thecol väärtustati
                   'Call informClient("Updating description!", True)
                   aData(getIndex(theROW, theCOL), cDESCRIPTION) = encryptedDescription 'uus crc32 desc
                   Call updateThis(getIndex(theROW, theCOL))
                End If
             End If
             handleDescription = True: Exit Function
          Else
            GetDescription = True
            Call SYNC_FALSE("Cannot read description! Retrying.")
          End If
       Else
         GetDescription = True
         Call SYNC_FALSE("Cannot read roomname! Retrying.")
       End If
   End If
End Function

Public Function getRoomDescription(ByRef text As String) As String
   getRoomDescription = vbNullString
   a = InStrB(1, text, lookColour, vbBinaryCompare)
   If a > 0 Then
      b = a + LenB(lookColour)
      c = InStrB(b, text, colourEndCode & vbCrLf, vbBinaryCompare) ' roomname colour end
      If c > 0 Then
         a = c + LenB(colourEndCode & vbCrLf) ' the beginning of description
         ' find last description row
         b = (InStrRev(text, roomdescriptionColour, , vbBinaryCompare) * 2) - 1  'the last descrption row
         If b > 0 Then
            c = InStrB(b, text, colourEndCode & vbCrLf, vbBinaryCompare) ' description colour end
            If c > 0 Then
               getRoomDescription = MidB(text, a, c - a)
               getRoomDescription = Replace(getRoomDescription, roomdescriptionColour, "", , , vbBinaryCompare)
               getRoomDescription = Replace(getRoomDescription, colourEndCode, "", , , vbBinaryCompare)
               getRoomDescription = Replace(getRoomDescription, vbCrLf, vbLf & vbCr, , , vbBinaryCompare)
               'getRoomDescription = Replace(getRoomDescription, vbLf, "<LF>" & vbLf, , , vbBinaryCompare)
               'getRoomDescription = Replace(getRoomDescription, vbCr, "<CR>" & vbCr, , , vbBinaryCompare)
            End If
         End If
      End If
   End If
End Function

Public Function getRoomname(ByRef text As String) As String
   getRoomname = vbNullString
   a = InStrB(1, text, lookColour, vbBinaryCompare)
   If a > 0 Then
      b = a + LenB(lookColour)
      c = InStrB(b, text, colourEndCode & vbCrLf, vbBinaryCompare)
      If c > 0 Then getRoomname = MidB(text, b, c - b)
   End If
End Function


'------------------------
         'b = (InStrB(a, text, vbCrLf, vbBinaryCompare) + 1) / 2 'convert to charIndex
         'If b > 0 Then
         '   c = (InStrRev(text, vbLf, b, vbBinaryCompare) * 2) - 1 'convert to byteIndex
         '   If c > a Then
         '      getRoomDescription = MidB(text, a, c - a)
         '   End If
         'End If
'------------------------
'        Dim rows As Variant, word As String
'        rows = Split(text, vbCrLf, , vbBinaryCompare)
'        Dim i As Integer, lastline As Integer
'        For i = 3 To UBound(rows)
'            Select Case True
'            Case InStrB(1, ">", rows(i), vbBinaryCompare) > 0 'kui tegemist on promptiga
'                lastline = i - 1: Exit For
'            Case LenB(rows(i)) = 0 'kui tegemist on tühja reaga, compact=false korral
'                lastline = i - 1: Exit For
'            Case Mid(rows(i), 1, 3) = "An " 'kui tegemist on arvatava objektiga
'                lastline = i - 1: Exit For
'            Case Mid(rows(i), 1, 2) = "A " 'kui tegemist on arvatava objektiga
'                lastline = i - 1: Exit For
'            Case Mid(rows(i), 1, 4) = "The " 'kui tegemist on arvatava objektiga
'                lastline = i - 1: Exit For
'            Case Else
'                lastline = 0
'            End Select
'        Next i
'
'        If lastline = 0 Then
'            getRoomDescription = vbNullString
'        Else
'            getRoomDescription = ""
'            For i = 1 To lastline
'                getRoomDescription = getRoomDescription & rows(i)
'                If i < lastline Then getRoomDescription = getRoomDescription & vbLf & vbCr
'            Next i
'        End If
'------------------------

