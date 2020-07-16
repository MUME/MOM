Attribute VB_Name = "MUME_Runtime"
Option Explicit
Dim n0 As Long, n1 As Long, n2 As Long, n3 As Long, n4 As Long
Dim a As Long, b As Long, c As Long

Public Function handleCollision(ByRef strData As String)
   
   handleCollision = False
   If Out_Of_Sync = False Then
      If AlasCount > 0 Then
         AlasCount = AlasCount - 1
         virtualRow = theRow
         virtualCol = theCol
      Else
         If checkString(strData, "Alas, you cannot go that way...") = True Or _
            checkString(strData, " seems to be closed.") = True Or _
            checkString(strData, "Your mount refuses to follow your orders!") = True Or _
            checkString(strData, "doesn't want you riding") = True Or _
            checkString(strData, "Oops! You cannot go there riding!") = True Or _
            checkString(strData, "You need to swim to go there.") = True Or _
            checkString(strData, "You need to climb to go there.") = True Or _
            checkString(strData, " too exhausted.") = True Then
               Call Collision
               handleCollision = True
               Exit Function
         End If
         If checkString(strData, "Maybe you should get on your feet first?") = True Or _
            checkString(strData, "Nah... You feel too relaxed to do that..") = True Or _
            checkString(strData, "In your dreams, or what?") = True Then
               Call resetBuffer
               handleCollision = True
               Exit Function
         End If
         If checkString(strData, "It is pitch black...") = True Then
            Call SYNC_FALSE
            Exit Function
         End If
      End If
   End If
   
End Function

Public Function handleRunMode(ByRef strData As String)

   handleRunMode = False
   n1 = InStr(strData, "[32")
   If n1 > 0 Then
      n2 = InStr(n1 + 5, strData, "[0m")
     If n2 > 0 Then
         n3 = InStr(n2 + 5, strData, "Exits:")
         If n3 > 0 Then
            currentRoomName = Mid(strData, n1 + 5, n2 - (n1 + 5))
            n4 = InStr(n3 + 6, strData, ".")
            currentExits = Mid(strData, n3 + 6, n4 - (n3 + 6)) '54
            currentString = Mid(strData, 1, n1)
            If Out_Of_Sync = False And checkString(currentString, "You flee head over heels.") = True Then
               fleeRadius = 0
               Call caseFleeHandler(currentRoomName, currentExits, 2, False)
               handleRunMode = True: Exit Function
            End If
            If Auto_Sync = True And Out_Of_Sync = True And MappingMode = False Then
               fleeRadius = 0
               Call caseFleeHandler(currentRoomName, currentExits, 5, False)
               handleRunMode = True: Exit Function
            End If
            If roomCount > 0 Then
               Call updateTheRoom
            End If
            handleRunMode = True: Exit Function
         Else
            If Out_Of_Sync = False Then
               n0 = InStr(1, strData, " leaves ")
               If n0 > 0 Then
                  If checkString(strData, " leaves north") = True Then
                     If checkTheMap(roomCount, N_MAP, virtualRow - 1, virtualCol, "n") = True Then Debug.Print "North"
                  End If
                  If checkString(strData, " leaves east") = True Then
                     If checkTheMap(roomCount, E_MAP, virtualRow, virtualCol + 1, "e") = True Then Debug.Print "East"
                  End If
                  If checkString(strData, " leaves south") = True Then
                     If checkTheMap(roomCount, S_MAP, virtualRow + 1, virtualCol, "s") = True Then Debug.Print "South"
                  End If
                  If checkString(strData, " leaves west") = True Then
                     If checkTheMap(roomCount, W_MAP, virtualRow, virtualCol - 1, "w") = True Then Debug.Print "West"
                  End If
                  If checkString(strData, " leaves up") = True Then
                     If checkTheMap(roomCount, U_MAP, virtualRow, virtualCol, "u") = True Then Debug.Print "Up"
                  End If
                  If checkString(strData, " leaves down") = True Then
                     If checkTheMap(roomCount, D_MAP, virtualRow, virtualCol, "d") = True Then Debug.Print "Down"
                  End If
                  If roomCount > 0 Then
                     Call updateTheRoom
                  End If
               End If
            End If
         End If
      End If
   End If

End Function

Public Function handleDescription(ByRef strData As String)

   handleDescription = False
   If Out_Of_Sync = True And GetDescription = True Then
      n1 = InStr(strData, "[32")
      If n1 > 0 Then
         n2 = InStr(n1 + 5, strData, "[0m")
         If n2 > 0 Then
            a = n2 + 6
            b = InStr(a, strData, vbCrLf)
            c = InStrRev(strData, vbLf, b)
            If c > a Then
               currentDesc = Mid(strData, a, c - a)
               GetDescription = False
               'Call Tools.Log("incoming", currentDesc)
               Call cmpFleeDesc(currentDesc)
            Else
               frmTools.status.Caption = "Retrying to catch description!"
               frmMap.tcpClient.SendData "EXAMINE" & vbLf
            End If
            handleDescription = True: Exit Function
         End If
      End If
   End If

End Function

