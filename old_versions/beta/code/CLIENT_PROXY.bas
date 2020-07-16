Attribute VB_Name = "CLIENT_RUNTIME"
Public Function handleSpecial(ByRef strData As String)

   handleSpecial = False
   If Len(strData) = 1 Then
      frmMap.tcpClient.SendData strData
      handleSpecial = True: Exit Function
   End If

End Function

Public Function handleRuntimeCommand(ByRef strData As String)
On Error GoTo errorhandler

Dim theCommand
Dim a As Long
Dim b As String
Dim c As String
Dim n As Long
   handleRuntimeCommand = False
   a = Len(strData)
   b = Mid(strData, 1, 1)
   c = Mid(strData, 1, 5)
   If a = 6 And b = "_" Then
      Select Case c
      Case "_show"
         Call frmMap.mnuOnTop_Click
      Case "_sync"
         Call frmMap.mnuLocate_Click
      Case "_tool"
         Call frmMap.mnuMapper_Click
      Case "_map1"
         Call frmMap.mnuSmall_Click
      Case "_map2"
         Call frmMap.mnuNormal_Click
      Case "_map3"
         Call frmMap.mnuLarge_Click
      Case "_port"
         Call frmMap.mnuPortals_Click
      Case "_door"
         Call frmMap.mnuDoornamesHide_Click
      Case "_move"
         Call frmMap.mnuMovement_Click
      Case "_help"
         With frmMap
            .tcpServer.SendData vbLf & "Mapping commands:"
            .tcpServer.SendData vbLf & "       Read room data       "
            .tcpServer.SendData vbLf & "_get"
            .tcpServer.SendData vbLf & "       Update room data       "
            .tcpServer.SendData vbLf & "_update"
            .tcpServer.SendData vbLf & "       Read room data and update       "
            .tcpServer.SendData vbLf & "_map"
            .tcpServer.SendData vbLf & "       Creating exits       "
            .tcpServer.SendData vbLf & "_n"
            .tcpServer.SendData vbLf & "_e"
            .tcpServer.SendData vbLf & "_s"
            .tcpServer.SendData vbLf & "_w"
            .tcpServer.SendData vbLf & "_u"
            .tcpServer.SendData vbLf & "_d"
            .tcpServer.SendData vbLf & "       Setting terrain type       "
            .tcpServer.SendData vbLf & "_t [road|plain|forest|swamp|hill|mountain|water|special]"
            .tcpServer.SendData vbLf & "_sun"
            .tcpServer.SendData vbLf & "_ride"
            .tcpServer.SendData vbLf & "       Setting direction door       "
            .tcpServer.SendData vbLf & "_nd [doorname]"
            .tcpServer.SendData vbLf & "_ed [doorname]"
            .tcpServer.SendData vbLf & "_sd [doorname]"
            .tcpServer.SendData vbLf & "_wd [doorname]"
            .tcpServer.SendData vbLf & "_ud [doorname]"
            .tcpServer.SendData vbLf & "_dd [doorname]"
            .tcpServer.SendData vbLf & "       Creating direction portal       "
            .tcpServer.SendData vbLf & "_np [row],[column]"
            .tcpServer.SendData vbLf & "_ep [row],[column]"
            .tcpServer.SendData vbLf & "_sp [row],[column]"
            .tcpServer.SendData vbLf & "_wp [row],[column]"
            .tcpServer.SendData vbLf & "_up [row],[column]"
            .tcpServer.SendData vbLf & "_dp [row],[column]"
            .tcpServer.SendData vbLf & "       Move on map       "
            .tcpServer.SendData vbLf & "_go [row],[column]"
            .tcpServer.SendData vbLf & "_movenorth"
            .tcpServer.SendData vbLf & "_moveeast"
            .tcpServer.SendData vbLf & "_movesouth"
            .tcpServer.SendData vbLf & "_movewest" & vbLf
         End With
         handleRuntimeCommand = True: Exit Function
      Case "_loc1"
         fleeRadius = 0
         Call caseFleeHandler(currentRoomName, currentExits, 20, True)
         handleRuntimeCommand = True: Exit Function
      Case "_locc2"
         Call WorldLocate(currentRoomName, currentExits)
         handleRuntimeCommand = True: Exit Function
      End Select
   End If

   If Out_Of_Sync = True Then
      frmMap.tcpClient.SendData strData
      handleRuntimeCommand = True: Exit Function
   End If

   If b <> "_" Then
      theCommand = Split(strData, vbLf)
      For n = LBound(theCommand) To UBound(theCommand) - 1
         If Len(theCommand(n)) = 1 Then
            Select Case theCommand(n)
            Case "n"
               If checkTheMap(roomCount, N_MAP, virtualRow - 1, virtualCol, "n") Then
                  If roomCount < limit Then frmMap.tcpClient.SendData theCommand(n) & vbLf
               End If
            Case "e"
               If checkTheMap(roomCount, E_MAP, virtualRow, virtualCol + 1, "e") Then
                  If roomCount < limit Then frmMap.tcpClient.SendData theCommand(n) & vbLf
               End If
            Case "s"
               If checkTheMap(roomCount, S_MAP, virtualRow + 1, virtualCol, "s") Then
                  If roomCount < limit Then frmMap.tcpClient.SendData theCommand(n) & vbLf
               End If
            Case "w"
               If checkTheMap(roomCount, W_MAP, virtualRow, virtualCol - 1, "w") Then
                  If roomCount < limit Then frmMap.tcpClient.SendData theCommand(n) & vbLf
               End If
            Case "u"
               If checkTheMap(roomCount, U_MAP, virtualRow, virtualCol, "u") Then
                  If roomCount < limit Then frmMap.tcpClient.SendData theCommand(n) & vbLf
               End If
            Case "d"
               If checkTheMap(roomCount, D_MAP, virtualRow, virtualCol, "d") Then
                  If roomCount < limit Then frmMap.tcpClient.SendData theCommand(n) & vbLf
               End If
            Case Else
               frmMap.tcpClient.SendData theCommand(n) & vbLf
            End Select
         Else
            frmMap.tcpClient.SendData theCommand(n) & vbCrLf
         End If
      Next
   End If
   
Exit Function
errorhandler:
   errorData = "Client_Runtime"
   writeError (errorData)
End Function

