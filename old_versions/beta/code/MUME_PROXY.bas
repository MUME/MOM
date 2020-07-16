Attribute VB_Name = "MUME_Mapping"
Option Explicit
Dim n0 As Long, n1 As Long, n2 As Long, n3 As Long, n4 As Long
Dim a As Long, b As Long, c As Long

Public Function handleMapping(ByRef strData As String)
On Error GoTo errorhandler
   handleMapping = False
   If MappingData = True Then
      Select Case MappingCase
      Case 1
         n1 = InStr(strData, "[32")
         If n1 > 0 Then
            n2 = InStr(n1 + 5, strData, "[0m")
            If n2 > 0 Then
               Call zeroMap
               mapRoomName = Mid(strData, n1 + 5, n2 - (n1 + 5))
               a = n2 + 6
               b = InStr(a, strData, vbCrLf)
               c = InStrRev(strData, vbLf, b)
               If c > a Then
                  mapDescription = Mid(strData, a, c - a)
               Else
                  frmTools.status.Caption = "Retrying..."
                  frmMap.tcpClient.SendData "EXAMINE" & vbLf
                  handleMapping = True
                  Exit Function
               End If
               'Debug.Print mapRoomName
               'Debug.Print mapDescription
               'Debug.Print ">" & EncryptDesc(mapDescription) & "<"
               frmTools.Roomname = mapRoomName
               frmTools.Description = mapDescription
               MappingCase = MappingCase + 1
               frmMap.tcpClient.SendData "EXITS" & vbLf
            End If
         Else
            frmMap.tcpClient.SendData "EXAMINE" & vbLf
         End If
         handleMapping = True: Exit Function
      Case 2
         If InStr(1, strData, "North - ") > 0 Then
            mapExitNorth = True
            frmTools.nExit = 1
         End If
         If InStr(1, strData, "East  - ") > 0 Then
            mapExitEast = True
            frmTools.eExit = 1
         End If
         If InStr(1, strData, "South - ") > 0 Then
            mapExitSouth = True
            frmTools.sExit = 1
         End If
         If InStr(1, strData, "West  - ") > 0 Then
            mapExitWest = True
            frmTools.wExit = 1
         End If
         If InStr(1, strData, "Up    - ") > 0 Then
            mapExitUp = True
            frmTools.uExit = 1
         End If
         If InStr(1, strData, "Down  - ") > 0 Then
            mapExitDown = True
            frmTools.dExit = 1
         End If
         MappingData = False
         frmTools.status.Caption = "Success..."
         If MappingGetUpdate = True Then
            MappingGetUpdate = False
            dataFromMUME = True
            Call mapUpdate
         End If
         handleMapping = True: Exit Function
      End Select
   End If

Exit Function
errorhandler:
   errorData = "Mume_Mapping handleMapping"
   writeError (errorData)
End Function

