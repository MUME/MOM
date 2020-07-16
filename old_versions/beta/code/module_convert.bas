Attribute VB_Name = "convert"
Public md5 As New MD5DLL.Crypt
Public md5Val As String
Public Base64 As New midori.Base64

Public Function EncryptDesc(tekst As String)
   'checkparam = (objb64.Base64Encode(md5.Encrypt(Text1.Text)))
   md5Val = md5.Encrypt(tekst)
   EncryptDesc = Base64.Base64Encode(md5Val)
End Function

Public Function Encrypt2(tekst As String)
   Encrypt2 = Base64.Base64Encode(tekst)
End Function


Public Sub DBConvert(ByRef rowOffset, ByRef colOffset)
   rowOffset = CLng(rowOffset)
   colOffset = CLng(colOffset)
   Dim cn As ADODB.Connection
   Set cn = New ADODB.Connection
   Dim rstWorld As New Recordset
   Set rstWorld = New ADODB.Recordset
   rstWorld.LockType = adLockBatchOptimistic
   rstWorld.CursorType = adOpenStatic
   rstWorld.CursorLocation = adUseClient
   Dim FileName As String
   Dim AccessConnStr As String
   FileName = App.Path & "\world.mdb"
   AccessConnStr = "Data Source=" & FileName & ";Provider=Microsoft.Jet.OLEDB.4.0;Mode=ReadWrite;Persist Security Info=False"
   cn.Open AccessConnStr
   rstWorld.Open "world", cn
   Set rstWorld.ActiveConnection = Nothing
   Dim n As Long
   Dim strDesc As String
   strDesc = ""
   Do While Not rstWorld.EOF
      rstWorld("row") = rstWorld("row") + rowOffset
      rstWorld("col") = rstWorld("col") + colOffset
      tmpDesc = Split(rstWorld("arrdesc"), ";")
      For n = 2 To 18 Step 3
         If tmpDesc(n) > 0 Then
            Stop
            tmpDesc(n) = tmpDesc(n) + rowOffset
            tmpDesc(n + 1) = tmpDesc(n + 1) + colOffset
         End If
      Next
      rstWorld("arrdesc") = Join(tmpDesc, ";")
      rstWorld.MoveNext
   Loop
   rstWorld.ActiveConnection = cn
   rstWorld.UpdateBatch
   rstWorld.Close
   Set rstWorld = Nothing
End Sub

Public Sub updateDB()

   Dim cn As ADODB.Connection
   Set cn = New ADODB.Connection
   Dim rstWorld As New Recordset
   Set rstWorld = New ADODB.Recordset
   rstWorld.LockType = adLockBatchOptimistic
   rstWorld.CursorType = adOpenStatic
   rstWorld.CursorLocation = adUseClient
   Dim FileName As String
   Dim AccessConnStr As String
   FileName = App.Path & "\world.mdb"
   AccessConnStr = "Data Source=" & FileName & ";Provider=Microsoft.Jet.OLEDB.4.0;Mode=ReadWrite;Persist Security Info=False"
   cn.Open AccessConnStr
   rstWorld.Open "world", cn
   Set rstWorld.ActiveConnection = Nothing
   
   Dim newData As Long
   Dim oldData As Long
   Dim tmpN, tmpE, tmpS, tmpW, tmpU, tmpD
   Do While Not rstWorld.EOF
      
      oldData = CLng(rstWorld("arr"))

      tmpN = updateDATA(oldData, 96, 64, N_MAP, N_exit, N_door, N_portal)
      tmpE = updateDATA(oldData, 384, 256, E_MAP, E_exit, E_door, E_portal)
      tmpS = updateDATA(oldData, 1536, 1024, S_MAP, S_exit, S_door, S_portal)
      tmpW = updateDATA(oldData, 6144, 4096, W_MAP, W_exit, W_door, W_portal)
      tmpU = updateDATA(oldData, 24576, 16384, U_MAP, U_exit, U_door, U_portal)
      tmpD = updateDATA(oldData, 98304, 65536, D_MAP, D_exit, D_door, D_portal)

      newData = (31 And oldData) Or tmpN Or tmpE Or tmpS Or tmpW Or tmpU Or tmpD
      rstWorld("arr") = newData

      rstWorld.MoveNext
   
   Loop
   rstWorld.ActiveConnection = cn
   rstWorld.UpdateBatch
   rstWorld.Close
   Set rstWorld = Nothing

End Sub

Public Function updateDATA( _
   ByRef oldData, ByRef oldMap, ByRef oldDoor, _
   ByRef newMap, ByRef newExit, ByRef newDoor, ByRef newPortal)
  
  Dim newData As Long
  
  If (oldData And oldMap) > 0 Then                 'there is an exit
      If (oldData And oldMap) = oldMap Then        'special case
         newData = newPortal
      Else
         If (oldData And oldMap) = oldDoor Then   'door
            newData = newDoor
         Else
            newData = newExit       'simple exit
         End If
      End If
   Else
      newData = 0
   End If
  
  updateDATA = newData

End Function

