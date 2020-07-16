Attribute VB_Name = "save"
Option Explicit

Public Sub TOTALCHANGE()
   Dim cn As ADODB.Connection
   Dim rstWorld As New Recordset
   Set cn = New ADODB.Connection
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
   Dim broken
   Do While Not rstWorld.EOF
'      broken = Split(rstWorld("arrdesc"), ";")
'      broken(0) = ""
'      broken(19) = ""
'      rstWorld("arrdesc") = Join(broken, ";")
      rstWorld("roomname") = Trim(rstWorld("roomname"))
'      rstWorld("description") = broken(19)
      rstWorld.Update
      rstWorld.MoveNext
   Loop
   rstWorld.ActiveConnection = cn
   rstWorld.UpdateBatch
   rstWorld.Close
   Set rstWorld = Nothing
   
End Sub

Public Sub saveWorld()
On Error GoTo errorhandler
   
   frmTools.status.ForeColor = &HC0FFC0
   frmTools.status = "Saving Arda! Please wait..."
   Dim row As Long, col As Long
   Dim oldRow As Long, oldCol As Long
   Dim warp As Boolean
   oldRow = arrMinRow
   oldCol = arrMinCol
   warp = True
   For row = arrMinRow To arrMaxRow
      For col = arrMinCol To arrMaxCol
         If arr(row, col) > 0 Then
            If warp = True Then
               arr(oldRow, oldCol) = -(row * 1000 + col)
               warp = False
            End If
         Else
            If warp = False Then
               oldRow = row
               oldCol = col
               warp = True
            End If
         End If
      Next
   Next
   arr(oldRow, oldCol) = -(arrMaxRow * 1000 + arrMaxCol)
   
   Dim cn As ADODB.Connection
   Set cn = New ADODB.Connection
   Dim rstWorld As New Recordset
   Set rstWorld = New ADODB.Recordset
   rstWorld.LockType = adLockBatchOptimistic
   rstWorld.CursorType = adOpenStatic
   rstWorld.CursorLocation = adUseClient
   Dim rstPortal As New Recordset
   Set rstPortal = New ADODB.Recordset
   rstPortal.LockType = adLockBatchOptimistic
   rstPortal.CursorType = adOpenStatic
   rstPortal.CursorLocation = adUseClient
   Dim FileName As String
   Dim AccessConnStr As String
   FileName = App.Path & "\world.mdb"
   AccessConnStr = "Data Source=" & FileName & ";Provider=Microsoft.Jet.OLEDB.4.0;Mode=ReadWrite;Persist Security Info=False"
   cn.Open AccessConnStr
   cn.Execute ("DELETE FROM world")
   cn.Execute ("DELETE FROM portal")
   
   rstWorld.Open "world", cn
   Set rstWorld.ActiveConnection = Nothing
   rstPortal.Open "portal", cn
   Set rstPortal.ActiveConnection = Nothing
 
   For row = arrMinRow To arrMaxRow
      For col = arrMinCol To arrMaxCol
         tempData = arr(row, col)
         If tempData < 0 Then
            rstPortal.AddNew
            rstPortal("portal") = tempData
            rstPortal("row") = row
            rstPortal("col") = col
            rstPortal.Update
            rstPortal.MoveNext
'            Debug.Print row & "," & col & " |" & tempData & " | " & arrDesc(row, col)
         End If
         If tempData > 0 Then
            rstWorld.AddNew
            rstWorld("row") = row
            rstWorld("col") = col
            rstWorld("arr") = tempData
            rstWorld("arrDesc") = arrDesc(row, col)
            rstWorld("roomname") = arrRoomname(row, col)
            rstWorld("description") = arrDescription(row, col)
            rstWorld.Update
            rstWorld.MoveNext
'           Debug.Print row & "," & col & " |" & tempData & " | " & arrDesc(row, col)
         End If
      Next
   Next
   rstWorld.ActiveConnection = cn
   rstWorld.UpdateBatch
   rstWorld.Close
   Set rstWorld = Nothing
   rstPortal.ActiveConnection = cn
   rstPortal.UpdateBatch
   rstPortal.Close
   Set rstPortal = Nothing
   frmTools.status.ForeColor = &HC0FFC0
   frmTools.status = "Arda has been saved!"
   cn.Close
   Set cn = Nothing

Exit Sub
errorhandler:
   errorData = "save SaveWorld"
   writeError (errorData)
   frmTools.status.ForeColor = &HFF&
   frmTools.status = "ARDA IS IN RUINS!"
End Sub

