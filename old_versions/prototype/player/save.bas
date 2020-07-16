Attribute VB_Name = "save"
Public Sub saveWorld()
On Error GoTo ErrorHandler
   BestEST.status.ForeColor = &HC0FFC0
   BestEST.status = "Saving Arda! Please wait..."
   BestEST.Refresh
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
'               Debug.Print "Warp Set > " & arr(oldRow, oldCol)
            End If
'            Debug.Print row & "," & col & " |" & arr(row, col) & " | " & arrDesc(row, col)
         Else
            If warp = False Then
               oldRow = row
               oldCol = col
'               Debug.Print "Warp Engaged > " & "row=" & oldRow & "   col=" & oldCol
               warp = True
            End If
         End If
      Next
   Next
   
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
   Dim filename As String
   Dim AccessConnStr As String
   filename = "C:\mume\world.mdb"
   AccessConnStr = "Data Source=" & filename & ";Provider=Microsoft.Jet.OLEDB.4.0;Mode=ReadWrite;Persist Security Info=False"
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
   BestEST.status.ForeColor = &HC0FFC0
   BestEST.status = "Arda has been saved!"
Exit Sub
ErrorHandler:
   BestEST.status.ForeColor = &HFF&
   BestEST.status = "ARDA IS IN RUINS!"
End Sub

