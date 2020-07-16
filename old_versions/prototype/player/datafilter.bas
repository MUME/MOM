Attribute VB_Name = "datafilter"
Option Explicit
Public theCommand

Public Function dataFilter(ByRef data As String)
   dataFilter = False
   theCommand = Split(data, vbCrLf)
   For n = LBound(theCommand) To UBound(theCommand)
      If Len(theCommand(n)) = 1 Then
         Select Case theCommand(n)
         Case "n"
            If checkTheMap(N_map, N_exit, virtualRow - 1, virtualCol, "n") = True Then
               If roomCount < 3 Then dataFilter = True
               Debug.Print "N - YES"
            Else
               Debug.Print "N - NO"
            End If
            Exit Function
         Case "e"
            If checkTheMap(E_map, E_exit, virtualRow, virtualCol + 1, "e") = True Then
               If roomCount < 3 Then dataFilter = True
               Debug.Print "E - YES"
            Else
               Debug.Print "E - NO"
            End If
            Exit Function
         Case "s"
            If checkTheMap(S_map, S_exit, virtualRow + 1, virtualCol, "s") = True Then
               If roomCount < 3 Then dataFilter = True
               Debug.Print "S - YES"
            Else
               Debug.Print "S - NO"
            End If
            Exit Function
         Case "w"
            If checkTheMap(W_map, W_exit, virtualRow, virtualCol - 1, "w") = True Then
               If roomCount < 3 Then dataFilter = True
            Else
               Debug.Print "W - NO"
            End If
            Exit Function
         Case "u"
            If checkTheMap(U_map, U_exit, virtualRow, virtualCol, "u") = True Then
               If roomCount < 3 Then dataFilter = True
            Else
            End If
            Exit Function
         Case "d"
            If checkTheMap(D_map, D_exit, virtualRow, virtualCol, "d") = True Then
               If roomCount < 3 Then dataFilter = True
            Else
            End If
            Exit Function
         Case Else
            dataFilter = True
         End Select
         If debug_mode = True Then Debug.Print ">>>>> " & roomCount & " <<<<<"
      Else
         dataFilter = True
      End If
   Next
End Function
