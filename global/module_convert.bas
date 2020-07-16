Attribute VB_Name = "convert"
Option Explicit
Public md5     '  As New MD5DLL.Crypt
Public cast128 '  As New cast.cipher
Public md5Val As Variant

Public Function getData(index As Integer) As Long
   If index = 0 Then
      getData = 0
   Else
      getData = getLng(aData(index, cDATA))
   End If
End Function
Public Function getRoom(index As Integer) As String
   If index = 0 Then
      getRoom = 0
   Else
      getRoom = aData(index, cROOMNAME)
   End If
End Function
Public Function getInt(ByRef s As String) As Integer
   If s = vbNullString Then getInt = 0 Else getInt = CInt(s)
End Function
Public Function getLng(ByRef s As String) As Long
   If s = vbNullString Then getLng = 0 Else getLng = CLng(s)
End Function

Public Function getIndex(ByRef row As Integer, ByRef col As Integer, Optional dir As String) As Integer
   If Not isValid(row, col) Then Exit Function  'vigane
   
   getIndex = getInt(aWorld(row, col, theLEVEL))
   
   'If MappingMode Then staticLevel = True
   'If Not staticLevel Then 'otsime ka teistest maailmadest
      'If getIndex = 0 Then getIndex = getInt(aWorld(row, col, Abs(theLEVEL - 1)))
   'End If
   
   'kui pole vaja portalist otsida, siis exit
   If dir = vbNullString Then Exit Function

   'check portals
   Dim r As Integer
   Dim c As Integer
   Dim cell As Long
   Dim map As Long
   Dim map_r As Integer
   Dim map_c As Integer
   Dim r_offset As Integer
   Dim c_offset As Integer
   Select Case LCase(dir)
   Case "n"
      map = N_MAP: map_r = cNPORTALR: map_c = cNPORTALC: r_offset = -1: c_offset = 0
   Case "e"
      map = E_MAP: map_r = cEPORTALR: map_c = cEPORTALC: r_offset = 0: c_offset = 1
   Case "s"
      map = S_MAP: map_r = cSPORTALR: map_c = cSPORTALC: r_offset = 1: c_offset = 0
   Case "w"
      map = W_MAP: map_r = cWPORTALR: map_c = cWPORTALC: r_offset = 0: c_offset = -1
   Case "u"
      map = U_MAP: map_r = cUPORTALR: map_c = cUPORTALC: r_offset = 0: c_offset = 0
   Case "d"
      map = D_MAP: map_r = cDPORTALR: map_c = cDPORTALC: r_offset = 0: c_offset = 0
   Case Else
      map = 0: map_r = 0: map_c = 0
   End Select
   If (aData(getIndex, cDATA) And map) = 0 Then 'exit puudub
      getIndex = 0
      Exit Function
   End If

   r = aData(getIndex(row, col), map_r)
   c = aData(getIndex(row, col), map_c)
   If isValid(r, c) Then
      'portal
   Else
      'kõrvalolev ruum
      r = row + r_offset
      c = col + c_offset
      If Not isValid(r, c) Then Exit Function
   End If 'kas on tavaline exit
   
   getIndex = getIndex(r, c)

End Function


Public Function encryptDesc(ByVal value As String)
   Dim i As Integer
   i = 0
   encryptDesc = Null
   If LenB(value) <> 0 Then
      Do While IsNull(encryptDesc)
         i = i + 1
         If i > 3 Then Exit Do
         
         md5Val = md5.Encrypt(value)                  'hash description
         encryptDesc = cast128.b64StrEnc(md5Val)      'encryption + converted to base64
         If IsNull(encryptDesc) Then
            value = value & "#"  'couldnt not encrypt, adding # to differ the input
            Call informClient("Description hash failed! (plan B)", True)
         End If
      
      Loop
   Else
      Call informClient("Description is empty.", True)
   End If
End Function

Public Function CRC32(ByVal value As String) As String
   Dim cCRC32 As New cCRC32
   Dim i As Integer: Dim j As Integer

   If LenB(value) <> 0 Then
      ReDim valueBytes(1 To LenB(value)) As Byte
      For j = LBound(valueBytes) To UBound(valueBytes)
         valueBytes(j) = AscB(MidB(value, j, 1))
      Next
      CRC32 = Hex(cCRC32.GetByteArrayCrc32(valueBytes()))
      If LenB(CRC32) = 0 Then Call informClient("CRC32 failed!", True)
   Else
      Call informClient("Description is empty.", True)
   End If
End Function
Public Sub makeportals() ' only same level
   For cursor = 1 To theCount
      'kui on exit ja pole portal, siis..
      If (getData(cursor) And N_MAP) > 0 And (getData(cursor) And N_MAP) < N_portal Then
         aData(cursor, cNPORTALR) = aData(cursor, cROW) - 1: aData(cursor, cNPORTALC) = aData(cursor, cCOL): aData(cursor, cNLEVEL) = theLEVEL
      End If
      If (getData(cursor) And E_MAP) > 0 And (getData(cursor) And E_MAP) < E_portal Then
         aData(cursor, cEPORTALR) = aData(cursor, cROW): aData(cursor, cEPORTALC) = aData(cursor, cCOL) + 1: aData(cursor, cELEVEL) = theLEVEL
      End If
      If (getData(cursor) And S_MAP) > 0 And (getData(cursor) And S_MAP) < S_portal Then
         aData(cursor, cSPORTALR) = aData(cursor, cROW) + 1: aData(cursor, cSPORTALC) = aData(cursor, cCOL): aData(cursor, cSLEVEL) = theLEVEL
      End If
      If (getData(cursor) And W_MAP) > 0 And (getData(cursor) And W_MAP) < W_portal Then
         aData(cursor, cWPORTALR) = aData(cursor, cROW): aData(cursor, cWPORTALC) = aData(cursor, cCOL) - 1: aData(cursor, cWLEVEL) = theLEVEL
      End If
      Call updateThis(cursor)
   Next
End Sub

Public Sub makeroads()
   For cursor = 1 To theCount
      If checkRoad(getRoom(cursor)) = True Then
         aData(cursor, cDATA) = aData(cursor, cDATA) Or ISROAD
         Call updateThis(cursor)
      End If
   Next
End Sub

Public Sub makedungeons()
'todo
   For cursor = 1 To theCount
      If (getData(cursor) And TERRAIN_MAP) = underground Then
         aData(cursor, cLEVEL) = 1
      Else
         aData(cursor, cLEVEL) = 0
      End If
      Call updateThis(cursor)
   Next
End Sub


Public Sub createportalMAPPING(cursor As Integer, dir As String)
   'kontrollitakse, kas exit on olemas, aga portalit veel pole tehtud
'
   If (getData(cursor) And N_MAP) > 0 Then
      aData(cursor, cNPORTALR) = aData(cursor, cROW) - 1: aData(cursor, cNPORTALC) = aData(cursor, cCOL): aData(cursor, cNLEVEL) = theLEVEL
   End If
   If (getData(cursor) And E_MAP) > 0 Then
      aData(cursor, cEPORTALR) = aData(cursor, cROW): aData(cursor, cEPORTALC) = aData(cursor, cCOL) + 1: aData(cursor, cELEVEL) = theLEVEL
   End If
   If (getData(cursor) And S_MAP) > 0 Then
      aData(cursor, cSPORTALR) = aData(cursor, cROW) + 1: aData(cursor, cSPORTALC) = aData(cursor, cCOL): aData(cursor, cSLEVEL) = theLEVEL
   End If
   If (getData(cursor) And W_MAP) > 0 Then
      aData(cursor, cWPORTALR) = aData(cursor, cROW): aData(cursor, cWPORTALC) = aData(cursor, cCOL) - 1: aData(cursor, cWLEVEL) = theLEVEL
   End If

   'kui mapitakse, tuleb arvestada kust tasemelt tuldi
   'oldrow, oldcol, oldlevel pannakse paika form.dblcick või roomload korral
   Dim oldCursor As Integer
   oldCursor = getInt(aWorld(oldRow, oldCol, oldLevel))
   If (dir = "n") And (getData(cursor) And N_MAP) > 0 Then
      aData(cursor, cNPORTALR) = oldRow: aData(cursor, cNPORTALC) = oldCol: aData(cursor, cNLEVEL) = oldLevel
      aData(oldCursor, cSPORTALR) = theROW: aData(oldCursor, cSPORTALC) = theCOL: aData(oldCursor, cSLEVEL) = theLEVEL
   End If
   If (dir = "e") And (getData(cursor) And E_MAP) > 0 Then
      aData(cursor, cEPORTALR) = oldRow: aData(cursor, cEPORTALC) = oldCol: aData(cursor, cELEVEL) = oldLevel
      aData(oldCursor, cWPORTALR) = theROW: aData(oldCursor, cWPORTALC) = theCOL: aData(oldCursor, cWLEVEL) = theLEVEL
   End If
   If (dir = "s") And (getData(cursor) And S_MAP) > 0 Then
      aData(cursor, cSPORTALR) = oldRow: aData(cursor, cSPORTALC) = oldCol: aData(cursor, cSLEVEL) = oldLevel
      aData(oldCursor, cNPORTALR) = theROW: aData(oldCursor, cNPORTALC) = theCOL: aData(oldCursor, cNLEVEL) = theLEVEL
   End If
   If (dir = "w") And (getData(cursor) And W_MAP) > 0 Then
      aData(cursor, cWPORTALR) = oldRow: aData(cursor, cWPORTALC) = oldCol: aData(cursor, cWLEVEL) = oldLevel
      aData(oldCursor, cEPORTALR) = theROW: aData(oldCursor, cEPORTALC) = theCOL: aData(oldCursor, cELEVEL) = theLEVEL
   End If


'Debug.Print "oldE (" & aData(oldCursor, cEPORTALR) & "," & aData(oldCursor, cWPORTALC) & ")"

   If (dir = "u") And (getData(cursor) And U_MAP) > 0 Then
      aData(cursor, cUPORTALR) = oldRow: aData(cursor, cUPORTALC) = oldCol: aData(cursor, cULEVEL) = oldLevel
      aData(oldCursor, cDPORTALR) = theROW: aData(oldCursor, cDPORTALC) = theCOL: aData(oldCursor, cDLEVEL) = theLEVEL
   End If
   If (dir = "d") And (getData(cursor) And D_MAP) > 0 Then
      aData(cursor, cDPORTALR) = oldRow: aData(cursor, cDPORTALC) = oldCol: aData(cursor, cDLEVEL) = oldLevel
      aData(oldCursor, cUPORTALR) = theROW: aData(oldCursor, cUPORTALC) = theCOL: aData(oldCursor, cULEVEL) = theLEVEL
   End If

End Sub
