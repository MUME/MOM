Attribute VB_Name = "select"
Option Explicit
Public localStartRow As Long
Public localStartCol As Long
Public localEndRow As Long
Public localEndCol As Long
Public oldLevel As Integer
Public selLevel As Integer
Public selectionEndLevel As Integer
Public selectionStartRow As Long
Public selectionStartCol As Long
Public selectionEndRow As Long
Public selectionEndCol As Long
Public tmpX As Long
Public tmpY As Long
Public selectType As Integer
Public Const selectCopy = 1
Public Const selectCut = 2
Public Const selectPaste = 3
Public Const selectDelete = 4
Public flagRow As Long
Public flagCol As Long

Public Sub handleSelection(Optional pasteSpecial As Boolean)

staticLevel = True

'If DEBUGMODE = False Then On Error GoTo errorhandler
Dim row As Integer
Dim col As Integer
Dim rowStart As Integer
Dim rowEnd As Integer
Dim colStart As Integer
Dim colEnd As Integer
'---
Dim n As Integer
Dim c As Integer
'---
errorData = errorData & "handleSelection -> "

If selectType > 0 Then

   If selectType = selectCopy Then pasteSpecial = False
   rowStart = selectionStartRow
   rowEnd = selectionEndRow
   If selectionStartRow > selectionEndRow Then
      rowStart = selectionEndRow
      rowEnd = selectionStartRow
   End If
   colStart = selectionStartCol
   colEnd = selectionEndCol
   If selectionStartCol > selectionEndCol Then
      colStart = selectionEndCol
      colEnd = selectionStartCol
   End If
Dim destRow As Integer
Dim destCol As Integer
Dim destRoom As Long
frmMap.Caption = "Please wait."

'------------------------------------------------------------
' moving
'------------------------------------------------------------
For row = rowStart To rowEnd
   For col = colStart To colEnd
      If isValid(row, col) Then
         If getInt(aWorld(row, col, selLevel)) > 0 Then 'room exists
            If selectType = selectCopy Or selectType = selectCut Then
            
            
            
            
               destRow = theROW + (row - rowStart)
               destCol = theCOL + (col - colStart)
               If getInt(aWorld(destRow, destCol, theLEVEL)) = 0 Then 'the new LOCATION must be created for pasted room
                  If canIncreaseTheCount Then ' also informs client
                     theCount = theCount + 1
                  Else
                     Call DrawMap: Exit Sub
                  End If
                  aWorld(destRow, destCol, theLEVEL) = theCount ' set the new source index
               End If
               ' WHERE:
               ' SOURCE ROOM      => getIndex(row, col)
               ' DESTINATION ROOM => getIndex(destRow, destCol)
               'destRoom = getIndex(destRow, destCol)
               aData(aWorld(destRow, destCol, theLEVEL), cROW) = destRow
               aData(aWorld(destRow, destCol, theLEVEL), cCOL) = destCol
               aData(aWorld(destRow, destCol, theLEVEL), cDATA) = aData(aWorld(row, col, selLevel), cDATA)
               aData(aWorld(destRow, destCol, theLEVEL), cROOMNAME) = aData(aWorld(row, col, selLevel), cROOMNAME)
               aData(aWorld(destRow, destCol, theLEVEL), cDESCRIPTION) = aData(aWorld(row, col, selLevel), cDESCRIPTION)
               aData(aWorld(destRow, destCol, theLEVEL), cNDOOR) = aData(aWorld(row, col, selLevel), cNDOOR)
               aData(aWorld(destRow, destCol, theLEVEL), cEDOOR) = aData(aWorld(row, col, selLevel), cEDOOR)
               aData(aWorld(destRow, destCol, theLEVEL), cSDOOR) = aData(aWorld(row, col, selLevel), cSDOOR)
               aData(aWorld(destRow, destCol, theLEVEL), cWDOOR) = aData(aWorld(row, col, selLevel), cWDOOR)
               aData(aWorld(destRow, destCol, theLEVEL), cUDOOR) = aData(aWorld(row, col, selLevel), cUDOOR)
               aData(aWorld(destRow, destCol, theLEVEL), cDDOOR) = aData(aWorld(row, col, selLevel), cDDOOR)
               aData(aWorld(destRow, destCol, theLEVEL), cNOTE) = aData(aWorld(row, col, selLevel), cNOTE)
               aData(aWorld(destRow, destCol, theLEVEL), cNPORTALR) = aData(aWorld(row, col, selLevel), cNPORTALR)
               aData(aWorld(destRow, destCol, theLEVEL), cNPORTALC) = aData(aWorld(row, col, selLevel), cNPORTALC)
               aData(aWorld(destRow, destCol, theLEVEL), cEPORTALR) = aData(aWorld(row, col, selLevel), cEPORTALR)
               aData(aWorld(destRow, destCol, theLEVEL), cEPORTALC) = aData(aWorld(row, col, selLevel), cEPORTALC)
               aData(aWorld(destRow, destCol, theLEVEL), cSPORTALR) = aData(aWorld(row, col, selLevel), cSPORTALR)
               aData(aWorld(destRow, destCol, theLEVEL), cSPORTALC) = aData(aWorld(row, col, selLevel), cSPORTALC)
               aData(aWorld(destRow, destCol, theLEVEL), cWPORTALR) = aData(aWorld(row, col, selLevel), cWPORTALR)
               aData(aWorld(destRow, destCol, theLEVEL), cWPORTALC) = aData(aWorld(row, col, selLevel), cWPORTALC)
               aData(aWorld(destRow, destCol, theLEVEL), cUPORTALR) = aData(aWorld(row, col, selLevel), cUPORTALR)
               aData(aWorld(destRow, destCol, theLEVEL), cUPORTALC) = aData(aWorld(row, col, selLevel), cUPORTALC)
               aData(aWorld(destRow, destCol, theLEVEL), cDPORTALR) = aData(aWorld(row, col, selLevel), cDPORTALR)
               aData(aWorld(destRow, destCol, theLEVEL), cDPORTALC) = aData(aWorld(row, col, selLevel), cDPORTALC)
               
               aData(aWorld(destRow, destCol, theLEVEL), cLEVEL) = theLEVEL
               aData(aWorld(destRow, destCol, theLEVEL), cNLEVEL) = aData(aWorld(row, col, selLevel), cNLEVEL)
               aData(aWorld(destRow, destCol, theLEVEL), cELEVEL) = aData(aWorld(row, col, selLevel), cELEVEL)
               aData(aWorld(destRow, destCol, theLEVEL), cSLEVEL) = aData(aWorld(row, col, selLevel), cSLEVEL)
               aData(aWorld(destRow, destCol, theLEVEL), cWLEVEL) = aData(aWorld(row, col, selLevel), cWLEVEL)
               aData(aWorld(destRow, destCol, theLEVEL), cULEVEL) = aData(aWorld(row, col, selLevel), cULEVEL)
               aData(aWorld(destRow, destCol, theLEVEL), cDLEVEL) = aData(aWorld(row, col, selLevel), cDLEVEL)
               'määrame ruumi tasemele
               aWorld(destRow, destCol, theLEVEL) = theCount
               
               If pasteSpecial Then
                  '#################################
                  'all portals coming to source room, their portal coordinates will be updated
                  '#################################
                  For c = 1 To theCount
                     For n = cNPORTALR To cDPORTALR Step 2
                        If LenB(aData(c, n)) > 2 Then
                           If LenB(aData(c, n + 1)) > 2 Then
                              If aData(c, n) = row Then
                                 If aData(c, n + 1) = col Then
                                    aData(c, n) = destRow
                                    aData(c, n + 1) = destCol
                                    Call updateThis(c)
                                 End If
                              End If
                           End If
                        End If
                     Next
                  Next
               End If ' pastespecial check
               
            End If ' copy/paste check
            If (selectType = selectCut Or selectType = selectDelete) Then Call clearArraySlot(row, col)
            If LenB(frmMap.Caption) > 42 Then frmMap.Caption = "Please wait" Else frmMap.Caption = frmMap.Caption & "."
         End If ' getIndex(row, col) check
      End If ' valid coordinates check
   Next
Next




'------------------------------------------------------------
' update
'------------------------------------------------------------
frmMap.Caption = frmMap.Caption & "."
For row = theROW To theROW + (rowEnd - rowStart)
   For col = theCOL To theCOL + (colEnd - colStart)
      If isValid(row, col) = True Then
         If selectType = selectCopy Or selectType = selectCut Then
            If getInt(aWorld(row, col, theLEVEL)) > 0 Then
               Call updateThis(getInt(aWorld(row, col, theLEVEL)))
            End If
         End If
      End If
   Next
Next

frmMap.Caption = frmMap.Caption & "."
'------------------------------------------------------------
' stack reorganizing - aWorld(aData(index, cROW), aData(index, cCOL), theLEVEL) = freeSlot BUGINE
'------------------------------------------------------------
'If False And (selectType = selectCut Or selectType = selectDelete) Then
'   Dim index As Integer
'   Dim freeSlot As Integer
'   Dim fld As Integer
'   Dim w As Integer
'   freeSlot = 0
'   For index = 1 To UBound(aData)
'      If LenB(aData(index, cDATA)) = 0 Then
'         freeSlot = index
'         Exit For
'      End If
'   Next
'   If freeSlot = 0 Then
'      frmMap.Caption = "Arda is full, - sail West?"
'   Else
'      theCount = freeSlot - 1
'      For index = freeSlot + 1 To UBound(aData)
'         If LenB(aData(index, cDATA)) <> 0 Then
'
'
'            aWorld(aData(index, cROW), aData(index, cCOL), theLEVEL) = freeSlot
'            For fld = LBound(aData, 2) To UBound(aData, 2)
'               aData(freeSlot, fld) = aData(index, fld)
'            Next
'
'            freeSlot = freeSlot + 1
'
'
'         End If
'      Next
'
'      'free slots
'      For index = freeSlot To UBound(aData)
'         For fld = LBound(aData, 2) To UBound(aData, 2)
'            aData(index, fld) = vbNullString
'         Next
'      Next
'      theCount = freeSlot - 1
'      For w = theCount + 1 To UBound(aData)
'         aData(w, cDATA) = vbNullString
'      Next
'
'
'   End If
'End If

Call DrawMap
frmMap.Caption = mapTitle
   
End If

staticLevel = False

Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "select"
   writeError (errorModule)
End Sub

Public Sub copySelect()
   selectType = selectCopy
End Sub

Public Sub cutselect()
   selectType = selectCut
End Sub

Public Sub selectStart(X, Y)
errorData = errorData & "selectStart -> "
   selLevel = theLEVEL
   localStartRow = Round((Y - frmMap.ScaleHeight / 2) / roomsize, 0)
   If Abs(localStartRow) > mapRadius Then localStartCol = 0: localStartRow = 0: Exit Sub
   localStartCol = Round((X - frmMap.ScaleWidth / 2) / roomsize, 0)
   If Abs(localStartCol) > mapRadius Then localStartCol = 0: localStartRow = 0: Exit Sub
   tmpX = X
   tmpY = Y
End Sub
   
Public Sub selectEnd(X, Y)
errorData = errorData & "selectEnd -> "
   localEndRow = Round((Y - frmMap.ScaleHeight / 2) / roomsize, 0)
   If Abs(localEndRow) > mapRadius Then localEndCol = 0: localEndRow = 0: Exit Sub
   localEndCol = Round((X - frmMap.ScaleWidth / 2) / roomsize, 0)
   If Abs(localEndCol) > mapRadius Then localEndCol = 0: localEndRow = 0: Exit Sub
   If localStartRow <> localEndRow Or localStartCol <> localEndCol Then
      If MappingMode Then
         selectionEndLevel = theLEVEL
         selectionStartRow = theROW + localStartRow
         selectionStartCol = theCOL + localStartCol
         selectionEndRow = theROW + localEndRow
         selectionEndCol = theCOL + localEndCol
         Call DrawMap
         frmMap.Line (tmpX, tmpY)-(X, Y), QBColor(15), B
         frmMap.Line (tmpX + 1, tmpY + 1)-(X - 1, Y - 1), QBColor(15), B
         frmTools.start.text = selectionStartRow & "," & selectionStartCol
         frmTools.finish.text = selectionEndRow & "," & selectionEndCol
         frmTools.aFrame.Caption = frmTools.start.text
         frmTools.bFrame.Caption = frmTools.finish.text
         Call myText(frmMap, "[" & selectionStartRow & "," & selectionStartCol & "]", tmpX, tmpY, True, 8)
         Call myText(frmMap, "[" & selectionEndRow & "," & selectionEndCol & "]", X, Y, True, 8)
         Call myText(frmMap, "(" & Abs(selectionStartRow - selectionEndRow) + 1 & "x" & Abs(selectionStartCol - selectionEndCol) + 1 & ")", X, Y + 14, True, 8)
         selectType = 0
      Else
         flagRow = theROW + localStartRow
         flagCol = theCOL + localStartCol
         Call DrawMap
      End If
   End If
End Sub


Public Sub setLevel(newLevel As Integer)
Dim row As Integer
Dim col As Integer
Dim rowStart As Integer
Dim rowEnd As Integer
Dim colStart As Integer
Dim colEnd As Integer

'set area
rowStart = selectionStartRow
rowEnd = selectionEndRow
If selectionStartRow > selectionEndRow Then
   rowStart = selectionEndRow
   rowEnd = selectionStartRow
End If
colStart = selectionStartCol
colEnd = selectionEndCol
If selectionStartCol > selectionEndCol Then
   colStart = selectionEndCol
   colEnd = selectionStartCol
End If

'update
For row = rowStart To rowEnd
   For col = colStart To colEnd
      If isValid(row, col) Then
         If getInt(aWorld(row, col, theLEVEL)) > 0 Then 'room exists
            If getInt(aWorld(row, col, newLevel)) = 0 Then 'is free
               'set room to new world
               aWorld(row, col, newLevel) = aWorld(row, col, theLEVEL)
               'clear room from old world
               aWorld(row, col, theLEVEL) = 0
               'update room data
               aData(getInt(aWorld(row, col, newLevel)), cLEVEL) = newLevel
               'update save data
               Call updateThis(getInt(aWorld(row, col, newLevel)))
            End If
         End If
      End If
   Next
Next

'draw
Call DrawMap

End Sub
