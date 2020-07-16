VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmMap 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   Caption         =   "TheBestest"
   ClientHeight    =   3210
   ClientLeft      =   6360
   ClientTop       =   630
   ClientWidth     =   3480
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6
      Charset         =   186
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   232
   ScaleMode       =   0  'User
   ScaleWidth      =   232
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock tcpPlayer 
      Left            =   5160
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpMUD 
      Left            =   4800
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuTest 
      Caption         =   "test"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuLocate 
      Caption         =   "[Run]"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuSave 
         Caption         =   "Save map"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuAddNote 
         Caption         =   "Add note"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut selection"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy selection"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPasteSpecial 
         Caption         =   "Paste everything"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste only rooms"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete selection"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDeleteRoom 
         Caption         =   "Delete room"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSetLevel0 
         Caption         =   "Set to Earth"
         Enabled         =   0   'False
         Shortcut        =   ^{F1}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSetLevel1 
         Caption         =   "Set to Dungeon"
         Enabled         =   0   'False
         Shortcut        =   ^{F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSetLevel2 
         Caption         =   "Set to Hell"
         Enabled         =   0   'False
         Shortcut        =   ^{F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWalkthrough 
         Caption         =   "Update mode"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuWorld 
         Caption         =   "World window"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuShowLevel0 
         Caption         =   "Earth"
         Enabled         =   0   'False
         Shortcut        =   {F1}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShowLevel1 
         Caption         =   "Dungeon"
         Enabled         =   0   'False
         Shortcut        =   {F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShowLevel2 
         Caption         =   "Hell"
         Enabled         =   0   'False
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMovement 
         Caption         =   "Movement"
      End
      Begin VB.Menu mnuNotes 
         Caption         =   "Notes"
      End
      Begin VB.Menu mnuDoornames 
         Caption         =   "Doornames"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEnemies 
         Caption         =   "History"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGroup 
         Caption         =   "Group window"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMapDescription 
         Caption         =   "Mapping description"
      End
      Begin VB.Menu mnuFeedback 
         Caption         =   "Map feedback"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPlayers 
         Caption         =   "Players on map"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPortals 
         Caption         =   "Portals"
      End
      Begin VB.Menu mnuGrid 
         Caption         =   "Grid"
      End
      Begin VB.Menu mnuGridXY 
         Caption         =   "GridXY"
      End
      Begin VB.Menu mnuFancy 
         Caption         =   "Fancy"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "Version"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Set"
      Begin VB.Menu mnuHere 
         Caption         =   "I am here (here)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuWalk 
         Caption         =   "Blindwalking (walk)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuMap1 
         Caption         =   "Small map (map1)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMap2 
         Caption         =   "Medium map  (map2)"
      End
      Begin VB.Menu mnuMap3 
         Caption         =   "Large map (map3)"
      End
      Begin VB.Menu mnuAutosync 
         Caption         =   "Autosync"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRoomsync 
         Caption         =   "Roomsync"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTarget 
         Caption         =   "Target"
      End
      Begin VB.Menu mnuFollow 
         Caption         =   "Leader (lead)"
      End
      Begin VB.Menu mnuRelocate 
         Caption         =   "Locate retry"
      End
      Begin VB.Menu mnuDescriptionColour 
         Caption         =   "Description background colour(syncing)"
      End
      Begin VB.Menu mnuRoomFgColour 
         Caption         =   "Roomname foreground colour(syncing)"
         Begin VB.Menu mnuLookBold 
            Caption         =   "Bold"
         End
         Begin VB.Menu mnuLookUnderline 
            Caption         =   "Underline"
         End
         Begin VB.Menu mnuLookBetween 
            Caption         =   "------------"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuLookBLACK 
            Caption         =   "BLACK"
         End
         Begin VB.Menu mnuLookRED 
            Caption         =   "RED"
         End
         Begin VB.Menu mnuLookGREEN 
            Caption         =   "GREEN"
         End
         Begin VB.Menu mnuLookYELLOW 
            Caption         =   "YELLOW"
         End
         Begin VB.Menu mnuLookBLUE 
            Caption         =   "BLUE"
         End
         Begin VB.Menu mnuLookMAGENTA 
            Caption         =   "MAGENTA"
         End
         Begin VB.Menu mnuLookCYAN 
            Caption         =   "CYAN"
         End
         Begin VB.Menu mnuLookWHITE 
            Caption         =   "WHITE"
         End
      End
      Begin VB.Menu mnuRoomBgColour 
         Caption         =   "Roomname background colour(syncing)"
         Begin VB.Menu mnuLookBgNONE 
            Caption         =   "none"
         End
         Begin VB.Menu mnuLookBgBLACK 
            Caption         =   "BLACK"
         End
         Begin VB.Menu mnuLookBgRED 
            Caption         =   "RED"
         End
         Begin VB.Menu mnuLookBgGREEN 
            Caption         =   "GREEN"
         End
         Begin VB.Menu mnuLookBgYELLOW 
            Caption         =   "YELLOW"
         End
         Begin VB.Menu mnuLookBgBLUE 
            Caption         =   "BLUE"
         End
         Begin VB.Menu mnuLookBgMAGENTA 
            Caption         =   "MAGENTA"
         End
         Begin VB.Menu mnuLookBgCYAN 
            Caption         =   "CYAN"
         End
         Begin VB.Menu mnuLookBgWHITE 
            Caption         =   "WHITE"
         End
      End
      Begin VB.Menu mnuTellcolour 
         Caption         =   "Change colour tell"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuTellBLACK 
            Caption         =   "BLACK"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTellRED 
            Caption         =   "RED"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTellGREEN 
            Caption         =   "GREEN"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTellYELLOW 
            Caption         =   "YELLOW"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTellBLUE 
            Caption         =   "BLUE"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTellMAGENTA 
            Caption         =   "MAGENTA"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTellCYAN 
            Caption         =   "CYAN"
            Checked         =   -1  'True
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTellWHITE 
            Caption         =   "WHITE"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuSpam 
         Caption         =   "SPAM MODE ON"
      End
      Begin VB.Menu mnuBrief 
         Caption         =   "BRIEF MODE OFF"
      End
      Begin VB.Menu mnuClient 
         Caption         =   "Client name"
      End
      Begin VB.Menu mnuAlwaysOnTop 
         Caption         =   "Always on top"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuReceiver 
         Caption         =   "Enable receiver"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInformer 
         Caption         =   "Enable informer"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "[Tools]"
   End
   Begin VB.Menu mnuMap 
      Caption         =   "[Map]"
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mouseX As Single
Public mouseY As Single

Public MUD_output As String
Public oldMUDOutput As String
Public buttonPressed As Integer

'Public WithEvents objInformer As Informer.cAsync
'Public WithEvents objReceiver As Informer.cWhere
'Public objInformer As Informer.cAsync
'Public objReceiver As Informer.cWhere

'Private Sub objReceiver_Complete()
'   'handleInformerWhere (objReceiver.Result)
'End Sub
'Private Sub objReceiver_Cancelled()
'   informClient "cancelled"
'End Sub

Private Sub mnuDescriptionColour_Click()
   MsgBox "Default roomdescription colour is 'white'" & vbCrLf & "In MUME type 'change colour roomdescription white'." & vbCrLf & "MOM config file 'MOM.ini' roomdescription is also 'white'." & vbCrLf & "To manually change MOM roomdescription colour, please read MOM manual." & vbCrLf & "NB! MUME and MOM roomdescription colours must match!", vbOKOnly, "Description colour"
End Sub

Private Sub mnuFancy_Click()
   If frmMap.mnuFancy.Checked = False Then
      frmMap.mnuFancy.Checked = True
   Else
      frmMap.mnuFancy.Checked = False
   End If
   Call DrawMap
End Sub

Private Sub mnuInformer_Click()
   If mnuInformer.Checked Then
      mnuInformer.Checked = False
   Else
      mnuInformer.Checked = True
   End If
End Sub

Public Sub mnuMapDescription_Click()
   If mnuMapDescription.Checked Then
      mnuMapDescription.Checked = False
   Else
      mnuMapDescription.Checked = True
   End If
   If WorldLoaded Then Call saveMOMini
End Sub

Private Sub mnuReceiver_Click()
   If mnuReceiver.Checked Then
      mnuReceiver.Checked = False
   Else
      mnuReceiver.Checked = True
   End If
End Sub

Public Sub mnuTest_Click()
'   Dim i As Integer
   
'   timer.StartTimer
'   For i = 0 To 10000
'        'If checkStringCS(MUD_output, "- breath of briskness") Then
'           MUD_output = Replace(MUD_output, "- breath of briskness", "- breath of briskness" & i, 1, 1, vbBinaryCompare)
'        'End If
'   Next
'   timer.StopTimer
'   Call informClient("ELAPSED: " & timer.ElapsedTime, True)
 
         For cursor = 1 To theCount
            aData(cursor, cDATA) = ((125829116 And aData(cursor, cDATA)) Or 3)
            aData(cursor, cNOTE) = vbNullString
         Next
'(((2147483647 AND NOT(1082130432)) AND NOT(939524096)) AND NOT(3))
'   timer.StartTimer
'   For i = 0 To 100
'      Call DrawMap
'   Next
'   timer.StopTimer
'
'   Call informClient("ELAPSED: " & timer.ElapsedTime, True)
   
End Sub

Public Sub tcpPlayer_DataArrival(ByVal bytesTotal As Long)
errorData = "tcpPlayer_DataArrival -> "
If DEBUGMODE = False Then On Error GoTo errorhandler
'Dim MUD_output As String

MUD_output = vbNullString
frmMap.tcpPlayer.getData MUD_output    'get data from MUME
MUD_output_length = LenB(MUD_output)

If handleDescription(MUD_output) Then Exit Sub

If LenB(arrTimer(0, 1)) > 0 Then
      MUD_output = Replace(MUD_output, "- armour", "- armour" & getMyTime(arrTimer(0, 1)), 1, 1, vbBinaryCompare)
End If
If LenB(arrTimer(1, 1)) > 0 Then
      MUD_output = Replace(MUD_output, "- shield", "- shield" & getMyTime(arrTimer(1, 1)), 1, 1, vbBinaryCompare)
End If
If LenB(arrTimer(2, 1)) > 0 Then
      MUD_output = Replace(MUD_output, "- strength", "- strength" & getMyTime(arrTimer(2, 1)), 1, 1, vbBinaryCompare)
End If
If LenB(arrTimer(3, 1)) > 0 Then
      MUD_output = Replace(MUD_output, "- sanctuary", "- sanctuary" & getMyTime(arrTimer(3, 1)), 1, 1, vbBinaryCompare)
End If
If LenB(arrTimer(4, 1)) > 0 Then
      MUD_output = Replace(MUD_output, "- breath of briskness", "- breath of briskness" & getMyTime(arrTimer(4, 1)), 1, 1, vbBinaryCompare)
End If

'Debug.Print "MappingMode: " & MappingMode
'Debug.Print "MappingData: " & MappingData
'Debug.Print getRoomDescription(MUD_output)

If MappingData Then
   If Not mnuMapDescription.Checked Then
      Dim mydesc As String
      Dim a, b, c As Integer
      mydesc = vbNullString
      a = InStrB(1, MUD_output, lookColour, vbBinaryCompare)
      If a > 0 Then
         b = a + LenB(lookColour)
         c = InStrB(b, MUD_output, colourEndCode & vbCrLf, vbBinaryCompare) ' roomname colour end
         If c > 0 Then
            a = c + LenB(colourEndCode & vbCrLf) ' the beginning of description
            ' find last description row
            b = (InStrRev(MUD_output, roomdescriptionColour, , vbBinaryCompare) * 2) - 1  'the last descrption row
            If b > 0 Then
               c = InStrB(b, MUD_output, colourEndCode, vbBinaryCompare) ' description colour end
               If c > 0 Then
                  mydesc = MidB(MUD_output, a, c - a)
                  frmMap.tcpMUD.SendData Replace(MUD_output, vbLf & mydesc, " (mapping)", 1, 1, vbBinaryCompare)
               End If
            End If
         End If
      End If
   Else
      frmMap.tcpMUD.SendData MUD_output 'send MUME data to Client
   End If
Else
   'MUD_output
   frmMap.tcpMUD.SendData MUD_output 'send MUME data to Client
End If







'for search optimization
   If InStrB(1, MUD_output, "You ", vbBinaryCompare) > 0 Then isYou1 = True Else isYou1 = False
   If InStrB(1, MUD_output, " you", vbBinaryCompare) > 0 Then isYou2 = True Else isYou2 = False
   
   oldMUDOutput = MUD_output
      
   If handleCollision(MUD_output) Then Exit Sub   '#viimane
   
   
   If handleRunMode(MUD_output) Then
      If GODMODE And mnuInformer.Checked Then
         If LOST = False Then
            Open App.Path & "\_me.txt" For Output As #1
            Print #1, "&row=" & theROW & "&col=" & theCOL & "&row2=" & virtualRow & "&col2=" & virtualCol
            Close #1
         End If
      End If
      Exit Sub     '#eelviimane
   End If
   
   
   If isYou2 Then If mnuPlayers.Checked Then If handleWhere(MUD_output) Then Exit Sub
   If MappingCase < 2 Then If MappingCase = 1 Then MappingCase = 2
   
   If handleMapping(MUD_output) = True Then Call DrawMap: Exit Sub
Exit Sub

errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "frmMap tcpPlayer_DataArrival"
   writeError (errorModule)
'Call handleTell(MUD_output)
'   If frmMap.mnuEnemies.Checked Then
'      If InStrB(1, MUD_output, "* leaves ", vbBinaryCompare) > 0 Then
'         Call handleEnemy(MUD_output, InStr(1, MUD_output, "* leaves ", vbBinaryCompare))
'      End If
'   End If
End Sub

Public Sub tcpMUD_ConnectionRequest(ByVal requestID As Long)
errorData = "tcpMUD_ConnectionRequest -> "
If DEBUGMODE = False Then On Error GoTo errorhandler
   If WorldLoaded = True And tcpMUD.State <> sckClosed Then
      frmMap.tcpMUD.Close
      frmMap.tcpMUD.Accept requestID
      frmMap.tcpPlayer.Connect
   End If
Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "frmMap tcpMUD_ConnectionRequest"
   writeError (errorModule)
End Sub

Public Sub Form_DblClick()
If DEBUGMODE = False Then On Error GoTo errorhandler
Dim tmpRow As Integer, tmpCol As Integer
   
staticLevel = True
  
   tmpRow = Round((mouseY - frmMap.ScaleHeight / 2) / roomsize, 0)
   tmpCol = Round((mouseX - frmMap.ScaleWidth / 2) / roomsize, 0)
   If isValid(theROW + tmpRow, theCOL + tmpCol) Then
      surfing = True
      theROW = theROW + tmpRow
      theCOL = theCOL + tmpCol
      
'Debug.Print " --- dblclick --- "
cursor = getInt(aWorld(theROW, theCOL, theLEVEL))
'Debug.Print "   N (" & aData(cursor, cNPORTALR) & "," & aData(cursor, cNPORTALC) & ") - " & aData(cursor, cNLEVEL)
'Debug.Print "   E (" & aData(cursor, cEPORTALR) & "," & aData(cursor, cEPORTALC) & ") - " & aData(cursor, cELEVEL)
'Debug.Print "   S (" & aData(cursor, cSPORTALR) & "," & aData(cursor, cSPORTALC) & ") - " & aData(cursor, cSLEVEL)
'Debug.Print "   W (" & aData(cursor, cWPORTALR) & "," & aData(cursor, cWPORTALC) & ") - " & aData(cursor, cWLEVEL)
'Debug.Print " =========== "
      
      
      If MappingMode Then
         'special mappingmode change
         virtualRow = theROW
         virtualCol = theCOL
         'special mappingmode change end
         If getIndex(theROW, theCOL) > 0 Then
            Call loadRoom(theROW, theCOL) 'loadold
            Call GetMapData
         End If
      Else
         If getIndex(theROW, theCOL) > 0 Then
            'set old values to link the portals
            oldLevel = theLEVEL
            oldRow = theROW
            oldCol = theCOL
            
            theDoornameNorth = aData(getIndex(theROW, theCOL), cNDOOR)
            theDoornameEast = aData(getIndex(theROW, theCOL), cEDOOR)
            theDoornameSouth = aData(getIndex(theROW, theCOL), cSDOOR)
            theDoornameWest = aData(getIndex(theROW, theCOL), cWDOOR)
            theDoornameUp = aData(getIndex(theROW, theCOL), cUDOOR)
            theDoornameDown = aData(getIndex(theROW, theCOL), cDDOOR)
         End If
      End If
      If getIndex(theROW, theCOL) > 0 Then
         frmMap.Caption = aData(getIndex(theROW, theCOL), cROOMNAME) & "[" & theROW & "," & theCOL & "]"
      Else
         frmMap.Caption = mapTitle
      End If
      Call DrawMap
   End If
   If MappingMode And selectType > 0 Then
      'do nothing
   Else
      If buttonPressed = 1 Then 'left button
         If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
      End If
   End If
   
   'Debug.Print theROW & "," & theCOL
staticLevel = False
   
Exit Sub
errorhandler:
End Sub

Public Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call selectStart(X, Y)
End Sub

Public Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call selectEnd(X, Y)
   mouseX = X
   mouseY = Y

   If Button = 1 Then
      buttonPressed = 1
   End If
   If Button = 2 Then
      buttonPressed = 2
   End If

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'errorData = "Form QueryUnload-> "
'If DEBUGMODE = False Then On Error GoTo errorhandler
   Dim reply
   reply = MsgBox("Programm will exit. Do you want to save the world?", vbYesNoCancel, "Quit")
   Select Case reply
   Case vbYes
      Call mnuSave_Click
      frmMap.tcpMUD.Close
      Call ZIP
   Case vbNo
      frmMap.tcpMUD.Close
      Call ZIP
   Case Else
      Cancel = True
   End Select
Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "Form_Unload"
   writeError (errorModule)
End Sub

Public Sub Form_Resize()
   BitBlt frmMap.hdc, 0, 0, frmMap.ScaleWidth, frmMap.ScaleHeight, 0, 0, 0, vbBlackness
   If WorldLoaded = True Then Call DrawMap
End Sub

Public Sub mnuAbout_Click()
   informClient ("Version " & App.Major & "." & App.Minor & "." & App.Revision & " (" & theCount & " rooms)")
   MsgBox vbCrLf & "http://mume.blogspot.com" & vbCrLf & "e-mail: jaanus@2in.ee" & vbCrLf & "subject: MUME_Online_Map_Feedback" & vbCrLf & App.Major & "." & App.Minor & "." & App.Revision & " (" & theCount & " rooms)", vbOKOnly, mapTitle
End Sub

Public Sub mnuAddNote_Click()
   If isValid(theROW, theCOL) = False Then Exit Sub
   If getIndex(theROW, theCOL) <> 0 Then
      Dim note As String
      note = Trim(InputBox(vbCrLf & "Please type your note:" & vbCrLf & "(up to 100 letters)", "New Note", aData(getIndex(theROW, theCOL), cNOTE)))
      aData(getIndex(theROW, theCOL), cNOTE) = Mid(note, 1, 100)
      Call updateThis(getIndex(theROW, theCOL))
      Call DrawMap
   End If
End Sub

Public Sub mnuAlwaysOnTop_Click()
   If mnuAlwaysOnTop.Checked Then
      mnuAlwaysOnTop.Checked = False
      alwaysOnTop = False
      WinTopMost.MakeNormal frmMap.hWnd
      WinTopMost.MakeNormal frmTools.hWnd
   Else
      mnuAlwaysOnTop.Checked = True
      alwaysOnTop = True
      WinTopMost.MakeTopMost frmMap.hWnd
      WinTopMost.MakeTopMost frmTools.hWnd
   End If
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "alwaysontop", frmMap.mnuAlwaysOnTop.Checked)
End Sub

Public Sub mnuAutosync_Click()
   If mnuAutosync.Checked Then
      mnuAutosync.Checked = False
      Autosync = False
   Else
      mnuAutosync.Checked = True
      Autosync = True
   End If
   If WorldLoaded Then Call saveMOMini
End Sub

Public Sub mnuBrief_Click()
   If mnuBrief.Checked Then
      isBriefMode = False
      mnuBrief.Checked = False
   Else
      isBriefMode = True
      mnuBrief.Checked = True
   End If
   If WorldLoaded Then Call saveMOMini     'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "brief", frmMap.mnuBrief.Checked)
End Sub

Public Sub mnuClient_Click()
If DEBUGMODE = False Then On Error GoTo errorhandler
   Dim name As String
   name = Mid(Trim(InputBox(vbCrLf & "Please type your client name:" & vbCrLf & vbCrLf & "" & theClientName & "" & vbCrLf & "telnet" & vbCrLf & "...")), 1, 30)
   theClientName = name
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "client", theClientName)
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
Exit Sub
errorhandler:
End Sub

Public Sub mnuCopy_Click()
   selectType = selectCopy
End Sub

Public Sub mnuCut_Click()
   selectType = selectCut
End Sub

Public Sub mnuDelete_Click()
   selectType = selectDelete
   Call handleSelection
   selectType = 0
End Sub

Public Sub mnuDeleteRoom_Click()
Call frmTools.button_reset_Click
End Sub

Public Sub mnuFollow_Click()
   If mnuFollow.Checked = True Then
      mnuFollow.Checked = False
      leader = ""
   Else
      mnuFollow.Checked = True
      leader = InputBox("Please type your leader name(case sensitive)", "Set Leader", leader)
      If Len(leader) = 0 Then mnuFollow.Checked = False
   End If
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub

'FOREGROUND
Public Sub mnuLookBLACK_Click()
   fgColour = BLACK: Call changeRoomColour
End Sub
Public Sub mnuLookBLUE_Click()
   fgColour = BLUE: Call changeRoomColour
End Sub
Public Sub mnuLookBold_Click()
   If mnuLookBold.Checked Then
      fgBold = ""
      Call changeRoomColour
   Else
      fgBold = BOLD
      Call changeRoomColour
   End If
End Sub
Public Sub mnuLookCYAN_Click()
   fgColour = CYAN: Call changeRoomColour
End Sub
Public Sub mnuLookGREEN_Click()
   fgColour = GREEN: Call changeRoomColour
End Sub
Public Sub mnuLookMAGENTA_Click()
   fgColour = MAGENTA: Call changeRoomColour
End Sub
Public Sub mnuLookRED_Click()
   fgColour = RED: Call changeRoomColour
End Sub
Public Sub mnuLookUnderline_Click()
   If mnuLookUnderline.Checked Then
      fgUnderline = ""
      Call changeRoomColour
   Else
      fgUnderline = UNDERLINE
      Call changeRoomColour
   End If
End Sub
Public Sub mnuLookYELLOW_Click()
   fgColour = YELLOW: Call changeRoomColour
End Sub
Public Sub mnuLookWHITE_Click()
   fgColour = WHITE: Call changeRoomColour
End Sub
'BACKGROUND
Public Sub mnuLookBgNONE_Click()
   bgColour = "": Call changeRoomColour
End Sub
Public Sub mnuLookBgBLACK_Click()
   bgColour = CStr(CInt(BLACK) + 10): Call changeRoomColour
End Sub
Public Sub mnuLookBgBLUE_Click()
   bgColour = CStr(CInt(BLUE) + 10): Call changeRoomColour
End Sub
Public Sub mnuLookBgCYAN_Click()
   bgColour = CStr(CInt(CYAN) + 10): Call changeRoomColour
End Sub
Public Sub mnuLookBgGREEN_Click()
   bgColour = CStr(CInt(GREEN) + 10): Call changeRoomColour
End Sub
Public Sub mnuLookBgMAGENTA_Click()
   bgColour = CStr(CInt(MAGENTA) + 10): Call changeRoomColour
End Sub
Public Sub mnuLookBgRED_Click()
   bgColour = CStr(CInt(RED) + 10): Call changeRoomColour
End Sub
Public Sub mnuLookBgYELLOW_Click()
   bgColour = CStr(CInt(YELLOW) + 10): Call changeRoomColour
End Sub
Public Sub mnuLookBgWHITE_Click()
   bgColour = CStr(CInt(WHITE) + 10): Call changeRoomColour
End Sub

Public Sub mnuMap_Click()
   If WorldLoaded = False Then Exit Sub
   Call setMapModeON
   Call DrawMap
   If alwaysOnTop Then WinTopMost.MakeTopMost frmMap.hWnd
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub

Public Sub mnuNotes_Click()
   If frmMap.mnuNotes.Checked Then
      frmMap.mnuNotes.Checked = False
      viewNotes = False
   Else
      frmMap.mnuNotes.Checked = True
      viewNotes = True
   End If
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "notes", frmMap.mnuNotes.Checked)
End Sub

Public Sub mnuPaste_Click()
   Call handleSelection
   selectType = 0
End Sub

Public Sub mnuQuit_Click()
   Call Form_QueryUnload(0, 0)
End Sub

Public Sub mnuRoomsync_Click()
   If mnuRoomsync.Checked = True Then
      mnuRoomsync.Checked = False
      Roomsync = False
   Else
      mnuRoomsync.Checked = True
      Roomsync = True
   End If
End Sub

Public Sub mnuSpam_Click()
   If mnuSpam.Checked Then
      isSpamMode = False
      mnuSpam.Checked = False
   Else
      isSpamMode = True
      mnuSpam.Checked = True
   End If
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "spam", frmMap.mnuSpam.Checked)
End Sub

Public Sub mnuTellBLACK_Click()
   Call changeTellColour(BLACK)
End Sub

Public Sub mnuTellBLUE_Click()
   Call changeTellColour(BLUE)
End Sub

Public Sub mnuTellCYAN_Click()
   Call changeTellColour(CYAN)
End Sub

Public Sub mnuTellGREEN_Click()
   Call changeTellColour(GREEN)
End Sub

Public Sub mnuTellMAGENTA_Click()
   Call changeTellColour(MAGENTA)
End Sub

Public Sub mnuTellRED_Click()
   Call changeTellColour(RED)
End Sub

Public Sub mnuTellWHITE_Click()
   Call changeTellColour(WHITE)
End Sub

Public Sub mnuTellYELLOW_Click()
   Call changeTellColour(YELLOW)
End Sub

Public Sub mnuTools_Click()
   If WorldLoaded = False Then Exit Sub
   If frmTools.Visible Then
      frmTools.Hide
'     frmMap.mnuTools.Checked = False
   Else
      If MappingMode = False Then
         Call setMapModeON
      End If
      frmTools.WindowState = vbNormal
      WinTopMost.MakeTopMost frmTools.hWnd
      frmTools.Show
'     frmMap.mnuTools.Checked = True
   End If
   Call DrawMap
   If alwaysOnTop Then WinTopMost.MakeTopMost frmMap.hWnd
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub
Public Sub mnuWorld_Click()
   Call frmWorld.drawWorld(0, 0, 1)
End Sub

Public Sub Form_Load()
errorData = "Form Load -> "
If DEBUGMODE = False Then On Error GoTo errorhandler

   frmLogin.Hide
   'Set objInformer = New Informer.cAsync
   'frmMap.objInformer.Init
   'frmMap.objInformer.start ("")
   'SleepAPI 100

   'Set objReceiver = New Informer.cWhere
   'frmMap.objReceiver.Init
   'frmMap.objReceiver.start ("")
   'SleepAPI 100
   
   Set md5 = CreateObject("MD5DLL.Crypt")
   Set cast128 = CreateObject("cast.cipher")
'   registryPath = "SOFTWARE\LangSoft\MUME Online Map"
   filePath = App.Path & "\map51.txt"
   Initialized = False
'show logo
   Me.ScaleMode = 3
   frmLogo.Visible = True
   frmLogo.Show
   frmLogo.SetFocus
   frmLogo.Refresh
'init critical variables
   arrMinRow = 1
   arrMinCol = 1
   arrMaxData = UBound(aData)
   arrMaxRow = UBound(aWorld, 1)
   arrMaxCol = UBound(aWorld, 2)
   arrMinRoom = LBound(arrRoomstack)
   arrMaxRoom = UBound(arrRoomstack)
   arrMinMove = LBound(arrMovestack)
   arrMaxMove = UBound(arrMovestack)
'load mom.ini
   Call loadVariables
'load world from file
   WorldLoaded = False
   Call loadWorld
   If Not (WorldLoaded) Then GoTo errorhandler
'load defaults
   Initialized = True
'init graphics
'WinTopMost.MakeTopMost frmMap.hWnd
   Call loadRoom(theROW, theCOL)
   Call DrawMap
'open port for client
   frmMap.tcpMUD.Listen
'runtime
   frmLogo.Hide
   frmMap.Show
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
Exit Sub

errorhandler:
   frmLogo.Hide
   frmLogo.Visible = False
   MsgBox "Invalid installation or corrupted database!" & vbCrLf & vbCrLf & Err.description & "(" & Err.Number & ")" & _
         vbCrLf & vbCrLf & "Executing MOM Setup (setup.exe)."
   Dim tulistaja
   tulistaja = ShellExecute(0, "", "setup.exe", "", App.Path, 0)
   errorModule = Err.description & "(" & Err.Number & ") -> " & "Program Load"
   writeError (errorModule)
   End
End Sub

Public Sub mnumap3_Click()
   Call loadGraphics(App.Path & mapLargePath)
   mnuMap1.Checked = False
   mnuMap2.Checked = False
   mnuMap3.Checked = True
   roomsize = 14
   mapRadius = Fix(((frmMap.ScaleWidth + frmMap.ScaleHeight) / 4) / roomsize)
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "map", "3")
   Call DrawMap
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub

Public Sub mnuLocate_Click()
   Call WorldLocate(currentRoomname, currentExits)   'Call AreaLocate
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub

Public Sub mnuMap2_Click()
   Call loadGraphics(App.Path & mapNormalPath)
   mnuMap1.Checked = False
   mnuMap2.Checked = True
   mnuMap3.Checked = False
   roomsize = 22
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "map", "2")
   Call DrawMap
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub

'Public Sub mnuOnTop_Click()
'   If mnuOnTop.Checked = True Then
'      mnuOnTop.Checked = False
'      WinTopMost.MakeNormal frmMap.hWnd
'      WinTopMost.MakeNormal frmTools.hWnd
'      frmMap.Hide
'   Else
'      mnuOnTop.Checked = True
'      If alwaysOnTop Then WinTopMost.MakeTopMost frmMap.hWnd
'      WinTopMost.MakeTopMost frmTools.hWnd
'      frmMap.Show
'   End If
'End Sub

Public Sub mnuMap1_Click()
   Call loadGraphics(App.Path & mapSmallPath)
   mnuMap1.Checked = True
   mnuMap2.Checked = False
   mnuMap3.Checked = False
   roomsize = 32
   mapRadius = Fix(((frmMap.ScaleWidth + frmMap.ScaleHeight) / 4) / roomsize)
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "map", "1")
   Call DrawMap
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub

Public Sub mnuDoornames_Click()
   If frmMap.mnuDoornames.Checked = False Then
      frmMap.mnuDoornames.Checked = True
      viewDoornames = True
   Else
      frmMap.mnuDoornames.Checked = False
      viewDoornames = False
   End If
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "doornames", frmMap.mnuDoornames.Checked)
   Call DrawMap
End Sub

Public Sub mnuMovement_Click()
   If frmMap.mnuMovement.Checked = False Then
      frmMap.mnuMovement.Checked = True
   Else
      frmMap.mnuMovement.Checked = False
   End If
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "movement", frmMap.mnuMovement.Checked)
   Call DrawMap
End Sub

Public Sub mnuPortals_Click()
   If frmMap.mnuPortals.Checked = False Then
      frmMap.mnuPortals.Checked = True
   Else
      frmMap.mnuPortals.Checked = False
   End If
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "portals", frmMap.mnuPortals.Checked)
   Call DrawMap
End Sub

Public Sub mnuSave_Click()
   Dim answer
   answer = MsgBox("Save world? All your changes will be made permanent", vbYesNo, "Save world")
   Select Case answer
      Case vbYes
         Call saveWorld
      Case vbNo
      Case Else
   End Select
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub

Public Sub changeRoomColour(Optional ByVal startup As Boolean, Optional ByVal s As String)
   frmMap.mnuLookBold.Checked = False
   frmMap.mnuLookUnderline.Checked = False
'foreground
   frmMap.mnuLookBLACK.Checked = False
   frmMap.mnuLookRED.Checked = False
   frmMap.mnuLookGREEN.Checked = False
   frmMap.mnuLookYELLOW.Checked = False
   frmMap.mnuLookBLUE.Checked = False
   frmMap.mnuLookMAGENTA.Checked = False
   frmMap.mnuLookCYAN.Checked = False
   frmMap.mnuLookWHITE.Checked = False
'background
   frmMap.mnuLookBgBLACK.Checked = False
   frmMap.mnuLookBgRED.Checked = False
   frmMap.mnuLookBgGREEN.Checked = False
   frmMap.mnuLookBgYELLOW.Checked = False
   frmMap.mnuLookBgBLUE.Checked = False
   frmMap.mnuLookBgMAGENTA.Checked = False
   frmMap.mnuLookBgCYAN.Checked = False
   frmMap.mnuLookBgWHITE.Checked = False

   If startup Then
      If Len(s) > 0 Then
         fgColour = Mid(s, 3, 2)
         If InStr(5, s, ";1", vbBinaryCompare) > 0 Then
            fgBold = BOLD
         Else
            fgBold = ""
         End If
         If InStr(5, s, ";4m", vbBinaryCompare) > 0 Then
            fgUnderline = UNDERLINE
         Else
            fgUnderline = ""
         End If
         If Mid(s, 6, 1) = "4" And Mid(s, 7, 1) <> "m" Then
            bgColour = Mid(s, 6, 2)
         Else
            bgColour = ""
         End If
      Else
         fgColour = GREEN
         fgBold = ""
         fgUnderline = ""
         bgColour = ""
      End If
   End If
   
   If fgBold = BOLD Then frmMap.mnuLookBold.Checked = True
   If fgUnderline = UNDERLINE Then frmMap.mnuLookUnderline.Checked = True
   'foreground
   If fgColour = BLACK Then frmMap.mnuLookBLACK.Checked = True
   If fgColour = RED Then frmMap.mnuLookRED.Checked = True
   If fgColour = GREEN Then frmMap.mnuLookGREEN.Checked = True
   If fgColour = YELLOW Then frmMap.mnuLookYELLOW.Checked = True
   If fgColour = BLUE Then frmMap.mnuLookBLUE.Checked = True
   If fgColour = MAGENTA Then frmMap.mnuLookMAGENTA.Checked = True
   If fgColour = CYAN Then frmMap.mnuLookCYAN.Checked = True
   If fgColour = WHITE Then frmMap.mnuLookWHITE.Checked = True
   'background
   If bgColour = CStr(CInt(BLACK) + 10) Then frmMap.mnuLookBgBLACK.Checked = True
   If bgColour = CStr(CInt(RED) + 10) Then frmMap.mnuLookBgRED.Checked = True
   If bgColour = CStr(CInt(GREEN) + 10) Then frmMap.mnuLookBgGREEN.Checked = True
   If bgColour = CStr(CInt(YELLOW) + 10) Then frmMap.mnuLookBgYELLOW.Checked = True
   If bgColour = CStr(CInt(BLUE) + 10) Then frmMap.mnuLookBgBLUE.Checked = True
   If bgColour = CStr(CInt(MAGENTA) + 10) Then frmMap.mnuLookBgMAGENTA.Checked = True
   If bgColour = CStr(CInt(CYAN) + 10) Then frmMap.mnuLookBgCYAN.Checked = True
   If bgColour = CStr(CInt(WHITE) + 10) Then frmMap.mnuLookBgWHITE.Checked = True
   'set look colour
   lookColour = lookHeader
   lookColour = lookColour & fgColour
   If Len(bgColour) > 0 Then lookColour = lookColour & ";" & bgColour
   If Len(fgBold) > 0 Then lookColour = lookColour & ";" & fgBold
   If Len(fgUnderline) > 0 Then lookColour = lookColour & ";" & fgUnderline
   lookColour = lookColour & lookFooter
   
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "lookColour", lookColour)
End Sub

Public Sub changeTellColour(Colour As String)
'   frmMap.mnuTellBLACK.Checked = False
'   frmMap.mnuTellRED.Checked = False
'   frmMap.mnuTellGREEN.Checked = False
'   frmMap.mnuTellYELLOW.Checked = False
'   frmMap.mnuTellBLUE.Checked = False
'   frmMap.mnuTellMAGENTA.Checked = False
'   frmMap.mnuTellCYAN.Checked = False
'   frmMap.mnuTellWHITE.Checked = False
'
'   If Colour = BLACK Then fgColour = BLACK: frmMap.mnuTellBLACK.Checked = True
'   If Colour = RED Then fgColour = RED: frmMap.mnuTellRED.Checked = True
'   If Colour = GREEN Then fgColour = GREEN: frmMap.mnuTellGREEN.Checked = True
'   If Colour = YELLOW Then fgColour = YELLOW: frmMap.mnuTellYELLOW.Checked = True
'   If Colour = BLUE Then fgColour = BLUE: frmMap.mnuTellBLUE.Checked = True
'   If Colour = MAGENTA Then fgColour = MAGENTA: frmMap.mnuTellMAGENTA.Checked = True
'   If Colour = CYAN Then fgColour = CYAN: frmMap.mnuTellCYAN.Checked = True
'   If Colour = WHITE Then fgColour = WHITE: frmMap.mnuTellWHITE.Checked = True
End Sub

Public Sub tcpMUD_Close()
   frmMap.tcpPlayer.Close
   frmMap.tcpMUD.Close
   frmMap.tcpMUD.Listen
End Sub

Public Sub tcpMUD_DataArrival(ByVal bytesTotal As Long)
errorData = "tcpMUD_DataArrival -> "
If DEBUGMODE = False Then On Error GoTo errorhandler
Dim Client_Output As String
frmMap.tcpMUD.getData Client_Output   'get client input data
   
   If frmMap.tcpPlayer.State = sckConnected Then
      
      'frmMap.objReceiver.Init
      'frmMap.objReceiver.start ("?where=" & characterName)
   
      atLf = InStrB(1, Client_Output, vbLf, vbBinaryCompare)
      atCrLf = InStrB(1, Client_Output, vbCrLf, vbBinaryCompare)
      If atLf > 0 Then
         Dim out As String
         out = tmpOutput & Client_Output
         tmpOutput = ""
         
         If atCrLf > 0 Then   ' zmud, telnet
            specialLen = LenB(vbCrLf)    ' CrLf
         Else                 'jmc
            specialLen = LenB(vbLf) ' Lf
         End If
         If handleSpecial(out) Then Exit Sub
         If handleMappingCommand(out) Then Exit Sub
         If handleRuntimeCommand(out) Then Exit Sub
      Else                    'telnet
         If Asc(Client_Output) = 8 Then
            If LenB(tmpOutput) > 1 Then tmpOutput = MidB(tmpOutput, 1, LenB(tmpOutput) - 1)
         Else
            tmpOutput = tmpOutput & Client_Output
         End If
      End If
   Else
      Call informClient(" disconnected." & vbCrLf & "#ZAP and reconnect from client to <localhost> <1001>")
   End If
Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "tcpMUD DataArrival"
   writeError (errorModule)
End Sub
'--------------------------------------------------
'sckClosed 0 Default. Closed
'sckOpen 1 Open
'sckListening 2 Listening
'sckConnectionPending  3 Connection pending
'sckResolvingHost  4 Resolving host
'sckHostResolved  5 Host resolved
'sckConnecting  6 Connecting
'sckConnected  7 Connected
'sckClosing  8 Peer is closing the connection
'sckError  9 Error

Public Sub mnuFeedback_Click()
   If mnuFeedback.Checked Then
      mnuFeedback.Checked = False
   Else
      mnuFeedback.Checked = True
   End If
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "feedback", frmMap.mnuFeedback.Checked)
End Sub

Public Sub mnuPlayers_Click()
   If mnuPlayers.Checked Then
      mnuPlayers.Checked = False
   Else
      mnuPlayers.Checked = True
   End If
   If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "players", frmMap.mnuPlayers.Checked)
End Sub

Public Sub mnuPasteSpecial_Click()
   Call handleSelection(True)
   selectType = 0
End Sub

Public Sub mnuWalk_Click()
   MappingMode = True
   MappingData = False
   Call SYNC_FALSE("I see dead people!")
End Sub

Public Sub mnuRelocate_Click()
On Error Resume Next
   locateRetry = CInt(InputBox("Please enter the maximum number for room relocating", "Set Relocate", locateRetry))
   If Err.Number <> 0 Then locateRetry = 3
   If Err.Number <> 0 Then locateRetry = 3
On Error GoTo 0
If WorldLoaded Then Call saveMOMini   'Call oldregsave(HKEY_LOCAL_MACHINE, registryPath, "locateretry", CStr(locateRetry))
If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub

Public Sub mnuHere_Click()
   If WorldLoaded = False Then Exit Sub
   LOST = True
   MappingMode = False
   MappingData = False
'   roomcount = 0
'   locatorCount = 0
'   frmTools.Hide
'   frmMap.mnuTools.Checked = False
'   frmMap.mnuEdit.Enabled = False
   virtualRow = theROW
   virtualCol = theCOL
   Call SYNC_TRUE
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub

Public Sub mnuTarget_Click()
   If mnuTarget.Checked Then
      viewTarget = False
      mnuTarget.Checked = False
   Else
      viewTarget = True
      mnuTarget.Checked = True
      target = InputBox("Type target to search? CASE SENSITIVE!", "SET TARGET", target)
      If Len(target) = 0 Then mnuTarget.Checked = False
   End If
   If WorldLoaded Then SetForegroundWindow FindWindowPartial(LCase("*" & theClientName & "*"), "*")
End Sub

'==================================================================

'Public Sub mnuGroup_Click()
'   If mnuGroup.Checked = True Then
'      mnuGroup.Checked = False
'      frmGroup.Hide
'   Else
'      mnuGroup.Checked = True
'      frmGroup.Show
'   End If
'End Sub


Public Sub mnuEnemies_Click()
   If frmMap.mnuEnemies.Checked = False Then
      frmMap.mnuEnemies.Checked = True
   Else
      frmMap.mnuEnemies.Checked = False
   End If
End Sub

Public Sub mnuGrid_Click()
   If frmMap.mnuGrid.Checked = False Then
      frmMap.mnuGrid.Checked = True
   Else
      frmMap.mnuGrid.Checked = False
   End If
   Call DrawMap
   If WorldLoaded Then Call saveMOMini
End Sub

Public Sub mnuGridXY_Click()
   If frmMap.mnuGridXY.Checked = False Then
      frmMap.mnuGridXY.Checked = True
   Else
      frmMap.mnuGridXY.Checked = False
   End If
   Call DrawMap
   If WorldLoaded Then Call saveMOMini
End Sub

Private Sub mnuWalkthrough_Click()
   If mnuWalkthrough.Checked Then
      mnuWalkthrough.Checked = False
   Else
      mnuWalkthrough.Checked = True
   End If
   Call DrawMap
End Sub


Private Sub mnuSetLevel0_Click()
   Call setLevel(0)
End Sub

Private Sub mnuSetLevel1_Click()
   Call setLevel(1)
End Sub

Private Sub mnuSetLevel2_Click()
   Call setLevel(2)
End Sub

Private Sub mnuShowLevel0_Click()
   theLEVEL = 0
   Call DrawMap
End Sub

Private Sub mnuShowLevel1_Click()
   theLEVEL = 1
   Call DrawMap
End Sub

Private Sub mnuShowLevel2_Click()
   theLEVEL = 2
   Call DrawMap
End Sub
