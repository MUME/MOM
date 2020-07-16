VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMap 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "TheBestest"
   ClientHeight    =   4560
   ClientLeft      =   10065
   ClientTop       =   1035
   ClientWidth     =   4560
   DrawWidth       =   3
   FontTransparent =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   5160
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   4800
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuLocate 
      Caption         =   "Locate"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuFrom 
         Caption         =   "From"
      End
      Begin VB.Menu mnuTo 
         Caption         =   "To"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuMapper 
         Caption         =   "Mapper"
      End
      Begin VB.Menu mnuPortals 
         Caption         =   "Portals"
      End
      Begin VB.Menu mnuMovement 
         Caption         =   "Movement"
      End
      Begin VB.Menu mnuDoornames 
         Caption         =   "Doornames"
         Begin VB.Menu mnuDoornamesSolid 
            Caption         =   "Solid"
         End
         Begin VB.Menu mnuDoornamesTransparent 
            Caption         =   "Transparent"
         End
         Begin VB.Menu mnuDoornamesHide 
            Caption         =   "Hide"
         End
      End
      Begin VB.Menu mnuWorld 
         Caption         =   "World"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuSmall 
         Caption         =   "Small Map"
      End
      Begin VB.Menu mnuNormal 
         Caption         =   "Normal Map"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLarge 
         Caption         =   "Large Map"
      End
      Begin VB.Menu mnuOnTop 
         Caption         =   "Always On Top"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub mnubuffer_Click()
   frmMapBuffer.Visible = True
   frmMapBuffer.Show
   frmMapBuffer.Refresh
End Sub

Private Sub tcpServer_Close()
   frmMap.tcpClient.Close
   frmMap.tcpServer.Close
   frmMap.tcpServer.Listen
End Sub
Private Sub tcpServer_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errorhandler
Dim Client_output  As String
   
   frmMap.tcpServer.GetData Client_output   'get client input data
   If frmMap.tcpClient.State = sckConnected Then
      If handleSpecial(Client_output) Then Exit Sub    'enter etc..
      If handleMappingCommand(Client_output) Then Exit Sub
      If handleRuntimeCommand(Client_output) Then Exit Sub
   Else
      frmMap.tcpServer.SendData "BestEST - NOT CONNECTED" & vbLf
   End If

Exit Sub
errorhandler:
   errorData = "tcpServer DataArrival"
   writeError (errorData)
   frmTools.status = "CRITICAL ERROR! CONNECTION DOWN?"
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errorhandler
Dim MUME_output As String
   
   frmMap.tcpClient.GetData MUME_output    'get data from MUME
   
   If handleDescription(MUME_output) Then Exit Sub
     
   frmMap.tcpServer.SendData MUME_output           'send MUME data to Client
   
   If handleMapping(MUME_output) Then Exit Sub
   If handleCollision(MUME_output) Then Exit Sub
   If handleRunMode(MUME_output) Then Exit Sub

Exit Sub
errorhandler:
   errorData = "frmMap tcpClient_DataArrival"
   writeError (errorData)
   frmTools.status = "CRITICAL ERROR! CONNECTION DOWN?"
End Sub

Private Sub tcpServer_ConnectionRequest(ByVal requestID As Long)
On Error GoTo errorhandler
   If WorldLoaded = True And tcpServer.State <> sckClosed Then
      frmMap.tcpServer.Close
      frmMap.tcpServer.Accept requestID
      frmMap.tcpClient.RemoteHost = "mume.pvv.org"
      frmMap.tcpClient.RemotePort = 4242
      frmMap.tcpClient.Connect
   End If

Exit Sub
errorhandler:
   errorData = "frmMap tcpServer_ConnectionRequest"
   writeError (errorData)
End Sub

Private Sub Form_Load()
On Error GoTo errorhandler

   Call initError
   
'   Call updateDB
'   End

   Me.ScaleMode = 3
   frmMap.tcpServer.LocalPort = 1001
   frmMap.tcpServer.Listen

   Call loadVariables
   Call loadWorld
   virtualRow = 80
   virtualCol = 250
   roomCount = 0
   theRow = virtualRow
   theCol = virtualCol

   Call loadGraphics(App.Path & mapNormalPath)
   
   frmMap.mnuNormal.Checked = True
   frmMap.mnuPortals.Checked = True
   frmMap.mnuMovement.Checked = True
   frmMap.mnuDoornamesHide.Checked = False
   
   Call loadRoom(theRow, theCol)
   Call DrawMap
   
   Auto_Sync = True
   Room_Sync = True
   
   Call SYNC_FALSE

   Dim CurStyle As Long
   Dim NewStyle As Long
'   CurStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
'   CurStyle = CurStyle And Not (WS_MINIMIZEBOX)
'   CurStyle = CurStyle And Not (WS_MAXIMIZEBOX)
'   CurStyle = CurStyle And Not (WS_THICKFRAME)
'   CurStyle = CurStyle And Not (WS_SYSMENU)
'   CurStyle = 0
'   NewStyle = SetWindowLong(Me.hWnd, GWL_STYLE, CurStyle)
   
   WinTopMost.MakeTopMost Me.hWnd

   frmMap.Show
   
'   frmMapBuffer.Visible = True
'   frmMapBuffer.Show

Exit Sub
errorhandler:
   errorData = "frmMap onLoad"
   writeError (errorData)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMap.tcpServer.Close
   Call ZIP
End Sub

Public Sub mnuLarge_Click()
   Call loadGraphics(App.Path & mapLargePath)
   mnuSmall.Checked = False
   mnuNormal.Checked = False
   mnuLarge.Checked = True
   Call DrawMap
End Sub

Public Sub mnuLocate_Click()
   Call WorldLocate(currentRoomName, currentExits)   'Call AreaLocate
End Sub

Public Sub mnuMapper_Click()
   If frmTools.Visible = True Then
      Call setMapModeOFF
      frmTools.Visible = False
      frmMap.mnuMapper.Checked = False
   Else
      If WorldLoaded = False Then Exit Sub
      Call setMapModeON
      frmTools.Left = 1000
      frmTools.Top = 1000
      frmTools.Visible = True
      frmMap.mnuMapper.Checked = True
   End If
   frmMap.SetFocus
End Sub

Public Sub mnuNormal_Click()
   Call loadGraphics(App.Path & mapNormalPath)
   mnuSmall.Checked = False
   mnuNormal.Checked = True
   mnuLarge.Checked = False
   Call DrawMap
End Sub

Public Sub mnuOnTop_Click()
'   If mnuOnTop.Checked = True Then
'      mnuOnTop.Checked = False
'      WinTopMost.MakeNormal frmMap.hWnd
'      WinTopMost.MakeNormal frmTools.hWnd
'      frmMap.Hide
'   Else
'      mnuOnTop.Checked = True
'      WinTopMost.MakeTopMost frmMap.hWnd
'      WinTopMost.MakeTopMost frmTools.hWnd
'      frmMap.Show
'   End If
End Sub

Private Sub mnuQuit_Click()
   End
End Sub

Public Sub mnuSmall_Click()
   Call loadGraphics(App.Path & mapSmallPath)
   mnuSmall.Checked = True
   mnuNormal.Checked = False
   mnuLarge.Checked = False
   Call DrawMap
End Sub

Public Sub mnuDoornamesSolid_Click()
   frmMap.mnuDoornamesSolid.Checked = True
   frmMap.mnuDoornamesTransparent.Checked = False
   frmMap.mnuDoornamesHide.Checked = False
   frmMapBuffer.FontTransparent = False
   Call DrawMap
End Sub

Public Sub mnuDoornamesTransparent_Click()
   frmMap.mnuDoornamesSolid.Checked = False
   frmMap.mnuDoornamesTransparent.Checked = True
   frmMap.mnuDoornamesHide.Checked = False
   frmMapBuffer.FontTransparent = True
   Call DrawMap
End Sub
Public Sub mnuDoornamesHide_Click()
   If frmMap.mnuDoornamesHide.Checked = False Then
      frmMap.mnuDoornamesHide.Checked = True
   Else
      frmMap.mnuDoornamesHide.Checked = False
   End If
   Call DrawMap
End Sub

Public Sub mnuMovement_Click()
   If frmMap.mnuMovement.Checked = False Then
      frmMap.mnuMovement.Checked = True
   Else
      frmMap.mnuMovement.Checked = False
   End If
   Call DrawMap
End Sub

Public Sub mnuPortals_Click()
   If frmMap.mnuPortals.Checked = False Then
      frmMap.mnuPortals.Checked = True
   Else
      frmMap.mnuPortals.Checked = False
   End If
   Call DrawMap
End Sub

Private Sub mnuSave_Click()
   frmMap.Caption = "Please wait. Saving . . ."
   Call saveWorld
   frmMap.Caption = "Arda has been saved!"
End Sub

