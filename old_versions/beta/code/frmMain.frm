VERSION 5.00
Begin VB.MDIForm frmTemp 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "vana"
   ClientHeight    =   8010
   ClientLeft      =   7065
   ClientTop       =   3345
   ClientWidth     =   9675
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   ScrollBars      =   0   'False
   Begin VB.PictureBox PictureBox 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   0
      Width           =   9675
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   240
         Left            =   3600
         TabIndex        =   1
         Top             =   -45
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Binary
Option Explicit

' for window movement
Dim tpoint As POINTAPI
Dim temp  As POINTAPI
Dim dpoint As POINTAPI

Dim fbox As RECT
Dim tbox As RECT
Dim oldbox As RECT

Dim TwipsPerPixelX
Dim TwipsPerPixelY



'Private Sub button_End_DblClick()
'   End
'End Sub

Private Sub Command1_Click()
   Debug.Print Encrypt2("9 TJ§yˆióÐ_áGR~")
   Dim abc, bbc, ccc
   abc = ">BûÎ°”ú Æ6(S"
   bbc = convert.Base64.Base64Encode(abc)
   ccc = convert.Base64.Base64Decode(bbc)
   MsgBox abc & vbLf & bbc & vbLf & ccc
End Sub
'
'Private Sub MDIForm_DragDrop(Source As Control, X As Single, Y As Single)
'   MsgBox "x=" & X & vbCrLf & "y=" & Y
'End Sub
'
'Private Sub MDIForm_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'   MsgBox "x=" & X & vbCrLf & "y=" & Y
'End Sub


' Private Sub MDIForm_Load()
'   Dim CurStyle As Long
'   Dim NewStyle As Long
'
'   CurStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
'   CurStyle = CurStyle And Not (WS_MINIMIZEBOX)
'   CurStyle = CurStyle And Not (WS_MAXIMIZEBOX)
'   CurStyle = CurStyle And Not (WS_THICKFRAME)
'   CurStyle = CurStyle And Not (WS_SYSMENU)
'   CurStyle = 0
'   NewStyle = SetWindowLong(Me.hWnd, GWL_STYLE, CurStyle)
'   frmTemp.Visible = True
'   frmTemp.BackColor = &H0&
'   tcpServer.LocalPort = 1001
'   tcpServer.Listen
'   Call Initialize
'   WinTopMost.MakeTopMost Me.hWnd
'   MappingData = False
'   Erase arr
'   Erase arrDesc
'   Erase arrRoomStack
'   Erase arrMoveStack
'   arrMinRow = LBound(arr, 1)
'   arrMinCol = LBound(arr, 2)
'   arrMaxRow = UBound(arr, 1)
'   arrMaxCol = UBound(arr, 2)
'   arrMinRoom = LBound(arrRoomStack)
'   arrMaxRoom = UBound(arrRoomStack)
'   arrMinMove = LBound(arrMoveStack)
'   arrMaxMove = UBound(arrMoveStack)
'   Call loadWorld
'   virtualRow = 80
'   virtualCol = 250
'   roomCount = 0
'   theRow = virtualRow
'   theCol = virtualCol
'   Call loadRoom(theRow, theCol)
'   Call DrawMap
'   Call SYNC_FALSE
'   MappingGetUpdate = False
'   compactMode = False
'   Call switchmappingmode
'End Sub
'
'Private Sub PictureBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    BeginFRDrag X, Y
'End Sub
'
'Private Sub PictureBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then DoFRDrag X, Y
'End Sub
'
'Private Sub PictureBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    EndFRDrag X, Y
'End Sub
'
'
'Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    BeginFRDrag X, Y
'End Sub
'
'Private Sub titleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then DoFRDrag X, Y
'End Sub
'
'Private Sub titleBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    EndFRDrag X, Y
'End Sub
'
'Private Sub BeginFRDrag(X As Single, Y As Single)
'    Dim tDc As Long
'    Dim sDc As Long
'    Dim d As Long
'
'   'convert points to POINTAPI struct
'    dpoint.X = X
'    dpoint.Y = Y
'
'   'get screen area of toolbar
'    GetWindowRect frmTemp.hWnd, fbox
'   'screen RECT of toolbar
'    TwipsPerPixelX = Screen.TwipsPerPixelX
'    TwipsPerPixelY = Screen.TwipsPerPixelY
'
'   'get point of MouseDown in screen coordinates
'    temp = dpoint
'    ClientToScreen frmTemp.hWnd, temp
'
'    sDc = GetDC(ByVal 0)
'    DrawFocusRect sDc, tbox
'    d = ReleaseDC(0, sDc)
'    oldbox = tbox
'
'End Sub
'
'Private Sub DoFRDrag(X As Single, Y As Single)
'
'    Dim tDc As Long
'    Dim sDc As Long
'    Dim d As Long
'
'    tpoint.X = X
'    tpoint.Y = Y
'
'    ClientToScreen frmTemp.hWnd, tpoint
'
'    tbox.Left = (fbox.Left + tpoint.X / TwipsPerPixelX) - temp.X / TwipsPerPixelX
'    tbox.Top = (fbox.Top + tpoint.Y / TwipsPerPixelY) - temp.Y / TwipsPerPixelY
'    tbox.Right = (fbox.Right + tpoint.X / TwipsPerPixelX) - temp.X / TwipsPerPixelX
'    tbox.Bottom = (fbox.Bottom + tpoint.Y / TwipsPerPixelY) - temp.Y / TwipsPerPixelY
'
'    sDc = GetDC(ByVal 0)
'    DrawFocusRect sDc, oldbox
'    DrawFocusRect sDc, tbox
'    d = ReleaseDC(0, sDc)
'    oldbox = tbox
'
'End Sub
'
'Private Sub EndFRDrag(X As Single, Y As Single)
'
'    Dim tDc As Long
'    Dim sDc As Long
'    Dim d As Long
'
'    Dim newleft As Single
'    Dim newtop As Single
'
'    sDc = GetDC(ByVal 0)
'    DrawFocusRect sDc, oldbox
'    d = ReleaseDC(0, sDc)
'
'    newleft = X + fbox.Left * TwipsPerPixelX - dpoint.X
'    newtop = Y + fbox.Top * TwipsPerPixelY - dpoint.Y
'
'    frmTemp.Move newleft, newtop
'
'End Sub
'
