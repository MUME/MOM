VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form BestEST 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "BestEST"
   ClientHeight    =   5700
   ClientLeft      =   4995
   ClientTop       =   615
   ClientWidth     =   9405
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   627
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton button_GetMapData 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Get Map Data"
      Height          =   375
      Left            =   6480
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5160
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton button_save 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "SAVE ARDA"
      Height          =   375
      Left            =   8160
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5160
      Width           =   1095
      Visible         =   0   'False
   End
   Begin VB.CommandButton button_reset 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Reset"
      Height          =   495
      Left            =   7080
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   120
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.Frame Options 
      BackColor       =   &H00A56E3A&
      Caption         =   "Save Arda"
      Height          =   1455
      Left            =   5280
      TabIndex        =   40
      Top             =   3720
      Width           =   2895
      Visible         =   0   'False
      Begin VB.CommandButton button_saveCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2040
         TabIndex        =   51
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton button_SaveWorld 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         Height          =   375
         Left            =   1080
         TabIndex        =   49
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00A56E3A&
         Caption         =   "Take a copy from <world.mdb>,   then press ""Save""."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton button_report 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Report"
      Height          =   375
      Left            =   6720
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3240
      Width           =   735
      Visible         =   0   'False
   End
   Begin VB.CommandButton button_GetData 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Get Room Data"
      Height          =   375
      Left            =   5280
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3240
      Width           =   1335
      Visible         =   0   'False
   End
   Begin VB.CheckBox MapMode 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Map Mode"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5280
      TabIndex        =   37
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton button_update 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "UPDATE ROOM"
      Height          =   375
      Left            =   7560
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3240
      Width           =   1335
      Visible         =   0   'False
   End
   Begin VB.CommandButton button_goto 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Goto"
      Height          =   255
      Left            =   6240
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   360
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.CheckBox Monster 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Monster"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      MaskColor       =   &H000000FF&
      TabIndex        =   34
      Top             =   1320
      Width           =   1110
      Visible         =   0   'False
   End
   Begin VB.CheckBox Sun 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Sun"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      MaskColor       =   &H000000FF&
      TabIndex        =   31
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CheckBox Ridable 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Ridable"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      MaskColor       =   &H000000FF&
      TabIndex        =   30
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1095
      Visible         =   0   'False
   End
   Begin VB.TextBox Roomname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5760
      TabIndex        =   29
      Top             =   720
      Width           =   3150
      Visible         =   0   'False
   End
   Begin VB.TextBox Description 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   960
      Width           =   3150
      Visible         =   0   'False
   End
   Begin VB.CheckBox dExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Down"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   2880
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.CheckBox uExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "Up"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   26
      Top             =   2640
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.CheckBox wExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "West"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   25
      Top             =   2400
      Value           =   1  'Checked
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.CheckBox sExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "South"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   24
      Top             =   2160
      Value           =   1  'Checked
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.CheckBox eExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "East"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   23
      Top             =   1920
      Value           =   1  'Checked
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.CheckBox nExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      Caption         =   "North"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      MaskColor       =   &H000000FF&
      TabIndex        =   22
      Top             =   1680
      Value           =   1  'Checked
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.TextBox dNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   8160
      TabIndex        =   21
      Top             =   2880
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.TextBox uNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   8160
      TabIndex        =   20
      Top             =   2640
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.TextBox wNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   8160
      TabIndex        =   19
      Top             =   2400
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.TextBox sNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   8160
      TabIndex        =   18
      Top             =   2160
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.TextBox eNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   8160
      TabIndex        =   17
      Top             =   1920
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.TextBox nNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   8160
      TabIndex        =   16
      Top             =   1680
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.TextBox dDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   2880
      Width           =   1980
      Visible         =   0   'False
   End
   Begin VB.TextBox uDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   2640
      Width           =   1980
      Visible         =   0   'False
   End
   Begin VB.TextBox wDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   2400
      Width           =   1980
      Visible         =   0   'False
   End
   Begin VB.TextBox sDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   2160
      Width           =   1980
      Visible         =   0   'False
   End
   Begin VB.TextBox eDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   1920
      Width           =   1980
      Visible         =   0   'False
   End
   Begin VB.TextBox nDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   1680
      Width           =   1980
      Visible         =   0   'False
   End
   Begin VB.TextBox GoToNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00A56E3A&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.PictureBox map 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   450
      ScaleHeight     =   286
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   286
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   540
      Width           =   4290
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      Caption         =   "Load"
      Height          =   240
      Left            =   4560
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5385
      Width           =   540
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   8760
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   8280
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label label_col 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "COL:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   47
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label label_row 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "ROW:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   46
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label col 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "COL"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   45
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label row 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "ROW"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   44
      Top             =   4080
      Width           =   495
   End
   Begin VB.Image pictureTerrain 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   8400
      Top             =   195
      Width           =   480
   End
   Begin VB.Label label_terrain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Terrain"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   43
      Top             =   360
      Width           =   495
      Visible         =   0   'False
   End
   Begin VB.Image expand2 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   4920
      Picture         =   "Form1.frx":FAF8
      Stretch         =   -1  'True
      Top             =   240
      Width           =   180
   End
   Begin VB.Image compact2 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   120
      Picture         =   "Form1.frx":FEAC
      Stretch         =   -1  'True
      Top             =   240
      Width           =   180
   End
   Begin VB.Image wMove 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   6360
      Picture         =   "Form1.frx":10260
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image eMove 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7320
      Picture         =   "Form1.frx":106A2
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image sMove 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   6840
      Picture         =   "Form1.frx":10AE4
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image nMove 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   6840
      Picture         =   "Form1.frx":10F26
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image button_End 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   4995
      Picture         =   "Form1.frx":11368
      Stretch         =   -1  'True
      Top             =   5100
      Width           =   180
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   30
      Picture         =   "Form1.frx":1171C
      Stretch         =   -1  'True
      Top             =   5100
      Width           =   180
   End
   Begin VB.Image compact 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   30
      Picture         =   "Form1.frx":11AD0
      Stretch         =   -1  'True
      Top             =   45
      Width           =   180
   End
   Begin VB.Image expand 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   4995
      Picture         =   "Form1.frx":11E84
      Stretch         =   -1  'True
      Top             =   45
      Width           =   180
   End
   Begin VB.Label label_desc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Desc"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   33
      Top             =   990
      Width           =   495
      Visible         =   0   'False
   End
   Begin VB.Label Label_room 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Room"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   32
      Top             =   750
      Width           =   495
      Visible         =   0   'False
   End
   Begin VB.Image Vnone 
      Height          =   390
      Left            =   9720
      Picture         =   "Form1.frx":12238
      Top             =   6480
      Width           =   30
   End
   Begin VB.Image Hnone 
      Height          =   30
      Left            =   9240
      Picture         =   "Form1.frx":1258C
      Top             =   6960
      Width           =   390
   End
   Begin VB.Label status 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Ok."
      ForeColor       =   &H00C0FFC0&
      Height          =   210
      Left            =   840
      TabIndex        =   9
      Top             =   5400
      Width           =   3615
   End
   Begin VB.Label u_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   1440
   End
   Begin VB.Label d_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   3360
      TabIndex        =   7
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label w_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   360
      TabIndex        =   6
      Top             =   4920
      Width           =   1440
   End
   Begin VB.Label e_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   3360
      TabIndex        =   5
      Top             =   4920
      Width           =   1440
   End
   Begin VB.Label s_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   1800
      TabIndex        =   4
      Top             =   5160
      Width           =   1560
   End
   Begin VB.Label n_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1560
   End
   Begin VB.Image special 
      Height          =   360
      Left            =   9000
      Picture         =   "Form1.frx":128E5
      Stretch         =   -1  'True
      Top             =   720
      Width           =   360
      Visible         =   0   'False
   End
   Begin VB.Image none 
      Height          =   360
      Left            =   9000
      Picture         =   "Form1.frx":12D02
      Stretch         =   -1  'True
      Top             =   360
      Width           =   360
      Visible         =   0   'False
   End
   Begin VB.Image water 
      Height          =   360
      Left            =   9000
      Picture         =   "Form1.frx":13279
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   360
      Visible         =   0   'False
   End
   Begin VB.Image mountain 
      Height          =   360
      Left            =   9000
      Picture         =   "Form1.frx":137F5
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   360
      Visible         =   0   'False
   End
   Begin VB.Image hill 
      Height          =   360
      Left            =   9000
      Picture         =   "Form1.frx":13D4F
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   360
      Visible         =   0   'False
   End
   Begin VB.Image swamp 
      Height          =   360
      Left            =   9000
      Picture         =   "Form1.frx":142D3
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   360
      Visible         =   0   'False
   End
   Begin VB.Image road 
      Height          =   360
      Left            =   9000
      Picture         =   "Form1.frx":14775
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   360
      Visible         =   0   'False
   End
   Begin VB.Image forest 
      Height          =   360
      Left            =   9000
      Picture         =   "Form1.frx":14BB0
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   360
      Visible         =   0   'False
   End
   Begin VB.Image field 
      Height          =   360
      Left            =   9000
      Picture         =   "Form1.frx":150FB
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   360
      Visible         =   0   'False
   End
End
Attribute VB_Name = "BestEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub button_End_DblClick()
   End
End Sub

Private Sub button_GetData_Click()
   Call mapGet
End Sub

Private Sub button_GetMapData_Click()
   Call GetMapData
End Sub

Private Sub button_goto_Click()
   Call gotoArea(GoToNumber.Text)
End Sub

Private Sub button_report_Click()
   Call MapReport
End Sub

Private Sub button_reset_Click()
   Set BestEST.pictureTerrain.Picture = pNone
   arr(theRow, theCol) = 0
   arrDesc(theRow, theCol) = ""
   Call zeroMap
   Call LoadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub button_save_Click()
   Options.Visible = True
End Sub

Private Sub button_saveCancel_Click()
   Options.Visible = False
End Sub

Private Sub button_SaveWorld_Click()
   Call saveWorld
   Options.Visible = False
End Sub

Private Sub button_update_Click()
   Call mapUpdate
End Sub

Private Sub Command1_Click()
   MAP_THE_DATA = False
   Erase arr
   Erase arrDesc
   Erase arrRoomStack
   Erase arrMoveStack
   arrMinRow = LBound(arr, 1)
   arrMinCol = LBound(arr, 2)
   arrMaxRow = UBound(arr, 1)
   arrMaxCol = UBound(arr, 2)
   arrMinRoom = LBound(arrRoomStack)
   arrMaxRoom = UBound(arrRoomStack)
   arrMinMove = LBound(arrMoveStack)
   arrMaxMove = UBound(arrMoveStack)
   'Call LoadArea("C:\mume\NOC.xls", 0, 0)
   Call LoadWorld
   virtualRow = 15
   virtualCol = 15
   roomCount = 0
   theRow = virtualRow
   theCol = virtualRow
   Call LoadRoom(theRow, theCol)
   Call DrawMap
   Call SYNC_FALSE
End Sub

Private Sub compact_Click()
compactMode = Not (compactMode)
   If compactMode = True Then
      BestEST.width = 5200
      BestEST.height = 5700
      BestEST.Left = 9900
   Else
      BestEST.width = 5200
      BestEST.height = 5700
      BestEST.Left = 9900
   End If
End Sub

Private Sub Description_Change()
On Error GoTo ErrorHandler
   If Len(Description.Text) > 0 Then
      mapDescription = Mid(Description.Text, 1, 50)
   Else
      mapDescription = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   Description = ""
   Call InvalidData
End Sub

Private Sub dExit_Click()
   If dExit.Value = 1 Then
      mapRoomDown = True
   Else
      mapRoomDown = False
   End If
End Sub

Private Sub eExit_Click()
   If eExit.Value = 1 Then
      mapRoomEast = True
   Else
      mapRoomEast = False
   End If
End Sub

Private Sub eMove_Click()
   theCol = theCol + 1
   Call LoadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub expand_Click()
   BestEST.width = 9400
   BestEST.height = 5700
   BestEST.Left = 5700
End Sub

Private Sub expand2_Click()
   Options.Visible = Not (Options.Visible)
   BestEST.width = 9400
   BestEST.height = 5700
   BestEST.Left = 5700
End Sub

Private Sub field_Click()
   Set BestEST.pictureTerrain.Picture = pField
   setMapTerrain ("field")
End Sub

Private Sub forest_Click()
   Set BestEST.pictureTerrain.Picture = pForest
   setMapTerrain ("forest")
End Sub

Private Sub Form_Load()
   tcpServer.LocalPort = 1001
   tcpServer.Listen
   Call Initialize
   WinFunc.MakeTopMost Me.hwnd
   Call compact_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
 tcpServer.Close
 Set arr = Nothing
 Set arrDesc = Nothing
End Sub

Private Sub hill_Click()
   Set BestEST.pictureTerrain.Picture = pHill
   setMapTerrain ("hill")
End Sub

Private Sub Image3_Click()
   If WorldLoaded = True Then
      MAP_MODE = False
      fleeRadius = 10
      Call caseFleeHandler(currentRoomName, currentExits)
   End If
End Sub

Private Sub MapMode_Click()
   If WorldLoaded = False Then
      MapMode.Value = 0
      Exit Sub
   End If
   If MapMode.Value = 0 Then
      Call setMapModeOFF
      button_save.Visible = False
      nMove.Visible = False
      eMove.Visible = False
      sMove.Visible = False
      wMove.Visible = False
      button_GetMapData.Visible = False
      label_terrain.Visible = False
      button_reset.Visible = False
      pictureTerrain.Visible = False
      button_GetData.Visible = False
      button_report.Visible = False
      button_goto.Visible = False
      GoToNumber.Visible = False
      Label_room.Visible = False
      label_desc.Visible = False
      Roomname.Visible = False
      Description.Visible = False
      Sun.Visible = False
      Ridable.Visible = False
      Monster.Visible = False
      nExit.Visible = False
      nDoor.Visible = False
      nNumber.Visible = False
      eExit.Visible = False
      eDoor.Visible = False
      eNumber.Visible = False
      sExit.Visible = False
      sDoor.Visible = False
      sNumber.Visible = False
      wExit.Visible = False
      wDoor.Visible = False
      wNumber.Visible = False
      uExit.Visible = False
      uDoor.Visible = False
      uNumber.Visible = False
      dExit.Visible = False
      dDoor.Visible = False
      dNumber.Visible = False
      button_update.Visible = False
'      none.Visible = False
      special.Visible = False
      road.Visible = False
      field.Visible = False
      forest.Visible = False
      swamp.Visible = False
      hill.Visible = False
      mountain.Visible = False
      water.Visible = False
      status.Caption = "Map Mode Off"
   Else
      Call setMapModeON
      button_save.Visible = True
      nMove.Visible = True
      eMove.Visible = True
      sMove.Visible = True
      wMove.Visible = True
      button_GetMapData.Visible = True
      label_terrain.Visible = True
      button_reset.Visible = True
      pictureTerrain.Visible = True
      button_GetData.Visible = True
      button_report.Visible = True
      button_goto.Visible = True
      GoToNumber.Visible = True
      Label_room.Visible = True
      label_desc.Visible = True
      Roomname.Visible = True
      Description.Visible = True
      Sun.Visible = True
      Ridable.Visible = True
      Monster.Visible = True
      nExit.Visible = True
      nDoor.Visible = True
      nNumber.Visible = True
      eExit.Visible = True
      eDoor.Visible = True
      eNumber.Visible = True
      sExit.Visible = True
      sDoor.Visible = True
      sNumber.Visible = True
      wExit.Visible = True
      wDoor.Visible = True
      wNumber.Visible = True
      uExit.Visible = True
      uDoor.Visible = True
      uNumber.Visible = True
      dExit.Visible = True
      dDoor.Visible = True
      dNumber.Visible = True
      button_update.Visible = True
'      none.Visible = True
      special.Visible = True
      road.Visible = True
      field.Visible = True
      forest.Visible = True
      swamp.Visible = True
      hill.Visible = True
      mountain.Visible = True
      water.Visible = True
      status.Caption = "Map Mode On"
   End If
End Sub

Private Sub Monster_Click()
   If Monster.Value = 1 Then
      mapMonster = True
   Else
      mapMonster = False
   End If
End Sub

Private Sub mountain_Click()
   Set BestEST.pictureTerrain.Picture = pMountain
   setMapTerrain ("mountain")
End Sub

Private Sub nDoor_LostFocus()
On Error GoTo ErrorHandler
   If Len(nDoor.Text) > 0 Then
      mapDoorNameNorth = nDoor.Text
   Else
      mapDoorNameNorth = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   nDoor = ""
   Call InvalidData
End Sub
Private Sub eDoor_LostFocus()
On Error GoTo ErrorHandler
   If Len(eDoor.Text) > 0 Then
      mapDoorNameEast = eDoor.Text
   Else
      mapDoorNameEast = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   eDoor = ""
   Call InvalidData
End Sub

Private Sub nNumber_LostFocus()
On Error GoTo ErrorHandler
   If Len(nNumber.Text) > 0 Then
      tempData = Split(nNumber.Text, ",")
      mapRowNorth = tempData(0)
      mapColNorth = tempData(1)
   Else
      mapRowNorth = ""
      mapColNorth = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   nNumber = ""
   Call InvalidData
End Sub
Private Sub eNumber_LostFocus()
On Error GoTo ErrorHandler
   If Len(eNumber.Text) > 0 Then
      tempData = Split(eNumber.Text, ",")
      mapRowEast = tempData(0)
      mapColEast = tempData(1)
   Else
      mapRowEast = ""
      mapColEast = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   eNumber = ""
   Call InvalidData
End Sub

Private Sub road_Click()
   Set BestEST.pictureTerrain.Picture = pRoad
   setMapTerrain ("road")
End Sub

Private Sub sNumber_LostFocus()
On Error GoTo ErrorHandler
   If Len(sNumber.Text) > 0 Then
      tempData = Split(sNumber.Text, ",")
      mapRowSouth = tempData(0)
      mapColSouth = tempData(1)
   Else
      mapRowSouth = ""
      mapColSouth = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   sNumber = ""
   Call InvalidData
End Sub

Private Sub special_Click()
   Set BestEST.pictureTerrain.Picture = pSpecial
   setMapTerrain ("special")
End Sub

Private Sub swamp_Click()
   Set BestEST.pictureTerrain.Picture = pSwamp
   setMapTerrain ("swamp")
End Sub

Private Sub water_Click()
   Set BestEST.pictureTerrain.Picture = pWater
   setMapTerrain ("water")
End Sub

Private Sub wNumber_LostFocus()
On Error GoTo ErrorHandler
   If Len(wNumber.Text) > 0 Then
      tempData = Split(wNumber.Text, ",")
      mapRowWest = tempData(0)
      mapColWest = tempData(1)
   Else
      mapRowWest = ""
      mapColWest = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   wNumber = ""
   Call InvalidData
End Sub
Private Sub uNumber_LostFocus()
On Error GoTo ErrorHandler
   If Len(uNumber.Text) > 0 Then
      tempData = Split(uNumber.Text, ",")
      mapRowUp = tempData(0)
      mapColUp = tempData(1)
   Else
      mapRowUp = ""
      mapColUp = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   uNumber = ""
   Call InvalidData
End Sub
Private Sub dNumber_LostFocus()
On Error GoTo ErrorHandler
   If Len(dNumber.Text) > 0 Then
      tempData = Split(dNumber.Text, ",")
      mapRowDown = tempData(0)
      mapColDown = tempData(1)
   Else
      mapRowDown = ""
      mapColDown = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   dNumber = ""
   Call InvalidData
End Sub
Private Sub sDoor_LostFocus()
On Error GoTo ErrorHandler
   If Len(sDoor.Text) > 0 Then
      mapDoorNameSouth = sDoor.Text
   Else
      mapDoorNameSouth = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   sDoor = ""
   Call InvalidData
End Sub
Private Sub wDoor_LostFocus()
On Error GoTo ErrorHandler
   If Len(wDoor.Text) > 0 Then
      mapDoorNameWest = wDoor.Text
   Else
      mapDoorNameWest = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   wDoor = ""
   Call InvalidData
End Sub
Private Sub uDoor_LostFocus()
On Error GoTo ErrorHandler
   If Len(uDoor.Text) > 0 Then
      mapDoorNameUp = uDoor.Text
   Else
      mapDoorNameUp = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   uDoor = ""
   Call InvalidData
End Sub
Private Sub dDoor_LostFocus()
On Error GoTo ErrorHandler
   If Len(dDoor.Text) > 0 Then
      mapDoorNameDown = dDoor.Text
   Else
      mapDoorNameDown = ""
   End If
Exit Sub
ErrorHandler:
   dDoor = ""
   Call InvalidData
End Sub
Private Sub nExit_Click()
   If nExit.Value = 1 Then
      mapRoomNorth = True
   Else
      mapRoomNorth = False
   End If
End Sub

Private Sub nMove_Click()
   theRow = theRow - 1
   Call LoadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub Ridable_Click()
   If Ridable.Value = 1 Then
      mapRide = True
   Else
      mapRide = False
   End If
End Sub

Private Sub Roomname_LostFocus()
On Error GoTo ErrorHandler
   If Len(Roomname.Text) > 0 Then
      mapRoomName = Mid(Roomname.Text, 1, 50)
   Else
      mapRoomName = ""
   End If
   status.Caption = "Ok."
Exit Sub
ErrorHandler:
   Roomname = ""
   Call InvalidData
End Sub

Private Sub sExit_Click()
   If sExit.Value = 1 Then
      mapRoomSouth = True
   Else
      mapRoomSouth = False
   End If
End Sub

Private Sub sMove_Click()
   theRow = theRow + 1
   Call LoadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub Sun_Click()
   If Sun.Value = 1 Then
      mapSun = True
   Else
      mapSun = False
   End If
End Sub

Private Sub uExit_Click()
   If uExit.Value = 1 Then
      mapRoomUp = True
   Else
      mapRoomUp = False
   End If
End Sub

Private Sub wExit_Click()
   If wExit.Value = 1 Then
      mapRoomWest = True
   Else
      mapRoomWest = False
   End If
End Sub

Private Sub wMove_Click()
   theCol = theCol - 1
   Call LoadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub tcpServer_DataArrival(ByVal bytesTotal As Long)
'On Error GoTo ErrorHandler
   Dim OtherData As Boolean
   Dim strData As String
   Dim n As Long
   tcpServer.GetData strData
'   If checkString(strData, "MAPON") = True Then
'      Call setMapModeON
'      Exit Sub
'   End If
'   If checkString(strData, "MAPOFF") = True Then
'      Call setMapModeOFF
'      Exit Sub
'   End If
   If MAP_MODE = False Then
      If checkString(strData, "GET_IN_SYNC") = True Then
         If debug_mode = True Then Debug.Print vbCrLf & ">>> GET_IN_SYNC BY USER CALL" & vbCrLf
         fleeRadius = 10
         Call caseFleeHandler(currentRoomName, currentExits)
         Exit Sub
      End If
      If Out_Of_Sync = True Then
         tcpClient.SendData strData
         Call SYNC_FALSE
         Exit Sub
      End If
      If Len(strData) = 1 Then
         tcpClient.SendData strData
         Exit Sub
      End If
'THE USUAL
      theCommand = Split(strData, vbLf)
      For n = LBound(theCommand) To UBound(theCommand) - 1
         If Len(theCommand(n)) = 1 Then
            Select Case theCommand(n)
            Case "n"
               If checkTheMap(roomCount, N_map, N_exit, virtualRow - 1, virtualCol, "n") = True Then
                  If roomCount < limit Then tcpClient.SendData theCommand(n) & vbLf
               End If
            Case "e"
               If checkTheMap(roomCount, E_map, E_exit, virtualRow, virtualCol + 1, "e") = True Then
                  If roomCount < limit Then tcpClient.SendData theCommand(n) & vbLf
               End If
            Case "s"
               If checkTheMap(roomCount, S_map, S_exit, virtualRow + 1, virtualCol, "s") = True Then
                  If roomCount < limit Then tcpClient.SendData theCommand(n) & vbLf
               End If
            Case "w"
               If checkTheMap(roomCount, W_map, W_exit, virtualRow, virtualCol - 1, "w") = True Then
                  If roomCount < limit Then tcpClient.SendData theCommand(n) & vbLf
               End If
            Case "u"
               If checkTheMap(roomCount, U_map, U_exit, virtualRow, virtualCol, "u") = True Then
                  If roomCount < limit Then tcpClient.SendData theCommand(n) & vbLf
               End If
            Case "d"
               If checkTheMap(roomCount, D_map, D_exit, virtualRow, virtualCol, "d") = True Then
                  If roomCount < limit Then tcpClient.SendData theCommand(n) & vbLf
               End If
            Case Else
               tcpClient.SendData theCommand(n) & vbLf
            End Select
         Else
            tcpClient.SendData theCommand(n) & vbCrLf
         End If
      Next
   Else
      Call checkMapCommand(strData)
   End If
Exit Sub
ErrorHandler:
   BestEST.status = "CRITICAL ERROR! CONNECTION DOWN?"
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
'On Error GoTo ErrorHandler
   Dim n1 As String, n2 As String, n3 As String
   Dim strData As String
   tcpClient.GetData strData
   tcpServer.SendData strData
' MAP_MODE
   If MAP_THE_DATA = True Then
      Select Case MAP_THE_CASE
      Case 1
         n1 = InStr(strData, "[32")
         If n1 > 0 Then
            n2 = InStr(n1 + 5, strData, "[0m")
            If n2 > 0 Then
               Call zeroMap
               mapRoomName = Mid(strData, n1 + 5, n2 - (n1 + 5))
               mapDescription = Mid(strData, n2 + 6, 50)
               BestEST.Roomname = mapRoomName
               BestEST.Description = mapDescription
               MAP_THE_CASE = MAP_THE_CASE + 1
               tcpClient.SendData "exits" & vbLf
            End If
         Else
            tcpClient.SendData "examine" & vbLf
         End If
         Exit Sub
      Case 2
         If InStr(1, strData, "North") > 0 Then
            mapRoomNorth = True
            BestEST.nExit = 1
         End If
         If InStr(1, strData, "East") > 0 Then
            mapRoomEast = True
            BestEST.eExit = 1
         End If
         If InStr(1, strData, "South") > 0 Then
            BestEST.sExit = 1
            mapRoomSouth = True
         End If
         If InStr(1, strData, "West") > 0 Then
            BestEST.wExit = 1
            mapRoomWest = True
         End If
         If InStr(1, strData, "Up") > 0 Then
            BestEST.uExit = 1
            mapRoomUp = True
         End If
         If InStr(1, strData, "Down") > 0 Then
            BestEST.dExit = 1
            mapRoomDown = True
         End If
         MAP_THE_DATA = False
      End Select
      Exit Sub
   End If
' SPECIAL_CASES
   If Out_Of_Sync = False Then
      If AlasCount > 0 Then
         AlasCount = AlasCount - 1
         virtualRow = theRow
         virtualCol = theCol
      Else
         If checkString(strData, " seems to be closed.") = True Then
           If debug_mode = True Then Debug.Print " seems to be closed."
            Call Collision
            Exit Sub
         End If
         If checkString(strData, "Alas, you cannot go that way...") = True Then
           If debug_mode = True Then Debug.Print "Alas, you cannot go that way..."
            Call Collision
            Exit Sub
         End If
         If checkString(strData, "Your mount refuses to follow your orders!") = True Then
            Call Collision
            Exit Sub
         End If
         If checkString(strData, "doesn't want you riding") = True Then
            Call Collision
            Exit Sub
         End If
         If checkString(strData, "Oops! You cannot go there riding!") = True Then
            Call Collision
            Exit Sub
         End If
         If checkString(strData, "You need to swim to go there.") = True Then
            Call Collision
            Exit Sub
         End If
         If checkString(strData, "you need to climb to go there.") = True Then
            Call Collision
            Exit Sub
         End If
         If checkString(strData, " too exhausted.") = True Then
            Call Collision
            Exit Sub
         End If
         If checkString(strData, "Maybe you should get on your feet first?") = True Then
            Call resetBuffer
            Exit Sub
         End If
         If checkString(strData, "Nah... You feel too relaxed to do that..") = True Then
            Call resetBuffer
            Exit Sub
         End If
         If checkString(strData, "In your dreams, or what?") = True Then
            Call resetBuffer
            Exit Sub
         End If
      End If
   End If
' RUN_MODE
   n1 = InStr(strData, "[32")
   If n1 > 0 Then
      n2 = InStr(n1 + 5, strData, "[0m")
     If n2 > 0 Then
         n3 = InStr(n2 + 5, strData, "Exits:")
         If n3 > 0 Then
            currentRoomName = Mid(strData, n1 + 5, n2 - (n1 + 5))
            currentExits = Mid(strData, n3 + 6, 54)
            If debug_mode = True Then Debug.Print ">>> Found next room >" & Mid(strData, n1 + 5, n2 - (n1 + 5)) & "<"
            currentString = Mid(strData, 1, n1)
            If checkString(currentString, "You flee head over heels.") = True Then
               If debug_mode = True Then Debug.Print "*************   CASE FLEE, TRYING TO SYNC   ************"
               Call caseFleeHandler(currentRoomName, currentExits)
               Exit Sub
            End If
            If roomCount > 0 Then
               Call updateTheRoom
            End If
            Exit Sub
         End If
      End If
    End If
Exit Sub
ErrorHandler:
   BestEST.status = "CRITICAL ERROR! CONNECTION DOWN?"
End Sub

Private Sub tcpServer_ConnectionRequest _
(ByVal requestID As Long)
If WorldLoaded = False Then Exit Sub
 If tcpServer.State <> sckClosed Then _
   tcpServer.Close
   tcpServer.Accept requestID
   tcpClient.RemoteHost = "mume.pvv.org"
   tcpClient.RemotePort = 4242
   tcpClient.Connect
End Sub

