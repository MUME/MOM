VERSION 5.00
Begin VB.Form frmTools 
   BackColor       =   &H00000000&
   Caption         =   "Tools"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   186
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   307
   ScaleMode       =   0  'User
   ScaleWidth      =   393.685
   Begin VB.CheckBox dHidden 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   3120
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.CheckBox uHidden 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   2880
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.CheckBox wHidden 
      BackColor       =   &H00C0C0C0&
      Caption         =   "West"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   2640
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.CheckBox sHidden 
      BackColor       =   &H00C0C0C0&
      Caption         =   "South"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   2400
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.CheckBox eHidden 
      BackColor       =   &H00C0C0C0&
      Caption         =   "East"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   2160
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.CheckBox nHidden 
      BackColor       =   &H00C0C0C0&
      Caption         =   "North"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      MaskColor       =   &H000000FF&
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   1920
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.TextBox dDoor 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   870
      TabIndex        =   13
      Top             =   3105
      Width           =   1425
   End
   Begin VB.TextBox uDoor 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   870
      TabIndex        =   11
      Top             =   2865
      Width           =   1425
   End
   Begin VB.TextBox wDoor 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   870
      TabIndex        =   9
      Top             =   2625
      Width           =   1425
   End
   Begin VB.TextBox sDoor 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   870
      TabIndex        =   7
      Top             =   2385
      Width           =   1425
   End
   Begin VB.TextBox eDoor 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   870
      TabIndex        =   5
      Top             =   2145
      Width           =   1425
   End
   Begin VB.TextBox nDoor 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   870
      TabIndex        =   3
      Top             =   1905
      Width           =   1425
   End
   Begin VB.Frame Options 
      Appearance      =   0  'Flat
      BackColor       =   &H00A56E3A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      TabIndex        =   42
      Top             =   5640
      Width           =   3975
      Begin VB.CommandButton button_SaveWorld 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3480
         Width           =   735
      End
      Begin VB.CommandButton button_saveCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3480
         Width           =   735
      End
      Begin VB.CheckBox check_RoomSync 
         Appearance      =   0  'Flat
         BackColor       =   &H00A56E3A&
         Caption         =   "Validate each room"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MaskColor       =   &H000000FF&
         TabIndex        =   49
         Top             =   240
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CommandButton button_convert 
         Caption         =   "Convert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1185
         TabIndex        =   48
         Top             =   3420
         Width           =   1575
         Visible         =   0   'False
      End
      Begin VB.TextBox rowConvert 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         TabIndex        =   47
         Top             =   3060
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.TextBox colConvert 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2025
         TabIndex        =   46
         Top             =   3060
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.CheckBox check_AutoSync 
         Appearance      =   0  'Flat
         BackColor       =   &H00A56E3A&
         Caption         =   "Sync each room"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MaskColor       =   &H000000FF&
         TabIndex        =   45
         Top             =   555
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.TextBox newRadius 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1515
         TabIndex        =   44
         Top             =   1575
         Width           =   735
      End
      Begin VB.CommandButton button_radius 
         Caption         =   "new size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1095
         TabIndex        =   43
         Top             =   1890
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00A56E3A&
         Caption         =   "Row"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1305
         TabIndex        =   53
         Top             =   2820
         Width           =   495
         Visible         =   0   'False
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00A56E3A&
         Caption         =   "Col"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2145
         TabIndex        =   52
         Top             =   2820
         Width           =   495
         Visible         =   0   'False
      End
   End
   Begin VB.CommandButton button_CutMapData 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cut Map Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4020
      MaskColor       =   &H00000000&
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2610
      Width           =   1530
   End
   Begin VB.CommandButton button_GetMapData 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Get Map Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4020
      MaskColor       =   &H00000000&
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1530
   End
   Begin VB.CommandButton button_reset 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reset Room"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4020
      MaskColor       =   &H00000000&
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1530
   End
   Begin VB.CommandButton button_update 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Update Map"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4020
      MaskColor       =   &H00000000&
      TabIndex        =   14
      Top             =   1710
      Width           =   1530
   End
   Begin VB.CommandButton button_GetData 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Get Room Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3990
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   120
      Width           =   1530
   End
   Begin VB.TextBox GoToNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   180
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3810
      Width           =   975
   End
   Begin VB.TextBox nPortal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3135
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1905
      Width           =   750
   End
   Begin VB.TextBox ePortal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3135
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2145
      Width           =   750
   End
   Begin VB.TextBox sPortal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3135
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2385
      Width           =   750
   End
   Begin VB.TextBox wPortal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3135
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2625
      Width           =   750
   End
   Begin VB.TextBox uPortal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3135
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2865
      Width           =   750
   End
   Begin VB.TextBox dPortal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3135
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3105
      Width           =   750
   End
   Begin VB.CheckBox nExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "North"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Top             =   1905
      Value           =   1  'Checked
      Width           =   750
   End
   Begin VB.CheckBox eExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "East"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   2145
      Value           =   1  'Checked
      Width           =   750
   End
   Begin VB.CheckBox sExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "South"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   6
      Top             =   2385
      Value           =   1  'Checked
      Width           =   750
   End
   Begin VB.CheckBox wExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "West"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   2625
      Value           =   1  'Checked
      Width           =   750
   End
   Begin VB.CheckBox uExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   2865
      Width           =   750
   End
   Begin VB.CheckBox dExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   12
      Top             =   3105
      Width           =   750
   End
   Begin VB.TextBox Description 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   585
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   855
      Width           =   2490
   End
   Begin VB.TextBox Roomname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   585
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   615
      Width           =   2490
   End
   Begin VB.CheckBox Ridable 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ridable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1170
      MaskColor       =   &H000000FF&
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1275
      Width           =   1455
   End
   Begin VB.CheckBox Sun 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   90
      MaskColor       =   &H000000FF&
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1275
      Width           =   945
   End
   Begin VB.CheckBox Monster 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Monster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      MaskColor       =   &H000000FF&
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1275
      Width           =   1125
   End
   Begin VB.CommandButton button_goto 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      MaskColor       =   &H00000000&
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3810
      Width           =   645
   End
   Begin VB.CommandButton nMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4560
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmMapper.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3420
      Width           =   390
   End
   Begin VB.CommandButton sMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4560
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmMapper.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4140
      Width           =   390
   End
   Begin VB.CommandButton wMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4200
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmMapper.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3780
      Width           =   390
   End
   Begin VB.CommandButton eMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4920
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmMapper.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3780
      Width           =   390
   End
   Begin VB.CommandButton ssMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2520
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmMapper.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4170
      Width           =   370
   End
   Begin VB.CommandButton eeMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmMapper.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3810
      Width           =   370
   End
   Begin VB.CommandButton nnMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2520
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmMapper.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3450
      Width           =   370
   End
   Begin VB.CommandButton wwMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmMapper.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3810
      Width           =   370
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   69
      Top             =   1650
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Hidden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   68
      Top             =   1650
      Width           =   855
   End
   Begin VB.Label n_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   4800
      TabIndex        =   61
      Top             =   6240
      Width           =   1500
   End
   Begin VB.Label s_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   4800
      TabIndex        =   60
      Top             =   6720
      Width           =   1500
   End
   Begin VB.Label e_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   5640
      TabIndex        =   59
      Top             =   6480
      Width           =   1500
   End
   Begin VB.Label w_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   4080
      TabIndex        =   58
      Top             =   6480
      Width           =   1500
   End
   Begin VB.Label d_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   4800
      TabIndex        =   57
      Top             =   7080
      Width           =   1500
   End
   Begin VB.Label u_doorname 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   4800
      TabIndex        =   56
      Top             =   5880
      Width           =   1500
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Portal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3060
      TabIndex        =   55
      Top             =   1650
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Doorname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   870
      TabIndex        =   54
      Top             =   1650
      Width           =   1335
   End
   Begin VB.Image pictureTerrain 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   4350
      Stretch         =   -1  'True
      Top             =   780
      Width           =   840
   End
   Begin VB.Image special 
      Height          =   375
      Left            =   3180
      Picture         =   "frmMapper.frx":2210
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image plain 
      Height          =   375
      Left            =   1380
      Picture         =   "frmMapper.frx":2914
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image forest 
      Height          =   375
      Left            =   1740
      Picture         =   "frmMapper.frx":3018
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image road 
      Height          =   375
      Left            =   1020
      Picture         =   "frmMapper.frx":3764
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image swamp 
      Height          =   375
      Left            =   2100
      Picture         =   "frmMapper.frx":3E68
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image hill 
      Height          =   375
      Left            =   2460
      Picture         =   "frmMapper.frx":456C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image mountain 
      Height          =   375
      Left            =   2820
      Picture         =   "frmMapper.frx":4C70
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image water 
      Height          =   375
      Left            =   3540
      Picture         =   "frmMapper.frx":5374
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Label status 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ok."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   38
      Top             =   5280
      Width           =   4545
   End
   Begin VB.Label Label_room 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   37
      Top             =   615
      Width           =   495
   End
   Begin VB.Label label_desc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Desc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   36
      Top             =   855
      Width           =   495
   End
   Begin VB.Label label_terrain 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Selected terrain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4110
      TabIndex        =   35
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label label_col 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "C="
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   73
      Top             =   3420
      Width           =   315
   End
   Begin VB.Label label_row 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "R="
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   72
      Top             =   3630
      Width           =   315
   End
   Begin VB.Label col 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "COL"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3690
      TabIndex        =   71
      Top             =   3420
      Width           =   570
   End
   Begin VB.Label row 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "ROW"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3690
      TabIndex        =   70
      Top             =   3630
      Width           =   570
   End
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub button_radius_Click()
On Error GoTo errorhandler

   mapRadius = CLng(newRadius.Text)
   RoomSize = CLng(288 / ((2 * mapRadius) + 1))
   Call DrawMap

Exit Sub
errorhandler:
   mapRadius = 4
   RoomSize = 32
   Call DrawMap
End Sub

Private Sub button_saveCancel_Click()
   Options.Visible = False
End Sub

Private Sub button_GetData_Click()
   Call getRoomData
End Sub

Private Sub button_GetMapData_Click()
   Call zeroMap
   Call GetMapData
End Sub

Private Sub button_goto_Click()
   Call gotoArea(GoToNumber.Text)
End Sub

Private Sub button_reset_Click()
   arr(theRow, theCol) = 0
   arrDesc(theRow, theCol) = ""
   Call zeroMap
   Call loadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub button_save_Click()
   Options.Visible = True
End Sub

Private Sub button_SaveWorld_Click()
   Call saveWorld
   Options.Visible = False
End Sub

Private Sub button_update_Click()
   Call mapUpdate
End Sub

Private Sub Description_Change()
On Error GoTo errorhandler
'   If Len(Description.Text) > 0 Then
'      mapDescription = Mid(Description.Text, 1, 50)
'   Else
'      mapDescription = ""
'   End If
'   status.Caption = "Ok."
Exit Sub
errorhandler:
   Description = ""
   Call InvalidData
End Sub

Private Sub dExit_Click()
   If dExit.Value = 1 Then
      mapExitDown = True
   Else
      mapExitDown = False
   End If
End Sub

Private Sub dHidden_Click()
   If dHidden.Value = 1 Then
      mapHiddendoorDown = True
   Else
      mapHiddendoorDown = False
   End If
End Sub

Private Sub eeMove_Click()
   theCol = theCol + 10
   Call loadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub eExit_Click()
   If eExit.Value = 1 Then
      mapExitEast = True
   Else
      mapExitEast = False
   End If
End Sub

Private Sub eHidden_Click()
   If eHidden.Value = 1 Then
      mapHiddendoorEast = True
   Else
      mapHiddendoorEast = False
   End If
End Sub

Private Sub eMove_Click()
   theCol = theCol + 1
   Call loadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub nHidden_Click()
   If nHidden.Value = 1 Then
      mapHiddendoorNorth = True
   Else
      mapHiddendoorNorth = False
   End If
End Sub

Private Sub plain_Click()
   setMapTerrain ("plain")
End Sub

Private Sub forest_Click()
   setMapTerrain ("forest")
End Sub

Private Sub Form_Load()
   Me.Top = 0
End Sub

Private Sub hill_Click()
   setMapTerrain ("hill")
End Sub

Private Sub Monster_Click()
   If Monster.Value = 1 Then
      mapMonster = True
   Else
      mapMonster = False
   End If
End Sub

Private Sub mountain_Click()
   setMapTerrain ("mountain")
End Sub

Private Sub nDoor_LostFocus()
On Error GoTo errorhandler
   If Len(nDoor.Text) > 0 Then
      mapDoornameNorth = nDoor.Text
      nHidden.Visible = True
   Else
      mapDoornameNorth = ""
      nHidden.Visible = False
      nHidden.Value = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   nDoor = ""
   Call InvalidData
End Sub
Private Sub eDoor_LostFocus()
On Error GoTo errorhandler
   If Len(eDoor.Text) > 0 Then
      mapDoornameEast = eDoor.Text
      eHidden.Visible = True
   Else
      mapDoornameEast = ""
      eHidden.Visible = False
      eHidden.Value = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   eDoor = ""
   Call InvalidData
End Sub

Private Sub nnMove_Click()
   theRow = theRow - 10
   Call loadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub nPortal_LostFocus()
On Error GoTo errorhandler
   If Len(nPortal.Text) > 0 Then
      Dim tempData
      tempData = Split(nPortal.Text, ",")
      mapRowNorth = tempData(0)
      mapColNorth = tempData(1)
   Else
      mapRowNorth = 0
      mapColNorth = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   nPortal = ""
   Call InvalidData
End Sub
Private Sub ePortal_LostFocus()
On Error GoTo errorhandler
   If Len(ePortal.Text) > 0 Then
      Dim tempData
      tempData = Split(ePortal.Text, ",")
      mapRowEast = tempData(0)
      mapColEast = tempData(1)
   Else
      mapRowEast = 0
      mapColEast = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   ePortal = ""
   Call InvalidData
End Sub

Private Sub road_Click()
   setMapTerrain ("road")
End Sub

Private Sub sHidden_Click()
   If sHidden.Value = 1 Then
      mapHiddendoorSouth = True
   Else
      mapHiddendoorSouth = False
   End If
End Sub

Private Sub sPortal_LostFocus()
On Error GoTo errorhandler
   If Len(sPortal.Text) > 0 Then
      Dim tempData
      tempData = Split(sPortal.Text, ",")
      mapRowSouth = tempData(0)
      mapColSouth = tempData(1)
   Else
      mapRowSouth = 0
      mapColSouth = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   sPortal = ""
   Call InvalidData
End Sub

Private Sub special_Click()
   setMapTerrain ("special")
End Sub

Private Sub ssMove_Click()
   theRow = theRow + 10
   Call loadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub swamp_Click()
   setMapTerrain ("swamp")
End Sub

Private Sub uHidden_Click()
   If uHidden.Value = 1 Then
      mapHiddendoorUp = True
   Else
      mapHiddendoorUp = False
   End If
End Sub

Private Sub water_Click()
   setMapTerrain ("water")
End Sub

Private Sub wHidden_Click()
   If wHidden.Value = 1 Then
      mapHiddendoorSouth = True
   Else
      mapHiddendoorSouth = False
   End If
End Sub

Private Sub wPortal_LostFocus()
On Error GoTo errorhandler
   If Len(wPortal.Text) > 0 Then
      Dim tempData
      tempData = Split(wPortal.Text, ",")
      mapRowWest = tempData(0)
      mapColWest = tempData(1)
   Else
      mapRowWest = 0
      mapColWest = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   wPortal = ""
   Call InvalidData
End Sub
Private Sub uPortal_LostFocus()
On Error GoTo errorhandler
   If Len(uPortal.Text) > 0 Then
      Dim tempData
      tempData = Split(uPortal.Text, ",")
      mapRowUp = tempData(0)
      mapColUp = tempData(1)
   Else
      mapRowUp = 0
      mapColUp = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   uPortal = ""
   Call InvalidData
End Sub
Private Sub dPortal_LostFocus()
On Error GoTo errorhandler
   If Len(dPortal.Text) > 0 Then
      Dim tempData
      tempData = Split(dPortal.Text, ",")
      mapRowDown = tempData(0)
      mapColDown = tempData(1)
   Else
      mapRowDown = 0
      mapColDown = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   dPortal = ""
   Call InvalidData
End Sub
Private Sub sDoor_LostFocus()
On Error GoTo errorhandler
   If Len(sDoor.Text) > 0 Then
      mapDoornameSouth = sDoor.Text
      sHidden.Visible = True
   Else
      mapDoornameSouth = ""
      sHidden.Visible = False
      sHidden.Value = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   sDoor = ""
   Call InvalidData
End Sub
Private Sub wDoor_LostFocus()
On Error GoTo errorhandler
   If Len(wDoor.Text) > 0 Then
      mapDoornameWest = wDoor.Text
      wHidden.Visible = True
   Else
      mapDoornameWest = ""
      wHidden.Visible = False
      wHidden.Value = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   wDoor = ""
   Call InvalidData
End Sub
Private Sub uDoor_LostFocus()
On Error GoTo errorhandler
   If Len(uDoor.Text) > 0 Then
      mapDoornameUp = uDoor.Text
      uHidden.Visible = True
   Else
      mapDoornameUp = ""
      uHidden.Visible = False
      uHidden.Value = 0
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   uDoor = ""
   Call InvalidData
End Sub
Private Sub dDoor_LostFocus()
On Error GoTo errorhandler
   If Len(dDoor.Text) > 0 Then
      mapDoornameDown = dDoor.Text
      dHidden.Visible = True
   Else
      mapDoornameDown = ""
      dHidden.Visible = False
      dHidden.Value = 0
   End If
Exit Sub
errorhandler:
   dDoor = ""
   Call InvalidData
End Sub
Private Sub nExit_Click()
   If nExit.Value = 1 Then
      mapExitNorth = True
   Else
      mapExitNorth = False
   End If
End Sub

Private Sub nMove_Click()
   theRow = theRow - 1
   Call loadRoom(theRow, theCol)
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
On Error GoTo errorhandler
   If Len(Roomname.Text) > 0 Then
      mapRoomName = Mid(Roomname.Text, 1, 50)
   Else
      mapRoomName = ""
   End If
   status.Caption = "Ok."
Exit Sub
errorhandler:
   Roomname = ""
   Call InvalidData
End Sub

Private Sub sExit_Click()
   If sExit.Value = 1 Then
      mapExitSouth = True
   Else
      mapExitSouth = False
   End If
End Sub

Private Sub sMove_Click()
   theRow = theRow + 1
   Call loadRoom(theRow, theCol)
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
      mapExitUp = True
   Else
      mapExitUp = False
   End If
End Sub

Private Sub wExit_Click()
   If wExit.Value = 1 Then
      mapExitWest = True
   Else
      mapExitWest = False
   End If
End Sub

Private Sub wMove_Click()
   theCol = theCol - 1
   Call loadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub wwMove_Click()
   theCol = theCol - 10
   Call loadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub button_CutMapData_Click()
   Call zeroMap
   Call GetMapData
   arr(theRow, theCol) = 0
   arrDesc(theRow, theCol) = ""
   Call loadRoom(theRow, theCol)
   Call DrawMap
End Sub

Private Sub button_convert_Click()
   Call DBConvert(frmTools.rowConvert.Text, frmTools.colConvert.Text)
End Sub
