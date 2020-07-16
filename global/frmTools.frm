VERSION 5.00
Begin VB.Form frmTools 
   Appearance      =   0  'Flat
   BackColor       =   &H00C4DEC4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mapping Tools"
   ClientHeight    =   5550
   ClientLeft      =   1545
   ClientTop       =   330
   ClientWidth     =   4035
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   186
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTools.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox wVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   55
      TabStop         =   0   'False
      ToolTipText     =   "Portal is visible"
      Top             =   3525
      Width           =   195
   End
   Begin VB.CheckBox sVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1455
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Portal is visible"
      Top             =   3525
      Width           =   195
   End
   Begin VB.CheckBox eVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2775
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Portal is visible"
      Top             =   3525
      Width           =   195
   End
   Begin VB.TextBox ePortal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DED4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3495
      Width           =   990
   End
   Begin VB.TextBox sPortal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DED4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3480
      Width           =   990
   End
   Begin VB.TextBox wPortal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DED4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3480
      Width           =   990
   End
   Begin VB.Frame aFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BorderStyle     =   0  'None
      Caption         =   "000,000"
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
      Height          =   1365
      Left            =   120
      TabIndex        =   39
      Top             =   3750
      Width           =   840
      Begin VB.OptionButton aEast 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   45
         Top             =   405
         Width           =   780
      End
      Begin VB.OptionButton aSouth 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   44
         Top             =   600
         Width           =   780
      End
      Begin VB.OptionButton aWest 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   43
         Top             =   795
         Width           =   780
      End
      Begin VB.OptionButton aUp 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   42
         Top             =   990
         Width           =   780
      End
      Begin VB.OptionButton aDown 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   41
         Top             =   1185
         Width           =   780
      End
      Begin VB.OptionButton aNorth 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   40
         Top             =   210
         Value           =   -1  'True
         Width           =   780
      End
   End
   Begin VB.Frame typeFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BorderStyle     =   0  'None
      Caption         =   "TYPE"
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
      Height          =   1170
      Left            =   885
      TabIndex        =   28
      Top             =   3750
      Width           =   630
      Begin VB.OptionButton abbaPortal 
         BackColor       =   &H00C4DEC4&
         Caption         =   "<=>"
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
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   540
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton abPortal 
         BackColor       =   &H00C4DEC4&
         Caption         =   "-->"
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
         Height          =   225
         Left            =   60
         TabIndex        =   30
         Top             =   735
         Width           =   555
      End
      Begin VB.OptionButton baPortal 
         BackColor       =   &H00C4DEC4&
         Caption         =   "<--"
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
         Height          =   195
         Left            =   60
         TabIndex        =   29
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.CommandButton sMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   525
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmTools.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6030
      Width           =   315
   End
   Begin VB.CommandButton nMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   525
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmTools.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5730
      Width           =   315
   End
   Begin VB.TextBox finish 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   705
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "999,888"
      ToolTipText     =   "Area selection end"
      Top             =   5190
      Width           =   570
   End
   Begin VB.TextBox start 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "999,888"
      ToolTipText     =   "Area selection start"
      Top             =   5190
      Width           =   570
   End
   Begin VB.CheckBox dVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2775
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "Portal is visible"
      Top             =   2730
      Width           =   195
   End
   Begin VB.CheckBox nVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1455
      TabIndex        =   57
      TabStop         =   0   'False
      ToolTipText     =   "Portal is visible"
      Top             =   2730
      Width           =   195
   End
   Begin VB.CheckBox uVisible 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "Portal is visible"
      Top             =   2730
      Width           =   195
   End
   Begin VB.CheckBox uExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Room has an exit"
      Top             =   2220
      Width           =   195
   End
   Begin VB.CheckBox uHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Door is hidden"
      Top             =   2475
      Width           =   195
   End
   Begin VB.CheckBox dHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2775
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Door is hidden"
      Top             =   2475
      Width           =   195
   End
   Begin VB.CheckBox wHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Door is hidden"
      Top             =   3270
      Width           =   195
   End
   Begin VB.CheckBox sHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1455
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Door is hidden"
      Top             =   3270
      Width           =   195
   End
   Begin VB.CheckBox eHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2775
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Door is hidden"
      Top             =   3270
      Width           =   195
   End
   Begin VB.CheckBox nHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1455
      MaskColor       =   &H000000FF&
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Door is hidden"
      Top             =   2475
      Width           =   195
   End
   Begin VB.CheckBox nExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1455
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Room has an exit"
      Top             =   2220
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox eExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2775
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Room has an exit"
      Top             =   3015
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox sExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1455
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Room has an exit"
      Top             =   3015
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox wExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Room has an exit"
      Top             =   3015
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox dExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2775
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Room has an exit"
      Top             =   2220
      Width           =   195
   End
   Begin VB.TextBox dDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2415
      Width           =   990
   End
   Begin VB.TextBox dPortal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DED4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2685
      Width           =   990
   End
   Begin VB.TextBox eDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3225
      Width           =   990
   End
   Begin VB.TextBox nDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2415
      Width           =   990
   End
   Begin VB.TextBox uDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   375
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2415
      Width           =   990
   End
   Begin VB.TextBox GoToNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2355
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Enter row and column ie [100,300]"
      Top             =   4890
      Width           =   750
   End
   Begin VB.Frame bFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BorderStyle     =   0  'None
      Caption         =   "000,000"
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
      Height          =   1410
      Left            =   1425
      TabIndex        =   32
      Top             =   3750
      Width           =   810
      Begin VB.OptionButton bDown 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   38
         Top             =   1170
         Width           =   780
      End
      Begin VB.OptionButton bUp 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   37
         Top             =   975
         Width           =   780
      End
      Begin VB.OptionButton bWest 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   36
         Top             =   780
         Width           =   780
      End
      Begin VB.OptionButton bSouth 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   35
         Top             =   585
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton bEast 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   34
         Top             =   405
         Width           =   780
      End
      Begin VB.OptionButton bNorth 
         BackColor       =   &H00C4DEC4&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   33
         Top             =   210
         Width           =   780
      End
   End
   Begin VB.TextBox wDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3210
      Width           =   990
   End
   Begin VB.TextBox sDoor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3210
      Width           =   990
   End
   Begin VB.TextBox nPortal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DED4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2685
      Width           =   990
   End
   Begin VB.TextBox uPortal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DED4&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   375
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2685
      Width           =   990
   End
   Begin VB.CheckBox Ridable 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   195
      Left            =   2190
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1905
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox Sun 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   195
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1905
      UseMaskColor    =   -1  'True
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton wMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   210
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmTools.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6075
      Width           =   375
   End
   Begin VB.CommandButton eMove 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   735
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmTools.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6075
      Width           =   390
   End
   Begin VB.Image button_goto 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   2355
      Picture         =   "frmTools.frx":154A
      ToolTipText     =   "Jump to coordinates"
      Top             =   5205
      Width           =   750
   End
   Begin VB.Image button_createPortal 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1440
      Picture         =   "frmTools.frx":1FA4
      ToolTipText     =   "Create portal"
      Top             =   5205
      Width           =   750
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      Height          =   210
      Left            =   15
      Picture         =   "frmTools.frx":29FE
      ToolTipText     =   "Portal coordinates"
      Top             =   3525
      Width           =   120
   End
   Begin VB.Image monster3 
      Height          =   360
      Left            =   3555
      Picture         =   "frmTools.frx":2B90
      ToolTipText     =   "Group of monsters"
      Top             =   1815
      Width           =   420
   End
   Begin VB.Image monster2 
      Height          =   360
      Left            =   3060
      Picture         =   "frmTools.frx":33B2
      ToolTipText     =   "Hard monster"
      Top             =   1815
      Width           =   420
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   0
      Picture         =   "frmTools.frx":3BD4
      ToolTipText     =   "Doornames"
      Top             =   3270
      Width           =   135
   End
   Begin VB.Image aportal 
      Appearance      =   0  'Flat
      Height          =   210
      Left            =   15
      Picture         =   "frmTools.frx":3DD6
      ToolTipText     =   "Portal coordinates"
      Top             =   2730
      Width           =   120
   End
   Begin VB.Image adoor 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   15
      Picture         =   "frmTools.frx":3F68
      ToolTipText     =   "Doornames"
      Top             =   2475
      Width           =   135
   End
   Begin VB.Image button_GetData 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3210
      Picture         =   "frmTools.frx":416A
      ToolTipText     =   "Read room from MUD"
      Top             =   4905
      Width           =   750
   End
   Begin VB.Image button_GetMapData 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3210
      Picture         =   "frmTools.frx":4BC4
      ToolTipText     =   "Load room from map"
      Top             =   5205
      Width           =   750
   End
   Begin VB.Image button_reset 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3210
      Picture         =   "frmTools.frx":561E
      ToolTipText     =   "Delete room"
      Top             =   4605
      Width           =   750
   End
   Begin VB.Image button_CutMapData 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3210
      Picture         =   "frmTools.frx":6078
      ToolTipText     =   "Cut data and use Update to copy"
      Top             =   4305
      Width           =   750
   End
   Begin VB.Image button_update 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   3210
      Picture         =   "frmTools.frx":6AD2
      Top             =   3825
      Width           =   735
   End
   Begin VB.Image monster1 
      Height          =   360
      Left            =   2535
      Picture         =   "frmTools.frx":7BD8
      ToolTipText     =   "Weak monster"
      Top             =   1815
      Width           =   420
   End
   Begin VB.Image rideZone 
      Height          =   330
      Left            =   1410
      Picture         =   "frmTools.frx":83FA
      ToolTipText     =   "Ridable room"
      Top             =   1845
      Width           =   1035
   End
   Begin VB.Image sunZone 
      Height          =   330
      Left            =   120
      Picture         =   "frmTools.frx":961C
      ToolTipText     =   "Sunny room"
      Top             =   1845
      Width           =   1215
   End
   Begin VB.Label Roomname 
      Alignment       =   2  'Center
      BackColor       =   &H00C4DEC4&
      Caption         =   "roomname"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   60
      TabIndex        =   46
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image flagNone 
      Height          =   330
      Left            =   3600
      Picture         =   "frmTools.frx":AB56
      ToolTipText     =   "Reset flags"
      Top             =   1485
      Width           =   330
   End
   Begin VB.Image flagQuest 
      Height          =   330
      Left            =   2535
      Picture         =   "frmTools.frx":B170
      ToolTipText     =   "Quest"
      Top             =   1485
      Width           =   330
   End
   Begin VB.Image flagItem 
      Height          =   330
      Left            =   585
      Picture         =   "frmTools.frx":B78A
      ToolTipText     =   "Items or mounts"
      Top             =   1485
      Width           =   330
   End
   Begin VB.Image flagHerb 
      Height          =   330
      Left            =   975
      Picture         =   "frmTools.frx":BDA4
      ToolTipText     =   "Herbs or food"
      Top             =   1485
      Width           =   330
   End
   Begin VB.Image flagWater 
      Height          =   330
      Left            =   195
      Picture         =   "frmTools.frx":C3BE
      ToolTipText     =   "Spring"
      Top             =   1485
      Width           =   330
   End
   Begin VB.Image flagTreasury 
      Height          =   330
      Left            =   1365
      Picture         =   "frmTools.frx":C9D8
      ToolTipText     =   "Treasury"
      Top             =   1485
      Width           =   330
   End
   Begin VB.Image flagKey 
      Height          =   330
      Left            =   1755
      Picture         =   "frmTools.frx":CFF2
      ToolTipText     =   "Key or Lock"
      Top             =   1485
      Width           =   330
   End
   Begin VB.Image flagMagic 
      Height          =   330
      Left            =   2145
      Picture         =   "frmTools.frx":D60C
      ToolTipText     =   "Special room"
      Top             =   1485
      Width           =   330
   End
   Begin VB.Image dungeon 
      Height          =   600
      Left            =   2355
      Picture         =   "frmTools.frx":DC26
      Top             =   855
      Width           =   525
   End
   Begin VB.Image city 
      Height          =   600
      Left            =   120
      Picture         =   "frmTools.frx":ED48
      Top             =   855
      Width           =   510
   End
   Begin VB.Image shop 
      Height          =   600
      Left            =   675
      Picture         =   "frmTools.frx":FDCA
      Top             =   855
      Width           =   510
   End
   Begin VB.Image inn 
      Appearance      =   0  'Flat
      Height          =   600
      Left            =   1785
      Picture         =   "frmTools.frx":10E4C
      Top             =   855
      Width           =   510
   End
   Begin VB.Image guild 
      Height          =   600
      Left            =   1230
      Picture         =   "frmTools.frx":11ECE
      Top             =   855
      Width           =   510
   End
   Begin VB.Image mountain 
      Height          =   600
      Left            =   2310
      Picture         =   "frmTools.frx":12F50
      Top             =   240
      Width           =   570
   End
   Begin VB.Image plain 
      Height          =   630
      Left            =   120
      Picture         =   "frmTools.frx":141B2
      Top             =   240
      Width           =   510
   End
   Begin VB.Image forest 
      Height          =   630
      Left            =   675
      Picture         =   "frmTools.frx":15304
      Top             =   240
      Width           =   510
   End
   Begin VB.Image road 
      Height          =   630
      Left            =   3015
      Picture         =   "frmTools.frx":16456
      ToolTipText     =   "Create road"
      Top             =   1095
      Width           =   510
   End
   Begin VB.Image swamp 
      Height          =   630
      Left            =   1230
      Picture         =   "frmTools.frx":175A8
      Top             =   240
      Width           =   510
   End
   Begin VB.Image hill 
      Height          =   630
      Left            =   1785
      Picture         =   "frmTools.frx":186FA
      Top             =   240
      Width           =   510
   End
   Begin VB.Image underground 
      Height          =   630
      Left            =   2910
      Picture         =   "frmTools.frx":1984C
      Top             =   240
      Width           =   510
   End
   Begin VB.Image water 
      Height          =   630
      Left            =   3465
      Picture         =   "frmTools.frx":1A99E
      Top             =   240
      Width           =   510
   End
   Begin VB.Label coordinates 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C4DEC4&
      Caption         =   "(123,456)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3315
      TabIndex        =   24
      Top             =   15
      Width           =   705
   End
   Begin VB.Image dZone 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   2775
      Picture         =   "frmTools.frx":1BAF0
      ToolTipText     =   "Room has an exit"
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Image nZone 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   1455
      Picture         =   "frmTools.frx":1C97E
      ToolTipText     =   "Room has an exit"
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Image uZone 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   150
      Picture         =   "frmTools.frx":1D80C
      ToolTipText     =   "Room has an exit"
      Top             =   2205
      Width           =   1200
   End
   Begin VB.Image eZone 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   2775
      Picture         =   "frmTools.frx":1E65E
      ToolTipText     =   "Room has an exit"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image sZone 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   1455
      Picture         =   "frmTools.frx":1F4EC
      ToolTipText     =   "Room has an exit"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image wZone 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   135
      Picture         =   "frmTools.frx":2037A
      ToolTipText     =   "Room has an exit"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image flagMudlle 
      Height          =   330
      Left            =   1455
      Picture         =   "frmTools.frx":21208
      Top             =   5805
      Width           =   330
      Visible         =   0   'False
   End
   Begin VB.Image flagQuestion 
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   1605
      Picture         =   "frmTools.frx":21822
      Top             =   5955
      Width           =   390
      Visible         =   0   'False
   End
   Begin VB.Image bridge 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   1335
      Picture         =   "frmTools.frx":21E3C
      Top             =   5655
      Width           =   570
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub button_GetData_Click()
If DEBUGMODE = False Then On Error GoTo errorhandler
   Call getRoomData
Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "frmTools GetData"
   writeError (errorModule)
   WorldLoaded = False
End Sub

Private Sub button_GetMapData_Click()
   Call GetMapData
End Sub

Private Sub button_goto_Click()
   Call gotoArea(GoToNumber.text)
End Sub

Private Sub button_update__Click()
If DEBUGMODE = False Then On Error GoTo errorhandler
   Call mapUpdate
Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "button_Update_Click"
   writeError (errorModule)
End Sub

Private Sub button_update_Click()
   Call mapUpdate
End Sub

Private Sub city_Click()
   setMapTerrain ("city")
      Call mapUpdate
End Sub

Private Sub dExit_Click()
   If dExit.value = 1 Then
      mapExitDown = True
   Else
      mapExitDown = False
   End If
End Sub

Private Sub dHidden_Click()
   If dHidden.value = 1 Then
      mapHiddendoorDown = True
   Else
      mapHiddendoorDown = False
   End If
End Sub

Private Sub dZone_Click()
   If dExit.value = 1 Then
      dExit.value = 0
      mapExitDown = False
   Else
      dExit.value = 1
      mapExitDown = True
   End If
   Call mapUpdate
End Sub

Private Sub dungeon_Click()
   setMapTerrain ("dungeon")
   Call mapUpdate
End Sub

Private Sub eExit_Click()
   If eExit.value = 1 Then
      mapExitEast = True
   Else
      mapExitEast = False
   End If
End Sub

Private Sub eHidden_Click()
   If eHidden.value = 1 Then
      mapHiddendoorEast = True
   Else
      mapHiddendoorEast = False
   End If
End Sub

Private Sub eMove_Click()
   theCOL = theCOL + 1
   Call loadRoom(theROW, theCOL)
   Call DrawMap
End Sub

Private Sub eZone_Click()
   If eExit.value = 1 Then
      eExit.value = 0
      mapExitEast = False
   Else
      eExit.value = 1
      mapExitEast = True
   End If
   Call mapUpdate
End Sub

Private Sub flagHerb_Click()
   setMapFlag ("herb")
      Call mapUpdate
End Sub

Private Sub flagItem_Click()
   setMapFlag ("item")
      Call mapUpdate
End Sub

Private Sub flagKey_Click()
   setMapFlag ("key")
      Call mapUpdate
End Sub

Private Sub flagMagic_Click()
   setMapFlag ("magic")
      Call mapUpdate
End Sub

Private Sub flagMudlle_Click()
   If theLEVEL = 0 Then
      frmTools.flagMudlle.BorderStyle = 1
      theLEVEL = 1
   Else
      theLEVEL = 0
      frmTools.flagMudlle.BorderStyle = 0
   End If
   
   Call DrawMap
End Sub

Private Sub flagNone_Click()
   setMapFlag ("none")
   mapMonsterValue = 0
   mapRoad = 0
   Call mapUpdate
End Sub

Private Sub flagQuest_Click()
   setMapFlag ("quest")
   Call mapUpdate
End Sub

Private Sub flagQuestion_Click()
   setMapFlag ("question")
   Call mapUpdate
End Sub

Private Sub flagTreasury_Click()
   setMapFlag ("treasury")
   Call mapUpdate
End Sub

Private Sub flagWater_Click()
   setMapFlag ("water")
   Call mapUpdate
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'frmMap.mnuTools.Checked = False
End Sub

Private Sub guild_Click()
   setMapTerrain ("guild")
   Call mapUpdate
End Sub

Private Sub Image3_Click()

End Sub

Private Sub Image4_Click()

End Sub

Private Sub inn_Click()
   setMapTerrain ("inn")
   Call mapUpdate
End Sub

Private Sub monster1_Click()
   mapMonster = True
   mapMonsterValue = MONSTER_EASY
   Call mapUpdate
End Sub

Private Sub monster2_Click()
   mapMonster = True
   mapMonsterValue = MONSTER_HARD
   Call mapUpdate
End Sub

Private Sub monster3_Click()
   mapMonster = True
   mapMonsterValue = MONSTER_GROUP
   Call mapUpdate
End Sub

Private Sub nHidden_Click()
   If nHidden.value = 1 Then
      mapHiddendoorNorth = True
   Else
      mapHiddendoorNorth = False
   End If
End Sub

Private Sub nZone_Click()
   If nExit.value = 1 Then
      nExit.value = 0
      mapExitNorth = False
   Else
      nExit.value = 1
      mapExitNorth = True
   End If
   Call mapUpdate
End Sub

Private Sub plain_Click()
   setMapTerrain ("plain")
   Call mapUpdate
End Sub

Private Sub forest_Click()
   setMapTerrain ("forest")
      Call mapUpdate
End Sub

Private Sub Form_Load()
   Me.Top = 0
End Sub

Private Sub hill_Click()
   setMapTerrain ("hill")
      Call mapUpdate
End Sub

Private Sub Monster_Click()
'   mapMonster = True
'   If Monster.value = 1 Then
'      Monster.value = 0
'      mapMonsterValue = 0
'   Else
'      Monster.value = 1
'      mapMonsterValue = MONSTER_EASY
'   End If
'   Call mapUpdate
End Sub

Private Sub rideZone_Click()
   If Ridable.value = 1 Then
      Ridable.value = 0
   Else
      Ridable.value = 1
   End If
   Call Ridable_Click
   Call mapUpdate
End Sub


Private Sub shop_Click()
   setMapTerrain ("shop")
      Call mapUpdate
End Sub

Private Sub sZone_Click()
   If sExit.value = 1 Then
      sExit.value = 0
      mapExitSouth = False
   Else
      sExit.value = 1
      mapExitSouth = True
   End If
   Call mapUpdate
End Sub

Private Sub sunZone_Click()
   If Sun.value = 1 Then
      Sun.value = 0
   Else
      Sun.value = 1
   End If
   Call Sun_Click
   Call mapUpdate
End Sub

Private Sub Underground_Click()
   setMapTerrain ("underground")
      Call mapUpdate
End Sub

Private Sub nDoor_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(nDoor.text) <> 0 Then
      mapDoornameNorth = nDoor.text
      nHidden.Visible = True
   Else
      mapDoornameNorth = vbNullString
      nHidden.Visible = False
      nHidden.value = 0
   End If
Exit Sub
errorhandler:
   nDoor = vbNullString
   Call InvalidData
End Sub
Private Sub eDoor_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(eDoor.text) <> 0 Then
      mapDoornameEast = eDoor.text
      eHidden.Visible = True
   Else
      mapDoornameEast = vbNullString
      eHidden.Visible = False
      eHidden.value = 0
   End If
Exit Sub
errorhandler:
   eDoor = vbNullString
   Call InvalidData
End Sub

Private Sub nPortal_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(nPortal.text) <> 0 Then
      Dim tempdata
      tempdata = Split(nPortal.text, ",", , vbBinaryCompare)
      mapRowNorth = tempdata(0)
      mapColNorth = tempdata(1)
   Else
      mapRowNorth = 0
      mapColNorth = 0
   End If
Exit Sub
errorhandler:
   nPortal = vbNullString
   Call InvalidData
End Sub
Private Sub ePortal_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(ePortal.text) <> 0 Then
      Dim tempdata
      tempdata = Split(ePortal.text, ",", , vbBinaryCompare)
      mapRowEast = tempdata(0)
      mapColEast = tempdata(1)
   Else
      mapRowEast = 0
      mapColEast = 0
   End If
Exit Sub
errorhandler:
   ePortal = vbNullString
   Call InvalidData
End Sub

Private Sub road_Click()
   mapRoad = ISROAD
   Call mapUpdate
End Sub

Private Sub sHidden_Click()
   If sHidden.value = 1 Then
      mapHiddendoorSouth = True
   Else
      mapHiddendoorSouth = False
   End If
End Sub

Private Sub sPortal_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(sPortal.text) <> 0 Then
      Dim tempdata
      tempdata = Split(sPortal.text, ",", , vbBinaryCompare)
      mapRowSouth = tempdata(0)
      mapColSouth = tempdata(1)
   Else
      mapRowSouth = 0
      mapColSouth = 0
   End If
Exit Sub
errorhandler:
   sPortal = vbNullString
   Call InvalidData
End Sub

Private Sub mountain_Click()
   setMapTerrain ("mountain")
   Call mapUpdate
End Sub

Private Sub swamp_Click()
   setMapTerrain ("swamp")
      Call mapUpdate
End Sub

Private Sub uHidden_Click()
   If uHidden.value = 1 Then
      mapHiddendoorUp = True
   Else
      mapHiddendoorUp = False
   End If
End Sub

Private Sub uZone_Click()
   If uExit.value = 1 Then
      uExit.value = 0
      mapExitUp = False
   Else
      uExit.value = 1
      mapExitUp = True
   End If
   Call mapUpdate
End Sub

Private Sub water_Click()
   setMapTerrain ("water")
      Call mapUpdate
End Sub

Private Sub wHidden_Click()
   If wHidden.value = 1 Then
      mapHiddendoorWest = True
   Else
      mapHiddendoorWest = False
   End If
End Sub

Private Sub wPortal_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(wPortal.text) <> 0 Then
      Dim tempdata
      tempdata = Split(wPortal.text, ",", , vbBinaryCompare)
      mapRowWest = tempdata(0)
      mapColWest = tempdata(1)
   Else
      mapRowWest = 0
      mapColWest = 0
   End If
Exit Sub
errorhandler:
   wPortal = vbNullString
   Call InvalidData
End Sub
Private Sub uPortal_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(uPortal.text) <> 0 Then
      Dim tempdata
      tempdata = Split(uPortal.text, ",", , vbBinaryCompare)
      mapRowUp = tempdata(0)
      mapColUp = tempdata(1)
   Else
      mapRowUp = 0
      mapColUp = 0
   End If
Exit Sub
errorhandler:
   uPortal = vbNullString
   Call InvalidData
End Sub
Private Sub dPortal_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(dPortal.text) <> 0 Then
      Dim tempdata
      tempdata = Split(dPortal.text, ",", , vbBinaryCompare)
      mapRowDown = tempdata(0)
      mapColDown = tempdata(1)
   Else
      mapRowDown = 0
      mapColDown = 0
   End If
Exit Sub
errorhandler:
   dPortal = vbNullString
   Call InvalidData
End Sub
Private Sub sDoor_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(sDoor.text) <> 0 Then
      mapDoornameSouth = sDoor.text
      sHidden.Visible = True
   Else
      mapDoornameSouth = vbNullString
      sHidden.Visible = False
      sHidden.value = 0
   End If
Exit Sub
errorhandler:
   sDoor = vbNullString
   Call InvalidData
End Sub
Private Sub wDoor_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(wDoor.text) <> 0 Then
      mapDoornameWest = wDoor.text
      wHidden.Visible = True
   Else
      mapDoornameWest = vbNullString
      wHidden.Visible = False
      wHidden.value = 0
   End If
Exit Sub
errorhandler:
   wDoor = vbNullString
   Call InvalidData
End Sub
Private Sub uDoor_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(uDoor.text) <> 0 Then
      mapDoornameUp = uDoor.text
      uHidden.Visible = True
   Else
      mapDoornameUp = vbNullString
      uHidden.Visible = False
      uHidden.value = 0
   End If
Exit Sub
errorhandler:
   uDoor = vbNullString
   Call InvalidData
End Sub
Private Sub dDoor_Validate(Cancel As Boolean)
If DEBUGMODE = False Then On Error GoTo errorhandler
   If LenB(dDoor.text) <> 0 Then
      mapDoornameDown = dDoor.text
      dHidden.Visible = True
   Else
      mapDoornameDown = vbNullString
      dHidden.Visible = False
      dHidden.value = 0
   End If
Exit Sub
errorhandler:
   dDoor = vbNullString
   Call InvalidData
End Sub
Private Sub nExit_Click()
   If nExit.value = 1 Then
      mapExitNorth = True
   Else
      mapExitNorth = False
   End If
End Sub

Private Sub nMove_Click()
   theROW = theROW - 1
   Call loadRoom(theROW, theCOL)
   Call DrawMap
End Sub

Public Sub Ridable_Click()
   If Ridable.value = 1 Then
      mapRide = True
   Else
      mapRide = False
   End If
End Sub

Private Sub sExit_Click()
   If sExit.value = 1 Then
      mapExitSouth = True
   Else
      mapExitSouth = False
   End If
End Sub

Private Sub sMove_Click()
   theROW = theROW + 1
   Call loadRoom(theROW, theCOL)
   Call DrawMap
End Sub

Public Sub Sun_Click()
   If Sun.value = 1 Then
      mapSun = True
   Else
      mapSun = False
   End If
End Sub

Private Sub uExit_Click()
   If uExit.value = 1 Then
      mapExitUp = True
   Else
      mapExitUp = False
   End If
End Sub

Private Sub wExit_Click()
   If wExit.value = 1 Then
      mapExitWest = True
   Else
      mapExitWest = False
   End If
End Sub

Private Sub wMove_Click()
   theCOL = theCOL - 1
   Call loadRoom(theROW, theCOL)
   Call DrawMap
End Sub
Private Sub button_CutMapData_Click()
   Call GetMapData
   aData(getIndex(theROW, theCOL), cDATA) = vbNullString
   aWorld(theROW, theCOL, theLEVEL) = 0
   Call loadRoom(theROW, theCOL)
   Call DrawMap
End Sub

Public Sub button_reset_Click()
   Call clearArraySlot(theROW, theCOL)
   Call DrawMap
End Sub

Private Sub button_createPortal_Click()
If DEBUGMODE = False Then On Error GoTo errorhandler
Dim wasRow As Integer
Dim wasCol As Integer
If MappingMode And selectionStartRow > 0 And selectionStartCol > 0 And selectionEndRow > 0 And selectionEndCol > 0 Then
   wasRow = theROW
   wasCol = theCOL
   
   'Create portal inside room A, the first selected room
   If frmTools.abPortal Or frmTools.abbaPortal Then
      theROW = selectionStartRow
      theCOL = selectionStartCol
      'loadroom
      If isValid(theROW, theCOL) Then
         Call loadRoom(theROW, theCOL)
         Call GetMapData
      Else
         frmTools.Caption = "Create Portal - Invalid selection!"
         Exit Sub
      End If
      
      Select Case True
      Case frmTools.aNorth
         frmTools.nVisible = 1
         mapRowNorth = selectionEndRow
         mapColNorth = selectionEndCol
      Case frmTools.aEast
         frmTools.eVisible = 1
         mapRowEast = selectionEndRow
         mapColEast = selectionEndCol
      Case frmTools.aSouth
         frmTools.sVisible = 1
         mapRowSouth = selectionEndRow
         mapColSouth = selectionEndCol
      Case frmTools.aWest
         frmTools.wVisible = 1
         mapRowWest = selectionEndRow
         mapColWest = selectionEndCol
      Case frmTools.aUp
         frmTools.uVisible = 1
         mapRowUp = selectionEndRow
         mapColUp = selectionEndCol
      Case frmTools.aDown
         frmTools.dVisible = 1
         mapRowDown = selectionEndRow
         mapColDown = selectionEndCol
      End Select
      Call mapUpdate
   End If
   
   'create portal in room B, the 2nd room
   If frmTools.baPortal Or frmTools.abbaPortal Then
      
      theROW = selectionEndRow
      theCOL = selectionEndCol
      If isValid(theROW, theCOL) Then
         Call loadRoom(theROW, theCOL)
         Call GetMapData
      Else
         frmTools.Caption = "Create Portal - Invalid selectionion!"
         Exit Sub
      End If
      Select Case True
      Case frmTools.bNorth
         frmTools.nVisible.value = 1
         mapRowNorth = selectionStartRow
         mapColNorth = selectionStartCol
      Case frmTools.bEast
         frmTools.eVisible = 1
         mapRowEast = selectionStartRow
         mapColEast = selectionStartCol
      Case frmTools.bSouth
         frmTools.sVisible = 1
         mapRowSouth = selectionStartRow
         mapColSouth = selectionStartCol
      Case frmTools.bWest
         frmTools.wVisible = 1
         mapRowWest = selectionStartRow
         mapColWest = selectionStartCol
      Case frmTools.bUp
         frmTools.uVisible = 1
         mapRowUp = selectionStartRow
         mapColUp = selectionStartCol
      Case frmTools.bDown
         frmTools.dVisible = 1
         mapRowDown = selectionStartRow
         mapColDown = selectionStartCol
      End Select
      Call mapUpdate
   End If
   theROW = wasRow
   theCOL = wasCol
   
   
   If isValid(theROW, theCOL) Then
      Call loadRoom(theROW, theCOL)
      Call GetMapData
   End If
   Call DrawMap
Else
   frmTools.Caption = "Please select area!"
End If
Exit Sub
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "frmtools CreatePortal"
   writeError (errorModule)
   WorldLoaded = False
End Sub

Private Sub wZone_Click()
   If wExit.value = 1 Then
      wExit.value = 0
      mapExitWest = False
   Else
      wExit.value = 1
      mapExitWest = True
   End If
   Call mapUpdate
End Sub
