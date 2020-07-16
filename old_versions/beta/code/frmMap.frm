VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   616
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   599
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox map 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   135
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   288
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4320
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuMappingMode 
         Caption         =   "Mapping Mode"
      End
      Begin VB.Menu mnuRadius 
         Caption         =   "Map radius"
         Begin VB.Menu mnuRadius3 
            Caption         =   "small"
         End
         Begin VB.Menu mnuRadius6 
            Caption         =   "medium"
         End
         Begin VB.Menu mnuRadius9 
            Caption         =   "large"
         End
         Begin VB.Menu mnuRadius18 
            Caption         =   "very large"
         End
      End
      Begin VB.Menu mnuWorldLocate 
         Caption         =   "World Locate"
      End
      Begin VB.Menu mnuAreaLocate 
         Caption         =   "Area Locate"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
