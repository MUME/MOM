VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1079
   ClientLeft      =   2834
   ClientTop       =   3484
   ClientWidth     =   1365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   636.685
   ScaleMode       =   0  'User
   ScaleWidth      =   1279.273
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   117
      TabIndex        =   2
      Top             =   585
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   234
      IMEMode         =   3  'DISABLE
      Left            =   117
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   351
      Width           =   1157
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sagapo"
      Height          =   247
      Index           =   1
      Left            =   117
      TabIndex        =   0
      Top             =   117
      Width           =   1079
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'JAANUS DEBUG
Private Sub cmdOK_Click()
   tsehhisalasona = txtPassword
   frmMap.Show
End Sub

