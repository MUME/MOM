VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "Labinaator"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox myUrl 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "http://infurni.ee/6061/mume.php"
      Top             =   0
      Width           =   3975
      Visible         =   0   'False
   End
   Begin VB.TextBox mySleep 
      Height          =   285
      Left            =   3240
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "1000"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox myCharacter 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   3
      Text            =   "nimi"
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":1F72
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "ms"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub SleepAPI Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)
    
Private WithEvents objMomToDB As momtodb.cAsync
Attribute objMomToDB.VB_VarHelpID = -1

'' Register a type library
'Sub RegisterTypeLib(ByVal TypeLibFile As String)
'    Dim TLI As New TLIApplication
'    ' raises an error if unable to register
'    ' (e.g. file not found or not a TLB)
'    TLI.TypeLibInfoFromFile(TypeLibFile).Register
'End Sub
'
'' Unregister a type library
'Sub UnregisterTypeLib(ByVal TypeLibFile As String)
'    Dim TLI As New TLIApplication
'    ' raises an error if unable to unregister
'    TLI.TypeLibInfoFromFile(TypeLibFile).UnRegister
'End Sub



Private Sub btnStart_Click()
   If LenB(myCharacter.Text) > 20 Or LenB(myCharacter.Text) < 1 Then
      myCharacter.Text = "name"
      Exit Sub
   End If
   If Not IsNumeric(mySleep.Text) Then
      mySleep.Text = 1000
      Exit Sub
   End If
   If CInt(mySleep.Text) < 0 Or CInt(mySleep.Text) > 9999 Then mySleep.Text = 1000
   btnStart.Enabled = False
   myCharacter.Enabled = False
   mySleep.Enabled = False
   btnStop.Enabled = True
   
   'On Error Resume Next
   Set objMomToDB = New momtodb.cAsync
   Call objMomToDB.Start(myCharacter.Text, myUrl.Text, mySleep.Text)
   If Err.Number <> 0 Then Call btnStop_Click
End Sub

Private Sub Form_Load()
Dim errno As Long
   'Call UnregisterTypeLib("C:\RUNNABLE.TLB")
   'Call RegisterTypeLib(App.Path & "\RUNNABLE.TLB")
   
   errno = Err.Number
   On Error Resume Next
   If errno <> 0 Then
      On Error GoTo 0
      MsgBox ("typelib error: " & Err.Description)
   Else
      Dim a As WinHttp.WinHttpRequest
      Set a = New WinHttp.WinHttpRequest
      a.SetTimeouts 100, 100, 100, 100
      Set a = Nothing
      If Err.Number <> 0 Then
         On Error GoTo 0
         MsgBox ("winhttp error: " & Err.Description)
      Else
         Dim momtodb As momtodb.cAsync
         Set momtodb = New momtodb.cAsync
         Set momtodb = Nothing
         If Err.Number <> 0 Then
            MsgBox ("momtodb error: " & Err.Description)
         Else
            MsgBox ("k6ik on OK")
         End If
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call btnStop_Click
End Sub

Private Sub objMomToDB_Complete()
   Form1.Text1.Text = Time() & vbCrLf & Replace(objMomToDB.Result, ";", vbCrLf, , , vbBinaryCompare)
   If btnStop.Enabled Then Call btnStart_Click
End Sub

Private Sub objMomToDB_Cancelled()
   Form1.Text1.Text = Time() & vbCrLf & Replace(objMomToDB.Result, ";", vbCrLf, , , vbBinaryCompare)
   If btnStop.Enabled Then Call btnStart_Click
End Sub


Private Sub btnStop_Click()
   'On Error Resume Next
   SleepAPI CInt(mySleep.Text + 1000)
   btnStop.Enabled = False
   btnStart.Enabled = True
   myCharacter.Enabled = True
   mySleep.Enabled = True
End Sub
