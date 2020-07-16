VERSION 5.00
Begin VB.Form main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUME Online Map Setup"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSetup.frx":1272
   ScaleHeight     =   6945
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Exit_button 
      BackColor       =   &H00000000&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   360
      MaskColor       =   &H00000000&
      TabIndex        =   9
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   1065
      Visible         =   0   'False
   End
   Begin VB.TextBox letter 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      ForeColor       =   &H0080C0FF&
      Height          =   1890
      Left            =   2895
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   4845
      Width           =   4830
      Visible         =   0   'False
   End
   Begin VB.TextBox Productname 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      ForeColor       =   &H0080C0FF&
      Height          =   285
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "..."
      Top             =   105
      Width           =   3495
      Visible         =   0   'False
   End
   Begin VB.TextBox Systemroot 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      ForeColor       =   &H0080C0FF&
      Height          =   285
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "..."
      Top             =   495
      Width           =   3495
      Visible         =   0   'False
   End
   Begin VB.TextBox Registeredowner 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      ForeColor       =   &H0080C0FF&
      Height          =   285
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "..."
      Top             =   885
      Width           =   3495
      Visible         =   0   'False
   End
   Begin VB.CommandButton Install_button 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   8340
      MaskColor       =   &H00000000&
      Picture         =   "frmSetup.frx":255F9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4965
      UseMaskColor    =   -1  'True
      Width           =   1905
   End
   Begin VB.Label productname_label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Productname: "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   315
      Left            =   30
      TabIndex        =   3
      Top             =   105
      Width           =   1395
      Visible         =   0   'False
   End
   Begin VB.Label systemroot_label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Systemroot: "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   495
      Width           =   1395
      Visible         =   0   'False
   End
   Begin VB.Label RegisteredOwner_label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Registered:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   885
      Width           =   1395
      Visible         =   0   'False
   End
   Begin VB.Label info 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00C0FFFF&
      Height          =   4410
      Left            =   3180
      TabIndex        =   7
      Top             =   1455
      Width           =   4275
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const w98_scriptingHTTP = "http://download.microsoft.com/download/4/c/9/4c9e63f1-617f-4c6d-8faf-c2868f670c1c/scr56en.exe"
Const w2K_scriptingHTTP = "http://download.microsoft.com/download/2/8/a/28a5a346-1be1-4049-b554-3bc5f3174353/scripten.exe"
'filesystem
Public fso As New Scripting.FileSystemObject
'internet
Public tcp As New MSWinsockLib.Winsock
'encryption
Public md5 As New MD5DLL.Crypt   'hash
Public md5Val As String
Public w98 As Boolean
Public wNT As Boolean

Private Sub Exit_button_Click()
   On Error Resume Next
   End
End Sub

Private Sub Form_Load()
   main.info.Caption = "Hello!" & vbCrLf & "This setup will check the registered components" & vbCrLf & "and create a map file."
   main.info.Caption = main.info.Caption & vbCrLf & "Please press [ RUN ] to start."
   main.Caption = main.Caption & "(" & App.Major & "." & App.Minor & "." & App.Revision & ")"
End Sub

Private Sub Install_button_Click()
   w98 = False
   wNT = False
   main.letter.Visible = False
   main.info = ""
   main.letter = ""
   main.Systemroot = ""
   main.Registeredowner = ""
   main.Productname = ""
   Call Install
End Sub

Public Sub Install()
On Error GoTo Registry_error

Dim errmsg As String
Dim path As String
errmsg = ""
main.info = ""
main.info = main.info & "Reading registry.." & vbCrLf
main.info.Refresh

'determine OS
   If Len(GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "Productname")) > 0 Then
      path = "SOFTWARE\Microsoft\Windows\CurrentVersion"    'w98,w98me,winNT
      w98 = True
   ElseIf Len(GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "Productname")) > 0 Then
      path = "SOFTWARE\Microsoft\Windows NT\CurrentVersion" 'w2k, wXP
      wNT = True
   End If
'read registry OS dependant values
   Productname = GetSettingString(HKEY_LOCAL_MACHINE, path, "Productname")
   Registeredowner = GetSettingString(HKEY_LOCAL_MACHINE, path, "RegisteredOwner")
   Systemroot = GetSettingString(HKEY_LOCAL_MACHINE, path, "SystemRoot")
   If Len(Productname) > 0 Then
      main.Productname = Productname
   Else
      main.Productname = "- error -"
      errmsg = errmsg & "Error reading registry. Productname not found." & vbCrLf
   End If
   If Len(Registeredowner) > 0 Then
      main.Registeredowner = Registeredowner
   Else
      main.Registeredowner = "- error -"
      errmsg = errmsg & "Error reading registry. Registered owner not found." & vbCrLf
   End If
   If Len(Systemroot) > 0 Then
      main.Systemroot = Systemroot
   Else
      main.Systemroot = "- error -"
      errmsg = errmsg & "Error reading registry. Systemroot not found." & vbCrLf
   End If
   
   If Len(errmsg) > 0 Then
      errmsg = errmsg & vbCrLf & "FAILED!" & vbCrLf & vbCrLf
      errmsg = errmsg & "You do not have the rights," & vbCrLf & "or this operating system is not supported." & vbCrLf
      main.info = main.info & errmsg & vbCrLf
      main.info = main.info & "Setup cancelled." & vbCrLf
      main.info.Refresh
      Exit Sub
   Else
'      main.info = main.info & "Writing to registry.." & vbCrLf
'      main.info.Refresh
'      Call SaveSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "name", "MUME Online Map")
'      Call SaveSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "remoteHost", "mume.pvv.org")
'      Call SaveSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "remotePort", "4242")
'      Call SaveSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "localHost", "localhost")
'      Call SaveSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "localPort", "1001")
'      Call SaveSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "version", CStr(App.Major & "." & App.Minor & "." & App.Revision))
'      Call SaveSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "systemRoot", Systemroot)
'      Call SaveSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "productName", Productname)
'      Call SaveSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "setupPath", App.path)
'      Call SaveSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "lookColour", "[32m")
   End If
Registry_error:
   If Err.Number <> 0 Then
      main.info = main.info & vbCrLf & "FAILED!" & vbCrLf & vbCrLf
      main.info = main.info & "Registry error. Setup cancelled. Check current user rights."
      main.info.Refresh
      Exit Sub
   End If

'test FSO installation
Err.Clear
On Error GoTo FSO_error
   main.info = main.info & "Creating FileSystemObject.." & vbCrLf
   main.info.Refresh
   Set fso = New Scripting.FileSystemObject
   Dim systemID As Variant
   systemID = CStr(fso.GetDrive(Mid(Systemroot, 1, 2)).SerialNumber)
FSO_error:
   If Err.Number <> 0 Then
      main.info = main.info & vbCrLf & "FAILED!" & vbCrLf & vbCrLf
      If w98 = True Then
         main.info = main.info & "This program needs MS Windows Scripting. Please visit the MUME Online Map website - download section." & vbCrLf
      ElseIf wNT = True Then
         main.info = main.info & "This program needs MS Windows Scripting. Please visit the MUME Online Map website - download section." & vbCrLf
      Else
         main.info = main.info & "Unsupported operating system. This program needs MS Windows Scripting. Please visit the MUME Online Map website - download section." & vbCrLf
      End If
      main.info.Refresh
      Exit Sub
   End If

Err.Clear
On Error GoTo Encrypt_error
   main.info = main.info & "Generating unique id.." & vbCrLf
   main.info.Refresh
   Set md5 = New MD5DLL.Crypt
   Dim test1 As String
   test1 = md5.Encrypt("kala")
   Dim cast128
   Set cast128 = CreateObject("cast.cipher")
   Dim test2 As Variant
   test2 = cast128.cast128encode("testing", "kollionu")
   
Encrypt_error:
   If Err.Number <> 0 Then
         main.info = main.info & vbCrLf & "FAILED!" & vbCrLf & vbCrLf
         main.info = main.info & "Unregistered system components(CAST128.DLL, MD5DLL.DLL)" & vbCrLf & "Try reinstalling the application." & vbCrLf
         main.info = main.info & "Setup cancelled." & vbCrLf
         main.info.Refresh
      Exit Sub
   End If
   
Err.Clear
'.........................................
On Error GoTo Convert_error
   Dim arrData(0 To 20000) As String
   Dim FileName As String
   FileName = App.path & "\setupmap.txt"
   Dim original As Variant
   Dim encrypted As Variant
   Dim theCount As Integer
   theCount = 0
   Open FileName For Input As #1
   Do While Not EOF(1)
      theCount = theCount + 1
      Line Input #1, encrypted
      arrData(theCount) = encrypted
   Loop
   Close #1
   
   'Systemroot = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "SystemRoot")
   systemID = fso.GetDrive(Mid(Systemroot, 1, 2)).SerialNumber
   Dim key As Variant
   key = getPassword()
   Dim cursor As Integer
   FileName = App.path & "\map.txt"
   Open FileName For Output As #1
   For cursor = 1 To theCount
      encrypted = arrData(cursor)
      original = cast128.cast128decode("780117demsi", encrypted)
      encrypted = cast128.cast128encode(key, original)
      If LenB(Trim(encrypted)) > 0 Then Print #1, encrypted
   Next
   Close #1
   
Convert_error:
   If Err.Number <> 0 Then
         main.info = main.info & vbCrLf & "FAILED!" & vbCrLf & vbCrLf
         main.info = main.info & "Corrupted or missing input data or file(s)." & vbCrLf & "Try reinstalling the application." & vbCrLf
         main.info = main.info & "Setup cancelled." & vbCrLf
         main.info.Refresh
      Exit Sub
   End If
'.........................................
Err.Clear
   main.info = main.info & vbCrLf & "Installation completed." & vbCrLf
   main.info = main.info & vbCrLf & "Please mail the message below."
   main.info = main.info & vbCrLf & "MUME Online Map is ready to run!" & vbCrLf
   main.letter.Visible = True
   main.letter = main.letter & "E-mail: jaanus@2in.ee" & vbCrLf
   main.letter = main.letter & "Subject: MUME_Online_Map_Registration" & vbCrLf
   main.letter = main.letter & vbCrLf
   main.letter = main.letter & "#OS# " & Productname & vbCrLf
   main.letter = main.letter & "#Name# " & Registeredowner & vbCrLf
   main.letter = main.letter & "#Key# " & cast128.cast128encode("780117demsi", key) & vbCrLf
   Install_button.Visible = False
   Exit_button.Visible = True
End Sub

Public Function getPassword()
On Error Resume Next
   Dim sysroot As String
   sysroot = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\LangSoft\MUME Online Map", "SystemRoot")
   Dim sysid As String
   sysid = fso.GetDrive(Mid(Systemroot, 1, 2)).SerialNumber
   getPassword = sysid
End Function

