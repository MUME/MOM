VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ==============================================================
' FileName:    cAsync.cls
' Author:      SP McMahon
' Date:        2 February 2000
'
' Very simple demo for mStart.bas.  See comments there
' for more information.  This demo just counts Interval
' 100 second intervals whenever the Start method is
' called.
'
' ==============================================================


Implements Runnable

Private Declare Sub SleepAPI Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)
    
Public Event Complete()
Public Event Cancelled()
Public Event Status(ByVal i As Long, ByRef Cancel As Boolean)

Private m_lInterval As Long
Private m_bRunning As Boolean

Public Property Get Interval() As Long
   Interval = m_lInterval
End Property
Public Property Let Interval(ByVal lInterval As Long)
   m_lInterval = lInterval
End Property

Public Sub Start()
   If Not m_bRunning Then
      m_bRunning = True
      ' Call the mStart module.  This uses a timer to
      ' fire the Runnable_Start() implementation,
      ' which ensures we yield control back to the
      ' caller before the processing starts.  This
      ' ensures that the processing runs asynchronously
      ' to the client.  Easy!!!
      mStart.Start Me
   Else
      ' Just checking....
      Err.Raise 32540, App.EXEName & ".cAsync", "Already running."
   End If
End Sub

Public Sub StartNonAsync()
   ' Just here to demonstrate what happens if
   ' you call a normal AX EXE method
   Runnable_Start
End Sub

Private Sub IRunnable_Start()
   Debug.Print "tere tali"
   End
   ' Do your code here.  It will run async against
   ' the calling app.
End Sub
