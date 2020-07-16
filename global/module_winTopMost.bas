Attribute VB_Name = "WinTopMost"
Option Explicit
'set window foregroundwindow
Private Const GW_HWNDNEXT = 2
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'------------------------
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'end of get rid of mdi border
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
'win on top constants
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const COLOR_ACTIVECAPTION = 2
Const SM_CXDLGFRAME = 7
Const SM_CYDLGFRAME = 8
' window movement without titlebar
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Public Type POINTAPI
   x As Long
   y As Long
End Type
Public Point As POINTAPI
' get rid of mdi border
Public Const MF_BYPOSITION = &H400&
Public Const GWL_STYLE = (-16)
Public Const WS_THICKFRAME = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_SYSMENU = &H80000

' end of window movement
Public Sub MakeNormal(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Function FindWindowPartial(ByVal Title As String, ByVal Class As String) As Long
    Dim hWndThis As Long
    hWndThis = FindWindow(vbNullString, vbNullString)
    While hWndThis
      Dim sTitle As String, sClass As String
      sTitle = Space$(255)
      sTitle = LCase(Left$(sTitle, GetWindowText(hWndThis, sTitle, Len(sTitle))))
      sClass = Space$(255)
      sClass = Left$(sClass, GetClassName(hWndThis, sClass, Len(sClass)))
      If sTitle Like Title Then
          FindWindowPartial = hWndThis
          Exit Function
      End If
      hWndThis = GetWindow(hWndThis, GW_HWNDNEXT)
    Wend
End Function

