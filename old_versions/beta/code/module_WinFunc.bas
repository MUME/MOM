Attribute VB_Name = "WinFunc"
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const COLOR_ACTIVECAPTION = 2
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8

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

' get rid of mdi border
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Const MF_BYPOSITION = &H400&
Public Const GWL_STYLE = (-16)
Public Const WS_THICKFRAME = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_SYSMENU = &H80000

' end of get rid of mdi border


Public Declare Function GetWindowRect Lib "user32" _
    (ByVal hWnd As Long, lpRect As RECT) As Long
    
Public Declare Function GetSysColor Lib "user32" _
    (ByVal nIndex As Long) As Long
    
Public Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long

Public Declare Function DrawFocusRect Lib "user32" _
    (ByVal hdc As Long, lpRect As RECT) As Long
    
Public Declare Function ClientToScreen Lib "user32" _
    (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    
Public Declare Function GetDC Lib "user32" _
    (ByVal hWnd As Long) As Long
    
Public Declare Function ReleaseDC Lib "user32" _
    (ByVal hWnd As Long, ByVal hdc As Long) As Long

' end of window movement

Public Sub MakeNormal(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
