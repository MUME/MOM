Attribute VB_Name = "WinDrawing"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long


Private Declare Function drawText Lib "user32" Alias "DrawTextA" _
  (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, _
  lpRect As RECT, ByVal wFormat As Long) As Long

'DrawText constants
Public Const DT_BOTTOM = &H8
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10

'Pen Constants
Public Const PS_SOLID = 0

'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTSIZE As Long = &H40
'****************************************

Public Function GenerateBitmapDC(FileName As String) As Long
On Error GoTo errorhandler
Dim DC As Long
Dim hBitmap As Long

   'Create a Device Context, compatible with the screen
   DC = CreateCompatibleDC(0)
   hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
   
   'Throw the Bitmap into the Device Context
   SelectObject DC, hBitmap
   
   'Return the device context
   GenerateBitmapDC = DC
   
   'Delte the bitmap handle object
   DeleteObject hBitmap
Exit Function
errorhandler:
   errorData = "WinDrawing GenerateBitmapDC"
   writeError (errorData)

End Function

Public Function DeleteGeneratedDC(DC As Long) As Long
   If DC > 0 Then
       DeleteGeneratedDC = DeleteDC(DC)
   Else
       DeleteGeneratedDC = 0
   End If
End Function

Public Function myLine(ByRef hWnd, ByRef x1, ByRef y1, ByRef x2, ByRef y2)
   MoveToEx hWnd, x1, y1, Point
   LineTo hWnd, x2, y2
End Function

Public Function myText(ByRef hWnd, ByRef sText As String, ByRef iHCenter, ByRef iVCenter)
On Error GoTo errorhandler
Dim MyRect As RECT
Dim sTextWidth As Integer
Dim sTextHeight As Integer
   hWnd.ScaleMode = 3
   sTextWidth = hWnd.TextWidth(sText)
   sTextHeight = hWnd.TextHeight(sText) + 4
   hWnd.Font.Size = 7
   hWnd.ScaleMode = vbPixels
   hWnd.ForeColor = &HFFFFFF
   MyRect.Left = iHCenter - (sTextWidth / 2)
   MyRect.Right = iHCenter + (sTextWidth / 2)
   MyRect.Top = iVCenter - (sTextHeight / 2) + 2
   MyRect.Bottom = iVCenter + (sTextHeight / 2) + 2
   Call drawText(hWnd.hdc, sText, Len(sText), MyRect, DT_TOP)    'Or DT_WORDBREAK

Exit Function
errorhandler:
   errorData = "winDrawing myText"
   writeError (errorData)
End Function
