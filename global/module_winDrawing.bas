Attribute VB_Name = "WinDrawing"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

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

Public Function GenerateBitmapDC(FileName As String) As Long
If DEBUGMODE = False Then On Error GoTo errorhandler
Dim DC As Long
Dim hBitmap As Long
   'Create a Device Context, compatible with the screen
   DC = CreateCompatibleDC(0)
   hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
   'Throw the Bitmap into the Device Context
   SelectObject DC, hBitmap
   'Return the device context
   GenerateBitmapDC = DC
   'Delete the bitmap handle object
   DeleteObject hBitmap
Exit Function
errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "WinDrawing GenerateBitmapDC"
   writeError (errorModule)
End Function

Public Function DeleteGeneratedDC(DC As Long) As Long
   If DC > 0 Then
       DeleteGeneratedDC = DeleteDC(DC)
   Else
       DeleteGeneratedDC = 0
   End If
End Function

Public Function myLine(hWnd, x1, y1, x2, y2)
   MoveToEx hWnd, x1, y1, Point
   LineTo hWnd, x2, y2
End Function

Public Function myText(hWnd, ByVal sText As String, iHCenter, iVCenter, Optional spaces As Boolean, Optional myfontsize As Integer, Optional myfontcolor As Long)
If DEBUGMODE = False Then On Error GoTo errorhandler
Dim MyRect As RECT
Dim sTextWidth As Integer
Dim sTextHeight As Integer
   hWnd.ScaleMode = 3
   
   If myfontsize = 0 Then
      hWnd.Font.Size = 7
   Else
      hWnd.Font.Size = myfontsize
      hWnd.Font.BOLD = True
   End If
   If spaces Then
      sText = " " & Trim$(sText) & " "
   Else
      sText = Trim$(sText)
   End If
   hWnd.ScaleMode = vbPixels
   If myfontcolor <> 0 Then
      hWnd.ForeColor = myfontcolor
   Else
      hWnd.ForeColor = &HFFFFFF
   End If
      
   sTextWidth = hWnd.TextWidth(sText) + Abs((hWnd.Font.Size - 6) * 3)
   sTextHeight = hWnd.TextHeight(sText) + 0
   
   MyRect.Left = iHCenter - (sTextWidth / 2)
   MyRect.Right = iHCenter + (sTextWidth / 2)
   MyRect.Top = iVCenter - (sTextHeight / 2) '+ 2
   MyRect.Bottom = iVCenter + (sTextHeight / 2) ' + 2
   Call drawText(hWnd.hdc, sText, Len(sText), MyRect, DT_TOP)  'Or DT_WORDBREAK
Exit Function

errorhandler:
   errorModule = Err.description & "(" & Err.Number & ") -> " & "winDrawing myText"
   writeError (errorModule)
End Function
