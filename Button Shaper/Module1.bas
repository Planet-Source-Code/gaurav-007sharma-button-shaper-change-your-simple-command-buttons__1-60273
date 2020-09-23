Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Function CreateREgion(picSkin As PictureBox) As Long
Dim X As Long, Y As Long, StartLineX As Long
Dim fullArea As Long, RegionalLine As Long
Dim TransparentColor As Long
Dim InFirstRegion As Boolean
Dim LineIn As Boolean
Dim hDC As Long
Dim PicWidth As Long
Dim PicHeight As Long
hDC = picSkin.hDC
PicWidth = picSkin.ScaleWidth
PicHeight = picSkin.ScaleHeight
InFirstRegion = True: LineIn = False
X = Y = StartLineX = 0
TransparentColor = GetPixel(hDC, 0, 0)
For Y = 0 To PicHeight - 1
For X = 0 To PicWidth - 1
If GetPixel(hDC, X, Y) = TransparentColor Or X = PicWidth Then
If LineIn Then
LineIn = False
RegionalLine = CreateRectRgn(StartLineX, Y, X, Y + 1)
If InFirstRegion Then
fullArea = RegionalLine
InFirstRegion = False
Else
CombineRgn fullArea, fullArea, RegionalLine, RGN_OR
DeleteObject RegionalLine
End If
End If
Else
If Not LineIn Then
LineIn = True
StartLineX = X
End If
End If
Next
Next
CreateREgion = fullArea
End Function



