Attribute VB_Name = "Module1"
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long ' Written by Yaniv Drukman.
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As Any, ByVal nCount As Long) As Long
Public Const LWA_ALPHA = &H2
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public LevelOfTranslucenceLevel As Integer
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public OPA As Integer
Const ALTERNATE = 1
Const WINDING = 2
Public FLAG  As Boolean
Public Pic1 As String
Public Pic2 As String
Public Function MakeTranslucent(hWnd As Long, ByVal LevelOfTranslucenceLevel As Byte) As Boolean

    SetWindowLong hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    On Error Resume Next
    SetLayeredWindowAttributes hWnd, 0, LevelOfTranslucenceLevel, LWA_ALPHA
End Function


