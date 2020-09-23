Attribute VB_Name = "Module1"
Public Snow(1000) As Projectile

Type Projectile
    Used As Boolean
    Type As Integer
    XSpeed As Double
    YSpeed As Double
    XPos As Double
    YPos As Double
End Type

Public Const SRCERASE As Long = &H440328
Public Const SRCINVERT As Long = &H660046
Public Const SRCPAINT  As Long = &HEE0086
Public Const SRCAND As Long = &H8800C6 'For BitBlitting (Merge two images)
Public Const SRCCOPY  As Long = &HCC0020 'For BitBlitting (Copy image ofer top of other)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
   ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
