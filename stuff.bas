Attribute VB_Name = "BitBltStuff"
' Graphics functions and constants used in the example.
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6 ' for masks
Public Const SRCPAINT = &HEE0086 ' to paint over masks on next blit
