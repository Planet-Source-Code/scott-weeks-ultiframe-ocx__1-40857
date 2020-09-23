Attribute VB_Name = "modEdges"
Option Explicit

'*********************************************'
'*       CreatePen fnPenStyle Constants      *'
'*********************************************'
Private Const PS_SOLID As Long = 0       ' The pen is solid.
Private Const PS_DASH As Long = 1        ' The pen is dashed. This style is valid only when the pen width is one or less in device units.
Private Const PS_DOT As Long = 2         ' The pen is dotted. This style is valid only when the pen width is one or less in device units.
Private Const PS_DASHDOT As Long = 3     ' The pen has alternating dashes and dots. This style is valid only when the pen width is one or less in device units.
Private Const PS_DASHDOTDOT As Long = 4  ' The pen has alternating dashes and double dots. This style is valid only when the pen width is one or less in device units.
Private Const PS_NULL As Long = 5        ' The pen is invisible.
Private Const PS_INSIDEFRAME As Long = 6 ' The pen is solid. When this pen is used in any GDI drawing function that takes a bounding rectangle, the dimensions of the figure are shrunk so that it fits entirely in the bounding rectangle, taking into account the width of the pen. This applies only to geometric pens.

'***********************************************'
'*              Win32 API Declares             *'
'***********************************************'
Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Sub Draw3DEdge(ByVal hDestDC As Long, ByRef Target As Control, _
                      ByRef rcFrame As RECT, ByRef rcCapBox As RECT, _
                      crColor1 As Long, crColor2 As Long, lRadius As Long)
                     
Dim lPen As Long
Dim lPenOld As Long
                     
ExcludeClipRect Target.hDC, rcCapBox.Left, rcCapBox.Top, rcCapBox.Right, rcCapBox.Bottom

lPen = CreatePen(PS_SOLID, 1, TranslateColor(crColor1))
lPenOld = SelectObject(Target.hDC, lPen)

RoundRect Target.hDC, rcFrame.Left + 1, rcFrame.Top + 1, rcFrame.Right, rcFrame.Bottom, lRadius, lRadius

SelectObject Target.hDC, lPenOld
DeleteObject lPen

lPen = CreatePen(PS_SOLID, 1, TranslateColor(crColor2))
lPenOld = SelectObject(Target.hDC, lPen)

RoundRect Target.hDC, rcFrame.Left, rcFrame.Top, rcFrame.Right - 1, rcFrame.Bottom - 1, lRadius, lRadius

SelectObject Target.hDC, lPenOld
DeleteObject lPen

BitBlt hDestDC, 0, 0, Target.ScaleWidth, Target.ScaleHeight, Target.hDC, 0, 0, vbSrcCopy

End Sub

Public Sub DrawFlatEdge(ByVal hDestDC As Long, ByRef Target As Control, _
                        ByRef rcFrame As RECT, ByRef rcCapBox As RECT, _
                        crColor As Long, lStyle As Long, lRadius As Long)

Dim lPen As Long
Dim lPenOld As Long
                     
lPen = CreatePen(lStyle, 1, TranslateColor(crColor))
lPenOld = SelectObject(Target.hDC, lPen)

RoundRect Target.hDC, rcFrame.Left, rcFrame.Top, rcFrame.Right, rcFrame.Bottom, lRadius, lRadius

SelectObject Target.hDC, lPenOld
DeleteObject lPen

BitBlt hDestDC, 0, 0, Target.ScaleWidth, Target.ScaleHeight, Target.hDC, 0, 0, vbSrcCopy

End Sub


