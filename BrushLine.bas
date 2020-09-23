Attribute VB_Name = "BrushLine"
'Author :Roberto Mior
'     reexre@gmail.com
'
'
'
'
'
'--------------------------------------------------------------------------------

Public Type POINTAPI
    x              As Long
    y              As Long
End Type

Public poi         As POINTAPI


Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function Arc Lib "gdi32" (ByVal hDC As Long, _
                                         ByVal xInizioRettangolo As Long, _
                                         ByVal yInizioRettangolo As Long, _
                                         ByVal xFineRettangolo As Long, _
                                         ByVal yFineRettangolo As Long, _
                                         ByVal xInizioArco As Long, _
                                         ByVal yInizioArco As Long, _
                                         ByVal xFineArco As Long, _
                                         ByVal yFineArco As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long

Public Const STRETCHMODE = vbPaletteModeNone    'You can find other modes in the "PaletteModeConstants" section of your Object Browser



'Declare Function Arc Lib "gdi32.dll" (ByVal HDC As Long, ByVal X1 As Long, _
 ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
 ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

Public PrevColor   As Long
Public PrevWidth   As Long


Public Sub SetBrush(ByVal hDC As Long, ByVal PenWidth As Long, ByVal PenColor As Long)


    DeleteObject (SelectObject(hDC, CreatePen(vbSolid, PenWidth, PenColor)))
    'kOBJ = SelectObject(hDC, CreatePen(vbSolid, PenWidth, PenColor))
    'SetBrush = kOBJ


End Sub



Public Sub FastLine(ByRef hDC As Long, ByRef x1 As Long, ByRef y1 As Long, _
                    ByRef x2 As Long, ByRef y2 As Long, ByRef w As Long, ByRef Color As Long)
Attribute FastLine.VB_Description = "disegna line veloce"

    Dim poi        As POINTAPI

    'SetBrush hdc, W, color
    'If color <> PrevColor Or w <> PrevWidth Then
    DeleteObject (SelectObject(hDC, CreatePen(vbSolid, w, Color)))
    '    PrevColor = color
    '    PrevWidth = w
    'End If

    MoveToEx hDC, x1, y1, poi
    LineTo hDC, x2, y2

End Sub

Sub MyCircle(ByRef hDC As Long, ByRef x As Long, ByRef y As Long, ByRef R As Long, w As Long, Color)
    Dim XpR        As Long

    'If color <> PrevColor Or w <> PrevWidth Then
    DeleteObject (SelectObject(hDC, CreatePen(vbSolid, w, Color)))
    '    PrevColor = color
    '    PrevWidth = w
    'End If

    XpR = x + R

    Arc hDC, x - R, y - R, XpR, y + R, XpR, y, XpR, y

End Sub


Public Sub bLOCK(ByRef hDC As Long, x As Long, y As Long, w As Long, Color As Long)

    DeleteObject (SelectObject(hDC, CreatePen(vbSolid, 1, Color)))

    Rectangle hDC, x, y, x + w, y + w

End Sub
