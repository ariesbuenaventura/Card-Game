Attribute VB_Name = "basCardStuff"
Option Explicit

Public Const HALFTONE As Long = 4

Public Const PS_SOLID As Long = 0

Public Const WM_KEYDOWN As Long = &H100

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As Any, ByVal lpOutput As Any, lpInitData As Any) As Long
Public Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function GetBrushOrgEx Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function PatBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Public Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Declare Function CopyImage Lib "user32.dll" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function CopyRect Lib "user32.dll" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function EqualRect Lib "user32.dll" (lpRect1 As RECT, lpRect2 As RECT) As Long
Public Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function InflateRect Lib "user32.dll" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function IntersectRect Lib "user32.dll" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function OffsetRect Lib "user32.dll" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetRectEmpty Lib "user32.dll" (lpRect As RECT) As Long
Public Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Function GetFocus Lib "user32.dll" () As Long

Public Enum BackgroundConstants
    bgNone
    bgCenter
    bgStretch
    bgTile
End Enum

Public Sub Background(lhDstDC As Long, Image As IPicture, nLeft As Long, nTop As Long, _
                      nWidth As Long, nHeight As Long, Position As BackgroundConstants, _
                      Optional dwRop As RasterOpConstants = vbSrcCopy)
    
    On Error GoTo ErrHandler
    
    If Image.handle = 0 Then Exit Sub
    If (nWidth <= 0) Or (nHeight <= 0) Then Exit Sub
    
    Dim lBmpW   As Long
    Dim lBmpH   As Long
    Dim lhBmp   As Long
    Dim lhDCC   As Long
    Dim lhSrcDC As Long
    Dim lPosX   As Long
    Dim lPosY   As Long
    Dim rcRect  As RECT
    Dim tBMP    As BITMAP
    
    GetObjectAPI Image.handle, Len(tBMP), tBMP
    If Position = bgNone Then
        lBmpW = IIf(tBMP.bmWidth > nWidth, nWidth, tBMP.bmWidth)
        lBmpH = IIf(tBMP.bmHeight > nHeight, nHeight, tBMP.bmHeight)
    ElseIf Position = bgCenter Then
        lBmpW = tBMP.bmWidth
        lBmpH = tBMP.bmHeight
    Else
        lBmpW = nWidth
        lBmpH = nHeight
    End If
    
    lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    lhSrcDC = CreateCompatibleDC(lhDCC)
        
    If (Position = bgNone) Or (Position = bgCenter) Then
        lhBmp = CopyImage(Image.handle, ByVal 0&, tBMP.bmWidth, tBMP.bmHeight, ByVal 0&)
        SelectObject lhSrcDC, lhBmp
        
        If Position = bgNone Then
            lPosX = 0: lPosY = 0
        Else
            lPosX = (nWidth - tBMP.bmWidth) / 2
            lPosY = (nHeight - tBMP.bmHeight) / 2
        End If
    Else
        Dim lhCopyDC  As Long
        Dim lhCopyBmp As Long
        
        lPosX = 0: lPosY = 0
        lhCopyDC = CreateCompatibleDC(lhDCC)
        lhCopyBmp = CopyImage(Image.handle, ByVal 0&, tBMP.bmWidth, tBMP.bmHeight, ByVal 0&)
        lhBmp = CreateCompatibleBitmap(lhDCC, lBmpW, lBmpH)
        SelectObject lhSrcDC, lhBmp
        SelectObject lhCopyDC, lhCopyBmp
            
        If Position = bgStretch Then
            StretchImage lhSrcDC, 0, 0, lBmpW, lBmpH, lhCopyDC, 0, 0, _
                         tBMP.bmWidth, tBMP.bmHeight, vbSrcCopy
        Else
            Dim lhBrush As Long
                
            SetRect rcRect, 0, 0, lBmpW, lBmpH
            lhBrush = CreatePatternBrush(lhCopyBmp)
            FillRect lhSrcDC, rcRect, lhBrush
            DeleteObject lhBrush
        End If
            
        DeleteDC lhCopyDC
        DeleteObject lhCopyBmp
    End If
       
    SetRect rcRect, 0, 0, lBmpW, lBmpH
    OffsetRect rcRect, nLeft + lPosX, nTop + lPosY
    BitBlt lhDstDC, rcRect.Left, rcRect.Top, lBmpW, lBmpH, _
           lhSrcDC, 0, 0, dwRop

    DeleteDC lhDCC
    DeleteDC lhSrcDC
    DeleteObject lhBmp
    Exit Sub
    
ErrHandler:
End Sub

Public Sub DrawFrameRect(lhDstDC As Long, lpRect As RECT, lColor As Long)
    Dim lhBrush As Long
    
    lhBrush = CreateSolidBrush(lColor)
    FrameRect lhDstDC, lpRect, lhBrush
    DeleteObject lhBrush
End Sub

Public Sub DrawLine(lhDstDC As Long, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, _
                    nWidth As Long, nColor As Long)
                    
    Dim lhBrush As Long
    Dim pts     As POINTAPI

    lhBrush = CreatePen(PS_SOLID, nWidth, nColor)
    SelectObject lhDstDC, lhBrush
    MoveToEx lhDstDC, X1, Y1, pts
    LineTo lhDstDC, X2, Y2
    DeleteObject lhBrush
End Sub

Public Sub CombinePicture(lhDstDC As Long, nWidth As Long, nHeight As Long, _
                          Sprite As IPicture, Mask As IPicture, Position As BackgroundConstants)
    
    On Error GoTo ErrHandler
    
    If (Sprite.handle = 0) And (Mask.handle = 0) Then Exit Sub
    
    Background lhDstDC, Mask, 0, 0, nWidth, nHeight, Position, vbSrcCopy
    Background lhDstDC, Sprite, 0, 0, nWidth, nHeight, Position, vbSrcPaint
    Exit Sub
    
ErrHandler:
End Sub

Public Sub RefreshWindow(lhWnd As Long)
    Dim rcRect As RECT
    
    GetClientRect lhWnd, rcRect
    InvalidateRect lhWnd, rcRect, False
End Sub

Public Sub StretchImage(lhDstDC As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, _
                        lhSrcDC As Long, xSrc As Long, ySrc As Long, nSrcWidth As Long, nSrcHeight As Long, _
                        dwRop As RasterOpConstants)
                        
    Dim lMode As Long
    Dim tPT   As POINTAPI
    Dim tMode As POINTAPI
            
    lMode = GetStretchBltMode(lhDstDC)
    SetStretchBltMode lhDstDC, HALFTONE
    GetBrushOrgEx lhDstDC, tMode
    SetBrushOrgEx lhDstDC, 0, 0, tPT
    StretchBlt lhDstDC, x, y, nWidth, nHeight, lhSrcDC, xSrc, ySrc, _
               nSrcWidth, nSrcHeight, dwRop
    SetStretchBltMode lhDstDC, lMode
    SetBrushOrgEx lhDstDC, tPT.x, tPT.y, tMode
End Sub

Public Function Largest(ByVal num1 As Long, ByVal num2 As Long) As Long
    If num1 < num2 Then
        Largest = num2
    Else
        Largest = num1
    End If
End Function

Public Function Smallest(ByVal num1 As Long, ByVal num2 As Long) As Long
    If num1 < num2 Then
        Smallest = num1
    Else
        Smallest = num2
    End If
End Function

