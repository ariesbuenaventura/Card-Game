VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAddControlEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ObjKey    As New Collection
Dim ColorBL1  As New Collection
Dim ColorBL2  As New Collection
Dim ColorCir  As New Collection
Dim LastHwnd  As Long

Private Sub ClearEffect()
    If LastHwnd <> 0 Then
        Dim rcRect As RECT
    
        ' refresh window
        GetClientRect LastHwnd, rcRect
        InvalidateRect LastHwnd, rcRect, False
    End If
End Sub

Private Sub DrawEffect(lhWnd As Long, cb1 As Long, cb2 As Long, cc As Long)
    Dim lhBrush As Long
    Dim lhDC    As Long
    Dim rcRect  As RECT
    
    GetClientRect lhWnd, rcRect
    ' get the object hDC(handle device object)
    lhDC = GetDC(lhWnd)
    
    ' once we have the hDC we can draw anything or
    ' add a pic in the control
    DrawFrameRect lhDC, rcRect, cb1
    InflateRect rcRect, -1, -1
    DrawFrameRect lhDC, rcRect, cb2
    
    If cc <> -1 Then
        lhBrush = CreateSolidBrush(cc)
        SelectObject lhDC, lhBrush
        Ellipse lhDC, 2, 2, 9, 9
        DeleteObject lhBrush
    End If
    
    ReleaseDC lhWnd, lhDC
End Sub

Public Sub RegisterControl(thisButton As Object, ColorBorder1 As Long, ColorBorder2 As Long, _
                           ColorCircle As Long)
                          
    ObjKey.Add thisButton.hwnd
    ColorBL1.Add ColorBorder1
    ColorBL2.Add ColorBorder2
    ColorCir.Add ColorCircle
End Sub

Public Sub StartEffect()
    If ObjKey.Count = 0 Then Exit Sub
        
    Dim lhWnd    As Long
    Dim MousePT  As POINTAPI
    Dim step     As Long
    Dim IsEffect As Boolean
    
    Static bVal    As Boolean
    Static OldHwnd As Long
    
    GetCursorPos MousePT
    lhWnd = WindowFromPoint(MousePT.X, MousePT.Y)
    
    If OldHwnd = lhWnd Then Exit Sub
    
    For step = 1 To ObjKey.Count
        If ObjKey(step) = lhWnd Then
            IsEffect = True
            DrawEffect lhWnd, ColorBL1(step), ColorBL2(step), _
                       ColorCir(step)
        End If
    Next step
    
    Call ClearEffect
    
    If IsEffect Then
        bVal = True: LastHwnd = lhWnd
    Else
        If bVal Then
            bVal = False: LastHwnd = 0
        End If
    End If
    
    OldHwnd = lhWnd
End Sub

Private Sub DrawFrameRect(lhDstDC As Long, lpRect As RECT, lColor As Long)
    Dim lhBrush As Long
    
    lhBrush = CreateSolidBrush(lColor)
    FrameRect lhDstDC, lpRect, lhBrush
    DeleteObject lhBrush
End Sub
