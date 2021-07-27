VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVictoryAni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Default_DistX = 5
Private Const Default_DistY = 5

Private Type DataInfo
    cx   As Integer
    cy   As Integer
    dx   As Integer
    dy   As Integer
    Card As Object
End Type

Dim cx         As Integer
Dim cy         As Integer
Dim dx         As Integer
Dim dy         As Integer
Dim deg        As Single
Dim oCard      As Object
Dim Temp       As Integer
Dim NewX       As Integer
Dim NewY       As Integer
Dim rcRect     As RECT
Dim RightEdge  As Integer
Dim BottomEdge As Integer
Dim IsAlign    As Boolean
Dim bSetCard   As Boolean

Public DistX    As Integer
Public DistY    As Integer
Public Reset    As Boolean
Public CardStat As Integer ' 0-normal, 1-close, 2-open

Dim Trigo As New cTrigonometry
Dim ScriptCalc As New MSScriptControl.ScriptControl

Dim AniInfo() As DataInfo

Public Function Linear(thisForm As Form, _
                       thisPictureBox As PictureBox, _
                       thisCard As Object, _
                       X1 As Integer, Y1 As Integer, _
                       X2 As Integer, Y2 As Integer) As Boolean
    
    Static X     As Single
    Static Y     As Single
    Static px    As Single
    Static py    As Single
    Static dx    As Single
    Static dy    As Single
    Static sdx   As Single
    Static sdy   As Single
    Static dxabs As Single
    Static dyabs As Single
    Static bInit As Boolean
    
    Dim Rate  As Integer
    
    On Error GoTo Errhandler
    
    Rate = 10
    If Not bInit Then
        X = 0: Y = 0
        px = X1: py = Y1
        dx = (X2 - X1) / Rate: dy = (Y2 - Y1) / Rate
        sdx = Sgn(dx): sdy = Sgn(dy)
        dxabs = Abs(dx): dyabs = Abs(dy)
        bInit = True: Linear = False
    Else
        Dim i As Single
        
        ' do not allow the user to close the form
        ' otherwise the program will crash.
        thisForm.Enabled = False
        thisCard.ZOrder 0
        If dxabs >= dyabs Then
            For i = 0 To dxabs
                Y = Y + dyabs * Rate
                If Y >= dxabs * Rate Then
                    Y = Y - dxabs * Rate
                    py = py + sdy * Rate
                End If
                
                thisCard.Move px, py
                thisPictureBox.Refresh
                px = px + sdx * Rate
                DoEvents ' this is the source of the crash
            Next i
            thisCard.Move X2, Y2
            bInit = False: Linear = True
        Else
            For i = 0 To dyabs
                X = X + dxabs * Rate
                If X >= dyabs * Rate Then
                    X = X - dyabs * Rate
                    px = px + sdx * Rate
                End If
                
                thisCard.Move px, py
                thisPictureBox.Refresh
                py = py + sdy * Rate
                DoEvents ' this is the source of the crash
            Next i
            thisCard.Move X2, Y2
            bInit = False: Linear = True
        End If
    End If
    thisForm.Enabled = True
    Exit Function
    
Errhandler:
End Function

Public Sub Bounce(thisForm As Form, _
                  thisPictureBox As PictureBox, _
                  MatchTag As String)
    
    On Error GoTo Errhandler
    
    GetClientRect thisPictureBox.hwnd, rcRect
    RightEdge = rcRect.Right - rcRect.Left
    BottomEdge = rcRect.Bottom - rcRect.Top
    
    If Reset Then
        cx = 0: cy = 0
        dx = DistX: dy = DistY
        Reset = False: IsAlign = False: bSetCard = True
    Else
        If bSetCard Then
            For Each oCard In thisForm.Controls
                If TypeName(oCard) = "Card" Then
                    If oCard.Tag = MatchTag Then
                        oCard.Visible = True
                    End If
                End If
            Next oCard
            bSetCard = False
        End If
        
        For Each oCard In thisForm.Controls
            If TypeName(oCard) = "Card" Then
                If oCard.Tag = MatchTag Then
                    If Abs(dx) <> Abs(DistX) Then
                        dx = IIf(dx < 0, -Abs(DistX), DistX)
                    End If
                    If Abs(dy) <> Abs(DistY) Then
                        dy = IIf(dy < 0, -Abs(DistY), DistY)
                    End If
   
                    Temp = cx + dx
                    If Temp + oCard.Width > RightEdge Then
                        dx = -Abs(dx)
                    ElseIf Temp < 0 Then
                        dx = Abs(dx)
                    End If
        
                    NewX = cx + dx
                    
                    Temp = cy + dy
                    If Temp + oCard.Height > BottomEdge Then
                        dy = -Abs(dy)
                    ElseIf Temp < 0 Then
                        dy = Abs(dy)
                    End If
            
                    NewY = cy + dy
               
                    cx = NewX
                    cy = NewY
                    If Not IsAlign Then
                        Call SetCardStat(oCard)
                        Do While Not Linear(thisForm, thisPictureBox, oCard, _
                                            oCard.Left, oCard.Top, cx, cy)
                        Loop
                    Else
                        oCard.Move cx, cy
                        oCard.ZOrder 0
                        thisPictureBox.Refresh
                    End If
                End If
            End If
        Next oCard
        IsAlign = True
    End If
    Exit Sub
    
Errhandler:
End Sub

Public Sub BounceTrail(thisForm As Form, _
                       thisPictureBox As PictureBox, _
                       thisCard As Object, _
                       ShowTrail As Boolean, _
                       MatchTag As String)
    
    On Error GoTo Errhandler

    GetClientRect thisPictureBox.hwnd, rcRect
    RightEdge = rcRect.Right - rcRect.Left
    BottomEdge = rcRect.Bottom - rcRect.Top

    If Reset Then
        cx = 0: cy = 0
        dx = DistX: dy = DistY
        Reset = False: IsAlign = False: bSetCard = True
    Else
        If bSetCard Then
            For Each oCard In thisForm.Controls
                If TypeName(oCard) = "Card" Then
                    If oCard.Tag = MatchTag Then
                        oCard.Visible = False
                    End If
                End If
            Next oCard
            thisCard.Visible = True
            bSetCard = False
        End If
        
        If Abs(dx) <> Abs(DistX) Then
            dx = IIf(dx < 0, -Abs(DistX), DistX)
        End If
        If Abs(dy) <> Abs(DistY) Then
            dy = IIf(dy < 0, -Abs(DistY), DistY)
        End If
        Temp = cx + dx
        If Temp + thisCard.Width > RightEdge Then
            dx = -Abs(dx)
            Call SetCard(thisCard)
        ElseIf Temp < 0 Then
            dx = Abs(dx)
            Call SetCard(thisCard)
        End If
        
        NewX = cx + dx
            
        Temp = cy + dy
        If Temp + thisCard.Height > BottomEdge Then
            dy = -Abs(dy)
            Call SetCard(thisCard)
        ElseIf Temp < 0 Then
            dy = Abs(dy)
            Call SetCard(thisCard)
        End If
            
        NewY = cy + dy
                    
        cx = NewX
        cy = NewY
                  
        If Not IsAlign Then
            Call SetCardStat(thisCard)
            Do While Not Linear(thisForm, thisPictureBox, thisCard, _
                                thisCard.Left, thisCard.Top, cx, cy)
            Loop
        Else
            If ShowTrail Then
                If thisCard.Visible Then
                    thisCard.Visible = False
                End If
                BitBlt thisPictureBox.hdc, cx, cy, thisCard.Width, thisCard.Height, _
                       thisCard.hdc, 0, 0, vbSrcCopy
                RefreshWindow thisPictureBox.hwnd
            Else
                If Not thisCard.Visible Then
                    thisCard.Visible = True
                End If
                thisCard.Move cx, cy
                thisPictureBox.Refresh
            End If
        End If
        IsAlign = True
    End If
    Exit Sub
    
Errhandler:
End Sub

Public Sub BounceScatter(thisForm As Form, _
                         thisPictureBox As PictureBox, _
                         ShowTrail As Boolean, _
                         MatchTag As String)
                     
    Dim i As Integer
    
    On Error GoTo Errhandler
    
    GetClientRect thisPictureBox.hwnd, rcRect
    RightEdge = rcRect.Right - rcRect.Left
    BottomEdge = rcRect.Bottom - rcRect.Top
    
    If Reset Then
        Dim TempColl As New Collection
        
        For Each oCard In thisForm.Controls
            If TypeName(oCard) = "Card" Then
                If oCard.Tag = MatchTag Then
                    TempColl.Add oCard
                End If
            End If
        Next oCard
        
        ReDim AniInfo(1 To TempColl.Count) As DataInfo
        
        For i = 1 To TempColl.Count
            Call Randomize
            AniInfo(i).cx = CInt(RightEdge * Rnd)
            AniInfo(i).cy = CInt(BottomEdge * Rnd)
            AniInfo(i).dx = CInt(Rnd * 5) + 6
            AniInfo(i).dy = CInt(Rnd * 5) + 6
            Set AniInfo(i).Card = TempColl(i)
            If Not AniInfo(i).Card.Visible Then
                AniInfo(i).Card.Visible = True
            End If
        Next i
        
        Set TempColl = Nothing
        Reset = False: IsAlign = False
    Else
        Dim bVal As Boolean
        
        For i = LBound(AniInfo()) To UBound(AniInfo())
            Temp = AniInfo(i).cx + AniInfo(i).dx
            If Temp + AniInfo(i).Card.Width > RightEdge Then
                bVal = True
                AniInfo(i).dx = -Abs(AniInfo(i).dx)
            ElseIf Temp < 0 Then
                AniInfo(i).dx = Abs(AniInfo(i).dx)
            End If
              
            NewX = AniInfo(i).cx + AniInfo(i).dx
            
            Temp = AniInfo(i).cy + AniInfo(i).dy
            If Temp + AniInfo(i).Card.Height > BottomEdge Then
                AniInfo(i).dy = -Abs(AniInfo(i).dy)
            ElseIf Temp < 0 Then
                AniInfo(i).dy = Abs(AniInfo(i).dy)
            End If
            
            NewY = AniInfo(i).cy + AniInfo(i).dy
                
            AniInfo(i).cx = NewX
            AniInfo(i).cy = NewY
            
            If Not IsAlign Then
                Call SetCardStat(AniInfo(i).Card)
                Do While Not Linear(thisForm, thisPictureBox, AniInfo(i).Card, _
                                    AniInfo(i).Card.Left, AniInfo(i).Card.Top, _
                                    AniInfo(i).cx, AniInfo(i).cy)
                Loop
            Else
                If ShowTrail Then
                    If AniInfo(i).Card.Visible Then
                        AniInfo(i).Card.Visible = False
                    End If
                    BitBlt thisPictureBox.hdc, AniInfo(i).cx, AniInfo(i).cy, _
                                               AniInfo(i).Card.Width, AniInfo(i).Card.Height, _
                                               AniInfo(i).Card.hdc, 0, 0, vbSrcCopy
                    RefreshWindow thisPictureBox.hwnd
                Else
                    If Not AniInfo(i).Card.Visible Then
                        AniInfo(i).Card.Visible = True
                    End If
                    AniInfo(i).Card.Move AniInfo(i).cx, AniInfo(i).cy
                    thisPictureBox.Refresh
                End If
            End If
        Next i
        IsAlign = True
    End If
    
    Exit Sub
    
Errhandler:
End Sub

Public Sub Spin(thisForm As Form, _
                thisPictureBox As PictureBox, _
                MatchTag As String)
    
    Dim RadiusX As Integer
    Dim RadiusY As Integer
    
    On Error GoTo Errhandler
    
    GetClientRect thisPictureBox.hwnd, rcRect
    RightEdge = rcRect.Right - rcRect.Left
    BottomEdge = rcRect.Bottom - rcRect.Top
    
    If Reset Then
        cx = 0: cy = 0
        Reset = False: IsAlign = False: bSetCard = True
    Else
        If bSetCard Then
            For Each oCard In thisForm.Controls
                If TypeName(oCard) = "Card" Then
                    If oCard.Tag = MatchTag Then
                        oCard.Visible = True
                    End If
                End If
            Next oCard
            bSetCard = False
        End If
 
        For Each oCard In thisForm.Controls
            If TypeName(oCard) = "Card" Then
                If oCard.Tag = MatchTag Then
                    RadiusX = (RightEdge - oCard.Width) / 2
                    RadiusY = (BottomEdge - oCard.Height) / 2
                    
                    cx = RadiusX + Cos(Rads(deg)) * RadiusX
                    cy = RadiusY + Sin(Rads(deg)) * RadiusY
                    
                    If Not IsAlign Then
                        Call SetCardStat(oCard)
                        Do While Not Linear(thisForm, thisPictureBox, oCard, _
                                            oCard.Left, oCard.Top, cx, cy)
                        Loop
                    Else
                        oCard.Move cx, cy
                        oCard.ZOrder 0
                        thisPictureBox.Refresh
                    End If
                    
                    deg = deg + DistX
                    If deg > 360 Then deg = 0
                End If
            End If
        Next oCard
        IsAlign = True
    End If
    Exit Sub
    
Errhandler:
End Sub

Public Sub SpinTrail(thisForm As Form, _
                     thisPictureBox As PictureBox, _
                     thisCard As Object, _
                     ShowTrail As Boolean, _
                     MatchTag As String)
                     
    Dim RadiusX As Integer
    Dim RadiusY As Integer
    
    On Error GoTo Errhandler
    
    GetClientRect thisPictureBox.hwnd, rcRect
    RightEdge = rcRect.Right - rcRect.Left
    BottomEdge = rcRect.Bottom - rcRect.Top
    
    If Reset Then
        For Each oCard In thisForm.Controls
            If TypeName(oCard) = "Card" Then
                If oCard.Tag = MatchTag Then
                    oCard.Visible = False
                End If
            End If
        Next oCard
        Reset = False: IsAlign = False: bSetCard = True
    Else
        If bSetCard Then
            For Each oCard In thisForm.Controls
                If TypeName(oCard) = "Card" Then
                    If oCard.Tag = MatchTag Then
                        oCard.Visible = False
                    End If
                End If
            Next oCard
            thisCard.Visible = True
            bSetCard = False
        End If
        
        RadiusX = (RightEdge - thisCard.Width) / 2
        RadiusY = (BottomEdge - thisCard.Height) / 2
                    
        cx = RadiusX + Cos(Rads(deg)) * RadiusX
        cy = RadiusY + Sin(Rads(deg)) * RadiusY
        
        
        If Not IsAlign Then
            Call SetCardStat(thisCard)
            Do While Not Linear(thisForm, thisPictureBox, thisCard, _
                                thisCard.Left, thisCard.Top, cx, cy)
            Loop
        Else
            If ShowTrail Then
                If thisCard.Visible Then
                    thisCard.Visible = False
                End If
                BitBlt thisPictureBox.hdc, cx, cy, thisCard.Width, thisCard.Height, _
                       thisCard.hdc, 0, 0, vbSrcCopy
                RefreshWindow thisPictureBox.hwnd
            Else
                If Not thisCard.Visible Then
                    thisCard.Visible = True
                End If
                thisCard.Move cx, cy
                thisPictureBox.Refresh
            End If
        End If
        
        deg = deg + DistX
        
        If deg > 360 Then
            deg = 0
            
            Call SetCard(thisCard)
        End If
        IsAlign = True
    End If
    Exit Sub

Errhandler:
End Sub

Public Sub Wave(thisForm As Form, _
                thisPictureBox As PictureBox, _
                MatchTag As String, _
                Expression As String, _
                CardWidth As Integer, _
                CardHeight As Integer, _
                Clip As Boolean)
    
    Dim oCard As Object
    Dim Crest As Single

    Static bVal As Boolean
    
    On Error GoTo Errhandler
    
    GetClientRect thisPictureBox.hwnd, rcRect
    RightEdge = rcRect.Right - rcRect.Left
    BottomEdge = rcRect.Bottom - rcRect.Top
    
    If Reset Then
        cx = 0: cy = 0
        dx = (RightEdge - CardWidth) / 2
        dy = (BottomEdge - CardHeight) / 2
        deg = -dx + CardWidth / 2
        bVal = False: IsAlign = False
        Reset = False: bSetCard = True
    Else
        If bSetCard Then
            For Each oCard In thisForm.Controls
                If TypeName(oCard) = "Card" Then
                    If oCard.Tag = MatchTag Then
                        oCard.Visible = True
                    End If
                End If
            Next oCard
            bSetCard = False
        End If
        
        For Each oCard In thisForm.Controls
            If TypeName(oCard) = "Card" Then
                If oCard.Tag = MatchTag Then
                    dx = (RightEdge - oCard.Width) / 2
                    dy = (BottomEdge - oCard.Height) / 2
                    Crest = dy / 10
                        
                    Call ScriptCalc.ExecuteStatement("x=" & Rads(deg))
                    cx = deg: cy = Crest * ScriptCalc.Eval(Expression)
                    cx = dx + cx: cy = dy - cy
                    
                    If Not IsAlign Then
                        Call SetCardStat(oCard)
                        Do While Not Linear(thisForm, thisPictureBox, oCard, _
                                            oCard.Left, oCard.Top, cx, cy)
                        Loop
                    Else
                        If Clip Then
                            If cy + oCard.Height > BottomEdge Then
                                cy = BottomEdge - oCard.Height
                            ElseIf cy < 0 Then
                                cy = 0
                            End If
                        End If
                    
                        oCard.Move cx, cy
                        oCard.ZOrder 0
                        thisPictureBox.Refresh
                    End If
                    If Not bVal Then
                        deg = deg + DistX
                        bVal = IIf(oCard.Left + oCard.Width > RightEdge, True, False)
                    Else
                        deg = deg - DistX
                        bVal = IIf(oCard.Left < 0, False, True)
                    End If
                End If
            End If
        Next oCard
        IsAlign = True
    End If
    Exit Sub
    
Errhandler:
End Sub

Private Sub SetCard(thisCard As Object)
    Call Randomize
    
    thisCard.Update = False
    thisCard.Rank = CInt(Rnd * 12)
    thisCard.Suit = CInt(Rnd * 3)
    thisCard.Update = True
    thisCard.Refresh
End Sub

Private Sub SetCardStat(thisCard As Object)
    Select Case CardStat
    Case Is = 0 ' normal
    Case Is = 1 ' close
        If thisCard.Face = 0 Then
            thisCard.Face = 1
        End If
    Case Is = 2 ' open
        If thisCard.Face = 1 Then
            thisCard.Face = 0
        End If
    End Select
End Sub

Private Sub Class_Initialize()
    Reset = True
    DistX = Default_DistX
    DistY = Default_DistY
    CardStat = 0
    
    ScriptCalc.Language = "VBScript"
    ScriptCalc.Timeout = NoTimeout
    ScriptCalc.AddObject "Trigo", Trigo, True
End Sub

