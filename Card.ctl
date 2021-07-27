VERSION 5.00
Begin VB.UserControl Card 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   825
   ScaleHeight     =   1080
   ScaleWidth      =   825
   ToolboxBitmap   =   "Card.ctx":0000
   Begin VB.PictureBox picHolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const Cards_Per_Suit = 13

Private Const OffsetRankResId = 101
Private Const OffsetDeckResId = 201
Private Const OffsetMaskResId = 301

Public Enum EffectConstants
    crd_Effect_None
    crd_Effect_Elevator
    crd_Effect_Flip
    crd_Effect_FlyIn
    crd_Effect_FlyOut
    crd_Effect_Gate
    crd_Effect_Split
    crd_Effect_Stretch
    crd_Effect_ThreeD
    crd_Effect_Wipe
    crd_Effect_Zoom
End Enum

Public Enum ElevatorConstants
    crd_Elevator_Down
    crd_Elevator_Left
    crd_Elevator_Right
    crd_Elevator_Up
End Enum

Public Enum FlipConstants
    crd_Flip_Bottom
    crd_Flip_Horizontal
    crd_Flip_Left
    crd_Flip_Right
    crd_Flip_Top
    crd_Flip_Vertical
End Enum

Public Enum FlyInConstants
    crd_FlyIn_Bottom
    crd_FlyIn_Bottom_Left
    crd_FlyIn_Bottom_Right
    crd_FlyIn_Left
    crd_FlyIn_Right
    crd_FlyIn_Top
    crd_FlyIn_Top_Left
    crd_FlyIn_Top_Right
End Enum

Public Enum FlyOutConstants
    crd_FlyOut_Bottom
    crd_FlyOut_Bottom_Left
    crd_FlyOut_Bottom_Right
    crd_FlyOut_Left
    crd_FlyOut_Right
    crd_FlyOut_Top
    crd_FlyOut_Top_Left
    crd_FlyOut_Top_Right
End Enum

Public Enum GateConstants
    crd_Gate_Horizontal_In
    crd_Gate_Horizontal_Out
    crd_Gate_Vertical_In
    crd_Gate_Vertical_Out
End Enum

Public Enum SplitConstants
    crd_Split_Horizontal_In
    crd_Split_Horizontal_Out
    crd_Split_Vertical_In
    crd_Split_Vertical_Out
End Enum

Public Enum StretchConstants
    crd_Stretch_Across
    crd_Stretch_From_Bottom
    crd_Stretch_From_Left
    crd_Stretch_From_Right
    crd_Stretch_From_Top
End Enum

Public Enum ThreeDConstants
    crd_ThreeD_From_Bottom
    crd_ThreeD_From_Left
    crd_ThreeD_From_Right
    crd_ThreeD_From_Top
End Enum

Public Enum WipeConstants
    crd_Wipe_Down
    crd_Wipe_Left
    crd_Wipe_Right
    crd_Wipe_Up
End Enum
    
Public Enum ZoomConstants
    crd_Zoom_In
    crd_Zoom_Out
End Enum

Public Enum DeckBackgroundConstants
    crd_BG_Stretch
    crd_BG_Tile
End Enum

Public Enum DeckConstants
    crd_Deck_1
    crd_Deck_2
    crd_Deck_Customize
End Enum

Public Enum DeckMaskStyleConstants
    crd_None
    crd_Angel
    crd_Bear
    crd_Bird
    crd_Dog_Paint
    crd_Dog_Butterfly
    crd_Leaf
    crd_Swan
    crd_Text
    crd_Customize
End Enum

Public Enum FaceConstants
    crd_Up
    crd_Down
End Enum

Public Enum RankConstants
    crd_Ace
    crd_Two
    crd_Three
    crd_Four
    crd_Five
    crd_Six
    crd_Seven
    crd_Eight
    crd_Nine
    crd_Ten
    crd_Jack
    crd_Queen
    crd_King
End Enum

Public Enum SuitConstants
    crd_Clubs
    crd_Spades
    crd_Hearts
    crd_Diamond
End Enum

Private Type CardProperties
    cArrowKeyFocus   As Boolean
    cAutoFlipCard    As Boolean
    cBorderLine      As Boolean
    sText         As String
    cData            As String
    cDeck            As DeckConstants
    cDeckBackground  As DeckBackgroundConstants
    cDeckMaskPicture As New StdPicture
    cDeckPicture     As New StdPicture
    cDeckMaskStyle   As DeckMaskStyleConstants
    cEffect          As EffectConstants
    cElevator        As ElevatorConstants
    cFace            As FaceConstants
    cFlip            As FlipConstants
    cFlyIn           As FlyInConstants
    cFlyOut          As FlyOutConstants
    cFramePerMoveX   As Integer
    cFramePerMoveY   As Integer
    cGate            As GateConstants
    cHotTracking     As Boolean
    cRank            As RankConstants
    cShowFocusRect   As Boolean
    cSelected        As Boolean
    cSpeed           As Integer
    cSplit           As SplitConstants
    cStopAni         As Boolean
    cStretch         As StretchConstants
    cSuit            As SuitConstants
    cThreeD          As ThreeDConstants
    cUpdate          As Boolean
    cValue           As Integer
    cWipe            As WipeConstants
    cZoom            As ZoomConstants
End Type

Private Type CardSlideVarHolder
    bPlayMode   As Boolean
    lAniEffect  As Long
    lBmpW       As Long
    lBmpH       As Long
    lhSrcDC     As Long
    lhSvBgDC    As Long
    lWinOffX    As Long
    lWinOffY    As Long
    lWinSzX     As Long
    lWinSzY     As Long
    rcDummy     As RECT
    rcRect      As RECT
    rcTemp      As RECT
End Type

Dim hCursor    As Long
Dim hOldCursor As Long
Dim MyProp     As CardProperties
Dim CSH        As CardSlideVarHolder

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Event AniPosition(ALeft As Single, ATop As Single, ARight As Single, ABottom As Single)

Public Property Get ArrowKeyFocus() As Boolean
    ArrowKeyFocus = MyProp.cArrowKeyFocus
End Property

Public Property Let ArrowKeyFocus(bArrowKeyFocus As Boolean)
    If bArrowKeyFocus <> MyProp.cArrowKeyFocus Then
        MyProp.cArrowKeyFocus = bArrowKeyFocus
        PropertyChanged "ArrowKeyFocus"
    End If
End Property

Public Property Get AutoFlipCard() As Boolean
    AutoFlipCard = MyProp.cAutoFlipCard
End Property

Public Property Let AutoFlipCard(bAutoFlipCard As Boolean)
    If bAutoFlipCard <> MyProp.cAutoFlipCard Then
        MyProp.cAutoFlipCard = bAutoFlipCard
        PropertyChanged "AutoFlipCard"
    End If
End Property

Public Property Get BorderLine() As Boolean
    BorderLine = MyProp.cBorderLine
End Property

Public Property Let BorderLine(bBorderLine As Boolean)
    If bBorderLine <> MyProp.cBorderLine Then
        MyProp.cBorderLine = bBorderLine
        PropertyChanged "BorderLine"
        Call RedrawCard
    End If
End Property

Public Property Get Text() As String
    Text = MyProp.sText
End Property

Public Property Let Text(sText As String)
Attribute Text.VB_UserMemId = -518
Attribute Text.VB_MemberFlags = "200"
    MyProp.sText = sText
    PropertyChanged "Text"
    Call RedrawCard
End Property

Public Property Get Data() As String
    Data = MyProp.cData
End Property

Public Property Let Data(sData As String)
    MyProp.cData = sData
    PropertyChanged "Data"
End Property

Public Property Get Deck() As DeckConstants
    Deck = MyProp.cDeck
End Property

Public Property Let Deck(DeckOp As DeckConstants)
    If DeckOp <> MyProp.cDeck Then
        MyProp.cDeck = DeckOp
        PropertyChanged "Deck"
        If MyProp.cFace = crd_Down Then Call RedrawCard
    End If
End Property

Public Property Get DeckBackground() As DeckBackgroundConstants
    DeckBackground = MyProp.cDeckBackground
End Property

Public Property Let DeckBackground(DeckBackgroundOp As DeckBackgroundConstants)
    If DeckBackgroundOp <> MyProp.cDeckBackground Then
        MyProp.cDeckBackground = DeckBackgroundOp
        PropertyChanged "DeckBackground"
        Call RedrawCard
    End If
End Property

Public Property Get DeckMaskPicture() As StdPicture
    Set DeckMaskPicture = MyProp.cDeckMaskPicture
End Property

Public Property Set DeckMaskPicture(ByVal New_DeckMaskPicture As StdPicture)
    Set MyProp.cDeckMaskPicture = New_DeckMaskPicture
    PropertyChanged "DeckMaskPicture"
    Call RedrawCard
End Property

Public Property Get DeckPicture() As StdPicture
    Set DeckPicture = MyProp.cDeckPicture
End Property

Public Property Set DeckPicture(ByVal New_DeckPicture As StdPicture)
    Set MyProp.cDeckPicture = New_DeckPicture
    PropertyChanged "DeckPicture"
    If (MyProp.cDeck = crd_Deck_Customize) And (MyProp.cFace = crd_Down) Then
        Call RedrawCard
    End If
End Property

Public Property Get DeckMaskStyle() As DeckMaskStyleConstants
    DeckMaskStyle = MyProp.cDeckMaskStyle
End Property

Public Property Let DeckMaskStyle(DeckMaskStyleOp As DeckMaskStyleConstants)
    If DeckMaskStyleOp <> MyProp.cDeckMaskStyle Then
        MyProp.cDeckMaskStyle = DeckMaskStyleOp
        PropertyChanged "DeckMaskStyle"
        If MyProp.cFace = crd_Down Then RedrawCard
    End If
End Property

Public Property Get Effect() As EffectConstants
    Effect = MyProp.cEffect
End Property

Public Property Let Effect(EffectOp As EffectConstants)
    If EffectOp <> MyProp.cEffect Then
        MyProp.cEffect = EffectOp
        PropertyChanged "Effect"
    End If
End Property

Public Property Get Elevator() As ElevatorConstants
    Elevator = MyProp.cElevator
End Property

Public Property Let Elevator(ElevatorOp As ElevatorConstants)
    If ElevatorOp <> MyProp.cElevator Then
        MyProp.cElevator = ElevatorOp
        PropertyChanged "Elevator"
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)
    UserControl.Enabled() = bEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get Face() As FaceConstants
    Face = MyProp.cFace
End Property

Public Property Let Face(FaceOp As FaceConstants)
    If FaceOp <> MyProp.cFace Then
        MyProp.cFace = FaceOp
        PropertyChanged "Face"
        Call RedrawCard
    End If
End Property

Public Property Get Flip() As FlipConstants
    Flip = MyProp.cFlip
End Property

Public Property Let Flip(FlipOp As FlipConstants)
    If FlipOp <> MyProp.cFlip Then
        MyProp.cFlip = FlipOp
        PropertyChanged "Flip"
    End If
End Property

Public Property Get FlyIn() As FlyInConstants
    FlyIn = MyProp.cFlyIn
End Property

Public Property Let FlyIn(FlyInOp As FlyInConstants)
    If FlyInOp <> MyProp.cFlyIn Then
        MyProp.cFlyIn = FlyInOp
        PropertyChanged "FlyIn"
    End If
End Property

Public Property Get FlyOut() As FlyOutConstants
    FlyOut = MyProp.cFlyOut
End Property

Public Property Let FlyOut(FlyOutOp As FlyOutConstants)
    If FlyOutOp <> MyProp.cFlyOut Then
        MyProp.cFlyOut = FlyOutOp
        PropertyChanged "FlyOut"
    End If
End Property

Public Property Get Font() As Font
    Set Font = picHolder.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set picHolder.Font = New_Font
    PropertyChanged "Font"
    If MyProp.cFace = crd_Down Then Call RedrawCard
End Property

Public Property Get FontBold() As Boolean
    FontBold = picHolder.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    If New_FontBold <> picHolder.FontBold Then
        picHolder.FontBold() = New_FontBold
        PropertyChanged "FontBold"
        If MyProp.cFace = crd_Down Then Call RedrawCard
    End If
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = picHolder.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    If New_FontItalic <> picHolder.FontItalic Then
        picHolder.FontItalic() = New_FontItalic
        PropertyChanged "FontItalic"
        If MyProp.cFace = crd_Down Then Call RedrawCard
    End If
End Property

Public Property Get FontName() As String
    FontName = picHolder.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    picHolder.FontName() = New_FontName
    PropertyChanged "FontName"
    If MyProp.cFace = crd_Down Then Call RedrawCard
End Property

Public Property Get FontSize() As Single
    FontSize = picHolder.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    picHolder.FontSize() = New_FontSize
    PropertyChanged "FontSize"
    If MyProp.cFace = crd_Down Then Call RedrawCard
End Property

Public Property Get FontStrikethru() As Boolean
    FontStrikethru = picHolder.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    If New_FontStrikethru <> picHolder.FontStrikethru Then
        picHolder.FontStrikethru() = New_FontStrikethru
        PropertyChanged "FontStrikethru"
        If MyProp.cFace = crd_Down Then Call RedrawCard
    End If
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = picHolder.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    If New_FontUnderline <> picHolder.FontUnderline Then
        picHolder.FontUnderline() = New_FontUnderline
        PropertyChanged "FontUnderline"
        If MyProp.cFace = crd_Down Then Call RedrawCard
    End If
End Property

Public Property Get FontTransparent() As Boolean
    FontTransparent = picHolder.FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    If New_FontTransparent <> picHolder.FontTransparent Then
        picHolder.FontTransparent() = New_FontTransparent
        PropertyChanged "FontTransparent"
        If MyProp.cFace = crd_Down Then Call RedrawCard
    End If
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = picHolder.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    If New_ForeColor <> picHolder.ForeColor Then
        picHolder.ForeColor() = New_ForeColor
        PropertyChanged "ForeColor"
        If MyProp.cFace = crd_Down Then Call RedrawCard
    End If
End Property

Public Property Let FramePerMoveX(nStep As Integer)
    If nStep <> MyProp.cFramePerMoveX Then
        MyProp.cFramePerMoveX = nStep
        PropertyChanged "FramePerMoveX"
    End If
End Property

Public Property Get FramePerMoveX() As Integer
    FramePerMoveX = MyProp.cFramePerMoveX
End Property

Public Property Get FramePerMoveY() As Integer
    FramePerMoveY = MyProp.cFramePerMoveY
End Property

Public Property Let FramePerMoveY(nStep As Integer)
    If nStep <> MyProp.cFramePerMoveY Then
        MyProp.cFramePerMoveY = nStep
        PropertyChanged "FramePerMoveY"
    End If
End Property

Public Property Get Gate() As GateConstants
    Gate = MyProp.cGate
End Property

Public Property Let Gate(GateOp As GateConstants)
    If GateOp <> MyProp.cGate Then
        MyProp.cGate = GateOp
        PropertyChanged "Gate"
    End If
End Property

Public Property Get HotTracking() As Boolean
    HotTracking = MyProp.cHotTracking
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Let HotTracking(bHotTracking As Boolean)
    If bHotTracking <> MyProp.cHotTracking Then
        MyProp.cHotTracking = bHotTracking
        PropertyChanged "HotTracking"
    End If
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Picture() As StdPicture
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Public Property Get Rank() As RankConstants
    Rank = MyProp.cRank
End Property

Public Property Let Rank(RankOp As RankConstants)
    If RankOp <> MyProp.cRank Then
        MyProp.cRank = RankOp
        PropertyChanged "Rank"
        If Face = crd_Up Then
            Call RedrawCard
        End If
        MyProp.cValue = MyProp.cRank + 1
    End If
End Property

Public Property Get Selected() As Boolean
    Selected = MyProp.cSelected
End Property

Public Property Let Selected(bSelected As Boolean)
    If bSelected <> MyProp.cSelected Then
        MyProp.cSelected = bSelected
        PropertyChanged "Selected"
        Call RedrawCard
    End If
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = MyProp.cShowFocusRect
End Property

Public Property Let ShowFocusRect(bShowFocusRect As Boolean)
    If bShowFocusRect <> MyProp.cShowFocusRect Then
        MyProp.cShowFocusRect = bShowFocusRect
        PropertyChanged "ShowFocusRect"
    End If
End Property

Public Property Get Speed() As Integer
    Speed = MyProp.cSpeed
End Property

Public Property Let Speed(nVal As Integer)
    If nVal <> MyProp.cSpeed Then
        MyProp.cSpeed = nVal
        PropertyChanged "Speed"
    End If
End Property

Public Property Get Split() As SplitConstants
    Split = MyProp.cSplit
End Property

Public Property Let Split(SplitOp As SplitConstants)
    If SplitOp <> MyProp.cSplit Then
        MyProp.cSplit = SplitOp
        PropertyChanged "Split"
    End If
End Property

Public Property Get StopAni() As Boolean
    StopAni = MyProp.cStopAni
End Property

Public Property Let StopAni(bStopAni As Boolean)
    If bStopAni <> MyProp.cStopAni Then
        MyProp.cStopAni = bStopAni
        PropertyChanged "StopAni"
    End If
End Property

Public Property Get Suit() As SuitConstants
    Suit = MyProp.cSuit
End Property

Public Property Let Suit(SuitOp As SuitConstants)
    If SuitOp <> MyProp.cSuit Then
        MyProp.cSuit = SuitOp
        PropertyChanged "Suit"
        If MyProp.cFace = crd_Up Then
            Call RedrawCard
        End If
    End If
End Property

Public Property Get ThreeD() As ThreeDConstants
    ThreeD = MyProp.cThreeD
End Property

Public Property Let ThreeD(ThreeDOp As ThreeDConstants)
    If ThreeDOp <> MyProp.cThreeD Then
        MyProp.cThreeD = ThreeDOp
        PropertyChanged "ThreeD"
    End If
End Property

Public Property Get Update() As Boolean
    Update = MyProp.cUpdate
End Property

Public Property Let Update(bUpdate As Boolean)
    If bUpdate <> MyProp.cUpdate Then
        MyProp.cUpdate = bUpdate
        PropertyChanged "Update"
    End If
End Property

Public Property Get Value() As Integer
    Value = MyProp.cValue
End Property

Public Property Let Value(nVal As Integer)
    If nVal <> MyProp.cValue Then
        MyProp.cValue = nVal
        PropertyChanged "Value"
    End If
End Property

Public Property Get Stretch() As StretchConstants
    Stretch = MyProp.cStretch
End Property

Public Property Let Stretch(StretchOp As StretchConstants)
    If StretchOp <> MyProp.cStretch Then
        MyProp.cStretch = StretchOp
        PropertyChanged "Stretch"
    End If
End Property

Public Property Get Wipe() As WipeConstants
    Wipe = MyProp.cWipe
End Property

Public Property Let Wipe(WipeOp As WipeConstants)
    If WipeOp <> MyProp.cWipe Then
        MyProp.cWipe = WipeOp
        PropertyChanged "Wipe"
    End If
End Property

Public Property Get Zoom() As ZoomConstants
    Zoom = MyProp.cZoom
End Property

Public Property Let Zoom(ZoomOp As ZoomConstants)
    If ZoomOp <> MyProp.cZoom Then
        MyProp.cZoom = ZoomOp
        PropertyChanged "Zoom"
    End If
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub
 
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
    If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeySpace) Then
        Call UserControl_Click
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Public Sub OLEDrag()
    UserControl.OLEDrag
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Public Property Get OLEDropMode() As Integer
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub
 
Private Sub UserControl_GotFocus()
    If Not MyProp.cShowFocusRect Then Exit Sub
    
    If GetProp(UserControl.hWnd, "Track") Then
        ' do nothing
    Else
        Call DrawFocusRect
    End If
    
    SetProp UserControl.hWnd, "Focus", True
End Sub

Private Sub UserControl_InitProperties()
    MyProp.cArrowKeyFocus = True
    MyProp.cAutoFlipCard = True
    MyProp.cBorderLine = True
    MyProp.cDeck = crd_Deck_1
    MyProp.cDeckBackground = crd_BG_Stretch
    MyProp.cDeckMaskStyle = crd_None
    MyProp.cEffect = crd_Effect_None
    MyProp.cFace = crd_Up
    MyProp.cFramePerMoveX = 1
    MyProp.cFramePerMoveY = 1
    MyProp.cHotTracking = False
    MyProp.cRank = crd_Ace
    MyProp.cSelected = False
    MyProp.cShowFocusRect = False
    MyProp.cSuit = crd_Clubs
    MyProp.cSpeed = 1
    MyProp.cStopAni = True
    MyProp.cUpdate = True
    MyProp.cValue = 1
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If MyProp.cArrowKeyFocus Then
        Dim lhWndParen  As Long
    
        lhWndParen = GetParent(UserControl.hWnd)
    
        Select Case KeyCode
        Case Is = vbKeyRight
            KeyCode = 0
            PostMessage lhWndParen, WM_KEYDOWN, ByVal &H27, ByVal &H4D0001
        Case Is = vbKeyDown
            KeyCode = 0
            PostMessage lhWndParen, WM_KEYDOWN, ByVal &H28, ByVal &H500001
        Case Is = vbKeyLeft
            KeyCode = 0
            PostMessage lhWndParen, WM_KEYDOWN, ByVal &H25, ByVal &H4B0001
        Case Is = vbKeyUp
            KeyCode = 0
            PostMessage lhWndParen, WM_KEYDOWN, ByVal &H26, ByVal &H480001
        End Select
    End If
    
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    If Not MyProp.cShowFocusRect Then Exit Sub
    
    If GetProp(UserControl.hWnd, "Track") Then
        ' do nothing
    Else
        UserControl.Cls
    End If
    
    SetProp UserControl.hWnd, "Focus", False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MyProp.cHotTracking Then
        Call SetHotTracking
    End If

    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MyProp.cArrowKeyFocus = PropBag.ReadProperty("ArrowKeyFocus", True)
    MyProp.cAutoFlipCard = PropBag.ReadProperty("AutoFlipCard", True)
    MyProp.cBorderLine = PropBag.ReadProperty("BorderLine", True)
    MyProp.sText = PropBag.ReadProperty("Text", "")
    MyProp.cData = PropBag.ReadProperty("Data", "")
    MyProp.cDeck = PropBag.ReadProperty("Deck", crd_Deck_1)
    MyProp.cDeckBackground = PropBag.ReadProperty("DeckBackground", crd_BG_Stretch)
    Set DeckMaskPicture = PropBag.ReadProperty("DeckMaskPicture", Nothing)
    Set DeckPicture = PropBag.ReadProperty("DeckPicture", Nothing)
    MyProp.cDeckMaskStyle = PropBag.ReadProperty("DeckMaskStyle", crd_None)
    MyProp.cEffect = PropBag.ReadProperty("Effect", crd_Effect_None)
    MyProp.cElevator = PropBag.ReadProperty("Elevator", crd_Elevator_Left)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    MyProp.cFace = PropBag.ReadProperty("Face", crd_Up)
    MyProp.cFlip = PropBag.ReadProperty("Flip", crd_Flip_Horizontal)
    MyProp.cFlyIn = PropBag.ReadProperty("FlyIn", crd_FlyIn_Left)
    MyProp.cFlyOut = PropBag.ReadProperty("FlyOut", crd_FlyOut_Left)
    MyProp.cFramePerMoveX = PropBag.ReadProperty("FramePerMoveX", 1)
    MyProp.cFramePerMoveY = PropBag.ReadProperty("FramePerMoveY", 1)
    MyProp.cGate = PropBag.ReadProperty("Gate", crd_Gate_Horizontal_In)
    MyProp.cHotTracking = PropBag.ReadProperty("HotTracking", False)
    MyProp.cRank = PropBag.ReadProperty("Rank", crd_Ace)
    MyProp.cSelected = PropBag.ReadProperty("Selected", False)
    MyProp.cShowFocusRect = PropBag.ReadProperty("ShowFocusRect", False)
    MyProp.cSpeed = PropBag.ReadProperty("Speed", 1)
    MyProp.cSplit = PropBag.ReadProperty("Split", crd_Split_Horizontal_In)
    MyProp.cStretch = PropBag.ReadProperty("Stretch", crd_Stretch_From_Left)
    MyProp.cStopAni = PropBag.ReadProperty("StopAni", True)
    MyProp.cSuit = PropBag.ReadProperty("Suit", crd_Clubs)
    MyProp.cThreeD = PropBag.ReadProperty("ThreeD", crd_ThreeD_From_Left)
    MyProp.cUpdate = PropBag.ReadProperty("Update", True)
    MyProp.cValue = PropBag.ReadProperty("Value", 1)
    MyProp.cWipe = PropBag.ReadProperty("Wipe", crd_Wipe_Left)
    MyProp.cZoom = PropBag.ReadProperty("Zoom", crd_Zoom_In)
    
    Set picHolder.Font = PropBag.ReadProperty("Font", Ambient.Font)
    picHolder.FontBold = PropBag.ReadProperty("FontBold", False)
    picHolder.FontItalic = PropBag.ReadProperty("FontItalic", False)
    picHolder.FontName = PropBag.ReadProperty("FontName", "Times New Roman")
    picHolder.FontSize = PropBag.ReadProperty("FontSize", 12)
    picHolder.FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
    picHolder.FontUnderline = PropBag.ReadProperty("FontUnderline", False)
    picHolder.FontTransparent = PropBag.ReadProperty("FontTransparent", True)
    picHolder.ForeColor = PropBag.ReadProperty("ForeColor", &H0)
    
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_Resize()
    If CSH.bPlayMode And (MyProp.cEffect <> crd_Effect_Flip) Then
        StopAni = True
    Else
        Call RedrawCard
    End If
End Sub

Private Sub UserControl_Show()
    Call RedrawCard
End Sub

Private Sub UserControl_Terminate()
    If Not StopAni Then StopAni = True
    
    DeleteDC CSH.lhSrcDC
    DeleteDC CSH.lhSvBgDC
    
    RemoveProp UserControl.hWnd, "ClassID"
    RemoveProp UserControl.hWnd, "Focus"
    RemoveProp UserControl.hWnd, "TimerActive"
    RemoveProp UserControl.hWnd, "Track"

    KillTimer UserControl.hWnd, GetProp(UserControl.hWnd, "TimerAniID")
    KillTimer UserControl.hWnd, GetProp(UserControl.hWnd, "TimerHotTrackingID")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ArrowKeyFocus", MyProp.cArrowKeyFocus, True)
    Call PropBag.WriteProperty("AutoFlipCard", MyProp.cAutoFlipCard, True)
    Call PropBag.WriteProperty("BorderLine", MyProp.cBorderLine, True)
    Call PropBag.WriteProperty("Text", MyProp.sText, "")
    Call PropBag.WriteProperty("Data", MyProp.cData, "")
    Call PropBag.WriteProperty("Deck", MyProp.cDeck, crd_Deck_1)
    Call PropBag.WriteProperty("DeckBackground", MyProp.cDeckBackground, crd_BG_Stretch)
    Call PropBag.WriteProperty("DeckMaskPicture", MyProp.cDeckMaskPicture, Nothing)
    Call PropBag.WriteProperty("DeckPicture", MyProp.cDeckPicture, Nothing)
    Call PropBag.WriteProperty("DeckMaskStyle", MyProp.cDeckMaskStyle, crd_None)
    Call PropBag.WriteProperty("Effect", MyProp.cEffect, crd_Effect_None)
    Call PropBag.WriteProperty("Elevator", MyProp.cElevator, crd_Elevator_Left)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Face", MyProp.cFace, crd_Up)
    Call PropBag.WriteProperty("Flip", MyProp.cFlip, crd_Flip_Horizontal)
    Call PropBag.WriteProperty("FlyIn", MyProp.cFlyIn, crd_FlyIn_Left)
    Call PropBag.WriteProperty("FlyOut", MyProp.cFlyOut, crd_FlyOut_Left)
    Call PropBag.WriteProperty("FramePerMoveX", MyProp.cFramePerMoveX, 1)
    Call PropBag.WriteProperty("FramePerMoveY", MyProp.cFramePerMoveY, 1)
    Call PropBag.WriteProperty("Gate", MyProp.cGate, crd_Gate_Horizontal_In)
    Call PropBag.WriteProperty("HotTracking", MyProp.cHotTracking, False)
    Call PropBag.WriteProperty("Rank", MyProp.cRank, crd_Ace)
    Call PropBag.WriteProperty("Selected", MyProp.cSelected, False)
    Call PropBag.WriteProperty("ShowFocusRect", MyProp.cShowFocusRect, False)
    Call PropBag.WriteProperty("Speed", MyProp.cSpeed, 1)
    Call PropBag.WriteProperty("Split", MyProp.cSplit, crd_Split_Horizontal_In)
    Call PropBag.WriteProperty("Stretch", MyProp.cStretch, crd_Stretch_From_Left)
    Call PropBag.WriteProperty("StopAni", MyProp.cStopAni, True)
    Call PropBag.WriteProperty("Suit", MyProp.cSuit, crd_Clubs)
    Call PropBag.WriteProperty("ThreeD", MyProp.cThreeD, crd_ThreeD_From_Left)
    Call PropBag.WriteProperty("Update", MyProp.cUpdate, True)
    Call PropBag.WriteProperty("Value", MyProp.cValue, 1)
    Call PropBag.WriteProperty("Wipe", MyProp.cWipe, crd_Wipe_Left)
    Call PropBag.WriteProperty("Zoom", MyProp.cZoom, crd_Zoom_In)

    Call PropBag.WriteProperty("Font", picHolder.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", picHolder.FontBold, False)
    Call PropBag.WriteProperty("FontItalic", picHolder.FontItalic, False)
    Call PropBag.WriteProperty("FontName", picHolder.FontName, "Times New Roman")
    Call PropBag.WriteProperty("FontSize", picHolder.FontSize, 12)
    Call PropBag.WriteProperty("FontStrikethru", picHolder.FontStrikethru, False)
    Call PropBag.WriteProperty("FontUnderline", picHolder.FontUnderline, False)
    Call PropBag.WriteProperty("FontTransparent", picHolder.FontTransparent, True)
    Call PropBag.WriteProperty("ForeColor", picHolder.ForeColor, &H0)
    
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

Public Sub Refresh()
    Call RedrawCard
    UserControl.Refresh
End Sub

Friend Sub TimerAniUpdate()
    If IsIconic(GetParent(UserControl.hWnd)) Then
        ' when parent window is minimized stop Aniation
    Else
        If SlideShow Then
            CSH.bPlayMode = False
            KillTimer UserControl.hWnd, TimerAniID
            Call RedrawCard
            StopAni = True
        Else
            Set UserControl.Picture = UserControl.Image
        End If
        
        If GetProp(UserControl.hWnd, "Focus") And MyProp.cShowFocusRect Then
            Call DrawFocusRect
        End If
        
        If GetProp(UserControl.hWnd, "Track") And MyProp.cHotTracking Then
            Call DrawTrackRect
        End If
    End If
End Sub

Friend Sub TimerHotTrackingUpdate()
    Dim MousePT As POINTAPI
    
    GetCursorPos MousePT
    If WindowFromPoint(MousePT.x, MousePT.y) <> UserControl.hWnd Then
        KillTimer UserControl.hWnd, TimerHotTrackingID
        SetProp UserControl.hWnd, "TimerActive", False
    End If
    
    Static bTrack As Boolean
    
    If GetProp(UserControl.hWnd, "TimerActive") Then
        If Not bTrack Then
            Call DrawTrackRect
            bTrack = True
            SetProp UserControl.hWnd, "Track", bTrack
        End If
    Else
        If bTrack Then
            If GetProp(UserControl.hWnd, "Focus") And MyProp.cShowFocusRect Then
                Call DrawFocusRect
            Else
                UserControl.Cls
            End If
            bTrack = False
            SetProp UserControl.hWnd, "Track", bTrack
        End If
    End If
End Sub

Private Sub GetCard(lhDstDC As Long, FaceOp As FaceConstants, bSelected As Boolean, _
                                     Optional bUserRect As Boolean = False)
    Dim Buffer   As Picture
    Dim rcRect   As RECT
    Dim ResId    As Integer
    Dim Position As DeckBackgroundConstants
    
    If FaceOp = crd_Up Then
        ResId = MyProp.cSuit * Cards_Per_Suit + MyProp.cRank + OffsetRankResId
    Else
        If MyProp.cDeck <> crd_Deck_Customize Then
            ResId = MyProp.cDeck + OffsetDeckResId
        Else
            ResId = OffsetDeckResId
        End If
    End If
    
    If (MyProp.cDeck = crd_Deck_Customize) And (FaceOp = crd_Down) Then
        If MyProp.cDeckPicture.handle = 0 Then
            Set Buffer = LoadResPicture(ResId, vbResBitmap)
        Else
            Set Buffer = MyProp.cDeckPicture
        End If
    Else
        Set Buffer = LoadResPicture(ResId, vbResBitmap)
    End If
    
    Position = IIf(MyProp.cDeckBackground = crd_BG_Stretch, bgStretch, bgTile)
    
    If bUserRect Then
        CopyRect rcRect, CSH.rcRect
    Else
        GetClientRect UserControl.hWnd, rcRect
    End If
    
    If (MyProp.cDeckMaskStyle <> crd_None) And (FaceOp = crd_Down) Then
        Dim ResIDMask As Integer
        Dim MaskBMP   As IPicture
        
        If MyProp.cDeckMaskStyle <> crd_Text Then
            If MyProp.cDeckMaskStyle <> crd_Customize Then
                ResIDMask = MyProp.cDeckMaskStyle + OffsetMaskResId - 1
                Set MaskBMP = LoadResPicture(ResIDMask, vbResBitmap)
            Else
                If MyProp.cDeckMaskPicture.handle = 0 Then
                    Set MaskBMP = LoadResPicture(OffsetMaskResId, vbResBitmap)
                Else
                    Set MaskBMP = MyProp.cDeckMaskPicture
                End If
            End If
        Else
            Dim tSize As Size
            
            picHolder.Cls
            Set picHolder.Picture = Nothing
            
            If MyProp.sText <> "" Then
                With picHolder
                    MoveWindow .hWnd, -rcRect.Left, -rcRect.Top, _
                                       rcRect.Right, rcRect.Bottom, True
                    GetTextExtentPoint32 .hdc, MyProp.sText, Len(MyProp.sText), tSize
                    TextOut .hdc, (rcRect.Right - tSize.cx) / 2, _
                                  (rcRect.Bottom - tSize.cy) / 2, MyProp.sText, Len(MyProp.sText)
                End With
            End If
            
            Set MaskBMP = picHolder.Image
            
            If Not picHolder.FontTransparent Then
                Set Buffer = picHolder.Image
            End If
        End If
        CombinePicture lhDstDC, rcRect.Right, rcRect.Bottom, Buffer, MaskBMP, Position
    Else
        If MyProp.cFace = crd_Down Then
            Background lhDstDC, Buffer, rcRect.Left, rcRect.Top, _
                       rcRect.Right, rcRect.Bottom, Position, vbSrcCopy
        Else
            Background lhDstDC, Buffer, rcRect.Left, rcRect.Top, _
                       rcRect.Right, rcRect.Bottom, bgStretch, vbSrcCopy
        End If
    End If
    
    If MyProp.cBorderLine Then
        Dim rcClone As RECT
        
        CopyRect rcClone, rcRect
        DrawFrameRect lhDstDC, rcClone, &H0
        InflateRect rcClone, -1, -1
        DrawFrameRect lhDstDC, rcClone, &HFFFFFF
    End If
    
    If bSelected Then
        PatBlt lhDstDC, rcRect.Left, rcRect.Top, _
               rcRect.Right, rcRect.Bottom, vbDstInvert
    End If
End Sub

Private Sub RedrawCard()
    If Not Update Then Exit Sub
    
    If CSH.bPlayMode Then
    Else
        GetCard UserControl.hdc, MyProp.cFace, MyProp.cSelected
        RefreshWindow UserControl.hWnd
        Set UserControl.Picture = UserControl.Image
    End If
    
    If GetProp(UserControl.hWnd, "Focus") And MyProp.cShowFocusRect Then
        Call DrawFocusRect
    End If
        
    If MyProp.cHotTracking Then
        Call SetHotTracking
    End If
End Sub

Public Sub PlayAni()
    If CSH.bPlayMode Then Exit Sub
    
    CSH.lAniEffect = GetAniEffect
    CSH.bPlayMode = True: StopAni = False
    
    If MyProp.cEffect <> crd_Effect_Flip Then
        Dim lhDCC As Long
        Dim lhBmp As Long
        Dim tBMP  As BITMAP
        Dim Temp  As FaceConstants
        
        DeleteDC CSH.lhSrcDC
        GetClientRect UserControl.hWnd, CSH.rcRect
        GetObjectAPI UserControl.Picture.handle, Len(tBMP), tBMP
        
        CSH.lBmpW = tBMP.bmWidth
        CSH.lBmpH = tBMP.bmHeight
        CSH.lWinSzX = CSH.rcRect.Right
        CSH.lWinSzY = CSH.rcRect.Bottom
    
        lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        CSH.lhSrcDC = CreateCompatibleDC(lhDCC)
        lhBmp = CreateCompatibleBitmap(lhDCC, CSH.lBmpW, CSH.lBmpH)
        SelectObject CSH.lhSrcDC, lhBmp
        
        Temp = IIf(MyProp.cAutoFlipCard, FlipCard(MyProp.cFace), MyProp.cFace)
        GetCard CSH.lhSrcDC, Temp, MyProp.cSelected
    
        DeleteDC lhDCC
        DeleteObject lhBmp
        
        If (MyProp.cEffect = crd_Effect_ThreeD) Or _
           (MyProp.cEffect = crd_Effect_Elevator) Or _
           (MyProp.cEffect = crd_Effect_Gate) Or _
           (MyProp.cEffect = crd_Effect_FlyOut) Then
           
            Dim lhSaveBmp As Long
            
            DeleteDC CSH.lhSvBgDC
            lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
            CSH.lhSvBgDC = CreateCompatibleDC(lhDCC)
            lhSaveBmp = CreateCompatibleBitmap(lhDCC, CSH.lWinSzX, CSH.lWinSzY)
            SelectObject CSH.lhSvBgDC, lhSaveBmp
            
            BitBlt CSH.lhSvBgDC, 0, 0, CSH.lWinSzX, CSH.lWinSzY, UserControl.hdc, _
                   0, 0, vbSrcCopy
            
            DeleteDC lhDCC
            DeleteObject lhSaveBmp
        End If
    Else
        CSH.lWinOffX = UserControl.Extender.Left
        CSH.lWinOffY = UserControl.Extender.Top
        CSH.lWinSzX = UserControl.Extender.Width
        CSH.lWinSzY = UserControl.Extender.Height
        
        GetClientRect UserControl.hWnd, CSH.rcRect
        CSH.lBmpW = CSH.rcRect.Right
        CSH.lBmpH = CSH.rcRect.Bottom
    End If
    
    SetProp UserControl.hWnd, "ClassID", ObjPtr(Me)
    SetTimer UserControl.hWnd, TimerAniID, MyProp.cSpeed, AddressOf TimerCallBack
End Sub

Private Function ElevatorAni() As Boolean
    Static bInit As Boolean
    
    With CSH
        If Not bInit Then
            SetRect .rcDummy, 0, 0, .lBmpW, .lBmpH
            CopyRect .rcTemp, .rcDummy
            If .lAniEffect = crd_Elevator_Left Then
                OffsetRect .rcDummy, -.lBmpW, .rcRect.Top
            ElseIf .lAniEffect = crd_Elevator_Down Then
                OffsetRect .rcDummy, .rcRect.Left, -.lBmpH
            ElseIf .lAniEffect = crd_Elevator_Up Then
                OffsetRect .rcDummy, .rcRect.Left, .lWinSzY
                If (.rcDummy.Top < .rcRect.Top) Then
                    CopyRect .rcDummy, .rcRect
                End If
            Else
                OffsetRect .rcDummy, .lWinSzX, .rcRect.Top
                If (.rcDummy.Left < .rcRect.Left) Then
                    CopyRect .rcDummy, .rcRect
                End If
            End If
            bInit = True
            Exit Function
        Else
            If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                bInit = False: ElevatorAni = True
                If MyProp.cAutoFlipCard Then
                    Face = FlipCard(MyProp.cFace)
                End If
                Exit Function
            Else
                If .lAniEffect = crd_Elevator_Left Then
                    OffsetRect .rcDummy, MyProp.cFramePerMoveX, 0
                    OffsetRect .rcTemp, MyProp.cFramePerMoveX, 0
                    If (.rcDummy.Left > .rcRect.Left) Or (.rcDummy.Right > .rcRect.Right) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Elevator_Down Then
                    OffsetRect .rcDummy, 0, MyProp.cFramePerMoveY
                    OffsetRect .rcTemp, 0, MyProp.cFramePerMoveY
                    If (.rcDummy.Top > .rcRect.Top) Or (.rcDummy.Bottom > .rcRect.Bottom) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Elevator_Up Then
                    OffsetRect .rcDummy, 0, -MyProp.cFramePerMoveY
                    OffsetRect .rcTemp, 0, -MyProp.cFramePerMoveY
                    If (.rcDummy.Top < .rcRect.Top) Or (.rcDummy.Bottom < .rcRect.Bottom) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                Else
                    OffsetRect .rcDummy, -MyProp.cFramePerMoveX, 0
                    OffsetRect .rcTemp, -MyProp.cFramePerMoveX, 0
                    If (.rcDummy.Left < .rcRect.Left) Or (.rcDummy.Right < .rcRect.Right) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                End If
            End If
        End If
            
        If (.rcDummy.Left < -.lBmpW) Or (.rcDummy.Top < -.lBmpH) Then
            CopyRect .rcDummy, .rcRect
        End If
        
        If EqualRect(.rcRect, .rcDummy) Then
            ' do nothing
        Else
            BitBlt UserControl.hdc, .rcDummy.Left, .rcDummy.Top, .lBmpW, .lBmpH, _
                  .lhSrcDC, 0, 0, vbSrcCopy
            BitBlt UserControl.hdc, .rcTemp.Left, .rcTemp.Top, .lBmpW, .lBmpH, _
                  .lhSvBgDC, 0, 0, vbSrcCopy
                  
            RaiseEvent AniPosition(ScaleX(CSng(.rcDummy.Left), vbPixels, vbTwips), _
                                    ScaleY(CSng(.rcDummy.Top), vbPixels, vbTwips), _
                                    ScaleX(CSng(.lBmpW), vbPixels, vbTwips), _
                                    ScaleY(CSng(.lBmpH), vbPixels, vbTwips))
        End If
    End With
End Function

Private Function FlipAni() As Boolean
    Dim sx, sy       As Long
    Dim lOffsetX     As Long
    Dim lOffsetY     As Long
    Dim lSizeX       As Long
    Dim lSizeY       As Long
    Static bFlipCard As Boolean
    Static bInit     As Boolean
    
    On Error GoTo ErrHandler
    
    If Not bInit Then
        Dim lhDCC As Long
        Dim lhBmp As Long
        
        DeleteDC CSH.lhSrcDC
        lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        CSH.lhSrcDC = CreateCompatibleDC(lhDCC)
        lhBmp = CreateCompatibleBitmap(lhDCC, CSH.lBmpW, CSH.lBmpH)
        SelectObject CSH.lhSrcDC, lhBmp
        GetCard CSH.lhSrcDC, MyProp.cFace, MyProp.cSelected, True
        DeleteDC lhDCC
        DeleteObject lhBmp
        
        bInit = True
    End If
    
    sx = ScaleX(MyProp.cFramePerMoveX, vbPixels, vbTwips)
    sy = ScaleY(MyProp.cFramePerMoveY, vbPixels, vbTwips)
    
    With UserControl.Extender
        Select Case CSH.lAniEffect
        Case Is = crd_Flip_Left
            If Not bFlipCard Then
                lOffsetX = .Left
                lOffsetY = .Top
                lSizeX = .Width - sx * 2
                lSizeY = .Height
                If lSizeX > 0 Then
                    ' do nothing
                Else
                    lSizeX = 0
                    bFlipCard = True
                    If MyProp.cAutoFlipCard Then
                        Face = FlipCard(MyProp.cFace)
                        bInit = False
                    End If
                End If
            Else
                lOffsetX = .Left
                lOffsetY = .Top
                lSizeX = .Width + sx * 2
                lSizeY = .Height
                If lSizeX <= CSH.lWinSzX Then
                    ' do nothing
                Else
                    lSizeX = CSH.lWinSzX
                    FlipAni = True: bFlipCard = False
                    bInit = False
                End If
            End If
        Case Is = crd_Flip_Top
            If Not bFlipCard Then
                lOffsetX = .Left
                lOffsetY = .Top
                lSizeX = .Width
                lSizeY = .Height - sy * 2
                If lSizeY > 0 Then
                    ' do nothing
                Else
                    lSizeY = 0
                    bFlipCard = True
                    If MyProp.cAutoFlipCard Then
                        Face = FlipCard(MyProp.cFace)
                        bInit = False
                    End If
                End If
            Else
                lOffsetX = .Left
                lOffsetY = .Top
                lSizeX = .Width
                lSizeY = .Height + sy * 2
                If lSizeY <= CSH.lWinSzY Then
                    ' do nothing
                Else
                    lSizeY = CSH.lWinSzY
                    FlipAni = True: bFlipCard = False
                    bInit = False
                End If
            End If
        Case Is = crd_Flip_Right
            If Not bFlipCard Then
                lSizeX = .Width - sx * 2
                lSizeY = .Height
                lOffsetX = CSH.lWinOffX + (CSH.lWinSzX - lSizeX)
                lOffsetY = .Top
                
                If lSizeX > 0 Then
                    ' do nothing
                Else
                    lOffsetX = CSH.lWinSzX
                    lSizeX = 0
                    bFlipCard = True
                    If MyProp.cAutoFlipCard Then
                        Face = FlipCard(MyProp.cFace)
                        bInit = False
                    End If
                End If
            Else
                lSizeX = .Width + sx * 2
                lSizeY = .Height
                lOffsetX = CSH.lWinOffX + (CSH.lWinSzX - lSizeX)
                lOffsetY = .Top
                If lSizeX <= CSH.lWinSzX Then
                   ' do nothing
                Else
                    lOffsetX = CSH.lWinOffX
                    lSizeX = CSH.lWinSzX
                    FlipAni = True: bFlipCard = False
                    bInit = False
                End If
            End If
        Case Is = crd_Flip_Bottom
            If Not bFlipCard Then
                lSizeX = .Width
                lSizeY = .Height - sy * 2
                lOffsetX = .Left
                lOffsetY = CSH.lWinOffY + (CSH.lWinSzY - lSizeY)
                
                If lSizeY > 0 Then
                    ' do nothing
                Else
                    lOffsetY = CSH.lWinSzY
                    lSizeY = 0
                    bFlipCard = True
                    If MyProp.cAutoFlipCard Then
                        Face = FlipCard(MyProp.cFace)
                        bInit = False
                    End If
                End If
            Else
                lSizeX = .Width
                lSizeY = .Height + sy * 2
                lOffsetX = .Left
                lOffsetY = CSH.lWinOffY + (CSH.lWinSzY - lSizeY)
                If lSizeY <= CSH.lWinSzY Then
                   ' do nothing
                Else
                    lOffsetY = CSH.lWinOffY
                    lSizeY = CSH.lWinSzY
                    FlipAni = True: bFlipCard = False
                    bInit = False
                End If
            End If
        Case Is = crd_Flip_Horizontal
            If Not bFlipCard Then
                lOffsetX = .Left + sx * 2
                lOffsetY = .Top
                lSizeX = .Width - sx * 4
                lSizeY = .Height
                If lSizeX > 0 Then
                    ' do nothing
                Else
                    lOffsetX = CSH.lWinOffX + CSH.lWinSzX / 2
                    lSizeX = 0
                    bFlipCard = True
                    If MyProp.cAutoFlipCard Then
                        Face = FlipCard(MyProp.cFace)
                        bInit = False
                    End If
                End If
            Else
                lOffsetX = .Left - sx * 2
                lOffsetY = .Top
                lSizeX = .Width + sx * 4
                lSizeY = .Height
                If lSizeX <= CSH.lWinSzX Then
                    ' do nothing
                Else
                    lOffsetX = CSH.lWinOffX
                    lSizeX = CSH.lWinSzX
                    FlipAni = True: bFlipCard = False
                    bInit = False
                End If
            End If
        Case Is = crd_Flip_Vertical
            If Not bFlipCard Then
                lOffsetX = .Left
                lOffsetY = .Top + sy * 2
                lSizeX = .Width
                lSizeY = .Height - sy * 4
                If lSizeY > 0 Then
                    ' do nothing
                Else
                    lOffsetY = CSH.lWinOffY + CSH.lWinSzY / 2
                    lSizeY = 0
                    bFlipCard = True
                    If MyProp.cAutoFlipCard Then
                        Face = FlipCard(MyProp.cFace)
                        bInit = False
                    End If
                End If
            Else
                lOffsetX = .Left
                lOffsetY = .Top - sy * 2
                lSizeX = .Width
                lSizeY = .Height + sy * 4
                If lSizeY <= CSH.lWinSzY Then
                    ' do nothing
                Else
                    lOffsetY = CSH.lWinOffY
                    lSizeY = CSH.lWinSzY
                    FlipAni = True: bFlipCard = False
                    bInit = False
                End If
            End If
        End Select
        
        .Move lOffsetX, lOffsetY, lSizeX, lSizeY
        GetClientRect UserControl.hWnd, CSH.rcTemp
        StretchImage UserControl.hdc, 0, 0, CSH.rcTemp.Right, CSH.rcTemp.Bottom, _
                     CSH.lhSrcDC, 0, 0, CSH.rcRect.Right, CSH.rcRect.Bottom, vbSrcCopy
        RefreshWindow UserControl.hWnd
            
        If StopAni Then
            UserControl.Extender.Left = CSH.lWinOffX
            UserControl.Extender.Top = CSH.lWinOffY
            UserControl.Extender.Width = CSH.lWinSzX
            UserControl.Extender.Height = CSH.lWinSzY
            FlipAni = True: bFlipCard = False
        End If
        
        RaiseEvent AniPosition(CSng(lOffsetX), CSng(lOffsetY), _
                                CSng(lSizeX), CSng(lSizeY))
    End With
    Exit Function

ErrHandler:
    FlipAni = True
End Function

Private Function FlyInAni() As Boolean
    Static bInit As Boolean
    
    With CSH
        Select Case .lAniEffect
        Case Is = crd_FlyIn_Left, crd_FlyIn_Right, crd_FlyIn_Top, crd_FlyIn_Bottom
            If Not bInit Then
                SetRect .rcDummy, 0, 0, .lBmpW, .lBmpH
                If .lAniEffect = crd_FlyIn_Left Then
                    OffsetRect .rcDummy, -.lBmpW, .rcRect.Top
                ElseIf .lAniEffect = crd_FlyIn_Top Then
                    OffsetRect .rcDummy, .rcRect.Left, -.lBmpH
                ElseIf .lAniEffect = crd_FlyIn_Right Then
                    OffsetRect .rcDummy, .lWinSzX, .rcRect.Top
                    If (.rcDummy.Left < .rcRect.Left) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                Else
                    OffsetRect .rcDummy, .rcRect.Left, .lWinSzY
                    If (.rcDummy.Top < .rcRect.Top) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                End If
                bInit = True
                Exit Function
            Else
                If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                    bInit = False: FlyInAni = True
                    If MyProp.cAutoFlipCard Then
                        Face = FlipCard(MyProp.cFace)
                    End If
                    Exit Function
                Else
                    If .lAniEffect = crd_FlyIn_Left Then
                        OffsetRect .rcDummy, MyProp.cFramePerMoveX, 0
                        If (.rcDummy.Left > .rcRect.Left) Or (.rcDummy.Right > .rcRect.Right) Then
                            CopyRect .rcDummy, .rcRect
                        End If
                    ElseIf .lAniEffect = crd_FlyIn_Top Then
                        OffsetRect .rcDummy, 0, MyProp.cFramePerMoveY
                        If (.rcDummy.Top > .rcRect.Top) Or (.rcDummy.Bottom > .rcRect.Bottom) Then
                            CopyRect .rcDummy, .rcRect
                        End If
                    ElseIf .lAniEffect = crd_FlyIn_Right Then
                        OffsetRect .rcDummy, -MyProp.cFramePerMoveX, 0
                        If (.rcDummy.Left < .rcRect.Left) Or (.rcDummy.Right < .rcRect.Right) Then
                            CopyRect .rcDummy, .rcRect
                        End If
                    Else
                        OffsetRect .rcDummy, 0, -MyProp.cFramePerMoveY
                        If (.rcDummy.Top < .rcRect.Top) Or (.rcDummy.Bottom < .rcRect.Bottom) Then
                            CopyRect .rcDummy, .rcRect
                        End If
                    End If
                End If
            End If
            
            If (.rcDummy.Left < -.lBmpW) Or (.rcDummy.Top < -.lBmpH) Then
                CopyRect .rcDummy, .rcRect
            End If
            
            If EqualRect(.rcRect, .rcDummy) Then
                ' do nothing
            Else
                BitBlt UserControl.hdc, .rcDummy.Left, .rcDummy.Top, .lBmpW, .lBmpH, _
                      .lhSrcDC, 0, 0, vbSrcCopy
            
                RaiseEvent AniPosition(ScaleX(CSng(.rcDummy.Left), vbPixels, vbTwips), _
                                        ScaleY(CSng(.rcDummy.Top), vbPixels, vbTwips), _
                                        ScaleX(CSng(.lBmpW), vbPixels, vbTwips), _
                                        ScaleY(CSng(.lBmpH), vbPixels, vbTwips))
            End If
        Case Is = crd_FlyIn_Top_Left, crd_FlyIn_Top_Right, crd_FlyIn_Bottom_Left, crd_FlyIn_Bottom_Right
            Static step As Integer
            
            If Not bInit Then
                If .lAniEffect = crd_FlyIn_Top_Left Then
                    SetRect .rcDummy, -.lBmpW, -.lBmpH, .lBmpW, .lBmpH
                ElseIf .lAniEffect = crd_FlyIn_Top_Right Then
                    SetRect .rcDummy, .lWinSzX, -.lBmpH, .lBmpW, .lBmpH
                ElseIf .lAniEffect = crd_FlyIn_Bottom_Left Then
                    SetRect .rcDummy, -.lBmpW, .lWinSzY, .lBmpW, .lBmpH
                ElseIf .lAniEffect = crd_FlyIn_Bottom_Right Then
                    SetRect .rcDummy, .lWinSzX, .lWinSzY, .lBmpW, .lBmpH
                End If
                
                bInit = True
                Exit Function
            Else
                Dim ll As Long
                Dim ls As Long
                Dim sx As Single
                Dim sy As Single
                
                ll = Largest(.lBmpW, .lBmpH)
                ls = Smallest(.lBmpW, .lBmpH)
        
                If .lBmpW <= .lBmpH Then
                    sx = 1
                    sy = CSng(ll / ls)
                Else
                    sx = CSng(ll / ls)
                    sy = 1
                End If
                
                If .lAniEffect = crd_FlyIn_Top_Left Then
                    sx = .rcDummy.Left + sx * step
                    sy = .rcDummy.Top + sy * step
                    If (sx > .rcRect.Left) Or (sy > .rcRect.Top) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_FlyIn_Top_Right Then
                    sx = .rcDummy.Left - sx * step
                    sy = .rcDummy.Top + sy * step
                    If (sx < .rcRect.Left) Or (sy > .rcRect.Top) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_FlyIn_Bottom_Left Then
                    sx = .rcDummy.Left + sx * step
                    sy = .rcDummy.Top - sy * step
                    If (sx > .rcRect.Left) Or (sy < .rcRect.Top) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                Else
                    sx = .rcDummy.Left - sx * step
                    sy = .rcDummy.Top - sy * step
                        If (sx < .rcRect.Left) Or (sy < .rcRect.Top) Then
                            CopyRect .rcDummy, .rcRect
                        End If
                End If
                
                step = step + MyProp.cFramePerMoveX
            End If
            
            If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                step = 0
                bInit = False: FlyInAni = True
                sx = .rcRect.Left: sy = .rcRect.Top
                If MyProp.cAutoFlipCard Then
                    Face = FlipCard(MyProp.cFace)
                End If
                Exit Function
            End If
        
            If EqualRect(.rcRect, .rcDummy) Then
                ' do nothing
            Else
                BitBlt UserControl.hdc, sx, sy, .lBmpW, .lBmpH, _
                       .lhSrcDC, 0, 0, vbSrcCopy
                                       
                RaiseEvent AniPosition(ScaleX(CSng(sx), vbPixels, vbTwips), _
                                        ScaleY(CSng(sy), vbPixels, vbTwips), _
                                        ScaleX(CSng(.lBmpW), vbPixels, vbTwips), _
                                        ScaleY(CSng(.lBmpH), vbPixels, vbTwips))
            End If
        End Select
    End With
End Function

Private Function FlyOutAni() As Boolean
    Static bInit As Boolean
    
    With CSH
        If Not bInit Then
            CopyRect .rcTemp, .rcRect
            SetRectEmpty .rcRect
            SetRect .rcDummy, .rcRect.Left, .rcRect.Top, .lBmpW, .lBmpH
            bInit = True
            Exit Function
        End If
    
        Select Case .lAniEffect
        Case Is = crd_FlyOut_Left, crd_FlyOut_Right, crd_FlyOut_Top, crd_FlyOut_Bottom
            If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                bInit = False: FlyOutAni = True
                If MyProp.cAutoFlipCard Then
                    Face = FlipCard(MyProp.cFace)
                End If
                Exit Function
            Else
                If .lAniEffect = crd_FlyOut_Left Then
                    OffsetRect .rcDummy, -MyProp.cFramePerMoveX, 0
                    If (.rcDummy.Left < -.lBmpW) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_FlyOut_Top Then
                    OffsetRect .rcDummy, 0, -MyProp.cFramePerMoveY
                    If (.rcDummy.Top < -.lBmpH) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_FlyOut_Right Then
                    OffsetRect .rcDummy, MyProp.cFramePerMoveX, 0
                    If (.rcDummy.Left > .lBmpW) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                Else
                    OffsetRect .rcDummy, 0, MyProp.cFramePerMoveY
                    If (.rcDummy.Top > .lBmpH) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                End If
            End If
                    
            If EqualRect(.rcRect, .rcDummy) Then
                ' do nothing
            Else
                      
                BitBlt UserControl.hdc, .rcTemp.Left, .rcTemp.Top, .lBmpW, .lBmpH, _
                      .lhSrcDC, 0, 0, vbSrcCopy
                BitBlt UserControl.hdc, .rcDummy.Left, .rcDummy.Top, .lBmpW, .lBmpH, _
                      .lhSvBgDC, 0, 0, vbSrcCopy
            
                RaiseEvent AniPosition(ScaleX(CSng(.rcDummy.Left), vbPixels, vbTwips), _
                                        ScaleY(CSng(.rcDummy.Top), vbPixels, vbTwips), _
                                        ScaleX(CSng(.lBmpW), vbPixels, vbTwips), _
                                        ScaleY(CSng(.lBmpH), vbPixels, vbTwips))
            End If
        Case Is = crd_FlyOut_Top_Left, crd_FlyOut_Top_Right, crd_FlyOut_Bottom_Left, crd_FlyOut_Bottom_Right
            Static step As Integer
            
            Dim ll As Long
            Dim ls As Long
            Dim sx As Single
            Dim sy As Single
                
            ll = Largest(.lBmpW, .lBmpH)
            ls = Smallest(.lBmpW, .lBmpH)
        
            If .lBmpW <= .lBmpH Then
                sx = 1
                sy = CSng(ll / ls)
            Else
                sx = CSng(ll / ls)
                sy = 1
            End If
                
            If .lAniEffect = crd_FlyOut_Top_Left Then
                sx = .rcDummy.Left - sx * step
                sy = .rcDummy.Top - sy * step
                If (sx < -.lBmpW) Or (sy < -.lBmpH) Then
                    CopyRect .rcDummy, .rcRect
                End If
            ElseIf .lAniEffect = crd_FlyOut_Top_Right Then
                sx = .rcDummy.Left + sx * step
                sy = .rcDummy.Top - sy * step
                If (sx > .lBmpW) Or (sy > .lBmpH) Then
                    CopyRect .rcDummy, .rcRect
                End If
            ElseIf .lAniEffect = crd_FlyOut_Bottom_Left Then
                sx = .rcDummy.Left - sx * step
                sy = .rcDummy.Top + sy * step
                If (sx < -.lBmpW) Or (sy > .lBmpH) Then
                    CopyRect .rcDummy, .rcRect
                End If
            Else
                sx = .rcDummy.Left + sx * step
                sy = .rcDummy.Top + sy * step
                If (sx > .lBmpW) Or (sy > .lBmpH) Then
                    CopyRect .rcDummy, .rcRect
                End If
            End If
                
            step = step + MyProp.cFramePerMoveX
            
            If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                step = 0
                bInit = False: FlyOutAni = True
                sx = .rcRect.Left: sy = .rcRect.Top
                If MyProp.cAutoFlipCard Then
                    Face = FlipCard(MyProp.cFace)
                End If
                Exit Function
            End If
            
            If EqualRect(.rcRect, .rcDummy) Then
                ' do nothing
            Else
                BitBlt UserControl.hdc, .rcTemp.Left, .rcTemp.Top, .lBmpW, .lBmpH, _
                      .lhSrcDC, 0, 0, vbSrcCopy
                BitBlt UserControl.hdc, sx, sy, .lBmpW, .lBmpH, _
                     .lhSvBgDC, 0, 0, vbSrcCopy
                
                RaiseEvent AniPosition(ScaleX(CSng(sx), vbPixels, vbTwips), _
                                        ScaleY(CSng(sy), vbPixels, vbTwips), _
                                        ScaleX(CSng(.lBmpW), vbPixels, vbTwips), _
                                        ScaleY(CSng(.lBmpH), vbPixels, vbTwips))
            End If
        End Select
    End With
End Function

Private Function GateAni() As Boolean
    Static bInit     As Boolean
    Static lTempBmpW As Long
    Static lTempBmpH As Long
    
    With CSH
        If Not bInit Then
            If .lAniEffect = crd_Gate_Horizontal_In Then
                lTempBmpW = .lBmpW: lTempBmpH = .lBmpH / 2
                SetRect .rcDummy, 0, 0, lTempBmpW, lTempBmpH
                CopyRect .rcTemp, .rcDummy
                
                OffsetRect .rcDummy, .rcRect.Left, -lTempBmpH
                OffsetRect .rcTemp, .rcRect.Left, CSH.lWinSzY
            ElseIf .lAniEffect = crd_Gate_Horizontal_Out Then
                lTempBmpW = .lBmpW: lTempBmpH = .lBmpH / 2
                SetRect .rcDummy, 0, 0, lTempBmpW, lTempBmpH
                CopyRect .rcTemp, .rcDummy
                
                OffsetRect .rcDummy, .rcRect.Left, 0
                OffsetRect .rcTemp, .rcRect.Left, lTempBmpH
            ElseIf .lAniEffect = crd_Gate_Vertical_In Then
                lTempBmpW = .lBmpW / 2: lTempBmpH = .lBmpH
                SetRect .rcDummy, 0, 0, lTempBmpW, lTempBmpH
                CopyRect .rcTemp, .rcDummy
                
                OffsetRect .rcDummy, -lTempBmpW, .rcRect.Top
                OffsetRect .rcTemp, CSH.lWinSzX, .rcRect.Top
            Else
                lTempBmpW = .lBmpW / 2: lTempBmpH = .lBmpH
                SetRect .rcDummy, 0, 0, lTempBmpW, lTempBmpH
                CopyRect .rcTemp, .rcDummy
                
                OffsetRect .rcDummy, 0, .rcRect.Top
                OffsetRect .rcTemp, lTempBmpW, .rcRect.Top
            End If
            bInit = True
            Exit Function
        Else
            If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                bInit = False: GateAni = True
                lTempBmpW = 0: lTempBmpH = 0
                If MyProp.cAutoFlipCard Then
                    Face = FlipCard(MyProp.cFace)
                End If
                Exit Function
            Else
                Dim rcBuffer     As RECT
                Dim lOffX        As Long
                Dim lOffY        As Long
                
                If .lAniEffect = crd_Gate_Horizontal_In Then
                    lOffX = 0: lOffY = lTempBmpH
                    OffsetRect .rcDummy, 0, MyProp.cFramePerMoveY
                    OffsetRect .rcTemp, 0, -MyProp.cFramePerMoveY
                    
                    If IntersectRect(rcBuffer, .rcDummy, .rcTemp) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Gate_Horizontal_Out Then
                    lOffX = 0: lOffY = lTempBmpH
                    OffsetRect .rcDummy, 0, -MyProp.cFramePerMoveY
                    OffsetRect .rcTemp, 0, MyProp.cFramePerMoveY
                    
                    If (.rcDummy.Top < -lTempBmpH) Or StopAni Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Gate_Vertical_In Then
                    lOffX = lTempBmpW: lOffY = 0
                    OffsetRect .rcDummy, MyProp.cFramePerMoveX, 0
                    OffsetRect .rcTemp, -MyProp.cFramePerMoveX, 0
                    
                    If IntersectRect(rcBuffer, .rcDummy, .rcTemp) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                Else
                    lOffX = lTempBmpW: lOffY = 0
                    OffsetRect .rcDummy, -MyProp.cFramePerMoveX, 0
                    OffsetRect .rcTemp, MyProp.cFramePerMoveX, 0
                    
                    If (.rcDummy.Left < -lTempBmpW) Or StopAni Then
                        CopyRect .rcDummy, .rcRect
                    End If
                End If
            End If
        End If
        
        If EqualRect(.rcRect, .rcDummy) Then
            ' do nothing
        Else
            If (.lAniEffect = crd_Gate_Horizontal_In) Or _
               (.lAniEffect = crd_Gate_Vertical_In) Then
                       
                BitBlt UserControl.hdc, .rcRect.Left, .rcRect.Top, .lBmpW, .lBmpH, _
                       .lhSvBgDC, 0, 0, vbSrcCopy
                BitBlt UserControl.hdc, .rcDummy.Left, .rcDummy.Top, lTempBmpW, lTempBmpH, _
                       .lhSrcDC, 0, 0, vbSrcCopy
                BitBlt UserControl.hdc, .rcTemp.Left, .rcTemp.Top, lTempBmpW, lTempBmpH, _
                       .lhSrcDC, lOffX, lOffY, vbSrcCopy
            Else
       
                BitBlt UserControl.hdc, .rcRect.Left, .rcRect.Top, .lBmpW, .lBmpH, _
                       .lhSrcDC, 0, 0, vbSrcCopy
                BitBlt UserControl.hdc, .rcDummy.Left, .rcDummy.Top, lTempBmpW, lTempBmpH, _
                       .lhSvBgDC, 0, 0, vbSrcCopy
                BitBlt UserControl.hdc, .rcTemp.Left, .rcTemp.Top, lTempBmpW, lTempBmpH, _
                       .lhSvBgDC, lOffX, lOffY, vbSrcCopy
            End If
        End If
                
        RaiseEvent AniPosition(ScaleX(CSng(.rcDummy.Left), vbPixels, vbTwips), _
                                ScaleY(CSng(.rcDummy.Top), vbPixels, vbTwips), _
                                ScaleX(CSng(lTempBmpW), vbPixels, vbTwips), _
                                ScaleY(CSng(lTempBmpH), vbPixels, vbTwips))
    End With
End Function

Private Function SplitAni() As Boolean
    Static bInit  As Boolean
    Static rcTemp As RECT
    
    With CSH
        If Not bInit Then
            If .lAniEffect = crd_Split_Horizontal_In Then
                SetRect .rcDummy, 0, 0, .lBmpW, 0
                SetRect rcTemp, 0, .lBmpH, .lBmpW, 0
            ElseIf .lAniEffect = crd_Split_Horizontal_Out Then
                SetRect .rcDummy, 0, .lBmpH / 2, .lBmpW, 0
            ElseIf .lAniEffect = crd_Split_Vertical_In Then
                SetRect .rcDummy, 0, 0, 0, .lBmpH
                SetRect rcTemp, .lBmpW, 0, 0, .lBmpH
            Else
                SetRect .rcDummy, .lBmpW / 2, 0, 0, .lBmpH
            End If
            bInit = True
            Exit Function
        Else
            If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                bInit = False: SplitAni = True
                If MyProp.cAutoFlipCard Then
                    Face = FlipCard(MyProp.cFace)
                End If
                Exit Function
            Else
                Dim rcBuffer As RECT
                
                If .lAniEffect = crd_Split_Horizontal_In Then
                    InflateRect .rcDummy, 0, MyProp.cFramePerMoveY
                    .rcDummy.Top = 0
                    InflateRect rcTemp, 0, MyProp.cFramePerMoveY
                        
                    If IntersectRect(rcBuffer, .rcDummy, rcTemp) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Split_Horizontal_Out Then
                    InflateRect .rcDummy, 0, MyProp.cFramePerMoveY
                    .rcDummy.Top = (.lBmpH - .rcDummy.Bottom) / 2
                    If .rcDummy.Bottom > .lBmpH Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Split_Vertical_In Then
                    InflateRect .rcDummy, MyProp.cFramePerMoveX, 0
                    .rcDummy.Left = 0
                    InflateRect rcTemp, MyProp.cFramePerMoveX, 0
                    
                    If IntersectRect(rcBuffer, .rcDummy, rcTemp) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                Else
                    InflateRect .rcDummy, MyProp.cFramePerMoveX, 0
                    .rcDummy.Left = (.lBmpW - .rcDummy.Right) / 2
                    If .rcDummy.Right > .lBmpW Then
                        CopyRect .rcDummy, .rcRect
                    End If
                End If
            End If
        End If
        
        If EqualRect(.rcRect, .rcDummy) Then
            ' do nothing
        Else
            BitBlt UserControl.hdc, .rcDummy.Left, .rcDummy.Top, .rcDummy.Right, .rcDummy.Bottom, _
                  .lhSrcDC, .rcDummy.Left, .rcDummy.Top, vbSrcCopy
        
            If (.lAniEffect = crd_Split_Horizontal_In) Or _
               (.lAniEffect = crd_Split_Vertical_In) Then
           
                BitBlt UserControl.hdc, rcTemp.Left, rcTemp.Top, rcTemp.Right, rcTemp.Bottom, _
                      .lhSrcDC, rcTemp.Left, rcTemp.Top, vbSrcCopy
            End If
        
            RaiseEvent AniPosition(ScaleX(CSng(.rcDummy.Left), vbPixels, vbTwips), _
                                    ScaleY(CSng(.rcDummy.Top), vbPixels, vbTwips), _
                                    ScaleX(CSng(.rcDummy.Right), vbPixels, vbTwips), _
                                    ScaleY(CSng(.rcDummy.Bottom), vbPixels, vbTwips))
        End If
    End With
End Function

Private Function StretchAni() As Boolean
    Static bInit As Boolean
    
    With CSH
        If Not bInit Then
            If .lAniEffect = crd_Stretch_Across Then
                SetRect .rcDummy, .lBmpW / 2, 0, 0, .lBmpH
            ElseIf .lAniEffect = crd_Stretch_From_Left Then
                SetRect .rcDummy, 0, 0, 0, .lBmpH
            ElseIf .lAniEffect = crd_Stretch_From_Top Then
                SetRect .rcDummy, 0, .lBmpH, .lBmpW, 0
            ElseIf .lAniEffect = crd_Stretch_From_Right Then
                SetRect .rcDummy, .lBmpW, 0, 0, .lBmpH
            Else
                SetRect .rcDummy, 0, .lBmpH, .lBmpW, 0
            End If
            bInit = True
            Exit Function
        Else
            If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                bInit = False: StretchAni = True
                If MyProp.cAutoFlipCard Then
                    Face = FlipCard(MyProp.cFace)
                End If
                Exit Function
            Else
                If .lAniEffect = crd_Stretch_Across Then
                    InflateRect .rcDummy, MyProp.cFramePerMoveX, 0
                    .rcDummy.Left = (.lBmpW - .rcDummy.Right) / 2
                    If .rcDummy.Right > .lBmpW Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Stretch_From_Left Then
                    InflateRect .rcDummy, MyProp.cFramePerMoveX, 0
                    SetRect .rcDummy, 0, 0, .rcDummy.Right, .rcDummy.Bottom
                    If .rcDummy.Right > .lBmpW Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Stretch_From_Right Then
                    InflateRect .rcDummy, MyProp.cFramePerMoveX, 0
                    .rcDummy.Left = .lBmpW - .rcDummy.Right
                    If .rcDummy.Right > .lBmpW Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Stretch_From_Top Then
                    InflateRect .rcDummy, 0, MyProp.cFramePerMoveY
                    SetRect .rcDummy, 0, 0, .rcDummy.Right, .rcDummy.Bottom
                    If .rcDummy.Bottom > .lBmpH Then
                        CopyRect .rcDummy, .rcRect
                    End If
                Else
                    InflateRect .rcDummy, 0, MyProp.cFramePerMoveY
                    .rcDummy.Top = .lBmpH - .rcDummy.Bottom
                    If .rcDummy.Bottom > .lBmpH Then
                        CopyRect .rcDummy, .rcRect
                    End If
                End If
            End If
        End If
        
        If EqualRect(.rcRect, .rcDummy) Then
            ' do nothing
        Else
            StretchImage UserControl.hdc, .rcDummy.Left, .rcDummy.Top, .rcDummy.Right, .rcDummy.Bottom, _
                        .lhSrcDC, 0, 0, .lBmpW, .lBmpH, vbSrcCopy
        
            RaiseEvent AniPosition(ScaleX(CSng(.rcDummy.Left), vbPixels, vbTwips), _
                                   ScaleY(CSng(.rcDummy.Top), vbPixels, vbTwips), _
                                   ScaleX(CSng(.rcDummy.Right), vbPixels, vbTwips), _
                                   ScaleY(CSng(.rcDummy.Bottom), vbPixels, vbTwips))
        End If
    End With
End Function

Private Function ThreeDAni() As Boolean
    Static bInit As Boolean
    
    With CSH
        If Not bInit Then
            If .lAniEffect = crd_ThreeD_From_Left Then
                SetRect .rcDummy, 0, 0, 0, .lBmpH
                SetRect .rcTemp, 0, 0, .lBmpW, .lBmpH
            ElseIf .lAniEffect = crd_ThreeD_From_Right Then
                SetRect .rcDummy, 0, 0, 0, .lBmpH
                SetRect .rcTemp, 0, 0, .lBmpW, .lBmpH
            ElseIf .lAniEffect = crd_ThreeD_From_Bottom Then
                SetRect .rcDummy, 0, .lBmpH, .lBmpW, 0
                SetRect .rcTemp, 0, 0, .lBmpW, .lBmpH
            Else
                SetRect .rcDummy, 0, .lBmpH, .lBmpW, 0
                SetRect .rcTemp, 0, 0, .lBmpW, .lBmpH
            End If
            bInit = True
            Exit Function
        Else
            If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                bInit = False: ThreeDAni = True
                If MyProp.cAutoFlipCard Then
                    Face = FlipCard(MyProp.cFace)
                End If
                Exit Function
            Else
                If .lAniEffect = crd_ThreeD_From_Left Then
                    InflateRect .rcTemp, -MyProp.cFramePerMoveX, 0
                    SetRect .rcTemp, 0, 0, .rcTemp.Right, .rcTemp.Bottom
                    
                    InflateRect .rcDummy, MyProp.cFramePerMoveX, 0
                    .rcDummy.Left = .lBmpW - .rcDummy.Right
                    If .rcDummy.Right > .lBmpW Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_ThreeD_From_Right Then
                    InflateRect .rcTemp, -MyProp.cFramePerMoveX, 0
                    .rcTemp.Left = .lBmpW - .rcTemp.Right
                    
                    InflateRect .rcDummy, MyProp.cFramePerMoveX, 0
                    SetRect .rcDummy, 0, 0, .rcDummy.Right, .rcDummy.Bottom
                    If .rcDummy.Right > .lBmpW Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_ThreeD_From_Bottom Then
                    InflateRect .rcTemp, 0, -MyProp.cFramePerMoveY
                    .rcTemp.Top = .lBmpH - .rcTemp.Bottom
                    
                    InflateRect .rcDummy, 0, MyProp.cFramePerMoveY
                    SetRect .rcDummy, 0, 0, .rcDummy.Right, .rcDummy.Bottom
                    If .rcDummy.Bottom > .lBmpH Then
                        CopyRect .rcDummy, .rcRect
                    End If
                Else
                    InflateRect .rcTemp, 0, -MyProp.cFramePerMoveY
                    SetRect .rcTemp, 0, 0, .rcTemp.Right, .rcTemp.Bottom
                    
                    InflateRect .rcDummy, 0, MyProp.cFramePerMoveY
                    .rcDummy.Top = .lBmpH - .rcDummy.Bottom
                    If .rcDummy.Bottom > .lBmpH Then
                        CopyRect .rcDummy, .rcRect
                    End If
                End If
            End If
        End If
                
        If EqualRect(.rcRect, .rcDummy) Then
            ' do nothing
        Else
            StretchImage UserControl.hdc, .rcTemp.Left, .rcTemp.Top, .rcTemp.Right, .rcTemp.Bottom, _
                        .lhSvBgDC, 0, 0, .lBmpW, .lBmpH, vbSrcCopy
        
            StretchImage UserControl.hdc, .rcDummy.Left, .rcDummy.Top, .rcDummy.Right, .rcDummy.Bottom, _
                        .lhSrcDC, 0, 0, .lBmpW, .lBmpH, vbSrcCopy
        
            RaiseEvent AniPosition(ScaleX(CSng(.rcDummy.Left), vbPixels, vbTwips), _
                                    ScaleY(CSng(.rcDummy.Top), vbPixels, vbTwips), _
                                    ScaleX(CSng(.rcDummy.Right), vbPixels, vbTwips), _
                                    ScaleY(CSng(.rcDummy.Bottom), vbPixels, vbTwips))
        End If
    End With
End Function

Private Function WipeAni() As Boolean
    Static bInit As Boolean
    
    With CSH
        If Not bInit Then
            If .lAniEffect = crd_Wipe_Left Then
                SetRect .rcDummy, .lBmpW, 0, 0, .lBmpH
            ElseIf .lAniEffect = crd_Wipe_Up Then
                SetRect .rcDummy, 0, .lBmpH, .lBmpW, 0
            ElseIf .lAniEffect = crd_Wipe_Right Then
                SetRect .rcDummy, 0, 0, 0, .lBmpH
            Else
                SetRect .rcDummy, 0, 0, .lBmpW, 0
            End If
            bInit = True
            Exit Function
        Else
            If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                bInit = False: WipeAni = True
                If MyProp.cAutoFlipCard Then
                    Face = FlipCard(MyProp.cFace)
                End If
                Exit Function
            Else
                If .lAniEffect = crd_Wipe_Left Then
                    InflateRect .rcDummy, MyProp.cFramePerMoveX, 0
                    .rcDummy.Left = .lBmpW - .rcDummy.Right
                    If .rcDummy.Right > .lBmpW Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Wipe_Up Then
                    InflateRect .rcDummy, 0, MyProp.cFramePerMoveY
                    .rcDummy.Top = .lBmpH - .rcDummy.Bottom
                    If .rcDummy.Bottom > .lBmpH Then
                        CopyRect .rcDummy, .rcRect
                    End If
                ElseIf .lAniEffect = crd_Wipe_Right Then
                    InflateRect .rcDummy, MyProp.cFramePerMoveX, 0
                    SetRect .rcDummy, 0, 0, .rcDummy.Right, .rcDummy.Bottom
                    If .rcDummy.Right > .lBmpW Then
                        CopyRect .rcDummy, .rcRect
                    End If
                Else
                    InflateRect .rcDummy, 0, MyProp.cFramePerMoveY
                    SetRect .rcDummy, 0, 0, .rcDummy.Right, .rcDummy.Bottom
                    If .rcDummy.Bottom > .lBmpH Then
                        CopyRect .rcDummy, .rcRect
                    End If
                End If
            End If
        End If
        
        If EqualRect(.rcRect, .rcDummy) Then
            ' do nothing
        Else
            BitBlt UserControl.hdc, .rcDummy.Left, .rcDummy.Top, .rcDummy.Right, .rcDummy.Bottom, _
                  .lhSrcDC, .rcDummy.Left, .rcDummy.Top, vbSrcCopy
        
            RaiseEvent AniPosition(ScaleX(CSng(.rcDummy.Left), vbPixels, vbTwips), _
                                    ScaleY(CSng(.rcDummy.Top), vbPixels, vbTwips), _
                                    ScaleX(CSng(.rcDummy.Right), vbPixels, vbTwips), _
                                    ScaleY(CSng(.rcDummy.Bottom), vbPixels, vbTwips))
        End If
    End With
End Function

Private Function ZoomAni() As Boolean
    Static bInit As Boolean
    Static step  As Integer
                
    With CSH
        If Not bInit Then
            If .lAniEffect = crd_Zoom_In Then
                SetRect .rcDummy, .lWinSzX / 2, .lWinSzX / 2, 0, 0
                step = 0
            ElseIf .lAniEffect = crd_Zoom_Out Then
                SetRect .rcDummy, (.lWinSzX - .lBmpW * 2) / 2, (.lWinSzY - .lBmpW * 2), _
                                  .lBmpW * 2, .lBmpH * 2
                step = IIf(.lBmpW < .lBmpH, .lBmpH * 2, .lBmpW * 2)
            End If
            bInit = True
            Exit Function
        Else
            If EqualRect(.rcRect, .rcDummy) Or StopAni Then
                bInit = False: ZoomAni = True
                If MyProp.cAutoFlipCard Then
                    Face = FlipCard(MyProp.cFace)
                End If
                Exit Function
            Else
                Dim ll As Long
                Dim ls As Long
                Dim sx As Single
                Dim sy As Single
                
                ll = Largest(.lBmpW, .lBmpH)
                ls = Smallest(.lBmpW, .lBmpH)
        
                If .lBmpW <= .lBmpH Then
                    sx = 1
                    sy = CSng(ll / ls)
                Else
                    sx = CSng(ll / ls)
                    sy = 1
                End If
                
                .rcDummy.Right = sx * step
                .rcDummy.Bottom = sy * step
                .rcDummy.Left = (.lWinSzX - .rcDummy.Right) / 2
                .rcDummy.Top = (.lWinSzY - .rcDummy.Bottom) / 2
                    
                If .lAniEffect = crd_Zoom_In Then
                    If (.rcDummy.Right > .lBmpW) Or (.rcDummy.Bottom > .lBmpH) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                    step = step + MyProp.cFramePerMoveX
                ElseIf .lAniEffect = crd_Zoom_Out Then
                    If (.rcDummy.Right < .lBmpW) Or (.rcDummy.Bottom < .lBmpH) Then
                        CopyRect .rcDummy, .rcRect
                    End If
                    step = step - MyProp.cFramePerMoveX
                End If
            End If
        End If
        
        If EqualRect(.rcRect, .rcDummy) Then
            ' do nothing
        Else
            StretchImage UserControl.hdc, .rcDummy.Left, .rcDummy.Top, .rcDummy.Right, .rcDummy.Bottom, _
                        .lhSrcDC, 0, 0, .lBmpW, .lBmpH, vbSrcCopy
        
            RaiseEvent AniPosition(ScaleX(CSng(.rcDummy.Left), vbPixels, vbTwips), _
                                    ScaleY(CSng(.rcDummy.Top), vbPixels, vbTwips), _
                                    ScaleX(CSng(.rcDummy.Right), vbPixels, vbTwips), _
                                    ScaleY(CSng(.rcDummy.Bottom), vbPixels, vbTwips))
        End If
    End With
End Function

Private Function SlideShow() As Boolean
    Select Case MyProp.cEffect
    Case Is = crd_Effect_Elevator
        SlideShow = ElevatorAni
    Case Is = crd_Effect_Flip
        SlideShow = FlipAni
    Case Is = crd_Effect_FlyIn
        SlideShow = FlyInAni
    Case Is = crd_Effect_FlyOut
        SlideShow = FlyOutAni
    Case Is = crd_Effect_Gate
        SlideShow = GateAni
    Case Is = crd_Effect_Split
        SlideShow = SplitAni
    Case Is = crd_Effect_Stretch
        SlideShow = StretchAni
    Case Is = crd_Effect_ThreeD
        SlideShow = ThreeDAni
    Case Is = crd_Effect_Wipe
        SlideShow = WipeAni
    Case Is = crd_Effect_Zoom
        SlideShow = ZoomAni
    End Select
    
    RefreshWindow UserControl.hWnd
End Function

Private Function GetAniEffect() As Long
    Select Case MyProp.cEffect
    Case Is = crd_Effect_Elevator
        GetAniEffect = MyProp.cElevator
    Case Is = crd_Effect_Flip
        GetAniEffect = MyProp.cFlip
    Case Is = crd_Effect_FlyIn
        GetAniEffect = MyProp.cFlyIn
    Case Is = crd_Effect_FlyOut
        GetAniEffect = MyProp.cFlyOut
    Case Is = crd_Effect_Gate
        GetAniEffect = MyProp.cGate
    Case Is = crd_Effect_Split
        GetAniEffect = MyProp.cSplit
    Case Is = crd_Effect_Stretch
        GetAniEffect = MyProp.cStretch
    Case Is = crd_Effect_ThreeD
        GetAniEffect = MyProp.cThreeD
    Case Is = crd_Effect_Wipe
        GetAniEffect = MyProp.cWipe
    Case Is = crd_Effect_Zoom
        GetAniEffect = MyProp.cZoom
    End Select
End Function

Private Function FlipCard(FaceOp As FaceConstants) As FaceConstants
    FlipCard = IIf(FaceOp = crd_Up, crd_Down, crd_Up)
End Function

Public Sub DrawFocusRect()
    If GetFocus = UserControl.hWnd Then
        Dim rcRect As RECT
        
        GetClientRect UserControl.hWnd, rcRect
        DrawLine UserControl.hdc, 0, 0, CInt(rcRect.Right), 0, 1, &HFF0000
        DrawLine UserControl.hdc, 0, 0, 0, CInt(rcRect.Bottom), 1, &HFF0000
        DrawLine UserControl.hdc, CInt(rcRect.Right) - 1, 1, CInt(rcRect.Right) - 1, _
                 CInt(rcRect.Bottom), 1, &HFF7979
        DrawLine UserControl.hdc, 1, CInt(rcRect.Bottom) - 1, CInt(rcRect.Right), _
                 CInt(rcRect.Bottom) - 1, 1, &HFF7979
        DrawLine UserControl.hdc, 1, 1, CInt(rcRect.Right) - 2, 1, 1, &HFFA8A8
        DrawLine UserControl.hdc, 1, 1, 1, CInt(rcRect.Bottom) - 2, 1, &HFFA8A8
        DrawLine UserControl.hdc, CInt(rcRect.Right) - 2, 1, CInt(rcRect.Right) - 2, _
                 CInt(rcRect.Bottom) - 2, 1, &HFFD5D5
        DrawLine UserControl.hdc, 1, CInt(rcRect.Bottom) - 2, CInt(rcRect.Right) - 1, _
                 CInt(rcRect.Bottom) - 2, 1, &HFFD5D5
        RefreshWindow UserControl.hWnd
    Else
        UserControl.Cls
    End If
End Sub

Public Sub DrawTrackRect()
    Dim rcRect As RECT
    
    GetClientRect UserControl.hWnd, rcRect
    DrawLine UserControl.hdc, 0, 0, CInt(rcRect.Right), 0, 1, &H4080FF
    DrawLine UserControl.hdc, 0, 0, 0, CInt(rcRect.Bottom), 1, &H4080FF
    DrawLine UserControl.hdc, CInt(rcRect.Right) - 1, 1, CInt(rcRect.Right) - 1, _
             CInt(rcRect.Bottom), 1, &H6C9CFF
    DrawLine UserControl.hdc, 1, CInt(rcRect.Bottom) - 1, CInt(rcRect.Right), _
             CInt(rcRect.Bottom) - 1, 1, &H6C9CFF
    DrawLine UserControl.hdc, 1, 1, CInt(rcRect.Right) - 2, 1, 1, &H64B1FF
    DrawLine UserControl.hdc, 1, 1, 1, CInt(rcRect.Bottom) - 2, 1, &H64B1FF
    DrawLine UserControl.hdc, CInt(rcRect.Right) - 2, 1, CInt(rcRect.Right) - 2, _
             CInt(rcRect.Bottom) - 2, 1, &HC4E1FF
    DrawLine UserControl.hdc, 1, CInt(rcRect.Bottom) - 2, CInt(rcRect.Right) - 1, _
             CInt(rcRect.Bottom) - 2, 1, &HC4E1FF
    RefreshWindow UserControl.hWnd
End Sub

Private Sub SetHotTracking()
    If GetProp(UserControl.hWnd, "TimerActive") Then
        ' do nothing
    Else
        SetProp UserControl.hWnd, "ClassID", ObjPtr(Me)
        SetProp UserControl.hWnd, "TimerActive", True
        SetTimer UserControl.hWnd, TimerHotTrackingID, 50, AddressOf TimerCallBack
    End If
End Sub





