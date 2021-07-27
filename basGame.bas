Attribute VB_Name = "basGame"
Option Explicit

Public Const CardWidth = 71
Public Const CardHeight = 96
Public Const TotalCards = 52

' use to identify the file
Public Const GameSignature = "65827383-76857383657148"

Private Type CardInfo
    curDeck         As Integer
    curMask         As Integer
    Deck            As Integer
    DeckBackground  As Integer
    DeckMaskStyle   As Integer
    DeckMaskPicture As String
    DeckPicture     As String
    Effect          As Integer
    FontBold        As Boolean
    FontItalic      As Boolean
    FontSize        As Integer
    FontName        As String
    FontTransparent As Boolean
    Forecolor       As Long
    FramePerMoveX   As Integer
    FramePerMoveY   As Integer
    Speed           As Integer
    Text            As String
    TypEffect       As Integer
End Type

Private Type GameSettingsInfo
    BkFile     As String  ' set background bitmap
    BkColor    As Long    ' set background color
    BkMode     As Integer ' 0-Bitmap, 1-Color
    Clip       As Boolean
    DistX      As Integer
    DistY      As Integer
    Speed      As Integer
    Trail      As Boolean
    VicAniMode As Integer ' Custom, Random
    VicAniSel  As Integer ' Bounce, Bounce (Stretch)...
    WaveExpr   As String
    Card       As CardInfo
End Type

Public Enum CardSizeConstants
    cs_Small
    cs_Standard
    cs_Large
End Enum

Public GSI As GameSettingsInfo

Public VBScript As New MSScriptControl.ScriptControl

Public Sub ResizeCard(thisForm As Form, ByVal MatchTag As String, _
                      Optional ByVal CardSizeOp As CardSizeConstants = cs_Standard)
                      
    Dim lcW         As Long
    Dim lcH         As Long
    Dim PlayingCard As Object
    
    If CardSizeOp = cs_Small Then
        lcW = CardWidth * 0.75
        lcH = CardHeight * 0.75
    ElseIf CardSizeOp = cs_Standard Then
        lcW = CardWidth
        lcH = CardHeight
    Else
        lcW = CardWidth * 1.25
        lcH = CardHeight * 1.25
    End If
    
    lcW = lcW * Screen.TwipsPerPixelX
    lcH = lcH * Screen.TwipsPerPixelY
    
    For Each PlayingCard In thisForm.Controls
        If MatchTag = PlayingCard.Tag Then
            If TypeName(PlayingCard) = "Card" Then
                PlayingCard.Move PlayingCard.Left, PlayingCard.Top, lcW, lcH
            End If
        End If
    Next PlayingCard
End Sub

Public Function GetTypeEffect(thisCard As Object)
     Select Case thisCard.Effect
     Case 0  ' None
     Case 1  ' Elevator
          GetTypeEffect = thisCard.Elevator
     Case 2  ' Flip
          GetTypeEffect = thisCard.Flip
     Case 3  ' FlyIn
          GetTypeEffect = thisCard.FlyIn
     Case 4  ' FlyOut
          GetTypeEffect = thisCard.FlyOut
     Case 5  ' Gate
          GetTypeEffect = thisCard.Gate
     Case 6  ' Split
          GetTypeEffect = thisCard.Split
     Case 7  ' Stretch
          GetTypeEffect = thisCard.Stretch
     Case 8  ' ThreeD
          GetTypeEffect = thisCard.ThreeD
     Case 9  ' Wipe
          GetTypeEffect = thisCard.Wipe
     Case 10 ' Zoom
          GetTypeEffect = thisCard.Zoom
     End Select
End Function

Public Sub SetTypeEffect(thisCard As Object, TypeEffect As Integer)
     Select Case thisCard.Effect
     Case 1  ' Elevator
          thisCard.Elevator = TypeEffect
     Case 2  ' Flip
          thisCard.Flip = TypeEffect
     Case 3  ' FlyIn
          thisCard.FlyIn = TypeEffect
     Case 4  ' FlyOut
          thisCard.FlyOut = TypeEffect
     Case 5  ' Gate
          thisCard.Gate = TypeEffect
     Case 6  ' Split
          thisCard.Split = TypeEffect
     Case 7  ' Stretch
          thisCard.Stretch = TypeEffect
     Case 8  ' ThreeD
          thisCard.ThreeD = TypeEffect
     Case 9  ' Wipe
          thisCard.Wipe = TypeEffect
     Case 10 ' Zoom
          thisCard.Zoom = TypeEffect
     End Select
End Sub
