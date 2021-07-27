VERSION 5.00
Object = "*\AprjCard.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCardOp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Card Options"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   2220
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrAni 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   540
      Top             =   4500
   End
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   60
      Top             =   4500
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   2940
      TabIndex        =   12
      Top             =   4140
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3960
      TabIndex        =   11
      Top             =   4140
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4980
      TabIndex        =   10
      Top             =   4140
      Width           =   975
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Height          =   315
      Left            =   60
      TabIndex        =   9
      Top             =   4140
      Width           =   975
   End
   Begin MSComctlLib.ImageList imlDeck 
      Left            =   1020
      Top             =   4500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlMask 
      Left            =   1620
      Top             =   4500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCardOp.frx":0000
            Key             =   "None"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCardOp.frx":015A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraAni 
      Caption         =   "Select Card Animation"
      Height          =   3615
      Left            =   180
      TabIndex        =   14
      Top             =   360
      Width           =   5655
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Frame fraTray 
         Caption         =   "Entry animation"
         Height          =   855
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   2520
         Width           =   5295
         Begin VB.ComboBox cmbType 
            Height          =   315
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   420
            Width           =   2355
         End
         Begin VB.ComboBox cmbEffect 
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   420
            Width           =   2475
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Effect"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.PictureBox picHolder 
         BackColor       =   &H00808080&
         FontTransparent =   0   'False
         Height          =   1875
         Index           =   1
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   1515
         TabIndex        =   15
         Top             =   300
         Width           =   1575
         Begin prjCard.Card crdAni 
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            DeckMaskPicture =   "frmCardOp.frx":06F4
            DeckPicture     =   "frmCardOp.frx":0710
            Elevator        =   0
            Flip            =   0
            FlyIn           =   0
            FlyOut          =   0
            HotTracking     =   -1  'True
            ShowFocusRect   =   -1  'True
            Stretch         =   0
            Suit            =   1
            ThreeD          =   0
            Wipe            =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Picture         =   "frmCardOp.frx":072C
         End
      End
      Begin VB.Frame fraTray 
         Caption         =   "Miscellaneous"
         Height          =   2175
         Index           =   1
         Left            =   1740
         TabIndex        =   20
         Top             =   300
         Width           =   3735
         Begin MSComctlLib.Slider sldSpeed 
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1680
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            Max             =   50
            SelStart        =   50
            Value           =   50
         End
         Begin VB.ListBox lstStepY 
            Height          =   255
            ItemData        =   "frmCardOp.frx":0AF2
            Left            =   900
            List            =   "frmCardOp.frx":0B14
            TabIndex        =   23
            Top             =   780
            Width           =   1875
         End
         Begin VB.ListBox lstStepX 
            Height          =   255
            ItemData        =   "frmCardOp.frx":0B37
            Left            =   900
            List            =   "frmCardOp.frx":0B59
            TabIndex        =   22
            Top             =   420
            Width           =   1875
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Step (Y) :"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   27
            Top             =   780
            Width           =   660
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Step (X) :"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   26
            Top             =   420
            Width           =   660
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speed"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   25
            Top             =   1440
            Width           =   465
         End
      End
   End
   Begin VB.Frame fraDeckBk 
      Height          =   3615
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   5655
      Begin VB.CommandButton cmdMask 
         Caption         =   "Add File"
         Height          =   315
         Index           =   0
         Left            =   3780
         TabIndex        =   39
         Top             =   2340
         Width           =   915
      End
      Begin VB.CommandButton cmdMask 
         Caption         =   "Remove"
         Height          =   315
         Index           =   1
         Left            =   4695
         TabIndex        =   38
         Top             =   2340
         Width           =   915
      End
      Begin VB.CommandButton cmdDeck 
         Caption         =   "Remove"
         Height          =   315
         Index           =   1
         Left            =   2835
         TabIndex        =   37
         Top             =   2340
         Width           =   915
      End
      Begin VB.CheckBox chkTrans 
         Caption         =   "Opaque"
         Height          =   255
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3240
         Width           =   1035
      End
      Begin VB.CommandButton cmdTextOp 
         Caption         =   "Color"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   32
         Top             =   3240
         Width           =   1035
      End
      Begin VB.CommandButton cmdTextOp 
         Caption         =   "Font"
         Height          =   255
         Index           =   0
         Left            =   2100
         TabIndex        =   31
         Top             =   3240
         Width           =   1035
      End
      Begin VB.TextBox txtMask 
         Height          =   285
         Left            =   2040
         TabIndex        =   30
         Top             =   2940
         Width           =   3435
      End
      Begin VB.CommandButton cmdDeck 
         Caption         =   "Add File"
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   13
         Top             =   2340
         Width           =   915
      End
      Begin VB.OptionButton optBk 
         Caption         =   "&Tile"
         Height          =   315
         Index           =   1
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2820
         Width           =   1275
      End
      Begin VB.OptionButton optBk 
         Caption         =   "&Stretch"
         Height          =   315
         Index           =   0
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2460
         Width           =   1275
      End
      Begin VB.PictureBox picHolder 
         BackColor       =   &H00808080&
         Height          =   1875
         Index           =   0
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   1515
         TabIndex        =   4
         Top             =   300
         Width           =   1575
         Begin prjCard.Card crdSelector 
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            AutoFlipCard    =   0   'False
            DeckMaskPicture =   "frmCardOp.frx":0B7C
            DeckPicture     =   "frmCardOp.frx":0B98
            Effect          =   1
            Elevator        =   0
            Face            =   1
            Flip            =   0
            FlyIn           =   0
            FlyOut          =   0
            HotTracking     =   -1  'True
            ShowFocusRect   =   -1  'True
            Stretch         =   0
            Suit            =   1
            ThreeD          =   0
            Wipe            =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Picture         =   "frmCardOp.frx":0BB4
         End
      End
      Begin VB.VScrollBar vsCard 
         Height          =   1755
         Left            =   1680
         TabIndex        =   3
         Top             =   300
         Width           =   195
      End
      Begin VB.HScrollBar hsCard 
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvwDeck 
         Height          =   2055
         Left            =   1920
         TabIndex        =   5
         Top             =   300
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Back Deck"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwMask 
         Height          =   2055
         Left            =   3765
         TabIndex        =   6
         Top             =   300
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Back Mask"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblText 
         Caption         =   "Text:"
         Height          =   195
         Left            =   2040
         TabIndex        =   34
         Top             =   2760
         Width           =   1215
      End
   End
   Begin MSComctlLib.TabStrip tbsCardOp 
      Height          =   4095
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7223
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Animation"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Deck"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2400
      TabIndex        =   36
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2400
      TabIndex        =   35
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmCardOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PathDeck As String
Dim PathMask As String
Dim DeckColl As New Collection
Dim MaskColl As New Collection

Dim CE As New cAddControlEffect

Private Sub chkTrans_Click()
    If chkTrans.Value = vbChecked Then
        crdSelector.FontTransparent = False
    ElseIf chkTrans.Value = vbUnchecked Then
        crdSelector.FontTransparent = True
    End If
    cmdApply.Enabled = True
End Sub

Private Sub cmbEffect_Click()
    cmbType.Clear
    
    Select Case cmbEffect.ListIndex
    Case Is = 0 ' none
        cmbType.AddItem "[None]"
    Case Is = 1  ' Elevator
        cmbType.AddItem "Down"
        cmbType.AddItem "Left"
        cmbType.AddItem "Right"
        cmbType.AddItem "Up"
    Case Is = 2  ' Flip
        cmbType.AddItem "Bottom"
        cmbType.AddItem "Horizontal"
        cmbType.AddItem "Left"
        cmbType.AddItem "Right"
        cmbType.AddItem "Top"
        cmbType.AddItem "Vertical"
    Case Is = 3  ' FlyIn
        cmbType.AddItem "Bottom"
        cmbType.AddItem "Bottom Left"
        cmbType.AddItem "Bottom Right"
        cmbType.AddItem "Left"
        cmbType.AddItem "Right"
        cmbType.AddItem "Top"
        cmbType.AddItem "Top Left"
        cmbType.AddItem "Top Right"
    Case Is = 4  ' FlyOut
        cmbType.AddItem "Bottom"
        cmbType.AddItem "Bottom Left"
        cmbType.AddItem "Bottom Right"
        cmbType.AddItem "Left"
        cmbType.AddItem "Right"
        cmbType.AddItem "Top"
        cmbType.AddItem "Top Left"
        cmbType.AddItem "Top Right"
    Case Is = 5  ' Gate
        cmbType.AddItem "Horizontal In"
        cmbType.AddItem "Horinzontal Out"
        cmbType.AddItem "Vertical In"
        cmbType.AddItem "Vertical Out"
    Case Is = 6  ' Split
        cmbType.AddItem "Horizontal In"
        cmbType.AddItem "Horinzontal Out"
        cmbType.AddItem "Vertical In"
        cmbType.AddItem "Vertical Out"
    Case Is = 7  ' Stretch
        cmbType.AddItem "Across"
        cmbType.AddItem "From Bottom"
        cmbType.AddItem "From Left"
        cmbType.AddItem "From Right"
        cmbType.AddItem "From Top"
    Case Is = 8  ' ThreeD
        cmbType.AddItem "From Bottom"
        cmbType.AddItem "From Left"
        cmbType.AddItem "From Right"
        cmbType.AddItem "From Top"
    Case Is = 9  ' Wipe
        cmbType.AddItem "Down"
        cmbType.AddItem "Left"
        cmbType.AddItem "Right"
        cmbType.AddItem "Up"
    Case Else    ' Zoom
        cmbType.AddItem "In"
        cmbType.AddItem "Out"
    End Select
    
    cmdPlay.Enabled = IIf(cmbEffect.ListIndex, True, False)
    crdAni.Effect = cmbEffect.ListIndex
    
    If cmbType.ListCount <> 0 Then
        cmbType.ListIndex = cmbType.TopIndex
    End If
    cmdApply.Enabled = True
End Sub

Private Sub cmbType_Click()
    lstStepX.Enabled = False
    lstStepY.Enabled = False
    
    Select Case cmbEffect.ListIndex
    Case Is = 0 ' none
    Case Is = 1 ' Elevator
        Select Case cmbType.ListIndex
        Case Is = crd_Elevator_Down, crd_Elevator_Up
            lstStepY.Enabled = True
        Case Else
            lstStepX.Enabled = True
        End Select
        crdAni.Elevator = cmbType.ListIndex
    Case Is = 2 ' Flip
        Select Case cmbType.ListIndex
        Case Is = crd_Flip_Bottom, crd_Flip_Vertical, crd_Flip_Top
            lstStepY.Enabled = True
        Case Else
            lstStepX.Enabled = True
        End Select
        crdAni.Flip = cmbType.ListIndex
    Case Is = 3 ' FlyIn
        Select Case cmbType.ListIndex
        Case Is = crd_FlyIn_Bottom, crd_FlyIn_Top
            lstStepY.Enabled = True
        Case Else
            lstStepX.Enabled = True
        End Select
        crdAni.FlyIn = cmbType.ListIndex
    Case Is = 4 ' FlyOut
        Select Case cmbType.ListIndex
        Case Is = crd_FlyOut_Bottom, crd_FlyOut_Top
            lstStepY.Enabled = True
        Case Else
            lstStepX.Enabled = True
        End Select
        crdAni.FlyOut = cmbType.ListIndex
    Case Is = 5 ' Gate
        Select Case cmbType.ListIndex
        Case Is = crd_Split_Vertical_In, crd_Split_Vertical_Out
            lstStepX.Enabled = True
        Case Else
            lstStepY.Enabled = True
        End Select
        crdAni.Gate = cmbType.ListIndex
    Case Is = 6 ' Split
        Select Case cmbType.ListIndex
        Case Is = crd_Split_Vertical_In, crd_Split_Vertical_Out
            lstStepX.Enabled = True
        Case Else
            lstStepY.Enabled = True
        End Select
        crdAni.Split = cmbType.ListIndex
    Case Is = 7 ' Stretch
        Select Case cmbType.ListIndex
        Case Is = crd_Stretch_From_Bottom, crd_Stretch_From_Top
            lstStepY.Enabled = True
        Case Else
            lstStepX.Enabled = True
        End Select
        crdAni.Stretch = cmbType.ListIndex
    Case Is = 8 ' ThreeD
        Select Case cmbType.ListIndex
        Case Is = crd_ThreeD_From_Bottom, crd_ThreeD_From_Top
            lstStepY.Enabled = True
        Case Else
            lstStepX.Enabled = True
        End Select
        crdAni.ThreeD = cmbType.ListIndex
    Case Is = 9 ' Wipe
        Select Case cmbType.ListIndex
        Case Is = crd_Wipe_Down, crd_Wipe_Up
            lstStepY.Enabled = True
        Case Else
            lstStepX.Enabled = True
        End Select
        crdAni.Wipe = cmbType.ListIndex
    Case Else   ' Zoom
        lstStepX.Enabled = True
        crdAni.Zoom = cmbType.ListIndex
    End Select
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
    Dim oCard As Object
    
    With frmMain
        ' set the main card only 'crdPlayingCard(0)'
        .crdPlayingCard(0).Update = False ' turn off update
        .crdPlayingCard(0).Deck = crdSelector.Deck
        .crdPlayingCard(0).DeckBackground = crdSelector.DeckBackground
        .crdPlayingCard(0).DeckMaskStyle = crdSelector.DeckMaskStyle
        Set .crdPlayingCard(0).DeckMaskPicture = crdSelector.DeckMaskPicture
        Set .crdPlayingCard(0).DeckPicture = crdSelector.DeckPicture
        .crdPlayingCard(0).Effect = crdAni.Effect
        .crdPlayingCard(0).FramePerMoveX = crdAni.FramePerMoveX
        .crdPlayingCard(0).FramePerMoveY = crdAni.FramePerMoveY
        Set .crdPlayingCard(0).Font = crdSelector.Font
        .crdPlayingCard(0).Forecolor = crdSelector.Forecolor
        .crdPlayingCard(0).FontTransparent = crdSelector.FontTransparent
        .crdPlayingCard(0).Speed = crdAni.Speed
        .crdPlayingCard(0).Text = crdSelector.Text
        .crdPlayingCard(0).Update = True  ' turn on update
        .crdPlayingCard(0).Refresh        ' update the card
        
        SetTypeEffect .crdPlayingCard(0), GetTypeEffect(crdAni)
    End With
    
    For Each oCard In frmMain
        With frmMain
            If oCard.Tag = "PlayingCard" Then
                If (oCard.Index = 0) And (oCard.Name = "crdPlayingCard") Then
                    ' do nothing
                Else
                    oCard.Update = False
                    oCard.Deck = .crdPlayingCard(0).Deck
                    oCard.DeckBackground = .crdPlayingCard(0).DeckBackground
                    oCard.DeckMaskStyle = .crdPlayingCard(0).DeckMaskStyle
                    Set oCard.DeckMaskPicture = .crdPlayingCard(0).DeckMaskPicture
                    oCard.DeckPicture = .crdPlayingCard(0).DeckPicture
                    oCard.Effect = .crdPlayingCard(0).Effect
                    oCard.FramePerMoveX = .crdPlayingCard(0).FramePerMoveX
                    oCard.FramePerMoveY = .crdPlayingCard(0).FramePerMoveY
                    Set oCard.Font = .crdPlayingCard(0).Font
                    oCard.Forecolor = .crdPlayingCard(0).Forecolor
                    oCard.FontTransparent = .crdPlayingCard(0).FontTransparent
                    oCard.Speed = .crdPlayingCard(0).Speed
                    oCard.Text = .crdPlayingCard(0).Text
                    oCard.Update = True
                    
                    If oCard.Face = crd_Down Then
                        Set oCard.Picture = .crdPlayingCard(0).Picture
                    End If

                    SetTypeEffect oCard, GetTypeEffect(.crdPlayingCard(0))
                End If
            End If
        End With
    Next oCard
    
    GSI.Card.curDeck = lvwDeck.SelectedItem.Index
    GSI.Card.curMask = lvwMask.SelectedItem.Index
    GSI.Card.DeckPicture = PathDeck
    GSI.Card.DeckMaskPicture = PathMask
    cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeck_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    Dim i   As Integer
    Dim key As String
    
    Select Case Index
    Case Is = 0 ' Add file
        With dlgDialog
            .Filter = "Bitmap Files (*.bmp) | *.bmp; | " _
                    & "JPEG (*.JPG,*.JPEG) | *.jpg; *.jpeg; | " _
                    & "GIF (*.GIF) | *.GIF; | " _
                    & "All Picture Files | *.bmp; *.gif; *.jpg; *.jpeg; | " _
                    & "All Files (*.*) | *.*"
            .FilterIndex = 1
            .InitDir = ""
            .Filename = ""
            .ShowOpen
            
            If .Filename <> "" Then
                key = "Cus" & imlDeck.ListImages.Count
                imlDeck.ListImages.Add , key, LoadPicture(.Filename)
                lvwDeck.ListItems.Add , key, Left$(.FileTitle, InStr(.FileTitle, ".") - 1), , imlDeck.ListImages.Count
                DeckColl.Add .Filename
            End If
        End With
    Case Is = 1 ' Remove file
        Dim nSum As Integer
        
        nSum = 0
        For i = 1 To lvwDeck.ListItems.Count
            If Left$(lvwDeck.ListItems(i).key, 3) <> "Cus" Then
                nSum = nSum + 1
            End If
        Next i
        
        DeckColl.Remove lvwDeck.SelectedItem.Index - nSum
        lvwDeck.ListItems.Remove lvwDeck.SelectedItem.Index
        hsCard.Value = 0
        hsCard_Change
    End Select
    
    hsCard.Max = lvwDeck.ListItems.Count - 1
    
    Dim InFile As Long
                
    InFile = FreeFile
    Open App.Path & "\Data\deck.dat" For Output As InFile
        For i = 1 To DeckColl.Count
            If Dir$(DeckColl(i)) <> "" Then
                Write #InFile, DeckColl(i)
            End If
        Next i
    Close InFile
    cmdApply.Enabled = True
    Exit Sub
    
ErrHandler:
    If InFile <> 0 Then Close InFile
    MsgBox Err.Description, vbOKOnly Or vbInformation, "Error"
End Sub

Private Sub cmdDefault_Click()
    chkTrans.Value = vbChecked
    cmbEffect.ListIndex = 0
    cmbType.ListIndex = 0
    lstStepX.ListIndex = 0
    lstStepY.ListIndex = 0
    optBk(0).Value = vbChecked
    txtMask.Text = ""
    txtMask.Forecolor = vbBlack
    sldSpeed.Value = sldSpeed.Max
    
    crdSelector.FontName = "Arial Black"
    crdSelector.Forecolor = vbBlack
    crdSelector.FontBold = False
    crdSelector.FontItalic = False
    crdSelector.FontSize = 8
    crdSelector.FontTransparent = False
    
    lvwMask_ItemClick lvwMask.ListItems(1)
    lvwDeck_ItemClick lvwDeck.ListItems(1)
    
    Call sldSpeed_Change
End Sub

Private Sub cmdMask_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    Dim i   As Integer
    Dim key As String
    
    Select Case Index
    Case Is = 0 ' Add file
        With dlgDialog
            .Filter = "Bitmap Files (*.bmp) | *.bmp; | " _
                    & "JPEG (*.JPG,*.JPEG) | *.jpg; *.jpeg; | " _
                    & "GIF (*.GIF) | *.GIF; | " _
                    & "All Picture Files | *.bmp; *.gif; *.jpg; *.jpeg; | " _
                    & "All Files (*.*) | *.*"
            .FilterIndex = 1
            .InitDir = ""
            .Filename = ""
            .ShowOpen
            
            If .Filename <> "" Then
                key = "Cus" & imlMask.ListImages.Count
                imlMask.ListImages.Add , key, LoadPicture(.Filename)
                lvwMask.ListItems.Add , key, _
                                      Left$(.FileTitle, InStr(.FileTitle, ".") - 1), , imlMask.ListImages.Count
                MaskColl.Add .Filename
            End If
        End With
    Case Is = 1 ' Remove file
        Dim nSum As Integer
        
        nSum = 0
        For i = 1 To lvwMask.ListItems.Count
            If Left$(lvwMask.ListItems(i).key, 3) <> "Cus" Then
                nSum = nSum + 1
            End If
        Next i
        
        MaskColl.Remove lvwMask.SelectedItem.Index - nSum
        lvwMask.ListItems.Remove lvwMask.SelectedItem.Index
        lvwMask.ListItems(1).Selected = True
        vsCard.Value = 0
        Call vsCard_Change
    End Select
    
    vsCard.Max = lvwMask.ListItems.Count - 1
    
    Dim InFile As Long
                
    InFile = FreeFile
    Open App.Path & "\Data\mask.dat" For Output As InFile
        For i = 1 To MaskColl.Count
            If Dir$(MaskColl(i)) <> "" Then
                Write #InFile, MaskColl(i)
            End If
        Next i
    Close InFile
    cmdApply.Enabled = True
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbOKOnly Or vbInformation, "Error"
End Sub

Private Sub cmdOk_Click()
    If cmdApply.Enabled Then cmdApply_Click
    Unload Me
End Sub

Private Sub cmdTextOp_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    With dlgDialog
        Select Case Index
        Case Is = 0 ' Font
            .FontName = crdSelector.FontName
            .FontBold = crdSelector.FontBold
            .FontItalic = crdSelector.FontItalic
            .FontSize = crdSelector.FontSize
            .Flags = 1
            .ShowFont
            
            crdSelector.FontName = .FontName
            crdSelector.FontBold = .FontBold
            crdSelector.FontItalic = .FontItalic
            crdSelector.FontSize = .FontSize
        Case Is = 1 ' Color
            .ShowColor
            crdSelector.Forecolor = .Color
            txtMask.Forecolor = .Color
        End Select
    End With
    cmdApply.Enabled = True
    Exit Sub

ErrHandler:
End Sub

Private Sub cmdPlay_Click()
    If cmbEffect.ListIndex <> 0 Then
        If cmdPlay.Caption = "Play" Then
            cmdPlay.Caption = "Stop"
            cmbEffect.Enabled = False
            cmbType.Enabled = False
            crdAni.PlayAni
        Else
            cmdPlay.Caption = "Play"
            crdAni.StopAni = True
            cmbEffect.Enabled = True
            cmbType.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    
    Call InitAniOp
    Call InitDeckOp

    With frmMain
        crdAni.Effect = .crdPlayingCard(0).Effect
        crdAni.FramePerMoveX = .crdPlayingCard(0).FramePerMoveX
        crdAni.FramePerMoveY = .crdPlayingCard(0).FramePerMoveY
        crdAni.Speed = .crdPlayingCard(0).Speed
        SetTypeEffect crdAni, GetTypeEffect(.crdPlayingCard(0))
        
        Set crdSelector.Font = .crdPlayingCard(0).Font
        crdSelector.Forecolor = .crdPlayingCard(0).Forecolor
        crdSelector.FontTransparent = .crdPlayingCard(0).FontTransparent
        crdSelector.Text = .crdPlayingCard(0).Text
        
        optBk(.crdPlayingCard(0).DeckBackground).Value = vbChecked
        chkTrans.Value = IIf(crdSelector.FontTransparent, vbUnchecked, vbChecked)
        cmbEffect.ListIndex = crdAni.Effect
        cmbType.ListIndex = GetTypeEffect(.crdPlayingCard(0))
        lstStepX.Text = crdAni.FramePerMoveX
        lstStepY.Text = crdAni.FramePerMoveY
        sldSpeed.Value = sldSpeed.Max - crdAni.Speed + 1
        txtMask.Text = .crdPlayingCard(0).Text
    End With
    
    If GSI.Card.curDeck <= 0 Then
        GSI.Card.curDeck = 1
    ElseIf GSI.Card.curDeck > lvwDeck.ListItems.Count Then
        GSI.Card.curDeck = 1
    End If
    
    If GSI.Card.curMask <= 0 Then
        GSI.Card.curMask = 1
    ElseIf GSI.Card.curMask > lvwMask.ListItems.Count Then
        GSI.Card.curMask = 1
    End If
    
    lvwDeck.ListItems(GSI.Card.curDeck).Selected = True
    lvwMask.ListItems(GSI.Card.curMask).Selected = True
    lvwDeck_ItemClick lvwDeck.ListItems(lvwDeck.SelectedItem.Index)
    lvwMask_ItemClick lvwMask.ListItems(lvwMask.SelectedItem.Index)
    
    Call RegsCtrlEffect
    cmdApply.Enabled = False
    tmrEffect.Enabled = True
End Sub

Private Sub hsCard_Change()
    If crdSelector.StopAni Then crdSelector.StopAni = True
    
    Dim key       As String
    Dim sType     As String
    Static OldVal As Integer
    
    If hsCard.Value > OldVal Then
        crdSelector.Elevator = crd_Elevator_Left
    Else
        crdSelector.Elevator = crd_Elevator_Right
    End If
    
    key = lvwDeck.ListItems(hsCard.Value + 1).key
    sType = Left$(key, 3)
    
    PathDeck = ""
    cmdDeck(1).Enabled = False
    crdSelector.Update = False
    If sType = "App" Then
        crdSelector.Deck = hsCard.Value
    ElseIf sType = "Cus" Then
        cmdDeck(1).Enabled = True
        crdSelector.Deck = Crd_Deck_Customize
        Set crdSelector.DeckPicture = imlDeck.ListImages(key).Picture
        Debug.Print lvwDeck.SelectedItem.Index - 2
        PathDeck = DeckColl(lvwDeck.SelectedItem.Index - 2)
    End If
    
    crdSelector.PlayAni
    crdSelector.Update = True
    crdSelector.Refresh
    lvwDeck.ListItems(hsCard.Value + 1).Selected = True
    lvwDeck.SelectedItem.EnsureVisible
    OldVal = hsCard.Value
    cmdApply.Enabled = True
End Sub

Private Sub lstStepX_Click()
    crdAni.FramePerMoveX = Val(lstStepX.List(lstStepX.ListIndex))
    cmdApply.Enabled = True
End Sub

Private Sub lstStepY_Click()
    crdAni.FramePerMoveY = Val(lstStepY.List(lstStepY.ListIndex))
    cmdApply.Enabled = True
End Sub

Private Sub lvwDeck_ItemClick(ByVal Item As MSComctlLib.ListItem)
    hsCard.Value = Item.Index - 1
    Call hsCard_Change
End Sub

Private Sub lvwMask_ItemClick(ByVal Item As MSComctlLib.ListItem)
    vsCard.Value = Item.Index - 1
    Call vsCard_Change
End Sub

Private Sub optBk_Click(Index As Integer)
    crdSelector.DeckBackground = Index
    cmdApply.Enabled = True
End Sub

Private Sub sldSpeed_Change()
    crdAni.Speed = sldSpeed.Max - sldSpeed.Value + 1
    cmdApply.Enabled = True
End Sub

Private Sub tbsCardOp_Click()
    Select Case tbsCardOp.SelectedItem.Index
    Case Is = 1 ' Animation
        crdAni.Update = False
        crdAni.Deck = crdSelector.Deck
        Set crdAni.DeckPicture = crdSelector.DeckPicture
        crdAni.DeckMaskStyle = crdSelector.DeckMaskStyle
        Set crdAni.DeckMaskPicture = crdSelector.DeckMaskPicture
        Set crdAni.Font = crdSelector.Font
        crdAni.FontTransparent = crdSelector.FontTransparent
        crdAni.Forecolor = crdSelector.Forecolor
        crdAni.Text = crdSelector.Text
        crdAni.Update = True
        crdAni.Refresh
        
        fraAni.ZOrder 0
    Case Is = 2 ' Deck
        fraDeckBk.ZOrder 0
    End Select
End Sub

Private Sub tmrAni_Timer()
    If crdAni.StopAni Then
        If cmdPlay.Caption <> "Play" Then
            cmdPlay.Caption = "Play"
            cmbEffect.Enabled = True
            cmbType.Enabled = True
        End If
    End If
End Sub

Private Sub tmrEffect_Timer()
    Call CE.StartEffect
End Sub

Private Sub txtMask_Change()
    If crdSelector.DeckMaskStyle = crd_Text Then
        crdSelector.Text = txtMask.Text
        cmdApply.Enabled = True
    End If
End Sub

Private Sub vsCard_Change()
    If crdSelector.StopAni Then crdSelector.StopAni = True
    
    Dim key       As String
    Dim sType     As String
    Static OldVal As Integer
    
    If vsCard.Value > OldVal Then
        crdSelector.Elevator = crd_Elevator_Up
    Else
        crdSelector.Elevator = crd_Elevator_Down
    End If
    
    key = lvwMask.ListItems(vsCard.Value + 1).key
    sType = Left$(key, 3)
    
    PathMask = ""
    cmdMask(1).Enabled = False
    crdSelector.Update = False
    If sType = "Non" Then
        crdSelector.DeckMaskStyle = crd_None
    ElseIf sType = "Txt" Then
        crdSelector.DeckMaskStyle = crd_Text
        crdSelector.Text = txtMask.Text
    ElseIf sType = "App" Then
        crdSelector.DeckMaskStyle = vsCard.Value - 1
    Else
        cmdMask(1).Enabled = True
        crdSelector.DeckMaskStyle = crd_Customize
        Set crdSelector.DeckMaskPicture = imlMask.ListImages(key).Picture
        PathMask = MaskColl(lvwMask.SelectedItem.Index - 9)
    End If
    
    crdSelector.PlayAni
    crdSelector.Update = True
    crdSelector.Refresh
    lvwMask.ListItems(vsCard.Value + 1).Selected = True
    lvwMask.SelectedItem.EnsureVisible
    OldVal = vsCard.Value
    cmdApply.Enabled = True
End Sub

Private Sub InitDeckOp()
    On Error GoTo ErrHandler
    
    Dim i      As Integer
    Dim lMidX  As Integer
    Dim lMidY  As Integer
    Dim cwTwip As Long
    Dim chTwip As Long
    
    cwTwip = ScaleX(CardWidth, vbPixels, vbTwips)
    chTwip = ScaleY(CardHeight, vbPixels, vbTwips)
    lMidX = (picHolder(0).ScaleWidth - cwTwip) / 2
    lMidY = (picHolder(0).ScaleHeight - chTwip) / 2
    crdSelector.Move lMidX, lMidY, cwTwip, chTwip
        
    lvwDeck.ColumnHeaders(1).Width = lvwDeck.Width - _
                                     ScaleX(5, vbPixels, vbTwips)
    lvwMask.ColumnHeaders(1).Width = lvwMask.Width - _
                                     ScaleX(5, vbPixels, vbTwips)
    
    For i = 201 To 202
        imlDeck.ListImages.Add , , LoadResPicture(i, vbResBitmap)
        If i = 201 Then Set lvwDeck.SmallIcons = imlDeck
        lvwDeck.ListItems.Add , "App" & i, LoadResString(i), , i Mod 200
    Next i
    
    Set lvwMask.SmallIcons = imlMask
    lvwMask.ListItems.Add , "Non", "[None]", , 1
    lvwMask.ListItems.Add , "Txt", "[Text]", , 2
    
    For i = 301 To 307
        imlMask.ListImages.Add , , LoadResPicture(i, vbResBitmap)
        lvwMask.ListItems.Add , "App" & i, LoadResString(i), , (i Mod 300) + 2
    Next i
    
    Dim Filename As String, Buffer As String
    Dim InFile As Long, key As String, Title As String
        
    Filename = App.Path & "\Data\deck.dat"
    If Dir$(Filename) <> "" Then
        InFile = FreeFile
        Open Filename For Input As InFile
            Do While Not EOF(InFile)
                Input #InFile, Buffer
                If Dir$(Buffer) <> "" Then
                    key = "Cus" & imlDeck.ListImages.Count
                    imlDeck.ListImages.Add , key, LoadPicture(Buffer)
                    Title = Right$(Buffer, Len(Buffer) - InStr(Buffer, "\"))
                    Title = Left$(Title, InStr(Title, ".") - 1)
                    lvwDeck.ListItems.Add , key, Title, , imlDeck.ListImages.Count
                    DeckColl.Add Buffer
                End If
            Loop
        Close InFile
    End If
    
    Filename = App.Path & "\Data\mask.dat"
    If Dir$(Filename) <> "" Then
        InFile = FreeFile
        Open Filename For Input As InFile
            Do While Not EOF(InFile)
                Input #InFile, Buffer
                If Dir$(Buffer) <> "" Then
                    key = "Cus" & imlMask.ListImages.Count
                    imlMask.ListImages.Add , key, LoadPicture(Buffer)
                    Title = Right$(Buffer, Len(Buffer) - InStr(Buffer, "\"))
                    Title = Left$(Title, InStr(Title, ".") - 1)
                    lvwMask.ListItems.Add , key, Title, , imlMask.ListImages.Count
                    MaskColl.Add Buffer
                End If
            Loop
        Close InFile
    End If
    
    hsCard.Max = lvwDeck.ListItems.Count - 1
    vsCard.Max = lvwMask.ListItems.Count - 1
    tmrEffect.Enabled = True
    Exit Sub
    
ErrHandler:
    If InFile <> 0 Then Close InFile
    MsgBox Err.Description, vbOKOnly Or vbInformation, "Error"
End Sub

Private Sub InitAniOp()
    Dim i      As Integer
    Dim lMidX  As Integer
    Dim lMidY  As Integer
    Dim cwTwip As Long
    Dim chTwip As Long
    
    cwTwip = ScaleX(CardWidth, vbPixels, vbTwips)
    chTwip = ScaleY(CardHeight, vbPixels, vbTwips)
    lMidX = (picHolder(1).ScaleWidth - cwTwip) / 2
    lMidY = (picHolder(1).ScaleHeight - chTwip) / 2
    crdAni.Move lMidX, lMidY, cwTwip, chTwip
    
    For i = 401 To 411
        cmbEffect.AddItem LoadResString(i)
    Next i
    
    cmbEffect.ListIndex = cmbEffect.TopIndex
    tmrAni.Enabled = True
End Sub

Private Sub RegsCtrlEffect()
    With CE
        .RegisterControl cmdApply, vbBlue, vbCyan, vbRed
        .RegisterControl cmdCancel, vbBlue, vbCyan, vbRed
        .RegisterControl cmdDeck(0), vbBlue, vbCyan, vbYellow
        .RegisterControl cmdDeck(1), vbBlue, vbCyan, vbYellow
        .RegisterControl cmdDefault, vbBlue, vbCyan, vbRed
        .RegisterControl cmdMask(0), vbBlue, vbCyan, vbMagenta
        .RegisterControl cmdMask(1), vbBlue, vbCyan, vbMagenta
        .RegisterControl cmdOk, vbBlue, vbCyan, vbRed
        .RegisterControl cmdPlay, vbBlue, vbCyan, vbWhite
        .RegisterControl cmdTextOp(0), vbBlue, vbCyan, vbGreen
        .RegisterControl cmdTextOp(1), vbBlue, vbCyan, vbGreen
        .RegisterControl cmbEffect, vbBlue, vbCyan, -1
        .RegisterControl cmbType, vbBlue, vbCyan, -1
        .RegisterControl chkTrans, vbBlue, vbCyan, vbGreen
        .RegisterControl lstStepX, vbBlue, vbCyan, -1
        .RegisterControl lstStepY, vbBlue, vbCyan, -1
        .RegisterControl optBk(0), vbBlue, vbCyan, vbWhite
        .RegisterControl optBk(1), vbBlue, vbCyan, vbWhite
    End With
End Sub
