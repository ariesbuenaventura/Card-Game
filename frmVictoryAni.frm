VERSION 5.00
Object = "*\AprjCard.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVictoryAni 
   Caption         =   "Customize Victory Animation"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   960
      Top             =   0
   End
   Begin VB.PictureBox picTray 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   9255
      TabIndex        =   7
      Top             =   5805
      Width           =   9255
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Left            =   8220
         TabIndex        =   15
         Top             =   60
         Width           =   975
      End
      Begin VB.Frame fraTray 
         Height          =   1695
         Index           =   0
         Left            =   60
         TabIndex        =   16
         Top             =   -60
         Width           =   6675
         Begin VB.Frame fraSettings 
            Caption         =   "Settings"
            Height          =   1395
            Left            =   2940
            TabIndex        =   27
            Top             =   180
            Width           =   2535
            Begin VB.CheckBox chkTrail 
               Caption         =   "Trail"
               Enabled         =   0   'False
               Height          =   195
               Left            =   1800
               TabIndex        =   37
               Top             =   1140
               Value           =   2  'Grayed
               Width           =   615
            End
            Begin VB.ListBox lstDistY 
               Height          =   255
               ItemData        =   "frmVictoryAni.frx":0000
               Left            =   1800
               List            =   "frmVictoryAni.frx":0022
               TabIndex        =   34
               Top             =   840
               Width           =   615
            End
            Begin VB.ListBox lstDistX 
               Height          =   255
               ItemData        =   "frmVictoryAni.frx":0045
               Left            =   600
               List            =   "frmVictoryAni.frx":0067
               TabIndex        =   33
               Top             =   840
               Width           =   615
            End
            Begin VB.ComboBox cmbAni 
               Height          =   315
               ItemData        =   "frmVictoryAni.frx":008A
               Left            =   120
               List            =   "frmVictoryAni.frx":00A0
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   420
               Width           =   2295
            End
            Begin MSComctlLib.Slider sldSpeed 
               Height          =   195
               Left            =   600
               TabIndex        =   30
               Top             =   1140
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   344
               _Version        =   393216
            End
            Begin VB.Label lblDist 
               AutoSize        =   -1  'True
               Caption         =   "Dist. Y:"
               Height          =   195
               Index           =   1
               Left            =   1260
               TabIndex        =   36
               Top             =   855
               Width           =   510
            End
            Begin VB.Label lblDist 
               AutoSize        =   -1  'True
               Caption         =   "Dist. X:"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   35
               Top             =   855
               Width           =   510
            End
            Begin VB.Label lblSpeed 
               AutoSize        =   -1  'True
               Caption         =   "Speed:"
               Height          =   195
               Left            =   60
               TabIndex        =   31
               Top             =   1140
               Width           =   510
            End
            Begin VB.Label lblAni 
               AutoSize        =   -1  'True
               Caption         =   "Animation:"
               Height          =   195
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame fraExpr 
            Caption         =   "Customize wave animation"
            Enabled         =   0   'False
            Height          =   1395
            Left            =   120
            TabIndex        =   19
            Top             =   180
            Width           =   2775
            Begin VB.CheckBox chkClip 
               Caption         =   "Clip"
               Enabled         =   0   'False
               Height          =   195
               Left            =   2100
               TabIndex        =   26
               Top             =   1080
               Value           =   2  'Grayed
               Width           =   555
            End
            Begin VB.CommandButton cmdExpr 
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   315
               Left            =   2340
               TabIndex        =   23
               Top             =   405
               Width           =   315
            End
            Begin VB.CommandButton cmdGraph 
               Caption         =   "&Graph"
               Enabled         =   0   'False
               Height          =   315
               Left            =   60
               TabIndex        =   22
               Top             =   1020
               Width           =   795
            End
            Begin VB.CommandButton cmdClear 
               Caption         =   "&Clear"
               Enabled         =   0   'False
               Height          =   315
               Left            =   900
               TabIndex        =   21
               Top             =   1020
               Width           =   795
            End
            Begin VB.TextBox txtExpr 
               Enabled         =   0   'False
               Height          =   285
               Left            =   60
               TabIndex        =   20
               Top             =   420
               Width           =   2235
            End
            Begin VB.Label lblNote 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "note: x is constant"
               Enabled         =   0   'False
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   960
               TabIndex        =   25
               Top             =   720
               Width           =   1290
            End
            Begin VB.Label lblExpr 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Expression:"
               Enabled         =   0   'False
               Height          =   195
               Left            =   60
               TabIndex        =   24
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "&Stop"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5640
            TabIndex        =   18
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdPlay 
            Caption         =   "&Play"
            Height          =   375
            Left            =   5640
            TabIndex        =   17
            Top             =   300
            Width           =   975
         End
      End
   End
   Begin VB.Timer tmrVictory 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrTitle 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1500
      ScaleHeight     =   855
      ScaleWidth      =   5835
      TabIndex        =   6
      Top             =   60
      Width           =   5835
      Begin VB.PictureBox picSpritePlane 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   0
         Picture         =   "frmVictoryAni.frx":00E8
         ScaleHeight     =   285
         ScaleWidth      =   465
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picMaskPlane 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   480
         Picture         =   "frmVictoryAni.frx":084A
         ScaleHeight     =   285
         ScaleWidth      =   465
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picSpritePlane 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   0
         Picture         =   "frmVictoryAni.frx":0FAC
         ScaleHeight     =   285
         ScaleWidth      =   465
         TabIndex        =   11
         Top             =   300
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picMaskPlane 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   480
         Picture         =   "frmVictoryAni.frx":170E
         ScaleHeight     =   285
         ScaleWidth      =   465
         TabIndex        =   10
         Top             =   300
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   480
         ScaleHeight     =   285
         ScaleWidth      =   465
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   465
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.PictureBox picBk 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   9255
      TabIndex        =   1
      Top             =   -60
      Width           =   9255
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "email:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   0
         Left            =   1260
         TabIndex        =   5
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ravemasterharuglory@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   0
         Left            =   1740
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   780
         Width           =   2385
      End
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ajb2001lg@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   4320
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   780
         Width           =   1665
      End
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "or"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   1
         Left            =   4140
         TabIndex        =   2
         Top             =   780
         Width           =   165
      End
   End
   Begin VB.PictureBox picViewer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      Height          =   4755
      Left            =   0
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   617
      TabIndex        =   0
      Top             =   1020
      Width           =   9315
      Begin prjCard.Card crdAni 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   450
         DeckMaskPicture =   "frmVictoryAni.frx":1E70
         DeckPicture     =   "frmVictoryAni.frx":1E8C
         Elevator        =   0
         Flip            =   0
         FlyIn           =   0
         FlyOut          =   0
         Stretch         =   0
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
         Picture         =   "frmVictoryAni.frx":1EA8
      End
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4755
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   613
      TabIndex        =   32
      Top             =   1020
      Width           =   9255
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuMenuSelect 
         Caption         =   "Select"
      End
      Begin VB.Menu mnuMenuRemove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "frmVictoryAni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PLANE = 6

Private Type PlaneInfo
    cx       As Integer
    cy       As Integer
    Width    As Integer
    Height   As Integer
    Speed    As Integer
End Type

Dim bSetCard     As Boolean
Dim curAni       As Integer
Dim curColorIdx  As Integer
Dim RightEdge    As Integer
Dim BottomEdge   As Integer
Dim Plane(MAX_PLANE - 1) As PlaneInfo

Dim VicAni    As New cVictoryAni
Dim ExprColl  As New Collection
Dim ColorColl As New Collection

Dim CE As New cAddControlEffect
Dim Trigo As New cTrigonometry
Dim ScriptCalc As New MSScriptControl.ScriptControl
    
Private Sub chkTrail_Click()
    Call DrawViewerDsgn
End Sub

Private Sub cmbAni_Click()
    If curAni = cmbAni.ListIndex Then Exit Sub
    Dim OldTimer As Boolean
    
    OldTimer = tmrVictory.Enabled
    tmrVictory.Enabled = False
    Call EnabledWaveControl(False)
    
    chkTrail.Enabled = False
    lstDistX.Enabled = True
    lstDistY.Enabled = True
    Select Case cmbAni.ListIndex
    Case Is = 0 ' Bounce
    Case Is = 1 ' Bounce (Scatter)
        chkTrail.Enabled = True
        lstDistX.Enabled = False
        lstDistY.Enabled = False
    Case Is = 2 ' Bounce (Trail)
        chkTrail.Enabled = True
    Case Is = 3 ' Spin
        lstDistY.Enabled = False
    Case Is = 4 ' Spin (Trail)
        chkTrail.Enabled = True
        lstDistY.Enabled = False
    Case Is = 5 ' Wave
        lstDistY.Enabled = False
        Call EnabledWaveControl(True)
    End Select
    
    Call DrawViewerDsgn
    curAni = cmbAni.ListIndex
    VicAni.Reset = True
    bSetCard = True
    tmrVictory.Enabled = OldTimer
End Sub

Private Sub EnabledWaveControl(bVal As Boolean)
    cmdGraph.Enabled = bVal
    cmdClear.Enabled = bVal
    txtExpr.Enabled = bVal
    cmdExpr.Enabled = bVal
    chkClip.Enabled = bVal
    lblNote.Enabled = bVal
    lblExpr.Enabled = bVal
    fraExpr.Enabled = bVal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    picGraph.Cls
    picGraph.ToolTipText = ""
    
    Set ExprColl = Nothing
    Set ColorColl = Nothing
End Sub

Private Sub cmdExpr_Click()
    Dim OldTimerVictory As Boolean
    
    OldTimerVictory = tmrVictory.Enabled
    
    tmrTitle.Enabled = False
    tmrVictory.Enabled = False
    frmExpr.Show vbModal, Me
    
    If frmExpr.RetVal <> "" Then
        txtExpr.Text = frmExpr.RetVal
    End If
    tmrTitle.Enabled = True
    tmrVictory.Enabled = OldTimerVictory
    cmdPlay.Enabled = Not tmrVictory.Enabled
    cmdStop.Enabled = tmrVictory.Enabled
End Sub

Private Sub cmdGraph_Click()
    Dim s1 As String
    Dim s2 As String
    Dim i  As Integer
    
    For i = 1 To ExprColl.Count
        s1 = RemovSpace(LCase(Trim$(txtExpr.Text)))
        s2 = RemovSpace(LCase(ExprColl(i)))
        ' if the graph is already exist then exit
        If s1 = s2 Then Exit Sub
    Next i
            
    Dim IsFound  As Boolean
    Dim Color    As Long
    Dim Counter As Integer
    
    Counter = 0
    Do While True
        Color = RGB(Int(125 * Rnd) + 50, _
                    Int(125 * Rnd) + 50, _
                    Int(125 * Rnd) + 50)
        
        IsFound = False
        For i = 1 To ColorColl.Count
            If ColorColl(i) = Color Then
                IsFound = True
            ElseIf ColorColl(i) = vbWhite Then
                IsFound = True ' background color
            ElseIf ColorColl(i) = vbBlack Then
                IsFound = True ' color coordinate sys
            ElseIf ColorColl(i) = &HD8E9EC Then
                IsFound = True ' color coordinate sys
            End If
        Next i
        
        If (Not IsFound) Or (Counter = 100) Then
            Exit Do
        End If
        
        Counter = Counter + 1 ' once we reached 100 we need
                              ' to end the loop otherwise it will
                              ' loop forever
    Loop
    
    If Counter <= 100 Then
        If PlotGraph(Trim$(txtExpr.Text), Color) Then
            ExprColl.Add Trim$(txtExpr.Text)
            ColorColl.Add Color
        End If
    
        tmrVictory.Enabled = False
        picGraph.ZOrder 0
    End If
End Sub

Private Sub cmdPlay_Click()
    cmdPlay.Enabled = False
    cmdStop.Enabled = True
    tmrVictory.Enabled = True
    picViewer.ZOrder 0
End Sub

Private Sub cmdStop_Click()
    cmdStop.Enabled = False
    cmdPlay.Enabled = True
    tmrVictory.Enabled = False
End Sub

Private Sub Form_Load()
    ScriptCalc.Language = "VBScript"
    ScriptCalc.Timeout = NoTimeout
    ScriptCalc.AddObject "Trigo", Trigo, True
    
    Call ShowTitle
    Call InitCard
    Call DrawCoordinateSys
    
    Set lblEmail(0).MouseIcon = LoadResPicture(101, vbResCursor)
    Set lblEmail(1).MouseIcon = LoadResPicture(101, vbResCursor)
    
    bSetCard = True
    cmbAni.ListIndex = cmbAni.TopIndex
    curAni = cmbAni.ListIndex

    lstDistX.Selected(GSI.DistX - 1) = True
    lstDistY.Selected(GSI.DistY - 1) = True
    VicAni.DistX = GSI.DistX
    VicAni.DistY = GSI.DistY
    Call RegsCtrlEffect
End Sub

Private Sub lblEmail_Click(Index As Integer)
    On Error Resume Next
    
    Call ShellExecute(0, "open", "mailto:" & lblEmail(Index).Caption, 0, 0, 0)
End Sub

Private Sub lblEmail_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set lblEmail(Index).MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub lblEmail_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set lblEmail(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub InitCard()
    Dim i As Integer, Temp As Collection
    
    crdAni(0).Width = CardWidth * 0.75
    crdAni(0).Height = CardHeight * 0.75
    CreateObject crdAni, 20
    
    Set Temp = Shuffle(crdAni.Count)
    For i = 1 To crdAni.Count - 1
        crdAni(i).Update = False
        crdAni(i).Rank = Temp(i) Mod 13
        crdAni(i).Suit = Temp(i) Mod 4
        crdAni(i).Tag = "PlayingCard"
        crdAni(i).Update = True
        crdAni(i).Refresh
        crdAni(i).Visible = True
    Next i
End Sub

Private Sub ShowTitle()
    Dim lhFont    As Long
    Dim lhrgn     As Long
    Dim lhOldFont As Long
    Dim rcRect    As RECT
    Dim xPos      As Integer
    Dim yPos      As Integer
    Dim sz        As Size
    Dim Title     As String
    
    Const CLIP_LH_ANGLES = 16
    Const PROOF_QUALITY As Long = 2
        
    lhFont = CreateFont(72, 17, 0, 0, 700, _
                        False, False, False, _
                        0, 0, CLIP_LH_ANGLES, PROOF_QUALITY, 0, "Arial Black")

    lhOldFont = SelectObject(picTitle.hdc, lhFont)
    
    Title = "Victory Animation"
    GetClientRect picTitle.hwnd, rcRect
    GetTextExtentPoint32 picTitle.hdc, Title, Len(Title), sz
    
    xPos = (rcRect.Right - rcRect.Left - sz.cx) / 2
    yPos = (rcRect.Bottom - rcRect.Top - sz.cy) / 2
    
    BeginPath picTitle.hdc
    TextOut picTitle.hdc, xPos, yPos, Title, Len(Title)
    EndPath picTitle.hdc
    lhrgn = PathToRegion(picTitle.hdc)
    SetWindowRgn picTitle.hwnd, lhrgn, True
    
    SelectObject picTitle.hdc, lhOldFont
    DeleteObject lhFont
    
    Dim i As Integer, OldScaleMode
    
    OldScaleMode = picTitle.ScaleMode
    picTitle.ScaleMode = vbUser
    picTitle.ScaleWidth = 510
    picTitle.ScaleHeight = 510
    
    For i = 0 To 255
        picTitle.Line (0, i)-(picTitle.ScaleWidth, i), RGB(0, i, 0)
    Next i
    For i = 256 To 510
        picTitle.Line (0, i)-(picTitle.ScaleWidth, i), RGB(510 - i, 510 - i, 0)
    Next i
    
    Set picTitle.Picture = picTitle.Image
    
    ' add glitters
    For i = 0 To 2000
        SetPixelV picTitle.hdc, CInt(Rnd * rcRect.Right), (CInt(Rnd * rcRect.Bottom)), _
                                RGB(CInt(Rnd * 125) + 125, CInt(Rnd * 125) + 125, CInt(Rnd * 125) + 125)
    Next i
    
    picTitle.ScaleMode = OldScaleMode
    Set picTitle.Picture = picTitle.Image
    tmrTitle.Enabled = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    picViewer.Width = Me.Width - Screen.TwipsPerPixelX * 7
    picViewer.Height = Me.ScaleHeight - picTray.ScaleHeight - _
                       picViewer.Top + Screen.TwipsPerPixelX * 3
    picGraph.Width = Me.Width - Screen.TwipsPerPixelX * 7
    picGraph.Height = Me.ScaleHeight - picTray.ScaleHeight - _
                       picGraph.Top + Screen.TwipsPerPixelX * 3
    picTitle.Move (Me.Width - picTitle.Width) / 2, -200
    picBk.Move 0, 0, Me.Width, picBk.Height
    
    Dim xmid As Integer
    
    xmid = (picBk.ScaleWidth - (lblPrompt(0).Width + lblPrompt(1).Width + _
            lblEmail(0).Width + lblEmail(1).Width)) / 2
    lblPrompt(0).Left = xmid
    lblEmail(0).Left = xmid + lblPrompt(0).Width
    lblPrompt(1).Left = xmid + lblPrompt(0).Width + lblEmail(0).Width
    lblEmail(1).Left = xmid + lblPrompt(0).Width + lblEmail(0).Width + _
                       lblPrompt(1).Width
                     
    Call DrawBkDsgn
    Call DrawViewerDsgn
    Call DrawCoordinateSys
    
    Dim i As Integer
    
    For i = 1 To ExprColl.Count
        Call PlotGraph(ExprColl(i), ColorColl(i))
    Next i
End Sub

Private Sub DrawBkDsgn()
    Dim OldScaleMode As Integer
    Dim i As Integer, rcRect As RECT
    
    OldScaleMode = picBk.ScaleMode
    picBk.ScaleMode = vbUser
    picBk.ScaleWidth = 255
    picBk.ScaleHeight = 255
    
    picBk.Cls
    For i = 0 To 255
        picBk.Line (0, i)-(picBk.ScaleWidth, i), _
                   RGB(0, 0, 255 - i)
    Next i
    picBk.ScaleMode = OldScaleMode
End Sub

Private Sub DrawViewerDsgn()
    Dim OldScaleMode As Integer
    Dim i As Integer, rcRect As RECT
    
    OldScaleMode = picViewer.ScaleMode
    picViewer.ScaleMode = vbUser
    picViewer.ScaleWidth = 125
    picViewer.ScaleHeight = 125
    
    GetClientRect picViewer.hwnd, rcRect
    
    picViewer.Cls
    For i = 0 To 125
        picViewer.Line (0, i)-(picViewer.ScaleWidth, i), _
                       RGB(125 - i, 125 - i, 125 - i)
    Next i
    picViewer.ScaleMode = OldScaleMode
End Sub

Private Sub PlayPlaneAni()
    Dim i    As Integer
    Dim Temp As Integer

    Static bInit As Boolean
    
    If Not bInit Then
        Dim BMP()  As BITMAP
        Dim rcRect As RECT
        Dim Speed  As Collection
        Dim DivH  As Integer
        
        ReDim BMP(UBound(Plane())) As BITMAP
    
        CreateObject picSprite, UBound(Plane())
        CreateObject picMask, UBound(Plane())
        
        GetClientRect picTitle.hwnd, rcRect
        RightEdge = rcRect.Right - rcRect.Left
        BottomEdge = rcRect.Bottom - rcRect.Top
        DivH = BottomEdge / 2 / UBound(Plane())
        
        Set Speed = Shuffle(UBound(Plane()) + 1)
        For i = LBound(Plane()) To UBound(Plane())
            GetObjectAPI picSpritePlane(0).Picture.handle, _
                         Len(BMP(i)), BMP(i)
        
            ' offset cx
            Temp = IIf(CInt(Rnd * -1), 1, -1)
            If Temp = 1 Then
                Plane(i).cx = -BMP(i).bmWidth - _
                              CInt(Rnd * 5) * CInt(Rnd * 10)
            Else
                Plane(i).cx = RightEdge - BMP(i).bmWidth - _
                              CInt(Rnd * 5) * CInt(Rnd * 10)
            End If
                   
            ' offset cy
            Plane(i).cy = DivH * i + 10
            
            Plane(i).Speed = (Speed(i + 1) Mod 4) + 4
            Plane(i).Width = BMP(i).bmWidth
            Plane(i).Height = BMP(i).bmHeight
            
            Call SetPlaneDir(i, 0)
        Next i
                
        bInit = True
    Else
        Dim NewX  As Integer
        
        picTitle.Cls
        For i = LBound(Plane()) To UBound(Plane())
            Temp = Plane(i).cx + Plane(i).Speed
            
            If Temp + Plane(i).Width > RightEdge Then
                Plane(i).Speed = -Abs(Plane(i).Speed)
                Call SetPlaneDir(i, 1)
            ElseIf Temp < 0 Then
                Plane(i).Speed = Abs(Plane(i).Speed)
                Call SetPlaneDir(i, 0)
            End If
                        
            NewX = Plane(i).cx + Plane(i).Speed
            
            BitBlt picTitle.hdc, NewX, Plane(i).cy, Plane(i).Width, Plane(i).Height, _
                   picMask(i).hdc, 0, 0, vbSrcAnd
            BitBlt picTitle.hdc, NewX, Plane(i).cy, Plane(i).Width, Plane(i).Height, _
                   picSprite(i).hdc, 0, 0, vbSrcInvert
            
            Plane(i).cx = NewX
        Next i
        
        RefreshWindow picTitle.hwnd
    End If
End Sub

Private Sub DrawCoordinateSys()
    Dim i As Integer
    Dim W As Integer
    Dim h As Integer
    Dim xmid    As Integer
    Dim ymid    As Integer
    Dim curStep As Integer
    Dim Crest   As Single
    Dim rcRect  As RECT
    Dim sz      As Size
    
    GetClientRect picGraph.hwnd, rcRect
    W = rcRect.Right - rcRect.Left
    h = rcRect.Bottom - rcRect.Top
    xmid = W / 2
    ymid = h / 2
    curStep = 10
    Crest = ymid / curStep

    picGraph.Cls
    Set picGraph = Nothing
    GetTextExtentPoint32 picGraph.hdc, "T", 1, sz
    
    For i = -curStep To curStep
        DrawLine picGraph.hdc, 0, ymid - i * Crest, _
                               W, ymid - i * Crest, 1, RGB(236, 233, 216)
        TextOut picGraph.hdc, xmid + IIf(Sgn(i) = 1, sz.cx, sz.cx / 2 + 1), _
                              ymid - i * Crest - sz.cy / 2, _
                              CStr(i), Len(CStr(i))
    Next i
    
    DrawLine picGraph.hdc, 0, ymid, W, ymid, 1, RGB(0, 0, 0)
    DrawLine picGraph.hdc, xmid, 0, xmid, h, 1, RGB(0, 0, 0)
    Set picGraph.Picture = picGraph.Image
End Sub

Private Function PlotGraph(Expr As String, Color As Long) As Boolean
    If Trim$(Expr) = "" Then Exit Function
    
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Crest  As Single
    Dim rcRect As RECT
    Dim xmid   As Integer
    Dim ymid   As Integer
    Dim bVal   As Boolean
    
    On Error GoTo EvalError
        
    GetClientRect picGraph.hwnd, rcRect
    xmid = (rcRect.Right - rcRect.Left) / 2
    ymid = (rcRect.Bottom - rcRect.Top) / 2
    Crest = ymid / 10
    
    For i = -xmid To xmid
        Call ScriptCalc.ExecuteStatement("x=" & Rads(CSng(i)))
        X = i: Y = Crest * ScriptCalc.Eval(Expr)
        SetPixelV picGraph.hdc, xmid + X, ymid - Y, Color
        
        If Not bVal Then
            DrawCircle picGraph.hdc, xmid + X + 6, ymid - Y, 6, Color, Color, 1
            bVal = True
        End If
    Next i
    
    RefreshWindow picGraph.hwnd
    PlotGraph = True
    Exit Function
    
EvalError:
    If Err.Number = 1002 Then
        MsgBox "Syntax Error", vbInformation Or vbOKOnly, "Error"
    Else
        Resume Next
    End If
End Function

Private Sub SetPlaneDir(Index As Integer, Direction As Integer)
    Set picSprite(Index).Picture = picSpritePlane(Direction).Picture
    Set picMask(Index).Picture = picMaskPlane(Direction).Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrTitle.Enabled = False
    tmrVictory.Enabled = False
    
    DestroyObject picSprite
    DestroyObject picMask
    
    Set frmVictoryAni = Nothing
End Sub

Private Sub lstDistX_Click()
    VicAni.DistX = Val(lstDistX.List(lstDistX.ListIndex))
End Sub

Private Sub lstDistY_Click()
    VicAni.DistY = Val(lstDistY.List(lstDistY.ListIndex))
End Sub

Private Sub mnuMenuRemove_Click()
    If curColorIdx <> -1 Then
        If ExprColl.Count > 0 Then
            ExprColl.Remove curColorIdx
            ColorColl.Remove curColorIdx
            Call DrawCoordinateSys
            
            Dim i As Integer
            For i = 1 To ExprColl.Count
                Call PlotGraph(ExprColl(i), ColorColl(i))
            Next i
        End If
    End If
End Sub

Private Sub mnuMenuSelect_Click()
    If curColorIdx <> -1 Then
        txtExpr.Text = ExprColl(curColorIdx)
    End If
End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picGraph.MouseIcon.handle <> 0 Then
        Set picGraph.MouseIcon = LoadResPicture(102, vbResCursor)
    End If
End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i    As Integer
    Dim bVal As Boolean
    
    For i = 1 To ColorColl.Count
        If ColorColl(i) = GetPixel(picGraph.hdc, X, Y) Then
            picGraph.ToolTipText = ExprColl(i)
            bVal = True
            Exit For
        End If
    Next i
    
    If bVal Then
        If picGraph.MouseIcon.handle = 0 Then
            Set picGraph.MouseIcon = LoadResPicture(101, vbResCursor)
        End If
    Else
        If picGraph.MouseIcon.handle <> 0 Then
            Set picGraph.MouseIcon = Nothing
        End If
    End If
End Sub

Private Sub picGraph_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    curColorIdx = -1
    For i = 1 To ColorColl.Count
        If ColorColl(i) = GetPixel(picGraph.hdc, X, Y) Then
            curColorIdx = i
            Me.PopupMenu mnuMenu
            Exit Sub
        End If
    Next i
    
    If picGraph.MouseIcon.handle <> 0 Then
        Set picGraph.MouseIcon = LoadResPicture(101, vbResCursor)
    End If
End Sub

Private Sub sldSpeed_Change()
    If sldSpeed.Value = 0 Then
        tmrVictory.Interval = 0
    ElseIf sldSpeed.Value = sldSpeed.Max Then
        tmrVictory.Interval = 1
    Else
        tmrVictory.Interval = (sldSpeed.Max - sldSpeed.Value) * 2
    End If
End Sub

Private Sub tmrTitle_Timer()
    If Me.WindowState <> vbMinimized Then
        Call PlayPlaneAni
    End If
End Sub

Private Sub tmrEffect_Timer()
    If Me.WindowState <> vbMinimized Then
        Call CE.StartEffect
    End If
End Sub

Private Sub tmrVictory_Timer()
    If Me.WindowState <> vbMinimized Then
        Select Case curAni
        Case Is = 0 ' Bounce
            Call VicAni.Bounce(Me, picViewer, "PlayingCard")
        Case Is = 1 ' Bounce (Scatter)
            Call VicAni.BounceScatter(Me, picViewer, chkTrail.Value, "PlayingCard")
        Case Is = 2 ' Bounce (Trail)
            Call VicAni.BounceTrail(Me, picViewer, crdAni(1), chkTrail.Value, "PlayingCard")
        Case Is = 3 ' Spin
            Call VicAni.Spin(Me, picViewer, "PlayingCard")
        Case Is = 4 ' Spin (Trail)
            Call VicAni.SpinTrail(Me, picViewer, crdAni(1), chkTrail.Value, "PlayingCard")
        Case Is = 5 ' Wave
            Call VicAni.Wave(Me, picViewer, "PlayingCard", Trim$(txtExpr.Text), _
                             crdAni(0).Width, crdAni(0).Height, chkClip.Value)
        End Select
    End If
End Sub

Private Sub txtExpr_Change()
    tmrVictory.Enabled = False
    cmdPlay.Enabled = True
    cmdStop.Enabled = False
    VicAni.Reset = True
End Sub

Private Sub RegsCtrlEffect()
    With CE
        .RegisterControl cmdCancel, vbBlue, vbCyan, vbWhite
        .RegisterControl cmdPlay, vbBlue, vbCyan, vbRed
        .RegisterControl cmdStop, vbBlue, vbCyan, vbRed
        .RegisterControl cmdGraph, vbBlue, vbCyan, vbYellow
        .RegisterControl cmdClear, vbBlue, vbCyan, vbYellow
        .RegisterControl cmdExpr, vbBlue, vbCyan, -1
        .RegisterControl cmbAni, vbBlue, vbCyan, -1
        .RegisterControl lstDistX, vbBlue, vbCyan, -1
        .RegisterControl lstDistY, vbBlue, vbCyan, -1
    End With
    
    tmrEffect.Enabled = True
End Sub
