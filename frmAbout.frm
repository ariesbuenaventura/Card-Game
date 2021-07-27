VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   Caption         =   "About Card Game version 1.0"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1020
      Top             =   60
   End
   Begin VB.Timer tmrTitle 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   540
      Top             =   60
   End
   Begin VB.Timer tmrAni3D 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   60
      Top             =   60
   End
   Begin VB.PictureBox picViewer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3435
      Left            =   0
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   438
      TabIndex        =   0
      Top             =   840
      Width           =   6630
      Begin VB.Label lblSchool 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Datapro Computer College, Inc. (Recto, Manila)"
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   1680
         TabIndex        =   20
         Top             =   3120
         Width           =   3345
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programmed by: Aris Buenaventura"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   2100
         TabIndex        =   19
         Top             =   2940
         Width           =   2490
      End
   End
   Begin VB.PictureBox picTray 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   6615
      TabIndex        =   1
      Top             =   4215
      Width           =   6615
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   5100
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox cmbObject 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1755
      End
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   195
         Left            =   1860
         TabIndex        =   5
         Top             =   600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   344
         _Version        =   393216
         Max             =   20
      End
      Begin VB.CommandButton cmdRot 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1260
         TabIndex        =   4
         Top             =   60
         Width           =   555
      End
      Begin VB.CommandButton cmdRot 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   660
         TabIndex        =   3
         Top             =   60
         Width           =   555
      End
      Begin VB.CommandButton cmdRot 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   555
      End
      Begin VB.CommandButton cmdArrow 
         Height          =   315
         Index           =   1
         Left            =   2460
         Picture         =   "frmAbout.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdArrow 
         Height          =   315
         Index           =   0
         Left            =   1920
         Picture         =   "frmAbout.frx":02E2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   495
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   420
         Width           =   465
      End
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6675
      TabIndex        =   10
      Top             =   0
      Width           =   6675
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1980
         Picture         =   "frmAbout.frx":05C4
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
         Top             =   60
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1680
         Picture         =   "frmAbout.frx":064E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         Top             =   60
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox picBk 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   -60
      ScaleHeight     =   855
      ScaleWidth      =   6675
      TabIndex        =   13
      Top             =   0
      Width           =   6675
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "email:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   18
         Top             =   600
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
         Left            =   1500
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   600
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
         Left            =   4080
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   600
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
         Left            =   3900
         TabIndex        =   15
         Top             =   600
         Width           =   165
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_BALL = 6
Private Const MAX_SNOW_BALL = 400

Private Type Point
    X1 As Single    ' Left
    Y1 As Single    ' Front
    Z1 As Single    ' Top
    X2 As Single    ' Right
    Y2 As Single    ' Back
    Z2 As Single    ' Bottom
End Type

Private Type SnowInfo
    xPos   As Integer
    yPos   As Integer
    Color  As Long
    Radius As Integer
    Speed  As Integer
    Weight As Integer
End Type

Private Type BallInfo
    cx      As Integer
    cy      As Integer
    dx      As Integer
    dy      As Integer
    Width   As Integer
    Height  As Integer
End Type

Dim SnowBallColor() As Variant

Dim Rotation As Integer
Dim ObjPts() As Point

Dim Ani3D      As New cGeometry
Dim RightEdge  As Integer
Dim BottomEdge As Integer
Dim DatObj     As New Collection
Dim CE         As New cAddControlEffect

Dim Ball(MAX_BALL - 1)  As BallInfo
Dim Snow(MAX_SNOW_BALL) As SnowInfo

Private Sub cmbObject_Click()
    On Error Resume Next
        
    tmrAni3D.Enabled = False
    Call SetObject(DatObj(cmbObject.ListIndex + 1))
    tmrAni3D.Enabled = True
End Sub

Private Sub cmdArrow_Click(Index As Integer)
    Ani3D.Angle = IIf(Index = 0, Rads(10), Rads(-10))
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRot_Click(Index As Integer)
    Rotation = Index
End Sub

Private Sub Form_Load()
    Set Me.Icon = frmMain.Icon
    
    Dim PathObjFile As String
    
    PathObjFile = App.Path & "\data\object.dat"
    If Dir$(PathObjFile) = "" Then
        MsgBox """" & PathObjFile & """ doest not exist!", _
               vbInformation Or vbOKOnly, "Error"
        Call DisabledControl
        Exit Sub
    End If
    
    Dim InFile As Long
    Dim Buffer As String
    
    InFile = FreeFile
    Open PathObjFile For Input As InFile
        Input #InFile, Buffer
        ' check the signature
        If CStr(Buffer) = GameSignature Then
            ' if signature is correct then read the file
            Do While Not EOF(InFile)
                Input #InFile, Buffer
                DatObj.Add Buffer
            
                cmbObject.AddItem "Object " & DatObj.Count
            Loop
        Else
            ' the file is not supported by this program
            MsgBox "File format error!", vbOKOnly Or vbInformation, "Error"
            Call DisabledControl
        End If
    Close InFile
    
    If DatObj.Count > 0 Then
        Rotation = 0
        sldSpeed.Value = sldSpeed.Max * 0.5
    
        Call sldSpeed_Change
        Call SetObject(DatObj(1))
    
        cmbObject.ListIndex = cmbObject.TopIndex
        tmrAni3D.Enabled = True
    End If
            
    Call BkGradient
    Call ShowTitle
    Call RegsCtrlEffect
    
    Set lblEmail(0).MouseIcon = LoadResPicture(101, vbResCursor)
    Set lblEmail(1).MouseIcon = LoadResPicture(101, vbResCursor)
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Error"
    Call DisabledControl
End Sub

Private Sub SetObject(Data As String)
    Dim arrPts() As String
    
    arrPts = Split(Data, ";")
    
    Dim i As Integer, n As Integer
    ReDim ObjPts(CInt(UBound(arrPts) / 6) - 1) As Point
    
    n = 0
    For i = LBound(ObjPts()) To UBound(ObjPts())
        ObjPts(i).X1 = CSng(arrPts(n + 0))
        ObjPts(i).Y1 = CSng(arrPts(n + 1))
        ObjPts(i).Z1 = CSng(arrPts(n + 2))
        ObjPts(i).X2 = CSng(arrPts(n + 3))
        ObjPts(i).Y2 = CSng(arrPts(n + 4))
        ObjPts(i).Z2 = CSng(arrPts(n + 5))
        n = n + 6
    Next
End Sub
 
Private Sub DrawObject(ByVal X1 As Integer, ByVal Y1 As Integer, _
                       ByVal X2 As Integer, ByVal Y2 As Integer, _
                       ByVal nWidth As Long, ByVal Color As Long)
                       
    Dim i      As Integer
    Dim lhDC   As Long
    Dim pX1    As Single
    Dim pY1    As Single
    Dim pX2    As Single
    Dim pY2    As Single
    Dim xmid   As Single
    Dim ymid   As Single
    
    lhDC = picViewer.hdc
    For i = LBound(ObjPts()) To UBound(ObjPts())
        ' scale the object
        pX1 = (X2 - X1) * ObjPts(i).X1 / 100
        pY1 = (Y2 - Y1) * ObjPts(i).Y1 / 100
        pX2 = (X2 - X1) * ObjPts(i).X2 / 100
        pY2 = (Y2 - Y1) * ObjPts(i).Y2 / 100
        
        xmid = X1 + (X2 - X1) / 2
        ymid = Y1 + (Y2 - Y1) / 2
        DrawLine lhDC, xmid + pX1, ymid + pY1, _
                       xmid + pX2, ymid + pY2, nWidth, Color
    Next i
    
    RefreshWindow picViewer.hwnd
End Sub

Private Sub PlayAni3D()
    Dim i As Integer
    
    For i = LBound(ObjPts()) To UBound(ObjPts())
        If Rotation = 0 Then
            Call Ani3D.rotaboutx(ObjPts(i).X1, ObjPts(i).Y1, ObjPts(i).Z1)
            Call Ani3D.rotaboutx(ObjPts(i).X2, ObjPts(i).Y2, ObjPts(i).Z2)
        ElseIf Rotation = 1 Then
            Call Ani3D.rotabouty(ObjPts(i).X1, ObjPts(i).Y1, ObjPts(i).Z1)
            Call Ani3D.rotabouty(ObjPts(i).X2, ObjPts(i).Y2, ObjPts(i).Z2)
        Else
            Call Ani3D.rotaboutz(ObjPts(i).X1, ObjPts(i).Y1, ObjPts(i).Z1)
            Call Ani3D.rotaboutz(ObjPts(i).X2, ObjPts(i).Y2, ObjPts(i).Z2)
        End If
    Next i
    
    Dim xmid   As Single
    Dim ymid   As Single
    Dim xWid   As Single
    Dim yWid   As Single
    
    Dim rcRect As RECT
    
    GetClientRect picViewer.hwnd, rcRect
    xWid = rcRect.Right - rcRect.Left
    yWid = rcRect.Bottom - rcRect.Top
    xmid = xWid / 2: ymid = yWid / 2
    
    Call DrawObject(0, 0, GetWinValX(15), GetWinValY(15), 1, vbYellow)
    Call DrawObject(xWid - GetWinValX(15), 0, xWid, GetWinValY(15), 1, vbGreen)
    Call DrawObject(0, yWid - GetWinValY(15), GetWinValX(15), yWid, 1, vbRed)
    Call DrawObject(xWid - GetWinValX(15), yWid - GetWinValY(15), _
                    xWid, yWid, 1, vbCyan)
    Call DrawObject(xmid - GetWinValX(35), ymid - GetWinValY(35), _
                    xmid + GetWinValX(35), ymid + GetWinValY(35), 1, vbWhite)
End Sub

Private Function GetWinValX(Percent As Integer) As Single
    Dim rcRect As RECT
    
    GetClientRect picViewer.hwnd, rcRect
    GetWinValX = (rcRect.Right - rcRect.Left) * Percent / 100
End Function

Private Function GetWinValY(Percent As Integer) As Single
    Dim rcRect As RECT
    
    GetClientRect picViewer.hwnd, rcRect
    GetWinValY = (rcRect.Bottom - rcRect.Top) * Percent / 100
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    picViewer.Width = Me.Width - Screen.TwipsPerPixelX * 6
    picTitle.Move (Me.Width - picTitle.Width) / 2, -110
    picBk.Move 0, 0, Me.Width, picBk.Height
    picViewer.Height = Me.ScaleHeight - picTray.ScaleHeight - _
                       picViewer.Top + Screen.TwipsPerPixelX * 4
    lblAuthor.Move (picViewer.ScaleWidth - lblAuthor.Width) / 2, _
                    picViewer.ScaleHeight - lblAuthor.Height - lblSchool.Height - 2
    lblSchool.Move (picViewer.ScaleWidth - lblSchool.Width) / 2, _
                    picViewer.ScaleHeight - lblSchool.Height - 2
    cmdClose.Left = Me.Width - cmdClose.Width - Screen.TwipsPerPixelX * 10
    
    Dim xmid As Integer
    
    xmid = (picBk.ScaleWidth - (lblPrompt(0).Width + lblPrompt(1).Width + _
            lblEmail(0).Width + lblEmail(1).Width)) / 2
    lblPrompt(0).Left = xmid
    lblEmail(0).Left = xmid + lblPrompt(0).Width
    lblPrompt(1).Left = xmid + lblPrompt(0).Width + lblEmail(0).Width
    lblEmail(1).Left = xmid + lblPrompt(0).Width + lblEmail(0).Width + _
                       lblPrompt(1).Width
                     
    Call BkGradient
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrAni3D.Enabled = False
    tmrEffect.Enabled = False
    tmrTitle.Enabled = False
    
    Set frmAbout = Nothing
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

Private Sub sldSpeed_Change()
    If sldSpeed.Value = 0 Then
        tmrAni3D.Interval = 0
    ElseIf sldSpeed.Value = sldSpeed.Max Then
        tmrAni3D.Interval = 1
    Else
        tmrAni3D.Interval = (sldSpeed.Max - sldSpeed.Value) * 2
    End If
End Sub

Private Sub DisabledControl()
    cmbObject.Enabled = False
    cmdRot(0).Enabled = False
    cmdRot(1).Enabled = False
    cmdRot(2).Enabled = False
    sldSpeed.Enabled = False
End Sub

Private Sub ShowTitle()
    Dim lhFont    As Long
    Dim lhrgn     As Long
    Dim lhOldFont As Long
    Dim rcRect    As RECT
    Dim xPos      As Integer
    Dim yPos      As Integer
    Dim sz        As Size
    
    Const CLIP_LH_ANGLES = 16
    Const PROOF_QUALITY As Long = 2
        
    lhFont = CreateFont(72, 17, 0, 0, 700, _
                        False, False, False, _
                        0, 0, CLIP_LH_ANGLES, PROOF_QUALITY, 0, "Arial Black")

    lhOldFont = SelectObject(picTitle.hdc, lhFont)
    
    GetClientRect picTitle.hwnd, rcRect
    GetTextExtentPoint32 picTitle.hdc, App.Title, Len(App.Title), sz
    
    xPos = (rcRect.Right - rcRect.Left - sz.cx) / 2
    yPos = (rcRect.Bottom - rcRect.Top - sz.cy) / 2
    
    BeginPath picTitle.hdc
    TextOut picTitle.hdc, xPos, yPos, App.Title, Len(App.Title)
    EndPath picTitle.hdc
    lhrgn = PathToRegion(picTitle.hdc)
    SetWindowRgn picTitle.hwnd, lhrgn, True
    
    SelectObject picTitle.hdc, lhOldFont
    DeleteObject lhFont
    
    Dim i As Integer, OldScaleMode As Integer
    
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
    
    ' add glitters
    For i = 0 To 2000
        SetPixelV picTitle.hdc, CInt(Rnd * rcRect.Right), (CInt(Rnd * rcRect.Bottom)), _
                                RGB(CInt(Rnd * 125) + 125, CInt(Rnd * 125) + 125, CInt(Rnd * 125) + 125)
    Next i
    
    picTitle.ScaleMode = OldScaleMode
    Set picTitle.Picture = picTitle.Image
    tmrTitle.Enabled = True
End Sub

Private Sub tmrAni3D_Timer()
    If Me.WindowState <> vbMinimized Then
        Dim rcRect        As RECT
        Dim TotalSnowBall As Integer
        Dim WinSzX        As Long
        Dim WinSzY        As Long
        
        picViewer.Cls
        
        GetClientRect picViewer.hwnd, rcRect
        WinSzX = rcRect.Right - rcRect.Left
        WinSzY = rcRect.Bottom - rcRect.Top
        If WinSzX < WinSzY Then
            TotalSnowBall = WinSzY * 0.4
        Else
            TotalSnowBall = WinSzX * 0.4
        End If
        If TotalSnowBall > MAX_SNOW_BALL Then
            TotalSnowBall = MAX_SNOW_BALL
        End If
    
        Call PlaySnowAni(TotalSnowBall, 2, 3, 3, 15, 15)
        Call PlayAni3D
    End If
End Sub

Private Sub PlaySnowAni(ByVal MaxSnowBalls As Integer, _
                        ByVal MaxRadius As Integer, _
                        ByVal MaxSpeed As Integer, _
                        ByVal MaxWeight As Integer, _
                        ByVal MaxWindVelocity As Integer, _
                        ByVal MaxWindLength As Integer)
                
    Static bInit As Boolean
    
    Dim i      As Integer
    Dim rcRect As RECT
    Dim WinSzX As Long
    Dim WinSzY As Long
    
    Static WindVel    As Integer ' Wind Velocity
    Static WindLen    As Integer ' Wind Length
    Static OldWindLen As Integer ' Old Wind Length
    
    GetClientRect picViewer.hwnd, rcRect
    WinSzX = rcRect.Right - rcRect.Left
    WinSzY = rcRect.Bottom - rcRect.Top
    
    If Not bInit Then
        WindVel = 0
        WindLen = 0
        OldWindLen = CInt(Rnd * MaxWindLength)
        SnowBallColor = Array(&HFFFFFF, &HF5F5F5, &HEBEBEB, &HE1E1E1, &HCDCDCD)
        
        For i = LBound(Snow()) To UBound(Snow())
            Snow(i).xPos = CInt(Rnd * WinSzX)
            Snow(i).yPos = CInt(Rnd * WinSzY)
            Snow(i).Color = CLng(SnowBallColor(Rnd * UBound(SnowBallColor)))
            Snow(i).Radius = CInt(Rnd * MaxRadius) + 1
            Snow(i).Speed = CInt(Rnd * MaxSpeed) + MaxSpeed / 2
            Snow(i).Weight = CInt(Rnd * MaxWeight) + 1
            DrawSnowBall Snow(i).xPos, Snow(i).yPos, _
                         Snow(i).Radius, Snow(i).Color
        Next i
        bInit = True
    Else
        Dim nStat   As Integer
        Dim OffSetX As Integer
        Dim OffsetY As Integer
        
        Static bVal    As Integer
        
        For i = 0 To MaxSnowBalls
            Snow(i).xPos = Snow(i).xPos + WindVel - Snow(i).Weight
            Snow(i).yPos = Snow(i).yPos + Snow(i).Speed + Snow(i).Weight
            
            If Snow(i).xPos < -Snow(i).Radius Then
                nStat = 0
            ElseIf Snow(i).xPos > WinSzX + Snow(i).Radius Then
                nStat = 1
            ElseIf Snow(i).yPos < -Snow(i).Radius Then
                nStat = 2
            ElseIf Snow(i).yPos > WinSzY + Snow(i).Radius Then
                nStat = 3
            Else
                nStat = -1
            End If
            
            If nStat <> -1 Then
                Snow(i).Radius = CInt(Rnd * MaxRadius) + 1
            End If
            
            Select Case nStat
            Case Is = 0
                OffSetX = WinSzX + Snow(i).Radius
                OffsetY = CInt(Rnd * WinSzY)
            Case Is = 1
                OffSetX = -Snow(i).Radius
                OffsetY = CInt(Rnd * WinSzY)
            Case Is = 2
                OffSetX = CInt(Rnd * WinSzX)
                OffsetY = WinSzY + Snow(i).Radius
            Case Is = 3
                OffSetX = CInt(Rnd * WinSzX)
                OffsetY = -Snow(i).Radius
            End Select
        
            If nStat <> -1 Then
                Snow(i).xPos = OffSetX
                Snow(i).yPos = OffsetY
                Snow(i).Color = CLng(SnowBallColor(Rnd * UBound(SnowBallColor)))
                Snow(i).Speed = CInt(Rnd * MaxSpeed) + MaxSpeed / 2
                Snow(i).Weight = CInt(Rnd * MaxWeight) + 1
            End If
            
            DrawSnowBall Snow(i).xPos, Snow(i).yPos, _
                         Snow(i).Radius, Snow(i).Color
        Next i
        
        If bVal Then
            If OldWindLen = WindLen Then WindVel = WindVel + 1
        Else
            If OldWindLen = WindLen Then WindVel = WindVel - 1
        End If
        
        If OldWindLen = WindLen Then
            WindLen = 0
            OldWindLen = CInt(Rnd * MaxWindLength)
        Else
            WindLen = WindLen + 1
        End If
        
        If (WindVel = MaxWindVelocity) And bVal Then
            bVal = False
        ElseIf (WindVel = -MaxWindVelocity) And Not bVal Then
            bVal = True
        End If
    End If
End Sub

Private Sub DrawSnowBall(ByVal X As Long, ByVal Y As Long, ByVal Radius, Color As Long)
    picViewer.ForeColor = Color
    picViewer.DrawWidth = IIf(Radius = 0, 1, Radius)
    picViewer.PSet (X, Y)
End Sub

Private Sub PlayBouncingBallAni()
    Dim i    As Integer
    Dim Temp As Integer
    
    Static bInit As Boolean
    
    If Not bInit Then
        Dim BMP()  As BITMAP
        Dim rcRect As RECT
        ReDim BMP(UBound(Ball())) As BITMAP
        
        GetClientRect picTitle.hwnd, rcRect
        RightEdge = rcRect.Right - rcRect.Left
        BottomEdge = rcRect.Bottom - rcRect.Top
            
        For i = LBound(Ball()) To UBound(Ball())
            GetObjectAPI picSprite.Picture.Handle, Len(BMP(i)), BMP(i)
        
            ' offset cx
            Temp = IIf(CInt(Rnd * -1), 1, -1)
            If Temp = 1 Then
                Ball(i).cx = 0
            Else
                Ball(i).cx = RightEdge
            End If
        
            ' offset cy
            Ball(i).cy = (i * 5) + BMP(i).bmHeight / 2 + 2
        
            Ball(i).dx = CInt(Rnd * 5) + 1 ' speed x
            Ball(i).dy = CInt(Rnd * 5) + 1 ' speed y
            Ball(i).Width = BMP(i).bmWidth
            Ball(i).Height = BMP(i).bmHeight
       Next i
                
        bInit = True
    Else
        Dim NewX As Integer
        Dim NewY As Integer
        
        picTitle.Cls

        For i = LBound(Ball()) To UBound(Ball())
            Temp = Ball(i).cx + Ball(i).dx
            If Temp + Ball(i).Width > RightEdge Then
                Ball(i).dx = -Abs(Ball(i).dx)
            ElseIf Temp < 0 Then
                Ball(i).dx = Abs(Ball(i).dx)
            End If
    
            NewX = Ball(i).cx + Ball(i).dx
        
            Temp = Ball(i).cy + Ball(i).dy
            If Temp + Ball(i).Height > BottomEdge Then
                Ball(i).dy = -Abs(Ball(i).dy)
            ElseIf Temp < 0 Then
                Ball(i).dy = Abs(Ball(i).dy)
            End If
        
            NewY = Ball(i).cy + Ball(i).dy
               
            BitBlt picTitle.hdc, NewX, NewY, Ball(i).Width, Ball(i).Height, _
                   picMask.hdc, 0, 0, vbSrcAnd
            BitBlt picTitle.hdc, NewX, NewY, Ball(i).Width, Ball(i).Height, _
                   picSprite.hdc, 0, 0, vbSrcInvert
            
            Ball(i).cx = NewX
            Ball(i).cy = NewY
        Next i
        
        RefreshWindow picTitle.hwnd
    End If
End Sub

Private Sub tmrEffect_Timer()
    If Me.WindowState <> vbMinimized Then
        Call CE.StartEffect
    End If
End Sub

Private Sub tmrTitle_Timer()
    If Me.WindowState <> vbMinimized Then
        Call PlayBouncingBallAni
    End If
End Sub

Private Sub BkGradient()
    Dim i As Integer, OldScaleMode As Integer
    
    OldScaleMode = picBk.ScaleMode
    picBk.ScaleMode = vbUser
    picBk.ScaleWidth = 255
    picBk.ScaleHeight = 255
    
    picBk.Cls
    For i = 0 To 255
        picBk.Line (0, i)-(picBk.ScaleWidth, i), RGB(0, 0, 255 - i)
    Next i
    
    picBk.ScaleMode = OldScaleMode
End Sub

Private Sub RegsCtrlEffect()
    With CE
        .RegisterControl cmdRot(0), vbBlue, vbCyan, vbRed
        .RegisterControl cmdRot(1), vbBlue, vbCyan, vbRed
        .RegisterControl cmdRot(2), vbBlue, vbCyan, vbRed
        .RegisterControl cmdArrow(0), vbBlue, vbCyan, vbYellow
        .RegisterControl cmdArrow(1), vbBlue, vbCyan, vbYellow
        .RegisterControl cmdClose, vbBlue, vbCyan, vbWhite
        .RegisterControl cmbObject, vbBlue, vbCyan, -1
    End With
    
    tmrEffect.Enabled = True
End Sub
