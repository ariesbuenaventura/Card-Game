VERSION 5.00
Begin VB.Form frmExpr 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customize Expression"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   480
      Top             =   0
   End
   Begin VB.Frame fraTray 
      Caption         =   "Enter your expression"
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   3600
      Width           =   4815
      Begin VB.CommandButton cmdCancelExpr 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   4080
         TabIndex        =   7
         Top             =   225
         Width           =   675
      End
      Begin VB.CommandButton cmdOkExpr 
         Caption         =   "Ok"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3360
         TabIndex        =   6
         Top             =   225
         Width           =   675
      End
      Begin VB.TextBox txtExpr 
         Height          =   285
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   3255
      End
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
      Height          =   795
      Left            =   0
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   17
      Top             =   0
      Width           =   4875
   End
   Begin VB.PictureBox picBk 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   4875
      TabIndex        =   12
      Top             =   0
      Width           =   4875
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "or"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   16
         Top             =   780
         Width           =   165
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
         Left            =   3120
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   780
         Width           =   1665
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
         Left            =   540
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   780
         Width           =   2385
      End
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "email:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   780
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Frame fraFrame 
      Height          =   2655
      Left            =   0
      TabIndex        =   10
      Top             =   900
      Width           =   3675
      Begin VB.PictureBox picTray 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   90
         ScaleHeight     =   375
         ScaleWidth      =   3375
         TabIndex        =   18
         Top             =   2220
         Width           =   3375
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2580
            TabIndex        =   4
            Top             =   0
            Width           =   795
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   60
            TabIndex        =   1
            Top             =   0
            Width           =   795
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   900
            TabIndex        =   2
            Top             =   0
            Width           =   795
         End
         Begin VB.CommandButton cmdRemov 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1740
            TabIndex        =   3
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.ListBox lstExpr 
         Height          =   1620
         ItemData        =   "frmExpr.frx":0000
         Left            =   60
         List            =   "frmExpr.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   540
         Width           =   3555
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the expression would you like to used, then click OK."
         Height          =   495
         Left            =   180
         TabIndex        =   11
         Top             =   120
         Width           =   3285
      End
   End
End
Attribute VB_Name = "frmExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_SNOW_BALL = 100

Private Type SnowInfo
    xPos   As Integer
    yPos   As Integer
    Color  As Long
    Radius As Integer
    Speed  As Integer
    Weight As Integer
End Type

Public RetVal As String

Dim IsInitAni           As Boolean
Dim IsChanged           As Boolean
Dim Filename            As String
Dim SnowBallColor()     As Variant
Dim Snow(MAX_SNOW_BALL) As SnowInfo

Dim CE As New cAddControlEffect

Private Sub DrawGradientDsgn()
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
    
    Title = "List Expression"
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
    
    For i = 0 To 2000
        SetPixelV picTitle.hdc, CInt(Rnd * rcRect.Right), (CInt(Rnd * rcRect.Bottom)), _
                                RGB(CInt(Rnd * 125) + 125, CInt(Rnd * 125) + 125, CInt(Rnd * 125) + 125)
    Next i
    
    picTitle.ScaleMode = OldScaleMode
    Set picTitle.Picture = picTitle.Image
    tmrTitle.Enabled = True
End Sub

Private Sub PlaySnowAni(ByVal MaxSnowBalls As Integer, _
                        ByVal MaxRadius As Integer, _
                        ByVal MaxSpeed As Integer, _
                        ByVal MaxWeight As Integer, _
                        ByVal MaxWindVelocity As Integer, _
                        ByVal MaxWindLength As Integer)
    
    Dim i      As Integer
    Dim rcRect As RECT
    Dim WinSzX As Long
    Dim WinSzY As Long
    
    Static WindVel    As Integer ' Wind Velocity
    Static WindLen    As Integer ' Wind Length
    Static OldWindLen As Integer ' Old Wind Length
    
    GetClientRect picTitle.hwnd, rcRect
    WinSzX = rcRect.Right - rcRect.Left
    WinSzY = rcRect.Bottom - rcRect.Top

    If Not IsInitAni Then
        WindVel = 0
        WindLen = 0
        OldWindLen = CInt(Rnd * MaxWindLength)
        SnowBallColor = Array(&HFF, &HFF0000, &HFFBBAA, &HFFFF00, &HFF00FF)

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
        IsInitAni = True
    Else
        Dim nStat   As Integer
        Dim OffSetX As Integer
        Dim OffsetY As Integer
        
        Static bVal    As Integer
        
        picTitle.Cls
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
    picTitle.ForeColor = Color
    picTitle.DrawWidth = IIf(Radius = 0, 1, Radius)
    picTitle.PSet (X, Y)
End Sub

Private Sub cmdAdd_Click()
    If Not fraTray.Enabled Then
        fraTray.Enabled = True
        fraTray.Tag = "Add"
        fraTray.Caption = "Enter your expression:"
        txtExpr.Text = ""
        txtExpr.SetFocus
        
        Me.Height = Me.Height + fraTray.Height + _
                    Screen.TwipsPerPixelY * 5
        Call EnabledControl(False)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelExpr_Click()
    If fraTray.Enabled Then
        fraTray.Enabled = False
        Me.Height = Me.Height - fraTray.Height - _
                    Screen.TwipsPerPixelY * 5
        Call EnabledControl(True)
    End If
End Sub

Private Sub cmdEdit_Click()
    If Not fraTray.Enabled Then
        If lstExpr.ListIndex = -1 Then
            If lstExpr.ListCount > 0 Then
                lstExpr.Selected(lstExpr.TopIndex) = True
            End If
        End If
        
        fraTray.Enabled = True
        fraTray.Tag = "Edit"
        fraTray.Caption = "Enter your new expression:"
        txtExpr.Text = lstExpr.List(lstExpr.ListIndex)
        txtExpr.SetFocus
        
        Me.Height = Me.Height + fraTray.Height + _
                    Screen.TwipsPerPixelY * 5
        Call EnabledControl(False)
    End If
End Sub

Private Sub cmdOk_Click()
    If lstExpr.ListIndex <> -1 Then
        RetVal = lstExpr.List(lstExpr.ListIndex)
    Else
        RetVal = ""
    End If

    Unload Me
End Sub

Private Sub cmdOkExpr_Click()
    If Trim$(txtExpr.Text) = "" Then Exit Sub
    
    Dim TempScript As New MSScriptControl.ScriptControl
        
    On Error Resume Next
    
    TempScript.Language = "VBScript"
    Call TempScript.Eval(Trim$(txtExpr.Text))
    If Err Then
        MsgBox Err.Description, vbOKOnly Or vbInformation, "Error"
    Else
        Dim i  As Integer
        Dim s1 As String
        Dim s2 As String
        
        For i = 0 To lstExpr.ListCount - 1
            s1 = RemovSpace(LCase(Trim$(txtExpr.Text)))
            s2 = RemovSpace(LCase(lstExpr.List(i)))
            
            If s1 = s2 Then
                MsgBox "Expression already exist!", _
                        vbInformation Or vbOKOnly, fraTray.Tag
                Exit Sub
            End If
        Next i
        
        If fraTray.Tag = "Add" Then
            lstExpr.AddItem Trim$(txtExpr.Text)
        ElseIf fraTray.Tag = "Edit" Then
            Dim curIndex As Integer
            
            curIndex = lstExpr.ListIndex
            lstExpr.RemoveItem curIndex
            lstExpr.AddItem Trim$(txtExpr.Text), curIndex
            lstExpr.ListIndex = curIndex
        End If
        
        txtExpr.Text = ""
        IsChanged = True
    End If
End Sub

Private Sub cmdRemov_Click()
    If lstExpr.ListIndex <> -1 Then
        If MsgBox("Are you sure?", vbYesNo Or _
                   vbQuestion, "Remove") = vbYes Then
            lstExpr.RemoveItem lstExpr.ListIndex
            lstExpr.SetFocus
            If lstExpr.ListCount = 0 Then
                cmdEdit.Enabled = False
                cmdRemov.Enabled = False
            End If
            
            IsChanged = True
            cmdSave.Enabled = True
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    Call FileSave
    
    IsChanged = False
    cmdSave.Enabled = False
End Sub

Private Sub Form_Load()
    Filename = App.Path & "\data\expr.dat"
    
    Call ShowTitle
    Call FileOpen
    Call RegsCtrlEffect
    
    If lstExpr.ListCount > 0 Then
        cmdEdit.Enabled = True
        cmdRemov.Enabled = True
    End If
    IsChanged = False
    IsInitAni = False
    
    Set lblEmail(0).MouseIcon = LoadResPicture(101, vbResCursor)
    Set lblEmail(1).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub FileOpen()
    On Error GoTo OpenErr
    
    If Dir$(Filename) <> "" Then
        Dim InFile As Integer
        Dim Buffer As String
        
        lstExpr.Clear
        
        InFile = FreeFile
        Open Filename For Input As InFile
            Input #InFile, Buffer
            If CStr(Buffer) = GameSignature Then
                Do While Not EOF(InFile)
                    Input #InFile, Buffer
                    lstExpr.AddItem CStr(Buffer)
                Loop
            End If
        Close InFile
    End If
    Exit Sub

OpenErr:
End Sub

Private Sub FileSave()
    On Error GoTo SaveErr
    
    Dim i      As Integer
    Dim InFile As Integer
    
    InFile = FreeFile
    Open Filename For Output As InFile
        Write #InFile, GameSignature
        For i = 0 To lstExpr.ListCount - 1
            Write #InFile, lstExpr.List(i)
        Next i
    Close InFile
    Exit Sub
    
SaveErr:
End Sub

Private Sub EnabledControl(bVal As Boolean)
    cmdAdd.Enabled = bVal
    
    If IsChanged And Not fraTray.Enabled Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
    
    If lstExpr.ListCount = 0 Then
        cmdEdit.Enabled = False
        cmdRemov.Enabled = False
    Else
        cmdEdit.Enabled = bVal
        cmdRemov.Enabled = bVal
    End If
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

Private Sub Form_Resize()
    Call DrawGradientDsgn
End Sub

Private Sub lstExpr_Click()
    If fraTray.Enabled Then
        If lstExpr.ListIndex <> -1 Then
            txtExpr.Text = lstExpr.List(lstExpr.ListIndex)
        End If
    End If
End Sub

Private Sub tmrEffect_Timer()
    If Me.WindowState <> vbMinimized Then
        Call CE.StartEffect
    End If
End Sub

Private Sub tmrTitle_Timer()
    If Me.WindowState <> vbMinimized Then
        Call PlaySnowAni(UBound(Snow()), 3, 2, 2, 15, 15)
    End If
End Sub

Private Sub txtExpr_Change()
    If Len(txtExpr.Text) > 0 Then
        cmdOkExpr.Enabled = True
    Else
        cmdOkExpr.Enabled = False
    End If
End Sub

Private Sub RegsCtrlEffect()
    With CE
        .RegisterControl cmdAdd, vbBlue, vbCyan, vbRed
        .RegisterControl cmdEdit, vbBlue, vbCyan, vbRed
        .RegisterControl cmdRemov, vbBlue, vbCyan, vbRed
        .RegisterControl cmdSave, vbBlue, vbCyan, vbRed
        .RegisterControl cmdOk, vbBlue, vbCyan, vbWhite
        .RegisterControl cmdCancel, vbBlue, vbCyan, vbWhite
        .RegisterControl cmdOkExpr, vbBlue, vbCyan, -1
        .RegisterControl cmdCancelExpr, vbBlue, vbCyan, -1
    End With
    
    tmrEffect.Enabled = True
End Sub


