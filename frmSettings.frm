VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3780
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   60
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imlDeck 
      Left            =   600
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlMask 
      Left            =   1200
      Top             =   4860
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
            Picture         =   "frmSettings.frx":0000
            Key             =   "None"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":015A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraAniOp 
      Height          =   3615
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   5475
      Begin VB.Frame fraMisc 
         Caption         =   "Miscellaneous"
         Height          =   735
         Left            =   180
         TabIndex        =   18
         Top             =   2700
         Width           =   5115
         Begin VB.ListBox lstDistY 
            Height          =   255
            ItemData        =   "frmSettings.frx":06F4
            Left            =   2040
            List            =   "frmSettings.frx":0716
            TabIndex        =   20
            Top             =   300
            Width           =   675
         End
         Begin VB.ListBox lstDistX 
            Height          =   255
            ItemData        =   "frmSettings.frx":0739
            Left            =   780
            List            =   "frmSettings.frx":075B
            TabIndex        =   19
            Top             =   300
            Width           =   675
         End
         Begin MSComctlLib.Slider sldSpeed 
            Height          =   195
            Left            =   3300
            TabIndex        =   21
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   344
            _Version        =   393216
         End
         Begin VB.Label lblDist 
            AutoSize        =   -1  'True
            Caption         =   "Dist. Y:"
            Height          =   195
            Index           =   1
            Left            =   1500
            TabIndex        =   24
            Top             =   315
            Width           =   510
         End
         Begin VB.Label lblDist 
            AutoSize        =   -1  'True
            Caption         =   "Dist. X:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   315
            Width           =   510
         End
         Begin VB.Label lblSpeed 
            AutoSize        =   -1  'True
            Caption         =   "Speed:"
            Height          =   195
            Left            =   2760
            TabIndex        =   22
            Top             =   315
            Width           =   510
         End
      End
      Begin VB.Frame fraOp 
         Caption         =   "Choose animation"
         Height          =   2475
         Left            =   180
         TabIndex        =   7
         Top             =   180
         Width           =   5115
         Begin VB.CheckBox chkAniOp 
            Caption         =   "Random"
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   14
            Top             =   2160
            Value           =   2  'Grayed
            Width           =   1695
         End
         Begin VB.CheckBox chkAniOp 
            Caption         =   "Custom"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   13
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox cmbAni 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmSettings.frx":077E
            Left            =   600
            List            =   "frmSettings.frx":0794
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1260
            Width           =   3195
         End
         Begin VB.CheckBox chkTrail 
            Caption         =   "Trail"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3900
            TabIndex        =   11
            Top             =   1320
            Value           =   2  'Grayed
            Width           =   615
         End
         Begin VB.CheckBox chkClip 
            Caption         =   "Clip"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3900
            TabIndex        =   10
            Top             =   1860
            Value           =   2  'Grayed
            Width           =   555
         End
         Begin VB.CommandButton cmdExpr 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            TabIndex        =   9
            Top             =   1785
            Width           =   315
         End
         Begin VB.TextBox txtExpr 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1800
            Width           =   2835
         End
         Begin VB.Label lblEffect 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Effect:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   600
            TabIndex        =   17
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lblInstruc 
            BackStyle       =   0  'Transparent
            Caption         =   "Select the type of animation would you like to play when user wins the game."
            Height          =   435
            Left            =   180
            TabIndex        =   16
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label lblExpr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expression:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   600
            TabIndex        =   15
            Top             =   1620
            Width           =   810
         End
      End
   End
   Begin VB.Frame fraBk 
      Height          =   3615
      Left            =   180
      TabIndex        =   25
      Top             =   480
      Width           =   5475
      Begin VB.Frame fraBkOp 
         Caption         =   "Options"
         Height          =   2775
         Left            =   3120
         TabIndex        =   27
         Top             =   240
         Width           =   2115
         Begin VB.CommandButton cmdBitmap 
            Caption         =   "Change to Bitmap"
            Height          =   315
            Left            =   240
            TabIndex        =   29
            Top             =   1380
            Width           =   1635
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "Change to Color"
            Height          =   315
            Left            =   240
            TabIndex        =   28
            Top             =   960
            Width           =   1635
         End
      End
      Begin VB.Frame fraPreview 
         Caption         =   "Preview"
         Height          =   2775
         Left            =   180
         TabIndex        =   26
         Top             =   240
         Width           =   2835
         Begin VB.PictureBox picPreview 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00008000&
            Enabled         =   0   'False
            Height          =   2295
            Left            =   180
            ScaleHeight     =   2235
            ScaleWidth      =   2415
            TabIndex        =   30
            Top             =   300
            Width           =   2475
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsSettings 
      Height          =   4215
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   7435
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Animation"
            Key             =   "Ani"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Background"
            Key             =   "Bk"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BkColor As Long
Dim BkFile  As String
Dim BkMode  As Integer
Dim Speed   As Integer

Dim CE As New cAddControlEffect

Private Sub chkAniOp_Click(Index As Integer)
    Select Case Index
    Case Is = 0 ' Customize
        If chkAniOp(0).Value = vbChecked Then
            chkAniOp(0).Enabled = True
            chkAniOp(1).Enabled = False
            chkAniOp(1).Value = vbUnchecked
        Else
            chkAniOp(0).Enabled = False
            chkAniOp(1).Enabled = True
            chkAniOp(1).Value = vbChecked
        End If
        
        Dim bVal As Boolean
        
        bVal = chkAniOp(0).Enabled
        
        cmbAni.Enabled = bVal
        lblEffect.Enabled = bVal
        
        If cmbAni.List(cmbAni.ListIndex) = "Wave" Then
            If chkAniOp(0).Enabled Then
                chkClip.Enabled = True
                chkTrail.Enabled = True
                cmdExpr.Enabled = True
                lblExpr.Enabled = True
                txtExpr.Enabled = True
            Else
                chkClip.Enabled = False
                chkTrail.Enabled = False
                cmdExpr.Enabled = False
                lblExpr.Enabled = False
                txtExpr.Enabled = False
            End If
        End If
    Case Is = 1 ' Randomize
        If chkAniOp(1).Value = vbChecked Then
            chkAniOp(0).Enabled = False
            chkAniOp(1).Enabled = True
            chkAniOp(0).Value = vbUnchecked
        Else
            chkAniOp(0).Enabled = True
            chkAniOp(1).Enabled = False
            chkAniOp(0).Value = vbChecked
        End If
    End Select
    
    cmdApply.Enabled = True
End Sub

Private Sub chkClip_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkTrail_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmbAni_Click()
    If cmbAni.List(cmbAni.ListIndex) = "Wave" Then
        chkClip.Enabled = True
        cmdExpr.Enabled = True
        lblExpr.Enabled = True
        txtExpr.Enabled = True
    End If
    
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
    End Select
    
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
    If BkMode = 0 Then
        If GSI.BkFile <> BkFile Then
            GSI.BkFile = BkFile
            Call frmMain.SetBkBmp
        End If
    Else
        If GSI.BkColor <> BkColor Then
            GSI.BkColor = BkColor
            Set frmMain.picTable.Picture = Nothing
            frmMain.picTable.BackColor = BkColor
        End If
    End If

    GSI.BkMode = BkMode
    GSI.DistX = Val(lstDistX.List(lstDistX.ListIndex))
    GSI.DistY = Val(lstDistY.List(lstDistY.ListIndex))
    GSI.Clip = IIf(chkClip.Value = 0, False, True)
    GSI.Speed = Speed
    GSI.Trail = IIf(chkTrail.Value = 0, False, True)
    GSI.VicAniSel = cmbAni.ListIndex
    GSI.WaveExpr = txtExpr.Text
    
    frmMain.tmrVictory.Interval = GSI.Speed
    If GSI.Trail = False Then frmMain.picTable.Cls
    
    If chkAniOp(0).Value = vbChecked Then
        GSI.VicAniMode = 1
    Else
        GSI.VicAniMode = 0
    End If
    
    cmdApply.Enabled = False
End Sub

Private Sub cmdBitmap_Click()
    On Error GoTo Errhandler
    
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
            Dim rcRect As RECT
    
            GetClientRect picPreview.hwnd, rcRect
            TileBmp picPreview.hdc, LoadPicture(.Filename), _
                    rcRect.Right, rcRect.Bottom
            RefreshWindow picPreview.hwnd
            BkFile = .Filename
            BkMode = 0
        End If
    End With
    cmdApply.Enabled = True
    Exit Sub
    
Errhandler:
    If Err.Number = 32755 Then ' Cancel Selected
    Else
        MsgBox Err.Description, vbOKOnly Or vbInformation, "Error"
    End If
End Sub

Private Sub cmdDefault_Click()
    cmbAni.ListIndex = cmbAni.TopIndex
    chkClip.Value = vbChecked
    chkTrail.Value = vbChecked
    lstDistX.ListIndex = 4
    lstDistY.ListIndex = 4
    sldSpeed.Value = 4
    txtExpr.Text = ""
    
    chkAniOp(1).Value = vbChecked
    chkAniOp_Click 1
    sldSpeed_Change
End Sub

Private Sub cmdOk_Click()
    If cmdApply.Enabled Then cmdApply_Click
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    With frmVictoryAni
        .chkClip.Value = chkClip.Value
        .chkTrail.Value = chkTrail.Value
        .cmbAni.ListIndex = cmbAni.ListIndex
        .lstDistX.ListIndex = lstDistX.ListIndex
        .lstDistY.ListIndex = lstDistY.ListIndex
        .sldSpeed.Value = sldSpeed.Value
        .txtExpr = txtExpr.Text
        
        .Show vbModal, Me
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdColor_Click()
    On Error GoTo Errhandler
    
    dlgDialog.ShowColor
    picPreview.BackColor = dlgDialog.Color
    BkColor = dlgDialog.Color
    BkMode = 1
    cmdApply.Enabled = True
    Exit Sub

Errhandler:
End Sub

Private Sub cmdExpr_Click()
    frmExpr.Show vbModal, Me
    If frmExpr.RetVal <> "" Then
        txtExpr.Text = frmExpr.RetVal
        cmdApply.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Errhandler
    
    chkClip.Value = IIf(GSI.Clip, vbChecked, vbUnchecked)
    chkTrail.Value = IIf(GSI.Trail, vbChecked, vbUnchecked)
    cmbAni.ListIndex = GSI.VicAniSel
    lstDistX.Selected(GSI.DistX - 1) = True
    lstDistY.Selected(GSI.DistY - 1) = True
    If GSI.Speed = 0 Then
        sldSpeed.Value = 0
    ElseIf GSI.Speed = sldSpeed.Max * 2 Then
        sldSpeed.Value = 1
    Else
        sldSpeed.Value = sldSpeed.Max - GSI.Speed / 2
    End If
    
    txtExpr.Text = GSI.WaveExpr
    
    If GSI.BkMode = 0 Then
        Dim rcRect As RECT
        
        GetClientRect picPreview.hwnd, rcRect
        If Dir$(GSI.BkFile) <> "" Then
            TileBmp picPreview.hdc, LoadPicture(GSI.BkFile), _
                    rcRect.Right, rcRect.Bottom
        End If
        RefreshWindow picPreview.hwnd
    Else
        picPreview.Cls
        picPreview.BackColor = GSI.BkColor
    End If
    
    chkAniOp_Click GSI.VicAniMode
    Call RegsCtrlEffect
    
    BkColor = GSI.BkColor
    BkMode = GSI.BkMode
    cmdApply.Enabled = False
    Exit Sub
    
Errhandler:
End Sub

Private Sub lstDistX_Click()
    cmdApply.Enabled = True
End Sub

Private Sub lstDistY_Click()
    cmdApply.Enabled = True
End Sub

Private Sub sldSpeed_Change()
    If sldSpeed.Value = 0 Then
        Speed = 0
    ElseIf sldSpeed.Value = sldSpeed.Max Then
        Speed = 1
    Else
        Speed = (sldSpeed.Max - sldSpeed.Value) * 2
    End If

    cmdApply.Enabled = True
End Sub

Private Sub tbsSettings_Click()
    Select Case tbsSettings.SelectedItem.Index
    Case Is = 1 ' Animation
        cmdPreview.Visible = True
        fraAniOp.ZOrder 0
    Case Is = 2 ' Background
        cmdPreview.Visible = False
        fraBk.ZOrder 0
    End Select
End Sub

Private Sub tmrEffect_Timer()
    If Me.WindowState <> vbMinimized Then
        Call CE.StartEffect
    End If
End Sub

Private Sub RegsCtrlEffect()
    With CE
        .RegisterControl cmbAni, vbBlue, vbCyan, -1
        .RegisterControl cmdApply, vbBlue, vbCyan, vbWhite
        .RegisterControl cmdCancel, vbBlue, vbCyan, vbWhite
        .RegisterControl cmdOk, vbBlue, vbCyan, vbWhite
        .RegisterControl cmdDefault, vbBlue, vbCyan, vbRed
        .RegisterControl cmdPreview, vbBlue, vbCyan, vbRed
        .RegisterControl cmdColor, vbBlue, vbCyan, vbYellow
        .RegisterControl cmdBitmap, vbBlue, vbCyan, vbYellow
        .RegisterControl cmdExpr, vbBlue, vbCyan, -1
        .RegisterControl lstDistX, vbBlue, vbCyan, -1
        .RegisterControl lstDistY, vbBlue, vbCyan, -1
    End With
    
    tmrEffect.Enabled = True
End Sub

