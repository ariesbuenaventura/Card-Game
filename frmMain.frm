VERSION 5.00
Object = "*\AprjCard.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Card Game version 1.0"
   ClientHeight    =   3870
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7125
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   475
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   120
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":107C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1920
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21C4
            Key             =   "About"
            Object.Tag             =   "About"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2616
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A68
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblToolbar 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   1111
      ButtonWidth     =   1852
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Select"
            Key             =   "Select"
            Object.ToolTipText     =   "Select Game"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "ReStart"
            Key             =   "ReStart"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "ReStart Game"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Open"
            Key             =   "Open"
            Object.ToolTipText     =   "Open Game"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            Key             =   "Save"
            Object.ToolTipText     =   "Save Game"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Card Options"
            Key             =   "Card"
            Object.ToolTipText     =   "Card Options"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            Key             =   "Settings"
            Object.ToolTipText     =   "Settings"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Rules"
            Key             =   "Rules"
            Object.ToolTipText     =   "Rules"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            Object.ToolTipText     =   "About"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrVictory 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1380
   End
   Begin MSComDlg.CommonDialog dlgGame 
      Left            =   120
      Top             =   1860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3615
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6906
            Text            =   "Welcome..."
            TextSave        =   "Welcome..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "6/21/2003"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:43 AM"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTable 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      Height          =   3015
      Left            =   0
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   1
      Top             =   600
      Width           =   7125
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   1920
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   27
         TabIndex        =   8
         Top             =   780
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   1440
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   27
         TabIndex        =   7
         Top             =   780
         Visible         =   0   'False
         Width           =   405
      End
      Begin prjCard.Card crdPlayingCard 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Tag             =   "PlayingCard"
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         DeckMaskPicture =   "frmMain.frx":521A
         DeckPicture     =   "frmMain.frx":5236
         Elevator        =   0
         Face            =   1
         FlyIn           =   0
         FlyOut          =   1
         HotTracking     =   -1  'True
         ShowFocusRect   =   -1  'True
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
         FontTransparent =   0   'False
         ForeColor       =   -2147483640
         MousePointer    =   99
         Picture         =   "frmMain.frx":5252
      End
      Begin prjCard.Card crdStock 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   4
         Tag             =   "PlayingCard"
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         AutoFlipCard    =   0   'False
         DeckMaskPicture =   "frmMain.frx":5618
         DeckPicture     =   "frmMain.frx":5634
         Elevator        =   0
         Face            =   1
         FlyIn           =   0
         FlyOut          =   0
         HotTracking     =   -1  'True
         ShowFocusRect   =   -1  'True
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
         MousePointer    =   99
         Picture         =   "frmMain.frx":5650
      End
      Begin prjCard.Card crdWaste 
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   3
         Tag             =   "PlayingCard"
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         AutoFlipCard    =   0   'False
         DeckMaskPicture =   "frmMain.frx":5A16
         DeckPicture     =   "frmMain.frx":5A32
         Elevator        =   0
         FlyIn           =   0
         FlyOut          =   0
         HotTracking     =   -1  'True
         ShowFocusRect   =   -1  'True
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
         MousePointer    =   99
         Picture         =   "frmMain.frx":5A4E
      End
      Begin prjCard.Card crdTemp 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Tag             =   "PlayingCard"
         Top             =   300
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         DeckMaskPicture =   "frmMain.frx":5E14
         DeckPicture     =   "frmMain.frx":5E30
         Elevator        =   0
         Face            =   1
         FlyIn           =   0
         FlyOut          =   0
         HotTracking     =   -1  'True
         ShowFocusRect   =   -1  'True
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
         MousePointer    =   99
         Picture         =   "frmMain.frx":5E4C
      End
      Begin VB.Image imgStock 
         Height          =   255
         Left            =   900
         MousePointer    =   99  'Custom
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgWaste 
         Height          =   255
         Left            =   1200
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape shpWaste 
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         FillStyle       =   7  'Diagonal Cross
         Height          =   255
         Left            =   1500
         Top             =   60
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape shpStock 
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         FillStyle       =   7  'Diagonal Cross
         Height          =   255
         Left            =   1800
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H0000C000&
         BorderWidth     =   7
         Height          =   255
         Index           =   0
         Left            =   2100
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameSelect 
         Caption         =   "&Select Game"
      End
      Begin VB.Menu mnuGameRestart 
         Caption         =   "&ReStart Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGameBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameOpen 
         Caption         =   "Open Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGameSave 
         Caption         =   "Save Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGameToolbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsCardSize 
         Caption         =   "Card Size"
         Begin VB.Menu mnuOptionsCardSizeOp 
            Caption         =   "Automatic"
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuOptionsCardSizeOp 
            Caption         =   "Small"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mnuOptionsCardSizeOp 
            Caption         =   "Standard"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu mnuOptionsCardSizeOp 
            Caption         =   "Large"
            Enabled         =   0   'False
            Index           =   3
         End
      End
      Begin VB.Menu mnuOptionsBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsCard 
         Caption         =   "Card Options"
      End
      Begin VB.Menu mnuOptionsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsSettings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHowToPlay 
         Caption         =   "How To Play"
         Enabled         =   0   'False
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Card Game"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curVicAni As Integer
Dim IsRestart As Boolean

Dim Trigo   As New cTrigonometry
Dim solGame As New cGame
Dim VicAni  As New cVictoryAni

Private Sub crdPlayingCard_Click(Index As Integer)
    If VBScript.Run("Process", crdPlayingCard(Index)) Then
        Call PlayVictoryAni
        Call EnabledControl(False)
    End If
    Call ShowScore
End Sub
    
Private Sub crdPlayingCard_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call VBScript.Run("DoKeyDown", Index, KeyCode)
    Call ShowScore
End Sub

Private Sub crdPlayingCard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set crdPlayingCard(Index).MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub crdPlayingCard_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set crdPlayingCard(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub crdStock_Click(Index As Integer)
    Call VBScript.Run("DoStockClick", crdStock(Index))
    Call ShowScore
End Sub

Private Sub crdStock_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set crdStock(Index).MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub crdStock_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set crdStock(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub crdWaste_Click(Index As Integer)
    If VBScript.Run("Process", crdWaste(Index)) Then
        Call PlayVictoryAni
        Call EnabledControl(False)
    End If
    Call ShowScore
End Sub

Private Sub crdWaste_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set crdPlayingCard(Index).MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub crdWaste_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set crdPlayingCard(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Form_Load()
    Dim oCard As Object
        
    Call OpenSettings
    Call InitLogo
    Call RestoreCardSettings
    Call CreateObject(crdWaste, 2)
    Call CreateObject(crdPlayingCard, TotalCards)
    For Each oCard In Me.Controls
        If oCard.Tag = "PlayingCard" Then
            Set oCard.MouseIcon = LoadResPicture(101, vbResCursor)
            oCard.Move -oCard.Width * 5, -oCard.Height * 5
        End If
    Next oCard
    
    imgStock.Move -imgStock.Width * 5, -imgStock.Height * 5
    imgWaste.Move -imgWaste.Width * 5, -imgWaste.Height * 5
    shpCircle(0).Move -shpCircle(0).Width * 5, -shpCircle(0).Height * 5

    Set imgStock.MouseIcon = LoadResPicture(101, vbResCursor)
    App.HelpFile = App.Path & "\card.hlp"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    picTable.Move 0, tblToolbar.Height, Me.ScaleWidth, _
                  Me.ScaleHeight - sbStatusBar.Height - tblToolbar.Height
    If solGame.Signature <> "" Then
        If GSI.BkMode = 0 Then Call SetBkBmp
    End If
    If Not tmrVictory.Enabled Then
        If solGame.Signature <> "" Then
            Dim lret As Long
            
            lret = LockWindowUpdate(Me.hwnd)
            Call VBScript.Run("AlignCards")
            If lret Then Call LockWindowUpdate(0)
        End If
    End If
    
    If solGame.Signature = "" Then Call RepaintLogo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim oCard As Object
    
    Call EndPlaySound
    WinHelp 0, App.HelpFile, HELP_QUIT, 0

    If solGame.Signature <> "" Then
        Do While VBScript.Run("IsCardAniOn", crdPlayingCard)
            For Each oCard In crdPlayingCard
                If oCard.Index <> 0 Then
                     oCard.StopAni = True
                End If
            Next oCard
    
            DoEvents
        Loop
    End If
    
    Call DestroyObject(crdPlayingCard)
    Call DestroyObject(crdWaste)
    Call DestroyObject(crdStock)
    Call DestroyObject(crdTemp)
    Call SaveCardSettings
    Call SaveSettings
    End
End Sub

Private Sub imgStock_Click()
    Call VBScript.Run("DoStockClick", crdStock(0))
End Sub

Private Sub imgStock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgStock.MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub imgStock_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgStock.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub mnuGameOpen_Click()
    On Error GoTo OpenErr
    
    With dlgGame
        .Filter = "Card Game File (*.cgf) | *.cgf; |All Files (*.*) | *.*"
        .FilterIndex = 1
        .Filename = ""
        .InitDir = App.Path & "\Save"
        .ShowOpen
    
        If .Filename <> "" Then
            Dim InFile         As Long
            Dim arrData()      As String
            Dim sData          As String
            Dim oCard          As Object
            Dim FileSignature  As String
            Dim lret           As Long
            
            InFile = FreeFile
            Open .Filename For Input As InFile
                Input #InFile, sData ' Game Signature
                If CStr(sData) <> GameSignature Then
                    MsgBox "File format error!", vbOKOnly Or vbCritical, "Error"
                    Exit Sub
                    Close InFile
                End If
                Input #InFile, FileSignature ' File Signature
                Input #InFile, sData ' Title
                If CStr(FileSignature) <> solGame.Signature Then
                    MsgBox "Game Type: " & CStr(sData) & ".", vbOKOnly Or vbInformation, "Error"
                    Close InFile
                    Exit Sub
                End If

                tmrVictory.Enabled = False
                picTable.ScaleMode = vbTwips
                For Each oCard In Me.Controls
                    If oCard.Tag = "PlayingCard" Then
                        oCard.Visible = False
                    End If
                Next oCard
                
                Input #InFile, sData ' Score
                solGame.Score = CLng(sData)
                Input #InFile, sData ' ValHolder1
                solGame.ValHolder1 = sData
                Input #InFile, sData ' ValHolder2
                solGame.ValHolder2 = sData
                Input #InFile, sData ' ValHolder3
                solGame.ValHolder3 = sData
                Call VBScript.Run("SetValHolder")
            
                Input #InFile, sData ' DataColl
                Call VBScript.Run("SetDataColl", CStr(sData))
                Input #InFile, sData ' TempColl
                Call VBScript.Run("SetTempColl", CStr(sData))
                Input #InFile, sData ' StockColl
                Call VBScript.Run("SetStockColl", CStr(sData))
                Input #InFile, sData ' WasteColl
                Call VBScript.Run("SetWasteColl", CStr(sData))
                
                lret = LockWindowUpdate(Me.hwnd)
                
                Do While Not EOF(InFile)
                    Input #InFile, sData
                    arrData = Split(sData, "/$")
                    For Each oCard In Me.Controls
                        If oCard.Tag = "PlayingCard" Then
                            If (arrData(0) = oCard.Name) And (arrData(1) = oCard.Index) Then
                                oCard.Update = False
                                oCard.Data = CStr(arrData(2))
                                oCard.Enabled = True
                                oCard.Face = CInt(arrData(3))
                                oCard.Left = CLng(arrData(4))
                                oCard.Top = CLng(arrData(5))
                                oCard.Rank = CInt(arrData(6))
                                oCard.Suit = CInt(arrData(7))
                                oCard.Selected = CBool(arrData(8))
                                oCard.Visible = CBool(arrData(9))
                                oCard.Update = True
                                oCard.ZOrder 0
                                oCard.Refresh
                            End If
                        End If
                    Next oCard
                Loop
            Close InFile
        
            Call VBScript.Run("AlignCards")
            Call VBScript.Run("UpdateControls")
            Call ShowScore
            If lret Then Call LockWindowUpdate(0)
        End If
    End With
    Exit Sub
    
OpenErr:
    If InFile <> 0 Then Close InFile
    If lret Then Call LockWindowUpdate(0)
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Error"
End Sub

Private Sub mnuGameSave_Click()
    If solGame.Signature = "" Then Exit Sub
    
    On Error GoTo SaveErr
    
    With dlgGame
        .Filter = "Card Game File (*.cgf) | *.cgf"
        .FilterIndex = 1
        .Filename = ""
        .InitDir = App.Path & "\Save"
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
        
        If .Filename = "" Then Exit Sub
        
        Dim oCard     As Object
        Dim InFile    As Long
        Dim sData     As String
        Dim Filename  As String
    
        InFile = FreeFile
        Open .Filename For Output As InFile
            Write #InFile, GameSignature
            Write #InFile, solGame.Signature
            Write #InFile, solGame.Title
            Write #InFile, solGame.Score
            Write #InFile, solGame.ValHolder1
            Write #InFile, solGame.ValHolder2
            Write #InFile, solGame.ValHolder3
        
            sData = VBScript.Run("GetDataColl")
            If sData <> "" Then Write #InFile, sData
            sData = VBScript.Run("GetTempColl")
            If sData <> "" Then Write #InFile, sData
            sData = VBScript.Run("GetStockColl")
            If sData <> "" Then Write #InFile, sData
            sData = VBScript.Run("GetWasteColl")
            If sData <> "" Then Write #InFile, sData
            
            For Each oCard In Me.Controls
                If oCard.Tag = "PlayingCard" Then
                    sData = oCard.Name & "/$" & _
                            oCard.Index & "/$" & _
                            oCard.Data & "/$" & _
                            oCard.Face & "/$" & _
                            oCard.Left & "/$" & _
                            oCard.Top & "/$" & _
                            oCard.Rank & "/$" & _
                            oCard.Suit & "/$" & _
                            oCard.Selected & "/$" & _
                            oCard.Visible
                    Write #InFile, sData
                End If
            Next oCard
            
        Close InFile
        Exit Sub
    End With
    
SaveErr:
    If InFile <> 0 Then Close InFile
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Error"
End Sub

Private Sub mnuGameSelect_Click()
    On Error GoTo OpenErr
    
    With dlgGame
        .Filter = "Card Game File (*.sol) | *.sol; |All Files (*.*) | *.*"
        .FilterIndex = 1
        .Filename = ""
        .InitDir = App.Path & "\Game"
        .ShowOpen
        
        If .Filename <> "" Then
            picLogo.Visible = False
            Set picLogo.Picture = Nothing
            Set picTable.Picture = Nothing
            picTable.Cls
            
            tmrVictory.Enabled = False
            picTable.ScaleMode = vbTwips
            
            Call InitSolGame
            Call VBScript.AddCode(OpenScript(.Filename))
            Call VBScript.Run("InitGame")
            If solGame.Signature <> "" Then
                If GSI.BkMode = 0 Then Call SetBkBmp
            End If
            Me.Caption = solGame.Title
            ResizeCard Me, "PlayingCard", solGame.CardSizeOp
            
            Call SetSolGame
        End If
    End With
    Exit Sub

OpenErr:
    MsgBox "Unrecognized file format", vbCritical, "Error"
End Sub

Private Sub mnuGameRestart_Click()
    IsRestart = True
    Call SetSolGame
End Sub

Private Sub InitSolGame()
    On Error GoTo Errhandler
    
    Set solGame = New cGame
    
    Call VBScript.Reset
    Call VBScript.AddObject("crdPlayingCard", crdPlayingCard)
    Call VBScript.AddObject("crdStock", crdStock)
    Call VBScript.AddObject("crdTemp", crdTemp)
    Call VBScript.AddObject("crdWaste", crdWaste)
    Call VBScript.AddObject("frmMain", Me)
    Call VBScript.AddObject("picTable", picTable)
    Call VBScript.AddObject("solGame", solGame, True)
    Call VBScript.AddObject("Trigo", Trigo, True)
    Call VBScript.AddObject("imgStock", imgStock)
    Call VBScript.AddObject("imgWaste", imgWaste)
    Call VBScript.AddObject("shpStock", shpStock)
    Call VBScript.AddObject("shpWaste", shpWaste)
    Call VBScript.AddObject("shpCircle", shpCircle)
    Call VBScript.AddCode(OpenScript(App.Path & "\gamescript.ajb"))
    Exit Sub
    
Errhandler:
    MsgBox Err.Description, vbCritical Or vbOKOnly, "Error"
    End
End Sub

Private Sub EnabledControl(bVal As Boolean)
    mnuGameSave.Enabled = bVal
    tblToolbar.Buttons("Save").Enabled = bVal
    
    mnuGameRestart.Enabled = True
    mnuGameOpen.Enabled = True
    mnuHelpHowToPlay.Enabled = True
    tblToolbar.Buttons("ReStart").Enabled = True
    tblToolbar.Buttons("Open").Enabled = True
    tblToolbar.Buttons("Rules").Enabled = True
        
    Dim i As Integer
        
    If solGame.AllowResize Then
        For i = mnuOptionsCardSizeOp.LBound To mnuOptionsCardSizeOp.UBound
            mnuOptionsCardSizeOp(i).Enabled = bVal
        Next i
    End If
End Sub

Public Sub SetSolGame()
    Dim i     As Integer
    Dim Temp  As New Collection
    Dim oCard As Object
    Dim lret  As Long

    sbStatusBar.Panels(1).Text = "Please Wait..."
    tmrVictory.Enabled = False
    
    picTable.Cls
    picTable.ScaleMode = vbTwips
    
    If IsRestart Then
        Call VBScript.Run("CloseIntro")
        IsRestart = False
    End If
    
    For Each oCard In crdStock
        oCard.Visible = False
    Next oCard
    For Each oCard In crdWaste
        oCard.Visible = False
    Next oCard
    For Each oCard In crdTemp
        oCard.Visible = False
    Next oCard
    For Each oCard In crdPlayingCard
        oCard.Visible = False
    Next
    
    imgStock.Visible = False
    imgWaste.Visible = False
    shpStock.Visible = False
    shpWaste.Visible = False
    shpCircle(0).Visible = False
      
    solGame.Score = 0
    solGame.ValHolder1 = 0
    solGame.ValHolder2 = 0
    solGame.ValHolder3 = 0

    Set solGame.DataColl = Nothing
    Set solGame.TempColl = Nothing
    Set solGame.StockColl = Nothing
    Set solGame.WasteColl = Nothing
      
    Set Temp = Shuffle(TotalCards)
    
    For i = 1 To Temp.Count
        crdPlayingCard(i).Update = False
        crdPlayingCard(i).Data = ""
        crdPlayingCard(i).Enabled = True
        crdPlayingCard(i).Face = crd_Down
        crdPlayingCard(i).Rank = Temp(i) Mod 13
        crdPlayingCard(i).Suit = Temp(i) Mod 4
        crdPlayingCard(i).Selected = False
        crdPlayingCard(i).ZOrder 0
        crdPlayingCard(i).Update = True
        crdPlayingCard(i).Refresh
    Next i
    
    For i = crdStock.LBound To crdStock.UBound
        crdStock(i).Face = crd_Down
    Next i
    
    For i = crdWaste.LBound To crdWaste.UBound
        crdWaste(i).Face = crd_Up
    Next i
    
    For Each oCard In Me.Controls
        If TypeName(oCard) = "Card" Then
            oCard.Tag = "PlayingCard"
        End If
    Next oCard
    
    Call VBScript.Run("ResetGame")
    Call VBScript.Run("OpenIntro")
    Call ShowScore
    Call EnabledControl(True)
    Me.mnuGameRestart.Enabled = True
End Sub

Public Sub ShowScore()
    sbStatusBar.Panels(1).Text = "Score: " & solGame.Score
End Sub

Private Sub mnuHelpAbout_Click()
    Dim OldTimer As Boolean
    
    OldTimer = tmrVictory.Enabled
    tmrVictory.Enabled = False
    frmAbout.Show vbModal, Me
    tmrVictory.Enabled = OldTimer
End Sub

Private Sub mnuHelpHowToPlay_Click()
    If solGame.Signature <> "" Then
        WinHelp 0, App.HelpFile, HELP_CONTEXT, solGame.HelpID
    End If
End Sub

Private Sub mnuOptionsCard_Click()
    Dim bOldTimer As Boolean
    
    bOldTimer = tmrVictory.Enabled
    tmrVictory.Enabled = False
    frmCardOp.Show vbModal, Me
    tmrVictory.Enabled = bOldTimer
End Sub

Private Sub mnuOptionsSettings_Click()
    Dim bOldTimer As Boolean
    
    bOldTimer = tmrVictory.Enabled
    tmrVictory.Enabled = False
    frmSettings.Show vbModal, Me
    tmrVictory.Enabled = bOldTimer
End Sub

Private Sub mnuOptionsCardSizeOp_Click(Index As Integer)
    On Error Resume Next
    
    Me.Enabled = False
    
    Dim i As Integer, OldScaleMode
    
    OldScaleMode = picTable.ScaleMode
    picTable.ScaleMode = vbTwips
    For i = mnuOptionsCardSizeOp.LBound To mnuOptionsCardSizeOp.UBound
        mnuOptionsCardSizeOp(i).Checked = False
    Next i
    
    Select Case Index
    Case Is = 0 ' Automatic
        If solGame.Signature <> "" Then
            ResizeCard Me, "PlayingCard", solGame.CardSizeOp
        End If
    Case Is = 1 ' Small
        ResizeCard Me, "PlayingCard", cs_Small
    Case Is = 2 ' Standard
        ResizeCard Me, "PlayingCard", cs_Standard
    Case Is = 3 ' Large
        ResizeCard Me, "PlayingCard", cs_Large
    End Select
    
    mnuOptionsCardSizeOp(Index).Checked = True
    
    If solGame.Signature <> "" Then Call VBScript.Run("AlignCards")
    picTable.ScaleMode = OldScaleMode
    Me.Enabled = True
End Sub

Private Sub PlayVictoryAni()
    Dim i       As Integer
    Dim oCard   As Object
    Dim Counter As Integer
    
    imgStock.Visible = False
    imgWaste.Visible = False
    shpWaste.Visible = False
    shpStock.Visible = False
    
    For i = shpCircle.LBound To shpCircle.UBound
        shpCircle(i).Visible = False
    Next i
    
    For Each oCard In Me.Controls
        If TypeName(oCard) = "Card" Then
            oCard.Tag = ""
        End If
    Next oCard
    
    Counter = 0
    For Each oCard In Me.Controls
        If TypeName(oCard) = "Card" Then
            If oCard.Visible Then
                If Counter <= 20 Then
                    oCard.Enabled = False
                    oCard.Tag = "PlayingCard"
                    Counter = Counter + 1 ' count card
                Else
                    If oCard.Visible Then
                        oCard.Visible = False
                    End If
                End If
            End If
        End If
    Next oCard
    
    If GSI.VicAniMode = 0 Then
        curVicAni = CInt(Rnd * 5)
    Else
        If GSI.VicAniSel <> curVicAni Then
            curVicAni = GSI.VicAniSel
        End If
    End If
    
    picTable.ScaleMode = vbPixels
    VicAni.Reset = True
    tmrVictory.Enabled = True
End Sub

Private Sub tblToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
    Case Is = "Select"
        Call mnuGameSelect_Click
    Case Is = "ReStart"
        Call mnuGameRestart_Click
    Case Is = "Open"
        Call mnuGameOpen_Click
    Case Is = "Save"
        mnuGameSave_Click
    Case Is = "Card"
        mnuOptionsCard_Click
    Case Is = "Settings"
        mnuOptionsSettings_Click
    Case Is = "Rules"
        mnuHelpHowToPlay_Click
    Case Is = "About"
        mnuHelpAbout_Click
    End Select
End Sub

Private Sub tmrVictory_Timer()
    If Me.WindowState <> vbMinimized Then
        If GSI.DistX <> VicAni.DistX Then VicAni.DistX = GSI.DistX
        If GSI.DistY <> VicAni.DistY Then VicAni.DistY = GSI.DistY
        
        Select Case curVicAni
        Case Is = 0 ' Bounce
            Call VicAni.Bounce(Me, picTable, "PlayingCard")
        Case Is = 1 ' Bounce (Scatter)
            Call VicAni.BounceScatter(Me, picTable, GSI.Trail, "PlayingCard")
        Case Is = 2 ' Bounce (Trail)
            Call VicAni.BounceTrail(Me, picTable, crdPlayingCard(1), GSI.Trail, "PlayingCard")
        Case Is = 3 ' Spin
            Call VicAni.Spin(Me, picTable, "PlayingCard")
        Case Is = 4 ' Spin (Trail)
            Call VicAni.SpinTrail(Me, picTable, crdPlayingCard(1), GSI.Trail, "PlayingCard")
        Case Is = 5 ' Wave
            Call VicAni.Wave(Me, picTable, "PlayingCard", GSI.WaveExpr, _
                             crdPlayingCard(0).Width, crdPlayingCard(0).Height, GSI.Clip)
        End Select
    End If
End Sub

Public Sub SetBkBmp()
    Dim rcRect As RECT

    On Error GoTo Errhandler
    
    picTable.Cls
    Set picTable.Picture = Nothing
    
    If Dir$(GSI.BkFile) <> "" Then
        GetClientRect picTable.hwnd, rcRect
        TileBmp picTable.hdc, LoadPicture(GSI.BkFile), _
                rcRect.Right, rcRect.Bottom
        RefreshWindow picTable.hwnd
        Call InitLogo
        Set picTable.Picture = picTable.Image
    Else
        Set picTable.Picture = Nothing
    End If
    Exit Sub
    
Errhandler:
    MsgBox Err.Description, vbOKOnly Or vbInformation, "Error"
End Sub

Private Sub RestoreCardSettings()
    With GSI.Card
        Dim oCard As Object
        
        For Each oCard In frmMain.Controls
            If TypeName(oCard) = "Card" Then
                oCard.Update = False
                oCard.Deck = .Deck
                oCard.DeckBackground = .DeckBackground
                oCard.DeckMaskStyle = .DeckMaskStyle
                oCard.Effect = .Effect
                oCard.FontBold = .FontBold
                oCard.FontItalic = .FontItalic
                oCard.FontName = .FontName
                oCard.FontSize = .FontSize
                oCard.FontTransparent = .FontTransparent
                oCard.Forecolor = .Forecolor
                oCard.FramePerMoveX = .FramePerMoveX
                oCard.FramePerMoveY = .FramePerMoveY
                oCard.Speed = .Speed
                oCard.Text = .Text
                oCard.Update = True
                oCard.Refresh
                
                If .Deck = Crd_Deck_Customize Then
                    If Dir$(.DeckPicture) <> "" Then
                        Set oCard.DeckPicture = LoadPicture(.DeckPicture)
                    Else
                        oCard.Deck = crd_Deck_1
                    End If
                End If
                
                If .DeckMaskStyle = crd_Customize Then
                    If Dir$(.DeckMaskPicture) <> "" Then
                        Set oCard.DeckMaskPicture = LoadPicture(.DeckMaskPicture)
                    Else
                        .DeckMaskStyle = crd_None
                    End If
                End If
                
                Call SetTypeEffect(oCard, .TypEffect)
            End If
        Next
    End With
    
    If GSI.BkMode = 1 Then picTable.BackColor = GSI.BkColor
    VicAni.CardStat = 2
    curVicAni = GSI.VicAniSel
    VicAni.DistX = GSI.DistX
    VicAni.DistY = GSI.DistY
End Sub

Private Sub SaveCardSettings()
    With GSI.Card
        .Deck = crdPlayingCard(0).Deck
        .DeckBackground = crdPlayingCard(0).DeckBackground
        .DeckMaskStyle = crdPlayingCard(0).DeckMaskStyle
        .Effect = crdPlayingCard(0).Effect
        .FontBold = crdPlayingCard(0).FontBold
        .FontItalic = crdPlayingCard(0).FontItalic
        .FontName = crdPlayingCard(0).FontName
        .FontSize = crdPlayingCard(0).FontSize
        .FontTransparent = crdPlayingCard(0).FontTransparent
        .Forecolor = crdPlayingCard(0).Forecolor
        .FramePerMoveX = crdPlayingCard(0).FramePerMoveX
        .FramePerMoveY = crdPlayingCard(0).FramePerMoveY
        .Speed = crdPlayingCard(0).Speed
        .Text = crdPlayingCard(0).Text
        .TypEffect = GetTypeEffect(crdPlayingCard(0))
    End With
End Sub

Private Sub InitLogo()
    If Dir$(PathLogo) <> "" Then
        Dim BM As BITMAP
        If GetObjectAPI(LoadPicture(PathLogo), Len(BM), BM) Then
            picLogo.Width = BM.bmWidth
            picLogo.Height = BM.bmHeight
            picLogo.BackColor = picTable.BackColor
            Call Logo(picLogo, picTemp)
            Set picTemp.Picture = Nothing
        End If
    End If
End Sub

Private Sub RepaintLogo()
    If picLogo.Picture Then
        picTable.Cls
        picTable.PaintPicture picLogo.Picture, _
            (picTable.ScaleWidth - picLogo.ScaleWidth) / 2, _
            (picTable.ScaleHeight - picLogo.ScaleHeight) / 2
    End If
End Sub


