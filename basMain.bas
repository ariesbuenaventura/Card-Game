Attribute VB_Name = "basMain"
Option Explicit

Public Enum BackgroundConstants
    bgNone
    bgCenter
    bgStretch
    bgTile
End Enum

Public PathLogo As String

Public Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    
    PathLogo = App.Path & "\logo.bmp"
    VBScript.Language = "VBScript"
    VBScript.UseSafeSubset = True
    VBScript.Timeout = NoTimeout
    Load frmMain
    frmMain.Show
    
    Unload frmSplash
End Sub

Public Sub CreateObject(objType As Object, ByVal Count As Integer)
    Dim i As Integer, j(3) As Integer
    
    j(0) = 11
    j(1) = 0
    
    For i = 1 To Count
        If objType.Count > 0 Then
            Load objType(objType.Count)
            objType(objType.Count - 1).Left = _
                -objType(objType.Count - 1).Width
            objType(objType.Count - 1).Top = _
                -objType(objType.Count - 1).Height
            objType(objType.Count - 1).ZOrder 0
        End If
    Next i
End Sub

Public Sub DestroyObject(objType As Object)
    Dim oCard As Object
    
    For Each oCard In objType
        If oCard.Index <> 0 Then
            Unload oCard
        End If
    Next oCard
End Sub

Public Function OpenScript(Filename As String) As String
    Dim sTemp As String, lfn As Long
    
    lfn = FreeFile
    Open Filename For Input As #lfn
        sTemp = Input(LOF(lfn), #lfn)
    Close #lfn
    
    OpenScript = sTemp
End Function

Public Function Sum(nStart, nEnd)
    Dim i As Integer
    Dim S As Integer
    
    S = 0
    For i = nStart To nEnd
        S = S + i
    Next i
    
    Sum = S
End Function

Public Function Shuffle(MxVal As Long) As Collection
    Dim RetVal  As Long
    Dim step    As Integer
    Dim cntr    As Integer
    Dim IsExist As Boolean
    Dim Data    As New Collection
    
    cntr = 0
    Randomize
    
    Do While cntr < MxVal
        IsExist = False
        RetVal = Int((MxVal * Rnd) + 1)
        For step = 1 To cntr
            If Data(step) = RetVal Then
                IsExist = True
                Exit For
            End If
        Next step
        If Not IsExist Then
            Data.Add RetVal
            cntr = cntr + 1
        End If
    Loop
    
    Set Shuffle = Data
End Function

Public Sub RefreshWindow(lhWnd As Long)
    Dim rcRect As RECT
    
    GetClientRect lhWnd, rcRect
    InvalidateRect lhWnd, rcRect, False
End Sub

Public Sub DrawLine(lhDstDC As Long, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, _
                    nWidth As Long, nColor As Long)
                    
    Dim lhBrush As Long
    Dim pts     As POINTAPI

    lhBrush = CreatePen(PS_SOLID, nWidth, nColor)
    SelectObject lhDstDC, lhBrush
    MoveToEx lhDstDC, X1, Y1, pts
    LineTo lhDstDC, X2, Y2
    DeleteObject lhBrush
End Sub

Public Sub DrawCircle(lhDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Radius As Long, _
                      ByVal Forecolor As Long, ByVal FillColor As Long, ByVal PenWidth As Integer)
                        
    Dim hPen      As Long
    Dim hOldPen   As Long
    Dim hBrush    As Long
    Dim hOldBrush As Long
    
    hPen = CreatePen(PS_SOLID, PenWidth, Forecolor)
    hBrush = CreateSolidBrush(FillColor)
    hOldPen = SelectObject(lhDC, hPen)
    hOldBrush = SelectObject(lhDC, hBrush)
    
    Ellipse lhDC, X - Radius, Y - Radius, X + Radius, Y + Radius
    
    SelectObject lhDC, hOldPen
    SelectObject lhDC, hOldBrush
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Public Function RemovSpace(Data As String) As String
    If Data = "" Then Exit Function
    
    Dim i      As Integer
    Dim Buffer As String
    
    Buffer = ""
    For i = 1 To Len(Data)
        If Mid$(Data, i, 1) <> " " Then
            Buffer = Buffer & Mid$(Data, i, 1)
        End If
    Next
    
    RemovSpace = Buffer
End Function

Public Function Rads(deg As Single) As Single
    Rads = deg * Pi / 180 ' convert the angle into radians
End Function

Public Function Pi() As Single
    Pi = 4 * Atn(1)
End Function

Public Sub ShowCards(thisCard As Object, _
                      bVal As Boolean, MatchTag)
    
    Dim oCard As Object
    
    For Each oCard In thisCard
        If oCard.Tag = MatchTag Then
            oCard.Visible = bVal
        End If
    Next oCard
End Sub

Public Sub TileBmp(lhDstDC As Long, Image As IPicture, nWidth As Long, nHeight As Long)
    On Error GoTo ErrHandler
    
    If Image.handle = 0 Then Exit Sub
    If (nWidth <= 0) Or (nHeight <= 0) Then Exit Sub
    
    Dim rcRect  As RECT
    Dim lhBrush As Long
    Dim lhBmp   As Long
    Dim BMP     As BITMAP
    
    GetObjectAPI Image.handle, Len(BMP), BMP
    lhBmp = CopyImage(Image.handle, ByVal 0&, BMP.bmWidth, BMP.bmHeight, ByVal 0&)
    SetRect rcRect, 0, 0, nWidth, nHeight
    
    lhBrush = CreatePatternBrush(lhBmp)
    FillRect lhDstDC, rcRect, lhBrush
    DeleteObject lhBrush
    
    DeleteObject lhBmp
    Exit Sub
    
ErrHandler:
End Sub

Public Sub DefaultSettings()
    GSI.BkColor = &H8000&
    GSI.BkFile = ""
    GSI.BkMode = 1
    GSI.Clip = True
    GSI.DistX = 4
    GSI.DistY = 4
    GSI.Speed = 4
    GSI.Trail = True
    GSI.VicAniMode = 0
    GSI.VicAniSel = 0
    GSI.WaveExpr = ""
    GSI.Card.curDeck = 1
    GSI.Card.curMask = 1
    GSI.Card.Deck = 0
    GSI.Card.DeckBackground = 0
    GSI.Card.DeckMaskPicture = ""
    GSI.Card.DeckMaskStyle = 0
    GSI.Card.DeckPicture = ""
    GSI.Card.Effect = 0
    GSI.Card.FontBold = False
    GSI.Card.FontItalic = False
    GSI.Card.FontName = "Arial Black"
    GSI.Card.FontSize = 8
    GSI.Card.FontTransparent = False
    GSI.Card.Forecolor = 0
    GSI.Card.FramePerMoveX = 1
    GSI.Card.FramePerMoveY = 1
    GSI.Card.Speed = 1
    GSI.Card.Text = ""
    GSI.Card.TypEffect = 0
    
    frmMain.crdPlayingCard(0).Effect = crd_Effect_None
End Sub

Public Sub OpenSettings()
    Dim Filename As String
    
    On Error GoTo OpenErr
    
    Filename = App.Path & "\data\setting.dat"
    
    If Dir$(Filename) <> "" Then
        Dim InFile As Integer
        Dim Buffer As String
        
        InFile = FreeFile
        Open Filename For Input As InFile
            Input #InFile, Buffer
            If CStr(Buffer) = GameSignature Then
                Input #InFile, GSI.BkColor
                Input #InFile, GSI.BkFile
                Input #InFile, GSI.BkMode
                Input #InFile, GSI.Clip
                Input #InFile, GSI.DistX
                Input #InFile, GSI.DistY
                Input #InFile, GSI.Speed
                Input #InFile, GSI.Trail
                Input #InFile, GSI.VicAniMode
                Input #InFile, GSI.VicAniSel
                Input #InFile, GSI.WaveExpr
                
                With GSI.Card
                    Input #InFile, .curDeck
                    Input #InFile, .curMask
                    Input #InFile, .Deck
                    Input #InFile, .DeckBackground
                    Input #InFile, .DeckPicture
                    Input #InFile, .DeckMaskStyle
                    Input #InFile, .DeckMaskPicture
                    Input #InFile, .Effect
                    Input #InFile, .FontBold
                    Input #InFile, .FontItalic
                    Input #InFile, .FontName
                    Input #InFile, .FontSize
                    Input #InFile, .FontTransparent
                    Input #InFile, .Forecolor
                    Input #InFile, .FramePerMoveX
                    Input #InFile, .FramePerMoveY
                    Input #InFile, .Speed
                    Input #InFile, .Text
                    Input #InFile, .TypEffect
                End With
            Else
                MsgBox "File format error!", vbOKOnly Or vbInformation, "Error"
            End If
        Close InFile
    Else
        Call DefaultSettings
    End If
    Exit Sub

OpenErr:
    If InFile <> 0 Then Close InFile
    MsgBox Err.Description, vbOKOnly & vbCritical, "Error"
End Sub

Public Sub SaveSettings()
    Dim Filename As String
    
    On Error GoTo SaveErr
    
    Filename = App.Path & "\data\setting.dat"
    
    Dim i As Integer, InFile As Integer
    
    InFile = FreeFile
    Open Filename For Output As InFile
        Write #InFile, GameSignature
        Write #InFile, GSI.BkColor
        Write #InFile, GSI.BkFile
        Write #InFile, GSI.BkMode
        Write #InFile, GSI.Clip
        Write #InFile, GSI.DistX
        Write #InFile, GSI.DistY
        Write #InFile, GSI.Speed
        Write #InFile, GSI.Trail
        Write #InFile, GSI.VicAniMode
        Write #InFile, GSI.VicAniSel
        Write #InFile, GSI.WaveExpr
        
        With GSI.Card
            Write #InFile, .curDeck
            Write #InFile, .curMask
            Write #InFile, .Deck
            Write #InFile, .DeckBackground
            Write #InFile, .DeckPicture
            Write #InFile, .DeckMaskStyle
            Write #InFile, .DeckMaskPicture
            Write #InFile, .Effect
            Write #InFile, .FontBold
            Write #InFile, .FontItalic
            Write #InFile, .FontName
            Write #InFile, .FontSize
            Write #InFile, .FontTransparent
            Write #InFile, .Forecolor
            Write #InFile, .FramePerMoveX
            Write #InFile, .FramePerMoveY
            Write #InFile, .Speed
            Write #InFile, .Text
            Write #InFile, .TypEffect
        End With
    Close InFile
    Exit Sub
SaveErr:
    If InFile <> 0 Then Close InFile
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Error"
End Sub

Public Sub BeginPlaySound(ByVal ResourceId As Integer, Optional ByVal SoundLoop As Boolean = False)
    Dim SoundBuffer() As Byte
    
    SoundBuffer = LoadResData(ResourceId, "SOUND")
    
    If SoundLoop Then
        sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_LOOP
    Else
        sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
    End If
End Sub

Public Sub EndPlaySound()
    sndPlaySound ByVal vbNullString, 0&
End Sub

Public Function TransBMP(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal TransColor As Long) As Long
    If DstW = 0 Or DstH = 0 Then Exit Function
    
    Dim B As Long, h As Long, F As Long, i As Long
    Dim TmpDC As Long, tmpBMP As Long, TmpObj As Long
    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
    Dim Data1() As Long, Data2() As Long
    Dim Info As BITMAPINFO
    
    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    tmpBMP = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, tmpBMP)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data1(DstW * DstH * 4 - 1)
    ReDim Data2(DstW * DstH * 4 - 1)
    Info.bmiHeader.biSize = Len(Info.bmiHeader)
    Info.bmiHeader.biWidth = DstW
    Info.bmiHeader.biHeight = DstH
    Info.bmiHeader.biPlanes = 1
    Info.bmiHeader.biBitCount = 32
    Info.bmiHeader.biCompression = 0

    Call BitBlt(TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy)
    Call BitBlt(Sr2DC, 0, 0, DstW, DstH, SrcDC, SrcX, SrcY, vbSrcCopy)
    Call GetDIBits(TmpDC, tmpBMP, 0, DstH, Data1(0), Info, 0)
    Call GetDIBits(Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0)
    
    For h = 0 To DstH - 1
        F = h * DstW
        For B = 0 To DstW - 1
            i = F + B
            If (Data2(i) And &HFFFFFF) = TransColor Then
            Else
                Data1(i) = Data2(i)
            End If
        Next B
    Next h

    Call SetDIBitsToDevice(DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0)

    Erase Data1
    Erase Data2
    Call DeleteObject(SelectObject(TmpDC, TmpObj))
    Call DeleteObject(SelectObject(Sr2DC, Sr2Obj))
    Call DeleteDC(TmpDC)
    Call DeleteDC(Sr2DC)
End Function

Public Sub Logo(DstPicBox As PictureBox, SrcPicBox As PictureBox)
    Dim BM As BITMAP
    
    If Dir$(PathLogo) <> "" Then
        With SrcPicBox
            Set .Picture = LoadPicture(PathLogo)
            
            If .Picture Then
                Dim OldAutoRedraw As Integer, OldScaleMode As Integer
                OldAutoRedraw = DstPicBox.AutoRedraw
                OldScaleMode = DstPicBox.ScaleMode
                DstPicBox.AutoRedraw = True
                DstPicBox.ScaleMode = vbPixels
                Call GetObjectAPI(.Picture, Len(BM), BM)
                Dim xmid As Integer, ymid As Integer
                xmid = (DstPicBox.ScaleWidth - BM.bmWidth) / 2
                ymid = (DstPicBox.ScaleHeight - BM.bmHeight) / 2
                Call TransBMP(DstPicBox.hdc, xmid, ymid, _
                    .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HFFFFFF)
                Set DstPicBox.Picture = DstPicBox.Image
                DstPicBox.AutoRedraw = OldAutoRedraw
                DstPicBox.ScaleMode = OldScaleMode
            End If
        End With
    End If
End Sub


