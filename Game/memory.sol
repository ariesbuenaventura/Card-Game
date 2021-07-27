Sub InitGame()
     Signature   = "776977798289" ' use to identify the game
     Title       = "Memory"       ' Game title
     CardSizeOp  = 0              ' Small size
     AllowResize = True 	    ' Resize card
     HelpID      = 100            ' Help ID  

     Call ResetGame
End Sub

Sub ResetGame()
     picTable.ScaleMode = vbTwips
     For Each oCard In crdPlayingCard
          If oCard.Index <> 0 Then
               oCard.ArrowKeyFocus = False
          End If
     Next
End Sub

Sub AlignCards()
     cW = crdPlayingCard(0).Width
     cH = crdPlayingCard(0).Height
            
     sx = cW + 60
     sy = cH + 480

     px = (picTable.ScaleWidth - sx * 13) / 2
     py = (picTable.ScaleHeight - sy * 4) / 2
     
     For row = 0 To 3
          For col = 0 To 12
	         crdPlayingCard((col + 1) + row * 13).Move px + sx * col, py + sy * row
               If Speed = 0 Then crdPlayingCard((col + 1) + row * 13).Visible = True
	    Next
     Next      
End Sub

Sub OpenIntro()
     Call LockForm(True)

     picTable.ScaleMode = vbTwips
     If Speed > 0 Then
          nWidth = picTable.ScaleWidth: nHeight = picTable.ScaleHeight
          cW = crdPlayingCard(0).Width: cH = crdPlayingCard(0).Height

          sx = cW + 60: sy = cH + 480
          nrow = 4: ncol = 13

          px = (nWidth - sx * ncol) / 2
          py = (nHeight - sy * nrow) / 2

          crdPlayingCard(1).Move -cW, py
          crdPlayingCard(26).Move nWidth + cW, py + sy * 1
          crdPlayingCard(27).Move -cW, py + sy * 2
          crdPlayingCard(52).Move nWidth + cW + sx * 12, py + sy * 3

          For i = 1 To ncol
               crdPlayingCard(i).Visible = True
               crdPlayingCard(ncol * 2 - i + 1).Visible = True
               crdPlayingCard(i + ncol * 2).Visible = True
               crdPlayingCard(ncol * 4 - i + 1).Visible = True

               Do While crdPlayingCard(i).Left < px + sx * (i - 1)
                    crdPlayingCard(i).Move crdPlayingCard(i).Left + gm_TwipsPerPixelX * Speed, py
                    crdPlayingCard(i).ZOrder 0
                    crdPlayingCard(ncol * 2 - i + 1).Move nWidth - cW - crdPlayingCard(i).Left - _
                                                          gm_TwipsPerPixelX * Speed, py + sy
                    crdPlayingCard(ncol * 2 - i + 1).ZOrder 0
                    crdPlayingCard(i + ncol * 2).Move crdPlayingCard(i).Left, py + sy * 2
                    crdPlayingCard(i + ncol * 2).ZOrder 0	 
                    crdPlayingCard(ncol * 4 - i + 1).Move nWidth - cW - crdPlayingCard(i).Left - _
                                                          gm_TwipsPerPixelX * Speed, py + sy * 3
                    crdPlayingCard(ncol * 4 - i + 1).ZOrder 0
                    Call gm_DoEvents()
	         Loop
		
               ' Make sure that all cards are not misaligned.		
	         crdPlayingCard(i).Move px + sx * (i - 1), py
    	         crdPlayingCard(ncol * 2 - i + 1).Move px + sx * (13 - i), py + sy
	         crdPlayingCard(i + ncol * 2).Move crdPlayingCard(i).Left, py + sy * 2
	         crdPlayingCard(ncol * 4 - i + 1).Move px + sx * (13 - i), py + sy * 3

	         For j = i To 13
	              crdPlayingCard(j).Move px + sx * (i - 1), py
	         Next
          Next
     Else
          Call AlignCards()
     End If

     Do While IsCardAniOnEx(frmMain)
          Call gm_DoEvents
     Loop
         
     Call LockForm(False)
     crdPlayingCard(1).SetFocus	
End Sub

Sub CloseIntro()
     Call LockForm(True)
      
     picTable.ScaleMode = vbTwips 
     If Speed > 0 Then
          nWidth = picTable.ScaleWidth: nHeight = picTable.ScaleHeight
          cW = crdPlayingCard(0).Width: cH = crdPlayingCard(0).Height

          sx = cW + 60: sy = cH + 480
          nrow = 4: ncol = 13

          px = (nWidth - sx * ncol) / 2
          py = (nHeight - sy * nrow) / 2

          curIndex = 1

          crdPlayingCard(13).ZOrder 0
          crdPlayingCard(14).ZOrder 0
          crdPlayingCard(39).ZOrder 0
          crdPlayingCard(40).ZOrder 0
 
          Do While crdPlayingCard(13).Left > -cW - cW * 0.1 
               crdPlayingCard(13).Move crdPlayingCard(13).Left - gm_TwipsPerPixelX * Speed, _
                                       crdPlayingCard(13).Top        
	         crdPlayingCard(14).Move crdPlayingCard(14).Left + gm_TwipsPerPixelX * Speed, _
                                       crdPlayingCard(14).Top
	         crdPlayingCard(39).Move crdPlayingCard(39).Left - gm_TwipsPerPixelX * Speed, _
                                       crdPlayingCard(39).Top
	         crdPlayingCard(40).Move crdPlayingCard(40).Left + gm_TwipsPerPixelX * Speed, _
                                       crdPlayingCard(40).Top
          
               If curIndex < 13 Then 
                    If crdPlayingCard(13).Left < px + sx * (13 - curIndex - 1) Then
                         crdPlayingCard(13 - curIndex).Visible = False
                         crdPlayingCard(14 + curIndex).Visible = False 
                         crdPlayingCard(39 - curIndex).Visible = False
                         crdPlayingCard(40 + curIndex).Visible = False 
                         curIndex = curIndex + 1     
                    End If
               End If
               Call gm_DoEvents()
          Loop
     Else
          Call AlignCards()
     End If

     Do While IsCardAniOnEx(frmMain)
          Call gm_DoEvents
     Loop
     Call LockForm(False)
End Sub

Sub DoKeyDown(Index, KeyCode)	
     ' make sure that all inputs are numeric
     If Not IsNumeric(Index)   Then Exit Sub ' if not numeric then exit sub
     If Not IsNumeric(KeyCode) Then Exit Sub ' if not numeric then exit sub
 
     If IsCardAniOn(crdPlayingCard) Then Exit Sub

     newFocus = -1

     Select Case KeyCode
     Case &H25 ' Left
          newFocus = Index - 1
          If (newFocus Mod 13) = 0  Then newFocus = Index - (Index Mod 13) + 13
     Case &H27 ' Right
          newFocus = Index + 1
          If (newFocus Mod 13) = 1 Then newFocus = Index - (Index Mod 13) - 13 + 1
     Case &H26 ' Up
          newFocus = Index - 13
          If newFocus < 1  Then newFocus = (13 * 3) + Index
     Case &H28 ' Down
          newFocus = Index + 13
          If newFocus > 52 Then newFocus = Index - (13 * 3) 
     Case &H0D, &H20 ' Enter, Space
          Call gm_DoEvents()
          Call gm_Delay(500) ' wait for 0.5 second  
     End Select
     
     If newFocus <> -1 Then 
          crdPlayingCard(newFocus).SetFocus
     End If
End Sub

Sub DoWasteClick(thisCard)
     ' do nothing	
End Sub

Sub DoStockClick(thisCard)
     ' do nothing
End Sub

Function Process(thisCard)
     Call LockForm(True)

     Process = False

     ' thisCard must be an object (ex. form, textbox, label...)
     If Not IsObject(thisCard) Then 
          Call LockForm(False)
          Exit Function
     Else
          ' check object type
          If TypeName(thisCard) <> "Card" Then
               ' if not a card then exit
               Call LockForm(False)
               Exit Function
          End If 
     End If
	
     If thisCard.Face = 0 Then ' Face Up
          If IsCardAniOn(crdPlayingCard) Then 
               Call LockForm(False)
               Exit Function
          End If 

          If thisCard.Data <> "" Then
               TempColl.Add thisCard
               For Each oCard In crdPlayingCard
                    If oCard.Index <> 0 Then
                         If oCard.Index <> thisCard.Index Then
                              If oCard.Data = thisCard.Data Then
                                   TempColl.Add oCard
                                   Exit For 
                              End If 
                         End If 
                    End If
               Next 
               
               For i = 1 To TempColl.Count
                    If TempColl(i).Effect <> 0 Then
                         TempColl(i).Update = False                    
                         TempColl(i).AutoFlipCard = False
                         TempColl(i).Selected = True
                         TempColl(i).Effect = 8 ' ThreeD
                         TempColl(i).ThreeD = 1 ' Left
                         TempColl(i).PlayAni
                    Else
                         TempColl(i).Selected = True 
                         Call gm_DoEvents ()
                    End If
               Next 

	         Do While IsCardAniOn(crdPlayingCard)
		        Call gm_DoEvents()
               Loop
                
               Call gm_Delay(500) ' wait for 0.5 second
               
               For i = 1 To TempColl.Count
                    If TempColl(i).Effect <> 0 Then
                         TempColl(i).Selected = False   
                         TempColl(i).ThreeD = 2 ' Right
                         TempColl(i).PlayAni
                    Else
                         TempColl(i).Selected = False
                    End If
               Next
               
	         Do While IsCardAniOn(crdPlayingCard)
		        Call gm_DoEvents()
               Loop

               For i = 1 To TempColl.Count
                   If TempColl(i).Effect <> 0 Then
                        TempColl(i).AutoFlipCard = True
                        TempColl(i).Update = True
                        TempColl(i).Refresh
                   End If
               Next 

               Set TempColl = Nothing
               Do While IsCardAniOnEx(frmMain)
                    Call gm_DoEvents
               Loop
               Call LockForm(False)
               Exit Function
          Else
               Call LockForm(False)
               Exit Function
          End If
     End If

     If thisCard.Effect <> 0 Then	
          If IsCardAniOn(crdPlayingCard) Then 
               Call LockForm(False)
	         Exit Function
	    Else
               thisCard.PlayAni
	         Do While IsCardAniOn(crdPlayingCard)
		        Call gm_DoEvents()
               Loop 
	    End If
     Else
          If thisCard.Face = 0 Then
		   thisCard.Face = 1 ' Face Down
	    Else	
		   thisCard.Face = 0 ' Face Up
	    End If
     End If	

     Call PlaySound(101)
     DataColl.Add thisCard

     If DataColl.Count = 2 then 
          If (DataColl(1).Rank = DataColl(2).Rank) Then
		   DataColl(1).Data = ValHolder1 
               DataColl(2).Data = ValHolder1
               ValHolder1 = ValHolder1 + 1
               
               If GetTotalCardsLeft("") = 0 Then
                    Process = True
               End If

               Score = Score + 50
          Else
               Call gm_DoEvents()
	         Call gm_Delay(1000) ' wait for 1 second before closing the card
               FlipCard crdPlayingCard, 1, ""		
               ' get 30% of total cards open then subtract from the score
               Score = Score - Int(ValHolder2 * 0.3)
               ' restore previous effect
          End If 

          Set DataColl = Nothing

          Do While IsCardAniOn(crdPlayingCard)
               Call gm_DoEvents()
          Loop
     End If 
     ValHolder2 = ValHolder2 + 1 ' total cards open

     Do While IsCardAniOnEx(frmMain)
          Call gm_DoEvents
     Loop
     
     Call LockForm(False)
End Function

Function GetDataColl()
     sData = ""
     For Each oCard In DataColl
          sData = sData & oCard.Name & "/$" & oCard.Index & "/$"
     Next

     If Right(sData, 2) = "/$" Then sData = Left(sData,Len(sData) - 2)
     If sData = "" Then sData = "/$"

     GetDataColl = sData
End Function

Function GetTempColl
    GetTempColl = "/$" 
End Function

Function GetStockColl
     GetStockColl = "/$"
End Function

Function GetWasteColl
     GetWasteColl = "/$"
End Function

Sub SetDataColl(sData)
     Set DataColl = Nothing

     If CStr(sData) = "/$" Then Exit Sub

     arrData = Split(sData, "/$")
     For i = LBound(arrData) To UBound(arrData) - 1
          For Each oCard In frmMain.Controls
               If oCard.Tag = "PlayingCard" Then
                   If (cStr(arrData(i)) = oCard.Name) And _
                      (cInt(arrData(i + 1)) = oCard.Index) Then
                           DataColl.Add oCard
                   End If 
               End If
          Next
     Next
End Sub

Sub SetTempColl(sData)
     Set TempColl = Nothing
End Sub

Sub SetStockColl(sData)
     Set StockColl = Nothing
End Sub	

Sub SetWasteColl(sData)
     Set WasteColl = Nothing
End Sub

Sub SetValHolder
     ValHolder1 = CInt(ValHolder1)
     ValHolder2 = CInt(ValHolder2)
End Sub

Sub UpdateControls
     ' do nothing
End Sub