Sub InitGame()
     Signature   = "808982657773" ' use to identify the game
     Title       = "Pyramid"      ' Game title
     CardSizeOp  = 0              ' Default size (0-Small,1-Standard,2-Large)
     AllowResize = True           ' resize the card
     HelpID      = 101            ' Help ID

     Call ResetGame
End Sub

Sub ResetGame
     picTable.ScaleMode = vbTwips

     For Each oCard In crdStock
          oCard.Visible = True
     Next

     For Each oCard In crdPlayingCard
          If (oCard.Index >= 29) And (oCard.Index <= 52) Then
               oCard.Visible = False 
		   StockColl.Add oCard.Index
	    Else
               If oCard.Index <> 0 Then
                    oCard.Visible = True
               End If
          End If 
     Next

     imgStock.Visible     = True
     imgWaste.Visible     = True
     shpStock.Visible     = True
     shpWaste.Visible     = True
     shpCircle(0).Visible = True   
End Sub

Sub AlignCards()
      nWidth = picTable.ScaleWidth: nHeight = picTable.ScaleHeight
      cW = crdPlayingCard(0).Width: cH = crdPlayingCard(0).Height
	
	curIndex = 1
	For row = 1 to 7
		For col = row to 1 Step -1
                  If crdPlayingCard(curIndex).Data = "" Then
			     sx = (nWidth - (cW + cW * 0.05) * row) / 2
			     sx = nWidth - (sx + col * (cW + cW * 0.05))
			     sy = (row - 1) * cH * 0.5 + 240

			     crdPlayingCard(curIndex).Move sx, sy
                  End If
			curIndex = curIndex + 1
		Next
		Call gm_DoEvents
	Next
			  
      For cntr = 0 To 1
           curpos = 0
	     If cntr = 0 Then
                xpos = 100
           Else
                xpos = picTable.ScaleWidth - crdPlayingCard(0).Width - 100
           End If

           For Each oCard In crdPlayingCard
                If oCard.Data = "@" & CStr(cntr) Then 
                     oCard.Move xpos, 100 + curpos * oCard.Height * 0.2
                     curpos = curpos + 1 
                End If 
           Next
      Next

      Call UpdateControls()
      Call AlignControls()
End Sub

Sub AlignControls
      nWidth = picTable.ScaleWidth: nHeight = picTable.ScaleHeight
      cW = crdPlayingCard(0).Width: cH = crdPlayingCard(0).Height

      px = (nWidth - cW * 3) / 2
      py = 7 * cH * 0.5 + cH
      crdStock(0).Move px, py
      crdWaste(0).Move px + cW * 2, py     
      crdWaste(1).Move crdWaste(0).Left + crdWaste(0).Width * 0.2, crdWaste(0).Top
      crdWaste(2).Move crdWaste(1).Left + crdWaste(1).Width * 0.2, crdWaste(1).Top
      imgStock.Move px, py, crdStock(0).Width, crdStock(0).Height
      imgWaste.Move px + cW * 2, py, crdWaste(0).Width, crdWaste(0).Height
      shpStock.Move px, py, crdStock(0).Width, crdStock(0).Height
      shpWaste.Move px + cW * 2, py, crdWaste(0).Width, crdWaste(0).Height

      shpCircle(0).Width = crdStock(0).Width -100
      shpCircle(0).Height =  crdStock(0).Width - 100
      shpCircle(0).Move crdStock(0).Left + (crdStock(0).Width - shpCircle(0).Width) /2,       crdStock(0).Top + (crdStock(0).Height - shpCircle(0).Height) / 2

      shpStock.ZOrder 1
      shpWaste.ZOrder 1
End Sub

Sub OpenIntro()
     Call LockForm(True)

     picTable.ScaleMode = vbTwips
     Call AlignControls()

     If Speed > 0 Then
          nWidth = picTable.ScaleWidth: nHeight = picTable.ScaleHeight
          cW = crdPlayingCard(0).Width: cH = crdPlayingCard(0).Height
	
          curIndex = 1
          For row = 1 to 7
               For col = row to 1 Step -1
	              sx = (nWidth - (cW + cW * 0.05) * row) / 2
		        sx = nWidth - (sx + col * (cW + cW * 0.05))
      		  sy = 6 * cH * 0.5 + 240

	              crdPlayingCard(curIndex).Move sx, sy
   		        curIndex = curIndex + 1
	         Next
          Next	
	
          For row = 5 to 0 Step -1
               dy = row * cH * 0.5 + 240
	         Do While crdPlayingCard(Sum(1,row + 1)).Top > dy
	              For curIndex = 1 To Sum(1,row + 1)	
		             crdPlayingCard(curIndex).Move crdPlayingCard(curIndex).Left, _
			               	                   crdPlayingCard(curIndex).Top - gm_TwipsPerPixelY * Speed
   		        Next 
		   Call gm_DoEvents()
    	         Loop
          Next
     Else
          Call AlignCards
     End If

     For curIndex = 22 to 28 
          If crdPlayingCard(curIndex).Effect <> 0 then
	         crdPlayingCard(curIndex).PlayAni
	    Else
		   crdPlayingCard(curIndex).Face = 0 ' Face Up
	    End If
     Next 

     Do While IsCardAniOnEx(frmMain)
          Call gm_DoEvents
     Loop
     Call LockForm(False)
End Sub

Sub CloseIntro()
     ' ***
End Sub

Sub DoKeyDown(Index, KeyCode)
     ' do nothing
End Sub

Sub DoWasteClick(thisCard)
     ' do nothing
End Sub

Sub DoStockClick(thisCard)
     Call LockForm(True)

     ' thisCard must be an object (ex. form, textbox, label...)
     If Not IsObject(thisCard) Then 
          Call LockForm(False)
          Exit Sub
     Else
          ' check object type
          If TypeName(thisCard) <> "Card" Then
               ' if not a card then exit
               Call LockForm(False)
               Exit Sub
          End If 
     End If

     IndexWasteCard = GetTopWasteCard

     If crdWaste(IndexWasteCard).Selected Then
          For i = 1 To DataColl.Count
               If DataColl(i).Name = crdWaste(IndexWasteCard).Name Then
                    DataColl.Remove i
               End If
          Next

          crdWaste(IndexWasteCard).Selected = False
    End If

    If ValHolder1 >= StockColl.Count Then
	    Set StockColl = WasteColl
          Set WasteColl = Nothing  
          
	    If StockColl.Count > 0 Then 
               thisCard.Visible = True
               For i = 0 To crdWaste.Count - 1
                    crdWaste(i).Visible = False
               Next    
               ValHolder1 = 0
          End If

          Call LockForm(False)
          Exit Sub
     End If	

     Call PlaySound(101)
     curIndex = ValHolder1 + 1
     For i = curIndex To StockColl.Count
          ValHolder1 = ValHolder1 + 1 
          If crdPlayingCard(StockColl(i)).Data <> "$" Then
               WasteColl.Add StockColl(i)
          End If
          If (ValHolder1 Mod 3) = 0 Then Exit For
     Next

     If ValHolder1 >= StockColl.Count Then thisCard.Visible = False

     crdTemp(0).Update = False     
     crdTemp(0).Face = 1 ' Face Down
     crdTemp(0).Rank = crdPlayingCard(WasteColl(WasteColl.Count)).Rank
     crdTemp(0).Suit = crdPlayingCard(WasteColl(WasteColl.Count)).Suit
     crdTemp(0).Update = True
     crdTemp(0).Refresh
     crdTemp(0).Move thisCard.Left, thisCard.Top, thisCard.Width, thisCard.Height
     crdTemp(0).ZOrder 0

     If Not crdTemp(0).Visible Then crdTemp(0).Visible = True

     cx = (crdWaste(0).Left - thisCard.Left) / 2
     sx = (cx / thisCard.Width) * gm_TwipsPerPixelX * 2
    
     Do While crdTemp(0).Left < crdWaste(0).Left
          If crdTemp(0).Left < cx + thisCard.Left Then
               w = crdTemp(0).Width - sx
               If w < 0 Then w = 0 
          Else
               If w <> thisCard.Width Then
                    w = crdTemp(0).Width + sx
                    If w > thisCard.Width Then w = thisCard.Width
                         If crdTemp(0).Face = 1 Then crdTemp(0).Face = 0
               End If
          End If

          crdTemp(0).Move crdTemp(0).Left + sx, thisCard.Top, w, thisCard.Height 
               Call gm_DoEvents
     Loop
     
     If crdTemp(0).Visible Then crdTemp(0).Visible = False 
 
     For i = 0 To crdWaste.Count - 1
          crdWaste(i).Visible = False
     Next 

     modWaste = (ValHolder1 Mod 3)
     If modWaste = 0 Then
          nLen = 2
     ElseIf modWaste = 1 Then
          nLen = 0
     Else
          nLen = 1
     End If           

     curIndex = 0
     For i = WasteColl.Count - nLen To WasteColl.Count
          crdWaste(curIndex).Update = False
          crdWaste(curIndex).Rank = crdPlayingCard(WasteColl(i)).Rank
          crdWaste(curIndex).Suit = crdPlayingCard(WasteColl(i)).Suit
          crdWaste(curIndex).Update = True
          crdWaste(curIndex).Refresh
          crdWaste(curIndex).ZOrder 0
          crdWaste(curIndex).Visible = True
          curIndex = curIndex + 1
     Next 

     If crdWaste(1).Visible Then
          crdWaste(1).Move crdWaste(0).Left + crdWaste(0).Width * 0.2, _
                           crdWaste(0).Top + gm_TwipsPerPixelY
     End If

     If crdWaste(2).Visible Then
          crdWaste(2).Move crdWaste(1).Left + crdWaste(1).Width * 0.2, _
                           crdWaste(1).Top + gm_TwipsPerPixelY
     End If

     Do While IsCardAniOnEx(frmMain)
          Call gm_DoEvents
     Loop
     
     Call LockForm(False)
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

     If (Left(thisCard.Data,1) = "@") Or (thisCard.Face = 1) Then
          Call LockForm(False)
          Exit Function 
     End If

     If thisCard.Name = "crdWaste" Then
          If thisCard.Index <> GetTopWasteCard Then 
               Call LockForm(False)
               Exit Function
          End If
     End If
 
     If thisCard.Selected Then
          For i = 1 To DataColl.Count
               If DataColl(i).Name = thisCard.Name Then
                    DataColl.Remove i
               End If
          Next

          For i = 1 To TempColl.Count
               If TempColl(i).Name = thisCard.Name Then
                    TempColl.Remove i
               End If
          Next

          If thisCard.Effect <> 0 Then
               ' set effect
               SetRevAni thisCard, GetRevAni(thisCard.Effect, _
                                             GetTypeEffect(thisCard))

               thisCard.Update = False
               thisCard.AutoFlipCard = False
               thisCard.Selected = False
               thisCard.PlayAni

               Do While IsCardAniOnEx(frmMain)
                    Call gm_DoEvents
               Loop
                          
               thisCard.AutoFlipCard =True
               thisCard.Update = True
               thisCard.Refresh 

               ' restore previous effect
               SetRevAni thisCard, GetRevAni(thisCard.Effect, _
                                             GetTypeEffect(thisCard)) 
          Else
               thisCard.Selected = False
          End If

          Call LockForm(False)
          Exit Function
     End If 

     If thisCard.Effect <> 0 Then	 
          If IsCardAniOnEx(frmMain) Then 
               Call LockForm(False)
               Exit Function
	    Else
               If Not thisCard.Selected Then
                    thisCard.Update = False
                    thisCard.AutoFlipCard = False
                    thisCard.Selected = Not thisCard.Selected
                    thisCard.PlayAni
			
                    Do While IsCardAniOnEx(frmMain)
 		             Call gm_DoEvents
                    Loop

                    thisCard.AutoFlipCard =True
                    thisCard.Update = True
                    thisCard.Refresh 
               End If
 	    End If
     Else
          thisCard.Selected = Not thisCard.Selected 	
	    thisCard.Refresh
          If DataColl.Count = 2 Then Call gm_Delay(1000)
     End If

     If DataColl.Count > 0 Then
          curIndex = 1

          Do
               If DataColl(curIndex).Name <> "crdWaste" Then
	              If DataColl(curIndex).Index = thisCard.Index Then 
	                   If DataColl(curIndex).Selected = False Then 
                              DataColl.Remove curIndex
				      curIndex = 1
                          End If
	              End If
              
                End If
 
                curIndex = curIndex + 1 
          Loop While (curIndex < DataColl.Count + 1)
     End If

     Call PlaySound(101)	
     If thisCard.Selected Then ' Face Up
          DataColl.Add thisCard
     End If

     If (DataColl.Count = 1) Or (DataColl.Count = 2) Then
          IsRule = False
          If DataColl.Count = 1 Then
               If DataColl(1).Value = 13 Then
                    IsRule = True
               End If
          End If

          If (DataColl.Count = 2) And Not IsRule Then  
               If (DataColl(1).Value + DataColl(2).Value) = 13 Then
                    IsRule = True
               End If 
          End If 
 
          If IsRule Then
               For i = 1 To DataColl.Count
                    Call gm_Delay(500) ' wait for 0.5 second
                    If DataColl(i).Name <> "crdWaste" Then
                         DataColl(i).Data = "@"
                         DataColl(i).Visible = False 
                    End If
               Next

		   If DataColl.Count = 1 Then
                    Score = Score + 25
               Else
                    Score = Score + 50
               End If	

               For i = 1 To DataColl.Count
                    If DataColl(i).Name <> "crdWaste" Then
                         pos = GetCurRow(DataColl(i).Index)
                    
                         If pos = GetCurRow(DataColl(i).Index - 1) Then ' Left
                              If Left(crdPlayingCard(DataColl(i).Index - 1).Data, 1) = "@" Then
                                   TempColl.Add crdPlayingCard(DataColl(i).Index - pos)
                              End If 
                         End If

                         If pos = GetCurRow(DataColl(i).Index + 1) Then ' Right
                              If Left(crdPlayingCard(DataColl(i).Index + 1).Data, 1) = "@" Then
                                   TempColl.Add crdPlayingCard(DataColl(i).Index - pos + 1)
                              End If
                         End If     
                    Else
                         TempColl.Add DataColl(i)
                    End If
               Next
               
               ' TempColl -> Total cards to be open
               For i = 1 To TempColl.Count
                    If TempColl(i).Name <> "crdWaste" Then
                         If TempColl(i).Face = 1 Then ' face down
                              If TempColl(i).Effect <> 0 Then
                                   TempColl(i).PlayAni ' face up with animation           
                              Else
                                   TempColl(i).Face = 0
                              End if 
                         End If
                    Else
                        IndexWasteCard = GetTopWasteCard 
		            If WasteColl.Count = 0 Then
                             ' do nothing
                        Else
                             DataColl.Add crdPlayingCard(WasteColl(WasteColl.Count))
                             WasteColl.Remove WasteColl.Count

                             If WasteColl.Count = 0 Then
                                  crdWaste(0).Visible = False
                             Else
                                  If crdWaste(1).Visible Or crdWaste(2).Visible Then    
                                       crdWaste(IndexWasteCard).Visible = False
                                  Else 
                                       crdWaste(0).Update = False 
                                       crdWaste(0).Rank = crdPlayingCard(WasteColl(WasteColl.Count)).Rank
                                       crdWaste(0).Suit = crdPlayingCard(WasteColl(WasteColl.Count)).Suit
                                       crdWaste(0).Update = True
                                       crdWaste(0).Refresh
                                  End If
                             End If
                        End If
                        If crdWaste(IndexWasteCard).Selected Then crdWaste(IndexWasteCard).Selected = False
                    End If         
                    Score = Score + 5 ' bonus score. 5 points multiply by total cards open 
               Next
               
               For i = 1 To DataColl.Count
                    If DataColl(i).Name <> "crdWaste" Then 
                         If ValHolder2 = 26 Then ValHolder2 = 0: ValHolder3 = 1
                         If DataColl(i).Selected Then DataColl(i).Selected = False
                         If DataColl(i).Data = "" Then DataColl(i).Data = "@"  & ValHolder3
                         If Len(DataColl(i).Data) = 1 Then DataColl(i).Data = DataColl(i).Data & ValHolder3

                         If ValHolder3 = 0 Then
                              xpos = 100 
                         Else                         
                              xpos = picTable.ScaleWidth - crdPlayingCard(0).Width - 100
                         End If

                         DataColl(i).Face = 0 ' Face Up
                         DataColl(i).Visible = True
                         DataColl(i).Move xpos, 100 + ValHolder2 * DataColl(i).Height * 0.2
                         DataColl(i).ZOrder 0
                         ValHolder2 = ValHolder2 + 1 ' save location
                    End If
               Next

               Set TempColl = Nothing

               Do While IsCardAniOnEx(frmMain)
 		        Call gm_DoEvents()
               Loop
          Else
               If DataColl.Count = 2 Then
                    Call gm_Delay(500) ' wait for 0.5 second

                    For i = 1 To DataColl.Count
                         If DataColl(i).Effect <> 0 Then
                              ' set effect
                              SetRevAni DataColl(i), GetRevAni(DataColl(i).Effect, _
                                                               GetTypeEffect(DataColl(i)))

                              DataColl(i).Update = False
                              DataColl(i).AutoFlipCard = False
                              DataColl(i).Selected = False
                              DataColl(i).PlayAni
                         Else
                              DataColl(i).Selected = False
                         End If
			  Next 

	              Do While IsCardAniOnEx(frmMain)
 		             Call gm_DoEvents
                    Loop

		        For i = 1 To DataColl.Count		
                         If DataColl(i).Effect <> 0 Then
                              DataColl(i).AutoFlipCard = True
                              DataColl(i).Update = True
                              DataColl(i).Refresh 

                              ' restore previous effect
                              SetRevAni DataColl(i), GetRevAni(DataColl(i).Effect, _
                                                               GetTypeEffect(DataColl(i)))
                         End If
                    Next
               End If
          End If 

          If IsRule Or (DataColl.Count = 2) Then
               nSum = 0
               For Each oCard In crdPlayingCard
                    If (oCard.Index >= 1) And (oCard.Index <=28) Then 
                         If oCard.Data = "" Then
                              nSum = nSum + 1
                         End If
                    End If
               Next

               If nSum = 0 Then Process = True
               Set DataColl = Nothing
          End If
     End If

     Do While IsCardAniOnEx(frmMain)
          Call gm_DoEvents
     Loop
     
     Call LockForm(False)
End Function

Function GetCurRow(nVal)
     GetCurRow = 0

     If Not IsNumeric(nVal) Then Exit Function ' if not numeric then exit sub
     
     cntr = 0
     For row = 1 To 7
          For col = 1 To row
               cntr = cntr + 1
               If cntr = nVal Then 
                    GetCurRow = row
                    Exit Function
               End If       
          Next
     Next	 
End Function

Function GetTopWasteCard()
     If crdWaste(2).Visible Then
          GetTopWasteCard = 2
     Else
          If crdWaste(1).Visible Then
               GetTopWasteCard = 1
          Else
               GetTopWasteCard = 0
          End If
     End If  
End Function

Function Sum(nStart, nEnd)
    s = 0
    For i = nStart To nEnd
        s = s + i
    Next
    
    Sum = s
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
     sData = ""
     For Each oCard In StockColl
          sData = sData & oCard & "/$"
     Next
     If Right(sData, 2) = "/$" Then sData = Left(sData,Len(sData) - 2)
     If sData = "" Then sData = "/$"

     GetStockColl = sData
End Function

Function GetWasteColl
     sData = ""
     For Each oCard In WasteColl
          sData = sData & oCard & "/$"
     Next
     If Right(sData, 2) = "/$" Then sData = Left(sData,Len(sData) - 2)
     If sData = "" Then sData = "/$"

     GetWasteColl = sData
End Function

Sub SetDataColl(sData)
     Set DataColl = Nothing
 
     If CStr(sData) = "/$" Then Exit Sub
     arrData = Split(sData, "/$")

     For i = LBound(arrData) To UBound(arrData) Step 2
          For Each oCard In frmMain.Controls
               If oCard.Tag = "PlayingCard" Then
                   If (CStr(arrData(i)) = oCard.Name) And _
                      (CInt(arrData(i + 1)) = oCard.Index) Then
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

     If CStr(sData) = "/$" Then Exit Sub

     arrData = Split(sData, "/$")
     For i = LBound(arrData) To UBound(arrData)
          StockColl.Add arrData(i)
     Next

    If StockColl.Count = 1 Then 
         If StockColl(1) = "/$" Then
              Set StockColl = Nothing
         End If
    End If
End Sub	

Sub SetWasteColl(sData)
     Set WasteColl = Nothing

     If CStr(sData) = "/$" Then Exit Sub

     arrData = Split(sData, "/$")
     For i = LBound(arrData) To UBound(arrData)
          WasteColl.Add arrData(i)
     Next
End Sub

Sub SetValHolder
     ValHolder1 = CInt(ValHolder1)
     ValHolder2 = CInt(ValHolder2)
     ValHolder3 = CInt(ValHolder3)
End Sub

Sub UpdateControls
     For cntr = 0 To 1
          Set TempColl = Nothing

          For Each oCard In crdPlayingCard
               If oCard.Data = "@" & CStr(cntr) Then
                    TempColl.Add oCard  
               End If
          Next

          If TempColl.Count > 0 Then
               For i = 1 To TempColl.Count
                    nIndex = 0: TempVal = TempColl(1).Top
                    For j = 1 To TempColl.Count
                         If TempVal <= TempColl(j).Top Then
                              nIndex = j: TempVal = TempColl(j).Top
                         End If
                    Next
                    TempColl(nIndex).ZOrder 1
                    TempColl.Remove nIndex
               Next 
          End If
     Next     
End Sub