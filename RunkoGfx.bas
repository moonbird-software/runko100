Attribute VB_Name = "RunkoGfx"
Option Explicit

' time constants
Public Const IDT_FLASH_CARD = 1500
Public Const IDT_SHOW_CARD = 500
Public Const IDT_MOVE_CARD = 2000
Public Const IDT_POP_CARD = 100
'Public Const IDT_POP_CARD_CPU = 25
Public Const IDT_POP_CARD_CPU = IDT_POP_CARD

' color constants
Public Const IDC_TABLEBG = &H8000&
Public Const IDC_TITLE = &H0
Public Const IDC_TITLE_SEL = &HFFFFFF
Sub AnimTakeCards(Trick As CardDeck, Player As CardDeck)
    Debug.Print Game.Turn + 1, "nostaa"
    SetStatus Replace(IDS_STATUS_TAKE, "%s", Game.Title(Player.Index))
    AnimPopCards Trick, Player, CountCards(Trick), True
End Sub
Sub AnimMoveCards(Source As CardDeck, Dest As CardDeck, ByVal Mode As Integer, Optional ByVal Rank As Integer, Optional ByVal Count As Integer, Optional ByVal fIncDest As Integer)
Dim nCards As Integer, nCardsDest As Integer, iDeck As Integer
Dim iCard As Integer, iCardTo As Integer, iCardMove As Integer, iMoved As Integer
Dim x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer
Dim X As Integer, Y As Integer
Dim fFaceUp As Boolean

    DoEvents
    
    nCards = CountCards(Source)
    iCardTo = CountCards(Dest)
    iDeck = Dest.Index
    
    ' animate card moves to dest deck
    For iCard = 0 To nCards - 1
        If Source.Card(iCard) <> -1 And Source.Mode(iCard) = Mode And _
        (GetRank(Source.Card(iCard)) = Rank Or Rank = 0) Then
            
            ' 1. draw sprite
            
            If IsCardFaceUp(Source, Deck(iDeck)) Then
                cdtDraw frmMain.picDeck(IDD_SPRITE).hdc, 0, 0, Source.Card(iCard), mdFaceUp, IDC_TABLEBG
            Else
                cdtDraw frmMain.picDeck(IDD_SPRITE).hdc, 0, 0, Game.CardBack, mdFaceDown, IDC_TABLEBG
            End If
            
            GetCardXY Source, iCard, nCards, X, Y
            
            x0 = frmMain.picDeck(Source.Index).Left + X
            y0 = frmMain.picDeck(Source.Index).Top + Y
          
            ' 2. hide selected card
            
            Source.Mode(iCard) = cmHidden
            DrawDeck Source
            
            With frmMain.picDeck(IDD_SPRITE)
                .Left = x0
                .Top = y0
                .Visible = True
            End With
            
            ' 3. move sprite
            
            nCardsDest = CountCards(Deck(iDeck))
            GetCardXY Deck(iDeck), nCardsDest, nCardsDest, X, Y
            x1 = frmMain.picDeck(Deck(iDeck).Index).Left + X
            y1 = frmMain.picDeck(Deck(iDeck).Index).Top + Y
            PlaySound IDSND_CARDMOVE
            AnimObject frmMain.picDeck(IDD_SPRITE), x0, y0, x1, y1
            
            ' 4. add card to dest deck and hide sprite
            
            Deck(iDeck).Card(iCardTo) = Source.Card(iCard)
            iCardTo = iCardTo + 1
            
            DrawDeck Deck(iDeck)
            With frmMain.picDeck(IDD_SPRITE)
                .Visible = False
            End With
            
            Delay IDT_SHOW_CARD
            DoEvents
            
            iMoved = iMoved + 1
            If iMoved = Count Then
                Exit For
            End If
            If fIncDest Then
                iDeck = iDeck + 1
                iCardTo = CountCards(Deck(iDeck))
            End If
        End If
    Next iCard

    ' remove cards from source deck
    iMoved = 0
    
    For iCard = nCards To 0 Step -1
        If Source.Card(iCard) <> -1 And Source.Mode(iCard) = cmHidden And (GetRank(Source.Card(iCard)) = Rank Or Rank = 0) Then
            For iCardMove = iCard To MAX_CARDS - 1
                Source.Card(iCardMove) = Source.Card(iCardMove + 1)
                Source.Mode(iCardMove) = Source.Mode(iCardMove + 1)
            Next iCardMove

            iMoved = iMoved + 1
            If iMoved = Count Then
                Exit For
            End If
        End If
    Next iCard
    
    DrawDeck Source
    
End Sub
Sub DrawCardBasic(Obj As Object, Deck As CardDeck, iCard As Integer, nCards As Integer, X As Integer, Y As Integer)
    If nCards > 0 Then
        Select Case Deck.Index
        Case IDD_USER
            Select Case Deck.Mode(iCard)
            Case cmNormal, cmSelected
                cdtDraw Obj.hdc, X, Y, Deck.Card(iCard), mdFaceUp, IDC_TABLEBG
            Case cmHilite
                cdtDraw Obj.hdc, X, Y, Deck.Card(iCard), mdHilite, IDC_TABLEBG
            Case cmHidden
            End Select
            Exit Sub
        Case IDD_PLAYER1, IDD_PLAYER2, IDD_PLAYER3, IDD_PLAYER4
            If Game.Demo Or Game.Debug Then
                Select Case Deck.Mode(iCard)
                Case cmNormal, cmSelected
                    cdtDraw Obj.hdc, X, Y, Deck.Card(iCard), mdFaceUp, IDC_TABLEBG
                Case cmHilite
                    cdtDraw Obj.hdc, X, Y, Deck.Card(iCard), mdHilite, IDC_TABLEBG
                Case cmHidden
                    '
                End Select
            Else
                Select Case Deck.Mode(iCard)
                Case cmNormal, cmHilite
                    cdtDraw Obj.hdc, X, Y, Game.CardBack, mdFaceDown, IDC_TABLEBG
                Case cmSelected
                    cdtDraw Obj.hdc, X, Y, Deck.Card(iCard), mdFaceUp, IDC_TABLEBG
                Case cmHidden
                
                End Select
            End If
        End Select
    End If
End Sub
Sub DrawTurnIcon(fCheckHand As Boolean)
'    With frmMain
'
'        .imaTurn(0).Visible = False
'        .imaTurn(1).Visible = False
'
'        .imaTurn(0).Left = .lblPlayer(Game.Turn).Left
'        .imaTurn(1).Left = .imaTurn(0).Left
'
'        Select Case Game.Turn
'        Case IDD_PLAYER1, IDD_PLAYER2
'            .imaTurn(0).Top = .lblPlayer(Game.Turn).Top - .imaTurn(0).Height - 5
'            .imaTurn(1).Top = .imaTurn(0).Top
'        Case IDD_PLAYER3, IDD_PLAYER4
'            .imaTurn(0).Top = .lblPlayer(Game.Turn).Top + .imaTurn(0).Height + 5
'            .imaTurn(1).Top = .imaTurn(0).Top
'        End Select
'
'        If fCheckHand Then
'            .imaTurn(0).Visible = True
'            .imaTurn(1).Visible = False
'        Else
'            .imaTurn(0).Visible = False
'            .imaTurn(1).Visible = True
'        End If
'
'    End With
End Sub
Function GetCardStepBasic(Deck As CardDeck, ByVal nCards As Integer, Optional ByRef fStepX As Boolean, Optional ByRef fReverse As Boolean) As Integer
    ' set card stepping and direction x/y
    Select Case Deck.Index
    Case IDD_PLAYER1, IDD_PLAYER3
        If nCards > 18 Then
            GetCardStepBasic = cdWidth / 8
        Else
            GetCardStepBasic = cdWidth / 4
        End If
        fStepX = True
        
    Case IDD_PLAYER2, IDD_PLAYER4
        If nCards > 18 Then
            GetCardStepBasic = cdHeight / 12
        Else
            GetCardStepBasic = cdHeight / 6
        End If
        fStepX = False
    End Select
    
    ' set card drawing order
    Select Case Deck.Index
    Case IDD_PLAYER1, IDD_PLAYER2
        fReverse = False
    
    Case IDD_PLAYER3, IDD_PLAYER4
        fReverse = True
        
    End Select
End Function
Sub GetCardXYBasic(Deck As CardDeck, ByVal iCard As Integer, ByVal nCards As Integer, ByRef X As Integer, ByRef Y As Integer)
Dim iStep As Integer, iPos As Integer
Dim fStepX As Boolean, fReverse As Boolean

    ' set card stepping
    iStep = GetCardStepBasic(Deck, nCards, fStepX, fReverse)
    
    ' get card coordinates
    iPos = iStep * iCard
    
    Select Case Deck.Index
    Case IDD_PLAYER1, IDD_PLAYER2, IDD_PLAYER3, IDD_PLAYER4
        If fReverse Then
            iPos = (nCards - iCard - 1) * iStep
        End If
        If fStepX Then
            X = iPos
            Y = 0
        Else
            X = 0
            Y = iPos
        End If
    End Select
    
    Select Case Deck.Index
    Case IDD_PLAYER1
        Select Case Deck.Mode(iCard)
        Case cmNormal, cmHilite
            Y = cdHeight / 6
        Case cmSelected
            Y = 0
        End Select
    Case IDD_PLAYER2
        Select Case Deck.Mode(iCard)
        Case cmNormal, cmHilite
            X = 0
        Case cmSelected
            X = cdWidth / 4
        End Select
    Case IDD_PLAYER3
        Select Case Deck.Mode(iCard)
        Case cmNormal, cmHilite
            Y = 0
        Case cmSelected
            Y = cdHeight / 6
        End Select
    Case IDD_PLAYER4
        Select Case Deck.Mode(iCard)
        Case cmNormal, cmHilite
            X = cdWidth / 4
        Case cmSelected
            X = 0
        End Select
    End Select
End Sub
Function IsCardFaceUpBasic(Source As CardDeck, Dest As CardDeck) As Boolean
    If Game.Demo Then
        IsCardFaceUpBasic = True
        Exit Function
    End If
    If IsPlayerDeck(Source) And Dest.Index = IDD_TRICK Then
        IsCardFaceUpBasic = True
    End If
End Function


Sub SetDeckProps(Deck As CardDeck)
Dim sToolTip As String
Dim nCards As Integer

    ' set tooltip text
    If Deck.Name <> "" Then
        
        nCards = CountCards(Deck)
        sToolTip = Deck.Name & ": " & Format(nCards) & " "
    
        Select Case nCards
        Case 0
            sToolTip = ""
        Case 1
            sToolTip = sToolTip & IDS_CARD
        Case Else
            sToolTip = sToolTip & IDS_CARDS
        End Select
        
        frmMain.picDeck(Deck.Index).ToolTipText = sToolTip
    
    End If
    
    ' resize decks
    FormMainResize
    
End Sub
Sub AnimHiliteDeck(Deck As CardDeck)
Dim iCard As Integer
    DoEvents
    If CountCards(Deck) > 0 Then
        For iCard = 0 To MAX_CARDS
            Deck.Mode(iCard) = cmHilite
        Next iCard
        DrawDeck Deck
        For iCard = 0 To MAX_CARDS
            Deck.Mode(iCard) = cmNormal
        Next iCard
    End If
End Sub
Sub DrawDeckBuffer(Deck As CardDeck)
Dim iCard As Integer
Dim nCards As Integer
Dim X As Integer, Y As Integer
    nCards = CountCards(Deck)
    'SetDeckPosBuffer Deck
    If nCards = 0 Then
        'DrawCard frmScreen.picScreen, Deck, iCard, nCards, Deck.X, Deck.Y
    Else
        For iCard = 0 To nCards - 1
            GetCardXY Deck, iCard, nCards, X, Y
            'DrawCard frmScreen.picScreen, Deck, iCard, nCards, Deck.X + X, Deck.Y + Y
        Next iCard
    End If
    'frmScreen.picScreen.Refresh
End Sub

Sub FormMainResizeBasic()
Dim iPlr As Integer, iFormWidth As Integer, iFormHeight As Integer, iCardStep As Integer
Dim nCards As Integer
    With frmMain
        If .WindowState <> vbMinimized Then
            If frmMain.Width < MIN_MAIN_WIDTH Then
                frmMain.Width = MIN_MAIN_WIDTH
            End If
            If frmMain.Height < MIN_MAIN_HEIGHT Then
                frmMain.Height = MIN_MAIN_HEIGHT
            End If
        End If
    End With
    For iPlr = 0 To MAX_PLAYERS - 1
        With frmMain.picDeck(iPlr)
            
            nCards = CountCards(Deck(iPlr))
            iCardStep = GetCardStep(Deck(iPlr), nCards)
            iFormWidth = frmMain.ScaleWidth
            iFormHeight = frmMain.ScaleHeight - frmMain.sbrStatus.Height
            
            Select Case Deck(iPlr).Index
            Case IDD_PLAYER1
                .Height = cdHeight + cdHeight / 6
                .Width = cdWidth + ((nCards - 1) * iCardStep)
                .Left = (iFormWidth - .Width) / 2
                .Top = iFormHeight - .Height - 8
                frmMain.lblPlayer(iPlr).Top = .Top + (cdWidth / 4)
                frmMain.lblPlayer(iPlr).Left = .Left - frmMain.lblPlayer(iPlr).Width - 5
            
            Case IDD_PLAYER2
                .Height = cdHeight + ((nCards - 1) * iCardStep)
                .Width = cdWidth + cdWidth / 4
                .Left = 8
                .Top = (iFormHeight - .Height) / 2
                frmMain.lblPlayer(iPlr).Top = .Top - 20
                frmMain.lblPlayer(iPlr).Left = .Left
            
            Case IDD_PLAYER3
                .Height = cdHeight + cdHeight / 6
                .Width = cdWidth + ((nCards - 1) * iCardStep)
                .Left = (iFormWidth - .Width) / 2
                .Top = 8
                frmMain.lblPlayer(iPlr).Top = .Top
                frmMain.lblPlayer(iPlr).Left = .Left + .Width + 5
            
            Case IDD_PLAYER4
                .Height = cdHeight + ((nCards - 1) * iCardStep)
                .Width = cdWidth + cdWidth / 4
                .Left = iFormWidth - .Width - 8
                .Top = (iFormHeight - .Height) / 2
                frmMain.lblPlayer(iPlr).Top = .Top + .Height + 5
                frmMain.lblPlayer(iPlr).Left = .Left + (cdWidth / 4)
            End Select
        End With
    Next iPlr
End Sub

Sub DrawAllDecksBuffer()
Dim iDeck As Integer
    'frmScreen.picScreen.Cls
    For iDeck = 0 To MAX_DECKS - 1
        DrawDeckBuffer Deck(iDeck)
    Next iDeck
End Sub
Function GetCardIndex(Deck As CardDeck, ByVal X As Single, ByVal Y As Single) As Integer
Dim iCard As Integer
Dim nCards As Integer
Dim x1 As Integer, x0 As Integer
Dim y1 As Integer, y0 As Integer

    ' count cards
    nCards = CountCards(Deck)
    
    ' go thru all cards and return index if found
    For iCard = 0 To nCards - 1
    
        x0 = x1
        y0 = y1
        
        GetCardXY Deck, iCard + 1, nCards, x1, y1
        
        Select Case Deck.Index
        Case IDD_PLAYER1
            If iCard = nCards - 1 Then
                x1 = x0 + cdWidth
            End If
            If X >= x0 And X < x1 Then
                GetCardIndex = iCard
                Exit Function
            End If
        Case IDD_PLAYER3
            If iCard = nCards - 1 Then
                x1 = x0 - cdWidth
            End If
            If X <= (x0 + cdWidth) And X > (x1 + cdWidth) Then
                GetCardIndex = iCard
                Exit Function
            End If
        Case IDD_PLAYER2
            If iCard = nCards - 1 Then
                y1 = y0 + cdHeight
            End If
            If Y >= y0 And Y < y1 Then
                GetCardIndex = iCard
                Exit Function
            End If
        Case IDD_PLAYER4
            If iCard = nCards - 1 Then
                y1 = y0 - cdHeight
            End If
            If Y <= (y0 + cdHeight) And Y > (y1 + cdHeight) Then
                GetCardIndex = iCard
                Exit Function
            End If
        End Select
        
    Next iCard
    
End Function


Sub AnimFlashCard(Deck As CardDeck, iCard As Integer)
    DoEvents
    If Game.Sound = True Then
        Beep
    End If
    Deck.Mode(iCard) = cmHilite
    DrawDeck Deck
    Delay IDT_FLASH_CARD
    Deck.Mode(iCard) = cmNormal
    DrawDeck Deck
End Sub
Sub AnimMoveCards2(Source As CardDeck, Dest As CardDeck, ByVal Mode As Integer, Optional ByVal Rank As Integer, Optional ByVal Count As Integer)
Dim nCards As Integer, nCardsDest As Integer
Dim iCard As Integer, iCardTo As Integer, iCardMove As Integer, iMoved As Integer
Dim x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer
Dim X As Integer, Y As Integer
Dim fFaceUp As Boolean
    
    DoEvents
    
    nCards = CountCards(Source)
    iCardTo = CountCards(Dest)
    
    For iCard = nCards To 0 Step -1
        If Source.Card(iCard) <> -1 And Source.Mode(iCard) = Mode And (GetRank(Source.Card(iCard)) = Rank Or Rank = 0) Then
            
            ' 1. kortti valittuna
            
            'DrawDeck Source
            Delay IDT_MOVE_CARD

            ' 2. sprite p‰‰lle
            
            'If IsPlayerDeck(Source) And IsPlayerDeck(Dest) And Source.Index <> IDD_USER And Dest.Index <> IDD_USER Then
            '    fFaceUp = False
            'End If
            
            If IsCardFaceUp(Source, Dest) Then
                cdtDraw frmMain.picDeck(IDD_SPRITE).hdc, 0, 0, Source.Card(iCard), mdFaceUp, IDC_TABLEBG
            Else
                cdtDraw frmMain.picDeck(IDD_SPRITE).hdc, 0, 0, Game.CardBack, mdFaceDown, IDC_TABLEBG
            End If
            
            GetCardXY Source, iCard, nCards, X, Y
            x0 = frmMain.picDeck(Source.Index).Left + X
            y0 = frmMain.picDeck(Source.Index).Top + Y
            frmMain.picDeck(IDD_SPRITE).Left = x0
            frmMain.picDeck(IDD_SPRITE).Top = y0
            frmMain.picDeck(IDD_SPRITE).Visible = True
            
            ' 3. valittu kortti piiloon
            
            Source.Mode(iCard) = cmHidden
            DrawDeck Source
            
            ' 4. sprite liikkeelle
            
            nCardsDest = CountCards(Dest)
            GetCardXY Dest, nCardsDest, nCardsDest, X, Y
            x1 = frmMain.picDeck(Dest.Index).Left + X
            y1 = frmMain.picDeck(Dest.Index).Top + Y
            PlaySound IDSND_CARDMOVE
            AnimObject frmMain.picDeck(IDD_SPRITE), x0, y0, x1, y1
            
            ' 5. kortti toiseen pakkaan
            
            Dest.Card(iCardTo) = Source.Card(iCard)
            iCardTo = iCardTo + 1
            For iCardMove = iCard To MAX_CARDS - 1
                Source.Card(iCardMove) = Source.Card(iCardMove + 1)
                Source.Mode(iCardMove) = Source.Mode(iCardMove + 1)
            Next iCardMove
            
            frmMain.picDeck(IDD_SPRITE).Visible = False
            DrawDeck Source
            DrawDeck Dest
            
            Delay IDT_SHOW_CARD
            DoEvents
            
            iMoved = iMoved + 1
            If iMoved = Count Then
                Exit For
            End If
        End If
    Next iCard
    
End Sub
Sub AnimMoveSelCards(Source As CardDeck, Dest As CardDeck, Optional iRankFirst As Integer)
    If iRankFirst > 0 Then
        AnimMoveCards Source, Dest, cmSelected, iRankFirst
    End If
    AnimMoveCards Source, Dest, cmSelected
End Sub
Sub AnimPopCards(Source As CardDeck, Dest As CardDeck, ByVal nCards As Integer, ByVal fSort As Boolean)
Dim iCard As Integer, iDelay As Integer, iCardSrc As Integer, iStep As Integer
Dim X As Integer, Y As Integer
Dim x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer
Dim nCards2 As Integer
    
    DoEvents
    
    For iCard = 1 To nCards
        
        ' draw sprite
        If IsCardFaceUp(Source, Dest) Then
            cdtDraw frmMain.picDeck(IDD_SPRITE).hdc, 0, 0, GetTopCard(Source), mdFaceUp, IDC_TABLEBG
        Else
            cdtDraw frmMain.picDeck(IDD_SPRITE).hdc, 0, 0, Game.CardBack, mdFaceDown, IDC_TABLEBG
        End If
        
'        Select Case Dest.Index
'        Case IDD_USER, IDD_TRICK
'            cdtDraw frmMain.picDeck(IDD_SPRITE).hdc, 0, 0, GetTopCard(Source), mdFaceUp, IDC_TABLEBG
'        Case Else
'            If Game.Demo Then
'                cdtDraw frmMain.picDeck(IDD_SPRITE).hdc, 0, 0, GetTopCard(Source), mdFaceUp, IDC_TABLEBG
'            Else
'                cdtDraw frmMain.picDeck(IDD_SPRITE).hdc, 0, 0, Game.CardBack, mdFaceDown, IDC_TABLEBG
'            End If
'        End Select
        
        ' set delay
        Select Case Dest.Index
        Case IDD_USER
            iDelay = IDT_POP_CARD
        Case IDD_TRICK
            iDelay = IDT_SHOW_CARD
        Case Else
            iDelay = IDT_POP_CARD_CPU
        End Select
        
        ' pop one card from source to dest
        PopCards Source, Dest, 1
        
        ' draw source deck
        Delay IDT_SHOW_CARD
        DrawDeck Source
        
        ' get start coordinates
        nCards2 = CountCards(Source)
        GetCardXY Source, nCards2, nCards2, X, Y
        x0 = frmMain.picDeck(Source.Index).Left + X
        y0 = frmMain.picDeck(Source.Index).Top + Y
        
        ' get end coordinates
        nCards2 = CountCards(Dest)
        GetCardXY Dest, nCards2 - 1, nCards2 - 1, X, Y
        'GetCardXY Dest, nCards2, nCards2 + 1, X, Y
        x1 = frmMain.picDeck(Dest.Index).Left + X
        y1 = frmMain.picDeck(Dest.Index).Top + Y
        
        ' animate sprite
        PlaySound IDSND_CARDMOVE
        AnimObject frmMain.picDeck(IDD_SPRITE), x0, y0, x1, y1, iStep, iDelay
        DoEvents
        'PopCards Source, Dest, 1
        
        ' draw dest deck
        frmMain.picDeck(IDD_SPRITE).Visible = False
        DrawTitles
        Delay IDT_SHOW_CARD
        If fSort Then
            SortDeck Dest
        End If
        DrawDeck Dest
    Next iCard
    
End Sub
Sub DrawTitles()
Dim iPlr As Integer
Dim nCards As Integer

    For iPlr = 0 To MAX_PLAYERS - 1
        
        With frmMain.lblPlayer(iPlr)
        
            ' draw player title
            .Caption = Game.Title(iPlr)
            
            ' hide empty players
            If Game.On Then
                If Game.Pos(iPlr) = 0 Then
                    .Visible = True
                Else
                    .Visible = False
                End If
            Else
                .Visible = False
            End If
            
            ' hilite current player
            If Game.Turn = iPlr Then
                .Font.Size = 10
                .ForeColor = IDC_TITLE_SEL
            Else
                .Font.Size = 8
                .ForeColor = IDC_TITLE
            End If
            
        End With
        
    Next iPlr
    
End Sub
Sub AnimObject(Obj As Object, ByVal x0 As Integer, ByVal y0 As Integer, ByVal x1 As Integer, ByVal y1 As Integer, Optional ByVal iStep As Integer, Optional ByVal iDelay As Integer)
Dim X As Single, Y As Single
Dim xStep As Single, yStep As Single
Dim xDist As Integer, yDist As Integer
    
    If Not Game.Animate Then Exit Sub
    
    DoEvents
    ' set starting coordinates
    Obj.Left = x0
    Obj.Top = y0
    X = x0
    Y = y0
    Obj.Visible = True
    Obj.Refresh
    
    ' set movement variables
    If x0 > x1 Then xStep = -1 Else xStep = 1
    If y0 > y1 Then yStep = -1 Else yStep = 1
    
    If (x1 - x0) > 0 Then xDist = x1 - x0 Else xDist = x0 - x1
    If (y1 - y0) > 0 Then yDist = y1 - y0 Else yDist = y0 - y1
    
    If iStep = 0 Then iStep = 50 / (Game.Speed / 2)
    xStep = xStep * (xDist / iStep)
    If xStep = 0 Then xStep = 1
    yStep = yStep * (yDist / iStep)
    If yStep = 0 Then yStep = 1
    
    If iDelay = 0 Then iDelay = IDT_POP_CARD
    
    ' move object from start to end coordinates
    Do
    
        Delay iDelay
        
        X = X + xStep
        Y = Y + yStep
        
        If xStep < 0 And X < x1 Then X = x1
        If xStep > 0 And X > x1 Then X = x1
        If yStep < 0 And Y < y1 Then Y = y1
        If yStep > 0 And Y > y1 Then Y = y1
        
        Obj.Left = X
        Obj.Top = Y
        
        DoEvents
        
    Loop Until X = x1 And Y = y1

End Sub
Sub DrawDeck(Deck As CardDeck)
Dim iCard As Integer
Dim nCards As Integer
Dim X As Integer, Y As Integer
    
    ' count cards in deck
    nCards = CountCards(Deck)
    
    ' set deck properties
    SetDeckProps Deck
    
    ' draw deck
    With frmMain.picDeck(Deck.Index)
        .Cls
        If nCards = 0 Then
            DrawCard frmMain.picDeck(Deck.Index), Deck, iCard, nCards, X, Y
        Else
            For iCard = 0 To nCards - 1
                GetCardXY Deck, iCard, nCards, X, Y
                DrawCard frmMain.picDeck(Deck.Index), Deck, iCard, nCards, X, Y
            Next iCard
        End If
        .Refresh
    End With
    
End Sub


Sub DrawAllDecks()
Dim iDeck As Integer
    For iDeck = 0 To MAX_DECKS - 1
        DrawDeck Deck(iDeck)
    Next iDeck
End Sub
