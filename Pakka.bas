Attribute VB_Name = "Pakka"
Option Explicit

Public Const MAX_CARDS = 52

' card deck struct
Type CardDeck
    Card(0 To MAX_CARDS) As Integer
    Mode(0 To MAX_CARDS) As Integer
    X As Integer
    Y As Integer
    Index As Integer
    Name As String
End Type

' card mode constants
Public Const cmNormal = 0
Public Const cmSelected = 1
Public Const cmHilite = 2
Public Const cmHidden = 3

' card rank name constants
Public Const A = 1
Public Const T = 10
Public Const J = 11
Public Const Q = 12
Public Const K = 13

' card sorting orders
Public Const csoNone = 0
Public Const csoRank = 1
Public Const csoSuit = 2
Public Const csoRank2TA = 3
Public Const csoSuitAis14 = 4
Function CardInDeck(Deck As CardDeck, ByVal Rank As Integer) As Boolean
Dim iCard As Integer
    For iCard = 0 To CountCards(Deck)
        If GetRank(Deck.Card(iCard)) = Rank Then
            CardInDeck = True
            Exit Function
        End If
    Next iCard
End Function
Function CardOnlyInDeck(Deck As CardDeck, ByVal Rank As Integer) As Boolean
Dim iRank As Integer
    If Not CardInDeck(Deck, Rank) Then
        Exit Function
    End If
    
    CardOnlyInDeck = True
    
    For iRank = A To K
        If CardInDeck(Deck, iRank) And iRank <> Rank Then
            CardOnlyInDeck = False
            Exit Function
        End If
    Next iRank
End Function
Function CardOnlyInDeck2(Deck As CardDeck, ByVal Rank As Integer) As Boolean
Dim iRank As Integer, iRankOld As Integer
Dim iCard As Integer
    If Not CardInDeck(Deck, Rank) Then
        Exit Function
    End If
        
    CardOnlyInDeck2 = True
    
    For iCard = 0 To CountCards(Deck) - 1
        iRank = GetRank(Deck.Card(iCard))
        If iRank <> Rank Then
            If iRankOld = 0 Then
                iRankOld = iRank
            Else
                If iRank <> iRankOld Then
                    CardOnlyInDeck2 = False
                    Exit Function
                End If
            End If
        End If
    Next iCard
End Function
Sub ClearDeck(Deck As CardDeck)
Dim iCard As Integer
    For iCard = 0 To MAX_CARDS
        Deck.Card(iCard) = -1
        Deck.Mode(iCard) = cmNormal
    Next iCard
End Sub
Function CopyDeck(Source As CardDeck) As CardDeck
Dim iCard As Integer
    With CopyDeck
        For iCard = 0 To MAX_CARDS - 1
            .Card(iCard) = Source.Card(iCard)
            .Mode(iCard) = Source.Mode(iCard)
        Next iCard
        .Index = Source.Index
        .Name = Source.Name
        .X = Source.X
        .Y = Source.Y
    End With
End Function
Function CountSelCards(Deck As CardDeck, Optional ByVal Rank As Integer) As Integer
Dim iCard As Integer
    Do While Deck.Card(iCard) <> -1
        If Deck.Mode(iCard) = cmSelected Then
            If Rank = 0 Then
                CountSelCards = CountSelCards + 1
            Else
                If GetRank(Deck.Card(iCard)) = Rank Then
                    CountSelCards = CountSelCards + 1
                End If
            End If
        End If
        iCard = iCard + 1
    Loop
End Function

Function GetCardName(ByVal Rank As Integer) As String
    Select Case Rank
    Case A
        GetCardName = IDS_CARD_ACE
    Case 2
        GetCardName = IDS_CARD_DEUX
    Case 3
        GetCardName = IDS_CARD_THREE
    Case 4
        GetCardName = IDS_CARD_FOUR
    Case 5
        GetCardName = IDS_CARD_FIVE
    Case 6
        GetCardName = IDS_CARD_SIX
    Case 7
        GetCardName = IDS_CARD_SEVEN
    Case 8
        GetCardName = IDS_CARD_EIGHT
    Case 9
        GetCardName = IDS_CARD_NINE
    Case 10
        GetCardName = IDS_CARD_TEN
    Case J
        GetCardName = IDS_CARD_JACK
    Case Q
        GetCardName = IDS_CARD_QUEEN
    Case K
        GetCardName = IDS_CARD_KING
    End Select
End Function
Sub ClearAllDecks()
Dim iDeck As Integer
    For iDeck = 0 To MAX_DECKS - 1
        ClearDeck Deck(iDeck)
        Deck(iDeck).Index = iDeck
    Next iDeck
End Sub
Function GetSuiteName(ByVal Suite As Integer) As String
    Select Case Suite
    Case suClub
        GetSuiteName = IDS_CARD_CLUB
    Case suDiamond
        GetSuiteName = IDS_CARD_DIAMOND
    Case suHeart
        GetSuiteName = IDS_CARD_HEART
    Case suSpade
        GetSuiteName = IDS_CARD_SPADE
    End Select
End Function
Sub InitDecks()
Dim iDeck As Integer
    For iDeck = 0 To MAX_DECKS - 1
        Deck(iDeck).Index = iDeck
    Next iDeck
End Sub


Function CountCards(Deck As CardDeck, Optional ByVal Rank As Integer) As Integer
Dim iCard As Integer
    Do While Deck.Card(iCard) <> -1 ' And iCard < MAX_CARDS
        iCard = iCard + 1
        If Rank = 0 Then
            CountCards = CountCards + 1
        Else
            If GetRank(Deck.Card(iCard)) = Rank Then
                CountCards = CountCards + 1
            End If
        End If
    Loop
End Function
Function GetRank(ByVal Card As Integer) As Integer
    GetRank = Int(Card / 4) + 1
End Function
Function GetSuite(ByVal Card As Integer) As Integer
    Select Case Card
    Case cdAClubs, cd2Clubs, cd3Clubs, cd4Clubs, cd5Clubs, cd6Clubs, cd7Clubs, cd8Clubs, cd9Clubs, cdTClubs, cdJClubs, cdQClubs, cdKClubs
        GetSuite = suClub
    Case cdADiamonds, cd2Diamonds, cd3Diamonds, cd4Diamonds, cd5Diamonds, cd6Diamonds, cd7Diamonds, cd8Diamonds, cd9Diamonds, cdTDiamonds, cdJDiamonds, cdQDiamonds, cdKDiamonds
        GetSuite = suDiamond
    Case cdAHearts, cd2Hearts, cd3Hearts, cd4Hearts, cd5Hearts, cd6Hearts, cd7Hearts, cd8Hearts, cd9Hearts, cdTHearts, cdJHearts, cdQHearts, cdKHearts
        GetSuite = suHeart
    Case cdASpades, cd2Spades, cd3Spades, cd4Spades, cd5Spades, cd6Spades, cd7Spades, cd8Spades, cd9Spades, cdTSpades, cdJSpades, cdQSpades, cdKSpades
        GetSuite = suSpade
    Case Else
        GetSuite = -1
    End Select
End Function
Function GetSelSuite(Deck As CardDeck)
Dim nCards As Integer
Dim iCard As Integer
    GetSelSuite = -1
    nCards = CountCards(Deck)
    For iCard = 0 To nCards - 1
        If Deck.Mode(iCard) = cmSelected Then
            GetSelSuite = GetSuite(Deck.Card(iCard))
            Exit For
        End If
    Next iCard
End Function
Function GetTopCard(Deck As CardDeck, Optional ByVal Index As Integer) As Integer
    If CountCards(Deck) > 0 Then
        GetTopCard = Deck.Card(CountCards(Deck) - 1 - Index)
    Else
        GetTopCard = -1
    End If
End Function
Function IsCardArray(Deck As CardDeck, ByVal Rank As Integer, ByVal Count As Integer) As Boolean
Dim iCard As Integer
Dim nCards As Integer
    If Count > CountCards(Deck) Then
        Exit Function
    End If
    For iCard = 0 To Count - 1
        If GetRank(GetTopCard(Deck, iCard)) = Rank Then
            nCards = nCards + 1
        End If
    Next iCard
    If nCards = Count Then
        IsCardArray = True
    End If
End Function
Function IsPlayerDeck(Deck As CardDeck) As Boolean
    If Deck.Index < MAX_PLAYERS Then
        IsPlayerDeck = True
    End If
End Function

Sub FillPlayerHand(Player As CardDeck, Dealer As CardDeck, CardsInHand As Integer)
Dim nCardsTake As Integer, nCardsPlr As Integer, nCardsDlr As Integer
    nCardsPlr = CountCards(Player)
    nCardsDlr = CountCards(Dealer)
    
    If nCardsPlr < CardsInHand Then
        If nCardsDlr > 0 Then
            nCardsTake = CardsInHand - nCardsPlr
            If nCardsTake > nCardsDlr Then
                nCardsTake = nCardsDlr
            End If
            AnimPopCards Dealer, Player, nCardsTake, True
        End If
    End If
End Sub

Sub MoveCards(Source As CardDeck, Dest As CardDeck, ByVal Mode As Integer, Optional ByVal Rank As Integer)
Dim nCards As Integer
Dim iCard As Integer, iCardTo As Integer, iCardMove As Integer
    nCards = CountCards(Source)
    iCardTo = CountCards(Dest)
    For iCard = nCards To 0 Step -1
        If Source.Mode(iCard) = Mode And (GetRank(Source.Card(iCard)) = Rank Or Rank = 0) Then
            Dest.Card(iCardTo) = Source.Card(iCard)
            Dest.Mode(iCardTo) = Source.Mode(iCardTo)
            
            iCardTo = iCardTo + 1
            For iCardMove = iCard To MAX_CARDS - 1
                Source.Card(iCardMove) = Source.Card(iCardMove + 1)
                Source.Mode(iCardMove) = Source.Mode(iCardMove + 1)
            Next iCardMove
        End If
    Next iCard
End Sub
Sub MoveSelCards(Source As CardDeck, Dest As CardDeck, Optional ByVal iRankFirst As Integer)
    If iRankFirst > 0 Then
        MoveCards Source, Dest, cmSelected, iRankFirst
    End If
    MoveCards Source, Dest, cmSelected
End Sub
Sub SortDeck(Deck As CardDeck, Optional ByVal SortStyle As Integer = -1)
Dim iSorted As Integer
Dim Deck2 As CardDeck

    Deck2 = CopyDeck(Deck)
    ClearDeck Deck2
    
    If SortStyle = -1 Then
        SortStyle = Game.CardSortOrder
    End If
    
    Select Case SortStyle
    Case csoNone
        Exit Sub
    
    Case csoRank
        SortDeck2 Deck, Deck2, iSorted, cdAClubs, cdKSpades
    
    Case csoSuit
        SortDeck2 Deck, Deck2, iSorted, cdAClubs, cdKClubs, 4
        SortDeck2 Deck, Deck2, iSorted, cdADiamonds, cdKDiamonds, 4
        SortDeck2 Deck, Deck2, iSorted, cdAHearts, cdKHearts, 4
        SortDeck2 Deck, Deck2, iSorted, cdASpades, cdKSpades, 4
        
    Case csoSuitAis14
        SortDeck2 Deck, Deck2, iSorted, cd2Clubs, cdKClubs, 4
        SortDeck2 Deck, Deck2, iSorted, cdAClubs, cdAClubs
        SortDeck2 Deck, Deck2, iSorted, cd2Diamonds, cdKDiamonds, 4
        SortDeck2 Deck, Deck2, iSorted, cdADiamonds, cdADiamonds
        SortDeck2 Deck, Deck2, iSorted, cd2Hearts, cdKHearts, 4
        SortDeck2 Deck, Deck2, iSorted, cdAHearts, cdAHearts
        SortDeck2 Deck, Deck2, iSorted, cd2Spades, cdKSpades, 4
        SortDeck2 Deck, Deck2, iSorted, cdASpades, cdASpades
    
    Case csoRank2TA
        SortDeck2 Deck, Deck2, iSorted, cd3Clubs, cd9Spades
        SortDeck2 Deck, Deck2, iSorted, cdJClubs, cdKSpades
        SortDeck2 Deck, Deck2, iSorted, cd2Clubs, cd2Spades
        SortDeck2 Deck, Deck2, iSorted, cdTClubs, cdTSpades
        SortDeck2 Deck, Deck2, iSorted, cdAClubs, cdASpades
    
    End Select
    
    Deck2.Index = Deck.Index
    Deck = Deck2
End Sub
Sub PopCards(Source As CardDeck, Dest As CardDeck, ByVal Count As Integer)
Dim nCards As Integer
Dim iCard As Integer, iCardTo As Integer
    nCards = CountCards(Source)
    If Count > nCards Then
        Exit Sub
    End If
    iCardTo = CountCards(Dest)
    For iCard = nCards - 1 To nCards - Count Step -1
        Dest.Card(iCardTo) = Source.Card(iCard)
        Dest.Mode(iCardTo) = Source.Mode(iCard)
        Source.Card(iCard) = -1
        Source.Mode(iCard) = cmNormal
        iCardTo = iCardTo + 1
    Next iCard
End Sub
Function SelCards(Deck As CardDeck, Optional ByVal Rank As Integer, Optional ByVal Count As Integer, Optional ByVal DontSelectCards As Boolean)
Dim iCard As Integer
Dim nCards As Integer, nSel As Integer
    nCards = CountCards(Deck)
    For iCard = 0 To nCards
        If Deck.Mode(iCard) <> cmSelected And (GetRank(Deck.Card(iCard)) = Rank Or Rank = 0) Then
            If Not DontSelectCards Then
                Deck.Mode(iCard) = cmSelected
            End If
            nSel = nSel + 1
            SelCards = nSel
            If nSel = Count Then
                Exit Function
            End If
        End If
    Next iCard
End Function
Function SelCardsRange(Deck As CardDeck, ByVal iRankLow As Integer, ByVal iRankHigh As Integer, Optional ByVal nCards As Integer) As Boolean
Dim iRank As Integer
    For iRank = iRankLow To iRankHigh
        If CardInDeck(Deck, iRank) Then
            SelCardsRange = True
            SelCards Deck, iRank, nCards
            Exit Function
        End If
    Next iRank
End Function
Sub ShuffleDeck(Deck As CardDeck)
Dim iCard As Integer, iPos As Integer
    ClearDeck Deck
    For iCard = cdAClubs To cdKSpades
        Do
            iPos = Int(MAX_CARDS * Rnd)
        Loop Until Deck.Card(iPos) = -1
        Deck.Card(iPos) = iCard
    Next iCard
End Sub
Sub OldSortDeck(Deck As CardDeck)
Dim iCard As Integer, iCard2 As Integer, iSorted As Integer
Dim nCards As Integer
Dim Deck2 As CardDeck

    nCards = CountCards(Deck)
    Deck2 = CopyDeck(Deck)
    ClearDeck Deck2
    
    Select Case Game.CardSortOrder
    Case csoNone
        Exit Sub
    
    Case csoRank
        For iCard = cdAClubs To cdKSpades
            For iCard2 = 0 To nCards
                If Deck.Card(iCard2) = iCard Then
                    Deck2.Card(iSorted) = iCard
                    iSorted = iSorted + 1
                End If
            Next iCard2
        Next iCard
    
    Case csoSuit, csoSuitAis14
        For iCard = cdAClubs To cdKClubs Step 4
            For iCard2 = 0 To nCards
                If Deck.Card(iCard2) = iCard Then
                    Deck2.Card(iSorted) = iCard
                    iSorted = iSorted + 1
                End If
            Next iCard2
        Next iCard
        For iCard = cdADiamonds To cdKDiamonds Step 4
            For iCard2 = 0 To nCards
                If Deck.Card(iCard2) = iCard Then
                    Deck2.Card(iSorted) = iCard
                    iSorted = iSorted + 1
                End If
            Next iCard2
        Next iCard
        For iCard = cdAHearts To cdKHearts Step 4
            For iCard2 = 0 To nCards
                If Deck.Card(iCard2) = iCard Then
                    Deck2.Card(iSorted) = iCard
                    iSorted = iSorted + 1
                End If
            Next iCard2
        Next iCard
        For iCard = cdASpades To cdKSpades Step 4
            For iCard2 = 0 To nCards
                If Deck.Card(iCard2) = iCard Then
                    Deck2.Card(iSorted) = iCard
                    iSorted = iSorted + 1
                End If
            Next iCard2
        Next iCard
        
    Case csoRank2TA
        For iCard = cd3Clubs To cd9Spades
            For iCard2 = 0 To nCards
                If Deck.Card(iCard2) = iCard Then
                    Deck2.Card(iSorted) = iCard
                    iSorted = iSorted + 1
                End If
            Next iCard2
        Next iCard
        For iCard = cdJClubs To cdKSpades
            For iCard2 = 0 To nCards
                If Deck.Card(iCard2) = iCard Then
                    Deck2.Card(iSorted) = iCard
                    iSorted = iSorted + 1
                End If
            Next iCard2
        Next iCard
        For iCard = cd2Clubs To cd2Spades
            For iCard2 = 0 To nCards
                If Deck.Card(iCard2) = iCard Then
                    Deck2.Card(iSorted) = iCard
                    iSorted = iSorted + 1
                End If
            Next iCard2
        Next iCard
        For iCard = cdTClubs To cdTSpades
            For iCard2 = 0 To nCards
                If Deck.Card(iCard2) = iCard Then
                    Deck2.Card(iSorted) = iCard
                    iSorted = iSorted + 1
                End If
            Next iCard2
        Next iCard
        For iCard = cdAClubs To cdASpades
            For iCard2 = 0 To nCards
                If Deck.Card(iCard2) = iCard Then
                    Deck2.Card(iSorted) = iCard
                    iSorted = iSorted + 1
                End If
            Next iCard2
        Next iCard
    
    End Select
    
    Deck2.Index = Deck.Index
    Deck = Deck2
    
End Sub
Sub SortDeck2(Deck As CardDeck, Deck2 As CardDeck, ByRef iSorted As Integer, ByVal iCardFirst As Integer, ByVal iCardLast As Integer, Optional ByVal iCardStep As Integer = 1)
Dim iCard As Integer, iCard2 As Integer
    For iCard = iCardFirst To iCardLast Step iCardStep
        For iCard2 = 0 To CountCards(Deck)
            If Deck.Card(iCard2) = iCard Then
                Deck2.Card(iSorted) = iCard
                iSorted = iSorted + 1
            End If
        Next iCard2
    Next iCard
End Sub
Function IsSuiteInDeck(Deck As CardDeck, ByVal Suite As Integer) As Boolean
Dim iCard As Integer
    For iCard = 0 To CountCards(Deck)
        If GetSuite(Deck.Card(iCard)) = Suite Then
            IsSuiteInDeck = True
            Exit Function
        End If
    Next iCard
End Function

Sub UnSelCards(Deck As CardDeck, Optional ByVal Rank As Integer)
Dim iCard As Integer
    For iCard = 0 To CountCards(Deck)
        If GetRank(Deck.Card(iCard)) = Rank Or Rank = 0 Then
            Deck.Mode(iCard) = cmNormal
        End If
    Next iCard
End Sub
