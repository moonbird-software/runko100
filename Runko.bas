Attribute VB_Name = "Runko100"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_NODEFAULT = &H2
Public Const SND_SYNC = &H0
Public Const SND_NOSTOP = &H10
Public Const SND_MEMORY = &H4

' applications
'Public Const IDA_PASKAHOUSU = 0
'Public Const IDA_RISTISEISKA = 1
'Public Const IDA_MUSTAMAIJA = 2

' cards32.dll
Public cdWidth As Long
Public cdHeight As Long

' help constants
Const HELP_TAB = &HF

' language constants
Public Const IDL_FINNISH = 0
Public Const IDL_ENGLISH = 1

' key constants
Public Const IDK_TRYFROMDECK = 167
Public Const IDK_PICKUP = 43

' sound constants
Public Const IDSND_CARDCLICK = 0
Public Const IDSND_CARDMOVE = 1
Public Const IDSND_CARDDROP = 2
Public Const IDSND_KOSH = 3
Public Const IDSND_GAMEOVER = 4

' player mode constants
Public Const IDP_HUMAN = 0
Public Const IDP_CPU = 1
Public Const IDP_NETWORK = 2

' game data type
Type GameData
    Demo As Boolean
    Debug As Boolean
    NetGame As Boolean
    New As Boolean
    Over As Boolean
    Inited As Boolean
    Dealing As Boolean
    On As Boolean
    
    Turn As Integer
    Dealer As Integer
    NextPos As Integer
    RoundNbr As Integer
    FirstTurn As Boolean
    FirstCard As Integer
    Mode As Integer
    Trump As Integer
    
    CardBack As Integer
    Animate As Boolean
    QuickDeal As Boolean
    Language As Integer
    Sound As Boolean
    Speed As Integer
    AutoStart As Boolean
    ShowScore As Boolean
    CardSortOrder As Integer
    ServerName As String
    PrivMsg As Boolean
    Counter As Integer
    Opener As Integer
    
    Score(0 To MAX_PLAYERS - 1) As Integer
    Value(0 To MAX_PLAYERS - 1) As Integer
    Value2(0 To MAX_PLAYERS - 1) As Integer
    AI(0 To MAX_PLAYERS - 1) As Integer
    Pos(0 To MAX_PLAYERS - 1) As Integer
    Title(0 To MAX_DECKS - 1) As String
    IP(0 To MAX_PLAYERS - 1) As String
    PlrMode(0 To MAX_PLAYERS - 1) As Integer
End Type

' game data
Public Deck(0 To MAX_DECKS - 1) As CardDeck
Public Game As GameData
Public Rules As RuleBook
Sub FormSettingsCardBackClick()
    With frmSettings
        cdtDraw .picDeck.hdc, cdWidth / 4, 0, cdASpades, mdFaceUp, IDC_TABLEBG
        cdtDraw .picDeck.hdc, 0, 0, .lstDeck.ListIndex + cdPlaid, mdFaceDown, IDC_TABLEBG
        .picDeck.Refresh
    End With
End Sub
Sub FormSettingsLoadBasic()
Dim iPlr As Integer, iAI As Integer
    With frmSettings
        .picSettings(0).ZOrder 0
        
        ' dynamic controls
        .fraRules(RUNKO_APP).Left = 120
        .fraRules(RUNKO_APP).Top = 0
        .fraRules(RUNKO_APP).Visible = True
        .fraPerformance.Left = 120
        .fraPerformance.Top = .fraRules(RUNKO_APP).Height + 105
        
        ' player names and AI
        For iPlr = 0 To MAX_PLAYERS - 1
            .txtName(iPlr).Text = Game.Title(iPlr)
            For iAI = 1 To 5
                .comAI(iPlr).AddItem iAI
            Next iAI
            .comAI(iPlr).ListIndex = Game.AI(iPlr)
        Next iPlr
        
        ' card back selection
        .lstDeck.AddItem IDS_CARD_BACK_0
        .lstDeck.AddItem IDS_CARD_BACK_1
        .lstDeck.AddItem IDS_CARD_BACK_2
        .lstDeck.AddItem IDS_CARD_BACK_3
        .lstDeck.AddItem IDS_CARD_BACK_4
        .lstDeck.AddItem IDS_CARD_BACK_5
        .lstDeck.AddItem IDS_CARD_BACK_6
        .lstDeck.AddItem IDS_CARD_BACK_7
        .lstDeck.AddItem IDS_CARD_BACK_8
        .lstDeck.AddItem IDS_CARD_BACK_9
        .lstDeck.AddItem IDS_CARD_BACK_10
        .lstDeck.AddItem IDS_CARD_BACK_11
        .lstDeck.ListIndex = Game.CardBack - cdPlaid
        
        ' performance
        .sldSpeed.Value = Game.Speed
        .chkCardAnim.Value = -Game.Animate
        .chkQuickDeal.Value = -Game.QuickDeal
        .chkAutoRestart.Value = -Game.AutoStart
        .chkShowScore.Value = -Game.ShowScore
    End With
End Sub
Sub FormSettingsSaveBasic()
Dim iPlr As Integer

    With frmSettings
        ' player names and AI
        For iPlr = 0 To MAX_PLAYERS - 1
            Game.Title(iPlr) = .txtName(iPlr).Text
            If Game.Title(iPlr) = "" Then
                Game.Title(iPlr) = IDS_PLAYER & " " & Format(iPlr + 1)
            End If
            Game.AI(iPlr) = .comAI(iPlr).ListIndex
        Next iPlr
        If frmMain.Visible Then
            DrawTitles
        End If
        
        ' card back selection
        Game.CardBack = .lstDeck.ListIndex + cdPlaid
        
        ' performance
        Game.Speed = .sldSpeed.Value
        Game.Animate = -.chkCardAnim.Value
        Game.QuickDeal = -.chkQuickDeal.Value
        Game.AutoStart = -.chkAutoRestart.Value
        Game.ShowScore = -.chkShowScore.Value
    End With
    
    SaveSettings
    
    Unload frmSettings
    
End Sub
Sub GameNetwork()
    'frmNetwork.Show vbModal
End Sub
Function GetPrevValidPlayerBasic(ByVal iPlr As Integer) As Integer
Dim iPanic As Integer
    Do
        iPlr = GetPrevPlayer(iPlr)
        iPanic = iPanic + 1
    Loop Until CountCards(Deck(iPlr)) > 0 Or iPanic = MAX_PLAYERS
    GetPrevValidPlayerBasic = iPlr
End Function
Sub InitLocaleBasic()
Dim iPlr As Integer

    ' get locale from Locale_xx.bas
    SetGameLanguage

    ' main form
    With frmMain
        .mnuGame.Caption = IDS_MENU_GAME
        .mnuGameNew.Caption = IDS_MENU_GAME_NEW
        .mnuGameNetwork.Caption = IDS_MENU_GAME_NETWORK
        .mnuGameSettings.Caption = IDS_MENU_GAME_SETTINGS
        .mnuGameScore.Caption = IDS_MENU_GAME_SCORE
        .mnuGameSound.Caption = IDS_MENU_GAME_SOUND
        .mnuGameDemo.Caption = IDS_MENU_GAME_DEMO
        .mnuGameExit.Caption = IDS_MENU_GAME_EXIT
        .mnuHelp.Caption = IDS_MENU_HELP
        .mnuHelpContents.Caption = IDS_MENU_HELP_CONTENTS
        .mnuHelpAbout.Caption = IDS_MENU_HELP_ABOUT
    End With
    
    ' settings
    With frmSettings
        .tabSettings.Tabs(1).Caption = IDS_DLG_SETTINGS_TAB_GENERAL
        .tabSettings.Tabs(2).Caption = IDS_DLG_SETTINGS_TAB_ADVANCED
    
        .fraPlayers.Caption = IDS_DLG_SETTINGS_PLAYERS
        For iPlr = 0 To MAX_PLAYERS - 1
            .lblName(iPlr).Caption = IDS_DLG_SETTINGS_PLAYER & " " & Format(iPlr + 1) & ":"
        Next iPlr
        
        .fraDeck.Caption = IDS_DLG_SETTINGS_DECK
        .lblDeck.Caption = IDS_DLG_SETTINGS_DECK_BACK
        
        .fraRules(0).Caption = IDS_DLG_SETTINGS_RULES
        .fraRules(1).Caption = IDS_DLG_SETTINGS_RULES
        
        .fraPerformance.Caption = IDS_DLG_SETTINGS_PERFORMANCE
        .lblSpeed.Caption = IDS_DLG_SETTINGS_GAME_SPEED
        .chkAutoRestart.Caption = IDS_DLG_SETTINGS_AUTOSTART
        .chkCardAnim.Caption = IDS_DLG_SETTINGS_ANIM_CARDS
        .chkQuickDeal.Caption = IDS_DLG_SETTINGS_QUICK_DEAL
        .chkShowScore.Caption = IDS_DLG_SETTINGS_SHOW_SCORE
        
        .cmdOK.Caption = IDS_OK
        .cmdCancel.Caption = IDS_CANCEL
    End With
End Sub

Function MakeWordWhose(ByVal sName As String) As String
    sName = MakeWordBend(sName)
    
    Select Case Game.Language
    Case IDL_FINNISH
        sName = sName & "n"
    Case IDL_ENGLISH
        If Right(sName, 1) = "s" Then
            sName = sName & "'"
        Else
            sName = sName & "'s"
        End If
    End Select
    
    MakeWordWhose = sName
End Function
Function MakeWordBend(ByVal sName As String) As String
    Select Case Game.Language
    Case IDL_FINNISH
        Select Case Left(Right(sName, 3), 2)
        Case "kk", "pp", "tt"
            sName = Left(sName, Len(sName) - 2) & Right(sName, 1)
        End Select
        
        Select Case Right(sName, 3)
        Case "tar"
            sName = Left(sName, Len(sName) - 3) & "ttare"
        Case "gas"
            sName = Left(sName, Len(sName) - 3) & "kaa"
        Case "nen"
            sName = Left(sName, Len(sName) - 3) & "se"
        Case "tys"
            sName = Left(sName, Len(sName) - 3) & "tykse"
        Case "yys"
            sName = Left(sName, Len(sName) - 3) & "yyde"
        Case "uus"
            sName = Left(sName, Len(sName) - 3) & "uude"
        End Select
        
        Select Case Right(sName, 2)
        Case "us"
            sName = Left(sName, Len(sName) - 2) & "ukse"
        Case "os"
            sName = Left(sName, Len(sName) - 2) & "okse"
'        Case "ke"
'            sName = Left(sName, Len(sName) - 2) & "kkee"
'        Case "ki"
'            sName = Left(sName, Len(sName) - 2) & "e"
        End Select
        
        Select Case Right(sName, 1)
        Case "b", "c", "d", "f", "g", "h", "k", "l", "m", "n", "p", "q", "r", "s", "t", "v", "w", "x", "z"
            sName = sName & "i"
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9
            sName = sName & ":"
        End Select
    End Select
    
    MakeWordBend = sName
End Function
Sub PlayerInputBasic(Enabled As Boolean)
    With frmMain
        If Enabled = True Then
            .MousePointer = vbNormal
        Else
            .MousePointer = vbHourglass
        End If
        
        .picDeck(IDD_USER).Enabled = Enabled
        If IDD_DEALER > 0 Then
            .picDeck(IDD_DEALER).Enabled = Enabled
        End If
        If IDD_TRICK > 0 Then
            .picDeck(IDD_TRICK).Enabled = Enabled
        End If
    End With
End Sub
Sub SetDeckNames()
Dim iPlr As Integer
    
    ' players
    For iPlr = 0 To MAX_PLAYERS - 1
        Deck(iPlr).Name = Game.Title(iPlr)
    Next iPlr
    
    ' basic decks
    Deck(IDD_DEALER).Name = IDS_DEALER
    If IDD_TRICK > 0 Then
        Deck(IDD_TRICK).Name = IDS_TRICK
    End If
    
    If IDD_TRASH > 0 Then
        Deck(IDD_TRASH).Name = IDS_TRASH
    End If
    
End Sub
Sub SetNextTurn(ByVal iPlr As Integer)
Dim nCards As Integer
    ' find next player with cards left
    iPlr = GetNextValidPlayer(iPlr)
    Game.Turn = iPlr
    
    ' animate changes
    DrawTitles
    
    If Game.Turn = IDD_USER Then
        SetStatus IDS_STATUS_CHOOSE_CARDS
    Else
        SetStatus Replace(IDS_STATUS_TURN, "%s", MakeWordWhose(Game.Title(iPlr)))
    End If
    
    DoEvents
    
    Delay IDT_MOVE_CARD
End Sub
Sub DealCards()
Dim iCard As Integer, iPlr As Integer
Dim fAnim As Boolean
    
    Debug.Print "Jaetaan kortit..."
    
    ' quick deal?
    With Game
        .On = True
        .Dealing = True
        fAnim = .Animate
        If .QuickDeal Then
            .Animate = False
        End If
    End With
    
    ' set status
    SetStatus IDS_STATUS_DEALING
    PlayerInput False
    With frmMain
        .mnuGameNew.Enabled = False
        .mnuGameNetwork.Enabled = False
        .mnuGameDemo.Enabled = False
        .mnuGameSettings.Enabled = False
    End With
    
    ' clear all decks and redraw
    ClearAllDecks
    DrawTitles
    DrawAllDecks
    DoEvents
    
    ' shuffle and draw deck
    ShuffleDeck Deck(IDD_DEALER)
    DrawDeck Deck(IDD_DEALER)
    DoEvents
    Delay IDT_MOVE_CARD
    
    ' do own hook routine
    DealCardsHook
    
    ' deal cards
    iPlr = GetNextPlayer(Game.Dealer)
    Game.Turn = Game.Dealer
    
    For iCard = 1 To (GetCardsInHand * MAX_PLAYERS)
        DrawTitles
        DoEvents
        PlaySound 1
        AnimPopCards Deck(IDD_DEALER), Deck(iPlr), 1, True
        iPlr = GetNextPlayer(iPlr)
    Next iCard
    
    ' set status
    With frmMain
        .mnuGameNew.Enabled = True
        .mnuGameNetwork.Enabled = True
        .mnuGameDemo.Enabled = True
        .mnuGameSettings.Enabled = True
    End With

    Game.Animate = fAnim
    Game.Dealing = False
    SetStatus
    
    ' do own hook routine
    DealCardsDoneHook

End Sub
Function MakeWordFrom(ByVal sName As String) As String
    sName = MakeWordBend(sName)
    
    Select Case Game.Language
    Case IDL_FINNISH
        sName = sName & "lta"
    End Select
    
    MakeWordFrom = sName
End Function
Function MakeWordTo(ByVal sName As String) As String
    sName = MakeWordBend(sName)
    
    Select Case Game.Language
    Case IDL_FINNISH
        sName = sName & "lle"
    End Select
    
    MakeWordTo = sName
End Function
Sub PlaySound(SoundIndex As Integer)
Dim SoundArray() As Byte
Dim sSndRes As String
Dim uFlags As Integer

    If Not Game.Sound Then
        Exit Sub
    End If
      
    uFlags = SND_ASYNC
    
    Select Case SoundIndex
    Case IDSND_CARDCLICK
        sSndRes = "IDSND_CARDCLICK"
    Case IDSND_CARDMOVE
        sSndRes = "IDSND_CARDMOVE"
    Case IDSND_CARDDROP
        sSndRes = "IDSND_CARDDROP"
    Case IDSND_KOSH
        sSndRes = "IDSND_KOSH"
        uFlags = SND_SYNC
    Case IDSND_GAMEOVER
        sSndRes = "IDSND_GAMEOVER"
        uFlags = SND_SYNC
    End Select
    
    SoundArray = LoadResData(sSndRes, "WAVE")
    
    sndPlaySound SoundArray(0), uFlags + SND_NODEFAULT + SND_MEMORY
    
End Sub
Sub ReadSettingsBasic()
Dim iPlr As Integer
    ' player titles
    For iPlr = 0 To MAX_PLAYERS - 1
        Game.Title(iPlr) = GetSetting("Moonbird Software", "Players", "Name" & Format(iPlr), IDS_PLAYER & " " & Format(iPlr + 1))
        Game.AI(iPlr) = GetSetting(App.Title, "Players", "AI" & Format(iPlr), 4)
    Next iPlr
    If Game.Title(0) = IDS_DEBUG_PLR_NAME Then
        Game.Debug = True
    End If
    
    ' settings
    Game.CardBack = GetSetting(App.Title, "Settings", "CardBack", cdPlaid)
    'Game.Language = GetSetting(App.Title, "Settings", "Language", IDL_FINNISH)
    Game.Sound = GetSetting(App.Title, "Settings", "Sound", True)
    Game.Speed = GetSetting(App.Title, "Settings", "Speed", 5)
    Game.Animate = GetSetting(App.Title, "Settings", "Animate", True)
    Game.QuickDeal = GetSetting(App.Title, "Settings", "QuickDeal", False)
    Game.AutoStart = GetSetting(App.Title, "Settings", "AutoStart", True)
    Game.ShowScore = GetSetting(App.Title, "Settings", "ShowScore", True)
    
    ' network
    'Game.ServerName = GetSetting(App.Title, "Network", "ServerName", App.Title)
    'Game.PrivMsg = GetSetting(App.Title, "Network", "PrivateMessages", False)
    
    frmMain.mnuGameSound.Checked = Game.Sound
    
    ' show settings if first startup
    If GetSetting(App.Title, "Settings", "FirstRun", True) Then
        MsgBox Replace(IDS_QUERY_FIRST_RUN, "%s", MakeWordWhose(App.Title)), vbInformation
        InitLocale
        Load frmSettings
        frmSettings.cmdCancel.Enabled = False
        frmSettings.Show vbModal
    End If
    
    SetDeckNames
End Sub
Sub SaveSettingsBasic()
Dim iPlr As Integer
    ' player titles
    For iPlr = 0 To MAX_PLAYERS - 1
        SaveSetting "Moonbird Software", "Players", "Name" & Format(iPlr), Game.Title(iPlr)
        SaveSetting App.Title, "Players", "AI" & Format(iPlr), Game.AI(iPlr)
    Next iPlr
    
    ' settings
    SaveSetting App.Title, "Settings", "CardBack", Game.CardBack
    'SaveSetting App.Title, "Settings", "Language", Game.Language
    SaveSetting App.Title, "Settings", "Sound", Game.Sound
    SaveSetting App.Title, "Settings", "Speed", Game.Speed
    SaveSetting App.Title, "Settings", "Animate", Game.Animate
    SaveSetting App.Title, "Settings", "QuickDeal", Game.QuickDeal
    SaveSetting App.Title, "Settings", "AutoStart", Game.AutoStart
    SaveSetting App.Title, "Settings", "ShowScore", Game.ShowScore
    
    ' network
    'SaveSetting App.Title, "Network", "ServerName", Game.ServerName
    'SaveSetting App.Title, "Network", "PrivateMessages", Game.PrivMsg
    
    ' disable first run
    SaveSetting App.Title, "Settings", "FirstRun", False

End Sub
Sub SetStatus(Optional ByVal sStatus As String, Optional ByVal fOverride As Boolean)
    If (Game.Demo = False Or fOverride = True) And frmMain.sbrStatus.SimpleText <> sStatus Then
        frmMain.sbrStatus.SimpleText = sStatus
    End If
End Sub
Function GetKeyValue(ByVal KeyAscii As Integer) As Integer
    Select Case KeyAscii
    Case vbKey1
        GetKeyValue = 1
    Case vbKey2
        GetKeyValue = 2
    Case vbKey3
        GetKeyValue = 3
    Case vbKey4
        GetKeyValue = 4
    Case vbKey5
        GetKeyValue = 5
    Case vbKey6
        GetKeyValue = 6
    Case vbKey7
        GetKeyValue = 7
    Case vbKey8
        GetKeyValue = 8
    Case vbKey9
        GetKeyValue = 9
    Case vbKey0
        GetKeyValue = 10
    Case vbKeyQ
        GetKeyValue = 11
    Case vbKeyW
        GetKeyValue = 12
    Case vbKeyE
        GetKeyValue = 13
    Case vbKeyR
        GetKeyValue = 14
    Case vbKeyT
        GetKeyValue = 15
    Case vbKeyY
        GetKeyValue = 16
    Case vbKeyU
        GetKeyValue = 17
    Case vbKeyI
        GetKeyValue = 18
    Case vbKeyO
        GetKeyValue = 19
    Case vbKeyP
        GetKeyValue = 20
    Case vbKeyA
        GetKeyValue = 21
    Case vbKeyS
        GetKeyValue = 22
    Case vbKeyD
        GetKeyValue = 23
    Case vbKeyF
        GetKeyValue = 24
    Case vbKeyG
        GetKeyValue = 25
    Case vbKeyH
        GetKeyValue = 26
    Case vbKeyJ
        GetKeyValue = 27
    Case vbKeyK
        GetKeyValue = 28
    Case vbKeyL
        GetKeyValue = 29
    Case vbKeyZ
        GetKeyValue = 30
    Case vbKeyX
        GetKeyValue = 31
    Case vbKeyC
        GetKeyValue = 32
    Case vbKeyV
        GetKeyValue = 33
    Case vbKeyB
        GetKeyValue = 34
    Case vbKeyN
        GetKeyValue = 35
    Case vbKeyM
        GetKeyValue = 36
    End Select
    
    KeyAscii = KeyAscii - 32
    
    Select Case KeyAscii
    Case vbKey1
        GetKeyValue = 1
    Case vbKey2
        GetKeyValue = 2
    Case vbKey3
        GetKeyValue = 3
    Case vbKey4
        GetKeyValue = 4
    Case vbKey5
        GetKeyValue = 5
    Case vbKey6
        GetKeyValue = 6
    Case vbKey7
        GetKeyValue = 7
    Case vbKey8
        GetKeyValue = 8
    Case vbKey9
        GetKeyValue = 9
    Case vbKey0
        GetKeyValue = 10
    Case vbKeyQ
        GetKeyValue = 11
    Case vbKeyW
        GetKeyValue = 12
    Case vbKeyE
        GetKeyValue = 13
    Case vbKeyR
        GetKeyValue = 14
    Case vbKeyT
        GetKeyValue = 15
    Case vbKeyY
        GetKeyValue = 16
    Case vbKeyU
        GetKeyValue = 17
    Case vbKeyI
        GetKeyValue = 18
    Case vbKeyO
        GetKeyValue = 19
    Case vbKeyP
        GetKeyValue = 20
    Case vbKeyA
        GetKeyValue = 21
    Case vbKeyS
        GetKeyValue = 22
    Case vbKeyD
        GetKeyValue = 23
    Case vbKeyF
        GetKeyValue = 24
    Case vbKeyG
        GetKeyValue = 25
    Case vbKeyH
        GetKeyValue = 26
    Case vbKeyJ
        GetKeyValue = 27
    Case vbKeyK
        GetKeyValue = 28
    Case vbKeyL
        GetKeyValue = 29
    Case vbKeyZ
        GetKeyValue = 30
    Case vbKeyX
        GetKeyValue = 31
    Case vbKeyC
        GetKeyValue = 32
    Case vbKeyV
        GetKeyValue = 33
    Case vbKeyB
        GetKeyValue = 34
    Case vbKeyN
        GetKeyValue = 35
    Case vbKeyM
        GetKeyValue = 36
    End Select
End Function
Sub Delay(ByVal Time As Integer)
    If Game.Debug = True Then
        Time = 0
    End If
    Sleep Time / Game.Speed / 8
End Sub
Sub GameInit()
    
    ' init cards32.dll
    If cdtInit(cdWidth, cdHeight) = False Then
        MsgBox IDS_ERROR_CARDS32, vbCritical
        End
    End If
    
    ' read settings from registry
    ReadSettings
    
    ' init locale
    InitLocale
    
    ' init random number generator
    Randomize
    
    ' init display
    InitGameData
    ClearAllDecks
    DrawTitles
    DrawAllDecks
    
    ' show main form
    frmMain.Show
    DoEvents
    
    ' start new game
    If Game.AutoStart Then
        GameNew
    Else
        SetStatus IDS_STATUS_PRESS_F2
    End If

End Sub
Sub GameNew(Optional ByVal CalledFromGameDemo As Boolean)
Dim iPlr As Integer
    
    ' set indicators
    
    PlayerInput False
    SetStatus IDS_STATUS_NEW_GAME
    Game.New = True
    
    ' clear game score
    'If CalledFromGameDemo Then
        For iPlr = 0 To MAX_PLAYERS - 1
            Game.Score(iPlr) = 0
        Next iPlr
        Game.RoundNbr = 1
    'End If
    
    If (CalledFromGameDemo = False And Game.Demo = False And Game.Turn = IDD_USER) Or Game.On = False Then
        RotateTurn
    End If
    
End Sub
Sub GameSettings()
    
    ' init locale strings
    InitLocale
    
    ' show frmSettings
    frmSettings.Show vbModal
    
    ' set status
    FormMainResize
    UpdateDebug
    SetDeckNames
    
    ' start new game if requested by frmSettings
    If Game.New = True And Game.Turn = IDD_USER Then
        SetStatus IDS_STATUS_NEW_GAME
        DoEvents
        RotateTurn
    End If

End Sub
Sub HelpAbout()
Dim sAbout As String
    sAbout = IDS_VERSION & " " & Format(App.Major) & "." & Format(App.Minor) & "." & Format(App.Revision) & vbCrLf & vbCrLf & IDS_COPYRIGHT & vbCrLf & vbCrLf & IDS_URL & vbCrLf & IDS_EMAIL
    MsgBox sAbout, vbInformation
End Sub
Sub HelpContents()
    WinHelp frmMain.hWnd, App.HelpFile, HELP_TAB, 0
End Sub
Sub GameDemo()
    ' set demo flag
    frmMain.mnuGameDemo.Checked = Not frmMain.mnuGameDemo.Checked
    Game.Demo = frmMain.mnuGameDemo.Checked

    ' start new game
    GameNew True
    
    ' rotate turn if player in control
    If Game.Demo = True And Game.Turn = IDD_USER Then
        RotateTurn
    End If
End Sub
Sub GameScore(Optional ByVal fShowAnyway As Boolean)
Dim iPlr As Integer
Dim nPtsMax As Integer, nPts As Integer
Dim sScore As String
    
    ' find highest score
    For iPlr = 0 To MAX_PLAYERS - 1
        If Game.Score(iPlr) > nPtsMax Then
            nPtsMax = Game.Score(iPlr)
        End If
    Next iPlr
    
    ' title row
    sScore = IDS_NAME & vbTab & vbTab & IDS_SCORE & vbCrLf & vbCrLf
    
    ' sort scores in descending order
    For nPts = nPtsMax To 0 Step -1
        For iPlr = 0 To MAX_PLAYERS - 1
            If Game.Score(iPlr) = nPts Then
                sScore = sScore & Game.Title(iPlr) & vbTab & vbTab & Format(Game.Score(iPlr)) & vbCrLf
            End If
        Next iPlr
    Next nPts
    
    ' round number
    sScore = sScore & vbCrLf & IDS_ROUND & ": " & Game.RoundNbr
    
    ' display scores
    Debug.Print sScore
    
    If fShowAnyway Or (Game.Debug = False And Game.Demo = False And Game.ShowScore) Then
        MsgBox sScore, vbInformation, IDS_SCOREBOARD
    End If
End Sub
Function GetNextPlayer(ByVal iPlr As Integer) As Integer
    iPlr = iPlr + 1
    If iPlr > MAX_PLAYERS - 1 Then iPlr = 0
    GetNextPlayer = iPlr
End Function
Function GetPrevPlayer(ByVal iPlr As Integer) As Integer
    iPlr = iPlr - 1
    If iPlr < 0 Then iPlr = MAX_PLAYERS - 1
    GetPrevPlayer = iPlr
End Function

Function CountPlayersBasic() As Integer
Dim iPlr As Integer
Dim nPlr As Integer
    nPlr = MAX_PLAYERS
    For iPlr = 0 To MAX_PLAYERS - 1
        If CountCards(Deck(iPlr)) = 0 Then
            nPlr = nPlr - 1
        End If
    Next iPlr
    CountPlayersBasic = nPlr
End Function
Sub ClearScores()
Dim iPlr As Integer
    For iPlr = 0 To MAX_PLAYERS - 1
        Game.Score(iPlr) = 0
    Next iPlr
End Sub
Function IsPlayerCPU(ByVal iPlr As Integer) As Boolean
    If Game.Demo Then
        IsPlayerCPU = True
    Else
        If iPlr <> IDD_USER Then
            IsPlayerCPU = True
        End If
    End If
End Function
Sub RotateTurn()
    
    PlayerInput False
    
    If Game.New Then
        ClearGameData
        DealCards
        GetFirstPlayer
    End If
1:  If AI_Turn Then
        Do
            If Game.New Then
                ClearGameData
                DealCards
                GetFirstPlayer
            End If
        Loop While AI_Turn
    End If
    If Game.Over Then
        GameEnd
    End If
    If Game.New Then
        ClearGameData
        If Game.AutoStart Or Game.Demo Then
            DealCards
            GetFirstPlayer
            GoTo 1
        Else
            ClearAllDecks
            DrawAllDecks
            DrawTitles
            frmMain.MousePointer = vbNormal
            SetStatus IDS_STATUS_PRESS_F2
        End If
    End If
End Sub
Function GetNextValidPlayerBasic(ByVal iPlr As Integer) As Integer
Dim iPanic As Integer
    Do
        iPlr = GetNextPlayer(iPlr)
        iPanic = iPanic + 1
    Loop Until CountCards(Deck(iPlr)) > 0 Or iPanic = MAX_PLAYERS
    GetNextValidPlayerBasic = iPlr
End Function
Sub GameEnd()
Dim iPlr As Integer
Dim nPts As Integer

    PlaySound IDSND_GAMEOVER
    
    Game.On = False
    
    'Debug.Print "Hold Down Counter: ", Game.HoldDown

    ' update player scores
    For iPlr = 0 To MAX_PLAYERS - 1
        If Game.Pos(iPlr) = 0 Then
            Game.Pos(iPlr) = Game.NextPos
        End If
        nPts = MAX_PLAYERS - Game.Pos(iPlr)
        Game.Score(iPlr) = Game.Score(iPlr) + nPts
        Game.Pos(iPlr) = 0
        
        ' the losing player is the next dealer
        If nPts = 0 Then
            Game.Dealer = iPlr
        End If
    Next iPlr
    Game.NextPos = 1
    
    ' show stats
    GameScore
    
    Game.RoundNbr = Game.RoundNbr + 1
    
    ' new game
    Game.New = True
End Sub
Sub GameUninit()
    cdtTerm
    End
End Sub
