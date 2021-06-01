Attribute VB_Name = "RunkoLocFIN"
Option Explicit

' FINNISH LOCALE

' menu strings

Public Const IDS_MENU_GAME = "&Peli"
Public Const IDS_MENU_GAME_NEW = "&Uusi peli"
Public Const IDS_MENU_GAME_NETWORK = "&Verkkopeli..."
Public Const IDS_MENU_GAME_SETTINGS = "&Asetukset..."
Public Const IDS_MENU_GAME_SCORE = "&Pisteet..."
Public Const IDS_MENU_GAME_SOUND = "&Äänet"
Public Const IDS_MENU_GAME_DEMO = "&Demo"
Public Const IDS_MENU_GAME_EXIT = "&Lopeta"
Public Const IDS_MENU_HELP = "&Ohje"
Public Const IDS_MENU_HELP_CONTENTS = "&Ohjeen aiheet"
Public Const IDS_MENU_HELP_ABOUT = "&Tietoja..."

' dialog strings

Public Const IDS_OK = "OK"
Public Const IDS_CANCEL = "Peruuta"

Public Const IDS_DLG_SETTINGS = "Asetukset"
Public Const IDS_DLG_SETTINGS_TAB_GENERAL = "Yleiset"
Public Const IDS_DLG_SETTINGS_TAB_ADVANCED = "Lisäasetukset"

Public Const IDS_DLG_SETTINGS_PLAYERS = "Pelaajien nimet ja tekoäly"
Public Const IDS_DLG_SETTINGS_PLAYER = "Pelaaja"

Public Const IDS_DLG_SETTINGS_DECK = "Pakka"
Public Const IDS_DLG_SETTINGS_DECK_BACK = "Valitse pakan taustakuva:"

Public Const IDS_DLG_SETTINGS_RULES = "Säännöt"

Public Const IDS_DLG_SETTINGS_PERFORMANCE = "Suorituskyky"
Public Const IDS_DLG_SETTINGS_GAME_SPEED = "Pelin nopeus:"
Public Const IDS_DLG_SETTINGS_ANIM_CARDS = "Animoi korttien siirrot"
Public Const IDS_DLG_SETTINGS_QUICK_DEAL = "Pikajako"
Public Const IDS_DLG_SETTINGS_AUTOSTART = "Aloita uusi peli automaattisesti"
Public Const IDS_DLG_SETTINGS_SHOW_SCORE = "Näytä pistetilanne erien välillä"

' card back names

Public Const IDS_CARD_BACK_0 = "Ristipunos"
Public Const IDS_CARD_BACK_1 = "Putkenpätkät"
Public Const IDS_CARD_BACK_2 = "Robotti"
Public Const IDS_CARD_BACK_3 = "Ruusut"
Public Const IDS_CARD_BACK_4 = "Muratti mustalla taustalla"
Public Const IDS_CARD_BACK_5 = "Muratti sinisellä taustalla"
Public Const IDS_CARD_BACK_6 = "Kalat syaanilla taustalla"
Public Const IDS_CARD_BACK_7 = "Kalat sinisellä taustalla"
Public Const IDS_CARD_BACK_8 = "Kotilo"
Public Const IDS_CARD_BACK_9 = "Öinen linna"
Public Const IDS_CARD_BACK_10 = "Aurinkoranta"
Public Const IDS_CARD_BACK_11 = "Korttihai"

' ui strings

Public Const IDS_COPYRIGHT = "Copyright © 2001-2003 Moonbird Software"
Public Const IDS_EMAIL = "moonbirdsoftware@hotmail.com"
Public Const IDS_URL = "http://www.geocities.com/moonbirdsoftware/"

' action button strings

Public Const IDS_ACTION_PLAY_CARD = "Pelaa kortti tikkiin"
Public Const IDS_ACTION_PLAY_CARDS = "Pelaa kortit tikkiin"
Public Const IDS_ACTION_TAKE_CARDS = "Nosta tikki"
Public Const IDS_ACTION_TRY_DECK = "Kokeile pakasta"
Public Const IDS_ACTION_CALL = "Epäile kortteja"
Public Const IDS_ACTION_PLACE_CARDS = "Laita kortit pöydälle"

' x

Public Const IDS_CARD = "kortti"
Public Const IDS_CARDS = "korttia"
Public Const IDS_DEALER = "Pakka"
Public Const IDS_NAME = "Nimi"
Public Const IDS_PLAYER = "Pelaaja"
Public Const IDS_SCORE = "Pisteet"
Public Const IDS_SCOREBOARD = "Pistetilanne"
Public Const IDS_TRICK = "Tikki"
Public Const IDS_TRASH = "Kaatopakka"
Public Const IDS_VERSION = "Versio"
Public Const IDS_ROUND = "Erä"
Public Const IDS_ROUND_2 = "Kierros"
Public Const IDS_TRUMP = "Valtti"

' player status

Public Const IDS_STATUS_TAKE = "%s nostaa kortit..."
Public Const IDS_STATUS_KILL = "%s kaataa..."
Public Const IDS_STATUS_TRY = "%s yrittää pakasta..."
Public Const IDS_STATUS_TAKE_FROM = "%s1 ottaa kortin %s2..."
Public Const IDS_STATUS_GIVE_TO = "%s1 antaa kortin %s2..."

' game status

Public Const IDS_STATUS_DEALING = "Jaetaan kortit..."
Public Const IDS_STATUS_DEMO = "Demo käynnissä..."
Public Const IDS_STATUS_NEW_GAME = "Aloitetaan uusi peli..."
Public Const IDS_STATUS_WAITING = "Odotetaan pelaajia..."
Public Const IDS_STATUS_PRESS_F2 = "Aloita uusi peli painamalla F2."

Public Const IDS_STATUS_TURN = "%s vuoro..."

Public Const IDS_ERROR_CARDS32 = "Kirjaston cards32.dll alustaminen epäonnistui."

Public Const IDS_QUERY_FIRST_RUN = "Käynnistät %s ensimmäistä kertaa. Voit nyt mukauttaa pelin asetukset haluamiksesi."
Public Const IDS_QUERY_RESTART_GAME = "Haluatko varmasti aloittaa uuden pelin? Et voi jatkaa nykyistä peliä, jos muutat sääntöjä."
Public Const IDS_QUERY_NEW_GAME = "Haluatko aloittaa uuden pelin?"
Public Const IDS_QUERY_CONTINUE_TURN = "Haluatko jatkaa vuoroasi?"

' card selection

Public Const IDS_STATUS_CHOOSE_CARD = "Valitse pelattava kortti."
Public Const IDS_STATUS_CHOOSE_CARDS = "Valitse pelattavat kortit."
Public Const IDS_STATUS_CHOOSE_LOWEST_RANK = "Pienin kortti aloittaa. Valitse %s."
Public Const IDS_STATUS_CHOOSE_RANK_OR_SUIT = "Valitse %s."
Public Const IDS_STATUS_CHOOSE_SAME_OR_HIGHER_RANK = "Valitse %s tai isompi kortti."
Public Const IDS_STATUS_CHOOSE_TRUMP_CARD = "Valitse valttikortti."
Public Const IDS_STATUS_CHOOSE_CARD_TO_GIVE = "%s1 ei voi pelata mitään kädessään olevia kortteja. Valitse %s2 annettava kortti."
Public Const IDS_STATUS_CHOOSE_CARD_TO_TAKE = "Et voi pelata mitään kädessäsi olevista korteista. Valitse %s otettava kortti."
Public Const IDS_STATUS_CANNOT_PLAY = "Et voi pelata mitään kädessäsi olevia kortteja. Nosta tikki tai yritä pakasta."

' card names

Public Const IDS_CARD_ACE = "ässä"
Public Const IDS_CARD_DEUX = "kakkonen"
Public Const IDS_CARD_THREE = "kolmonen"
Public Const IDS_CARD_FOUR = "nelonen"
Public Const IDS_CARD_FIVE = "viitonen"
Public Const IDS_CARD_SIX = "kuutonen"
Public Const IDS_CARD_SEVEN = "seiska"
Public Const IDS_CARD_EIGHT = "kasi"
Public Const IDS_CARD_NINE = "ysi"
Public Const IDS_CARD_TEN = "kymppi"
Public Const IDS_CARD_JACK = "jätkä"
Public Const IDS_CARD_QUEEN = "rouva"
Public Const IDS_CARD_KING = "kuningas"

' suit names

Public Const IDS_CARD_CLUB = "risti"
Public Const IDS_CARD_DIAMOND = "ruutu"
Public Const IDS_CARD_HEART = "hertta"
Public Const IDS_CARD_SPADE = "pata"

' debug player name - enter this as player 1's name to activate debug mode

Public Const IDS_DEBUG_PLR_NAME = "Aku Ankka"
Sub SetGameLanguage()
    Game.Language = IDL_FINNISH
End Sub
