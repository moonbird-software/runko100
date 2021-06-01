VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asetukset"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   1
      Left            =   240
      ScaleHeight     =   5055
      ScaleWidth      =   5055
      TabIndex        =   15
      Top             =   600
      Width           =   5055
      Begin VB.Frame fraRules 
         Caption         =   "Säännöt"
         Height          =   1455
         Index           =   3
         Left            =   2280
         TabIndex        =   39
         Top             =   3360
         Width           =   3615
         Visible         =   0   'False
         Begin VB.Image imaRules 
            Height          =   480
            Index           =   3
            Left            =   240
            Picture         =   "Settings.frx":0442
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraRules 
         Caption         =   "Säännöt"
         Height          =   1455
         Index           =   2
         Left            =   2280
         TabIndex        =   38
         Top             =   1920
         Width           =   3615
         Visible         =   0   'False
         Begin VB.Image imaRules 
            Height          =   480
            Index           =   2
            Left            =   240
            Picture         =   "Settings.frx":1084
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraRules 
         Caption         =   "Säännöt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   1
         Left            =   600
         TabIndex        =   35
         Top             =   0
         Width           =   4815
         Visible         =   0   'False
         Begin VB.OptionButton optRule 
            Caption         =   "Antopeli (jos mikään kortti ei käy, edellinen pelaaja antaa sinulle valitsemansa kortin)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   960
            TabIndex        =   37
            Top             =   360
            Width           =   3615
         End
         Begin VB.OptionButton optRule 
            Caption         =   "Ottopeli (jos mikään kortti ei käy, voit valita minkä kortin otat edelliseltä pelaajalta)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   960
            TabIndex        =   36
            Top             =   960
            Width           =   3615
         End
         Begin VB.Image imaRules 
            Height          =   480
            Index           =   1
            Left            =   240
            Picture         =   "Settings.frx":1CC6
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraPerformance 
         Caption         =   "Suorituskyky"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   4815
         Begin VB.CheckBox chkQuickDeal 
            Caption         =   "Pikajako"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   30
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chkShowScore 
            Caption         =   "Näytä pistetilanne erien välillä"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   29
            Top             =   1800
            Width           =   3495
         End
         Begin VB.CheckBox chkAutoRestart 
            Caption         =   "Aloita uusi erä automaattisesti"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   28
            Top             =   1440
            Width           =   3495
         End
         Begin VB.CheckBox chkCardAnim 
            Caption         =   "Animoi korttien siirrot"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   10
            Top             =   1080
            Width           =   2415
         End
         Begin MSComctlLib.Slider sldSpeed 
            Height          =   375
            Left            =   2040
            TabIndex        =   9
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            Min             =   1
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Label lblSpeed 
            Caption         =   "Pelin nopeus:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   27
            Top             =   540
            Width           =   975
         End
         Begin VB.Image imaPerformance 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":2908
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraRules 
         Caption         =   "Säännöt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   4815
         Visible         =   0   'False
         Begin VB.CheckBox chkRule 
            Caption         =   "Tikin voi aloittaa kuvakorteilla"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   6
            Top             =   960
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.CheckBox chkRule 
            Caption         =   "Kymppi on ainoa kaatokortti"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   7
            Top             =   1320
            Width           =   3615
         End
         Begin VB.CheckBox chkRule 
            Caption         =   "Kakkosen päälle käy mikä tahansa kortti"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   8
            Top             =   1680
            Width           =   3615
         End
         Begin VB.ComboBox cboGameType 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   2895
         End
         Begin VB.Image imaRules 
            Height          =   480
            Index           =   0
            Left            =   240
            Picture         =   "Settings.frx":354A
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblGameType 
            Caption         =   "Peli:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   17
            Top             =   420
            Width           =   735
         End
      End
   End
   Begin MSComctlLib.TabStrip tabSettings 
      Height          =   5655
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9975
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Yleiset"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lisäasetukset"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   0
      Left            =   240
      ScaleHeight     =   5055
      ScaleWidth      =   5055
      TabIndex        =   14
      Top             =   600
      Width           =   5055
      Begin VB.Frame fraDeck 
         Caption         =   "Pakka"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   4815
         Begin VB.ListBox lstDeck 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1425
            ItemData        =   "Settings.frx":418C
            Left            =   960
            List            =   "Settings.frx":418E
            TabIndex        =   4
            Top             =   840
            Width           =   2055
         End
         Begin VB.PictureBox picDeck 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   3240
            ScaleHeight     =   97
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   89
            TabIndex        =   25
            Top             =   840
            Width           =   1335
         End
         Begin VB.Image imaDeck 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":4190
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblDeck 
            Caption         =   "Valitse pakan taustakuva:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   26
            Top             =   480
            Width           =   3615
         End
      End
      Begin VB.Frame fraPlayers 
         Caption         =   "Pelaajat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   4815
         Begin VB.ComboBox comAI 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1800
            Width           =   495
         End
         Begin VB.ComboBox comAI 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1320
            Width           =   495
         End
         Begin VB.ComboBox comAI 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   840
            Width           =   495
         End
         Begin VB.ComboBox comAI 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   0
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   1
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   2
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   3
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblName 
            Caption         =   "Pelaaja 1:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   23
            Top             =   390
            Width           =   735
         End
         Begin VB.Label lblName 
            Caption         =   "Pelaaja 2:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   22
            Top             =   870
            Width           =   735
         End
         Begin VB.Label lblName 
            Caption         =   "Pelaaja 3:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   21
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label lblName 
            Caption         =   "Pelaaja 4:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   20
            Top             =   1830
            Width           =   735
         End
         Begin VB.Image imaPlayers 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":4DD2
            Top             =   360
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboGameType_Click()
    Select Case cboGameType.ListIndex
    'Case IDG_NORMAL
    Case 0
        chkRule(0).Value = 1
        chkRule(1).Value = 0
        chkRule(2).Value = 0
    'Case IDG_FAKE
    Case 1
        chkRule(0).Value = 1
        chkRule(1).Value = 1
        chkRule(2).Value = 1
    'Case IDG_SPANISH
    Case 2
        chkRule(0).Value = 0
        chkRule(1).Value = 0
        chkRule(2).Value = 0
    End Select
End Sub
Private Sub cboGameType_KeyDown(KeyCode As Integer, Shift As Integer)
    cboGameType_Click
End Sub


Private Sub chkCardAnim_Click()
    chkQuickDeal.Enabled = -chkCardAnim.Value
End Sub

Private Sub chkRule_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    cboGameType.ListIndex = cboGameType.ListCount - 1
End Sub
Private Sub chkRule_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    cboGameType.ListIndex = cboGameType.ListCount - 1
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    FormSettingsSave
End Sub

Private Sub Form_Load()
    FormSettingsLoad
End Sub

Private Sub lstDeck_Click()
    FormSettingsCardBackClick
End Sub
Private Sub tabSettings_Click()
    picSettings(tabSettings.SelectedItem.Index - 1).ZOrder 0
End Sub
Private Sub txtName_GotFocus(Index As Integer)
    txtName(Index).SelStart = 0
    txtName(Index).SelLength = Len(txtName(Index).Text)
End Sub
