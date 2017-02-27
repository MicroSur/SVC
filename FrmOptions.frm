VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   11175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15135
   Icon            =   "FrmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrExport 
      Height          =   5475
      Left            =   3000
      TabIndex        =   37
      Top             =   3000
      Width           =   7395
      Begin SurVideoCatalog.XpB comOptDelPreset 
         Height          =   315
         Left            =   4920
         TabIndex        =   5
         Top             =   2400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         Caption         =   "Del"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.OptionButton OptHtml 
         Appearance      =   0  'Flat
         Caption         =   "Title"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   3660
         TabIndex        =   9
         Top             =   3525
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.OptionButton OptHtml 
         Appearance      =   0  'Flat
         Caption         =   "File name"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   8
         Top             =   3525
         Width           =   1635
      End
      Begin VB.OptionButton OptHtml 
         Appearance      =   0  'Flat
         Caption         =   "Random"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   3525
         Width           =   1695
      End
      Begin VB.CheckBox chExpFolders 
         Appearance      =   0  'Flat
         Caption         =   "SubFolders"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   3900
         Width           =   6915
      End
      Begin VB.TextBox tExpFolders 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   4920
         TabIndex        =   13
         Top             =   4260
         Width           =   2175
      End
      Begin VB.TextBox tExpFolders 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2580
         TabIndex        =   12
         Top             =   4260
         Width           =   2175
      End
      Begin VB.TextBox tExpFolders 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   4260
         Width           =   2175
      End
      Begin VB.ListBox lstOptPreset 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   1395
         Left            =   4920
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox tExpDelim 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox TxtNnOnPage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         TabIndex        =   15
         Text            =   "30"
         Top             =   4980
         Width           =   435
      End
      Begin VB.ComboBox CombTemplate 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   4980
         Width           =   3495
      End
      Begin VB.ListBox LstExport 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Columns         =   3
         Height          =   2055
         ItemData        =   "FrmOptions.frx":000C
         Left            =   120
         List            =   "FrmOptions.frx":000E
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   300
         Width           =   4635
      End
      Begin SurVideoCatalog.XpB ComUnmarkAll 
         Height          =   315
         Left            =   6120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "Nothing"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComMarkAll 
         Height          =   315
         Left            =   4920
         TabIndex        =   1
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "All"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB comOptSavePreset 
         Height          =   315
         Left            =   4920
         TabIndex        =   3
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         Caption         =   "Save"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   7380
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Label lExpChB 
         Caption         =   "Export Fields"
         Height          =   255
         Left            =   180
         TabIndex        =   62
         Top             =   2460
         Width           =   4575
      End
      Begin VB.Label lExpDelim 
         Caption         =   "delimiter"
         Height          =   255
         Left            =   1860
         TabIndex        =   56
         Top             =   2820
         Width           =   5295
      End
      Begin VB.Label LblOptHtml 
         Caption         =   "Pictures file names"
         Height          =   255
         Left            =   180
         TabIndex        =   39
         Top             =   3300
         Width           =   6855
      End
      Begin VB.Label LTempl 
         Caption         =   "Template"
         Height          =   255
         Left            =   180
         TabIndex        =   38
         Top             =   4680
         Width           =   6975
      End
   End
   Begin VB.Frame FrGlobal 
      Height          =   5475
      Left            =   3660
      TabIndex        =   34
      Top             =   1500
      Width           =   7395
      Begin MSComctlLib.TreeView tvOpt 
         Height          =   4395
         Left            =   60
         TabIndex        =   87
         Top             =   180
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   7752
         _Version        =   393217
         LabelEdit       =   1
         Style           =   1
         FullRowSelect   =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.ComboBox ComboCDHid 
         Height          =   315
         Left            =   60
         TabIndex        =   35
         Text            =   "D:\"
         Top             =   4980
         Width           =   6870
      End
      Begin SurVideoCatalog.XpB cOptBrowseHid 
         Height          =   315
         Left            =   6990
         TabIndex        =   89
         Top             =   4980
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         Caption         =   "..."
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.Label LCD 
         Caption         =   "Path"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   4680
         Width           =   5415
      End
   End
   Begin VB.Frame FrFont 
      Height          =   5475
      Left            =   6480
      TabIndex        =   40
      Top             =   1080
      Width           =   7395
      Begin VB.CheckBox chLVGrid 
         Appearance      =   0  'Flat
         Caption         =   "LVGrid"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         TabIndex        =   88
         Top             =   3360
         Width           =   6975
      End
      Begin VB.CheckBox chNoLVSelFr 
         Appearance      =   0  'Flat
         Caption         =   "NoLVSelFrame"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         TabIndex        =   86
         Top             =   3060
         Width           =   6975
      End
      Begin VB.CheckBox chStripedLV 
         Appearance      =   0  'Flat
         Caption         =   "StripedLV"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         TabIndex        =   85
         Top             =   2760
         Width           =   6975
      End
      Begin VB.ComboBox ComboLangHid 
         Height          =   315
         ItemData        =   "FrmOptions.frx":0010
         Left            =   420
         List            =   "FrmOptions.frx":0012
         TabIndex        =   65
         Top             =   4560
         Width           =   3750
      End
      Begin VB.CheckBox chVMcolor 
         Appearance      =   0  'Flat
         Caption         =   "VertMenu"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         TabIndex        =   55
         Top             =   3720
         Width           =   6975
      End
      Begin VB.TextBox TextFontH 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   420
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1320
         Width           =   3765
      End
      Begin VB.TextBox TextFontV 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   420
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   540
         Width           =   3765
      End
      Begin VB.TextBox TextFontLV 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   420
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   2220
         Width           =   3765
      End
      Begin SurVideoCatalog.XpB ComCoverHorFillColor 
         Height          =   315
         Left            =   5940
         TabIndex        =   41
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   ""
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComCoverVertFillColor 
         Height          =   315
         Left            =   5940
         TabIndex        =   42
         Top             =   540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "ForeColor"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComColorPick 
         Height          =   315
         Left            =   5940
         TabIndex        =   43
         Top             =   2220
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   ""
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComFontLV 
         Height          =   315
         Left            =   4620
         TabIndex        =   45
         Top             =   2220
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   ""
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComFontV 
         Height          =   315
         Left            =   4620
         TabIndex        =   46
         Top             =   540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "Font"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComFontH 
         Height          =   315
         Left            =   4620
         TabIndex        =   47
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   ""
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComLangOk 
         Height          =   315
         Left            =   4620
         TabIndex        =   66
         Top             =   4560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Caption         =   "Ok"
         ButtonStyle     =   3
         Picture         =   "FrmOptions.frx":0014
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB comHiLightColor 
         Height          =   315
         Left            =   2280
         TabIndex        =   84
         Top             =   5040
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Caption         =   ""
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.Label lblLang 
         Caption         =   "Language"
         Height          =   255
         Left            =   180
         TabIndex        =   67
         Top             =   4200
         Width           =   6375
      End
      Begin VB.Label LHFont 
         Caption         =   "Cover Hor Font"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1020
         Width           =   2895
      End
      Begin VB.Label LVFont 
         Caption         =   "Cover Vert Font"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label LLVFont 
         Caption         =   "Table Font"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1920
         Width           =   2880
      End
   End
   Begin VB.Frame frInet 
      Height          =   5475
      Left            =   8580
      TabIndex        =   57
      Top             =   9600
      Width           =   7395
      Begin VB.TextBox tSVCNewVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   75
         Top             =   4020
         Width           =   1635
      End
      Begin VB.TextBox tSVCCurVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   74
         Top             =   4380
         Width           =   1635
      End
      Begin VB.TextBox txtSVCSite 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "http://sur.hotbox.ru/"
         Top             =   3660
         Width           =   3495
      End
      Begin VB.OptionButton optProxy 
         Appearance      =   0  'Flat
         Caption         =   "My"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   18
         Top             =   1080
         Width           =   6855
      End
      Begin VB.OptionButton optProxy 
         Appearance      =   0  'Flat
         Caption         =   "IE"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   17
         Top             =   660
         Width           =   6855
      End
      Begin VB.OptionButton optProxy 
         Appearance      =   0  'Flat
         Caption         =   "No"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   16
         Top             =   360
         Width           =   6855
      End
      Begin VB.CheckBox chSecure 
         Appearance      =   0  'Flat
         Caption         =   "Secure"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   22
         Top             =   3060
         Width           =   6735
      End
      Begin VB.TextBox tProxyPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   2460
         Width           =   4395
      End
      Begin VB.TextBox tProxyUserName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   1980
         Width           =   4395
      End
      Begin VB.TextBox tProxyServerPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   1500
         Width           =   4395
      End
      Begin SurVideoCatalog.XpB comFindProxy 
         Height          =   285
         Left            =   6420
         TabIndex        =   61
         Top             =   1500
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         Caption         =   "Find"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB comInetVerCheck 
         Height          =   285
         Left            =   3720
         TabIndex        =   68
         Top             =   4020
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         Caption         =   "CheckNew"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB comSiteGo 
         Height          =   285
         Left            =   5580
         TabIndex        =   70
         Top             =   3660
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         Caption         =   "Go"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.Label lSVCCurVer 
         Caption         =   "CurVer"
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label lSVCNewVer 
         Caption         =   "NewVer"
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label lSVCSite 
         Caption         =   "SVCSite"
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   3720
         Width           =   1515
      End
      Begin VB.Label lpPass 
         Caption         =   "Password"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Label lpUser 
         Caption         =   "User"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label lpServerPort 
         Caption         =   "Server:Port"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   1560
         Width           =   1515
      End
   End
   Begin VB.Frame FrameBD 
      Height          =   5475
      Left            =   1200
      TabIndex        =   24
      Top             =   720
      Width           =   7395
      Begin VB.TextBox TextQJPGHid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         MaxLength       =   3
         TabIndex        =   63
         Text            =   "80"
         Top             =   4800
         Width           =   375
      End
      Begin VB.ListBox LstBases 
         Appearance      =   0  'Flat
         DragIcon        =   "FrmOptions.frx":0A26
         Height          =   1785
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Width           =   7155
      End
      Begin SurVideoCatalog.XpB ComDelBD 
         Height          =   375
         Left            =   2640
         TabIndex        =   25
         Top             =   2700
         Width           =   2055
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "Exclude"
         ButtonStyle     =   3
         Picture         =   "FrmOptions.frx":0D30
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComCompactA 
         Height          =   375
         Left            =   180
         TabIndex        =   27
         Top             =   4080
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   661
         Caption         =   "Compact Actors"
         ButtonStyle     =   3
         Picture         =   "FrmOptions.frx":1742
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComCompact 
         Height          =   375
         Left            =   180
         TabIndex        =   28
         Top             =   3480
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   661
         Caption         =   "Compact base"
         ButtonStyle     =   3
         Picture         =   "FrmOptions.frx":1CDC
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComOpenBD 
         Height          =   375
         Left            =   180
         TabIndex        =   29
         Top             =   2700
         Width           =   2055
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "Add"
         ButtonStyle     =   3
         Picture         =   "FrmOptions.frx":2276
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
         MaskColor       =   16711935
      End
      Begin SurVideoCatalog.XpB ComNewBD 
         Height          =   375
         Left            =   5100
         TabIndex        =   30
         Top             =   2700
         Width           =   2055
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "New"
         ButtonStyle     =   3
         Picture         =   "FrmOptions.frx":25CA
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.Label LQJPG 
         Caption         =   "JPEG (0-100)"
         Height          =   285
         Left            =   660
         TabIndex        =   64
         Top             =   4860
         Width           =   6435
      End
      Begin VB.Label LabelCurrBDHid 
         Caption         =   ">"
         Height          =   225
         Left            =   240
         TabIndex        =   33
         Top             =   2220
         Width           =   5670
      End
      Begin VB.Label LBDSizeHid 
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label LABDSizeHid 
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   4200
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList imlOpt 
      Left            =   360
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":2FDC
            Key             =   "Bases"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":3576
            Key             =   "Fonts"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":3B10
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":40AA
            Key             =   "Other"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":4644
            Key             =   "Internet"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":5056
            Key             =   "Combos"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":55F0
            Key             =   "About"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":5B8A
            Key             =   "Plus"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":6124
            Key             =   "Minus"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":66BE
            Key             =   "Yes"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOptions.frx":6C58
            Key             =   "No"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabSOpt 
      Height          =   735
      Left            =   60
      TabIndex        =   54
      Top             =   60
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   1296
      MultiRow        =   -1  'True
      Style           =   1
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   2099
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bases"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fonts"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Other"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Internet"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Combos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "?"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frCombo 
      Height          =   5475
      Left            =   240
      TabIndex        =   76
      Top             =   840
      Width           =   7395
      Begin VB.TextBox txtComboView 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   4140
         Width           =   5235
      End
      Begin VB.ListBox lstComboNames 
         Appearance      =   0  'Flat
         Height          =   1785
         Left            =   720
         TabIndex        =   82
         Top             =   2100
         Visible         =   0   'False
         Width           =   1755
      End
      Begin SurVideoCatalog.XpB comComboDel 
         Height          =   285
         Left            =   5580
         TabIndex        =   78
         Top             =   4140
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         Caption         =   "Del"
         ButtonStyle     =   3
         Picture         =   "FrmOptions.frx":71F2
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.ListBox lstComboVal 
         Appearance      =   0  'Flat
         Height          =   3735
         Left            =   3000
         TabIndex        =   81
         Top             =   240
         Width           =   4275
      End
      Begin VB.ListBox lstComboName 
         Appearance      =   0  'Flat
         Height          =   3735
         Left            =   120
         TabIndex        =   80
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtComboAdd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   77
         Top             =   4680
         Width           =   5235
      End
      Begin SurVideoCatalog.XpB comComboAdd 
         Height          =   285
         Left            =   5580
         TabIndex        =   79
         Top             =   4680
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         Caption         =   "Add"
         ButtonStyle     =   3
         Picture         =   "FrmOptions.frx":7C04
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
   End
   Begin SurVideoCatalog.XpB comOptRet 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5460
      TabIndex        =   23
      Top             =   6240
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      Caption         =   "Cancel"
      ButtonStyle     =   3
      Picture         =   "FrmOptions.frx":8616
      PictureWidth    =   16
      PictureHeight   =   16
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB ComOptSave 
      Height          =   375
      Left            =   360
      TabIndex        =   53
      Top             =   6240
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   661
      Caption         =   "Save"
      ButtonStyle     =   3
      Picture         =   "FrmOptions.frx":9028
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      MaskColor       =   16711935
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'? Залочить форму опций модально
Private DragIndex As Integer 'для драг дропа в списке баз

Private FillLstProgFlag As Boolean 'по lvopt кликаем не сами, а прога
Private opt_bdname As String
Private ini_opt As String ' типа INIFILE
Private OptCaption As String 'название окна
Private oldComboCDHid As String 'запомненное до выбора сидюка
Private Manual_1_flag As Boolean ' как кликнули на chVMcolor и chStripedLV и chNoLVSelFr


Private ComDelBD_RealClick As Boolean
Private NStoreOpt(9) As String 'от 0, доп. фразы

Private UserIniCurrSection As String 'текущая секция в user.lng, открытая в настройках комбо

Private Sub chExpFolders_Click()
If FrmOptions.Visible Then optsaved = False
Opt_ExpUseFolders = chExpFolders.Value
End Sub

Private Sub chLVGrid_Click()
If Manual_1_flag Then
    Opt_ShowLVGrid = CBool(chLVGrid.Value)
    NoSetColorFlag = False
    ApplyOpt
    optsaved = False
End If
Manual_1_flag = True

End Sub

Private Sub chLVGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Manual_1_flag = True
End Sub

Private Sub chLVGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Manual_1_flag = True
End Sub

Private Sub chNoLVSelFr_Click()
If Manual_1_flag Then
    NoLVSelFrame = CBool(chNoLVSelFr.Value)
    NoSetColorFlag = False
    ApplyOpt
    optsaved = False
End If
Manual_1_flag = True
End Sub

Private Sub chNoLVSelFr_KeyDown(KeyCode As Integer, Shift As Integer)
Manual_1_flag = True
End Sub

Private Sub chNoLVSelFr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Manual_1_flag = True
End Sub

Private Sub chSecure_Click()
If FrmOptions.Visible Then optsaved = False
Opt_InetSecureFlag = chSecure.Value
End Sub

Private Sub chStripedLV_Click()
If Manual_1_flag Then
    StripedLV = CBool(chStripedLV.Value)
    NoSetColorFlag = False
    ApplyOpt
    optsaved = False
End If
Manual_1_flag = True
End Sub

Private Sub chStripedLV_KeyDown(KeyCode As Integer, Shift As Integer)
Manual_1_flag = True
End Sub

Private Sub chStripedLV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Manual_1_flag = True
End Sub

Private Sub chVMcolor_Click()
If Manual_1_flag Then
    VMSameColor = CBool(chVMcolor.Value)
    NoSetColorFlag = False
    ApplyOpt
    optsaved = False
End If
Manual_1_flag = True
End Sub

Private Sub chVMcolor_KeyDown(KeyCode As Integer, Shift As Integer)
Manual_1_flag = True
End Sub

Private Sub chVMcolor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Manual_1_flag = True
End Sub

Private Sub ComboCDHid_Click()
ComboCDHid = ComboCDHid & ";" & oldComboCDHid
End Sub

Private Sub comComboAdd_Click()
Dim kName As String    'имя ключа
'Dim kVal As String 'значение
If Len(txtComboAdd) = 0 Then Exit Sub

If FileExists(userFile) Then
    'если нет такого
    If SearchListBox(lstComboVal, txtComboAdd, False) = -1 Then

        'записать ключ
        kName = GetAutoIniKeyName(lstComboNames)

        If WriteKey(UserIniCurrSection, kName, txtComboAdd, userFile) > 0 Then

            'внести в список
            lstComboVal.AddItem txtComboAdd
            lstComboNames.AddItem kName

            'перезагрузить новость
            Call FillUserCombo(UserIniCurrSection, lstComboVal)
            
            'и сразу в список редактора
            With frmEditor
            Select Case UserIniCurrSection
            Case "Genre":  .ComboGenre.AddItem txtComboAdd
            Case "Country":  .ComboCountry.AddItem txtComboAdd
            Case "Language":  .TextLang.AddItem txtComboAdd
            Case "Subtitle":  .TextSubt.AddItem txtComboAdd
            Case "Comments":  .ComboOther.AddItem txtComboAdd
            Case "Media":  .ComboNos.AddItem txtComboAdd
            Case "Site"
                .ComboSites.AddItem txtComboAdd
                .cBasePicURL.AddItem txtComboAdd
            End Select
            End With
        End If

    End If
End If
End Sub

Private Sub comComboDel_Click()
'удалить выбранный пункт из выбранного комбо
Dim ret As Long
Dim pname As String 'имя текущего ключа
Dim pval As String 'значение -..-

If lstComboVal.ListIndex < 0 Then Exit Sub
If lstComboNames.ListCount <> lstComboVal.ListCount Then Exit Sub

pname = lstComboNames.List(lstComboVal.ListIndex) 'lstComboVal.ListIndex + 1
pval = lstComboVal.List(lstComboVal.ListIndex)

'убить ключ
ret = DeleteKey(pname, UserIniCurrSection, userFile)

If ret > 0 Then
    'удалить из списков
    ret = SearchListBox(lstComboVal, pval, False)
    If ret > -1 Then
        lstComboVal.RemoveItem ret
    End If
    ret = SearchListBox(lstComboNames, pname, False)
    If ret > -1 Then
        lstComboNames.RemoveItem ret
    End If
    
    'на всякий пожарный оставить txtComboView = vbNullString
End If
End Sub

Private Sub ComDelBD_KeyDown(KeyCode As Integer, Shift As Integer)
ComDelBD_RealClick = True
End Sub

Private Sub ComDelBD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ComDelBD_RealClick = True
End Sub

Private Sub comFindProxy_Click()
Dim result As Long
Dim RegKey As String
Dim RegRoot As Long ' Registry Root z.B. HKEY_CURRENT_USER
Dim retstr As String
Dim pos As Long

RegRoot = HKEY_CURRENT_USER

' Dieser Schlьssel wird unter Windows 95 fьr MS Internet Explorer verwendet.
' Andere Betriebssyteme verwenden eventuell einen anderen Schьssel.
RegKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"

result = RegKeyExist(RegRoot, RegKey)
If result <> 0 Then
    Exit Sub    '   MsgBox "Fehler!"
End If



Dim ret As Long
ret = RegValueGet(RegRoot, RegKey, "ProxyServer", retstr)
pos = InStr(1, retstr, "http=", 1)
If pos > 0 Then
    retstr = Mid$(retstr, pos + 5)     ' das "http:" entfernen
End If
pos = InStr(1, retstr, ";", 1)
If pos > 0 Then
    retstr = Left$(retstr, pos - 1)
End If

tProxyServerPort = retstr


End Sub



Private Sub comHiLightColor_Click()
Dim c As Long
Dim cd As cCommonDialog
Dim ret As Long

Set cd = New cCommonDialog
cd.CustomColor(0) = LVHighLightLong

ret = cd.VBChooseColor(c, True, True, False, Me.hwnd)
If ret = 0 Then Exit Sub
LVHighLightLong = c
'Set cd = Nothing
NoSetColorFlag = False

'setForeColorOpt
ApplyOpt

optsaved = False
End Sub

Private Sub comInetVerCheck_Click()
'получить из файла на сервере данные о версии на сайте
Dim url As String
Dim tmp As String

Screen.MousePointer = vbHourglass

If Len(txtSVCSite) <> 0 Then
    If Right$(txtSVCSite, 1) <> "/" Then txtSVCSite = txtSVCSite & "/"
    url = txtSVCSite & "svcversioninfo"
    tmp = OpenURLProxy(url, "txt")

    If Len(tmp) > 10 Then
        tSVCNewVer.Text = "?.?.?"
    Else
        tSVCNewVer.Text = tmp
    End If
End If  'Len(txtSVCSite) <> 0

Screen.MousePointer = vbNormal
End Sub

Private Sub ComLangOk_Click()
'менаем язык

GetLang 'в опциях
Var2tvOpt 'показать чеки в иконках
Var2ListExport 'в ленге все чистили, вернуть
'убить лишнюю пустышку
'LstExport.List(LstExport.ListCount) = vbNullString    ' пустышка для ловли селекта
'LstExport.RemoveItem (LstExport.ListCount - 1)
'LstExport.ListIndex = LstExport.ListCount - 1

FrmMain.LangChange 'в мэйне
FrmMain.GetLangEditor

If Opt_UCLV_Vis Then FrmMain.UCLV.Refresh

LastLanguage = ComboLangHid.Text
End Sub


Private Sub FillTVOpt()
'Dim tmp As String

'FillTVOpt
'по умолч ветки раскрыты, чеки сняты

tvOpt.ImageList = imlOpt
tvOpt.Indentation = 0
tvOpt.Nodes.Clear
    
'Opt_InetGetPicUseTempFile=False

'1

tvOpt.Nodes.Add(, , "Opt_AutoSaveOpt", ReadLangOpt("Opt_AutoSaveOpt")).Checked = False

'2
tvOpt.Nodes.Add(, , "_List", ReadLangOpt("Opt_List"), "Minus").Expanded = True
tvOpt.Nodes.Add("_List", tvwChild, "Opt_SortOnStart", ReadLangOpt("Opt_SortOnStart")).Checked = False
tvOpt.Nodes.Add("_List", tvwChild, "Opt_SortLVAfterEdit", ReadLangOpt("Opt_SortLVAfterEdit")).Checked = False
tvOpt.Nodes.Add("_List", tvwChild, "Opt_SortLabelAsNum", ReadLangOpt("Opt_SortLabelAsNum")).Checked = False
tvOpt.Nodes.Add("_List", tvwChild, "Opt_Debtors_Colorize", ReadLangOpt("Opt_Debtors_Colorize")).Checked = False
tvOpt.Nodes.Add("_List", tvwChild, "Opt_LoadOnlyTitles", ReadLangOpt("Opt_LoadOnlyTitles")).Checked = False
tvOpt.Nodes.Add("_List", tvwChild, "Opt_LoanAllSameLabels", ReadLangOpt("Opt_LoanAllSameLabels")).Checked = False

'3
tvOpt.Nodes.Add(, , "_Video", ReadLangOpt("Opt_Video"), "Minus").Expanded = True
tvOpt.Nodes.Add("_Video", tvwChild, "Opt_PicRealRes", ReadLangOpt("Opt_PicRealRes")).Checked = False
tvOpt.Nodes.Add("_Video", tvwChild, "Opt_UseAspect", ReadLangOpt("Opt_UseAspect")).Checked = False
tvOpt.Nodes.Add("_Video", tvwChild, "Opt_UseOurMpegFilters", ReadLangOpt("Opt_UseOurMpegFilters")).Checked = False
tvOpt.Nodes.Add("_Video", tvwChild, "Opt_AviDirectShow", ReadLangOpt("Opt_AviDirectShow")).Checked = False

'4
tvOpt.Nodes.Add(, , "_Interface", ReadLangOpt("Opt_Interface"), "Minus").Expanded = True
tvOpt.Nodes.Add("_Interface", tvwChild, "Opt_NoSlideShow", ReadLangOpt("Opt_NoSlideShow")).Checked = False
tvOpt.Nodes.Add("_Interface", tvwChild, "Opt_CenterShowPic", ReadLangOpt("Opt_CenterShowPic")).Checked = False
tvOpt.Nodes.Add("_Interface", tvwChild, "Opt_UCLV_Vis", ReadLangOpt("Opt_UCLV_Vis")).Checked = False
tvOpt.Nodes.Add("_Interface", tvwChild, "Opt_Group_Vis", ReadLangOpt("Opt_Group_Vis")).Checked = False
tvOpt.Nodes.Add("_Interface", tvwChild, "Opt_PutOtherInAnnot", ReadLangOpt("Opt_PutOtherInAnnot")).Checked = False

'5
tvOpt.Nodes.Add(, , "_Export", ReadLangOpt("Opt_Export"), "Minus").Expanded = True
tvOpt.Nodes.Add("_Export", tvwChild, "Opt_ShowColNames", ReadLangOpt("Opt_ShowColNames")).Checked = False

'6
tvOpt.Nodes.Add(, , "_Media", ReadLangOpt("Opt_Media"), "Minus").Expanded = True
tvOpt.Nodes.Add("_Media", tvwChild, "Opt_GetMediaType", ReadLangOpt("Opt_GetMediaType")).Checked = False
tvOpt.Nodes.Add("_Media", tvwChild, "Opt_GetVolumeInfo", ReadLangOpt("Opt_GetVolumeInfo")).Checked = False
tvOpt.Nodes.Add("_Media", tvwChild, "Opt_QueryCancelAutoPlay", ReadLangOpt("Opt_QueryCancelAutoPlay")).Checked = False

'7
tvOpt.Nodes.Add(, , "_Editor", ReadLangOpt("Opt_Editor"), "Minus").Expanded = True
tvOpt.Nodes.Add("_Editor", tvwChild, "Opt_LVEDIT", ReadLangOpt("Opt_LVEDIT")).Checked = False
tvOpt.Nodes.Add("_Editor", tvwChild, "Opt_FileWithPath", ReadLangOpt("Opt_FileWithPath")).Checked = False

'tvOpt.SelectedItem.Selected = False
tvOpt.Nodes(1).Selected = True



End Sub

Private Sub comOptDelPreset_Click()
Dim ret As Long
Dim pname As String

If lstOptPreset.ListIndex < 0 Then Exit Sub
pname = lstOptPreset.List(lstOptPreset.ListIndex)
'убить ключ
ret = DeleteKey(pname, "ExportPreset", userFile)
If ret > 0 Then
    'удалить из списка lstOptPreset
    ret = SearchListBox(lstOptPreset, pname, False)
    If ret > -1 Then
        lstOptPreset.RemoveItem ret
    End If
End If

End Sub

Private Sub comOptRet_Click()

Unload Me
End Sub

Private Sub ApplyOpt()
With FrmMain
    'если база = текущей
    If opt_bdname = bdname Then
        setForeColor

        If Opt_ShowLVGrid Then
            .ListView.GridLines = True
            .tvGroup.GridLines = True
            .LVActer.GridLines = True
        Else
            .ListView.GridLines = False
            .tvGroup.GridLines = False
            .LVActer.GridLines = False
        End If

        If Opt_Group_Vis Then
            'FrmMain.FillTVGroup 'каждый раз при клике на опции
            If GroupedFlag Then InitFlag = False
        Else
            If GroupedFlag Then InitFlag = True
            'Dim ret As Long
            'ret = myMsgBox(msgsvc(42), vbYesNo, , FrmMain.hWnd) 'оставить группировку?
            'If ret = vbNo Then InitFlag = True 'перечитаем базу на выходе из опций
        End If

        If Opt_NoSlideShow Then
            .picScrollBoxV.Visible = True: .TextVAnnot.Visible = True
            If .FrameView.Visible Then .ComShowFa_Click
        Else
            .ComShowAn.Visible = False: .ComShowFa.Visible = True
        End If


        If ShowCoverFlag Then
            ExportDelim = getExpDelim(tExpDelim.Text)
            Select Case .TabStripCover.SelectedItem.Index
            Case 1
                Call ShowCoverStandard
            Case 2
                Call ShowCoverConvert
                
            Case 3
                Call ShowCoverDVD(False)
            Case 4
                Call ShowCoverDVD(True)
                
            Case 5
                'ChPrintChecked.Value = 0: ChPrintChecked.Enabled = False     'убрать галку печатать все
                Call ShowCoverSpisok
            End Select
        End If

    End If

End With
End Sub

Private Sub comOptSavePreset_Click()
Dim tmp As String, pname As String
Dim i As Integer
Dim oldName As String

If lstOptPreset.ListIndex > -1 Then
    oldName = lstOptPreset.List(lstOptPreset.ListIndex)
Else
    oldName = NStoreOpt(0)
End If

pname = myInputBox(NStoreOpt(1), , FrmOptions.hwnd, , , , oldName)
If Len(pname) = 0 Then Exit Sub
'If StrPtr(Dolg) = 0 Then Exit Sub 'cancel

'записать пресет в user.lng
If FileExists(userFile) Then

    For i = 0 To UBound(LstExport_Arr)
        If LstExport_Arr(i) Then
            tmp = tmp & "1,"
        Else
            tmp = tmp & "0,"
        End If
    Next i
    tmp = Left$(tmp, Len(tmp) - 1)

    WriteKey "ExportPreset", pname, tmp, userFile

    'внести с список пресетов
    If SearchListBox(lstOptPreset, pname, False) = -1 Then lstOptPreset.AddItem pname
End If

End Sub

Private Sub comSiteGo_Click()
Dim temp As Long
Dim strPath As String
Dim site As String

strPath = Space$(255)
site = txtSVCSite.Text

temp = FindExecutable(txtSVCSite, "", strPath)
Select Case temp
Case 31
    myMsgBox msgsvc(17), vbInformation, , Me.hwnd
    Exit Sub
Case 2
End Select

temp = ShellExecute(GetDesktopWindow(), "open", site, vbNull, vbNull, 1)
End Sub

Private Sub cOptBrowseHid_Click()
Dim sdir As String

sdir = BrowseForFolderByPath(vbNullString, vbNullString, FrmOptions.hwnd)
If Len(sdir) = 0 Then Exit Sub

If InStr(ComboCDHid.Text, sdir) Then
Else
    If Right$(ComboCDHid.Text, 1) <> ";" Then ComboCDHid.Text = ComboCDHid.Text & ";"
    ComboCDHid.Text = ComboCDHid.Text & sdir & ";"
End If

End Sub

Private Sub Form_Activate()
optsaved = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 112 'F1
    If FrmMain.ChBTT.Value = 0 Then FrmMain.ChBTT.Value = 1 Else FrmMain.ChBTT.Value = 0
    Me.SetFocus
End Select
End Sub

Private Sub Form_Load()
'Dim i As Integer

GetLang    '1 локализация

'при старте ини не читается -
'ReadINIOpt GetNameFromPathAndName(opt_bdname)
'берутся текущие переменные
Var2Options (True)    '2 заполнить поля данными из переменных

''иконки табам
'For i = 1 To TabSOpt.Tabs.Count
'TabSOpt.Tabs(i).Image = i
'Next i


TabSOpt.ZOrder 0
FrmOptions.Width = 7590 '8600 '7590
FrmOptions.Height = 7200 '7230 '6965

'TabSOpt.Tabs(1).Selected = True
'кликнуть на таб
If TabSOptLast = 0 Then TabSOptLast = 1
If TabSOpt.Tabs.Count >= TabSOptLast Then
    TabSOpt.Tabs(TabSOptLast).Selected = True
Else
    TabSOpt.Tabs(1).Selected = True
End If

'не давать перебор баз в главной
FrmMain.TabLVHid.Enabled = False

frmOptFlag = True

End Sub

Private Sub Options2Var()
'Процедура - вернуть значения переменным

'LstOpt2Var - каждый раз
'LstExport2Var - каждый раз

Dim i As Integer

'Имена картинок
OptHtml(Opt_HtmlJpgName) = True
For i = 0 To 2
    If OptHtml(i) = True Then Opt_HtmlJpgName = i: Exit For
Next i

QJPG = TextQJPGHid
ComboCDHid_Text = ComboCDHid

'текущий HTML шаблон
CurrentHtmlTemplate = CombTemplate.Text
TxtNnOnPage_Text = TxtNnOnPage.Text

ExportDelim = getExpDelim(tExpDelim.Text)

Opt_ExpFolder1 = tExpFolders(0).Text 'html
Opt_ExpFolder2 = tExpFolders(1).Text 'cover
Opt_ExpFolder3 = tExpFolders(2).Text 'sshots
End Sub
Private Sub Var2Options(glob As Boolean)
'Процедура - Из переменных в контролы - при этом форма невидима (проверить)
'mzt Dim temp As String
Dim i As Integer

On Error Resume Next

Select Case glob
Case True
    '                                                       Общие настройки
    ini_opt = INIFILE
    ' список                                                    шаблонов
    FillTemplateCombo App.Path & "\Templates\", CombTemplate

    '                                                           список баз
    LstBases.Clear
    For i = 1 To LstBases_ListCount
        LstBases.AddItem LstBases_List(i)
    Next i
    SetListboxScrollbar LstBases
    opt_bdname = LstBases.List(CurrentBaseIndex - 1)

    '                                                           язык интерфейса
    ComboLangHid.Clear
    For i = 1 To LangCount: ComboLangHid.AddItem ComboLang(i): Next i
    ComboLangHid.Text = LastLanguage

    If Len(abdname) <> 0 Then LABDSizeHid.Caption = FileLen(abdname) / 1024 & " K"

    'inet
    If Opt_InetUseProxy = 0 Or Opt_InetUseProxy = 1 Or Opt_InetUseProxy = 2 Then
        optProxy(Opt_InetUseProxy).Value = True
    End If
    tProxyServerPort = Opt_InetProxyServerPort
    tProxyUserName = Opt_InetUserName
    tProxyPass = Opt_InetPassword
    If Opt_InetSecureFlag Then chSecure.Value = vbChecked Else chSecure.Value = vbUnchecked

'списки, названия комбиков
lstComboName.Clear
lstComboName.AddItem NStoreOpt(2) 'жанр
lstComboName.AddItem NStoreOpt(3) 'страна
lstComboName.AddItem NStoreOpt(4) 'язык
lstComboName.AddItem NStoreOpt(5) 'субтитры
lstComboName.AddItem NStoreOpt(6) 'качество
lstComboName.AddItem NStoreOpt(7) 'носитель
lstComboName.AddItem NStoreOpt(8) 'сайты


    'кликнуть текущую базу
    If LstBases_ListCount > 0 Then LstBases.Selected(CurrentBaseIndex - 1) = True

Case False
    '                                                           настройки конкретной базы

    ' размер и имя базы
    If Len(opt_bdname) <> 0 Then
        LabelCurrBDHid.Caption = opt_bdname
        Me.Caption = OptCaption & ": " & opt_bdname

        LBDSizeHid.Caption = FileLen(opt_bdname) / 1024 & " K"
        'ComCompact.Enabled = True
    End If

    'цвета и шрифты
    setForeColorOpt
    If VMSameColor Then
        Manual_1_flag = False
        chVMcolor.Value = vbChecked
    Else
        chVMcolor.Value = vbUnchecked
    End If


    If StripedLV Then
        Manual_1_flag = False
        chStripedLV.Value = vbChecked
    Else
        chStripedLV.Value = vbUnchecked
    End If
    
    If NoLVSelFrame Then
        Manual_1_flag = False
        chNoLVSelFr.Value = vbChecked
    Else
        chNoLVSelFr.Value = vbUnchecked
    End If
    
    If Opt_ShowLVGrid Then
        Manual_1_flag = False
        chLVGrid.Value = vbChecked
    Else
        chLVGrid.Value = vbUnchecked
    End If


    '
    TextQJPGHid.Text = QJPG


'tvOpt другие
Var2tvOpt
'Поля экспорта
Var2ListExport

    'Пути path
    ComboCDHid.Text = ComboCDHid_Text

    'читать список пресетов экспорта
    Call ReadExportPresets

    'Имена картинок
    OptHtml(Opt_HtmlJpgName) = True

    'текущий HTML шаблон
    CombTemplate.Text = CurrentHtmlTemplate

    TxtNnOnPage.Text = TxtNnOnPage_Text
    'Разделитель
    tExpDelim.Text = putExpDelim(ExportDelim)

If Opt_ExpUseFolders Then chExpFolders.Value = vbChecked Else chExpFolders.Value = vbUnchecked
tExpFolders(0).Text = Opt_ExpFolder1 'html
tExpFolders(1).Text = Opt_ExpFolder2 'cover
tExpFolders(2).Text = Opt_ExpFolder3 'sshots

End Select
End Sub
Private Sub Var2ListExport()
Dim i As Integer
FillLstProgFlag = True
For i = 0 To LstExport_ListCount    '- 1
    LstExport.Selected(i) = LstExport_Arr(i)
Next i
FillLstProgFlag = False
End Sub

Private Sub Var2tvOpt()
tvOpt.Nodes("Opt_AutoSaveOpt").Checked = Opt_AutoSaveOpt
tvOpt.Nodes("Opt_SortOnStart").Checked = Opt_SortOnStart
tvOpt.Nodes("Opt_Debtors_Colorize").Checked = Opt_Debtors_Colorize
tvOpt.Nodes("Opt_LoadOnlyTitles").Checked = Opt_LoadOnlyTitles
tvOpt.Nodes("Opt_LoanAllSameLabels").Checked = Opt_LoanAllSameLabels
tvOpt.Nodes("Opt_PicRealRes").Checked = Opt_PicRealRes
tvOpt.Nodes("Opt_UseAspect").Checked = Opt_UseAspect
tvOpt.Nodes("Opt_UseOurMpegFilters").Checked = Opt_UseOurMpegFilters
tvOpt.Nodes("Opt_AviDirectShow").Checked = Opt_AviDirectShow
tvOpt.Nodes("Opt_NoSlideShow").Checked = Opt_NoSlideShow
tvOpt.Nodes("Opt_CenterShowPic").Checked = Opt_CenterShowPic
tvOpt.Nodes("Opt_UCLV_Vis").Checked = Opt_UCLV_Vis
tvOpt.Nodes("Opt_Group_Vis").Checked = Opt_Group_Vis
tvOpt.Nodes("Opt_ShowColNames").Checked = Opt_ShowColNames
tvOpt.Nodes("Opt_GetMediaType").Checked = Opt_GetMediaType
tvOpt.Nodes("Opt_GetVolumeInfo").Checked = Opt_GetVolumeInfo
tvOpt.Nodes("Opt_QueryCancelAutoPlay").Checked = Opt_QueryCancelAutoPlay

tvOpt.Nodes("Opt_LVEDIT").Checked = Opt_LVEDIT
tvOpt.Nodes("Opt_FileWithPath").Checked = Opt_FileWithPath

tvOpt.Nodes("Opt_SortLVAfterEdit").Checked = Opt_SortLVAfterEdit
tvOpt.Nodes("Opt_SortLabelAsNum").Checked = Opt_SortLabelAsNum
tvOpt.Nodes("Opt_PutOtherInAnnot").Checked = Opt_PutOtherInAnnot

tvOpt_Check2Ico 'показать чеки в иконках

End Sub
Private Sub GetLang()
Dim Contrl As Control
Dim i As Integer

On Error Resume Next

If Dir(lngFileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = vbNullString Or Len(lngFileName) = 0 Then
    Call myMsgBox("Не найден файл локализации! Исправьте параметр LastLang в global.ini" & vbCrLf & "Language file not found: " & vbCrLf & lngFileName, vbCritical, , Me.hwnd)
    'Unload FrmOptions - хреново
    Exit Sub
End If

'Screen.MousePointer = vbHourglass
'ToDebug "Чтение файла локализации: " & lngFileName

'LockWindowUpdate FrmOptions.hWnd

For Each Contrl In FrmOptions.Controls

    If TypeOf Contrl Is Label Then    '                           Label
        Contrl.Caption = ReadLangOpt(Contrl.name & ".Caption", Contrl.Caption)
    End If

    'If TypeOf Contrl Is Frame Then '                            Frame
    'Contrl.Caption = ReadLangOpt(Contrl.name & ".Caption", Contrl.Caption)
    'End If

    If TypeOf Contrl Is XpB Then    '                            XPB
        Contrl.Caption = ReadLangOpt(Contrl.name & ".Caption", Contrl.Caption)
        Contrl.pInitialize
        'Contrl.ToolTipText = ReadLangOpt(Contrl.name & ".ToolTip")
    End If

    'If TypeOf Contrl Is ComboBox Then '                         Combo
    'End If

    If TypeOf Contrl Is OptionButton Then    '                   OptionButton
        For i = 0 To 2
            If Contrl.name = "OptHtml" Then OptHtml(i).Caption = ReadLangOpt(Contrl.name & i & ".Caption", OptHtml(i).Caption)
            If Contrl.name = "optProxy" Then optProxy(i).Caption = ReadLangOpt(Contrl.name & i & ".Caption", optProxy(i).Caption)
        Next
    End If

    If TypeOf Contrl Is CheckBox Then    '                        CheckBox
        Contrl.Caption = ReadLangOpt(Contrl.name & ".Caption", Contrl.Caption)
    End If

Next    'Contrl


'                                                           lstExport
LstExport.Clear
With FrmMain
    For i = 0 To LstExport_ListCount - 1    'без аннотации
        LstExport.List(i) = .ListView.ColumnHeaders(i + 1).Text
    Next i
    LstExport.List(i) = frmEditor.LFilm(9).Caption    '+ аннотация
End With
LstExport.List(LstExport_ListCount + 1) = vbNullString ' пустышка для ловли селекта
On Error GoTo 0
LstExport.ListIndex = LstExport.ListCount


''                                                      tvOpt
FillTVOpt


''''''''''
OptCaption = ReadLang("VerticalMenu.MenuItemCaption6")

'                                                       NamesStore()
For i = 0 To UBound(NStoreOpt)
    NStoreOpt(i) = ReadLangOpt("NStoreOpt" & i)
Next i
'Tabs после NStoreOpt
TabSOpt.ImageList = imlOpt
TabSOpt.Tabs(1).Caption = ReadLangOpt("FrameBD.Caption", TabSOpt.Tabs(1).Caption)
TabSOpt.Tabs(1).Image = "Bases"
TabSOpt.Tabs(2).Caption = ReadLangOpt("FrFont.Caption", TabSOpt.Tabs(2).Caption)
TabSOpt.Tabs(2).Image = "Fonts"
TabSOpt.Tabs(3).Caption = ReadLangOpt("FrExport.Caption", TabSOpt.Tabs(3).Caption)
TabSOpt.Tabs(3).Image = "Export"
TabSOpt.Tabs(4).Caption = ReadLangOpt("FrGlobal.Caption", TabSOpt.Tabs(4).Caption)
TabSOpt.Tabs(4).Image = "Other"
TabSOpt.Tabs(5).Caption = ReadLangOpt("FrInet.Caption", TabSOpt.Tabs(5).Caption)
TabSOpt.Tabs(5).Image = "Internet"
TabSOpt.Tabs(6).Caption = ReadLangOpt("frCombo.Caption", TabSOpt.Tabs(6).Caption)
TabSOpt.Tabs(6).Image = "Combos"
TabSOpt.Tabs(7).Caption = NStoreOpt(9)    'о программе
TabSOpt.Tabs(7).Image = "About"


tSVCCurVer.Text = App.Major & "." & App.Minor & "." & App.Revision

Me.Icon = FrmMain.Icon

'LockWindowUpdate 0
'Screen.MousePointer = vbNormal
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DoEvents

'''''''''''''''из кнопки возврат'
'возврат
FrmOptions.Visible = False    'hide

If LstBases.ListIndex > -1 Then
    If Opt_AutoSaveOpt And (Not optsaved) Then ComOptSave_Click    ' автосохранение
    optReadIniFlag = True
    FrmMain.TabLVHid.Tabs(LstBases.ListIndex + 1).Selected = True
End If

'вернуть глобальное в переменные
Options2Var

optReadIniFlag = False
'''''''''''

FrmMain.TabLVHid.Enabled = True

With FrmMain
    If InitFlag And (Not NoDBFlag) Then
        If (opt_bdname = bdname) And (Len(bdname) <> 0) Then
            .TabLVHid.Tabs(CurrentBaseIndex).Selected = True
        End If
    End If

    If actInitFlag Then
        OpenActDB
        If .FrameActer.Visible Then .FillActListView
    End If
    If NoDBFlag Then
        If .FrameActer.Visible Then
            If Not LVActerFilled Then .FillActListView
        End If
    End If

End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmOptFlag = False
MakeNormal FrmMain.hwnd 'это выведет окно в фокус
'Call SendMessage(FrmMain.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Private Sub LstBases_Click()
If Len(LstBases.Text) = 0 Then Exit Sub

'если опции уже видны
If FrmOptions.Visible Then

    If delBaseFlag Then
        opt_bdname = LstBases.Text
        ReadINIOpt GetNameFromPathAndName(opt_bdname)
        Exit Sub
    End If

    opt_bdname = LstBases.Text '1
    'автосохранение
    If Opt_AutoSaveOpt Then ComOptSave_Click '2



    ReadINIOpt GetNameFromPathAndName(opt_bdname)

    '''@@If opt_bdname = bdname Then FrmMain.setForeColor

    LstBases_ListIndex = LstBases.ListIndex

End If

'подгрузить опции конкретных баз
Var2Options False

If Opt_UCLV_Vis Then FrmMain.UCLV.Refresh

LstExport.ListIndex = LstExport.ListCount - 1: LstExport.TopIndex = 0    'чтоб не видно выделения
'LstExport.Selected(LstExport.ListCount - 1) = False

'''optsaved = True 'хочется передать на запись опций и извне
End Sub
Private Sub ComOpenBD_Click()
'Dim cd As New cCommonDialog
Dim cd As cCommonDialog

Dim sfile As String
Dim i As Integer
Dim temp As String

Set cd = New cCommonDialog

If (cd.VBGetOpenFileName( _
    sfile, _
    Filter:="MDB Files (*.mdb)|*.mdb|All Files (*.*)|*.*", _
    FilterIndex:=1, _
    DefaultExt:="mdb", _
    Owner:=Me.hwnd)) Then
    temp = sfile
End If

DoEvents
ToDebug "Добавить базу..."

If (temp <> vbNullString) And (sfile <> vbNullString) Then

    For i = 0 To LstBases.ListCount - 1
        If LstBases.List(i) = temp Then
            ToDebug temp & "... уже есть в списке"
            'Set cd = Nothing
            Exit Sub    'повтор
        End If
    Next i

    'add to list
    LstBases.AddItem temp: SetListboxScrollbar LstBases

    '''в переменные
    LstBases_ListCount = LstBases_ListCount + 1
    ReDim Preserve LstBases_List(LstBases_ListCount)
    LstBases_List(LstBases_ListCount) = temp

    '''LstBases.Selected(LstBases.ListCount - 1) = True 'клик для подгрузки ини

    ''opt_bdname = temp

    Call AddTabsLV

    '''InitFlag = True
    NoDBFlag = False
    INIFileFlagRW = True    'для начала

    'FrmMain.ReadINI GetNameFromPathAndName(opt_bdname)
    NoSetColorFlag = False

    ''LabelCurrBDHid.Caption = opt_bdname
    ''LBDSizeHid.Caption = FileLen(opt_bdname) / 1024 & " K"
    ComCompact.Enabled = True

    ToDebug "... ok"
    ComOptSave.Enabled = True

    optsaved = False
    'Else
    pwd = vbNullString
    'End If 'open db


End If    'null ret

'Set cd = Nothing
End Sub

Private Sub ComOptSave_Click()
Dim WFD As WIN32_FIND_DATA
Dim ret As Long
Dim i As Integer


'If Not INIFileFlagRW Then ComOptSave.Enabled = False: Exit Sub

On Error Resume Next

'''If optsaved Then Exit Sub

Screen.MousePointer = vbHourglass

'                                       -------      Global
'check ini
iniFileName = App.Path
If Right$(iniFileName, 1) <> "\" Then iniFileName = iniFileName & "\"
iniFileName = iniFileName & "global.ini"
ret = FindFirstFile(iniFileName, WFD)
If ret < 0 Then
    MakeINI "global.ini"
Else
    '        ' проверить ридонли
    '        If WFD.dwFileAttributes And FILE_ATTRIBUTE_READONLY Then
    '            GlobalFileFlagRW = False
    '        Else
    '            GlobalFileFlagRW = True
    '        End If
End If

FindClose ret

ToDebug "SaveOptsIn global"
'выбранный из списка и не примененный язык не запомнится
WriteKey "Language", "LastLang", LastLanguage, iniFileName 'LastLanguage = ComboLangHid.Text


WriteKey "GLOBAL", "BDCount", LstBases.ListCount, iniFileName
For i = 0 To LstBases.ListCount - 1
    WriteKey "GLOBAL", "BDname" & i + 1, LstBases.List(i), iniFileName
Next i

'inet
WriteKey "GLOBAL", "UseProxy", CStr(Opt_InetUseProxy), iniFileName
WriteKey "GLOBAL", "ProxyServerPort", Opt_InetProxyServerPort, iniFileName
WriteKey "GLOBAL", "ProxyUserName", Opt_InetUserName, iniFileName
'WriteKey "GLOBAL", "ProxyPassword", Opt_InetPassword, iniFileName
WriteKey "GLOBAL", "ProxySecure", CStr(Opt_InetSecureFlag), iniFileName

'общие галочки
WriteKey "GLOBAL", "GetMediaType", CStr(Opt_GetMediaType), iniFileName
WriteKey "GLOBAL", "AviDirectShow", CStr(Opt_AviDirectShow), iniFileName
WriteKey "GLOBAL", "CancelAutoPlay", CStr(Opt_QueryCancelAutoPlay), iniFileName
WriteKey "GLOBAL", "GetVolumeInfo", CStr(Opt_GetVolumeInfo), iniFileName


'
If Len(opt_bdname) = 0 Then    'не выбрана база - выход
    ComOptSave.Enabled = False
    optsaved = True
    Screen.MousePointer = vbNormal
    Exit Sub
End If


'                                    -------     `             Current Base ini
iniFileName = App.Path
If Right$(iniFileName, 1) <> "\" Then iniFileName = iniFileName & "\"
iniFileName = iniFileName & ini_opt
ret = FindFirstFile(iniFileName, WFD)
If ret < 0 Then MakeINI ini_opt
FindClose ret

ToDebug "SaveOptsIn " & ini_opt


'еще tv галочки
WriteKey "GLOBAL", "SaveOptOnExit", CStr(Opt_AutoSaveOpt), iniFileName
WriteKey "GLOBAL", "SortOnStart", CStr(Opt_SortOnStart), iniFileName
WriteKey "LIST", "ColorDebt", CStr(Opt_Debtors_Colorize), iniFileName
WriteKey "GLOBAL", "LVLoadOnlyTitle", CStr(Opt_LoadOnlyTitles), iniFileName
WriteKey "GLOBAL", "LoanAllSameLabels", CStr(Opt_LoanAllSameLabels), iniFileName
WriteKey "GLOBAL", "SaveBigPix", CStr(Opt_PicRealRes), iniFileName
WriteKey "GLOBAL", "UseAspect", CStr(Opt_UseAspect), iniFileName
WriteKey "GLOBAL", "FreeDVDFilters", CStr(Opt_UseOurMpegFilters), iniFileName
WriteKey "GLOBAL", "SlideShowWindow", CStr(Opt_NoSlideShow), iniFileName
WriteKey "GLOBAL", "CenterShowPic", CStr(Opt_CenterShowPic), iniFileName
WriteKey "GLOBAL", "ListAndInfo", CStr(Opt_UCLV_Vis), iniFileName
WriteKey "GLOBAL", "GroupWindow", CStr(Opt_Group_Vis), iniFileName
WriteKey "COVER", "ShowColNames", CStr(Opt_ShowColNames), iniFileName

WriteKey "GLOBAL", "LVEDIT", CStr(Opt_LVEDIT), iniFileName
WriteKey "GLOBAL", "SaveFileWithPath", CStr(Opt_FileWithPath), iniFileName

WriteKey "GLOBAL", "SortLVAfterEdit", CStr(Opt_SortLVAfterEdit), iniFileName
WriteKey "GLOBAL", "SortLabelAsNum", CStr(Opt_SortLabelAsNum), iniFileName
WriteKey "GLOBAL", "PutOtherInAnnot", CStr(Opt_PutOtherInAnnot), iniFileName


'                                                                          LV
WriteKey "LIST", "LVGrid", CStr(Opt_ShowLVGrid), iniFileName
'                                                                   сплиттеры
WriteKey "LIST", "LVWidth%", CStr(LVWidth), iniFileName
WriteKey "LIST", "TVWidth", CStr(TVWidth), iniFileName    '                     group
WriteKey "LIST", "ScrShotWidth%", CStr(SplitLVD), iniFileName    '           SS




'---------

WriteKey "CD", "CDdrive", ComboCDHid, iniFileName

'                                                                           Font
'With FrmMain
    WriteKey "FONT", "VFontName", FontVert.name, iniFileName
    WriteKey "FONT", "VFontSize", FontVert.Size, iniFileName
    WriteKey "FONT", "VFontBold", FontVert.Bold, iniFileName
    WriteKey "FONT", "VFontItalic", FontVert.Italic, iniFileName
    WriteKey "FONT", "VFontColor", str$(VFontColor), iniFileName

    WriteKey "FONT", "HFontName", FontHor.name, iniFileName
    WriteKey "FONT", "HFontSize", FontHor.Size, iniFileName
    WriteKey "FONT", "HFontBold", FontHor.Bold, iniFileName
    WriteKey "FONT", "HFontItalic", FontHor.Italic, iniFileName
    WriteKey "FONT", "HFontColor", str$(HFontColor), iniFileName

    WriteKey "FONT", "LVFontName", FontListView.name, iniFileName
    WriteKey "FONT", "LVFontSize", FontListView.Size, iniFileName
    WriteKey "FONT", "LVFontBold", FontListView.Bold, iniFileName
    WriteKey "FONT", "LVFontItalic", FontListView.Italic, iniFileName
    WriteKey "FONT", "LVFontColor", str$(LVFontColor), iniFileName

    WriteKey "FONT", "LVBackColor", str$(LVBackColor), iniFileName
    WriteKey "FONT", "CoverHorBackColor", str$(CoverHorBackColor), iniFileName
    WriteKey "FONT", "CoverVertBackColor", str$(CoverVertBackColor), iniFileName
'End With
    WriteKey "FONT", "LVHighLight", str$(LVHighLightLong), iniFileName

WriteKey "FONT", "VMcolor", CStr(VMSameColor), iniFileName
WriteKey "FONT", "StripedLV", CStr(StripedLV), iniFileName
WriteKey "FONT", "NoLVSelFrame", CStr(NoLVSelFrame), iniFileName
WriteKey "LIST", "Opt_ShowLVGrid", CStr(Opt_ShowLVGrid), iniFileName


'End If 'Not frmOptFlag Then

'                                                                   Export
For i = 0 To LstExport.ListCount - 2    '(- пустышку)
    WriteKey "EXPORT", "L" & i, LstExport.Selected(i), iniFileName
Next
'                                                                   имена картинок
For i = OptHtml.LBound To OptHtml.UBound
    DeleteKey "OptHtml" & i & ".Caption", "EXPORT", iniFileName
    If OptHtml(i).Value = True Then
        WriteKey "EXPORT", "OptHtml" & i & ".Caption", True, iniFileName
    End If
Next
'                                                               Разделитель полей
WriteKey "EXPORT", "ExportDelimiter", """" & tExpDelim & """", iniFileName


WriteKey "EXPORT", "Template", CombTemplate.Text, iniFileName
WriteKey "EXPORT", "NumsOnPage", TxtNnOnPage.Text, iniFileName

'сабфолдеры html
WriteKey "EXPORT", "UseSubFolders", CStr(Opt_ExpUseFolders), iniFileName
WriteKey "EXPORT", "SubFolder1", tExpFolders(0).Text, iniFileName
WriteKey "EXPORT", "SubFolder2", tExpFolders(1).Text, iniFileName
WriteKey "EXPORT", "SubFolder3", tExpFolders(2).Text, iniFileName

WriteKey "GLOBAL", "QJPG", TextQJPGHid, iniFileName


optsaved = True
Screen.MousePointer = vbNormal

If err.Number = 0 Then
    ToDebug "...ok"
Else
    ToDebug "...err: " & err.Description
End If
Call err.Clear
End Sub


Private Sub ComFontV_Click()
'Dim cd As New cCommonDialog
Dim cd As cCommonDialog

Dim sFnt As StdFont
Dim temp As String

With FrmMain
    Set sFnt = FontVert

    If VFontColor = 0 Then VFontColor = 1

    Set cd = New cCommonDialog

    If (cd.VBChooseFont(sFnt, , Me.hwnd, VFontColor, , , CF_NoOemFonts Or CF_ScalableOnly)) Then

        'FontVert = sFnt

        'font
        temp = " " & FontVert.Size
        If FontVert.Bold Then temp = temp + " Bold"
        If FontVert.Italic Then temp = temp + " Italic"
        TextFontV.Text = FontVert.name & temp
    End If

End With

setForeColorOpt


RefreshCover
ApplyOpt
optsaved = False

'Set cd = Nothing
Set sFnt = Nothing
End Sub
Private Sub RefreshCover()
'рефрешить обложку
With FrmMain
If .FrameCover.Visible Then
    .TabStripCover_Click
End If
End With
End Sub
Private Sub setForeColorOpt()
Dim temp As String

'цвета
TextFontV.BackColor = CoverVertBackColor
TextFontV.ForeColor = VFontColor

TextFontH.BackColor = CoverHorBackColor
TextFontH.ForeColor = HFontColor

TextFontLV.BackColor = LVBackColor
TextFontLV.ForeColor = LVFontColor

'в текст шрифта
With FrmMain
 temp = " " & FontVert.Size
 If FontVert.Bold Then temp = temp + " Bold"
 If FontVert.Italic Then temp = temp + " Italic"
 TextFontV.Text = FontVert.name & temp
 '
 temp = " " & FontHor.Size
 If FontHor.Bold Then temp = temp + " Bold"
 If FontHor.Italic Then temp = temp + " Italic"
 TextFontH.Text = FontHor.name & temp
 '
 temp = " " & FontListView.Size
 If FontListView.Bold Then temp = temp + " Bold"
 If FontListView.Italic Then temp = temp + " Italic"
 TextFontLV.Text = FontListView.name & temp
End With


End Sub

Private Sub ComCoverVertFillColor_Click()
Dim c As Long
'Dim cd As New cCommonDialog
Dim cd As cCommonDialog
Dim ret As Long

Set cd = New cCommonDialog

cd.CustomColor(0) = CoverVertBackColor
ret = cd.VBChooseColor(c, True, True, False, Me.hwnd)
If ret = 0 Then Exit Sub
CoverVertBackColor = c
'PicCoverPaper.Line (35, 145)-(172, 262), c, BF

'Set cd = Nothing
setForeColorOpt

RefreshCover
ApplyOpt

optsaved = False
End Sub
Private Sub ComFontH_Click()
'Dim cd As New cCommonDialog
Dim cd As cCommonDialog

Dim sFnt As StdFont
Dim temp As String

With FrmMain

    Set sFnt = FontHor
    If HFontColor = 0 Then HFontColor = 1

    Set cd = New cCommonDialog
    If (cd.VBChooseFont(sFnt, , Me.hwnd, HFontColor)) Then

        'FontHor = sFnt

        temp = " " & FontHor.Size
        If FontHor.Bold Then temp = temp + " Bold"
        If FontHor.Italic Then temp = temp + " Italic"
        TextFontH.Text = FontHor.name & temp

    End If

End With

setForeColorOpt

RefreshCover
ApplyOpt

optsaved = False
'Set cd = Nothing
Set sFnt = Nothing
End Sub
Private Sub ComCoverHorFillColor_Click()
Dim c As Long
Dim cd As cCommonDialog
Dim ret As Long

Set cd = New cCommonDialog
cd.CustomColor(0) = CoverHorBackColor
ret = cd.VBChooseColor(c, True, True, False, Me.hwnd)
If ret = 0 Then Exit Sub
CoverHorBackColor = c

'Set cd = Nothing
setForeColorOpt

RefreshCover
ApplyOpt

optsaved = False
End Sub
Private Sub ComFontLV_Click()
Dim cd As cCommonDialog
Dim sFnt As StdFont
Dim temp As String

With FrmMain

Set sFnt = FontListView
'Set sFnt = New StdFont
'sFnt.name = FontListView.name
'sFnt.Size = FontListView.Size

If LVFontColor = 0 Then LVFontColor = 1

Set cd = New cCommonDialog
If (cd.VBChooseFont(sFnt, , Me.hwnd, LVFontColor, 6, 16)) Then
 
temp = " " & FontListView.Size
If FontListView.Bold Then temp = temp + " Bold"
If FontListView.Italic Then temp = temp + " Italic"
TextFontLV.Text = FontListView.name & temp
'Set TextFontLV.Font = FontListView

 End If

End With

NoSetColorFlag = False

setForeColorOpt
ApplyOpt

optsaved = False
Set cd = Nothing
Set sFnt = Nothing
End Sub

Private Sub ComColorPick_Click()
Dim c As Long
Dim cd As cCommonDialog
Dim ret As Long

Set cd = New cCommonDialog
cd.CustomColor(0) = LVBackColor

ret = cd.VBChooseColor(c, True, True, False, Me.hwnd)
If ret = 0 Then Exit Sub
LVBackColor = c

NoSetColorFlag = False

setForeColorOpt
ApplyOpt

optsaved = False
Set cd = Nothing
End Sub
Private Sub ComDelBD_Click()

If Not ComDelBD_RealClick Then Exit Sub
ComDelBD_RealClick = False 'для RealClick до того дб нажаты vousedown или keydown

optsaved = False


If LstBases.ListCount = 0 Then Exit Sub
If opt_bdname = bdname Then NoBaseClear
delBaseFlag = True

If LstBases.ListIndex > -1 Then LstBases.RemoveItem LstBases.ListIndex

If (LstBases.ListCount > 0) Then
    LstBases.Selected(0) = True '1 - клик
    opt_bdname = LstBases.List(0) '2
    CurrentBaseIndex = 1
    LabelCurrBDHid.Caption = opt_bdname
    Me.Caption = OptCaption & ": " & opt_bdname
    If Len(opt_bdname) <> 0 Then LBDSizeHid.Caption = FileLen(opt_bdname) / 1024 & " K"
    ComCompact.Enabled = True
  
Else
    'нет баз
    opt_bdname = vbNullString
    LabelCurrBDHid.Caption = vbNullString
    Me.Caption = OptCaption
    LBDSizeHid.Caption = vbNullString
    ini_opt = vbNullString
    ComCompact.Enabled = False
    NoDBFlag = True
    InitFlag = True
    ToDebug "Нет баз..."
    FrmMain.VerticalMenu_MenuItemClick 5, 0 'goto Actors
End If

'''в переменные
Dim i As Integer
LstBases_ListCount = LstBases.ListCount
ReDim LstBases_List(LstBases_ListCount)
For i = 1 To LstBases_ListCount
LstBases_List(i) = LstBases.List(i - 1)
Next i


'переделать табы
Call AddTabsLV
    'кликнуть на новый
    '''FrmMain.TabLVHid.Tabs(LstBases.ListIndex + 1).Selected = True

delBaseFlag = False
If opt_bdname <> bdname Then InitFlag = True

End Sub
Private Sub ComNewBD_Click()
Dim a() As Byte
Dim fn As Integer
Dim cd As cCommonDialog
Dim sfile As String
  
ToDebug "Создать новую базу фильмов..."
Set cd = New cCommonDialog

   If (cd.VBGetSaveFileName( _
      sfile, _
      Filter:="MDB (*.mdb)|*.mdb|All Files (*.*)|*.*", _
      FilterIndex:=1, _
      DefaultExt:="mdb", _
      Owner:=Me.hwnd)) Then
   End If
   
a() = LoadResData("SVC", "CUSTOM")

fn = FreeFile
If sfile <> vbNullString Then

'Set cd = Nothing
Open sfile For Binary Access Write As fn
     Put #fn, , a()
Close #fn

ToDebug "...ok"
End If
End Sub


Private Sub ComCompact_Click()
Dim tempDBname As String

If BaseReadOnly Then myMsgBox msgsvc(24), vbInformation, , Me.hwnd: Exit Sub
If BaseReadOnlyU Then myMsgBox msgsvc(22), vbInformation, , Me.hwnd: Exit Sub

ToDebug "Сжатие базы фильмов " & opt_bdname
'rs.Close: DB.Close
Set rs = Nothing
Set DB = Nothing
Screen.MousePointer = vbHourglass

tempDBname = App.Path + "\svcmdb.tmp"

On Error Resume Next 'password
If Dir(opt_bdname) <> vbNullString Then
 If Dir(tempDBname) <> vbNullString Then Kill tempDBname
 DBEngine.CompactDatabase opt_bdname, tempDBname
 If err.Number = 3031 Then 'пассворд
  SetTimer hwnd, NV_INPUTBOX, 10, AddressOf TimerProc
  pwd = myInputBox(NamesStore(5) & vbCrLf & opt_bdname)
  On Error GoTo err
  DBEngine.CompactDatabase opt_bdname, tempDBname, , , ";PWD=" & pwd
 End If

 If Dir(tempDBname) <> vbNullString Then Kill opt_bdname
 If Dir(opt_bdname) = vbNullString Then Name tempDBname As opt_bdname
 If Dir(opt_bdname) = vbNullString Then FileCopy tempDBname, opt_bdname: ToDebug "...ok"
End If

LBDSizeHid = FileLen(opt_bdname) / 1024 & " K"

CurSearch = 0 'потом будет 1
InitFlag = True

OpenNewDataBase

Screen.MousePointer = vbNormal

Exit Sub
err:
Screen.MousePointer = vbNormal
ToDebug "..." & err.Description
MsgBox err.Description, vbCritical
End Sub

Private Sub ComCompactA_Click()
Dim tempDBname As String
If BaseAReadOnly Then myMsgBox msgsvc(25), vbInformation, , Me.hwnd: Exit Sub
If BaseAReadOnlyU Then myMsgBox msgsvc(23), vbInformation, , Me.hwnd: Exit Sub

ToDebug "Сжатие базы актеров..."
'ars.Close: ADB.Close
Set ars = Nothing
Set ADB = Nothing

Screen.MousePointer = vbHourglass

tempDBname = App.Path + "\pmdb.tmp"

On Error GoTo err
If Dir(abdname) <> vbNullString Then
 If Dir(tempDBname) <> vbNullString Then Kill tempDBname
 DBEngine.CompactDatabase abdname, tempDBname
 If Dir(tempDBname) <> vbNullString Then Kill abdname
 If Dir(abdname) = vbNullString Then Name tempDBname As abdname
 If Dir(abdname) = vbNullString Then FileCopy tempDBname, abdname: ToDebug "...ok"
End If

LABDSizeHid.Caption = FileLen(abdname) / 1024 & " K"
Set ADB = DBEngine.OpenDatabase(abdname, False)

'LVActerFilled = False
subActFiltCancel
''там (ChActOnlyFoto) произойдет Set ars и заполнение списка, если смена
' FrmMain.OptActOnlyFotoHid(0).Value = True
 
If ars Is Nothing Then
 Set ars = ADB.OpenRecordset("Acter", dbOpenTable)
 ars.Index = "KeyAct"
End If

actInitFlag = True
Screen.MousePointer = vbNormal

Exit Sub
err:
Screen.MousePointer = vbNormal
MsgBox err.Description, vbCritical
ToDebug "..." & err.Description
End Sub


Private Sub LstOpt2Var()
'передача галочек опций в переменные
On Error Resume Next
'

Opt_AutoSaveOpt = CBool(tvOpt.Nodes("Opt_AutoSaveOpt").Checked)
Opt_SortOnStart = CBool(tvOpt.Nodes("Opt_SortOnStart").Checked)
Opt_Debtors_Colorize = CBool(tvOpt.Nodes("Opt_Debtors_Colorize").Checked)
Opt_LoadOnlyTitles = CBool(tvOpt.Nodes("Opt_LoadOnlyTitles").Checked)
Opt_LoanAllSameLabels = CBool(tvOpt.Nodes("Opt_LoanAllSameLabels").Checked)
Opt_PicRealRes = CBool(tvOpt.Nodes("Opt_PicRealRes").Checked)
Opt_UseAspect = CBool(tvOpt.Nodes("Opt_UseAspect").Checked)
Opt_UseOurMpegFilters = CBool(tvOpt.Nodes("Opt_UseOurMpegFilters").Checked)
Opt_AviDirectShow = CBool(tvOpt.Nodes("Opt_AviDirectShow").Checked)
Opt_NoSlideShow = CBool(tvOpt.Nodes("Opt_NoSlideShow").Checked)
Opt_CenterShowPic = CBool(tvOpt.Nodes("Opt_CenterShowPic").Checked)
Opt_UCLV_Vis = CBool(tvOpt.Nodes("Opt_UCLV_Vis").Checked)
Opt_Group_Vis = CBool(tvOpt.Nodes("Opt_Group_Vis").Checked)
Opt_ShowColNames = CBool(tvOpt.Nodes("Opt_ShowColNames").Checked)
Opt_GetMediaType = CBool(tvOpt.Nodes("Opt_GetMediaType").Checked)
Opt_GetVolumeInfo = CBool(tvOpt.Nodes("Opt_GetVolumeInfo").Checked)
Opt_QueryCancelAutoPlay = CBool(tvOpt.Nodes("Opt_QueryCancelAutoPlay").Checked)

Opt_LVEDIT = CBool(tvOpt.Nodes("Opt_LVEDIT").Checked)
Opt_FileWithPath = CBool(tvOpt.Nodes("Opt_FileWithPath").Checked)

Opt_SortLVAfterEdit = CBool(tvOpt.Nodes("Opt_SortLVAfterEdit").Checked)
Opt_SortLabelAsNum = CBool(tvOpt.Nodes("Opt_SortLabelAsNum").Checked)
Opt_PutOtherInAnnot = CBool(tvOpt.Nodes("Opt_PutOtherInAnnot").Checked)


'в меню почекать
FrmMain.mnuCard.Checked = Opt_UCLV_Vis
FrmMain.mnuGroup.Checked = Opt_Group_Vis

ApplyOpt
If opt_bdname = bdname Then FrmMain.Form_Resize

End Sub


Private Sub ComboCDHid_Change() 'TxtNnOnPage_Text
If FrmOptions.Visible Then optsaved = False
oldComboCDHid = ComboCDHid
End Sub

Private Sub ComboCDHid_DropDown()
Call drives
optsaved = False
End Sub

Private Sub LstBases_DragDrop(Source As Control, X As Single, Y As Single)
Dim CurTabTxt As String
Dim i As Integer

On Error Resume Next

ListRowMove Source, DragIndex, ListRowCalc(Source, Y)

'в массив
For i = 1 To LstBases_ListCount
    LstBases_List(i) = LstBases.List(i - 1)
Next i

With FrmMain
    'какой таб активен? oldTabLVInd
    CurTabTxt = .TabLVHid.Tabs(.TabLVHid.SelectedItem.Index).Caption

    'переделать табы
    Call AddTabsLV

    'восстановить кликнутую кнопку таба
    For i = 1 To .TabLVHid.Tabs.Count
        If .TabLVHid.Tabs(i).Caption = CurTabTxt Then
            .TabLVHid.Tabs(i).Selected = True
            'чтоб не перечитывать
            oldTabLVInd = i
            Exit For
        End If
    Next i
End With

'кликнуть на бывший помеченный таб
'нет FrmMain.TabLVHid.Tabs(curTab).Selected = True
'If (LstBases.ListIndex + 1) = curTab Then
'If CurrentBaseIndex = DragIndex + 1 Then
'FrmMain.TabLVHid.Tabs(LstBases.ListIndex + 1).Selected = True
'If curTab <> j + 1 And curTab <> DragIndex Then
''If curTab <> j + 1 Then
'If curTab <> j + 1 And DragIndex = j + 1 Then
'FrmMain.TabLVHid.Tabs(j + 1).Selected = True

'Debug.Print
'Debug.Print "curTab = " & curTab
'Debug.Print "s = " & s
'Debug.Print "CurrentBaseIndex = " & CurrentBaseIndex
'Debug.Print "oldTabLVInd = " & oldTabLVInd
'Debug.Print "DragIndex = " & DragIndex
'Debug.Print
End Sub

Private Sub LstBases_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    DragIndex = ListRowCalc(LstBases, Y)
    LstBases.Drag
End If

End Sub

Private Sub lstComboName_Click()
'составить список значений выбранной секции
Dim selItem As Integer

On Error GoTo err

lstComboVal.Clear
lstComboNames.Clear

If lstComboName.ListCount < 1 Then Exit Sub

selItem = lstComboName.ListIndex
If selItem < 0 Then selItem = 0

Select Case selItem
Case 0: UserIniCurrSection = "Genre" 'жанр
Case 1: UserIniCurrSection = "Country"       'страна
Case 2: UserIniCurrSection = "Language"       'язык
Case 3: UserIniCurrSection = "Subtitle"       'субтитры
Case 4: UserIniCurrSection = "Comments"       'качество
Case 5: UserIniCurrSection = "Media"      'носитель
Case 6: UserIniCurrSection = "Site"       'сайты
End Select

Call FillUserCombo(UserIniCurrSection, lstComboVal)

lstComboName.Selected(selItem) = True

Exit Sub

err:
ToDebug "lCN_Err: " & err.Description
End Sub

Private Sub lstComboVal_Click()
If lstComboVal.ListIndex < 0 Then Exit Sub
txtComboView = lstComboVal.List(lstComboVal.ListIndex)
End Sub

Private Sub LstExport_LostFocus()
On Error Resume Next
    If LstExport.ListCount > 1 Then
    LstExport.ListIndex = LstExport.ListCount - 1: LstExport.TopIndex = 0 'чтобы не было видно выделения
    End If
End Sub


Private Sub lstOptPreset_Click()
If lstOptPreset.ListIndex < 0 Then Exit Sub
Call GetExportPreset(lstOptPreset.ListIndex)
End Sub
Public Sub GetExportPreset(ind As Long)
Dim tmp As String
Dim tmpArr() As String
Dim i As Integer

'читать userFile
tmp = VBGetPrivateProfileString("ExportPreset", lstOptPreset.List(ind), userFile)
If Len(tmp) = 0 Then Exit Sub

tmpArr = Split(tmp, ",")

If UBound(tmpArr) = LstExport_ListCount Then
    For i = 0 To LstExport_ListCount
        LstExport.Selected(i) = CBool(tmpArr(i))
        'LstExport_Arr(i) = CBool(tmpArr(i))
    Next i
End If
LstExport.ListIndex = LstExport.ListCount - 1: LstExport.TopIndex = 0

'RefreshCover - уже, при селекте
End Sub

Private Sub ReadExportPresets()
'читать из user.lng список пресетов в lstOptPreset
Dim arrPresets() As String
Dim i As Long

If Not FileExists(userFile) Then Exit Sub
If GetKeyNames("ExportPreset", userFile, arrPresets) > -1 Then
    lstOptPreset.Clear
    For i = 0 To UBound(arrPresets)
        lstOptPreset.AddItem arrPresets(i)
    Next i
End If
End Sub
Private Sub optProxy_Click(Index As Integer)
If FrmOptions.Visible Then optsaved = False
Opt_InetUseProxy = Index
End Sub

Private Sub TabSOpt_Click()
Dim l As Integer, t As Integer
l = 60 'Me.Width / 2 - 3696 '60 '1000
t = 700 '780 '480
'l = 1080
FrameBD.Visible = False
FrGlobal.Visible = False
FrFont.Visible = False
FrExport.Visible = False
frInet.Visible = False
frCombo.Visible = False

Select Case TabSOpt.SelectedItem.Index
Case 1
    FrameBD.Move l, t: FrameBD.Visible = True: TabSOptLast = 1
    FrameBD.Enabled = True
    If addflag Or editFlag Then 'нет смены баз при редактировании
        FrameBD.Enabled = False
    End If
Case 2: FrFont.Move l, t: FrFont.Visible = True: TabSOptLast = 2
Case 3
    FrExport.Move l, t: FrExport.Visible = True: TabSOptLast = 3
Case 4: FrGlobal.Move l, t: FrGlobal.Visible = True: TabSOptLast = 4
Case 5: frInet.Move l, t: frInet.Visible = True: TabSOptLast = 5
Case 6: frCombo.Move l, t: frCombo.Visible = True: Call lstComboName_Click: TabSOptLast = 6

Case 7 'About
 TabSOpt.Tabs(TabSOptLast).Selected = True
 If myMsgBox("V " & App.Major & "." & App.Minor & "." & App.Revision & " (C)Lebedev Alexander aka SuR " & msgsvc(7), vbOKCancel, , FrmOptions.hwnd) = vbOK Then
 Shell "notepad.exe " & App.Path & "\readme.txt", vbNormalFocus
 End If
 
End Select

End Sub

Private Sub tExpDelim_Change()
If FrmOptions.Visible Then optsaved = False
End Sub


Private Sub tExpDelim_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ApplyOpt
End If
End Sub

Private Sub tExpDelim_LostFocus()
If Not optsaved Then ApplyOpt
End Sub

Private Sub tExpFolders_Change(Index As Integer)
If FrmOptions.Visible Then optsaved = False
End Sub

Private Sub TextQJPGHid_LostFocus()
If (Val(TextQJPGHid) < 1) Or (Val(TextQJPGHid) > 100) Then TextQJPGHid = 80
QJPG = TextQJPGHid
optsaved = False
End Sub

Private Sub ComboLangHid_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub ComboLangHid_LostFocus()
'get language from ini
Dim temp As String
Dim i As Integer

'ComboLangHid.Clear
'LangCount = Int(Val(VBGetPrivateProfileString("Language", "LCount", iniFileName)))
LangCount = Int(Val(VBGetPrivateProfileString("Language", "LCount", iniGlobalFileName)))

If LangCount < 1 Then LangCount = 1
For i = 1 To LangCount
temp = VBGetPrivateProfileString("Language", "L" & i, iniGlobalFileName)
If temp = ComboLangHid.Text Then lngFileName = App.Path + "\" + _
VBGetPrivateProfileString("Language", "L" & i & "File", iniGlobalFileName)
Next i
optsaved = False
End Sub


Private Sub ComMarkAll_Click()

If LstExport.ListCount < 1 Then Exit Sub
optsaved = False
Call SendMessage(LstExport.hwnd, LB_SETSEL, True, ByVal -1)
LstExport.Selected(LstExport.ListCount - 1) = False 'не выд пустышку
LstExport2Var

'снять селект с пресетов
If lstOptPreset.ListIndex > -1 Then
'lstOptPreset.List (lstOptPreset.ListIndex)
lstOptPreset.Selected(lstOptPreset.ListIndex) = False
End If
End Sub
Private Sub ComUnmarkAll_Click()
If LstExport.ListCount < 1 Then Exit Sub
optsaved = False
Call SendMessage(LstExport.hwnd, LB_SETSEL, False, ByVal -1)
LstExport2Var

'снять селект с пресетов
If lstOptPreset.ListIndex > -1 Then
'lstOptPreset.List (lstOptPreset.ListIndex)
lstOptPreset.Selected(lstOptPreset.ListIndex) = False
End If
End Sub

Private Sub LstExport_ItemCheck(Item As Integer)
If Item = LstExport.ListCount - 1 Then
LstExport.Selected(Item) = False
Exit Sub
End If

If Not FillLstProgFlag Then LstExport2Var: optsaved = False
End Sub
Private Sub LstExport2Var()
Dim i As Integer

'Поля экспорта в массив полей для экспорта
For i = 0 To LstExport_ListCount
LstExport_Arr(i) = LstExport.Selected(i)
Next i

RefreshCover

End Sub

Private Sub OptHtml_Click(Index As Integer)
If FrmOptions.Visible Then
optsaved = False
Opt_HtmlJpgName = Index
End If
End Sub

Private Sub CombTemplate_Click()
If FrmOptions.Visible Then optsaved = False
End Sub
Private Sub CombTemplate_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub tProxyPass_Change()
If FrmOptions.Visible Then optsaved = False
Opt_InetPassword = tProxyPass
End Sub

Private Sub tProxyServerPort_Change()
If FrmOptions.Visible Then optsaved = False
Opt_InetProxyServerPort = tProxyServerPort
End Sub

Private Sub tProxyUserName_Change()
If FrmOptions.Visible Then optsaved = False
Opt_InetUserName = tProxyUserName
End Sub



Private Sub tvOpt_Collapse(ByVal Node As MSComctlLib.Node)
       Node.Image = "Plus"
End Sub

Private Sub tvOpt_DblClick()
'Dim sItem As String

If tvOpt.SelectedItem.Children > 0 Then
    'This node has children, so we ignore the dbl -Click
    '- в своих событиях    сменить ему иконку - раскрыто-закрыто
'    If tvOpt.SelectedItem.Expanded Then
'        tvOpt.SelectedItem.Image = "Minus"
'    Else
'        tvOpt.SelectedItem.Image = "Plus"
'    End If
Else

   ' Debug.Print tvOpt.SelectedItem.Index
    'This node doesn't have children, so we process check
    tvOpt.SelectedItem.Checked = Not tvOpt.SelectedItem.Checked
        If tvOpt.SelectedItem.Checked Then
        tvOpt.SelectedItem.Image = "Yes"
    Else
        tvOpt.SelectedItem.Image = "No" '0
    End If
    
    '
    LstOpt2Var
End If

optsaved = False
End Sub
Private Sub tvOpt_Check2Ico()
'по чеку менять иконку
Dim n As Node
For Each n In tvOpt.Nodes
If n.Children = 0 Then
'бездетный - наш клиент
    If n.Checked Then
        n.Image = "Yes"
    Else
        n.Image = "No" '0
    End If
End If
Next
End Sub

Private Sub tvOpt_Expand(ByVal Node As MSComctlLib.Node)
Node.Image = "Minus"
End Sub



Private Sub tvOpt_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
If tvOpt.SelectedItem.Children = 0 Then
    tvOpt.SelectedItem.Checked = Not tvOpt.SelectedItem.Checked
        If tvOpt.SelectedItem.Checked Then
        tvOpt.SelectedItem.Image = "Yes"
    Else
        tvOpt.SelectedItem.Image = "No" '0
    End If
    
    '
    LstOpt2Var
    
End If
End If

End Sub


Private Sub TxtNnOnPage_Change()
If Not IsNumeric(TxtNnOnPage.Text) Then TxtNnOnPage.Text = 30: Exit Sub
If (Val(TxtNnOnPage.Text) < 1) Then TxtNnOnPage.Text = 30
If FrmOptions.Visible Then optsaved = False
End Sub
' Set the list box's horizontal extent
Public Sub SetListboxScrollbar(lB As ListBox)
Dim i As Integer
Dim new_len As Long
Dim max_len As Long

For i = 0 To lB.ListCount - 1
 new_len = 10 + ScaleX(TextWidth(lB.List(i)), ScaleMode, vbPixels)
 If max_len < new_len Then max_len = new_len
Next i

SendMessage lB.hwnd, LB_SETHORIZONTALEXTENT, max_len, 0
End Sub
Private Sub ReadINIOpt(iFn As String)
Dim WFD As WIN32_FIND_DATA
Dim ret As Long
Dim temp As String
Dim i As Integer ', tmpi As Integer 'mzt , j As Integer
Dim glob As String
'mzt Dim flag As Boolean

ToDebug "ReadIniFromOptBases:" & ini_opt

With FrmMain

glob = "global"
ini_opt = iFn & ".ini"
'check ini
iniFileName = App.Path
If Right$(iniFileName, 1) <> "\" Then iniFileName = iniFileName & "\"
iniFileName = iniFileName & ini_opt
ret = FindFirstFile(iniFileName, WFD)

If iFn = glob Then
    'GlobalFileFlagRW = True
    If ret < 0 Then
        If MakeINI(ini_opt) Then GlobalFileFlagRW = True
    Else
        ' проверить ридонли
        If WFD.dwFileAttributes And FILE_ATTRIBUTE_READONLY Then
            GlobalFileFlagRW = False
        Else
            GlobalFileFlagRW = True
        End If
    End If
Else
    'INIFileFlagRW = True
    If ret < 0 Then
        'если нет спец инишника базы, сделать и перечитать
        If MakeINI(ini_opt) Then INIFileFlagRW = True
        ReadINIOpt GetNameFromPathAndName(opt_bdname): Exit Sub
    Else
        ' проверить ридонли
        If WFD.dwFileAttributes And FILE_ATTRIBUTE_READONLY Then
            INIFileFlagRW = False
        Else
            INIFileFlagRW = True
        End If
    End If
End If
FindClose ret


If iFn <> glob Then                                                 'Конкретная база ---- только

 'сортировать ли
 temp = VBGetPrivateProfileString("GLOBAL", "SortOnStart", iniFileName)
 If Len(temp) <> 0 Then Opt_SortOnStart = CBool(temp) Else Opt_SortOnStart = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Фонты
'vert
temp = VBGetPrivateProfileString("FONT", "VFontName", iniFileName)
    If LenB(temp) = 0 Then
        FontVert.name = "Arial"
    Else
        FontVert.name = temp
    End If
temp = VBGetPrivateProfileString("FONT", "VFontSize", iniFileName)
    If Not IsNumeric(temp) Then
        FontVert.Size = 12
    Else
        FontVert.Size = CDbl(Replace2Regional(temp))
    End If
temp = VBGetPrivateProfileString("FONT", "VFontBold", iniFileName)
    FontVert.Bold = False
    If StrComp(temp, "true", vbTextCompare) = 0 Then FontVert.Bold = True
temp = VBGetPrivateProfileString("FONT", "VFontItalic", iniFileName)
    FontVert.Italic = False
    If StrComp(temp, "true", vbTextCompare) = 0 Then FontVert.Italic = True
temp = VBGetPrivateProfileString("FONT", "VFontColor", iniFileName)
    If Not IsNumeric(temp) Then
        VFontColor = 0
    Else
        VFontColor = Val(temp)
    End If
    
'hor cover font
temp = VBGetPrivateProfileString("FONT", "HFontName", iniFileName)
    If LenB(temp) = 0 Then
         FontHor.name = "MS Sans Serif"
     Else
         FontHor.name = temp
    End If
temp = VBGetPrivateProfileString("FONT", "HFontSize", iniFileName)
    If Not IsNumeric(temp) Then
        FontHor.Size = CDbl(Replace2Regional("9.75"))
    Else
        FontHor.Size = CDbl(Replace2Regional(temp))
    End If
temp = VBGetPrivateProfileString("FONT", "HFontBold", iniFileName)
    FontHor.Bold = False
    If StrComp(temp, "true", vbTextCompare) = 0 Then FontHor.Bold = True
temp = VBGetPrivateProfileString("FONT", "HFontItalic", iniFileName)
    FontHor.Italic = False
    If StrComp(temp, "true", vbTextCompare) = 0 Then FontHor.Italic = True
temp = VBGetPrivateProfileString("FONT", "HFontColor", iniFileName)
    If Not IsNumeric(temp) Then
        HFontColor = 0
    Else
        HFontColor = Int(CDbl(temp))
    End If

'LV font
temp = VBGetPrivateProfileString("FONT", "LVFontName", iniFileName)
    If LenB(temp) = 0 Then
         FontListView.name = "MS Sans Serif"
     Else
         FontListView.name = temp
    End If
temp = VBGetPrivateProfileString("FONT", "LVFontSize", iniFileName)
    If Not IsNumeric(temp) Then
        FontListView.Size = CDbl(Replace2Regional("9.75"))
    Else
        FontListView.Size = CDbl(Replace2Regional(temp)) 'не ставить val
    End If
temp = VBGetPrivateProfileString("FONT", "LVFontBold", iniFileName)
    FontListView.Bold = False
    If StrComp(temp, "true", vbTextCompare) = 0 Then FontListView.Bold = True
temp = VBGetPrivateProfileString("FONT", "LVFontItalic", iniFileName)
    FontListView.Italic = False
    If StrComp(temp, "true", vbTextCompare) = 0 Then FontListView.Italic = True
temp = VBGetPrivateProfileString("FONT", "LVFontColor", iniFileName)
    If Not IsNumeric(temp) Then
        LVFontColor = 0
    Else
        LVFontColor = CDbl(temp)
    End If


'lv backcolor
temp = VBGetPrivateProfileString("FONT", "LVBackColor", iniFileName)
    If Not IsNumeric(temp) Then
        LVBackColor = 15000275 '12648447
    Else
        LVBackColor = Int(CDbl(temp))
    End If

'CoverVertBackColor обложка
temp = VBGetPrivateProfileString("FONT", "CoverVertBackColor", iniFileName)
    If Not IsNumeric(temp) Then
        CoverVertBackColor = vbWhite
    Else
        CoverVertBackColor = CDbl(temp)
    End If

'CoverHorBackColor
temp = VBGetPrivateProfileString("FONT", "CoverHorBackColor", iniFileName)
    If Not IsNumeric(temp) Then
        CoverHorBackColor = vbWhite
    Else
        CoverHorBackColor = temp
    End If
    
temp = VBGetPrivateProfileString("FONT", "LVHighLight", iniFileName)
    If Not IsNumeric(temp) Then
        LVHighLightLong = &HF0CAA6
    Else
        LVHighLightLong = temp
    End If
    
  
temp = VBGetPrivateProfileString("FONT", "VMcolor", iniFileName)
If Len(temp) <> 0 Then VMSameColor = CBool(temp) Else VMSameColor = False
temp = VBGetPrivateProfileString("FONT", "StripedLV", iniFileName)
If Len(temp) <> 0 Then StripedLV = CBool(temp) Else StripedLV = False
temp = VBGetPrivateProfileString("FONT", "NoLVSelFrame", iniFileName)
If Len(temp) <> 0 Then NoLVSelFrame = CBool(temp) Else NoLVSelFrame = False
temp = VBGetPrivateProfileString("LIST", "LVGrid", iniFileName)
If Len(temp) <> 0 Then Opt_ShowLVGrid = CBool(temp) Else Opt_ShowLVGrid = False
    
   
'                                                                Export
For i = 0 To LstExport_ListCount
temp = VBGetPrivateProfileString("EXPORT", "L" & i, iniFileName)
Select Case Trim$(LCase$(temp))
 Case "true", "1"
  LstExport_Arr(i) = True
 Case Else
  LstExport_Arr(i) = False
End Select
Next

For i = 0 To 2 '012
temp = VBGetPrivateProfileString("EXPORT", "OptHtml" & i & ".Caption", iniFileName)
If StrComp(temp, "true", vbTextCompare) = 0 Then Opt_HtmlJpgName = i 'OptHtml(i).Value = True
Next

CurrentHtmlTemplate = VBGetPrivateProfileString("EXPORT", "Template", iniFileName)
'If Len(CombTemplate.Text) = 0 And CombTemplate.ListCount > 0 Then CombTemplate.Text = CombTemplate.List(1)

TxtNnOnPage_Text = VBGetPrivateProfileString("EXPORT", "NumsOnPage", iniFileName)
If IsNumeric(TxtNnOnPage_Text) Then
 If Val(TxtNnOnPage_Text) < 1 Then TxtNnOnPage_Text = "30"
Else
 TxtNnOnPage_Text = "30"
End If

'разделитель полей экспорта
ExportDelim = VBGetPrivateProfileString("EXPORT", "ExportDelimiter", iniFileName)
tExpDelim.Text = putExpDelim(ExportDelim)

'сабфолдеры html
'chExpFolders
temp = VBGetPrivateProfileString("EXPORT", "UseSubFolders", iniFileName)
        If Len(temp) <> 0 Then Opt_ExpUseFolders = CBool(temp) Else Opt_ExpUseFolders = False
Opt_ExpFolder1 = VBGetPrivateProfileString("EXPORT", "SubFolder1", iniFileName)
tExpFolders(0).Text = Opt_ExpFolder1
Opt_ExpFolder2 = VBGetPrivateProfileString("EXPORT", "SubFolder2", iniFileName)
tExpFolders(1).Text = Opt_ExpFolder2
Opt_ExpFolder3 = VBGetPrivateProfileString("EXPORT", "SubFolder3", iniFileName)
tExpFolders(2).Text = Opt_ExpFolder3

ComboCDHid_Text = VBGetPrivateProfileString("CD", "CDdrive", iniFileName)
If Len(ComboCDHid_Text) = 0 Then ComboCDHid_Text = "D:\;C:\Video;C:\DVD;"

'качество сжатия
QJPG = Int(Val(VBGetPrivateProfileString("GLOBAL", "QJPG", iniFileName)))
If (QJPG < 1) Or (QJPG > 100) Then QJPG = 80




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   Опции
temp = VBGetPrivateProfileString("GLOBAL", "SaveOptOnExit", iniFileName)
If Len(temp) <> 0 Then Opt_AutoSaveOpt = CBool(temp) Else Opt_AutoSaveOpt = True
temp = VBGetPrivateProfileString("GLOBAL", "SortOnStart", iniFileName)
If Len(temp) <> 0 Then Opt_SortOnStart = CBool(temp) Else Opt_SortOnStart = False
temp = VBGetPrivateProfileString("LIST", "ColorDebt", iniFileName)
If Len(temp) <> 0 Then Opt_Debtors_Colorize = CBool(temp) Else Opt_Debtors_Colorize = True
temp = VBGetPrivateProfileString("GLOBAL", "LVLoadOnlyTitle", iniFileName)
If Len(temp) <> 0 Then Opt_LoadOnlyTitles = CBool(temp) Else Opt_LoadOnlyTitles = False
temp = VBGetPrivateProfileString("GLOBAL", "LoanAllSameLabels", iniFileName)
If Len(temp) <> 0 Then Opt_LoanAllSameLabels = CBool(temp) Else Opt_LoanAllSameLabels = True
temp = VBGetPrivateProfileString("GLOBAL", "SaveBigPix", iniFileName)
If Len(temp) <> 0 Then Opt_PicRealRes = CBool(temp) Else Opt_PicRealRes = True
temp = VBGetPrivateProfileString("GLOBAL", "UseAspect", iniFileName)
If Len(temp) <> 0 Then Opt_UseAspect = CBool(temp) Else Opt_UseAspect = True
temp = VBGetPrivateProfileString("GLOBAL", "FreeDVDFilters", iniFileName)
If Len(temp) <> 0 Then Opt_UseOurMpegFilters = CBool(temp) Else Opt_UseOurMpegFilters = False
'В глобал ини temp = VBGetPrivateProfileString("GLOBAL", "AviDirectShow", iniFileName)
'If Len(temp) <> 0 Then Opt_AviDirectShow = CBool(temp) Else Opt_AviDirectShow = False
temp = VBGetPrivateProfileString("GLOBAL", "SlideShowWindow", iniFileName)
If Len(temp) <> 0 Then Opt_NoSlideShow = CBool(temp) Else Opt_NoSlideShow = False
temp = VBGetPrivateProfileString("GLOBAL", "CenterShowPic", iniFileName)
If Len(temp) <> 0 Then Opt_CenterShowPic = CBool(temp) Else Opt_CenterShowPic = False
temp = VBGetPrivateProfileString("GLOBAL", "ListAndInfo", iniFileName)
If Len(temp) <> 0 Then Opt_UCLV_Vis = CBool(temp) Else Opt_UCLV_Vis = False
temp = VBGetPrivateProfileString("GLOBAL", "GroupWindow", iniFileName)
If Len(temp) <> 0 Then Opt_Group_Vis = CBool(temp) Else Opt_Group_Vis = False
temp = VBGetPrivateProfileString("COVER", "ShowColNames", iniFileName)
If Len(temp) <> 0 Then Opt_ShowColNames = CBool(temp) Else Opt_ShowColNames = True
'В глобал ини temp = VBGetPrivateProfileString("GLOBAL", "GetMediaType", iniFileName)
'If Len(temp) <> 0 Then Opt_GetMediaType = CBool(temp) Else Opt_GetMediaType = True

temp = VBGetPrivateProfileString("GLOBAL", "LVEDIT", iniFileName)
If Len(temp) <> 0 Then Opt_LVEDIT = CBool(temp) Else Opt_LVEDIT = False
temp = VBGetPrivateProfileString("GLOBAL", "SaveFileWithPath", iniFileName)
If Len(temp) <> 0 Then Opt_FileWithPath = CBool(temp) Else Opt_FileWithPath = False

temp = VBGetPrivateProfileString("GLOBAL", "SortLVAfterEdit", iniFileName)
If Len(temp) <> 0 Then Opt_SortLVAfterEdit = CBool(temp) Else Opt_SortLVAfterEdit = False
temp = VBGetPrivateProfileString("GLOBAL", "SortLabelAsNum", iniFileName)
If Len(temp) <> 0 Then Opt_SortLabelAsNum = CBool(temp) Else Opt_SortLabelAsNum = False
temp = VBGetPrivateProfileString("GLOBAL", "PutOtherInAnnot", iniFileName)
If Len(temp) <> 0 Then Opt_PutOtherInAnnot = CBool(temp) Else Opt_PutOtherInAnnot = False

'                                                                       сплиттеры
temp = VBGetPrivateProfileString("LIST", "LVWidth%", iniFileName)
If IsNumeric(temp) Then LVWidth = temp
temp = VBGetPrivateProfileString("LIST", "TVWidth", iniFileName)       '       TV
If IsNumeric(temp) Then TVWidth = temp
temp = VBGetPrivateProfileString("LIST", "ScrShotWidth%", iniFileName)       '       SS
If IsNumeric(temp) Then SplitLVD = temp

End If 'не global ----------------------------------------------------------------------------------------

''In                                                                                 Global
'If iFn = glob Then
''кажется мы сюда не попадаем
''                                                                           language
'temp = VBGetPrivateProfileString("Language", "LCount", iniGlobalFileName)
'If IsNumeric(temp) Then LangCount = Int(Val(temp)) Else LangCount = 2
'LastLanguage = VBGetPrivateProfileString("Language", "LastLang", iniFileName)
''If LastLanguage <> vbNullString Then
'    ReDim ComboLang(LangCount)
'    For i = 1 To LangCount
'        temp = VBGetPrivateProfileString("Language", "L" & i, iniGlobalFileName)
'        If Len(temp) <> 0 Then
'        ComboLang(i) = temp
'            If temp = LastLanguage Then
'                lngFileName = App.Path & "\" & VBGetPrivateProfileString("Language", "L" & i & "File", iniGlobalFileName)
'            End If
'        End If
'    Next i
''End If
'
''скоко                                                                          баз в списке
'temp = VBGetPrivateProfileString("GLOBAL", "BDCount", iniFileName)
'If IsNumeric(temp) Then tmpi = Int(Val(temp)) Else tmpi = 1
'Erase LstBases_List: ReDim LstBases_List(tmpi)
'For i = 1 To tmpi
'temp = VBGetPrivateProfileString("GLOBAL", "BDname" & i, iniFileName)
'If Len(temp) <> 0 Then
'    ret = FindFirstFile(temp, WFD)
'    If ret >= 0 Then
'        'ReadINI GetNameFromPathAndName(temp)
'        'дать имя базы с путем
'        If Len(GetPathFromPathAndName(temp)) = 0 Then temp = App.Path & "\" & temp
'        LstBases_List(i) = temp
''Debug.Print "+ база " & temp
'        'flag = True
'        'For j = 1 To tmpi 'UBound(LstBases_List) - 1
'        '    'нет повторам
'        '    If LstBases_List(j) = temp Then flag = False: Exit For
'        'Next j
'        'If flag Then LstBases_List(i) = temp
'    Else
'        temp = vbNullString
'    End If
'    FindClose ret
'End If
'Next i
'LstBases_ListCount = UBound(LstBases_List)
''Debug.Print "Баз всего:" & LstBases_ListCount
'
'End If 'global 2

End With
'Debug.Print "ReadIniOpt " & iniFileName
End Sub



