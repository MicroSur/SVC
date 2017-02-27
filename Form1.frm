VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "SurVideoCatalog"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   -180
   ClientWidth     =   12435
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   623
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   829
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tPlay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   8580
   End
   Begin VB.CheckBox ChBTT 
      Caption         =   "?"
      Height          =   315
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   1170
   End
   Begin SurVideoCatalog.UCVMSUR VerticalMenu 
      Height          =   8476
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1209
      _ExtentX        =   2143
      _ExtentY        =   14949
      ScaleWidth      =   190,539
      ScaleHeight     =   1512,689
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   8580
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   240
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   42
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030A
            Key             =   "CHECK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08A4
            Key             =   "LArr"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BBE
            Key             =   "AVI"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DC7
            Key             =   "DIVX"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FDC
            Key             =   "DVD"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11EC
            Key             =   "MOV"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1404
            Key             =   "MPG"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1612
            Key             =   "WMV"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1834
            Key             =   "XVID"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A44
            Key             =   "mnuPlayM"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1FDE
            Key             =   "mGotoURL"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29F0
            Key             =   "mnuEditMov"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F8A
            Key             =   "mnuAddNewMov"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3524
            Key             =   "mnuAddNewAuto"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3ABE
            Key             =   "mSR"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4058
            Key             =   "mTools"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":45F2
            Key             =   "mConvert"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":474C
            Key             =   "mValid"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":48A6
            Key             =   "mnuAutoSizeLV"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A00
            Key             =   "mnuDolgCheck"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4F9A
            Key             =   "mnuLabelCheck"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5534
            Key             =   "mGetCoverCh"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5ACE
            Key             =   "mnuExportCheckClip"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":64E0
            Key             =   "mCh2Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":663A
            Key             =   "mnuExportCheckHTML"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6BD4
            Key             =   "mDelCh"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":716E
            Key             =   "mCombine"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7708
            Key             =   "mnuDolgSel"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7CA2
            Key             =   "mnuLabelSel"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":823C
            Key             =   "mGetCoverSel"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":87D6
            Key             =   "mnuCopyLV"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":91E8
            Key             =   "mSel2Excel"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9342
            Key             =   "mnuHTML"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":98DC
            Key             =   "mDelSel"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9E76
            Key             =   "mnuCopyRow"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A410
            Key             =   "mnuLVChecked"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A9AA
            Key             =   "mnuLVSelected"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AF44
            Key             =   "mFiltActAll"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B4DE
            Key             =   "mFiltAct"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BA78
            Key             =   "SAVE_ICON"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C012
            Key             =   "mnuKillPic"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CA24
            Key             =   "mPutThisActer"
         EndProperty
      EndProperty
   End
   Begin SurVideoCatalog.MyFrame FrameView 
      Height          =   8535
      Left            =   1260
      TabIndex        =   40
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   15055
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleWidth      =   4254,553
      ScaleHeight     =   4068,98
      Begin SurVideoCatalog.XpB comHistory 
         Height          =   315
         Left            =   9300
         TabIndex        =   79
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         Caption         =   "H&istory"
         ButtonStyle     =   3
         Picture         =   "Form1.frx":CFBE
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         Height          =   705
         Left            =   3360
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   1080
         Visible         =   0   'False
         Width           =   1275
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   4455
         Left            =   2700
         TabIndex        =   72
         Top             =   600
         Width           =   3555
         _ExtentX        =   6297
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         OLEDragMode     =   1
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483629
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin SurVideoCatalog.UCLVaddon UCLV 
         Height          =   4455
         Left            =   6420
         TabIndex        =   59
         Top             =   600
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   7858
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
      Begin MSComctlLib.ListView tvGroup 
         Height          =   4455
         Left            =   60
         TabIndex        =   58
         Top             =   600
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483629
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "<>"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.Frame PicSplitLVDHid 
         BorderStyle     =   0  'None
         Height          =   3435
         Left            =   60
         TabIndex        =   41
         Top             =   5100
         Width           =   10575
         Begin VB.ListBox LstFiles 
            Appearance      =   0  'Flat
            Height          =   615
            ItemData        =   "Form1.frx":D558
            Left            =   4200
            List            =   "Form1.frx":D55A
            TabIndex        =   75
            Top             =   1260
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox TextVAnnot 
            BackColor       =   &H80000013&
            ForeColor       =   &H00000000&
            Height          =   2895
            HideSelection   =   0   'False
            Left            =   4680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   71
            Top             =   420
            Width           =   3510
         End
         Begin VB.PictureBox picScrollBoxV 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Height          =   2880
            Left            =   3420
            ScaleHeight     =   2820
            ScaleWidth      =   4215
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   420
            Visible         =   0   'False
            Width           =   4275
            Begin VB.PictureBox PicTempHid 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1395
               Index           =   1
               Left            =   3300
               ScaleHeight     =   93
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   85
               TabIndex        =   70
               Top             =   0
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.PictureBox PicFaceV 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00808080&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   2295
               Left            =   0
               MouseIcon       =   "Form1.frx":D55C
               MousePointer    =   99  'Custom
               ScaleHeight     =   40.481
               ScaleMode       =   6  'Millimeter
               ScaleWidth      =   47.096
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   0
               Width           =   2670
            End
            Begin VB.PictureBox PicTempHid 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1395
               Index           =   0
               Left            =   1740
               ScaleHeight     =   93
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   85
               TabIndex        =   68
               Top             =   0
               Visible         =   0   'False
               Width           =   1275
            End
         End
         Begin VB.Frame FrameImageHid 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   2895
            Left            =   0
            TabIndex        =   65
            Top             =   420
            Width           =   3360
            Begin VB.PictureBox Image0 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ClipControls    =   0   'False
               FillColor       =   &H00808080&
               ForeColor       =   &H00C0C0C0&
               Height          =   1755
               Left            =   0
               MouseIcon       =   "Form1.frx":DE26
               MousePointer    =   99  'Custom
               ScaleHeight     =   1755
               ScaleWidth      =   2760
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   0
               Width           =   2760
            End
         End
         Begin VB.Frame FrSplitD_Vert 
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   2355
            Left            =   3360
            TabIndex        =   64
            Top             =   720
            Width           =   60
         End
         Begin VB.Frame FrFindViewHid 
            BorderStyle     =   0  'None
            Height          =   2832
            Left            =   8340
            TabIndex        =   43
            Top             =   420
            Width           =   1935
            Begin VB.Frame FrameSearch 
               Caption         =   "Search"
               Height          =   1695
               Left            =   0
               TabIndex        =   44
               Top             =   1080
               Width           =   1935
               Begin VB.CheckBox ChMarkFindHid 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   47
                  Top             =   1380
                  Width           =   195
               End
               Begin VB.TextBox TextFind 
                  BackColor       =   &H00FFFFFF&
                  Height          =   345
                  Left            =   120
                  TabIndex        =   46
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.ComboBox CombFind 
                  Height          =   315
                  ItemData        =   "Form1.frx":E6F0
                  Left            =   120
                  List            =   "Form1.frx":E6F2
                  TabIndex        =   45
                  Top             =   240
                  Width           =   1695
               End
               Begin SurVideoCatalog.XpB ComNext 
                  Height          =   255
                  Left            =   420
                  TabIndex        =   48
                  Top             =   1380
                  Width           =   1395
                  _ExtentX        =   265
                  _ExtentY        =   265
                  Caption         =   ">"
                  ButtonStyle     =   3
                  XPColor_Pressed =   15116940
                  XPColor_Hover   =   4692449
               End
               Begin SurVideoCatalog.XpB ComFind 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   49
                  Top             =   1020
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  Caption         =   "Find"
                  ButtonStyle     =   3
                  Picture         =   "Form1.frx":E6F4
                  PictureWidth    =   16
                  PictureHeight   =   16
                  XPColor_Pressed =   15116940
                  XPColor_Hover   =   4692449
               End
            End
            Begin SurVideoCatalog.XpB ComPlay 
               Height          =   315
               Left            =   0
               TabIndex        =   50
               Top             =   720
               Width           =   1875
               _ExtentX        =   265
               _ExtentY        =   265
               Caption         =   "Play"
               ButtonStyle     =   3
               Picture         =   "Form1.frx":E84E
               PictureWidth    =   16
               PictureHeight   =   16
               XPColor_Pressed =   15116940
               XPColor_Hover   =   4692449
            End
            Begin SurVideoCatalog.XpB ComFilter 
               Height          =   315
               Left            =   0
               TabIndex        =   51
               Top             =   360
               Width           =   1875
               _ExtentX        =   265
               _ExtentY        =   265
               Caption         =   "Filter"
               ButtonStyle     =   3
               Picture         =   "Form1.frx":EDE8
               PictureWidth    =   16
               PictureHeight   =   16
               XPColor_Pressed =   15116940
               XPColor_Hover   =   4692449
            End
            Begin SurVideoCatalog.XpB ComShowFa 
               Height          =   315
               Left            =   0
               TabIndex        =   52
               Top             =   0
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               Caption         =   "Cover"
               ButtonStyle     =   3
               Picture         =   "Form1.frx":F382
               PictureWidth    =   16
               PictureHeight   =   16
               XPColor_Pressed =   15116940
               XPColor_Hover   =   4692449
            End
            Begin SurVideoCatalog.XpB ComShowAn 
               Height          =   315
               Left            =   0
               TabIndex        =   53
               Top             =   0
               Width           =   1875
               _ExtentX        =   265
               _ExtentY        =   265
               Caption         =   "Descr"
               ButtonStyle     =   3
               Picture         =   "Form1.frx":F91C
               PictureWidth    =   16
               PictureHeight   =   16
               XPColor_Pressed =   15116940
               XPColor_Hover   =   4692449
            End
         End
         Begin VB.TextBox TextItemHid 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   0
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   0
            Width           =   9255
         End
         Begin MSComctlLib.ProgressBar PBar 
            Height          =   255
            Left            =   8280
            TabIndex        =   54
            Top             =   60
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Min             =   1
         End
      End
      Begin MSComctlLib.TabStrip TabLVHid 
         Height          =   315
         Left            =   60
         TabIndex        =   56
         Top             =   240
         Width           =   7875
         _ExtentX        =   13917
         _ExtentY        =   556
         Style           =   1
         ShowTips        =   0   'False
         HotTracking     =   -1  'True
         TabMinWidth     =   2293
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
      Begin VB.Frame FrLV_Vert 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   3495
         Left            =   6300
         TabIndex        =   55
         Top             =   600
         Width           =   75
      End
      Begin VB.Frame FrTV_Vert 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   3495
         Left            =   2580
         TabIndex        =   57
         Top             =   600
         Width           =   75
      End
   End
   Begin VB.Frame FrameActer 
      Caption         =   "Actors"
      Height          =   8535
      Left            =   1260
      TabIndex        =   23
      Top             =   0
      Width           =   10665
      Begin VB.Frame FrActBio 
         BorderStyle     =   0  'None
         Height          =   4155
         Left            =   4500
         TabIndex        =   35
         Top             =   4140
         Width           =   5895
         Begin VB.TextBox TextActBio 
            BackColor       =   &H80000013&
            Height          =   3255
            Left            =   0
            MultiLine       =   -1  'True
            OLEDropMode     =   1  'Manual
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   900
            Width           =   5895
         End
         Begin VB.CommandButton ComRHid 
            Caption         =   "G"
            Height          =   315
            Index           =   0
            Left            =   5520
            TabIndex        =   12
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox TextActName 
            BackColor       =   &H80000013&
            Height          =   345
            Left            =   0
            MaxLength       =   100
            OLEDropMode     =   1  'Manual
            TabIndex        =   11
            Top             =   240
            Width           =   5475
         End
         Begin VB.Label LActBio 
            BackStyle       =   0  'Transparent
            Caption         =   "Bio"
            Height          =   195
            Left            =   0
            TabIndex        =   37
            Top             =   660
            Width           =   5355
         End
         Begin VB.Label LActName 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   4035
         End
      End
      Begin VB.Frame FrActButtons 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   8580
         TabIndex        =   34
         Top             =   240
         Width           =   1815
         Begin SurVideoCatalog.XpB comActFilt 
            Height          =   375
            Left            =   0
            TabIndex        =   73
            Top             =   3480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            ButtonStyle     =   3
            Picture         =   "Form1.frx":FEB6
            PictureWidth    =   16
            PictureHeight   =   16
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
         End
         Begin VB.Frame FrameFoto 
            Caption         =   "Photo"
            Height          =   1035
            Left            =   0
            TabIndex        =   60
            Top             =   1320
            Width           =   1815
            Begin VB.CheckBox chActFotoScale 
               Caption         =   "Scale"
               Height          =   195
               Left            =   120
               TabIndex        =   76
               Top             =   780
               Width           =   1635
            End
            Begin VB.CommandButton ComActFile 
               Height          =   435
               Left            =   120
               MousePointer    =   1  'Arrow
               Picture         =   "Form1.frx":10450
               Style           =   1  'Graphical
               TabIndex        =   63
               Top             =   240
               Width           =   435
            End
            Begin VB.CommandButton ComActPast 
               Height          =   435
               Left            =   720
               MousePointer    =   1  'Arrow
               Picture         =   "Form1.frx":114D2
               Style           =   1  'Graphical
               TabIndex        =   62
               Top             =   240
               Width           =   435
            End
            Begin VB.CommandButton ComActFotoDel 
               Height          =   435
               Left            =   1320
               MousePointer    =   1  'Arrow
               Picture         =   "Form1.frx":12554
               Style           =   1  'Graphical
               TabIndex        =   61
               Top             =   240
               Width           =   375
            End
         End
         Begin SurVideoCatalog.XpB ComCancelAct 
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   2940
            Width           =   1815
            _ExtentX        =   3228
            _ExtentY        =   688
            Caption         =   "Cancel"
            ButtonStyle     =   3
            Picture         =   "Form1.frx":12F56
            PictureWidth    =   24
            PictureHeight   =   24
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
            MaskColor       =   16711935
         End
         Begin SurVideoCatalog.XpB ComActDel 
            Height          =   375
            Left            =   0
            TabIndex        =   9
            Top             =   2460
            Width           =   1815
            _ExtentX        =   3228
            _ExtentY        =   688
            Caption         =   "Del"
            ButtonStyle     =   3
            Picture         =   "Form1.frx":13FE8
            PictureWidth    =   16
            PictureHeight   =   16
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
         End
         Begin SurVideoCatalog.XpB ComActSave 
            Height          =   375
            Left            =   0
            TabIndex        =   8
            Top             =   840
            Width           =   1815
            _ExtentX        =   265
            _ExtentY        =   265
            Caption         =   "Save"
            ButtonStyle     =   3
            Picture         =   "Form1.frx":149FA
            PictureWidth    =   16
            PictureHeight   =   16
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
            MaskColor       =   16711935
         End
         Begin SurVideoCatalog.XpB ComActEdit 
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   420
            Width           =   1815
            _ExtentX        =   265
            _ExtentY        =   265
            Caption         =   "Edit"
            ButtonStyle     =   3
            Picture         =   "Form1.frx":14D4E
            PictureWidth    =   16
            PictureHeight   =   16
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
         End
         Begin SurVideoCatalog.XpB ComAddAct 
            Height          =   375
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   1815
            _ExtentX        =   265
            _ExtentY        =   265
            Caption         =   "Add"
            ButtonStyle     =   3
            Picture         =   "Form1.frx":14EA8
            PictureWidth    =   16
            PictureHeight   =   16
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
            MaskColor       =   16711935
         End
      End
      Begin VB.Frame FrActLeft 
         BorderStyle     =   0  'None
         Height          =   8055
         Left            =   60
         TabIndex        =   30
         Top             =   240
         Width           =   4455
         Begin VB.Frame FrActSelect 
            BorderStyle     =   0  'None
            Height          =   2235
            Left            =   0
            TabIndex        =   31
            Top             =   5820
            Width           =   4335
            Begin SurVideoCatalog.XpB ComSelMovIcon 
               Height          =   375
               Left            =   0
               TabIndex        =   5
               Top             =   1800
               Width           =   4335
               _ExtentX        =   265
               _ExtentY        =   265
               Caption         =   "Select"
               ButtonStyle     =   3
               Enabled         =   0   'False
               XPColor_Pressed =   15116940
               XPColor_Hover   =   4692449
            End
            Begin VB.TextBox TextSearchLVActTypeHid 
               Height          =   315
               Left            =   0
               TabIndex        =   2
               Top             =   0
               Width           =   3795
            End
            Begin VB.ListBox ListBActHid 
               Height          =   1230
               ItemData        =   "Form1.frx":151FC
               Left            =   2520
               List            =   "Form1.frx":151FE
               MultiSelect     =   1  'Simple
               TabIndex        =   4
               Top             =   420
               Width           =   1755
            End
            Begin SurVideoCatalog.XpB comActSearchInBIO 
               Height          =   315
               Left            =   3900
               TabIndex        =   3
               Top             =   0
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   556
               Caption         =   ">"
               ButtonStyle     =   3
               XPColor_Pressed =   15116940
               XPColor_Hover   =   4692449
            End
            Begin VB.Label LActMarkCount 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Selected: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   1380
               Width           =   2175
            End
            Begin VB.Label LActMarkHelp 
               BackStyle       =   0  'Transparent
               Caption         =   "Names"
               Height          =   735
               Left            =   120
               TabIndex        =   32
               Top             =   420
               Width           =   2235
            End
         End
         Begin MSComctlLib.ListView LVActer 
            Height          =   4815
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   4395
            _ExtentX        =   7779
            _ExtentY        =   8493
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483629
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   9701
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.PictureBox PicActFotoScroll 
         BackColor       =   &H80000013&
         Height          =   3735
         Left            =   4560
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   3675
         ScaleWidth      =   3915
         TabIndex        =   24
         Top             =   240
         Width           =   3975
         Begin VB.PictureBox PicActFoto 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2175
            Left            =   120
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   2175
            ScaleWidth      =   2235
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   120
            Width           =   2235
         End
      End
   End
   Begin VB.Frame FrameCover 
      Caption         =   "Cover"
      Height          =   8535
      Left            =   1260
      TabIndex        =   26
      Top             =   0
      Width           =   10665
      Begin VB.PictureBox ImBlankHid 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1155
         Left            =   3180
         ScaleHeight     =   20.373
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   25.665
         TabIndex        =   78
         Top             =   2820
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox PicCoverTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1635
         Left            =   540
         ScaleHeight     =   1635
         ScaleWidth      =   2295
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2820
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.PictureBox PicPrintScroll 
         Height          =   4185
         Left            =   60
         ScaleHeight     =   4125
         ScaleWidth      =   10155
         TabIndex        =   28
         Top             =   240
         Width           =   10215
         Begin VB.PictureBox PicCoverPaper 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   16860
            Left            =   0
            ScaleHeight     =   297.392
            ScaleMode       =   6  'Millimeter
            ScaleWidth      =   210.609
            TabIndex        =   29
            Top             =   0
            Width           =   11940
            Begin VB.PictureBox PicCoverTextWnd 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000013&
               ForeColor       =   &H00000000&
               Height          =   2115
               Left            =   540
               ScaleHeight     =   36.777
               ScaleMode       =   6  'Millimeter
               ScaleWidth      =   81.227
               TabIndex        =   38
               Top             =   720
               Visible         =   0   'False
               Width           =   4635
            End
         End
      End
      Begin VB.Frame FrPrintBotHid 
         BorderStyle     =   0  'None
         Height          =   1665
         Left            =   180
         TabIndex        =   27
         Top             =   6696
         Width           =   10155
         Begin VB.CheckBox ChCentrTitle 
            Caption         =   "CentreT"
            Height          =   255
            Left            =   3420
            TabIndex        =   77
            Top             =   1440
            Width           =   3975
         End
         Begin SurVideoCatalog.XpB CmdPrint 
            Height          =   495
            Left            =   8040
            TabIndex        =   22
            Top             =   840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            Caption         =   "Printer"
            ButtonStyle     =   3
            Picture         =   "Form1.frx":15200
            PictureWidth    =   16
            PictureHeight   =   16
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
            MaskColor       =   16711935
         End
         Begin MSComctlLib.ImageList ImListCover 
            Left            =   7380
            Top             =   840
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483633
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":15552
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":15AEC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":16086
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":16620
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":16BBA
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.CheckBox chPrnAllOne 
            Caption         =   "All in 1"
            Height          =   255
            Left            =   3420
            TabIndex        =   21
            Top             =   1020
            Width           =   4035
         End
         Begin VB.CheckBox ChCentrP 
            Caption         =   "CentreC"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1755
         End
         Begin VB.CheckBox ChPrintChecked 
            Caption         =   "All"
            Height          =   255
            Left            =   3420
            TabIndex        =   20
            Top             =   720
            Width           =   3855
         End
         Begin VB.CheckBox ChPrintPix 
            Caption         =   "Cover"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   720
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CheckBox ChScaleP 
            Caption         =   "Scale"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   960
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CheckBox ChPropP 
            Caption         =   "Aspect"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   1200
            Width           =   3015
         End
         Begin MSComctlLib.TabStrip TabStripCover 
            Height          =   375
            Left            =   0
            TabIndex        =   15
            Top             =   180
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   688
            HotTracking     =   -1  'True
            TabMinWidth     =   2205
            ImageList       =   "ImListCover"
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   5
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Standart"
                  ImageVarType    =   2
                  ImageIndex      =   1
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Letter"
                  ImageVarType    =   2
                  ImageIndex      =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "DVD"
                  ImageVarType    =   2
                  ImageIndex      =   3
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "DVDSlim"
                  ImageVarType    =   2
                  ImageIndex      =   4
               EndProperty
               BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "List"
                  ImageVarType    =   2
                  ImageIndex      =   5
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Menu popCoverHid 
      Caption         =   "popCover"
      Visible         =   0   'False
      Begin VB.Menu mnuCoverSave 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuCoverCopy 
         Caption         =   "Copy"
      End
   End
   Begin VB.Menu popMovieHid 
      Caption         =   "popMovie"
      Visible         =   0   'False
      Begin VB.Menu mnuMovieSaveFrame 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuMovieCopyClip 
         Caption         =   "Copy"
      End
   End
   Begin VB.Menu popFaceHid 
      Caption         =   "popFace"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveFace 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuCopyFace 
         Caption         =   "Copy"
      End
   End
   Begin VB.Menu popActHid 
      Caption         =   "popAct"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveFoto 
         Caption         =   "Save as"
      End
      Begin VB.Menu mnuCopyFoto 
         Caption         =   "Copy"
      End
   End
   Begin VB.Menu popLVHid 
      Caption         =   "popLV"
      Visible         =   0   'False
      Begin VB.Menu mnuPlayM 
         Caption         =   "Play"
         Shortcut        =   ^M
      End
      Begin VB.Menu mGotoURL 
         Caption         =   "URL"
      End
      Begin VB.Menu m1Hid 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMov 
         Caption         =   "Edit"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuAddNewMov 
         Caption         =   "Add"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAddNewAuto 
         Caption         =   "Add Auto"
      End
      Begin VB.Menu mSR 
         Caption         =   "Search"
         Shortcut        =   ^F
      End
      Begin VB.Menu m20Hid 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCard 
         Caption         =   "Movie Card"
      End
      Begin VB.Menu mnuGroup 
         Caption         =   "Groups"
      End
      Begin VB.Menu m2Hid 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLVChecked 
         Caption         =   "Checked"
         Begin VB.Menu mChSel 
            Caption         =   "Check selected"
            Shortcut        =   ^Z
         End
         Begin VB.Menu mUnChSel 
            Caption         =   "UnCheck selected"
         End
         Begin VB.Menu mnuCheckAll 
            Caption         =   "All"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuCheckNone 
            Caption         =   "None"
            Shortcut        =   ^D
         End
         Begin VB.Menu mInvCh 
            Caption         =   "Invert"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuSortChecked 
            Caption         =   "Sort"
            Shortcut        =   ^S
         End
         Begin VB.Menu m3Hid 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDolgCheck 
            Caption         =   "Set Debtor"
            Shortcut        =   ^Y
         End
         Begin VB.Menu mnuLabelCheck 
            Caption         =   "Set Label"
            Shortcut        =   ^K
         End
         Begin VB.Menu mGetCoverCh 
            Caption         =   "Get Cover"
            Shortcut        =   ^R
         End
         Begin VB.Menu m4Hid 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExportCheckClip 
            Caption         =   "Clipboard"
            Shortcut        =   ^X
         End
         Begin VB.Menu mCh2Excel 
            Caption         =   "Excel"
            Shortcut        =   ^O
         End
         Begin VB.Menu mnuExportCheckHTML 
            Caption         =   "to HTML"
            Shortcut        =   ^G
         End
         Begin VB.Menu m113Hid 
            Caption         =   "-"
         End
         Begin VB.Menu mDelCh 
            Caption         =   "Delete"
            Shortcut        =   +{DEL}
         End
         Begin VB.Menu m66Hid 
            Caption         =   "-"
         End
         Begin VB.Menu mCombine 
            Caption         =   "Combine"
         End
      End
      Begin VB.Menu mnuLVSelected 
         Caption         =   "Selected"
         Begin VB.Menu mSelCh 
            Caption         =   "Select checked"
            Shortcut        =   ^V
         End
         Begin VB.Menu mUnSelCh 
            Caption         =   "UnSelect checked"
         End
         Begin VB.Menu mnuSelectAllLV 
            Caption         =   "All"
            Shortcut        =   ^A
         End
         Begin VB.Menu mInvSel 
            Caption         =   "Invert"
            Shortcut        =   ^J
         End
         Begin VB.Menu m5Hid 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDolgSel 
            Caption         =   "Set Debtor"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuLabelSel 
            Caption         =   "Set Label"
            Shortcut        =   ^L
         End
         Begin VB.Menu mGetCoverSel 
            Caption         =   "Get Cover"
            Shortcut        =   ^T
         End
         Begin VB.Menu m6Hid 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCopyLV 
            Caption         =   "to Clipboard"
            Shortcut        =   ^C
         End
         Begin VB.Menu mSel2Excel 
            Caption         =   "Excel"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuHTML 
            Caption         =   "to HTML"
            Shortcut        =   ^H
         End
         Begin VB.Menu m114Hid 
            Caption         =   "-"
         End
         Begin VB.Menu mDelSel 
            Caption         =   "Delete"
            Shortcut        =   {DEL}
         End
         Begin VB.Menu m78Hid 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCopyRow 
            Caption         =   "Copy row"
            Shortcut        =   ^W
         End
      End
      Begin VB.Menu m77Hid 
         Caption         =   "-"
      End
      Begin VB.Menu mTools 
         Caption         =   "Tools"
         Begin VB.Menu mConvert 
            Caption         =   "Convert"
         End
         Begin VB.Menu mValid 
            Caption         =   "Validation"
         End
         Begin VB.Menu mnuAutoSizeLV 
            Caption         =   "Auto size"
         End
      End
   End
   Begin VB.Menu mnuTextMenuHid 
      Caption         =   "mnuTextMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuShowThisActer 
         Caption         =   "Select"
         Enabled         =   0   'False
      End
      Begin VB.Menu mPutThisActer 
         Caption         =   "Put"
         Enabled         =   0   'False
      End
      Begin VB.Menu mFiltActAll 
         Caption         =   "ShowAll"
      End
      Begin VB.Menu m777 
         Caption         =   "-"
      End
      Begin VB.Menu mFiltAct 
         Caption         =   "Filter"
      End
      Begin VB.Menu m7Hid 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu m8Hid 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
      End
   End
   Begin VB.Menu PopShowPicHid 
      Caption         =   "PopShowPicHid"
      Visible         =   0   'False
      Begin VB.Menu mnuSavePic 
         Caption         =   "Save as"
      End
      Begin VB.Menu mnuCopyPic 
         Caption         =   "Copy"
      End
      Begin VB.Menu m9Hid 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKillPic 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu PopActInetHid 
      Caption         =   "PopActInetHid"
      Visible         =   0   'False
      Begin VB.Menu mActGoogleHid 
         Caption         =   "Google"
      End
      Begin VB.Menu mWWW1Hid 
         Caption         =   "Kinopoisk"
      End
      Begin VB.Menu mWWW2Hid 
         Caption         =   "World-Art"
      End
      Begin VB.Menu mWWW3Hid 
         Caption         =   "www3"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mPopGroup 
      Caption         =   "PopGroup"
      Visible         =   0   'False
      Begin VB.Menu mGroup 
         Caption         =   "0"
         Index           =   0
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   15
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   16
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   17
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   18
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   19
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   20
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   21
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   22
      End
      Begin VB.Menu mGroup 
         Caption         =   ""
         Index           =   23
      End
      Begin VB.Menu mGroup 
         Caption         =   "24"
         Index           =   24
      End
   End
   Begin VB.Menu mPopHistory 
      Caption         =   "History"
      Visible         =   0   'False
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   15
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   16
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   17
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   18
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   19
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   20
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   21
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   22
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   23
      End
      Begin VB.Menu mHist 
         Caption         =   ""
         Index           =   24
      End
      Begin VB.Menu mh1Hid 
         Caption         =   "-"
      End
      Begin VB.Menu mHistClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'153 - 148
'283 - 278
'ImBlankHid.BackColor = CoverHorBackColor
'mnuShowThisActer_Click
'UCLVShowPersonFlag = False
'FillLVAdd
'picUCLV.Move lblL, PicLVAddon.Top + PicLVAddon.Height, PicLVAddon.Width - FixTxtLeft
'PutCoverUCLV
'MenuActSelect
'FrmMain.UCLV.Controls("tBIO").Visible = False



'перевод твипов-пикселей
'pwidth = ScaleX(img.Width, vbPixels, vbTwips) =
'    pwidth = img.Width * Screen.TwipsPerPixelX
'MovieWidth = ScaleX(movie.Width, vbTwips, vbPixels) =
'    MovieWidth = .movie.Width / Screen.TwipsPerPixelX

'!!!!!!!!!!!!!!!!!!
'Если менять основное меню - корректировать:
' Case 91 To 116 порядковые номера меню группировки, меньше высота в MeasureItem


'                                   LAST MODIFIED
'SearchString = Replace(SearchString, "&", "%26")
'+ по альттабу при редактировании можем потерять не записанное (кликался лв)
'   'кликаем по стрелкам, начальным буквам
'   Select Case KeyCode
'   Case 18, 17 'чистые alt и шифт

'29.04.2007
'снята постоянная сортировка с tv в ide
'при заполнении tv сортировка снимается (устанавливается при сотрировке чисел)
'внедрена сортировка запроса
'strSQL = SQL1 & " Group By " & GroupField & " Order By " & GroupField '& " Desc"
'теперь если поле числовое и поле не потрошено, сортируется как число в группировке

'28.03.2007 MMI_Format As Currency 'Single  As Currency 'Single
'возможна замена, если встретится нечитаемый старым фильтром двд
'было Private Const cstrSVDName As String = "Universal Open Source MPEG Source"
'на Private Const cstrSVDName As String = "Mpeg Source" + новый соурс фильтр

'''возможные неисправности''''''
'Инет поиск: терминатор
'10:44:30> InetOpen: Ok
'10:44:30> Входной URL =
'10:44:30> Internet Connect Error: 87
' --------- не зарегистрирован msscript.ocx
''''''''''
'!!! myMsgBox применять правильный hwnd если меняем форму
'!!! frmAutoFlag Then FrmMain.SetFocus тоже

'!! с 1501 актеры worldart
'актеры кинопоиск, Acter -..- shura_le@


'''''''''''''''''''''''''''''''''''''''''''''''''''''''TODO

'+в редакторе открывать фильм без спроса еще если путь указан абсолютный сразу
'+ для интерактива создавать не svc1 а index.htm
'- потестить кансел автоплей для редактора
'+ дизейблить ползунки и ключи при входе а редактор
'+ всю форму... дизейблить кнопки при диалогах открытия в автодоб
'+ в скрипте для Пост метода (не для прямого соединения!)
'  Dim PostMethod 'True for POST method and False for GET method
'  PostMethod = True
'-хистори ,
' - сабкласс не видно, если в меню 1 буква (при учете что нет других длинных пунктов меню)
'  +очистить (пункт в меню),
'  +что делать при смене баз? таки запоминать в юсерини
'+ метаданные ави, менять найденный язык Russian на русский

'-как показать сист окно ассоциаций файлов?
'+ после нового, нет мультиселекта
'+ автод, серить галочки требующие шаблона если не в редакторе
'? автод, пропускать большие файлы
'+ автод, не добавлять дубли по названию файла

'- с одного двд Добавляли однофайловые вобы, время суммировалось и затем попался двухфайловый - время не суммировалось
'+ доработал If MenuItem = 3 Then и ctrl+N в редакторе .пусть можно Новый при редактировании
'+ редактор актеров - менять на системный цвет шрифта в редакторе
'+ html2text замена &#39; на '
'- в группировке, по алфавиту, рус/англ
'   да нет вроде\ может хоть так ?как отсортировать по оригинальному названию?
'+галка - не сортировать список автоматически после редактирования строк - после автодоба все равно не сортировать
'+галочку - сортировать метку в списке как текст/число.
'+галка в настройках = помещать примечания в конец описания в окне фильмов
'~всетаки 1:1 и перемещение стрелками влево-вправо иногда неадекватно (такси 4), и бегунок наверно он не дает по 1 му клику закрыть 1:1





'5.4.4 патч
'+ печать нескольких, не меняется обложка,
'+ печать, у второго и след. dvd меняется ориентация на портрет
'+ печать нескольких, неверная высота скриншотов

'+шаблон печати dvd slim
'-переименовывание файла по названию в базе, а в идеале по шаблонам(например английское название - русское название.avi)
' думать... риск...
'+ не сворачивать прогу при плее файла (отключил таймер)
'+ 'вставил проверку видимости верт скролла в сабкласс (так стало из-за системного (синенького )скролла )  почему скроллом крутится карточка, если нет ползунка
'+ при ctrl клике на актера в карточке попытаться выделить имя актера ограниченное ()/,
'+ запомнить фильтр персон на тек сессию
'+ гугл и 3 доп сайта.   кнопка поиск в гугле расширить с меню выбором сайтов,
' ? сайты в юсер ини
'   mActGoogleHid_Click
'+ пораньше брать картинку в карточке фильмов
'+ убирать окно фильтра при выходе из актеров
'~ что-то сделал, проверять...  если при отмене фильтра актеров вызвать помеченного (режиссера), то он не найдется
'  нужно будет перевыделять, тогда работает
'+ если не нашли персону в базе актеров - предложить внести
'+ фильтр в актерах: показать c/без био
'+правка эксп в excel
' +'запросом. баг rr_QuickExcel:Method 'CopyFromRecordset' of object 'Range' failed при больших
'   описаниях (911) и если кроме описания еще чтото есть
'+ наверно вернуть: смотрю 1:1, листаем, нет картинки, потом есть и фокус вернуть в список для листания дальше
' Form_Activate() в frmShowPic
'+ не держит окно с обложкой при перем. по записям: If ComShowFa.Visible Or Opt_NoSlideShow Then Call SendMessage(TextVAnnot.hwnd, WM_SETREDRAW, True, ByVal 0&)
'+ попытка оторвать редактор от осовного окна программы
'+ убрать использование меню LV и опасных горячих клавиш при редактировании
'- попытка в фильтре сравнивать числа если дали мат знак CheckForShablon итд (оставлен дебаг запроса)
'   - у sql (.) у val (.) у нумерик и русских (,) cdbl у sql не канает - в итоге sql режер после запятой
'+ вставить фотку актера (и био )и обложку в карточку (боль 1024 высота)
'+ масштабирование фото актера для удобства просмотра
'= мульти группы пока не делаем, слишком большая строка запроса
'  + почему шифт в группе иногда выделяет несколько в лв?  SelectLVItemFromKey LastKey в filllistview
'+нет описания в обложке у вновь (или еще когда...) созданном инишнике
'+нет видео в окне редактора на некоторых DS если открывали не во вкладке видео
'+ нет, не успевает с odd сейчас стоит слип 300 ... уже наверно можно убрать sleep 500 из захвата 3 кадра

'патч 543
'+переделать в firstrun поле файлов и другие в мемо
'+правка работы экспорта в текст неверно работал, если не выбрано название или индекс в списке был не в конце.
'+вложенные папки для файлов при экспорте в html
'+галка - запоминать абс путь к файлу и применение
'+ Редактирование в гриде, галка
'+нтмл по алфавиту
'?новая строка рейтинга кинопоиска
'   <td width=30% style="padding-left:20px"><div style="color:#f60;font:800 23px tahoma, verdana"><a href="/level/83/film/444/" class="continue">8.659</a><span style="color:#999;font:800 14px tahoma, verdana">&nbsp;(941)</div><div style="color:#999;font:100 11px tahoma, verdana">IMDB: 8.30 (127 475)</div></td>
'?А вот создавать и класть рядом с фильмом текстовую карточку-описание - мысль интересная.
'+И можно ли добавить в статистику поле Примечания?
'+ И ГОД
'+ фотики должны тоже дизейблить редактор
'+ фильтры актеров поимели окно
'+ не удалялся серийник в редакторе
'   + расковырял все про dshgetinfo
'+ renderauto вставил самплграббер : получше стабильность захвата basiccapture
'   + всунуть везьде
'   + захват делать после старт-паузы
'?букмарки в списках актеров и фильмов
'?сортировка актеров
'?Print показать.поместить; не; найденных; актеров - доп; форма
'+ сохранять бекап последнего лога svcdebug.old
'+ первый кадр - черный If Not m_cAVI.AVIStreamIsKeyFrame(5000) Then - теперь ави стартуют не с 0
'+ правка - могли пропадать цветовая диф должников при новых настройках вида списка
'+ в автодоб определять картинку на диске и вставлять
'   ? также файл(ы) субтитров ssm, sub, srt - в субтитры (после метатегов если можно)
' + по имени файла искать текстовый файл и кинуть в описание
'   настройки - искать картинку по имени файла... взять любую из папки ... искать расширения (jpg, jpeg, bmp, gif, png)
'+ переделка окна дерево Настройки/Другие
'+ not в фильтре
'    передать в CheckForShablon  а или о - номер, чтобы знать какое поле Not
'    в CheckForShablon проверять надоли нот и менять Like на not Like
'    и инвертировать ><=
'+ замена space$ на AllocString_ADV при чтении из файлов
'     и KeyValue = String$(1024, 0)'    KeyValue = AllocString_ADV(1024)


'''патч 5.4.2
'+не для рекламы - referer
'- Если делать новые поля (Моя категория, группа типа любимые, Весь Холмс итд)
'       -в настроки вкладку какие поля показывать в списке - сложно решиться
'- группировка по алфавиту названия
'? а если несколько серийников... отдавать диск должнику по серийнику
'+ правка, программа не догадывалась, что можно выделять записи одной мышкой в "лассо" (10x Andrey Konishev)
'+правка рефреша. смена языка и lvaddon?
'+не так калечить название при автозаполнениии
'+если была группировка и вызвать/убрать ее из настроек, то на каждую галку будет вопрос.
'+вызов актеров, если список пуст - вылет
'+ нет автозапуску cd в автодобавлении
'+ вызов настроек -клавиша F9

'''''''''''
'+ из актеров при выборе персоны идет теперь не пометка а фильтрация фильмов (Пометить в каталоге - Оставить...)
'+ показ всех актеров в базе актеров для текущего фильма (меню карточки)
'- поиск по шаблону не только всего поля - как делать Replace with pattern
'- фильтр нет био в актерах - не куда пока, думать над инструментами для актеров, отдельное окно
'   только русские, иностранные
'+ Нет в списке фильмов без названия, правка взад - проверка на нулл при fill lv в названии
'+ править LstBases_DragDrop перетаскивание (будет все равно не верно, если вкладки имеют одинаковые названия)
'+ Статистика показ кол-ва дисков по серийке
'+ Дополнительный Инструмент/Проверка меток и серийников в меню списка фильмов
'+ Правка для статистики по выделенному - не применялось, если снимали выделение с ctrl
'+ правка для f11 чтоб не мигал (гор прокрутка своя)
'+ kinopoisk для фильмов и актеров

'? обновить рейтинки с ибдм по аналогии с
'  Sub InetGetPics(ch As Boolean)       http://us.imdb.com/title/tt0477347/
'~ не уменьшать первую букву в скобках в актерах: Эрик (теодор) Картман - это при драгдропе
'? окошечко пока только одно в технической информации на movie site..
'? почему группа метки показ 22 и 2,22 при выборе 22
'? нафик, склеивать   .Индексированный поиск по всем базам (вкладка в ctrl+F)
'~ при опред типа dvd опред буктип, а не конкретно dvd+r например
'? как массово переопределить например все метки? к чему привязаться
'? как определить, что метка уже есть при заполнении нового фильма
'+ как определить, что метка уже дана неправильно(уже была) серийник?
'~ При объединении суммируется только 255 символов в поле имя файла - если поле базы txt (и у времени)
'? ну lcase там зачем то ... в автодобаавлении названия файлов маленькими буквами?
'- экспорт в HTML по алфавиту?
'? как-то отследить, что global.ini поврежден, и если так, сделать новый с бекапом
'   ~ до того посмотреть, что там без него плывет
'? предупредить, что ини файл ридонли, если так. INIFileFlagRW и для опций
'~ не засек ... не удалялась firstrun после добавления новых баз и удаления firstrun
'? проверить... автодобавлению очень мешает показ скриншотов в окне просмотра.  Как то его заморозить...
'? Могу подумать о пункте в меню типа "Открыть папку с фильмом в проводнике". - длительный поиск...
'- GetWord жрет двойные и более пробелы
' - через WIA
'   ~ обложки в печати - есть в двд
' ? вертик. расположение кадров в печати обложки
' ? сделать www.torrents.ru
' ~ OriginalTitleFirst = True (Шаререактор, ...
' ? произвести сортировку после глобальной замены
' ? дома по правому клику после отработки меню происходит клик на lv
' ? не вышло создавать базу программно - unicode compression как поставить?
' ? почему передается фокус в 1:1 - изза постоянной модальности окна 1:1 - поставлен флаг
' ? Почемужтак тормозит большие Био в актерах - зависит от шрифта
' = картинки записанные в базу с реальным разр. при html не масштабируются
' ? Совместная база из всех?
' ? TextGenre_KeyDown -> TextGenre_KeyPress для ловли ctrl c
' ? как бы перенести обложку в район карточки фильма
' ? антифлик в актерах? ато тормозит на больших текстах
' ? после удаления и перезапуска может не найти последнее поле - подумать о записи ключа, а не позиции
' ~ брать картинки для печати обложки - 1 раз - тока обложку
' ? не меняется ли global при инсталле? - нет
' - И при выводе в HTML тоже сделать возможность группировки по какому-либо полю.
' ? в окне редактирования добавить бы кнопки перехода от фильма к фильму
' ? редактирование в таблице
' ? Изучить кварц события и ваще
'   http://msdn.microsoft.com/library/default.asp?url=/library/en-us/directshow/htm/geteventmethodimediaeventobject.asp
' ? тестировать - почему долго выходит при опто инфо - не чистить объекты ?
' ~ врет опто инфо при nero drivespeed
' ? не работает соурс кодек дома (разрешение svcd?) - тогда проверить еще и наши кодеки при определении svcd
' + ? убрал cancelupdate удаление фильмов - баг - удаляет, после компакта поля снова есть
' ? в настройки - добавлять к должнику дату (и в редакторе)
' - Люди кино - просить сохранить при переходе меж строк и пункт автосохранения в настройках
' ? Сохранять кадры с реальным разрешением, показывать пропорционально.
'    нужно хранить пропорции
'    а если картинка не скриншот... - беда
' ? добавить форму для выбора и заполнения должника с адресами - таблицу в базе people.mdb
' ? при поиске в описании продолжить поиск не со след записи а попытаться в тойже.
''''''''''''''''''



'При изменении полей базы:
'+ При изменении названий полей в базе - менять в SQL запросах
'+ Внести индекс - Public dbMovieNameInd As Integer
'+ Поменять кол-во
'    Public Const lvHeaderIndexPole As Integer = 25
'    Public Const lvIndexPole As Integer = 24 ' 18
'    Public Const SVCBaseFielsCount = 31 '25 'полей в базе фильмов
'+ Добавить проверку в OpenDB
'+ Изменить файлы локализации ListView.CH1 и новые лейблы
'+ Добавить новое в FillLVSubs
'+ Засунуть новую пустышку в ресурсы
'+ Поменять FirstRun.mdb
'+ Если надо подогнать сортировку ListView_ColumnClick
'+ добавить в ExportHTML
'   + поменять шаблон HTML
'   + добавить в ридми  ($SUBTITLE$    - Субтитры)
'+ проверить индексы GetCoverSpisok и mnuCopyLV_Click
'  + LstExport_ListCount = 25 '19       +текстовые поля + аннотация
'  + LstExport_Arr(25) As Boolean ' с 0
'+ проверить на тестовом test.html
'+ Разместить и сконнектить поля в редакторе
'  + GetFields()
'  + PutFields()
'  + ClearFields
'+ GetPutDD копирование
'+ Автоматически заполнять новые поля
'+ добавить в инет скрипты и подпрограммы с этим связанные SetFromScript
'+ добавить в UCLV
'
''''''''''''''''''''''''''''''''''''''''''''
'
'Для новой галочки в Настройки.Другие надо
'
'ModOpt           - внести новый булин (в основном его имя будет такое же как в ини только с Opt_)
'FillTVOpt        - поместить в опред раздел со значением по умолчанию
'Var2tvOpt        - присвоить ноде значение переменной
'ComOptSave_Click - запись переменной в ини
'LstOpt2Var       - присвоить переменной чек ноды
'ReadIniOpt       - читать из ини в переменную (со значением по умолчанию)
'ReadIni          - читать из ини в переменную (со значением по умолчанию)
'MakeINI          - Перенос текущего значения в новый ини
'OpenNewDataBase  - для ToDebug если охото
'Прописать название галки в файлах локализаций
'Применить в коде
'.
'
'старый   Для новой опции надо
'+ Декларировать переменную-флаг в модуле Опций
'+ прописать в файлы локализаций (LstOpt14=Центровать 1:1 окно изображений)
'+ запись переменной в ини ComOptSave_Click
'+ чтение переменной ReadINI ReadINIOpt
'+ передача ее в опции Var2Options
'+ возврат из контрола в переменную Options2Var - LstOpt2Var - LstExport2Var
'           или возврат и применение сразу по изменению контрола
'+ процедура изменения контрола, если надо chVMcolor_Click()
'        VMSameColor = chVMcolor.Value
'        ApplyOpt
'        optsaved = False
'+ Применить в коде
'+ makeini дефолт


'узнать цвет картинки для установки цвета разрешения
'Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Private Type TGDIBitmapInfoHeader
'    Size As Long
'    Width As Long
'    Height As Long
'    Planes As Integer
'    BitCount As Integer
'    Compression As Long
'    SizeImage As Long
'    XPelsPerMeter As Long
'    YPelsPerMeter As Long
'    ClrUsed As Long
'    ClrImportant As Long
'End Type
'Implements ISampleGrabberCB
'Private mobjSampleGrabber As ISampleGrabber
'Private mudtBitmapInfo As TGDIBitmapInfoHeader

''WB
'Public WithEvents WBBRDoc As MSHTML.HTMLDocument 'правый клик

'''''

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'коллекция форм, для тултипов
Private FormObjects As New Collection

'маустраки для сплиттеров
Private WithEvents m_LV_Vert As cMouseTrack
Attribute m_LV_Vert.VB_VarHelpID = -1
Private WithEvents m_TV_Vert As cMouseTrack
Attribute m_TV_Vert.VB_VarHelpID = -1
Private WithEvents m_SS_Vert As cMouseTrack    'скриншот/обложка - аннот/обложка
Attribute m_SS_Vert.VB_VarHelpID = -1

'Dim LV_VDragFlag As Boolean
'Private LV_VDrag As Integer ' = x при перетаскивании
Private MainWidth As Long    'ширина окна без вертменю
Private MainWidthPix As Long    'ширина окна без вертменю в пикселях
Private MainHeightPix As Long
Private MainHeight As Long


Private mnuLVCheckedCaption As String    'для изменяемых названий в меню LV
Private mnuLVSelectedCaption As String
'Private mSumChCaption  As String
'Private mSumSelCaption  As String

Private fr_acter As Boolean    'ресайзить актеров



'smess for listbox
'Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam As String) As Long

'''


'Private m_wCurOptIdx As Integer 'for browse for folder
'Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'Private GoNextKey As Boolean 'true - двигались вправо по кеям ави

'
'Private not2save As Boolean 'не просить сохранять настройки
Private oldTabStripCoverInd As Integer    'текущий таб обложки


'тултипы
'Private iniTTFileName As String 'Путь и имя инишника тултипов
Private TT As New cToolTipEx

'Private CheckedInLV As Integer 'показ помеченых галкой строк при запуске


'Private Const CB_FINDSTRING = &H14C
'Private Const CB_FINDSTRINGEXACT = &H158
'Private Const CB_ERR = (-1)

'для ограничения рисайза
'Private Type POINTAPI
'    x As Long
'    y As Long
'End Type
Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Private Const WM_GETMINMAXINFO = &H24
Implements ISubclass
Private m_emr As EMsgResponse


'Private Const CB_SHOWDROPDOWN = &H14F 'комбо поиск



Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1
Private WithEvents ma_cScroll As cScrollBars
Attribute ma_cScroll.VB_VarHelpID = -1
Private WithEvents map_cScroll As cScrollBars    'print
Attribute map_cScroll.VB_VarHelpID = -1

'lview click
Private Type lvwMsgInfo
    X As Long
    Y As Long
    Flgs As Long
    Itm As Long
    SubItm As Long
End Type
Private lvwMsg As lvwMsgInfo
'''''''''''''''''''''''''''''''''''''''''''''

Private ChPrintCheckedCaption As String
Public LActMarkCountCaption As String
Private mcount As Long    ' счет помеченных фильмов

Private cIM As cIconMenu    'отключен в ide

'lved
Private mblnEditing As Boolean
Private mlngIndex As Long
Private mlngSubIndex As Long
Private mlngTBoxH As Long    'высота tb

'Private rectLabelLeft As Long

'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

'Private Sub ISampleGrabberCB_BufferCB(ByVal SampleTime As Double, ByVal BufferPointer As Long, ByVal BufferLength As Long)
'    Dim hGlobal As OLE_HANDLE
'    Dim lngPointer As Long
'    Dim lngLength As Long
'
'    On Error Resume Next
'    lngLength = Len(mudtBitmapInfo)
'    hGlobal = GlobalAlloc(, BufferLength + lngLength)
'    If (0& <> hGlobal) Then
'        lngPointer = GlobalLock(hGlobal)
'        If (0& <> lngPointer) Then
'            Call CopyMemory(ByVal lngPointer, mudtBitmapInfo, lngLength)
'            Call CopyMemory(ByVal (lngPointer + lngLength), ByVal BufferPointer, BufferLength)
'            Call GlobalUnlock(hGlobal)
'            If (DIBDataCopy(hGlobal, FrmMain.PicTempHid(0))) Then
'                Call mobjSampleGrabber.SetCallback(Nothing)
'                Call mobjSampleGrabber.SetBufferSamples(0&)
'                Exit Sub
'            End If
'        End If
'    End If
'    Call mobjSampleGrabber.SetCallback(Nothing)
'    Call mobjSampleGrabber.SetBufferSamples(0&)
'End Sub

Friend Function WindowProc(hwnd As Long, Msg As Long, wp As Long, lp As Long) As Long
Dim result As Long
Select Case Msg
Case m_RegMsg       ' QueryCancelAutoPlay
    ' TRUE: cancel AutoRun
    ' *must* be 1, not -1!
    ' FALSE: allow AutoRun
    result = 1
    ToDebug "CancelAutoPlayMain"
Case Else
    ' Pass along to default window procedure.
    result = InvokeWindowProc(hwnd, Msg, wp, lp)
End Select
' Return desired result code to Windows.
WindowProc = result
End Function



Private Sub FillLVAdd()    '(ind As Long)
'filluclv
'заполнение карточку UCLVAddon текущими данными
Dim pole As String
Dim tmp As String, tmps As String
Dim stars As Single
Dim ind As Integer
Dim mType As String

'может брать из базы поля?

On Error Resume Next
'DoEvents

'обложку
UCLVShowPersonFlag = False
FrmMain.UCLV.Controls("tBIO").Visible = False    'UCLV.Controls("tBIO") = vbNullString
If Opt_UCLVPic_Vis Or FirstActivateFlag Then
    If GetPic(PicTempHid(1), 1, "FrontFace") Then
        ' ResizeWIA PicTempHid(1), UCLV.Controls("picUCLV").ScaleWidth, UCLV.Controls("picUCLV").ScaleHeight, aratio:=True
        ResizeWIA PicTempHid(1), UCLV.Controls("picUCLV").ScaleHeight, UCLV.Controls("picUCLV").ScaleHeight, aratio:=True

        UCLV.Controls("picUCLV").Width = PicTempHid(1).Width
        UCLV.Controls("picUCLV").Picture = PicTempHid(1).Picture
    Else
        UCLV.Controls("picUCLV").Picture = Nothing
    End If
End If

'Call SendMessage(UCLV.hwnd, WM_SETREDRAW, False, ByVal 0&)
'UCLV.Visible = False 'мерцает группировка
With ListView.ListItems(ListView.SelectedItem.Index)

    'UCLV.Controls("textMName") = CheckNoNullVal(dbMovieNameInd)
    UCLV.Controls("textMName") = .Text

    '''саб поля

    UCLV.Controls("textGenre") = .SubItems(dbGenreInd)    'CheckNoNullVal(dbGenreInd)

    '    tmp = .SubItems(dbYearInd)
    '    pole = .SubItems(dbCountryInd)
    '    If Len(pole) <> 0 Then
    '        pole = pole & ", " & tmp
    '    Else
    '        pole = tmp
    '    End If
    tmp = .SubItems(dbCountryInd)
    pole = .SubItems(dbYearInd)
    If Len(pole) <> 0 And Len(tmp) <> 0 Then
        pole = pole & ", " & tmp
    ElseIf Len(tmp) <> 0 Then
        pole = tmp
        'иначе остается pole
    End If
    UCLV.Controls("textCountry") = pole

    UCLV.Controls("textAuthor") = .SubItems(dbDirectorInd)

    tmp = .SubItems(dbActerInd)
    '    tmp = Replace(tmp, " ,", ",")
    '    tmp = Replace(tmp, ", ", ",")
    '    tmp = Replace(tmp, ",", vbCrLf)
    UCLV.Controls("textRole") = tmp

    UCLV.Controls("textTime") = .SubItems(dbTimeInd)

    tmp = .SubItems(dbResolutionInd)
    pole = .SubItems(dbFpsInd)
    If Len(pole) <> 0 Then
        If Len(tmp) <> 0 Then
            tmp = tmp & ", " & pole
        Else
            tmp = pole
        End If
    End If
    pole = .SubItems(dbVideoInd)
    If Len(pole) <> 0 Then
        If Len(tmp) <> 0 Then
            UCLV.Controls("textVideo") = tmp & ", " & pole
        Else
            UCLV.Controls("textVideo") = pole
        End If
    Else
        UCLV.Controls("textVideo") = tmp
    End If

    UCLV.Controls("textAudio") = .SubItems(dbAudioInd)

    'pole = .SubItems(dbFileLenInd)

    pole = .SubItems(dbMediaTypeInd)
    'tmp = .SubItems(dbsnDiskInd)
    'UCLV.Controls("textCDN") = Trim$(.SubItems(dbCDNInd) & " " & pole & " " & tmp)
    UCLV.Controls("textCDN") = Trim$(.SubItems(dbCDNInd) & " " & pole)

    'pole = .SubItems(dbFileNameInd)
    'кто это    UCLV.Controls("textUser") = .SubItems(dbDebtorInd)
    'pole = .SubItems(dbsnDiskInd)
    UCLV.Controls("textOther") = .SubItems(dbOtherInd)

    UCLV.Controls("textLang") = .SubItems(dbLanguageInd)
    UCLV.Controls("textSubt") = .SubItems(dbSubTitleInd)

    UCLV.Controls("textFile") = .SubItems(dbFileNameInd)
    UCLV.Controls("textLabel") = .SubItems(dbLabelInd)
    UCLV.Controls("textDebtor") = .SubItems(dbDebtorInd)



    'рейтинг
    pole = .SubItems(dbRatingInd)
    pole = Replace2Regional(pole)
    If IsNumeric(pole) Then
        stars = Abs(CSng(pole))
    Else
        stars = 0
    End If
    UCLV.ShowStars stars

    'определить тип видео
    'dvd, avi, vcd, svcd
    'пока только dvd
    ind = 0
    tmp = UCLV.Controls("textFile")
    tmps = UCLV.Controls("textVideo")

    'иконки форматам в UCLV
    'dvd
    If InStr(1, tmp, "vts_", vbTextCompare) > 0 Then
        If InStr(1, tmp, ".ifo", vbTextCompare) > 0 Then
            ind = 3
            mType = "DVD"
        ElseIf InStr(1, tmp, ".vob", vbTextCompare) > 0 Then
            ind = 3
            mType = "DVD"
        End If
    End If
    'mpg
    If ind = 0 Then
        If LCase$(right$(tmp, 4)) = ".mpg" Then
            mType = "MPG"
            ind = 4
        End If
    End If
    'divx
    If ind = 0 Then
        If InStr(1, tmps, "divx", vbTextCompare) > 0 Then
            ind = 6
            mType = "DIVX"
        End If
    End If
    'divx
    If ind = 0 Then
        If InStr(1, tmps, "xvid", vbTextCompare) > 0 Then
            ind = 7
            mType = "XVID"
        End If
    End If
    'divx
    If ind = 0 Then
        If InStr(1, tmps, "wmv", vbTextCompare) > 0 Then
            mType = "WMV"
            ind = 8
        End If
    End If
    'avi, после divx и xvid и  wmv
    If ind = 0 Then
        If LCase$(right$(tmp, 4)) = ".avi" Then
            ind = 5
            mType = "AVI"
        End If
    End If
    'mov
    If ind = 0 Then
        If LCase$(right$(tmp, 4)) = ".mov" Then
            ind = 9
            mType = "MOV"
        End If
    End If
    If ind <> 0 Then
        'UCLV.Controls("imgType").Picture = ImageList.ListImages(ind).Picture
        'по ключу
        UCLV.Controls("imgType").Picture = ImageList.ListImages(mType).Picture
    Else
        Set UCLV.Controls("imgType").Picture = Nothing
    End If


End With
'Call SendMessage(UCLV.hwnd, WM_SETREDRAW, True, ByVal 0&)
'UCLV.Refresh

End Sub

Public Sub LVActClick()
'по сути кликает на помеченную в LV строку
Dim temp As String, temp2 As String
Dim i As Integer
Dim txtlen As Integer
Dim ch As String

If ars.RecordCount = 0 Then Exit Sub
If LVActer.ListItems.Count < 1 Then Exit Sub
'If LVActer.SelectedItem.Index < 1 Then Exit Sub
'If ActEditFlag Then ComCancelAct_Click: Exit Sub
ActNewFlag = False
If ars.EditMode Then ActEditFlag = True


RSGotoAct LVActer.SelectedItem.Key
CurAct = LVActer.SelectedItem.Index
LastIndAct = LVActer.ListItems(LVActer.SelectedItem.Index).SubItems(1)
CurActKey = LVActer.SelectedItem.Key

'Screen.MousePointer = vbHourglass
'картинку
Set PicActFoto.Picture = Nothing
If GetPic(PicActFoto, 2, "Face") Then  '2
    'PicActFoto.Picture = PicActFoto.Image
    'PicActFoto.Refresh
Else
    PicActFoto.Height = 0: PicActFoto.Width = 0    'убирает прокрутку
End If

PicActFotoScroll_Resize
'If ma_cScroll.Visible(efsHorizontal) Then PicActFoto.Left = -Screen.TwipsPerPixelX * ma_cScroll.Value(efsHorizontal)
'If ma_cScroll.Visible(efsVertical) Then PicActFoto.Top = -Screen.TwipsPerPixelY * ma_cScroll.Value(efsVertical)
If ma_cScroll.Visible(efsHorizontal) Then ma_cScroll.Value(efsHorizontal) = 0
If ma_cScroll.Visible(efsVertical) Then ma_cScroll.Value(efsVertical) = 0

'поля
'Call SendMessage(TextActBio.hwnd, WM_SETREDRAW, 0, 0)
TextActName.Text = vbNullString: TextActBio.Text = vbNullString
GetAFields
'Call SendMessage(TextActBio.hwnd, WM_SETREDRAW, 1, 0)

'заполнить listb
'DoEvents
ListBActHid.Clear

temp = Trim$(LVActer.ListItems(LVActer.SelectedItem.Index).Text)
Do
    temp2 = GetWord(temp)
    txtlen = Len(temp2)
    For i = 1 To txtlen
        ch = Mid$(temp2, i, 1)
        If ch < "A" Then Mid$(temp2, i, 1) = " "
        'If ch < Chr(13) Then Mid$(temp2, i, 1) = " "
        'Or ch > "~"
    Next i
    'Debug.Print Asc("A"), Asc(" "), Asc("А")
    temp2 = Trim$(temp2)
    If LenB(temp2) = 0 Then Exit Do
    ListBActHid.AddItem temp2
Loop

ActNotManualClick = False
ComSelMovIcon.Enabled = False
Screen.MousePointer = vbNormal

If ActEditFlag Then
    'Debug.Print "Редактирование Key=" & LVActer.SelectedItem.Key
    ToDebug "РедActKey=" & LVActer.SelectedItem.Key
Else
    'Debug.Print "Click L=" & LastIndAct & " C=" & CurAct & " Key=" & LVActer.SelectedItem.Key
End If

End Sub




Private Sub chActFotoScale_Click()
'If ActFlag Then Exit Sub
'картинку
Set PicActFoto.Picture = Nothing
If GetPic(PicActFoto, 2, "Face") Then  '2
Else
    PicActFoto.Height = 0: PicActFoto.Width = 0    'убирает прокрутку
End If

PicActFotoScroll_Resize
If ma_cScroll.Visible(efsHorizontal) Then ma_cScroll.Value(efsHorizontal) = 0
If ma_cScroll.Visible(efsVertical) Then ma_cScroll.Value(efsVertical) = 0
End Sub

Private Sub ChBTT_Click()
Dim c As Control
Dim rflag As Boolean
Dim i As Integer, j As Integer

VerticalMenu.SetFocus
If ChBTT.Value = vbUnchecked Then
    '                                                   Убрать
    Screen.MousePointer = vbNormal

    FormObjects.Add FrmMain, "FrmMain"
    If frmOptFlag Then FormObjects.Add FrmOptions, "FrmOptions"
    If frmFilterFlag Then FormObjects.Add FrmFilter, "FrmFilter"
    If frmSRFlag Then FormObjects.Add frmSR, "frmSR"
    FormObjects.Add frmEditor, "frmEditor"

    'If frmOptFlag Then FormObjects.Add FrmOptions, "FrmOptions"
    'If frmOptFlag Then FormObjects.Add FrmOptions, "FrmOptions"

    For j = 1 To FormObjects.Count    'для форм в коллекции

        For Each c In FormObjects(j).Controls    'Frame1.Parent
            'Debug.Print ctrl.Name
            rflag = False

            If TypeOf c Is CommandButton Then rflag = True: GoTo nextc
            If TypeOf c Is ComboBox Then rflag = True: GoTo nextc
            If TypeOf c Is CheckBox Then rflag = True: GoTo nextc
            If TypeOf c Is TextBox Then rflag = True: GoTo nextc
            'If TypeOf c Is Menu Then rflag = True: GoTo nextc
            If TypeOf c Is PictureBox Then rflag = True: GoTo nextc
            If c.name = "VerticalMenu" Then rflag = True: GoTo nextc
            If TypeOf c Is ListBox Then rflag = True: GoTo nextc
            If TypeOf c Is TreeView Then rflag = True: GoTo nextc
            If TypeOf c Is ListView Then rflag = True: GoTo nextc
            If TypeOf c Is Slider Then rflag = True: GoTo nextc
            If TypeOf c Is TabStrip Then rflag = True: GoTo nextc
            If TypeOf c Is XpB Then rflag = True: GoTo nextc
            If TypeOf c Is UCLVaddon Then rflag = True: GoTo nextc
            If TypeOf c Is OptionButton Then rflag = True: GoTo nextc

nextc:

            If rflag Then
                If c.name = "VerticalMenu" Then
                    For i = 1 To 8
                        Select Case i
                        Case 1: TT.RemoveToolTip (c.hwnd1)
                        Case 2: TT.RemoveToolTip (c.hwnd2)
                        Case 3: TT.RemoveToolTip (c.hwnd3)
                        Case 4: TT.RemoveToolTip (c.hwnd4)
                        Case 5: TT.RemoveToolTip (c.hwnd5)
                        Case 6: TT.RemoveToolTip (c.hwnd6)
                        Case 7: TT.RemoveToolTip (c.hwnd7)
                        Case 8: TT.RemoveToolTip (c.hwnd8)
                        End Select
                    Next i

                Else

                    TT.RemoveToolTip (c.hwnd)
                End If

            End If    'flag
        Next    'control
    Next j

    Set FormObjects = Nothing

    'не дестроить - убивает остальные сабклассы
    'Call TT.DestroyToolTip
    'If Not DebugMode Then
    '   AttachMessage Me, Me.hwnd, WM_GETMINMAXINFO
    'End If

Else
    '                                                      Установить
    'Debug.Print lngFileName
    Call TT.CreateToolTip(Me.hwnd, TTS_BALLOON, icoTTInfo, "Sur Video Catalog", 80, 1, -1, RGB(224, 240, 255), RGB(0, 0, 78))
    TT.DelayTime(TTDT_AUTOMATIC) = &HFFFF
    TT.DelayTime(TTDT_AUTOPOP) = &H7FFF    '16384
    TT.DelayTime(TTDT_INITIAL) = 0
    TT.DelayTime(TTDT_RESHOW) = 0

    ReadTTINI lngFileName & ".btt"

End If

End Sub
Private Sub ReadTTINI(iFn As String)
'Dim WFD As WIN32_FIND_DATA
'Dim ret As Long
Dim temp As String
Dim c As Control
Dim rflag As Boolean
Dim i As Integer, j As Integer

On Error Resume Next

'check ini
If Not FileExists(iFn) Then Exit Sub    'не делать

FormObjects.Add FrmMain, "FrmMain"
If frmOptFlag Then FormObjects.Add FrmOptions, "FrmOptions"
If frmFilterFlag Then FormObjects.Add FrmFilter, "FrmFilter"
If frmSRFlag Then FormObjects.Add frmSR, "frmSR"
FormObjects.Add frmEditor, "frmEditor"

'FormObjects.Add FrmOptions, "FrmOptions"
'If frmAutoFlag Then FormObjects.Add FrmAuto, "FrmAuto"
'If frmOptFlag Then FormObjects.Add FrmOptions, "FrmOptions"

For j = 1 To FormObjects.Count    'для форм в коллекции
    For Each c In FormObjects(j).Controls
        rflag = False
        If TypeOf c Is CommandButton Then rflag = True: GoTo nextc
        If TypeOf c Is ComboBox Then rflag = True: GoTo nextc
        If TypeOf c Is CheckBox Then rflag = True: GoTo nextc
        If TypeOf c Is TextBox Then rflag = True: GoTo nextc
        'If TypeOf c Is Menu Then rflag = True: GoTo nextc
        If TypeOf c Is PictureBox Then rflag = True: GoTo nextc
        If c.name = "VerticalMenu" Then rflag = True: GoTo nextc
        If TypeOf c Is ListBox Then rflag = True: GoTo nextc
        If TypeOf c Is TreeView Then rflag = True: GoTo nextc
        If TypeOf c Is ListView Then rflag = True: GoTo nextc
        If TypeOf c Is Slider Then rflag = True: GoTo nextc
        If TypeOf c Is TabStrip Then rflag = True: GoTo nextc
        If TypeOf c Is XpB Then rflag = True: GoTo nextc
        If TypeOf c Is UCLVaddon Then rflag = True: GoTo nextc
        If TypeOf c Is OptionButton Then rflag = True: GoTo nextc

nextc:

        If rflag Then
            If c.name = "VerticalMenu" Then
                For i = 1 To 8
                    temp = VBGetPrivateProfileString("TTGlobal", c.name & i, iFn)
                    temp = Change2lfcr(temp)
                    Select Case i
                    Case 1: Call TT.AddToolTip(c.hwnd1, temp)
                    Case 2: Call TT.AddToolTip(c.hwnd2, temp)
                    Case 3: Call TT.AddToolTip(c.hwnd3, temp)
                    Case 4: Call TT.AddToolTip(c.hwnd4, temp)
                    Case 5: Call TT.AddToolTip(c.hwnd5, temp)
                    Case 6: Call TT.AddToolTip(c.hwnd6, temp)
                    Case 7: Call TT.AddToolTip(c.hwnd7, temp)
                    Case 8: Call TT.AddToolTip(c.hwnd8, temp)
                    End Select
                Next i
            Else
                temp = VBGetPrivateProfileString("TTGlobal", c.name, iFn)
                temp = Change2lfcr(temp)
                'Debug.Print c.name
                Call TT.AddToolTip(c.hwnd, temp)
            End If
        End If
    Next    'control
Next j

Set FormObjects = Nothing
End Sub

Private Sub ChCentrP_Click()
If NoPicFrontFaceFlag Or ChPrintPix.Value = vbUnchecked Then
Else
    Select Case TabStripCover.SelectedItem.Index
    Case 1: PutPrintPixStandard
    Case 2: PutPrintPixConvert
    Case 3: PutPrintPixDVD slim:=False
    Case 4: PutPrintPixDVD slim:=True
    End Select
End If
End Sub





Private Sub ChCentrTitle_Click()
TabStripCover_Click
End Sub

Private Sub ChPrintChecked_Click()
If TabStripCover.SelectedItem.Index <> 5 Then ChPrintCheckedFlag = ChPrintChecked.Value
End Sub

Private Sub ChPrintPix_Click()
Select Case TabStripCover.SelectedItem.Index
Case 1
    If NoPicFrontFaceFlag Or ChPrintPix.Value = vbUnchecked Then
        PicCoverPaper.PaintPicture ImBlankHid.Image, 35.1, 25.1, 119.9, 119.9
        PicFaceV.Picture = PicFaceV.Image
    Else
        PutPrintPixStandard
    End If
Case 2
    If NoPicFrontFaceFlag Or ChPrintPix.Value = vbUnchecked Then
        PicCoverPaper.PaintPicture ImBlankHid.Image, 35.1, 145.2, 119.9, 119.9
        PicFaceV.Picture = PicFaceV.Image
    Else
        PutPrintPixConvert
    End If

Case 3
    If NoPicFrontFaceFlag Or ChPrintPix.Value = vbUnchecked Then
        PicCoverPaper.PaintPicture ImBlankHid.Image, 153.1, 15.3, 130, DVD_Height - 0.3    '179.7
        PicFaceV.Picture = PicFaceV.Image
    Else
        PutPrintPixDVD slim:=False
    End If
Case 4
    If NoPicFrontFaceFlag Or ChPrintPix.Value = vbUnchecked Then
        PicCoverPaper.PaintPicture ImBlankHid.Image, 148.1, 15.3, 130, DVD_Height - 0.3    '179.7
        PicFaceV.Picture = PicFaceV.Image
    Else
        PutPrintPixDVD slim:=True
    End If

End Select

End Sub

Private Sub chPrnAllOne_Click()
'If TabStripCover.SelectedItem.Index <> 5 Then ChPrintAllOneFlag = chPrnAllOne.Value
TabStripCover_Click
End Sub

Private Sub ChPropP_Click()

If NoPicFrontFaceFlag Or ChPrintPix.Value = vbUnchecked Then
Else
    Select Case TabStripCover.SelectedItem.Index
    Case 1
        PutPrintPixStandard
    Case 2
        PutPrintPixConvert
    Case 3
        PutPrintPixDVD slim:=False
    Case 4
        PutPrintPixDVD slim:=True

    End Select
End If
End Sub

Private Sub ChScaleP_Click()
If ChScaleP.Value = Unchecked Then
    ChPropP.Value = Checked
    ChPropP.Enabled = False
Else
    ChPropP.Enabled = True
End If

If NoPicFrontFaceFlag Or ChPrintPix.Value = vbUnchecked Then
Else
    Select Case TabStripCover.SelectedItem.Index
    Case 1
        PutPrintPixStandard
    Case 2
        PutPrintPixConvert
    Case 3
        PutPrintPixDVD slim:=False
    Case 4
        PutPrintPixDVD slim:=True

    End Select
End If
End Sub

Private Sub CmdPrint_Click()
Dim printDlg As PrinterDlg
Dim strsetting As String
Dim M As Integer
Dim NewPrinterName As String
Dim objPrinter As Printer

Set printDlg = New PrinterDlg
ToDebug "Принтеры"
On Error Resume Next

printDlg.PrinterName = Printer.DeviceName
printDlg.DriverName = Printer.DriverName
printDlg.Port = Printer.Port
printDlg.PaperSize = 9    '(A4)
printDlg.PaperBin = Printer.PaperBin

Select Case TabStripCover.SelectedItem.Index
Case 1, 2, 5
    printDlg.Orientation = vbPRORPortrait    '1 '"Portrait."
Case 3, 4
    printDlg.Orientation = vbPRORLandscape  '2  '"Landscape." dvd
End Select

printDlg.Flags = VBPrinterConstants.cdlPDNoSelection _
                 Or VBPrinterConstants.cdlPDNoPageNums _
                 Or VBPrinterConstants.cdlPDReturnDC
Printer.TrackDefault = False

' When CancelError is set to True the ShowPrinterDlg will return error
' 32755. You can handle the error to know when the Cancel button was
' clicked. Enable this by uncommenting the lines prefixed with "'**".
'**printDlg.CancelError = True

' Add error handling for Cancel.
'**On Error GoTo Cancel
If Not printDlg.ShowPrinter(Me.hwnd) Then
    ToDebug "Отмена печати: " & err.Description
    Exit Sub
End If

'Turn off Error Handling for Cancel.
'**On Error GoTo 0
'Dim strsetting As String

' Locate the printer that the user selected in the Printers collection.
NewPrinterName = UCase$(printDlg.PrinterName)
If Printer.DeviceName <> NewPrinterName Then
    For Each objPrinter In Printers
        If UCase$(objPrinter.DeviceName) = NewPrinterName Then
            Set Printer = objPrinter
        End If
    Next
End If

' Copy user input from the dialog box to the properties of the selected printer.
Printer.Copies = printDlg.Copies
Printer.Orientation = printDlg.Orientation
Printer.ColorMode = printDlg.ColorMode
Printer.Duplex = printDlg.Duplex
Printer.PaperBin = printDlg.PaperBin
Printer.PaperSize = printDlg.PaperSize
Printer.PrintQuality = printDlg.PrintQuality
' Display the results in the immediate (Debug) window.
' NOTE: Supported values for PaperBin and Size are printer specific. Some
' common defaults are defined in the Win32 SDK in MSDN and in Visual Basic.
' Print quality is the number of dots per inch.
With Printer
    'Debug.Print .DeviceName
    ToDebug "Принтер:" & .DeviceName
    If .Orientation = 1 Then
        strsetting = "Portrait"
    Else
        strsetting = "Landscape"
    End If
    'Debug.Print "Copies = " & .Copies, "Orientation = " &  strsetting
    ToDebug "копий:" & .Copies & ", ориент:" & strsetting
    If .ColorMode = 1 Then
        strsetting = "B/W"
    Else
        strsetting = "Color"
    End If
    'Debug.Print "ColorMode = " & strsetting
    ToDebug "ColorMode:" & strsetting
    '    If .Duplex = 1 Then
    '        strsetting = "None. "
    '    ElseIf .Duplex = 2 Then
    '        strsetting = "Horizontal/Long Edge. "
    '    ElseIf .Duplex = 3 Then
    '        strsetting = "Vertical/Short Edge. "
    '    Else
    '        strsetting = "Unknown. "
    '    End If
    'Debug.Print "Duplex = " & strsetting
    'Debug.Print "PaperBin = " & .PaperBin
    'Debug.Print "PaperSize = " & .PaperSize
    ToDebug "PaperSize=" & .PaperSize
    'Debug.Print "PrintQuality = " & .PrintQuality
    ToDebug "PrintQuality = " & .PrintQuality
    If (printDlg.Flags And VBPrinterConstants.cdlPDPrintToFile) = _
       VBPrinterConstants.cdlPDPrintToFile Then
        'Debug.Print "Print to File Selected"
        ToDebug "Print to File"
    Else
        ToDebug "Not Print to File"
    End If
    'Debug.Print "hDC = " & printDlg.hdc
End With

MousePointer = vbHourglass

If ChPrintChecked Then    'печать помеченных обложек
    ToDebug "печать помеченных: " & CheckCount
    CmdPrint.Enabled = False
    Screen.MousePointer = vbHourglass

    For M = 1 To UBound(CheckRows)
        'rs.MoveFirst: rs.Move CheckRows(M)

        'RSMOVE CheckRows(M), "CmdPrint_Click"
        RSGoto CheckRowsKey(M)

        If TabStripCover.SelectedItem.Index = 1 Then Call ShowCoverStandard
        If TabStripCover.SelectedItem.Index = 2 Then Call ShowCoverConvert
        If TabStripCover.SelectedItem.Index = 3 Then Call ShowCoverDVD(False)
        If TabStripCover.SelectedItem.Index = 4 Then Call ShowCoverDVD(True)

        PicCoverPaper.Picture = PicCoverPaper.Image
        PicCoverPaper.Refresh

        DoEvents
        Printer.PaintPicture PicCoverPaper.Picture, 0, 0

        '  Printer.EndDoc
        Printer.NewPage

        'еще раз, а то dvd менял ориентацию (из-за Printer.EndDoc?)
        ' Copy user input from the dialog box to the properties of the selected printer.
        'Printer.Copies = printDlg.Copies
        'Printer.Orientation = printDlg.Orientation
        'Printer.ColorMode = printDlg.ColorMode
        'Printer.Duplex = printDlg.Duplex
        'Printer.PaperBin = printDlg.PaperBin
        'Printer.PaperSize = printDlg.PaperSize
        'Printer.PrintQuality = printDlg.PrintQuality

    Next M

    Printer.EndDoc

    CmdPrint.Enabled = True

    'восстановить текущ. поз. в базе
    RestoreBasePos

    Screen.MousePointer = vbNormal

Else    'печать этой обложки

    DoEvents
    ToDebug "печать текущей"
    PicCoverPaper.Picture = PicCoverPaper.Image
    Printer.PaintPicture PicCoverPaper.Picture, 0, 0    'Printer.CurrentX, Printer.CurrentY ', 0, 0
    'Debug.Print Printer.CurrentX, Printer.CurrentY
    Printer.EndDoc
End If

MousePointer = vbDefault

'Exit Sub
'**Cancel:
'**If Err.Number = 32755 Then
'**Debug.Print "Cancel Selected"
'**Else
'**Debug.Print "A nonCancel Error Occured - "; Err.Number
'**End If
End Sub


Private Sub ComActDel_Click()
Dim okk As Integer

If BaseAReadOnly Then myMsgBox msgsvc(25), vbInformation, , Me.hwnd: Exit Sub
If BaseAReadOnlyU Then myMsgBox msgsvc(23), vbInformation, , Me.hwnd: Exit Sub
If ars.RecordCount = 0 Then Exit Sub
If LVActer.SelectedItem Is Nothing Then Exit Sub

okk = myMsgBox(LVActer.SelectedItem.Text & vbCrLf & vbCrLf & msgsvc(15), vbOKCancel, , Me.hwnd)
If okk = vbCancel Then Exit Sub

ActFlag = False

'LockWindowUpdate TextActBio.hWnd

ars.Delete

'удалить из списка
LVActer.ListItems.Remove LVActer.SelectedItem.Index

If LVActer.ListItems.Count = 0 Then
    'почистить окна
    Set PicActFoto = Nothing
    TextActBio = vbNullString: TextActName = vbNullString
    ListBActHid.Clear
    ComActDel.Enabled = False
    ComActEdit.Enabled = False

Else
    'позиционировать список
    'наоборот в клике GotoLVAct ars("Key")
    LVActClick
End If

FrameActer.Caption = FrameActerCaption & LVActer.ListItems.Count & ")"

'LockWindowUpdate 0
End Sub

Private Sub ComActEdit_Click()
'Dim i As Integer

If BaseAReadOnly Then myMsgBox msgsvc(25), vbInformation, , Me.hwnd: Exit Sub
If BaseAReadOnlyU Then myMsgBox msgsvc(23), vbInformation, , Me.hwnd: Exit Sub
If ars.RecordCount = 0 Then Exit Sub
If LVActer.ListItems.Count = 0 Then Exit Sub

Call ActEditorColors(True)


ActEditFlag = True
ComActSave.Enabled = True
ComActPast.Enabled = True
ComActFile.Enabled = True
ComActFotoDel.Enabled = True
ComCancelAct.Enabled = True
ComActDel.Enabled = False
ComAddAct.Enabled = False
ComActEdit.Enabled = False
TextSearchLVActTypeHid.Enabled = False
comActFilt.Enabled = False
chActFotoScale.Enabled = False

TextActName.SetFocus
ActFlag = True
LastIndAct = LVActer.ListItems(LVActer.SelectedItem.Index).SubItems(1)
CurActKey = LVActer.SelectedItem.Key

RSGotoAct CurActKey

'ars.MoveFirst
'ars.Move LVActer.ListItems(LVActer.SelectedItem.Index).SubItems(1) - 1



ars.Edit
'Debug.Print "Редактирование актера Key=" & CurActKey
ToDebug "CActEdKey=" & CurActKey
End Sub

Private Sub ComActFile_Click()
Dim iFile As String
Dim TifPngFlag As Boolean

On Error GoTo err    'инвалид пикча

iFile = pLoadPixDialog
If LCase$(getExtFromFile(iFile)) = "png" Then TifPngFlag = True
If left$(LCase$(getExtFromFile(iFile)), 3) = "tif" Then TifPngFlag = True

If iFile <> vbNullString Then

    If TifPngFlag Then
        PicActFoto.Picture = LoadPictureWIA(iFile)
    Else
        PicActFoto.Picture = LoadPicture(iFile)
    End If

    SavePicActFlag = True
    'ComActSave.BackColor = &HC0C0E0
    PicActFotoScroll_Resize
End If

Exit Sub
err:
ToDebug "Err_ComAF: " & err.Description
MsgBox err.Description, vbCritical
End Sub

Private Sub comActFilt_Click()
On Error Resume Next
frmActFilt.Show 0, FrmMain
End Sub


Private Sub ComActFotoDel_Click()

If myMsgBox(msgsvc(4), vbOKCancel, , Me.hwnd) = vbCancel Then Exit Sub

If ars.EditMode Then
Else
    ars.Edit
End If

'ComActSave.BackColor = &HC0C0E0
Set PicActFoto = Nothing
ars.Fields("Face") = vbNullString
PicActFoto.Width = 0: PicActFoto.Height = 0: PicActFotoScroll_Resize
End Sub
Private Sub ComActPast_Click()
'Dim f As Integer
'Clipboard.GetFormat (f)
'Debug.Print f
If Clipboard.GetFormat(vbCFDIB) Then
    PicActFoto.Picture = Clipboard.GetData
    SavePicActFlag = True
    'ComActSave.BackColor = &HC0C0E0
    PicActFotoScroll_Resize
End If
End Sub


Private Sub ComActSave_Click()
Dim akey As String    'ключ до апдейта (новый)
'Dim i As Integer

If BaseAReadOnly Then myMsgBox msgsvc(25), vbInformation, , Me.hwnd: Exit Sub
'If BaseAReadOnlyU Then myMsgBox msgsvc(23), vbInformation, , Me.hwnd: Exit Sub
If Len(TextActName.Text) = 0 Then myMsgBox msgsvc(39), vbInformation, , Me.hwnd: Exit Sub

Call ActEditorColors(False)


ComActSave.Enabled = False
ComActPast.Enabled = False
ComActFile.Enabled = False
ComActFotoDel.Enabled = False
ComActEdit.Enabled = True
ComActDel.Enabled = True
ComAddAct.Enabled = True
ComCancelAct.Enabled = False
TextSearchLVActTypeHid.Enabled = True
PicActFotoScroll.Visible = False
comActFilt.Enabled = True
chActFotoScale.Enabled = True

If ars.EditMode = 0 Then ars.Edit

'картинку в базу
If SavePicActFlag Then Pic2JPG PicActFoto, 2, "Face"
SavePicActFlag = False

'поля в базу
PutAFields

akey = ars("Key") & """"    '  1
ars.Update    '                2

'Debug.Print LastIndAct

If ActNewFlag Then
    LVActer.Visible = False: LVActer.Sorted = False

    'добавить в ALV
    LastIndAct = LVActer.ListItems.Count + 1
    LVActer.ListItems.Add(, akey, TextActName.Text).ListSubItems.Add 1, , LastIndAct
    'LVActer.ListItems(LastIndAct).SubItems(1) = LastIndAct

    GotoLVAct akey    'встать на... в списке
    'Set LVActer.SelectedItem = LVActer.ListItems(LastIndAct)

    RSGotoAct akey    'встать на созданную запись

    LVActer.Sorted = True: LVActer.Visible = True
    CurAct = LVActer.SelectedItem.Index
    FrameActer.Caption = FrameActerCaption & LastIndAct & ")"

    'ars.MoveFirst: ars.MoveLast
    ActNewFlag = False

    'Debug.Print "Save New L=" & LastIndAct & " C=" & CurAct & " Key=" & akey
    'Debug.Print "Save New Field = " & ars.Fields("Name")

ElseIf ActEditFlag Then

    LVActer.ListItems(CurAct).Text = TextActName.Text
    ActEditFlag = False
    'LVActer.Sorted = True 'или потом

End If

ActFlag = False

'LVActer.SelectedItem.EnsureVisible
LV_EnsureVisible LVActer, LVActer.SelectedItem.Index

PicActFotoScroll_Resize
PicActFotoScroll.Visible = True

ComActSave.BackColor = &HFFFFFF
FrameActer.Caption = FrameActerCaption & LVActer.ListItems.Count & ")"

'Debug.Print "Актер сохранен. Key=" & akey
ToDebug "ActSavedKey=" & akey
End Sub

Private Sub comActSearchInBIO_Click()
Dim itmX As ListItem
'in annot
'mzt Dim nxt As Long
Dim temp As String
Dim ret As Long
Dim i As Long

If LVActer.ListItems.Count = 0 Then Exit Sub

Screen.MousePointer = vbHourglass

For i = LVActer.SelectedItem.Index + 1 To LVActer.ListItems.Count

    Set itmX = LVActer.ListItems(i)

    'ars.MoveFirst: ars.Move itmX.SubItems(1) - 1
    RSGotoAct LVActer.ListItems(i).Key


    'Debug.Print ars("Name")
    If ars.Fields("Bio").Value <> vbNullString Then
        temp = ars.Fields("Bio")
    Else
        temp = vbNullString
    End If

    ret = InStr(1, temp, TextSearchLVActTypeHid.Text, vbTextCompare)
    If ret > 0 Then

        Set LVActer.SelectedItem = itmX            'LVActer.ListItems.Item(nxt)
        LVActClick    'itmX
        LVActer.SetFocus
        'LVActer.ListItems(LVActer.SelectedItem.Index).EnsureVisible
        LV_EnsureVisible LVActer, LVActer.SelectedItem.Index

        TextActBio.SetFocus
        TextActBio.SelStart = ret - 1
        TextActBio.SelLength = Len(TextSearchLVActTypeHid.Text)

        Exit For
    End If
Next

Set itmX = Nothing
Screen.MousePointer = vbNormal
End Sub
Private Sub ActEditorColors(edflag As Boolean)
If edflag Then
    ComActSave.BackColor = &HC0C0E0

    TextActName.BackColor = &HFFFFFF
    TextActName.ForeColor = &H0

    TextActBio.BackColor = &HFFFFFF
    TextActBio.ForeColor = &H0

    PicActFotoScroll.BackColor = &HFFFFFF
Else

    TextActName.BackColor = LVBackColor
    TextActName.ForeColor = LVFontColor

    TextActBio.BackColor = LVBackColor
    TextActBio.ForeColor = LVFontColor

    PicActFotoScroll.BackColor = LVBackColor

    'ComActSave там остался в конце
End If

End Sub
Private Sub ComAddAct_Click()
If BaseAReadOnly Then myMsgBox msgsvc(25), vbInformation, , Me.hwnd: Exit Sub

Dim i As Integer

Call ActEditorColors(True)

Set PicActFoto.Picture = Nothing

TextActName.Text = vbNullString
TextActBio.Text = vbNullString
ListBActHid.Clear
PicActFoto.Width = 0: PicActFoto.Height = 0: PicActFotoScroll_Resize

If rs.EditMode Then rs.CancelUpdate

TextSearchLVActTypeHid.Enabled = False

ComActSave.Enabled = True
ComAddAct.Enabled = False
ComActEdit.Enabled = False
ComActDel.Enabled = False
ActNewFlag = True
ComActPast.Enabled = True
ComActFile.Enabled = True
ComActFotoDel.Enabled = True
ComCancelAct.Enabled = True
comActFilt.Enabled = False
chActFotoScale.Enabled = False
ActFlag = True

TextActName.SetFocus

ars.AddNew
ToDebug "AddActNewKey=" & ars("Key")    '& CurActKey
End Sub

Private Sub CombFind_Click()
ComNext.Enabled = True
'показать описание
If CombFind.ItemData(CombFind.ListIndex) = 5 Then ComShowAn_Click
If FrameView.Visible Then
    If CombFind.ItemData(CombFind.ListIndex) = 6 Then TextFind.Text = Hex$(GetSerialNumber(ComboCDHid_Text))
End If
End Sub

Private Sub CombFind_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub ComFilter_ShiftClick(Shift As Integer)
'в хелпе пишу контрол 2 нажимать с кликом
'Отменить фильтр
'аналог frmfilter.CommShowAll

Dim strSQL As String
On Error Resume Next    ' надо если после фильтров нет помеченных полей

If FiltPersonFlag Or FiltValidationFlag Or FilteredFlag Then
    'заходим
Else
    Exit Sub
End If

LastInd = FrmMain.ListView.SelectedItem.SubItems(lvIndexPole)

If GroupedFlag And Len(LastSQLGroupString) <> 0 Then
    'применить запрос от группировки
    strSQL = "Select * From Storage Where " & LastSQLGroupString
Else
    'вернуть все
    strSQL = "Select * From Storage"
End If

Set rs = DB.OpenRecordset(strSQL)

FiltPersonFlag = False    'снять флаг фильтрации по актеру
FiltValidationFlag = False
FilteredFlag = False    'флаг неполного показа (фильтрация) 1

FillListView

FrmMain.ComFilter.BackColor = &HFFFFFF
End Sub




Private Sub comHistory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'arrHistory()
'показ меню истории
Dim i As Integer
Dim doflag As Boolean

If Button = vbLeftButton Then
'mHistClear.Enabled = False
    For i = 0 To nHistory
        If Len(arrHistoryKeys(i)) = 0 Then
            'пустые не покажем сабклассом
            'mHist(i).Visible = False
            mHist(i).Caption = vbNullString
            mHist(i).Tag = vbNullString
        Else
        'заполнение меню из массива истории
            mHist(i).Caption = arrHistoryTitles(i)
            mHist(i).Tag = arrHistoryKeys(i)
            doflag = True
            'mHistClear.Enabled = True
        End If
    Next i

    If doflag Then PopupMenu mPopHistory, vbPopupMenuRightAlign, (comHistory.left + 2 * comHistory.Width) / Screen.TwipsPerPixelX, (comHistory.top + comHistory.Height) / Screen.TwipsPerPixelY
End If

End Sub

Private Sub ComRHid_Click(Index As Integer)
Dim tmp As String
Dim temp As Long
Dim strPath As String
Dim site As String

On Error GoTo err

strPath = Space$(255)

'Select Case Index
'Case 0 'картинку для актера
tmp = Replace(TextActName.Text, "(", vbNullString)
tmp = Replace(tmp, ")", vbNullString)
tmp = Replace(tmp, "/", vbNullString)

'site = "http://images.yandex.ru/yandsearch?stype=image&text=" & tmp
'site = "http://images.google.com/images?q=" & tmp
site = ActWWWsite & tmp

temp = FindExecutable(site, "", strPath)
Select Case temp
Case 31
    myMsgBox msgsvc(17), vbInformation, , Me.hwnd
    Exit Sub
Case 2
End Select

temp = ShellExecute(GetDesktopWindow(), "open", site, vbNull, vbNull, 1)
ToDebug "crh_ret=" & temp

Exit Sub
err:
ToDebug "Err_RHid: " & err.Description
End Sub

Private Sub ComRHid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim p As POINTAPI
'GetCursorPos p
If Button = vbRightButton Then PopupMenu PopActInetHid    ', , (X - ComRHid(0).Left) / Screen.TwipsPerPixelX ', ComRHid(0).Top + ComRHid(0).Height
'If Button = vbRightButton Then
'PopupMenu PopActInetHid, , ComRHid(0).left / 8 ', Y ', ComRHid(0).Top + ComRHid(0).Height
'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Debug.Print KeyAscii

On Error Resume Next

Select Case KeyAscii

Case 14    '^N
    'Select Case True
    '    Case FrameView.Visible
    Call mnuAddNewMov_Click
    '    Case FrameActer.Visible
    '        ComAddAct_Click
    '    Case FrameAddEdit.Visible
    '        ComOpen_Click
    'End Select

Case 15    '^O
    Select Case True
        '    Case FrameView.Visible
        '        Call mnuAddNewMov_Click
    Case FrameActer.Visible
        ComAddAct_Click
    Case frmEditorFlag    'FrameAddEdit.Visible
        ComOpen_Click
    End Select

Case 19    '^S
    Select Case True
        '    Case FrameView.Visible
        '        Call mnuAddNewMov_Click
    Case FrameActer.Visible
        ComActSave_Click
    Case frmEditorFlag    'FrameAddEdit.Visible
        SaveFromEditor
    End Select

Case 5    '^E
    If addflag Or editFlag Then Exit Sub    'не делать в редакторе
    If FrameView.Visible Then
        Call mnuEditMov_Click
    End If

Case 6    'crtl+F
    If addflag Or editFlag Then Exit Sub    'не делать в редакторе
    If FrameView.Visible Then
        Call mSR_Click    'поиск
    End If

Case 17    'crtl+Q
    If addflag Or editFlag Then Exit Sub    'не делать в редакторе
    If FrameView.Visible Then
        Call ComFilter_Click    'Фильтр
    End If

End Select
End Sub



Private Sub FrSplitD_Vert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Tracking is initialised by entering the control:
FrSplitD_Vert.MousePointer = 9
If Not (m_SS_Vert.Tracking) Then m_SS_Vert.StartMouseTracking
End Sub

Private Sub FrSplitD_Vert_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'двигаем вертик сплиттер
If Opt_NoSlideShow Then
    SplitLVD = ((X + picScrollBoxV.Width) * 100) / SSCoverAnnotW
Else
    SplitLVD = ((X + FrameImageHid.Width) * 100) / SSCoverAnnotW
End If
Form_Resize
End Sub



Private Sub ListBActHid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'выборка по именам
If Button = vbRightButton Then ActOtherFilters 4
End Sub

Private Sub ListView_Click()
LVManualClickFlag = True
LVCLICK
'DoEvents
'загрузить в редактор
If frmEditorFlag Then
    Screen.MousePointer = vbHourglass
    addflag = False: editFlag = True
    GetEditPix
    frmEditor.ImgPrCov.Picture = frmEditor.PicFrontFace.Picture
    Mark2SaveFlag = False
    GetFields
    Mark2SaveFlag = True
    frmEditor.ComSaveRec.BackColor = &HC0E0C0
    frmEditor.ComDel.Visible = True: frmEditor.ComDel.Enabled = True
    ToDebug "LV_Ed_Key: " & rs("key")
    Screen.MousePointer = vbNormal
End If
End Sub




Private Sub ListView_KeyDown(KeyCode As Integer, Shift As Integer)
Timer2.Enabled = False

'Debug.Print KeyCode, Shift

Select Case KeyCode

Case 46    'Del
    If Shift = 0 Then    'удал sel
        Get_LV_Selections
        mDelSel_Click
    Else    'удал ch
        mDelCh_Click
    End If

End Select

End Sub

Private Sub ListView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'поймать фокус, если можно
On Error Resume Next
'If Not FormShowPicLoaded Then
If txtEdit.Visible Then Exit Sub
If LstFiles.Visible Then Exit Sub
If frmSRFlag Then
    If frmSR.Visible Then Exit Sub
End If
If TextItemHid.SelLength > 0 Then Exit Sub

If GetForegroundWindow = Me.hwnd Then
    If ActiveControl.name <> "ListView" Then ListView.SetFocus
End If
End Sub

Private Sub ListView_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
'Debug.Print Time; " sd"
'Dim Itm As ListItem
'Dim tmp As Long
On Error Resume Next
SelCount = SendMessage(ListView.hwnd, LVM_GETSELECTEDCOUNT, 0&, ByVal 0&)

If SelCount = 1 Then
    'тащимое = текущее
    If ListView.SelectedItem.Index <> CurSearch Then
        LVCLICK
    End If
End If
'
'Else
''для мультиков перечитать массивы пометок
''не надо, при dd запрос каждого опять по ключу
'

End Sub

Private Sub LstFiles_DblClick()
If LstFiles.ListIndex > -1 Then
    LstFiles.Visible = False
    PlayMovie LstFiles.Text
End If
End Sub

Private Sub LstFiles_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then LstFiles.Visible = False: Exit Sub

If LstFiles.ListIndex > -1 Then
    LstFiles.Visible = False
    PlayMovie LstFiles.Text
End If

End Sub

Private Sub LstFiles_LostFocus()
LstFiles.Visible = False
End Sub

Private Sub LVActer_Click()
LVActClick
End Sub

Private Sub LVActer_DblClick()
If ars.RecordCount = 0 Then Exit Sub
If LVActer.ListItems.Count < 1 Then Exit Sub

TextSearchLVActTypeHid = LVActer.ListItems(LVActer.SelectedItem.Index)

End Sub

Private Sub LVActer_KeyUp(KeyCode As Integer, Shift As Integer)
LVActClick
End Sub

Private Sub LVActer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If ActNewFlag Or ActEditFlag Then Exit Sub
If TextActBio.SelLength <> 0 Then Exit Sub
If TextActName.SelLength <> 0 Then Exit Sub

If GetForegroundWindow = Me.hwnd Then
    If ActiveControl.name = "TextSearchLVActTypeHid" Then Exit Sub
    If ActiveControl.name <> "LVActer" Then LVActer.SetFocus
End If
End Sub


Private Sub mActGoogleHid_Click()
ActWWWsite = "http://images.google.com/images?q="
ComRHid(0).Caption = "G"
End Sub

Private Sub mCh2Excel_Click()
If CheckCount = 0 Then Exit Sub
Export2Excel True
End Sub

Private Sub mChSel_Click()
'пометить выделенное
If SelCount < 1 Then Exit Sub

Dim Itm As ListItem

For Each Itm In ListView.ListItems
    If Itm.Selected Then
        Itm.Checked = True
        ListView_ItemCheck Itm
    End If
Next
LVCLICK
End Sub

Private Sub mCombine_Click()
If CheckCount < 2 Then Exit Sub
JoinMovies True    'for checked
End Sub

Private Sub mConvert_Click()
Dim tmp As String
On Error Resume Next
tmp = App.Path & "\convert2svc.exe"
If FileExists(tmp) Then
    Shell tmp, vbNormalFocus
End If
End Sub

Private Sub mDelCh_Click()
If CheckCount = 0 Then Exit Sub
DelMovies True    'for checked
End Sub
Private Sub mDelSel_Click()
If SelCount < 1 Then Exit Sub
DelMovies False    'for selected
End Sub

Private Sub mFiltAct_Click()
'оставить в лв только фильмы с этим актером
'искать в актерах и режиссерах
'фильтрация  будет отменена, группировку оставим

If Len(sPerson) = 0 Then Exit Sub

Dim strSQL As String

Screen.MousePointer = vbHourglass
LastSQLPersonString = "Director Like '*" & sPerson & "*' Or Acter Like '*" & sPerson & "*'"
strSQL = "SELECT * FROM Storage WHERE (" & LastSQLPersonString & ")"

If GroupedFlag Then
    'добавить запрос от группировки
    strSQL = strSQL & " AND (" & LastSQLGroupString & ")"
    'Else
    '    strSQL = strSQL & ")"
End If

'Debug.Print "mFA:" & strSQL

On Error GoTo err
Set rs = DB.OpenRecordset(strSQL)

FillListView
'свой флаг
FiltPersonFlag = True
FrmMain.ComFilter.BackColor = &HC0C0FF

Screen.MousePointer = vbNormal
Exit Sub

err:
Screen.MousePointer = vbNormal
ToDebug "Err_mFiltA:" & err.Description
End Sub

Private Sub mFiltActAll_Click()
'взять текущих актеров и режиссеров, распотрошить и скормить на фильтрацию в базе актеров
'те показ только текущих персон в актерах

Dim j As Long
Dim R() As String
Dim ArrFlag As Boolean
Dim rsArr() As String
Dim Pers As String
Dim sSQL As String
Dim tmp As String

Screen.MousePointer = vbHourglass
DoEvents

Pers = CheckNoNull("Director") & "," & CheckNoNull("Acter")
Pers = Replace(Pers, "(", ",")
Pers = Replace(Pers, ")", ",")
Pers = Replace(Pers, "/", ",")
'Pers = Replace(Pers, "-", ",") 'жан-луи превратится в жан и найдем всех жанов
'Pers = Replace(Pers, ".", ",")
Pers = Replace(Pers, "jr", ",")
Pers = Replace(Pers, "младший", ",")
Pers = Replace(Pers, "старший", ",")
SQLCompatible Pers

ReDim rsArr(0)

If Len(Pers) < 5 Then      'не брать коротких
    Screen.MousePointer = vbNormal
Else
    If Tokenize04(Pers, R(), ",", False) > -1 Then                  ' False Пустышек не должно.
        TriQuickSortString R    'sorts your string array
        remdups R    'removes all duplicates

        For j = 0 To UBound(R)
            If ArrFlag Then    'не первая
                ReDim Preserve rsArr(UBound(rsArr) + 1)
                'join ставит пробел
                If Len(R(j)) > 6 Then
                    tmp = Trim$(R(j))
                    tmp = Replace(tmp, " ", "*")    'чтобы найти с отчеством всередине *Александр*Лебедев*
                    rsArr(UBound(rsArr)) = "Or Name Like '*" & tmp & "*'"
                End If
            Else    'первая
                ArrFlag = True
                'можем потерять, потом все равно проверять на первый Or
                If Len(R(j)) > 6 Then
                    tmp = Trim$(R(j))
                    tmp = Replace(tmp, " ", "*")
                    rsArr(UBound(rsArr)) = " Name Like '*" & tmp & "*'"
                End If
            End If

        Next j
    End If

    '      TriQuickSortString rsArr    'sorts your string array
    '      remdups rsArr    'removes all duplicates

    If (UBound(rsArr) = 0) And (Not ArrFlag) Then
        'пусто
        Screen.MousePointer = vbNormal
    Else
        'UBound(rsArr) + 1 штук, фильтровать
        Pers = Join(rsArr)
        If left$(Pers, 3) = " Or" Then Pers = right$(Pers, Len(Pers) - 3)

        ' не sSQL = "Select * From Acter Where Name In (" & Pers & ")"

        sSQL = "Select * From Acter Where" & Pers

        'Debug.Print sSQL

        On Error GoTo err
        Set ars = ADB.OpenRecordset(sSQL)

        ActOtherFilters 5
        VerticalMenu_MenuItemClick 5, 0
        Screen.MousePointer = vbNormal

    End If
End If
Exit Sub

err:
'попадаем сюда и по ошибке в VerticalMenu_MenuItemClick 5, 0
VerticalMenu_MenuItemClick 1, 0
Screen.MousePointer = vbNormal
'MsgBox err.Description
ToDebug "Err_FActAll"
MsgBox msgsvc(46), vbExclamation    ': ToDebug err.Description

End Sub

Private Sub mGetCoverCh_Click()
InetGetPics True
End Sub

Private Sub mGetCoverSel_Click()
If SelCount < 1 Then Exit Sub
InetGetPics False
End Sub

Private Sub mGotoURL_Click()
Dim temp As Long
Dim strPath As String
Dim site As String

On Error GoTo err
If SelCount < 1 Then Exit Sub
strPath = Space$(255)

site = CheckNoNull("MovieURL")
If Len(site) = 0 Then Exit Sub

temp = FindExecutable(site, "", strPath)
Select Case temp
Case 31
    myMsgBox msgsvc(17), vbInformation, , Me.hwnd
    Exit Sub
Case 2
End Select

temp = ShellExecute(GetDesktopWindow(), "open", site, vbNull, vbNull, 1)
ToDebug "goMURL=" & temp

Exit Sub
err:
ToDebug "Err_goMURL: " & err.Description
End Sub

Public Sub mGroup_Click(Index As Integer)
'GroupedFlag = True
GroupColumnHeader = vbNullString
GroupInd = Index - 1
Select Case GroupInd
Case -1
    '    GroupedFlag = False
    GroupColumnHeader = NamesStore(8)
    GroupField = vbNullString
    '    InitFlag = True
    '    TabLVHid_Click
    '    Exit Sub
Case dbMovieNameInd: GroupField = "MovieName": GroupColumnHeader = ListView.ColumnHeaders(dbMovieNameInd + 1)
Case dbLabelInd: GroupField = "Label": GroupColumnHeader = ListView.ColumnHeaders(dbLabelInd + 1)
Case dbGenreInd: GroupField = "Genre": GroupColumnHeader = ListView.ColumnHeaders(dbGenreInd + 1)
Case dbYearInd: GroupField = "Year": GroupColumnHeader = ListView.ColumnHeaders(dbYearInd + 1)
Case dbCountryInd: GroupField = "Country": GroupColumnHeader = ListView.ColumnHeaders(dbCountryInd + 1)
Case dbDirectorInd: GroupField = "Director": GroupColumnHeader = ListView.ColumnHeaders(dbDirectorInd + 1)
Case dbActerInd: GroupField = "Acter": GroupColumnHeader = ListView.ColumnHeaders(dbActerInd + 1)
Case dbTimeInd: GroupField = "Time": GroupColumnHeader = ListView.ColumnHeaders(dbTimeInd + 1)
Case dbResolutionInd: GroupField = "Resolution": GroupColumnHeader = ListView.ColumnHeaders(dbResolutionInd + 1)
Case dbAudioInd: GroupField = "Audio": GroupColumnHeader = ListView.ColumnHeaders(dbAudioInd + 1)
Case dbFpsInd: GroupField = "FPS": GroupColumnHeader = ListView.ColumnHeaders(dbFpsInd + 1)
Case dbFileLenInd: GroupField = "FileLen": GroupColumnHeader = ListView.ColumnHeaders(dbFileLenInd + 1)
Case dbCDNInd: GroupField = "CDN": GroupColumnHeader = ListView.ColumnHeaders(dbCDNInd + 1)
Case dbMediaTypeInd: GroupField = "MediaType": GroupColumnHeader = ListView.ColumnHeaders(dbMediaTypeInd + 1)
Case dbVideoInd: GroupField = "Video": GroupColumnHeader = ListView.ColumnHeaders(dbVideoInd + 1)
Case dbSubTitleInd: GroupField = "SubTitle": GroupColumnHeader = ListView.ColumnHeaders(dbSubTitleInd + 1)
Case dbLanguageInd: GroupField = "Language": GroupColumnHeader = ListView.ColumnHeaders(dbLanguageInd + 1)
Case dbRatingInd: GroupField = "Rating": GroupColumnHeader = ListView.ColumnHeaders(dbRatingInd + 1)
Case dbFileNameInd: GroupField = "FileName": GroupColumnHeader = ListView.ColumnHeaders(dbFileNameInd + 1)
Case dbDebtorInd: GroupField = "Debtor": GroupColumnHeader = ListView.ColumnHeaders(dbDebtorInd + 1)
Case dbsnDiskInd: GroupField = "snDisk": GroupColumnHeader = ListView.ColumnHeaders(dbsnDiskInd + 1)
Case dbOtherInd: GroupField = "Other": GroupColumnHeader = ListView.ColumnHeaders(dbOtherInd + 1)
Case dbCoverPathInd: GroupField = "CoverPath": GroupColumnHeader = ListView.ColumnHeaders(dbCoverPathInd + 1)
Case dbMovieURLInd: GroupField = "MovieURL": GroupColumnHeader = ListView.ColumnHeaders(dbMovieURLInd + 1)

End Select
If right$(GroupColumnHeader, 2) = " >" Or right$(GroupColumnHeader, 2) = " <" Then GroupColumnHeader = left$(GroupColumnHeader, Len(GroupColumnHeader) - 2)
GroupColumnHeader = "<" & GroupColumnHeader & ">"
tvGroup.ColumnHeaders(1).Text = GroupColumnHeader
FillTVGroup
End Sub

Private Sub mHist_Click(Index As Integer)
Dim tmpLVCur As Long
If rs Is Nothing Then Exit Sub
If rs.RecordCount = 0 Then Exit Sub
If ListView.ListItems.Count = 0 Then Exit Sub
If Len(mHist(Index).Tag) = 0 Then Exit Sub

tmpLVCur = GotoLV(mHist(Index).Tag)
If tmpLVCur > -1 Then    'если есть в списке
    CurSearch = tmpLVCur
    RSGoto mHist(Index).Tag

    ListView.MultiSelect = False
    Set ListView.SelectedItem = FrmMain.ListView.ListItems(CurSearch)
    ListView.MultiSelect = True

    If FrameView.Visible Then FrmMain.ListView.SelectedItem.EnsureVisible    ': LVCLICK

    LVCLICK
End If
End Sub

Private Sub mHistClear_Click()
Dim i As Integer
'очистить историю
For i = 0 To nHistory
    arrHistoryKeys(i) = vbNullString
    arrHistoryTitles(i) = vbNullString
Next i
'mHistClear.Enabled = False
End Sub

Private Sub mInvCh_Click()
'инверт пометок
If ListView.ListItems.Count < 1 Then Exit Sub

Dim Itm As ListItem

For Each Itm In ListView.ListItems
    Itm.Checked = Not Itm.Checked
    ListView_ItemCheck Itm
Next
LVCLICK
End Sub

Private Sub mInvSel_Click()
'инвертировать выделенное
If ListView.ListItems.Count < 1 Then Exit Sub

Dim SelIt As Boolean
Dim Itm As ListItem

For Each Itm In ListView.ListItems
    If Itm.Index = CurSearch Then
        If Itm.Selected Then
            SelIt = False
        Else
            SelIt = True
        End If
    End If

    Itm.Selected = Not Itm.Selected
Next

ListView.ListItems(CurSearch).Selected = True
ListView.ListItems(CurSearch).Selected = SelIt

LVCLICK
End Sub

Private Sub mnuAddNewAuto_Click()
If BaseReadOnly Or BaseReadOnlyU Then
    myMsgBox msgsvc(24), vbInformation, , Me.hwnd
    Exit Sub
End If
frmAuto.Show 1, Me
'If LastVMI <> 1 Then VerticalMenu_MenuItemClick 1
End Sub

Private Sub mnuCard_Click()
DoEvents
mnuCard.Checked = Not mnuCard.Checked
Opt_UCLV_Vis = mnuCard.Checked
If Not Opt_UCLV_Vis Then UCLV.Clear
Form_Resize
LVCLICK

optsaved = False
End Sub

Private Sub mnuCopyRow_Click()
'сделать копию текущей строки
If SelCount < 1 Then Exit Sub
Call MakeDupCurrent
End Sub

Private Sub mnuCoverCopy_Click()
Clipboard.Clear
'PicFaceV.Picture = PicFaceV.Image
Clipboard.SetData PicCoverPaper.Picture    ', vbCFBitmap
End Sub

Private Sub mnuCoverSave_Click()
'сохранить из картинки в файл обложку печати
SavePicFromPic PicCoverPaper, FrmMain.hwnd
End Sub

Private Sub mnuGroup_Click()
Dim ret As Long

mnuGroup.Checked = Not mnuGroup.Checked
Opt_Group_Vis = mnuGroup.Checked
Form_Resize
'Form_Resize
If Opt_Group_Vis Then
    'FillTVGroup
    If Not GroupedFlag Then mGroup_Click 0    'очистить
Else
    If GroupedFlag Then
        ret = myMsgBox(msgsvc(42), vbYesNo, , FrmMain.hwnd)    'оставить группировку?
        If ret = vbNo Then InitFlag = True    'перечитаем базу на выходе из опций
        TabLVHid.Tabs(CurrentBaseIndex).Selected = True
    End If
End If

optsaved = False
End Sub

Private Sub mnuLabelCheck_Click()
If CheckCount > 0 Then Call MNU_Label(False)    'Check
End Sub

Private Sub mnuLabelSel_Click()
If SelCount < 1 Then Exit Sub
Call MNU_Label(True)    'sel
End Sub


Private Sub mnuMovieCopyClip_Click()
Clipboard.Clear
Clipboard.SetData m_cAVI.FramePicture(frmEditor.Position.Value, AviWidth, AviHeight)
End Sub

Private Sub mnuMovieSaveFrame_Click()
'сохранить в файл bmp кадр из окна movie avi
Dim dib As cDIB
Dim pDIB As Long    'pointer to packed DIB in memory

'create a DIB class to load the frames into
Set dib = New cDIB
pDIB = AVIStreamGetFrame(From_m_pGF, frmEditor.Position.Value)  'returns "packed DIB"
If dib.CreateFromPackedDIBPointer(pDIB) Then
    'Call dib.WriteToFile(App.Path & "\" & Position.Value& ".bmp")
    Call dib.WriteToFile(pSaveDialogBMP(DTitle:=frmEditor.ComSaveRec.Caption))
End If

Set dib = Nothing
End Sub

Private Sub mnuUndo_Click()
Call SendMessage(ActiveControl.hwnd, WM_UNDO, 0, ByVal 0&)
End Sub

Private Sub mPutThisActer_Click()
'поместить выделенного ненайденного актера в базу актеров
If Len(sPerson) = 0 Then Exit Sub

If PutActName(sPerson) Then
    mnuShowThisActer.Enabled = True
Else
    mnuShowThisActer.Enabled = False
End If
mPutThisActer.Enabled = False

End Sub

Private Sub mSel2Excel_Click()
If SelCount < 1 Then Exit Sub
DoEvents
Export2Excel False
End Sub

Private Sub mSelCh_Click()
'выделить помеченное
If CheckCount = 0 Then Exit Sub
Dim SelIt As Boolean
Dim Itm As ListItem

For Each Itm In ListView.ListItems

    If Itm.Index = CurSearch Then
        If Itm.Selected Then
            SelIt = True
        Else
            If Itm.Checked Then SelIt = True Else SelIt = False
        End If
    End If

    If Itm.Checked Then
        Itm.Selected = True
        'ListView_ItemCheck Itm
    End If
Next

ListView.ListItems(CurSearch).Selected = True
ListView.ListItems(CurSearch).Selected = SelIt
LVCLICK

End Sub

Private Sub mSR_Click()
If ListView.ListItems.Count < 1 Then Exit Sub
frmSR.Show , Me
End Sub

Private Sub mUnChSel_Click()
'не пометить выделенное
If SelCount < 1 Then Exit Sub

Dim Itm As ListItem

For Each Itm In ListView.ListItems
    If Itm.Selected Then
        Itm.Checked = False
        ListView_ItemCheck Itm
    End If
Next
LVCLICK

End Sub

Private Sub mUnSelCh_Click()
'не выделять помеченное
If CheckCount = 0 Then Exit Sub

Dim Itm As ListItem

For Each Itm In ListView.ListItems
    If Itm.Checked Then
        Itm.Selected = False
        'ListView_ItemCheck Itm
    End If
Next
LVCLICK

End Sub

Private Sub mValid_Click()


'Показать только проблемные строки, с несоответствием метка-серийник
Dim UniSelect As String    'уникальные связки
'1             a
'1             b
'2             c
'2             d
'3             e
'4             e
'5             f
'6             g
Dim strSQL1 As String, strSQL2 As String    'выбор простых (метка1, серийник2) из уник. связок только с каунтом > 1
'это и есть проблемные строки
'1   e
'2
Dim strSQL As String    'выборка всего In простых

DoEvents

UniSelect = "(select label,sndisk from storage group by label,sndisk)"
strSQL1 = "(select label from " & UniSelect & " group by label HAVING Count(Label) > 1)"
strSQL2 = "(select sndisk from " & UniSelect & " group by sndisk HAVING Count(sndisk) > 1)"
strSQL = "select * From Storage Where (Label In (" & strSQL1 & ") Or sndisk In (" & strSQL2 & "))"

'Debug.Print strSQL

Screen.MousePointer = vbHourglass
On Error GoTo err
Set rs = DB.OpenRecordset(strSQL)

FillListView
'свой флаг
FiltValidationFlag = True
FrmMain.ComFilter.BackColor = &HC0C0FF

Screen.MousePointer = vbNormal
Exit Sub

err:
Screen.MousePointer = vbNormal
ToDebug "Err_mValid:" & err.Description
End Sub


Private Sub mWWW1Hid_Click()
ActWWWsite = "http://www.kinopoisk.ru/index.php?level=7&m_act%5Bfrom%5D=forma&m_act%5Bwhat%5D=actor&m_act%5Bfind%5D="
ComRHid(0).Caption = "K"
End Sub
Private Sub mWWW2Hid_Click()
ActWWWsite = "http://www.world-art.ru/search.php?global_sector=people&name="
ComRHid(0).Caption = "W"
End Sub



Private Sub PicActFoto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If ActNewFlag Or ActEditFlag Then Exit Sub
If GetForegroundWindow = Me.hwnd Then
    If ActiveControl.name <> "PicActFoto" Then PicActFoto.SetFocus
End If
End Sub



Private Sub PicActFoto_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
PicActFoto.Picture = Data.GetData(vbCFDIB)
SavePicActFlag = True
' ComActSave.BackColor = &HC0C0E0
PicActFotoScroll_Resize
End Sub

Private Sub PicActFoto_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
' A drop is OK only if bitmap data is available.
If Data.GetFormat(vbCFBitmap) Or Data.GetFormat(vbCFDIB) Then
    If ActFlag Then Effect = vbDropEffectCopy Else Effect = vbDropEffectNone
Else
    Effect = vbDropEffectNone
End If

End Sub

Private Sub PicActFotoScroll_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
PicActFoto.Picture = Data.GetData(vbCFDIB)
SavePicActFlag = True
' ComActSave.BackColor = &HC0C0E0
PicActFotoScroll_Resize
End Sub

Private Sub PicActFotoScroll_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
' A drop is OK only if bitmap data is available.
If Data.GetFormat(vbCFBitmap) Or Data.GetFormat(vbCFDIB) Then
    If ActFlag Then Effect = vbDropEffectCopy Else Effect = vbDropEffectNone
Else
    Effect = vbDropEffectNone
End If

End Sub



Private Sub PicCoverPaper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Me.PopupMenu Me.popCoverHid, vbPopupMenuCenterAlign
End If
End Sub



Private Sub TextActBio_GotFocus()
'проверить есьь ли уже в списке
'Dim word As String
Dim temp As String
Dim itmX As ListItem
Dim i As Long    ', j As Integer
Dim VsegoSlov As Integer, Podhodit As Integer

If ActNewFlag Then
    TextActName.Text = sTrimChars(TextActName.Text, vbNewLine)
    If TextActName.Text <> vbNullString Then

        For Each itmX In LVActer.ListItems
            i = i + 1
            temp = TextActName.Text
            VsegoSlov = 0: Podhodit = 0

            Do While temp <> vbNullString
                VsegoSlov = VsegoSlov + 1
                If InStr(1, itmX.Text, GetWord(temp), vbTextCompare) <> 0 Then Podhodit = Podhodit + 1
            Loop    'temp <> vbNullString

            If Podhodit = VsegoSlov Then

                ActNotManualClick = True
                Set LVActer.SelectedItem = LVActer.ListItems.Item(i)
                'LVActer.ListItems(i).EnsureVisible
                LV_EnsureVisible LVActer, i
                ActNotManualClick = False
                Exit Sub
            End If

            'Debug.Print "itmX.Text=" & itmX.Text, "VsegoSlov=" & VsegoSlov, "Podhodit=" & Podhodit, "pos=" & i
        Next    'itmX


    End If    'TextActName.Text <> vbNullString
End If    'actflag
End Sub

Private Sub TextActBio_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim plus As String
'IsTypingAct = False
'TextActBio.Text = sTrimChars(TextActBio.Text, vbNewLine)
If Data.GetFormat(1) Then
    If TextActBio.Text <> vbNullString Then plus = TextActBio.Text & vbCrLf
    TextActBio.Text = plus & Data.GetData(1)
End If
'TextActBio.Text = StrConv(TextActBio.Text, vbProperCase, LCID)

End Sub

Private Sub TextActBio_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone
If Not ActFlag Then Effect = vbDropEffectNone
End Sub

Private Sub TextActName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim plus As String
'IsTypingAct = False
TextActName.Text = sTrimChars(TextActName.Text, vbNewLine)
If Data.GetFormat(1) Then
    If X > TextActName.Width / 2 Then
        If TextActName.Text <> vbNullString Then plus = TextActName.Text & " "
        TextActName.Text = plus & Data.GetData(1)
    Else
        If TextActName.Text <> vbNullString Then plus = " " & TextActName.Text
        TextActName.Text = Data.GetData(1) & plus
    End If

End If
TextActName.Text = StrConv(TextActName.Text, vbProperCase, LCID)
TextActName.Text = Replace(TextActName.Text, "  ", " ")
TextActName.Text = Trim$(TextActName.Text)

End Sub

Private Sub TextActName_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
If Not Data.GetFormat(vbCFText) Then Effect = vbDropEffectNone
If Not ActFlag Then Effect = vbDropEffectNone

End Sub

Private Sub ComCancelAct_Click()
ComCancelAct.Enabled = False

Call ActEditorColors(False)

ComActSave.Enabled = False
ComActPast.Enabled = False
ComActFile.Enabled = False
ComActFotoDel.Enabled = False
ComActEdit.Enabled = True
ComActDel.Enabled = True
ComAddAct.Enabled = True
TextSearchLVActTypeHid.Enabled = True
'PicActFotoScroll.Visible = False
comActFilt.Enabled = True
chActFotoScale.Enabled = True

If ars.EditMode Then ars.CancelUpdate
ActEditFlag = False: ActNewFlag = False: ActFlag = False
If ars.RecordCount = 0 Then Exit Sub

LVActClick

ComActSave.BackColor = &HFFFFFF

ToDebug "Отмена изменения"
End Sub



Private Sub ComFilter_Click()
On Error Resume Next
'If Opt_LoadOnlyTitles Then
'    FrmFilter.Show 0, FrmMain
'Else
'    'если заполнено все
FrmFilter.FillFilter
FrmFilter.Show 0, FrmMain
'End If
End Sub


Private Sub ComFind_Click()
If ListView.ListItems.Count < 1 Then Exit Sub

Timer2.Enabled = False
ComNext.Enabled = True
'LockWindowUpdate ListView.hwnd
Dim FindText As String
'FindText = UCase$(TextFind.Text)
FindText = TextFind.Text

Screen.MousePointer = vbHourglass
ToDebug "Поиск: " & FindText
ListView.MultiSelect = False

If CombFind.ListIndex < 0 Then CombFind.ListIndex = 0

Select Case CombFind.ItemData(CombFind.ListIndex)
Case 0
    SearchLV dbMovieNameInd, FindText, vbTextCompare
Case 1
    SearchLV dbLabelInd, FindText, vbTextCompare
Case 2
    SearchLV dbDirectorInd, FindText, vbTextCompare
Case 3
    SearchLV dbActerInd, FindText, vbTextCompare
Case 4
    SearchLV dbGenreInd, FindText, vbTextCompare
Case 5
    'аннотация
    ComShowAn_Click
    SearchNextDB dbAnnotationInd, FindText, 1, True, vbTextCompare
Case 6
    SearchLV dbsnDiskInd, FindText, vbTextCompare
Case 7
    SearchLV dbFileNameInd, FindText, vbTextCompare
Case 8
    SearchLV dbDebtorInd, FindText, vbTextCompare
End Select

ListView.MultiSelect = True

Screen.MousePointer = vbNormal
If FrameView.Visible Then Timer2.Enabled = True
End Sub











Private Sub ComNext_Click()
If ListView.ListItems.Count < 1 Then Exit Sub

Dim FindText As String

Timer2.Enabled = False

Screen.MousePointer = vbHourglass

'LockWindowUpdate ListView.hwnd
'FindText = UCase$(TextFind.Text)
FindText = TextFind.Text

'ToDebug "Продолжить поиск: " & FindText
If ChMarkFindHid Then: Else ListView.MultiSelect = False

If CombFind.ListIndex < 0 Then CombFind.ListIndex = 0

Select Case CombFind.ItemData(CombFind.ListIndex)
Case 0
    FindNextLV 0, FindText
Case 1
    FindNextLV dbLabelInd, FindText
Case 2
    FindNextLV dbDirectorInd, FindText
Case 3
    FindNextLV dbActerInd, FindText
Case 4
    FindNextLV dbGenreInd, FindText
Case 5
    ComShowAn_Click
    SearchNextDB dbAnnotationInd, FindText, ListView.SelectedItem.Index, False, vbTextCompare
Case 6
    FindNextLV dbsnDiskInd, FindText

Case 7
    FindNextLV dbFileNameInd, FindText

Case 8
    FindNextLV dbDebtorInd, FindText

End Select

ListView.MultiSelect = True
'LvClick
If FrameView.Visible Then Timer2.Enabled = True
'LockWindowUpdate 0
Screen.MousePointer = vbNormal

End Sub

Public Sub GetLangEditor()
Dim Contrl As Control
Dim i As Integer
'Dim tmp As String, R_lang As String, E_lang As String
'Dim E_langFlag As Boolean
'Dim iPos As Integer

On Error GoTo err

If Not FileExists(userFile) Then
    'создать файл user.lng
    MakeUserFile userFile
End If

If Len(lngFileName) <> 0 Then
Else
    Exit Sub
End If

Screen.MousePointer = vbHourglass

'ToDebug "Locale_Editor " & lngFileName
'LockWindowUpdate Me.hwnd

With frmEditor
    For Each Contrl In frmEditor.Controls
        If TypeOf Contrl Is Label Then        '                           Label
            If Contrl.name = "LTech" Then
                .LTech(Contrl.Index).Caption = ReadLang("LTech(" & Contrl.Index & ").Caption", .LTech(Contrl.Index).Caption)
            ElseIf Contrl.name = "LFilm" Then
                .LFilm(Contrl.Index).Caption = ReadLang("LFilm(" & Contrl.Index & ").Caption", .LFilm(Contrl.Index).Caption)
            Else
                Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
            End If
        End If

        If TypeOf Contrl Is Frame Then        '                           Frame
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
        End If
        AddEditCapt = .FrameAddEdit.Caption

        If TypeOf Contrl Is CommandButton Then        '                   CommandButton
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
            ' Contrl.ToolTipText = ReadLang(Contrl.name & ".ToolTip", Contrl.ToolTipText)
        End If


        If TypeOf Contrl Is XpB Then        '                               XPB
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
            ' Contrl.ToolTipText = ReadLang(Contrl.name & ".ToolTip", Contrl.ToolTipText)
            Contrl.pInitialize
        End If


        If TypeOf Contrl Is ComboBox Then             'Combo

            If Contrl.name = "CombFind" Then
                CombFind.Clear
                For i = 0 To 8
                    CombFind.AddItem (ReadLang("CombFind.Item" & i))
                    CombFind.ItemData(i) = i
                Next i
            End If

        End If

        If TypeOf Contrl Is CheckBox Then        '                         CheckBox
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
        End If

        If TypeOf Contrl Is TabStrip Then        '                        TabStrAdEd
            Select Case Contrl.name
            Case "TabStrAdEd"
                Contrl.Tabs(1).Caption = ReadLang(Contrl.name & ".Tabs(1).Caption", Contrl.Tabs(1).Caption)
                Contrl.Tabs(2).Caption = ReadLang(Contrl.name & ".Tabs(2).Caption", Contrl.Tabs(2).Caption)
                Contrl.Tabs(3).Caption = ReadLang(Contrl.name & ".Tabs(3).Caption", Contrl.Tabs(3).Caption)
            End Select
        End If        '(TypeOf Contrl

    Next
End With

'Добавить пользовательские
If FileExists(userFile) Then
    With frmEditor
        Call FillUserCombo("Genre", .ComboGenre)  'жанр
        Call FillUserCombo("Country", .ComboCountry)  'страна
        Call FillUserCombo("Language", .TextLang)  'язык
        Call FillUserCombo("Subtitle", .TextSubt)  'субтитры
        Call FillUserCombo("Comments", .ComboOther)  'качество
        Call FillUserCombo("Media", .ComboNos)  'носитель
        Call FillUserCombo("Site", .ComboSites)  'сайты
        Call FillUserCombo("Site", .cBasePicURL)  'сайты
    End With
End If


With frmEditor                                      'UserControl LVaddon
    UCLV.Controls("LAct") = .LFilm(5).Caption & ":"
    UCLV.Controls("LAudio") = .LTech(6).Caption & ":"        'LAudio.Caption & ":"
    UCLV.Controls("LCountry") = .LFilm(3).Caption & ":"
    UCLV.Controls("LGenre") = .LFilm(2).Caption & ":"
    UCLV.Controls("LMName") = .LFilm(0).Caption & ":"
    UCLV.Controls("LNcd") = .LTech(0).Caption & ":"        'LNcd.Caption & ":"
    UCLV.Controls("LOther") = .LFilm(8).Caption & ":"
    UCLV.Controls("LRes") = .LFilm(4).Caption & ":"
    UCLV.Controls("LTime") = .LTech(2).Caption & ":"        'LTime.Caption & ":"
    UCLV.Controls("LVideo") = .LTech(3).Caption & ":"        'LVideo.Caption & ":"
    UCLV.Controls("LLang") = .LFilm(10).Caption & ":"
    UCLV.Controls("LSubt") = .LFilm(11).Caption & ":"
    UCLV.Controls("LRate") = .LFilm(7).Caption & ":"
    UCLV.Controls("LFile") = .LTech(7).Caption & ":"        'LFile.Caption & ":"
    UCLV.Controls("LLabel") = .LFilm(1).Caption & ":"
    UCLV.Controls("LDebtor") = .LTech(11).Caption & ":"        'LUser.Caption & ":"
End With

For Each Contrl In UCLV.Controls
    If TypeOf Contrl Is Label Then
        'сменить курсор
        Contrl.MousePointer = 99
        Contrl.MouseIcon = FrmMain.ImageList.ListImages("LArr").Picture
    End If
Next


frmEditor.Caption = "SurVideoCatalog: " & VerticalMenu.Controls("LVMB")(1)
frmEditor.Icon = FrmMain.Icon

Screen.MousePointer = vbNormal

Exit Sub

err:
Screen.MousePointer = vbNormal
If err <> 0 Then
    ToDebug "Err_EdLCh:" & err.Description
    Debug.Print "Err_EdLCh:" & err.Description
    On Error Resume Next
    Resume Next
End If
End Sub



Public Sub LangChange()
Dim Contrl As Control
Dim i As Integer
'Dim temp As String
Dim tmp As String, R_lang As String, E_lang As String
Dim E_langFlag As Boolean

Dim iPos As Integer

On Error GoTo err

If Not FileExists(lngFileName) Then
    '?не нашли параметра \последний язык\ в глобале или первый запуск

    'применить русский, если есть
    R_lang = App.Path & "\rus.lng"
    E_lang = App.Path & "\eng.lng"
    If FileExists(E_lang) Then E_langFlag = True

    If FileExists(R_lang) Then
        'нашли русский
        lngFileName = R_lang
    ElseIf E_langFlag Then
        'нашли англ
        lngFileName = E_lang
    Else
        'продолжить без языка
        Call myMsgBox("Не найден файл локализации! Переустановите программу." & vbCrLf & "Reinstall application! Language file not found: " & vbCrLf & lngFileName, vbCritical, , Me.hwnd)
        lngFileName = vbNullString
    End If

    If Len(lngFileName) <> 0 Then
        If LCID = 1049 And lngFileName = R_lang Then
            'уникоды для России и русский файл есть (и уже выбран выше)
        ElseIf E_langFlag Then
            'установить инглиш
            lngFileName = E_lang
        End If
    End If
End If
'lngFileName прописать в глобал ?


If Not FileExists(userFile) Then
    'создать файл user.lng
    MakeUserFile userFile
End If

If Len(lngFileName) <> 0 Then
    'не перечитывать
    If LastLangFile = lngFileName Then Exit Sub
Else
    Exit Sub
End If

Screen.MousePointer = vbHourglass

ToDebug "локализация " & lngFileName

'Call SendMessage(Me.hWnd, WM_SETREDRAW, False, ByVal 0&)
'LockWindowUpdate Me.hWnd

For Each Contrl In FrmMain.Controls
    'If (TypeOf Contrl Is ComboBox _
     Or TypeOf Contrl Is ListView _
     Or TypeOf Contrl Is Label _
     Or TypeOf Contrl Is Frame _
     Or TypeOf Contrl Is Slider _
     Or TypeOf Contrl Is CommandButton _
     Or TypeOf Contrl Is PictureBox _
     Or TypeOf Contrl Is Image _
     Or TypeOf Contrl Is VerticalMenu _
     ) Then
    'msgbox
    'popup menu
    'checkbox

    If right$(Contrl.name, 3) <> "Hid" Then

        If Contrl.name = "VerticalMenu" Then        '                      VerticalMenu
            VerticalMenu.Controls("LVMB")(0) = ReadLang(Contrl.name & ".MenuItemCaption" & 1, VerticalMenu.Controls("LVMB")(0).Caption)
            VerticalMenu.Controls("LVMB")(1) = ReadLang(Contrl.name & ".MenuItemCaption" & 2, VerticalMenu.Controls("LVMB")(1).Caption)
            VerticalMenu.Controls("LVMB")(2) = ReadLang(Contrl.name & ".MenuItemCaption" & 3, VerticalMenu.Controls("LVMB")(2).Caption)
            VerticalMenu.Controls("LVMB")(3) = ReadLang(Contrl.name & ".MenuItemCaption" & 4, VerticalMenu.Controls("LVMB")(3).Caption)
            VerticalMenu.Controls("LVMB")(4) = ReadLang(Contrl.name & ".MenuItemCaption" & 5, VerticalMenu.Controls("LVMB")(4).Caption)
            VerticalMenu.Controls("LVMB")(5) = ReadLang(Contrl.name & ".MenuItemCaption" & 6, VerticalMenu.Controls("LVMB")(5).Caption)
            VerticalMenu.Controls("LVMB")(6) = ReadLang(Contrl.name & ".MenuItemCaption" & 7, VerticalMenu.Controls("LVMB")(6).Caption)
            VerticalMenu.Controls("LVMB")(7) = ReadLang(Contrl.name & ".MenuItemCaption" & 8, VerticalMenu.Controls("LVMB")(7).Caption)
        End If        'Contrl Is VerticalMenu

        If TypeOf Contrl Is Label Then        '                           Label
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
        End If

        If TypeOf Contrl Is Frame Then        '                           Frame
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
        End If




        If TypeOf Contrl Is MyFrame Then        '                           MyFrame
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
        End If

        If TypeOf Contrl Is ListView Then        '                        ListView
            If Contrl.name = "ListView" Then

                For i = 1 To ListView.ColumnHeaders.Count
                    ListView.ColumnHeaders(i).Text = ReadLang(ListView.name & ".CH" & i, ListView.ColumnHeaders(i).Text)
                    'ListView.ColumnHeaders(i).Icon = "AVI"

                Next i
            End If
            If Contrl.name = "tvGroup" Then Contrl.ColumnHeaders(1).Text = ReadLang(Contrl.name & ".CH1", Contrl.ColumnHeaders(1).Text)
        End If

        If TypeOf Contrl Is CommandButton Then        '                   CommandButton
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
            ' Contrl.ToolTipText = ReadLang(Contrl.name & ".ToolTip", Contrl.ToolTipText)
        End If


        If TypeOf Contrl Is XpB Then        '                               XPB
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
            ' Contrl.ToolTipText = ReadLang(Contrl.name & ".ToolTip", Contrl.ToolTipText)
            Contrl.pInitialize
        End If


        If TypeOf Contrl Is ComboBox Then             'Combo

            If Contrl.name = "CombFind" Then
                CombFind.Clear
                For i = 0 To 8
                    CombFind.AddItem (ReadLang("CombFind.Item" & i))
                    CombFind.ItemData(i) = i
                Next i
            End If

        End If

        'If TypeOf Contrl Is OptionButton Then '                     OptionButton
        'For i = OptHtml.LBound To OptHtml.UBound
        'OptHtml(i).Caption = ReadLang(Contrl.name & i & ".Caption")
        'Next
        'End If

        If TypeOf Contrl Is Menu Then        '                                     Menu
            If Contrl.name <> "mGroup" Then
                Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
                'Debug.Print Contrl.name, Contrl.Caption

                On Error Resume Next
                'отрезать название от шортката, скормить в меню
                tmp = Contrl.Caption: iPos = InStr(tmp, vbTab)
                If iPos > 0 Then tmp = left$(tmp, iPos - 1)

                Select Case Contrl.name
                Case "mnuCoverSave", "mnuMovieSaveFrame", "mnuSaveFace", "mnuSaveFoto", "mnuSavePic"
                    cIM.IconIndex(tmp) = ImageList.ListImages("SAVE_ICON").Index - 1
                Case "mnuCoverCopy", "mnuMovieCopyClip", "mnuCopyFace", "mnuCopyFoto", "mnuCopyPic"
                    cIM.IconIndex(tmp) = ImageList.ListImages("mnuCopyLV").Index - 1
                Case Else
                    cIM.IconIndex(tmp) = ImageList.ListImages(Contrl.name).Index - 1
                End Select
                '            If err = 0 Then
                '                Debug.Print
                '            End If
                err.Clear
                On Error GoTo err

            End If    ' Contrl.name <> "mGroup"
        End If


        If TypeOf Contrl Is CheckBox Then        '                         CheckBox
            Contrl.Caption = ReadLang(Contrl.name & ".Caption", Contrl.Caption)
        End If

        If TypeOf Contrl Is TabStrip Then        '                        TabStripCover
            Select Case Contrl.name
            Case "TabStripCover"
                Contrl.Tabs(1).Caption = ReadLang(Contrl.name & ".Tabs(1).Caption", Contrl.Tabs(1).Caption)
                Contrl.Tabs(2).Caption = ReadLang(Contrl.name & ".Tabs(2).Caption", Contrl.Tabs(2).Caption)
                Contrl.Tabs(3).Caption = ReadLang(Contrl.name & ".Tabs(3).Caption", Contrl.Tabs(3).Caption)
                Contrl.Tabs(4).Caption = ReadLang(Contrl.name & ".Tabs(4).Caption", Contrl.Tabs(4).Caption)
                Contrl.Tabs(5).Caption = ReadLang(Contrl.name & ".Tabs(5).Caption", Contrl.Tabs(5).Caption)
            End Select
        End If        '(TypeOf Contrl

    End If        '<> "Hid"
Next                                            '                               конец


'                                                           NamesStore()
For i = 1 To ListView.ColumnHeaders.Count - 1        'mGroup.Count - 1 'indexpole
    mGroup(i).Caption = ListView.ColumnHeaders(i)
Next i
mGroup(0).Caption = "* " & ReadLang("NamesStore8") & " *"


For i = 0 To UBound(NamesStore)
    NamesStore(i) = ReadLang("NamesStore" & i)
Next i

'имена полей , переведенные названия полей
ReDim TranslatedFieldsNames(ListView.ColumnHeaders.Count - 1)
For i = 0 To ListView.ColumnHeaders.Count - 2
    TranslatedFieldsNames(i) = ListView.ColumnHeaders(i + 1).Text
Next i
TranslatedFieldsNames(i) = NamesStore(11)    'frmEditor.LFilm(9).Caption    'LAnnot.Caption        '+ аннотация

'                                                           msgbox
For i = 1 To UBound(msgsvc)
    msgsvc(i) = Change2lfcr(ReadLang("msg" & i))
Next i

'переменные для изменяемых названий

mnuLVCheckedCaption = mnuLVChecked.Caption
mnuLVSelectedCaption = mnuLVSelected.Caption

FrameViewCaption = FrameView.Caption
FrameView.Caption = FrameViewCaption & " " & ListView.ListItems.Count & " )"

FrameActerCaption = FrameActer.Caption
FrameActer.Caption = FrameActerCaption & LVActer.ListItems.Count & ")"


LActMarkCountCaption = LActMarkCount.Caption
ChPrintCheckedCaption = ChPrintChecked.Caption

'LockWindowUpdate 0
'Call SendMessage(Me.hWnd, WM_SETREDRAW, True, ByVal 0&)
'Me.Refresh

LastLangFile = lngFileName

Screen.MousePointer = vbNormal

Exit Sub

err:
'Call SendMessage(Me.hWnd, WM_SETREDRAW, True, ByVal 0&)
'LockWindowUpdate 0
Screen.MousePointer = vbNormal
If err <> 0 Then
    ToDebug "Err_MainLCh:" & err.Description
    Debug.Print "Err_MainLCh:" & err.Description
End If
End Sub





Private Sub ComPlay_ShiftClick(Shift As Integer)
PlayMovieFolderFlag = True
Call ComPlay_Click
'не тут, а в коде PlayMovieFolderFlag = False
End Sub
Private Sub ComPlay_Click()
Dim a() As String
Dim tmp As String
Dim i As Integer

LstFiles.Visible = False
If SelCount < 1 Then Exit Sub
LstFiles.Clear
tmp = CheckNoNullVal(dbFileNameInd)
If Len(tmp) <> 0 Then    'не пусто
    If Tokenize04(tmp, a(), "|", False) > -1 Then
        If UBound(a) > 0 Then
            For i = 0 To UBound(a)
                LstFiles.AddItem Trim$(a(i))
            Next i
            LstFiles.ZOrder 0
            LstFiles.Visible = True
            SetListboxScrollbar LstFiles, FrmMain
            LstFiles.SetFocus
        Else
            PlayMovie tmp
        End If
    End If
Else    'спросить
    PlayMovie vbNullString
End If

End Sub


Private Sub ComOpen_Click()
OpenNewMovie
End Sub


Private Sub ComSelMovIcon_Click()
'подготовка к показу актера в фильмах
Dim j As Integer, i As Integer
Dim tmpd As String, tmpa As String, tmp As String

If rs Is Nothing Then Exit Sub
If ListBActHid.ListCount < 0 Then Exit Sub

tmpd = "Director Like "
tmpa = "Acter Like "
For j = 0 To ListBActHid.ListCount - 1
    If ListBActHid.Selected(j) Then
        i = i + 1
        tmp = ListBActHid.List(j)
        SQLCompatible tmp
        If i = ListBActHid.SelCount Then  'последняя
            tmpd = tmpd & "'*" & tmp & "*'"
            tmpa = tmpa & "'*" & tmp & "*'"
        Else
            tmpd = tmpd & "'*" & tmp & "*' And Director Like "
            tmpa = tmpa & "'*" & tmp & "*' And Acter Like "
        End If
    End If    'selected
Next j

'Debug.Print "((" & tmpd & ") Or (" & tmpa & "))"
'Debug.Print
Call FilterMovieWithPers("(" & tmpd & ") Or (" & tmpa & ")")

'заменено на FilterMovieWithPers
'clearLVIcon
'
'Dim itmX As ListItem
'Dim i As Long, j As Integer
'Dim temp As String
'Dim MarkSelect As Boolean
'
'InitFlag = False
'
'If rs Is Nothing Then
'    'подгрузить базу
'Debug.Print "Test: rs = nothing, ComSelMovIcon_Click"
'    If LstBases_ListIndex = -1 Then
'        TabLVHid.Tabs(1).Selected = True
'    Else
'        TabLVHid.Tabs(LstBases_ListIndex + 1).Selected = True
'    End If
'End If
'
'i = 0: mcount = 0
'
''Dim itmX As ListItem
'For Each itmX In ListView.ListItems
'    i = i + 1
'    temp = itmX.SubItems(dbDirectorInd) + " " + itmX.SubItems(dbActerInd)     ' соединить режиссеров и актеров
'
'    MarkSelect = True
'    If ListBActHid.ListCount < 0 Then Set itmX = Nothing: Exit Sub
'
'    For j = 0 To ListBActHid.ListCount - 1
'        If ListBActHid.Selected(j) Then
'            If InStr(1, temp, ListBActHid.List(j), vbTextCompare) = 0 Then MarkSelect = False
'        End If    'selected
'    Next j
'
'    If MarkSelect Then
'        ListView.ListItems(i).SmallIcon = 1
'        ListView.ListItems(i).Checked = True
'
'        mcount = mcount + 1
'    End If    'MarkSelect
'
'    LActMarkCount.Caption = LActMarkCountCaption + " " + str$(mcount)
'Next        'For Each
'
'Set itmX = Nothing
End Sub

Public Sub ComShowAn_Click()
On Error Resume Next    'кнопки неактивны если в редакторе - ошибка setfocus

ComShowFa.Visible = True
ComShowAn.Visible = False

TextVAnnot.Visible = True

If Not Opt_NoSlideShow Then picScrollBoxV.Visible = False

If FrameView.Visible And (Not frmOptFlag) Then ComShowFa.SetFocus

End Sub

Public Sub ComShowFa_Click()
'If Not MultiSel Then
Set PicFaceV = Nothing

'If picScrollBoxV.Visible Then
GetPic PicFaceV, 1, "FrontFace"    '22

If NoPicFrontFaceFlag Then
    PicFaceV.Width = 0: PicFaceV.Height = 0    'если не exit: picScrollBoxV_Resize
    ''GetPic Image0, 1, "SnapShot1" '19
    If Not Opt_NoSlideShow Then Timer2.Enabled = True
    'Exit Sub
End If

picScrollBoxV_Resize

ComShowAn.Visible = True: ComShowFa.Visible = False
If FrameView.Visible And (Not frmOptFlag) Then ComShowAn.SetFocus

If Not Opt_NoSlideShow Then picScrollBoxV.Visible = True: TextVAnnot.Visible = False

'End If 'multisel
End Sub




Private Sub Form_Activate()
On Error Resume Next

If FirstActivateFlag Then
    If Opt_NoSlideShow Then picScrollBoxV_Resize
    ComShowAn_Click    'убрать из видимости обложку, надо
    Timer2_Timer
End If

FirstActivateFlag = False

End Sub

Public Sub ReadINI(iFn As String)
Dim WFD As WIN32_FIND_DATA
Dim ret As Long
Dim temp As String
Dim i As Integer, tmpi As Integer    'mzt , j As Integer
Dim glob As String
Const t = 1000
On Error GoTo err

glob = "global"
INIFILE = iFn & ".ini"
ToDebug INIFILE
'check ini
iniFileName = App.Path
If right$(iniFileName, 1) <> "\" Then iniFileName = iniFileName & "\"
iniFileName = iniFileName & INIFILE
ret = FindFirstFile(iniFileName, WFD)

If iFn = glob Then
    'GlobalFileFlagRW = True
    If ret < 0 Then
        If MakeINI(INIFILE) Then GlobalFileFlagRW = True
    Else
        ' проверить ридонли
        If WFD.dwFileAttributes And FILE_ATTRIBUTE_READONLY Then
            GlobalFileFlagRW = False
        Else
            GlobalFileFlagRW = True
        End If
        If Not GlobalFileFlagRW Then ToDebug "ROnlyINI: " & iniFileName
    End If
Else
    'INIFileFlagRW = True
    If ret < 0 Then
        'если нет спец инишника базы, сделать и перечитать
        If MakeINI(INIFILE) Then INIFileFlagRW = True
        ReadINI GetNameFromPathAndName(bdname): Exit Sub
    Else
        ' проверить ридонли
        If WFD.dwFileAttributes And FILE_ATTRIBUTE_READONLY Then
            INIFileFlagRW = False
        Else
            INIFileFlagRW = True
        End If
        If Not INIFileFlagRW Then ToDebug "Read only INI: " & iniFileName
    End If
End If
FindClose ret

'общее чтение фонтов и цветов
With ListView

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
        FontListView.Size = CDbl(Replace2Regional(temp))    'не ставить val
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
        LVBackColor = 15000275    '12648447
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



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If iFn <> glob Then                                             '-----------Конкретная база

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      Вначале  Опции

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


        'галки вкладки шрифтов
        temp = VBGetPrivateProfileString("FONT", "VMcolor", iniFileName)
        If Len(temp) <> 0 Then VMSameColor = CBool(temp) Else VMSameColor = False
        temp = VBGetPrivateProfileString("FONT", "StripedLV", iniFileName)
        If Len(temp) <> 0 Then StripedLV = CBool(temp) Else StripedLV = False
        temp = VBGetPrivateProfileString("FONT", "NoLVSelFrame", iniFileName)
        If Len(temp) <> 0 Then NoLVSelFrame = CBool(temp) Else NoLVSelFrame = False
        temp = VBGetPrivateProfileString("LIST", "LVGrid", iniFileName)
        If Len(temp) <> 0 Then Opt_ShowLVGrid = CBool(temp) Else Opt_ShowLVGrid = False


        'применить опции
        mnuCard.Checked = Opt_UCLV_Vis
        mnuGroup.Checked = Opt_Group_Vis
        If Opt_NoSlideShow Then
            FrmMain.picScrollBoxV.Visible = True: FrmMain.TextVAnnot.Visible = True
        Else
            If Me.Visible Then ComShowAn_Click    'для соответствия
        End If

        If Opt_ShowLVGrid Then
            ListView.GridLines = True
            tvGroup.GridLines = True
            LVActer.GridLines = True
        Else
            ListView.GridLines = False
            tvGroup.GridLines = False
            LVActer.GridLines = False
        End If


        '                                                                      Список  ListView
        'читать для базы позицию колонок LV 'Читать размеры полей LV
        .Visible = False
        For i = 1 To .ColumnHeaders.Count
            temp = VBGetPrivateProfileString("LIST", "P" & i, iniFileName)
            If IsNumeric(temp) Then .ColumnHeaders(i).Position = Int(Val(temp))
        Next i
        'второй раз позиционирование (с первого раза не все на местах, через ж. но...)
        For i = 1 To .ColumnHeaders.Count
            temp = VBGetPrivateProfileString("LIST", "P" & i, iniFileName)
            If IsNumeric(temp) Then .ColumnHeaders(i).Position = Int(Val(temp))
            'Debug.Print i, .ColumnHeaders(i).Position, .ColumnHeaders(i).Text
            temp = VBGetPrivateProfileString("LIST", "C" & i, iniFileName)
            If IsNumeric(temp) Then .ColumnHeaders(i).Width = Int(Val(temp)) Else .ColumnHeaders(i).Width = 1500
        Next i
        'ниже .Visible = True

        If Opt_SortOnStart Then
            'сортировка
            temp = VBGetPrivateProfileString("LIST", "LVSortColl", iniFileName)
            If IsNumeric(temp) Then LVSortColl = Int(Val(temp)) Else LVSortColl = 0
            temp = VBGetPrivateProfileString("LIST", "LVSortOrder", iniFileName)
            If IsNumeric(temp) Then LVSortOrder = Int(Val(temp)) Else LVSortOrder = 1

            ListView.SortOrder = LVSortOrder    'lvwAscending =0 ' lvwDescending =1
            'сама сортировка - в заполнении списка
            'поставить метку сортировки <>
            If LVSortColl > 0 Then .ColumnHeaders(LVSortColl).Text = ChangeLVHSortMark(.ColumnHeaders(LVSortColl))
        Else
            LVSortColl = 0: ListView.SortOrder = 1
            'перечитать названия хедеров - убрать значки сортировки ">"
            For i = 1 To ListView.ColumnHeaders.Count
                ListView.ColumnHeaders(i).Text = ReadLang("ListView.CH" & i)
            Next i

        End If

        .Visible = True

        '                                                                               Export
        For i = 0 To LstExport_ListCount    '- 1
            temp = VBGetPrivateProfileString("EXPORT", "L" & i, iniFileName)
            Select Case Trim$(LCase$(temp))
            Case "true", "1"
                LstExport_Arr(i) = True
            Case Else
                LstExport_Arr(i) = False
            End Select

        Next

        For i = 0 To 2    '012
            temp = VBGetPrivateProfileString("EXPORT", "OptHtml" & i & ".Caption", iniFileName)
            If StrComp(temp, "true", vbTextCompare) = 0 Then Opt_HtmlJpgName = i    'OptHtml(i).Value = True
        Next

        CurrentHtmlTemplate = VBGetPrivateProfileString("EXPORT", "Template", iniFileName)
        If Len(CurrentHtmlTemplate) = 0 Then
            temp = App.Path & "\Templates\svc_компакт.htm"
            If FindFirstFile(temp, WFD) <> -1 Then
                CurrentHtmlTemplate = "svc_компакт.htm"
            End If
        End If

        TxtNnOnPage_Text = VBGetPrivateProfileString("EXPORT", "NumsOnPage", iniFileName)
        If IsNumeric(TxtNnOnPage_Text) Then
            If Val(TxtNnOnPage_Text) < 1 Then TxtNnOnPage_Text = "30"
        Else
            TxtNnOnPage_Text = "30"
        End If

        ExportDelim = getExpDelim(VBGetPrivateProfileString("EXPORT", "ExportDelimiter", iniFileName))
        If Len(ExportDelim) = 0 Then ExportDelim = vbCrLf

        'chExpFolders
        temp = VBGetPrivateProfileString("EXPORT", "UseSubFolders", iniFileName)
        If Len(temp) <> 0 Then Opt_ExpUseFolders = CBool(temp) Else Opt_ExpUseFolders = False
        Opt_ExpFolder1 = VBGetPrivateProfileString("EXPORT", "SubFolder1", iniFileName)
        Opt_ExpFolder2 = VBGetPrivateProfileString("EXPORT", "SubFolder2", iniFileName)
        Opt_ExpFolder3 = VBGetPrivateProfileString("EXPORT", "SubFolder3", iniFileName)


        ComboCDHid_Text = VBGetPrivateProfileString("CD", "CDdrive", iniFileName)
        If Len(ComboCDHid_Text) = 0 Then ComboCDHid_Text = "D:\;C:\Video;C:\DVD;"

        'последние были
        temp = Int(Val(VBGetPrivateProfileString("LIST", "LastItem", iniFileName)))
        If IsNumeric(temp) Then LastInd = Int(Val(temp)) Else LastInd = 1
        'и для актеров
        temp = Int(Val(VBGetPrivateProfileString("LIST", "LastItemAct", iniFileName)))
        If IsNumeric(temp) Then CurActKey = temp & """" Else CurActKey = vbNullString

        'качество сжатия
        QJPG = Int(Val(VBGetPrivateProfileString("GLOBAL", "QJPG", iniFileName)))
        If (QJPG < 1) Or (QJPG > 100) Then QJPG = 80



        '                                                                       сплиттеры
        temp = VBGetPrivateProfileString("LIST", "LVWidth%", iniFileName)
        If IsNumeric(temp) Then LVWidth = temp
        temp = VBGetPrivateProfileString("LIST", "TVWidth", iniFileName)    '             TV
        If IsNumeric(temp) Then
            If temp <> 0 Then
                TVWidth = temp
            Else
                TVWidth = 2800
            End If
        Else
            TVWidth = 2800
        End If
        temp = VBGetPrivateProfileString("LIST", "ScrShotWidth%", iniFileName)    ' SShots width
        If IsNumeric(temp) Then
            If temp <> 0 Then
                SplitLVD = temp
            Else
                SplitLVD = 40
            End If
        Else
            SplitLVD = 40
        End If


        'в фореколор tvGroup.ColumnHeaders(1).Width = TVWidth - tvGroup.ColumnHeaders(2).Width 'мерцает


        'в readini нули подменят дефолтом, но в makeini может быть записаны нули (первый старт без ини)
        '                                                   обложка-растяжка
        '1000 - против явного бреда (были раз слишком высоки при отладке...)
        temp = VBGetPrivateProfileString("COVER", "txt_Stan_L", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_stan.l = temp Else cov_stan.l = 35.4
        If cov_stan.l > t Then cov_stan.l = 35.4
        temp = VBGetPrivateProfileString("COVER", "txt_Stan_T", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_stan.t = temp Else cov_stan.t = 179.9
        If cov_stan.t > t Then cov_stan.t = 179.9
        temp = VBGetPrivateProfileString("COVER", "txt_Stan_W", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_stan.w = temp Else cov_stan.w = 136.2
        If cov_stan.w > t Then cov_stan.w = 136.2
        temp = VBGetPrivateProfileString("COVER", "txt_Stan_H", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_stan.H = temp Else cov_stan.H = 81.6
        If cov_stan.H > t Then cov_stan.H = 81.6

        temp = VBGetPrivateProfileString("COVER", "txt_Conv_L", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_conv.l = temp Else cov_conv.l = 35.4
        If cov_conv.l > t Then cov_conv.l = 35.4
        temp = VBGetPrivateProfileString("COVER", "txt_Conv_T", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_conv.t = temp Else cov_conv.t = 45.8
        If cov_conv.t > t Then cov_conv.t = 45.8
        temp = VBGetPrivateProfileString("COVER", "txt_Conv_W", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_conv.w = temp Else cov_conv.w = 119.4
        If cov_conv.w > t Then cov_conv.w = 119.4
        temp = VBGetPrivateProfileString("COVER", "txt_Conv_H", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_conv.H = temp Else cov_conv.H = 98
        If cov_conv.H > t Then cov_conv.H = 98

        temp = VBGetPrivateProfileString("COVER", "txt_Dvd_L", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_dvd.l = temp Else cov_dvd.l = 11
        If cov_dvd.l > t Then cov_dvd.l = 11
        temp = VBGetPrivateProfileString("COVER", "txt_Dvd_T", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_dvd.t = temp Else cov_dvd.t = 65.7
        If cov_dvd.t > t Then cov_dvd.t = 65.7
        temp = VBGetPrivateProfileString("COVER", "txt_Dvd_W", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_dvd.w = temp Else cov_dvd.w = 128
        If cov_dvd.w > t Then cov_dvd.w = 128
        temp = VBGetPrivateProfileString("COVER", "txt_Dvd_H", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_dvd.H = temp Else cov_dvd.H = 128.2
        If cov_dvd.H > t Then cov_dvd.H = 128.2
        If cov_dvd.l + cov_dvd.w > 140 Then cov_dvd.l = 11: cov_dvd.w = 128
        If cov_dvd.t + cov_dvd.H > DVD_BotY Then cov_dvd.t = 65.7: cov_dvd.H = 128.2

        temp = VBGetPrivateProfileString("COVER", "txt_List_L", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_list.l = temp Else cov_list.l = 5
        If cov_list.l > t Then cov_list.l = 5
        temp = VBGetPrivateProfileString("COVER", "txt_List_T", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_list.t = temp Else cov_list.t = 10
        If cov_list.t > t Then cov_list.t = 10
        temp = VBGetPrivateProfileString("COVER", "txt_List_W", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_list.w = temp Else cov_list.w = 110
        If cov_list.w > t Then cov_list.w = 110
        temp = VBGetPrivateProfileString("COVER", "txt_List_H", iniFileName)
        If IsNumeric(temp) And temp <> 0 Then cov_list.H = temp Else cov_list.H = 280
        If cov_list.H > t Then cov_list.H = 280

        'начальные значения
        Select Case TabStripCover.SelectedItem.Index
        Case 1
            PicCoverTextWnd.left = cov_stan.l
            PicCoverTextWnd.top = cov_stan.t
            PicCoverTextWnd.Width = cov_stan.w
            PicCoverTextWnd.Height = cov_stan.H
        Case 2
            PicCoverTextWnd.left = cov_conv.l
            PicCoverTextWnd.top = cov_conv.t
            PicCoverTextWnd.Width = cov_conv.w
            PicCoverTextWnd.Height = cov_conv.H
        Case 3, 4
            PicCoverTextWnd.left = cov_dvd.l
            PicCoverTextWnd.top = cov_dvd.t
            PicCoverTextWnd.Width = cov_dvd.w
            PicCoverTextWnd.Height = cov_dvd.H
        Case 5
            PicCoverTextWnd.left = cov_list.l
            PicCoverTextWnd.top = cov_list.t
            PicCoverTextWnd.Width = cov_list.w
            PicCoverTextWnd.Height = cov_list.H
        End Select


        '                Галочки в Auto Add Movie                                    Галочки в Auto Add Movie
        temp = VBGetPrivateProfileString("AutoAdd", "chSubFolders", iniFileName)
        If Val(temp) = 1 Then ch_chSubFolders = 1 Else ch_chSubFolders = 0

        temp = VBGetPrivateProfileString("AutoAdd", "chAvi", iniFileName)
        If Val(temp) = 1 Then ch_chAviHid = 1 Else ch_chAviHid = 0
        temp = VBGetPrivateProfileString("AutoAdd", "chDS", iniFileName)
        If Val(temp) = 1 Then ch_chDSHid = 1 Else ch_chDSHid = 0
        temp = VBGetPrivateProfileString("AutoAdd", "chShots", iniFileName)
        If Val(temp) = 1 Then ch_chShots = 1 Else ch_chShots = 0
        temp = VBGetPrivateProfileString("AutoAdd", "chNoMess", iniFileName)
        If Val(temp) = 1 Then ch_chNoMess = 1 Else ch_chNoMess = 0
        temp = VBGetPrivateProfileString("AutoAdd", "cAutoClose", iniFileName)
        If Val(temp) = 1 Then ch_cAutoClose = 1 Else ch_cAutoClose = 0
        temp = VBGetPrivateProfileString("AutoAdd", "cEjectMedia", iniFileName)
        If Val(temp) = 1 Then ch_cEjectMedia = 1 Else ch_cEjectMedia = 0

        temp = VBGetPrivateProfileString("AutoAdd", "cAddCoverExt", iniFileName)
        If Val(temp) = 1 Then ch_chAutoFiles0 = 1 Else ch_chAutoFiles0 = 0
        temp = VBGetPrivateProfileString("AutoAdd", "cAddCoverAny", iniFileName)
        If Val(temp) = 1 Then ch_chAutoFiles1 = 1 Else ch_chAutoFiles1 = 0
        temp = VBGetPrivateProfileString("AutoAdd", "cAddTXTDescr", iniFileName)
        If Val(temp) = 1 Then ch_chAutoFiles2 = 1 Else ch_chAutoFiles2 = 0

        temp = VBGetPrivateProfileString("AutoAdd", "cPixTemplChange", iniFileName)
        If Val(temp) = 1 Then ch_chAutoFiles3 = 1 Else ch_chAutoFiles3 = 0
        temp = VBGetPrivateProfileString("AutoAdd", "cTxtTemplChange", iniFileName)
        If Val(temp) = 1 Then ch_chAutoFiles4 = 1 Else ch_chAutoFiles4 = 0

'        temp = VBGetPrivateProfileString("AutoAdd", "cNoAutoDups", iniFileName)
'        If Val(temp) = 1 Then ch_chAutoFiles5 = 1 Else ch_chAutoFiles5 = 0

        'Select Case False
        'Case Opt_UCLV_Vis = Old_Opt_UCLV_Vis
        If Not FirstActivateFlag Then Form_Resize
        'Case Opt_NoSlideShow = Old_Opt_NoSlideShow
        'Form_Resize
        'End Select
    End If    '                                            ----------------не global




    '                                                       ----------------In Global
    If iFn = glob Then
        '                                                                    в глобальном ини
        'LV
        'Очистка/установка хедеров списка фильмов
        ReloadLVHeaders lvHeaderIndexPole


        'get language
        temp = VBGetPrivateProfileString("Language", "LCount", iniGlobalFileName)
        If IsNumeric(temp) Then LangCount = Int(Val(temp)) Else LangCount = 2

        temp = VBGetPrivateProfileString("Language", "LastLang", iniFileName)
        'If Len(temp) = 0 Then LastLanguage = "Русский" Else LastLanguage = temp
        If Len(temp) = 0 Then LastLanguage = vbNullString Else LastLanguage = temp

        'If LastLanguage <> vbNullString Then 'ну нет языка, почему бы список не создать
        ReDim ComboLang(LangCount)
        For i = 1 To LangCount
            temp = VBGetPrivateProfileString("Language", "L" & i, iniGlobalFileName)
            If Len(temp) <> 0 Then
                ComboLang(i) = temp
                If temp = LastLanguage Then
                    lngFileName = App.Path & "\" & VBGetPrivateProfileString("Language", "L" & i & "File", iniGlobalFileName)
                End If
            End If
        Next i
        'End If



        LangChange    'change Language полюбому


        'загружать ли последнюю БД                                                         BASES
        'temp = VBGetPrivateProfileString("GLOBAL", "LoadLastBD", iniFileName)
        'If Len(temp) <> 0 Then Opt_GoInCatalog = CBool(temp) Else Opt_GoInCatalog = True

        'скоко баз в списке
        temp = VBGetPrivateProfileString("GLOBAL", "BDCount", iniFileName)
        If IsNumeric(temp) Then tmpi = Int(Val(temp)) Else tmpi = 1
        Erase LstBases_List: ReDim LstBases_List(tmpi)    ':  'LstBases.Clear
        For i = 1 To tmpi    ' i c 1 for BDname & i
            temp = VBGetPrivateProfileString("GLOBAL", "BDname" & i, iniFileName)
            'check and add to list
            If Len(temp) <> 0 Then
                ret = FindFirstFile(temp, WFD)
                If ret > -1 Then
                    'ReadINI GetNameFromPathAndName(temp)
                    'дать имя базы с путем
                    If Len(GetPathFromPathAndName(temp)) = 0 Then temp = App.Path & "\" & temp
                    LstBases_List(i) = temp
                Else
                    LstBases_List(i) = temp
                    ToDebug "Err_NoBase: " & temp
                    temp = vbNullString

                End If
                FindClose ret
            Else    'нет записи в ини с таким индексом или пустая

            End If
        Next i
        LstBases_ListCount = UBound(LstBases_List)
        'Debug.Print "Баз всего:" & LstBases_ListCount


        'last base index
        temp = VBGetPrivateProfileString("GLOBAL", "LastBaseInd", iniFileName)
        If IsNumeric(temp) Then
            'LastBaseInd = temp
            CurrentBaseIndex = temp
        Else
            'LastBaseInd = 1
            CurrentBaseIndex = 1
        End If

        'window size
        temp = VBGetPrivateProfileString("GLOBAL", "WindowWidth", iniFileName)
        If IsNumeric(temp) Then
            tmpi = Int(Val(temp))
            If tmpi > 12075 Then
                MeWidth = tmpi
            Else
                MeWidth = Me.Width
            End If
        Else
            MeWidth = 12075
        End If
        temp = VBGetPrivateProfileString("GLOBAL", "WindowHeight", iniFileName)
        If IsNumeric(temp) Then
            tmpi = Int(Val(temp))
            If tmpi > 9105 Then
                MeHeight = tmpi
            Else
                MeHeight = Me.Height
            End If
        Else
            MeHeight = 9105
        End If
        'winState
        temp = VBGetPrivateProfileString("GLOBAL", "WindowState", iniFileName)
        If IsNumeric(temp) Then
            Select Case temp
            Case vbNormal, vbMaximized
                Me.WindowState = temp
                'If temp = vbMaximized Then Me.top = 0
            Case 1
                Me.WindowState = vbNormal
            End Select
        End If


        'InetProxy
        temp = VBGetPrivateProfileString("GLOBAL", "UseProxy", iniFileName)
        If IsNumeric(temp) Then
            If Val(temp) = 0 Or Val(temp) = 1 Or Val(temp) = 2 Then
                Opt_InetUseProxy = Val(temp)
            Else
                Opt_InetUseProxy = 1    'IE
            End If
        End If
        'server:port
        Opt_InetProxyServerPort = VBGetPrivateProfileString("GLOBAL", "ProxyServerPort", iniFileName)
        'user
        Opt_InetUserName = VBGetPrivateProfileString("GLOBAL", "ProxyUserName", iniFileName)
        'Pass
        'Opt_InetPassword = VBGetPrivateProfileString("GLOBAL", "ProxyPassword", iniFileName)
        'Secure
        temp = VBGetPrivateProfileString("GLOBAL", "ProxySecure", iniFileName)
        If Len(temp) <> 0 Then Opt_InetSecureFlag = CBool(temp) Else Opt_InetSecureFlag = False



        'Общие настройки
        temp = VBGetPrivateProfileString("GLOBAL", "GetMediaType", iniFileName)
        If Len(temp) <> 0 Then Opt_GetMediaType = CBool(temp) Else Opt_GetMediaType = True

        temp = VBGetPrivateProfileString("GLOBAL", "AviDirectShow", iniFileName)
        If Len(temp) <> 0 Then Opt_AviDirectShow = CBool(temp) Else Opt_AviDirectShow = False

        temp = VBGetPrivateProfileString("GLOBAL", "CancelAutoPlay", iniFileName)
        If Len(temp) <> 0 Then Opt_QueryCancelAutoPlay = CBool(temp) Else Opt_QueryCancelAutoPlay = True

        temp = VBGetPrivateProfileString("GLOBAL", "GetVolumeInfo", iniFileName)
        If Len(temp) <> 0 Then Opt_GetVolumeInfo = CBool(temp) Else Opt_GetVolumeInfo = True


        '-------это только читать------
        'брать картинку из инета через темп файл, иначе потоком. по умолч = false
        temp = VBGetPrivateProfileString("GLOBAL", "InetGetPicUseTempFile", iniFileName)
        If Len(temp) <> 0 Then Opt_InetGetPicUseTempFile = CBool(temp) Else Opt_InetGetPicUseTempFile = False



        '''''''''''''''''

    End If    'global
End With

'Debug.Print "прочитан " & iniFileName

Exit Sub
err:
If err.Number <> 0 Then
    Debug.Print "ReadINI.Err: " & err.Description
    ToDebug "Err_ReadINI: " & err.Description
End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'On Error Resume Next
Timer2.Enabled = False

'Debug.Print KeyCode, Shift

Select Case KeyCode

Case 81    'alt+q
    If addflag Or editFlag Then Exit Sub    'не делать в редакторе
    If Shift = 4 Then
        If FrameActer.Visible Then
            subActFiltCancel
        Else
            Call ComFilter_ShiftClick(2)
        End If
    End If

Case 45    'Insert auto add
    If Shift = 0 Then mnuAddNewAuto_Click    'там проверка
    ' frmAuto.Show 1, Me

Case 116    'F5 'группировка
    If addflag Or editFlag Then Exit Sub    'не делать в редакторе
    If FrameView.Visible Then mnuGroup_Click

Case 117    'F6 'показ скриншотов/обложки
    Opt_NoSlideShow = Not Opt_NoSlideShow
    If Opt_NoSlideShow Then
        picScrollBoxV.Visible = True: TextVAnnot.Visible = True
        If FrameView.Visible Then ComShowFa_Click
        'не Set Image0 = Nothing
    Else

        ComShowAn.Visible = False: ComShowFa.Visible = True
        Timer2_Timer    'обновить быстро
    End If
    Form_Resize

Case 119    'F8 карточка
    If FrameView.Visible Then mnuCard_Click

Case 114    'найти далее F3
    If addflag Or editFlag Then Exit Sub    'не делать в редакторе
    If FrameView.Visible Then ComNext_Click

Case 27    'esc
    If addflag Or editFlag Then Exit Sub    'не делать в редакторе
    If FormShowPicLoaded And FrameView.Visible Then Unload FormShowPic
    If LastVMI <> 1 Then VerticalMenu_MenuItemClick 1, 0

Case 123    'f12
    FormDebug.Show , FrmMain    ': FrmMain.SetFocus ' F12

Case 122    'f11
    If addflag Or editFlag Then Exit Sub    'не делать в редакторе
    FrmPeople.Show , FrmMain   ': FrmMain.SetFocus ' F11

Case 120    'F9 options
    If Not frmOptFlag Then
        VerticalMenu_MenuItemClick 6, 0
        'FrmOptions.SetFocus
    End If

Case 112    'F1
    If ChBTT.Value = 0 Then ChBTT.Value = 1 Else ChBTT.Value = 0
End Select
End Sub

Private Sub Form_Load()
'FrmMain.WBBR.Navigate2 "about:blank" 'in load
'Set WBBRDoc = FrmMain.WBBR.Document 'после навигейта
''Dim TempPicPath As String
''если нет, то создать
'TempPicPath = App.Path & "\templates\temp\"

'Dim WFD As WIN32_FIND_DATA
'пусть правит, а то бывает заскок вверх шапки ..?может уже нафик
KeepFormOnScreen FrmMain


'флажки и начальные переменные

'граббер
'    With mudtBitmapInfo
'        .Size = Len(mudtBitmapInfo)
'        .Planes = 1
'        .BitCount = 24
'    End With

AutoShots = True
GroupInd = -1
'PicSplitLVUHid.BorderStyle = 0:
PicSplitLVDHid.BorderStyle = 0
FirstLVFill = True    'для: куда ставить курсор в filllv

ActWWWsite = "http://images.google.com/images?q="    'сайт поиска актеров по кнопке

'цвет фона обложки белый
CoverVertBackColor = vbWhite
CoverHorBackColor = vbWhite

oldTabStripCoverInd = 1    'текущий таб обложки
oldTabLVInd = 0    'текущий таб LV

ChBTT.Move 2, 2

Randomize Timer

iniGlobalFileName = App.Path & "\global.ini"
CodecsFileName = App.Path & "\AVCodecList.svc"
userFile = App.Path & "\user.lng"

SendMessage PBar.hwnd, &H2001, 0, ByVal RGB(255, 255, 100)    'PBar Forecolor
SendMessage PBar.hwnd, &H409, 0, ByVal ForeColor    'RGB(50, 150, 0) 'PBar Backcolor

'файл для дебага
fnd = FreeFile
DebugFileFlagRW = True
On Error Resume Next
FileCopy App.Path & "\svcdebug.log", App.Path & "\svcdebug.old"
err.Clear
Open App.Path & "\svcdebug.log" For Output As fnd Len = 1
If err.Number <> 0 Then DebugFileFlagRW = False
err.Clear
On Error GoTo 0

ListView.MultiSelect = False
FirstActivateFlag = True

ToDebug Date & " SVC ver: " & App.Major & "." & App.Minor & "." & App.Revision
InitFlag = True
isWindowsNt = Environ$("OS") <> vbNullString
ToDebug "WinNT = " & isWindowsNt & " (" & winver & ")"
Me.ScaleMode = vbPixels
Ratio = 1

'узнать разделитель - до readINI
'SeparadorDecimal = GetDecimalSymbol
SeparadorDecimal = Format$(0, ".")
ToDebug "Separator - (" & SeparadorDecimal & ")"

ToDebug "ACP: " & GetACP
LCID = GetSystemDefaultLCID
ToDebug "LCID: " & LCID

'For Each Contrl In FrmMain.Controls
'If TypeOf Contrl Is TextBox Then 'Or (TypeOf Contrl Is ComboBox) Then
'Contrl.Font.Charset = 204
'End If
'Next
'FontVert.name = "Arial Cyr"
'FontHor.name = "Arial Cyr"



'верт слайдер
Set m_LV_Vert = New cMouseTrack
m_LV_Vert.AttachMouseTracking FrLV_Vert
Set m_TV_Vert = New cMouseTrack
m_TV_Vert.AttachMouseTracking FrTV_Vert
Set m_SS_Vert = New cMouseTrack
m_SS_Vert.AttachMouseTracking FrSplitD_Vert



''' открыть базу                                        актеров
OpenActDB

'иконки меню
If Not DebugMode Then
    SetMenuIcon    '(до readini, языка)
End If


ReadINI "global"

'загрузить редактор
Load frmEditor    'после координат из ини
'язык редактора
GetLangEditor

'сообщение и хук - no auto CD, ресайз (после readini)
If Not DebugMode Then
    If Opt_QueryCancelAutoPlay Then
        m_RegMsg = RegisterWindowMessage(RegMsg)
        ToDebug "AutoPlayMain = try2cancel"
    End If
    Call HookWindow(FrmMain.hwnd, Me)    'Friend Function WindowProc
    AttachMessage Me, FrmMain.hwnd, WM_GETMINMAXINFO
End If

'                                                       Прочитан язык
If Not FileExists(CodecsFileName) Then MsgBox msgsvc(49), vbExclamation

If LstBases_ListCount > 0 Then
    'Заполнить табы LV
    Call AddTabsLV
    'If LstBases_ListCount > 0 Then bdname = LstBases_List(LastBaseInd)
    If CurrentBaseIndex > LstBases_ListCount Then CurrentBaseIndex = LstBases_ListCount
    If LstBases_ListCount > 0 Then bdname = LstBases_List(CurrentBaseIndex)
    LstBases_ListIndex = CurrentBaseIndex - 1    'LastBaseInd
Else
    NoDBFlag = True
End If

'''''''''''
'Set up scroll bars просмотр
'ToDebug "Создание полос прокрутки для картинок"
Set m_cScroll = New cScrollBars
m_cScroll.Create picScrollBoxV.hwnd, False, True
'm_cScroll.style = efsEncarta

'm_cScroll.VBGColor = RGB(0, 0, 0)
'm_cScroll.VBGPalette = RGB(0, 0, 0)


PicFaceV.Move 0, 0

' Set up scroll bars актер
Set ma_cScroll = New cScrollBars
ma_cScroll.Create PicActFotoScroll.hwnd, False, True
'ma_cScroll.style = efsEncarta
PicActFoto.Move 0, 0



'LVSortColl = 0
'ListView.SortOrder = LVSortOrder 'lvwAscending =0 ' lvwDescending =1 'чтоб при нажатии было по алфавиту

frmMainFlag = True

'If Opt_GoInCatalog And Not NoDBFlag Then
If Not NoDBFlag Then
    'ToDebug "Переход к просмотру списка фильмов"
    VerticalMenu_MenuItemClick 1, 0    'goto View
Else
    InitFlag = True
    ToDebug "Не заданы базы..."
    VerticalMenu_MenuItemClick 5, 0    'goto Actors
End If

Image0.top = FrameImageHid.Height / 2 - Image0.Height / 2    'по центру

'Размеры
'окна скриншотов редактора
ScrShotEd_W = 3360
ScrShotEd_H = 2775
'окна видео
MovieEd_W = 4680
MovieEd_H = 3525

'BackGround
If FileExists(App.Path & "\Templates\background.jpg") Then
    lngBrush = CreatePatternBrush(LoadPicture(App.Path & "\Templates\background.jpg"))
End If

'NoPicture 'вставить из файла в конец ImageList
LastImageListInd = ImageList.ListImages.Count
If FileExists(App.Path & "\Templates\nopicture.jpg") Then
    ImageList.ListImages.Add LastImageListInd + 1, "NoPic", LoadPicture(App.Path & "\Templates\nopicture.jpg")
End If
LastImageListInd = ImageList.ListImages.Count

If CombFind.ListIndex > -1 Then CombFind.ListIndex = 0    'текст в списке колонок поиска

''доступны ли dll
'ToDebug "wiaaut.dll - " & IsDLLAvailable(App.Path & "\wiaaut.dll")
''Debug.Print IsDLLAvailable("mediainfo.dll")

Unload frmSplash: Set frmSplash = Nothing

'Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ExitSVC = True
On Error Resume Next

'запись в ини
If ListView.ListItems.Count > 0 Then
    LastInd = FrmMain.ListView.SelectedItem.SubItems(lvIndexPole)
    WriteKey "LIST", "LastItem", str$(LastInd), iniFileName
End If
WriteKey "GLOBAL", "WindowState", Me.WindowState, iniGlobalFileName
If Me.WindowState = vbNormal Then
    WriteKey "GLOBAL", "WindowWidth", Me.Width, iniGlobalFileName
    WriteKey "GLOBAL", "WindowHeight", Me.Height, iniGlobalFileName
End If

If TabLVHid.Tabs.Count > 0 Then
    WriteKey "GLOBAL", "LastBaseInd", TabLVHid.SelectedItem.Index, iniGlobalFileName
End If

If LVActer.ListItems.Count > 0 Then
    WriteKey "LIST", "LastItemAct", CStr(Val(CurActKey)), iniFileName
End If

If Opt_AutoSaveOpt Then SaveInterface

'сохранить историю
Call SaveHistory(GetNameFromPathAndName(bdname))

'дестрой
ModLVSubClass.UnAttach FrameView.hwnd
'cManager.Goodbye
DetachMessage Me, Me.hwnd, WM_GETMINMAXINFO
'меню-иконки
If Not (cIM Is Nothing) Then cIM.Detach
Call TT.DestroyToolTip


End Sub

Public Sub Form_Resize()
Dim X As Long, Y As Long, w As Long
Dim temp As Single

'Наблюдения
'VerticalMenu.Width - в пикселях
FrmMain.ScaleMode = vbPixels    'точно!

Dim LVHeight As Long
'Debug.Print Time
'содержимое окна при таскании
'pTurnOffFullDrag
'''''''''''''''''''''''''
If NoResizePlease Then Exit Sub

On Error Resume Next
If kDPI <= 0 Then kDPI = 1
If txtEdit.Visible Then txtEdit_LostFocus

''Background
'If lngBrush <> 0 Then
'GetClientRect hwnd, rctMain
'FillRect hdc, rctMain, lngBrush

''GetClientRect FrameView.hWnd, rctMain
''FillRect FrameView.hDC, rctMain, lngBrush
'End If

'DoEvents:
'TextVAnnot.Visible = False
'txt = TextVAnnot.Text: TextVAnnot.Text = vbNullString
'LockWindowUpdate TextVAnnot.hwnd


'FrmMainState = Me.WindowState - путаница c Bin
If Not NoDBFlag Then
    If Me.WindowState = 1 Then
        Timer2.Enabled = False
        Exit Sub    'спрятан
    Else
        Timer2.Enabled = True
    End If
End If    'nodb

MainWidth = Me.Width - VerticalMenu.Width - 1440    '- 1500
MainWidthPix = Me.ScaleWidth - VerticalMenu.Width / Screen.TwipsPerPixelX - 78 * kDPI
MainHeightPix = Me.ScaleHeight

MainHeight = Me.Height
If MainHeight < 12000 Then
    LVHeight = MainHeight / 1.95
Else
    LVHeight = MainHeight / 1.75    '1.75 '8
End If
'LVHeight = MainHeight / 1.95

FrameView.Move VerticalMenu.Width + 2, FrameView.top, MainWidthPix, MainHeightPix
VerticalMenu.Move 0, 0, VerticalMenu.Width, MainHeightPix    ' + 1


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Фильмы
If FrameView.Visible Or FirstActivateFlag Then
    'Call SendMessage(ListView.hwnd, WM_SETREDRAW, 0, 0)
    If Not FirstLVFill Then If ComShowFa.Visible Then Call SendMessage(TextVAnnot.hwnd, WM_SETREDRAW, 0, 0)

    'растяжка Листвью
    'FrameView.Visible = False
    ': ListView.Visible = False обратно трудно

    If Opt_UCLV_Vis Then
        If Opt_Group_Vis Then

            tvGroup.Move 60, tvGroup.top, TVWidth, LVHeight
            temp = Abs(LVWidth / 100 * (MainWidth - tvGroup.Width))
            If LVWidth >= 99 Then temp = (MainWidth - tvGroup.Width) * 0.4
            If LVWidth <= 1 Then temp = (MainWidth - tvGroup.Width) * 0.4

            ListView.Move TVWidth + 120, ListView.top, temp, LVHeight
            FrTV_Vert.Move tvGroup.left + TVWidth, tvGroup.top, FrTV_Vert.Width, LVHeight
            UCLV.Move TVWidth + temp + 180, ListView.top, MainWidth - temp - TVWidth - 90, LVHeight
            tvGroup.Visible = True: FrTV_Vert.Visible = True
        Else
            temp = Abs(LVWidth / 100 * MainWidth)
            If LVWidth >= 99 Then temp = MainWidth * 0.4
            If LVWidth <= 1 Then temp = MainWidth * 0.4

            ListView.Move 60, ListView.top, temp, LVHeight
            UCLV.Move ListView.Width + ListView.left + 60, ListView.top, MainWidth - ListView.Width - 30, LVHeight
            tvGroup.Visible = False: FrTV_Vert.Visible = False
        End If

        FrLV_Vert.Move ListView.Width + ListView.left, ListView.top, FrLV_Vert.Width, LVHeight
        UCLV.Visible = True: FrLV_Vert.Visible = True
        UCLV.Refresh

    Else    'нет карточки

        If Opt_Group_Vis Then
            tvGroup.Move 60, tvGroup.top, TVWidth, LVHeight
            ListView.Move tvGroup.Width + tvGroup.left + 60, ListView.top, MainWidth - tvGroup.Width - 30, LVHeight
            FrTV_Vert.Move tvGroup.Width + tvGroup.left, ListView.top, FrTV_Vert.Width, LVHeight
            tvGroup.Visible = True: FrTV_Vert.Visible = True
        Else
            tvGroup.Visible = False: FrLV_Vert.Visible = False
            ListView.Move 60, ListView.top, MainWidth + 30, LVHeight
        End If

        UCLV.Visible = False: FrLV_Vert.Visible = False

    End If    'Opt_UCLV_Vis

    TabLVHid.Move TabLVHid.left, TabLVHid.top, MainWidth - comHistory.Width
    comHistory.left = TabLVHid.left + TabLVHid.Width + 30






    PicSplitLVDHid.Visible = False
    PicSplitLVDHid.Move PicSplitLVDHid.left, ListView.top + ListView.Height + 50, MainWidth + 30, ScaleY(FrameView.Height, vbPixels, vbTwips) - LVHeight - 700    '750

    TextItemHid.Move TextItemHid.left, TextItemHid.top, PicSplitLVDHid.Width
    PBar.Move TextItemHid.left + 50, TextItemHid.top + 60, TextItemHid.Width - 130, TextItemHid.Height - 130    'mline TextItemHid.Width - 350
    '
    SSCoverAnnotW = PicSplitLVDHid.Width - FrFindViewHid.Width - 130
    SSCoverAnnotT = TextItemHid.Height + 60
    SSCoverAnnotH = PicSplitLVDHid.Height - SSCoverAnnotT - 30

    If SplitLVD > 80 Then
        temp = Abs(80 / 100 * SSCoverAnnotW)
    Else
        temp = Abs(SplitLVD / 100 * SSCoverAnnotW)
    End If

    If Opt_NoSlideShow Then    'без скриншотов
        FrameImageHid.Visible = False
        picScrollBoxV.Move 0, SSCoverAnnotT, temp, SSCoverAnnotH   '- 10 по идее
        TextVAnnot.Move picScrollBoxV.Width + 60, SSCoverAnnotT, SSCoverAnnotW - picScrollBoxV.Width, SSCoverAnnotH
        'двинуть слайдер
        FrSplitD_Vert.Move picScrollBoxV.Width, SSCoverAnnotT, FrSplitD_Vert.Width, SSCoverAnnotH

    Else    'показ скриншотов
        'temp = 4000
        FrameImageHid.Move 0, SSCoverAnnotT, temp, SSCoverAnnotH
        Image0.Width = FrameImageHid.Width
        FrameImageHid.Visible = True

        picScrollBoxV.Visible = False
        picScrollBoxV.Move FrameImageHid.Width + 60, SSCoverAnnotT, SSCoverAnnotW - FrameImageHid.Width, SSCoverAnnotH
        picScrollBoxV.Visible = True

        TextVAnnot.Move picScrollBoxV.left, SSCoverAnnotT, picScrollBoxV.Width, picScrollBoxV.Height
        'двинуть слайдер
        FrSplitD_Vert.Move FrameImageHid.Width, SSCoverAnnotT, FrSplitD_Vert.Width, FrameImageHid.Height
        'If ComShowFa.Visible Then picScrollBoxV.Visible = False

    End If
    PicSplitLVDHid.Visible = True

    LstFiles.Move TextVAnnot.left + 500, TextVAnnot.top + 200, TextVAnnot.Width - 1000, TextVAnnot.Height - 400

    'коробочка с поиском
    FrFindViewHid.left = SSCoverAnnotW + 140

    'высота комбиков
    X = ScaleX(CombFind.left, vbTwips, vbPixels)
    Y = ScaleY(CombFind.top, vbTwips, vbPixels)
    w = ScaleY(CombFind.Width, vbTwips, vbPixels)
    SetWindowPos CombFind.hwnd, 0, X, Y, w, 500, 0
    CombFind.SelLength = 0

    'Call SendMessage(ListView.hwnd, WM_SETREDRAW, 1, 0)
    If ComShowFa.Visible Then Call SendMessage(TextVAnnot.hwnd, WM_SETREDRAW, 1, 0)  'или SWP_NOZORDER ?

    If FirstLVFill Then
        'nopic
        Image0.Move 0, 0, FrameImageHid.Width, FrameImageHid.Height
        If ImageList.ListImages.Count >= 3 Then
            Image0.PaintPicture ImageList.ListImages(LastImageListInd).Picture, 0, 0, Image0.Width, Image0.Height
        End If
    End If

End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''обложка
'If fr_cover Then
If FrameCover.Visible Then

    FrameCover.Move FrameView.left, FrameView.top, FrameView.Width, FrameView.Height
    PicPrintScroll.Move PicPrintScroll.left, PicPrintScroll.top, ScaleX(FrameCover.Width - 7, vbPixels, vbTwips), ScaleY(FrameCover.Height, vbPixels, vbTwips) - FrPrintBotHid.Height - 200
    'PicPrintScroll.Move PicPrintScroll.Left, PicPrintScroll.Top, ScaleX(FrameCover.Width - 7, vbPixels, vbTwips), ScaleY(FrameCover.Height, vbPixels, vbTwips) - FrPrintBotHid.Height - 200

    'PicPrintScroll.Move PicPrintScroll.Left, PicPrintScroll.Top, FrameCover.Width * Screen.TwipsPerPixelX - 25 * Screen.TwipsPerPixelX, ScaleY(FrameCover.Height, vbPixels, vbTwips) - FrPrintBotHid.Height - 200
    'PicPrintScroll.Move PicPrintScroll.Left, PicPrintScroll.Top, FrameCover.Width - 100, ScaleY(FrameCover.Height, vbPixels, vbTwips) - FrPrintBotHid.Height - 200

    FrPrintBotHid.Move FrPrintBotHid.left, ScaleY(FrameCover.Height, vbPixels, vbTwips) - FrPrintBotHid.Height - 100, FrameCover.Width * Screen.TwipsPerPixelX - 300 * kDPI
    CmdPrint.Move FrPrintBotHid.left + FrPrintBotHid.Width - CmdPrint.Width - 300 * kDPI
    TabStripCover.Move TabStripCover.left, TabStripCover.top, FrPrintBotHid.Width - 100

    ' 'высота комбиков
    'x = ScaleX(CombFind.Left, vbTwips, vbPixels)
    'y = ScaleY(CombFind.Top, vbTwips, vbPixels)
    'w = ScaleY(CombFind.Width, vbTwips, vbPixels)
    'SetWindowPos CombFind.hwnd, 0, x, y, w, 500, 0

End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Актеры
'If fr_acter Then
If FrameActer.Visible Or fr_acter Then

    Call SendMessage(LVActer.hwnd, WM_SETREDRAW, 0, 0)
    Call SendMessage(TextActBio.hwnd, WM_SETREDRAW, 0, 0)



    FrameActer.Move FrameView.left, FrameView.top, FrameView.Width, FrameView.Height
    FrActLeft.Move FrActLeft.left, FrActLeft.top, ScaleX(FrameActer.Width, vbPixels, vbTwips) / 2.2 - 880, ScaleY(FrameActer.Height, vbPixels, vbTwips) - 350

    LVActer.Height = FrActLeft.Height - 2260
    FrActSelect.top = FrActLeft.Height - 2200

    LVActer.Width = FrActLeft.Width - 50
    LVActer.ColumnHeaders(1).Width = LVActer.Width - 400    'и в филе актеров - дрожит даже для myframe

    FrActButtons.Move ScaleX(FrameActer.Width, vbPixels, vbTwips) - 1930    '1900

    FrActBio.Move FrActLeft.Width + 90, ScaleY(FrameView.Height, vbPixels, vbTwips) / 2, ScaleX(FrameActer.Width, vbPixels, vbTwips) - FrActLeft.Width - 150, ScaleY(FrameActer.Height, vbPixels, vbTwips) / 2 - 100

    FrActSelect.Width = FrActLeft.Width - 40
    TextSearchLVActTypeHid.Width = FrActSelect.Width - comActSearchInBIO.Width - 60
    comActSearchInBIO.left = TextSearchLVActTypeHid.Width + 30
    ListBActHid.Width = FrActSelect.Width - LActMarkHelp.Width - 300
    ComSelMovIcon.Width = FrActSelect.Width

    TextActName.Width = FrActBio.Width - ComRHid(0).Width - 60
    ComRHid(0).left = TextActName.left + TextActName.Width + 20

    'LockWindowUpdate TextActBio.hwnd
    'TextActBio.Move TextActBio.Left, TextActBio.Top, FrActBio.Width - 40, FrActBio.Height - 900
    'LockWindowUpdate 0
    'Call SendMessage(TextActBio.hwnd, WM_SETREDRAW, False, ByVal 0&)
    TextActBio.Width = FrActBio.Width - 40
    TextActBio.Height = FrActBio.Height - 900
    'Call SendMessage(TextActBio.hwnd, WM_SETREDRAW, True, ByVal 0&)

    PicActFotoScroll.Move FrActLeft.Width + 90, PicActFotoScroll.top, ScaleX(FrameActer.Width, vbPixels, vbTwips) - FrActLeft.Width - FrActButtons.Width - 300, FrActBio.top - 300


    ComSelMovIcon.pInitialize

    Call SendMessage(LVActer.hwnd, WM_SETREDRAW, 1, 0)
    Call SendMessage(TextActBio.hwnd, WM_SETREDRAW, 1, 0)
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Редактор

'If FrameAddEdit.Visible Then
'FrameAddEdit.Move (FrameView.Width - FrameAddEdit.Width) / 2 + VerticalMenu.Width + 1, (FrameView.Height - FrameAddEdit.Height) / 2
'End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
ExitSVC = True
On Error Resume Next

'ModLVSubClass.UnAttach FrameView.hWnd

Timer2.Enabled = False

If DebugFileFlagRW Then Close #fnd    'debug

Unload frmEditor
Unload FormDebug    ': Set FormDebug = Nothing
Unload FormShowPic    ': Set FormShowPic = Nothing
Unload frmAuto
Unload FrmBin    ': Set FrmBin = Nothing
Unload FrmFilter    ': Set FrmFilter = Nothing
Unload FrmOptions
Unload FrmPeople    ': Set FrmPeople = Nothing
Unload FrmStat
Unload frmSR

'Set m_cAVI = Nothing
'Set m_cDib = Nothing

'Инет
Set SC = Nothing

'Set mobjManager = Nothing

rs.Close
ars.Close
DB.Close
ADB.Close
Set rs = Nothing
Set DB = Nothing
Set ars = Nothing
Set ADB = Nothing

Call UnhookWindow(Me.hwnd)

ListView.ListItems.Clear    'по идее ускорит выгрузку
LVActer.ListItems.Clear

'close mutex from main.bas
EndApp

End Sub


Private Sub Image0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If rs Is Nothing Then Exit Sub
If rs.RecordCount < 1 Then Exit Sub

Select Case Button
Case 2    'правая
    'Timer2.Enabled = False
    Timer2_Timer
    'Exit Sub

Case 4    'средняя
    Timer2.Enabled = False

Case Else
    PicManualFlag = True
    If ListView.ListItems.Count > 0 Then
        ScrShotClick
    End If
    PicManualFlag = False

End Select
End Sub


Private Sub ListBActHid_Click()

If NoDBFlag Then Exit Sub

If ListBActHid.SelCount = 0 Then
    ComSelMovIcon.Enabled = False
Else
    ComSelMovIcon.Enabled = True
End If

End Sub

Public Sub LVCLICK()
Dim Itm As ListItem
Dim tmp As Long

'Debug.Print Time & "lvclick"
'mzt Dim netkart As Boolean

LVSelectChanged = True    'меняется в frmStat (из айтемклика ctrl+ клик не срабатывает, когда снимаем выделение)

SelCount = 0: CheckCount = 0
ReDim CheckRows(0): ReDim SelRows(0)
ReDim CheckRowsKey(0): ReDim SelRowsKey(0)

If (rs Is Nothing) Then
    'nopic
    Image0.Move 0, 0, FrameImageHid.Width, FrameImageHid.Height
    Image0.PaintPicture ImageList.ListImages(LastImageListInd).Picture, 0, 0, Image0.Width, Image0.Height
    NoListClear
    Exit Sub
End If
If (rs.RecordCount < 1) Or (ListView.ListItems.Count < 1) Then
    'nopic
    Image0.Move 0, 0, FrameImageHid.Width, FrameImageHid.Height
    Image0.PaintPicture ImageList.ListImages(LastImageListInd).Picture, 0, 0, Image0.Width, Image0.Height
    NoListClear
    If frmEditorFlag Then
        'убрать редактор
        Unload frmEditor
    End If
    Exit Sub
End If

Timer2.Enabled = False
'ListView.Refresh

ComNext.Enabled = True

''при каждом клике заполнять таблицы чеков и селектов
'For Each Itm In ListView.ListItems
'    With Itm
'        If .Selected Then
'            SelCount = SelCount + 1
'            ReDim Preserve SelRows(UBound(SelRows) + 1)
'            SelRows(SelCount) = .ListSubItems(lvIndexPole)
'            ReDim Preserve SelRowsKey(UBound(SelRowsKey) + 1)
'            SelRowsKey(SelCount) = Itm.Key
'        End If
'        If .Checked Then
'            CheckCount = CheckCount + 1
'            ReDim Preserve CheckRows(UBound(CheckRows) + 1)
'            CheckRows(CheckCount) = .ListSubItems(lvIndexPole)
'            ReDim Preserve CheckRowsKey(UBound(CheckRowsKey) + 1)
'            CheckRowsKey(CheckCount) = Itm.Key
'        End If
'    End With
'Next

tmp = SendMessage(ListView.hwnd, LVM_GETSELECTEDCOUNT, 0&, ByVal 0&)
ReDim SelRows(tmp)
ReDim SelRowsKey(tmp)
For Each Itm In ListView.ListItems
    With Itm
        If .Selected Then
            SelCount = SelCount + 1
            SelRows(SelCount) = .ListSubItems(lvIndexPole)
            SelRowsKey(SelCount) = .Key
        End If
        If .Checked Then
            CheckCount = CheckCount + 1
            ReDim Preserve CheckRows(UBound(CheckRows) + 1)
            CheckRows(CheckCount) = .ListSubItems(lvIndexPole)
            ReDim Preserve CheckRowsKey(UBound(CheckRowsKey) + 1)
            CheckRowsKey(CheckCount) = .Key
        End If
    End With
Next

'SelCount = SendMessage(ListView.hWnd, LVM_GETSELECTEDCOUNT, 0&, ByVal 0&)
If SelCount > 1 Then MultiSel = True Else MultiSel = False

PicManualFlag = False


'If Not MultiSel Then

Select Case rs.Type
Case dbOpenTable
    rs.Seek "=", Val(ListView.SelectedItem.Key)    'seek не пашет после sql
    If rs.NoMatch Then Exit Sub
    'Case dbOpenDynamic
    'Case dbOpenSnapshot
Case dbOpenDynaset, dbOpenSnapshot
    rs.FindFirst "[Key] = " & Val(ListView.SelectedItem.Key)
    If rs.NoMatch Then Exit Sub
    'Case dbOpenForwardOnly
Case Else
    'Не по ключу, а по порядку
    Debug.Print "b_типтаб"
    Exit Sub
End Select

CurSearch = ListView.SelectedItem.Index
LastInd = ListView.SelectedItem.Index - 1
LastKey = Val(ListView.SelectedItem.Key)

'ListView.SelectedItem.Selected = True     'надо, не помню зачем, отключим пока

On Error Resume Next

If Opt_LoadOnlyTitles Then
    'If Not lvItemLoaded(ListView.SelectedItem.Index) Then
    'ну и пусть грузит всегда - надо грузить для обновления при замене

    FillLvSubs ListView.SelectedItem.Index         'подгрузка LV до аддона
    'там    lvItemLoaded(ListView.SelectedItem.Index) = True
    'End If
End If

SlideShowFlag = 0
timerflag = False

'If Not MultiSel Then

Set PicTempHid(0) = Nothing: Set PicTempHid(1) = Nothing
'скриншоты
NoPic1Flag = False: NoPic2Flag = False: NoPic3Flag = False

SShotsCount = 3
If rs.Fields(dbSnapShot1Ind).FieldSize = 0 Then NoPic1Flag = True: SShotsCount = SShotsCount - 1
If rs.Fields(dbSnapShot2Ind).FieldSize = 0 Then NoPic2Flag = True: SShotsCount = SShotsCount - 1
If rs.Fields(dbSnapShot3Ind).FieldSize = 0 Then NoPic3Flag = True: SShotsCount = SShotsCount - 1

If FrameView.Visible Then
    If Not Opt_NoSlideShow Then
        If Not NoPic3Flag Then SlideShowFlag = 2
        If Not NoPic2Flag Then SlideShowFlag = 1
        If Not NoPic1Flag Then SlideShowFlag = 0

        If Not MultiSel Then Timer2_Timer    'показать новый скриншот

        'SlideShowLastFlag = 0
        'GetPic Image0, 1, "SnapShot1"
        'GetPic PicTempHid(0), 1, "SnapShot1"

        If FormShowPicLoaded Then
            If NoPic1Flag And NoPic2Flag And NoPic3Flag Then
                If ViewScrShotFlag And Not IsCoverShowFlag Then FormShowPic.Hide    'спрятать окно
            End If
        End If

        If FormShowPicLoaded And Not IsCoverShowFlag Then
            SlideShowLastGoodPic = 0    'устарело
            SlideShowFlag = 1
            'кликнуть на скриншот, там показ в 1:1
            ShowPicFocus = False    'отнять фокус у 1:1

            On Error GoTo 0
            ScrShotClick
            ViewScrShotFlag = True: IsCoverShowFlag = False
        End If
    End If
End If


'обложка
Set PicFaceV = Nothing

If picScrollBoxV.Visible Or Opt_NoSlideShow Then
    GetPic PicFaceV, 1, "FrontFace"
    picScrollBoxV_Resize
    If m_cScroll.Visible(efsHorizontal) Then PicFaceV.left = -Screen.TwipsPerPixelX * m_cScroll.Value(efsHorizontal)
    If m_cScroll.Visible(efsVertical) Then PicFaceV.top = -Screen.TwipsPerPixelY * m_cScroll.Value(efsVertical)
End If
'ListView.Picture = PicFaceV.Picture

If FormShowPicLoaded And Not ViewScrShotFlag Then
    'Показать текущую обложку в 1:1 окне
    If GetPic(PicTempHid(1), 1, "FrontFace") Then
        IsCoverShowFlag = True
        ViewScrShotFlag = False
        ShowInShowPic 1, FrmMain
    Else
        FormShowPic.Hide        'спрятать окно
        'Me.SetFocus 'в той форме
    End If
End If
If NoPicFrontFaceFlag Then
    PicFaceV.Width = 0: PicFaceV.Height = 0: picScrollBoxV_Resize
    'If FormShowPicFlag And Not ViewScrShotFlag Then FormShowPic.Hide 'спрятать окно
    If FormShowPicLoaded And Not ViewScrShotFlag Then FormShowPic.Hide        'спрятать окно
End If

'End If 'If Not MultiSel Then

'ОПИсалово
'TextVAnnot.Text = CheckNoNull("Annotation")
Call SendMessage(TextVAnnot.hwnd, WM_SETREDRAW, False, ByVal 0&)
If IsNull(rs.Fields("Annotation")) Then TextVAnnot.Text = vbNullString Else TextVAnnot.Text = rs.Fields("Annotation")
'примечания
If Opt_PutOtherInAnnot Then
    If Len(TextVAnnot.Text) = 0 Then
        If IsNull(rs.Fields("Other")) Then TextVAnnot.Text = vbNullString Else TextVAnnot.Text = rs.Fields("Other")
    Else
        'дописать
        If Not IsNull(rs.Fields("Other")) Then TextVAnnot.Text = TextVAnnot.Text & vbCrLf & vbCrLf & rs.Fields("Other")
    End If
End If
If ComShowFa.Visible Or Opt_NoSlideShow Then Call SendMessage(TextVAnnot.hwnd, WM_SETREDRAW, True, ByVal 0&)


'загрузка аддона (после FillLvSubs)
DoEvents
If Opt_UCLV_Vis Then FillLVAdd


'запомнить в историю
If LVManualClickFlag Then
    StoreHistory ListView.SelectedItem.Text, ListView.SelectedItem.Key    'lastkey
End If
LVManualClickFlag = False

'ShowRowInWB
End Sub

'Private Function wbbrDoc_oncontextmenu() As Boolean
''правый клик на WB отключен
' WBBRDoc.parentWindow.event.returnValue = False
' PopupMenu Me.popFaceHid
' 'Debug.Print WBBRDoc.embeds(
'
'End Function























Private Sub Get_LV_Selections()
'поиметь на текущий момент состояние селекта
Dim Itm As ListItem
Dim SC As Long

SC = SendMessage(ListView.hwnd, LVM_GETSELECTEDCOUNT, 0&, ByVal 0&)
ReDim SelRows(SC): ReDim SelRowsKey(SC)
SelCount = 0
For Each Itm In ListView.ListItems
    With Itm
        If .Selected Then
            SelCount = SelCount + 1
            SelRows(SelCount) = .ListSubItems(lvIndexPole)
            SelRowsKey(SelCount) = .Key
        End If
    End With
Next
If SelCount > 1 Then MultiSel = True Else MultiSel = False
End Sub
Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Dim ind As Long    'ColumnHeader.Index - 1

If ListView.ListItems.Count < 1 Then Exit Sub

On Error GoTo errh

'sort
LVSortColl = ColumnHeader.Index
ComNext.Enabled = True

'LockWindowUpdate ListView.hwnd

'Менять порядок сортировки
ListView.SortOrder = (ListView.SortOrder + 1) Mod 2
LVSortOrder = ListView.SortOrder

ToDebug "SortHeader/Order: " & ColumnHeader.Text & "/" & LVSortOrder

'Поменять <, >
ColumnHeader.Text = ChangeLVHSortMark(ColumnHeader.Text)

'ind = ColumnHeader.Index - 1
LVSOrt ColumnHeader.Index

'Select Case ind
'Case dbTimeInd  'time
'    SortByDates ind
'Case dbYearInd, dbFileLenInd, dbCDNInd, dbRatingInd, lvIndexPole  'lvHeaderIndexPole 'год 4, размер 12, cdn 13, рейтинг, индекс
'    SortByNumber ind, ListView
'Case Else
'    If ind = dbLabelInd Then
'        Select Case rs(dbLabelInd).Type
'        Case 2, 3, 4, 6, 7 'метка -числовые
'            SortByNumber ind, ListView
'        Case Else
'            'как текст
'            SortByString ind
'        End Select
'    Else
'        'как текст
'        SortByString ind
'    End If
'End Select

'LockWindowUpdate 0

CurSearch = ListView.SelectedItem.Index
If MultiSel Then
    ListView.ListItems(CurSearch).EnsureVisible
Else
    LV_EnsureVisible ListView, CurSearch
End If

Exit Sub

errh:
ToDebug "Err_SortLV: " & err.Description

End Sub

Private Sub ListView_DblClick()
'редактировать
If GetKeyState(vbKeyMenu) < 0 Then
    'alt
    VerticalMenu_MenuItemClick 2, 0
    Exit Sub
End If

If Opt_LVEDIT Then
    ListViewEdit ListView
Else
    VerticalMenu_MenuItemClick 2, 0
End If
End Sub

Private Sub ListView_ItemCheck(ByVal Item As ListItem)
If rs Is Nothing Then Exit Sub

'rs.MoveFirst:rs.Move Item.ListSubItems(lvIndexPole)
'RSMOVE Item.ListSubItems(lvIndexPole), "ListView_ItemCheck"
RSGoto Item.Key

If BaseReadOnly Then Exit Sub

rs.Edit
If Item.Checked Then
    rs.Fields(dbCheckedInd) = "1"
Else
    rs.Fields(dbCheckedInd) = vbNullString
End If

rs.Update
End Sub




Private Sub ListView_KeyPress(KeyAscii As Integer)
'Debug.Print KeyAscii
'Dim Itm As ListItem
Dim ClickFlag As Boolean    'иногда не кликать LV (глюк с ctrl+N)

ClickFlag = False

'Debug.Print "Кнопки: " & KeyAscii  '& Chr$(KeyAscii)
Select Case KeyAscii

Case 26    'ctrl+Z пометить выделенное
    Get_LV_Selections
    Call mChSel_Click

Case 22    'ctrl+v
    Call mSelCh_Click    'выделить помеченное

Case 23    'ctrl+w 'сделать дубликат текущей записи
    Get_LV_Selections
    mnuCopyRow_Click    'Call MakeDupCurrent

    'в форме Case 6 'crtl+F
    '    mSR_Click 'поиск

Case 18    'ctrl R
    InetGetPics True    'помеченные

Case 20    'ctrl T
    Get_LV_Selections
    mGetCoverSel_Click    'InetGetPics False 'выделенные mGetCoverSel_Click

Case 43    ' + объединение
    JoinMovies (True)    'for checked



Case 1    'ctrl+A
    Call mnuSelectAllLV_Click
    '    For Each Itm In ListView.ListItems
    '        Itm.Selected = True
    '    Next
    '    ClickFlag = True

Case 2    'ctrl+B
    Call mnuCheckAll_Click

Case 4    'ctrl+D
    Call mnuCheckNone_Click

Case 13    'ctrl+M Enter
    Call mnuPlayM_Click

    'в форме Case 5 '^E
    '    Call mnuEditMov_Click

    'в форме Case 14 '^N
    'Call mnuAddNewMov_Click

Case 19    'S
    Call mnuSortChecked_Click

Case 9    'ctrl+i
    Call mInvCh_Click

Case 10    'ctrl+j
    Call mInvSel_Click

Case 21    'U
    Get_LV_Selections
    Call mnuDolgSel_Click

Case 12    'L
    Get_LV_Selections
    Call mnuLabelSel_Click

Case 3    'C
    Get_LV_Selections
    Call mnuCopyLV_Click

Case 8    'H и ?BackSpace
    Get_LV_Selections
    Call mnuHTML_Click

Case 25    ' Y
    Call mnuDolgCheck_Click

Case 11    ' K
    Call mnuLabelCheck_Click

Case 24    'X
    Call mnuExportCheckClip_Click

Case 7    'G
    Call mnuExportCheckHTML_Click

    'уже нет Case 6 '^F
    '    ComFind_Click

Case 15    '^O Excel
    Call mCh2Excel_Click

Case 16    '^P
    Get_LV_Selections
    Call mSel2Excel_Click

End Select

If ClickFlag Then LVCLICK
End Sub

Private Sub ListView_KeyUp(KeyCode As Integer, Shift As Integer)
If rs Is Nothing Then Exit Sub
If rs.RecordCount = 0 Then Exit Sub

'при нажатии Горячих комбинаций - после отпускания всех
'Debug.Print KeyCode

'Debug.Print CurSearch, ListView.SelectedItem.index
Timer2.Enabled = False
'CurSearch = ListView.SelectedItem.index
'Set ListView.SelectedItem = ListView.ListItems(LastInd)
'Debug.Print CurSearch, ListView.SelectedItem.index


If Shift = 0 Then
    'кликаем по стрелкам, начальным буквам
    Select Case KeyCode
    Case 18, 17    'чистые alt и шифт
        Exit Sub
    End Select

    LVCLICK

    DoEvents
    'загрузить в редактор
    If frmEditorFlag Then
        Screen.MousePointer = vbHourglass
        addflag = False: editFlag = True
        GetEditPix
        frmEditor.ImgPrCov.Picture = frmEditor.PicFrontFace.Picture
        Mark2SaveFlag = False
        GetFields
        Mark2SaveFlag = True
        frmEditor.ComSaveRec.BackColor = &HC0E0C0
        ToDebug "LV_Ed_Key: " & rs("key")
        Screen.MousePointer = vbNormal
    End If
End If

End Sub

Private Sub ListView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer2.Enabled = False
If rs Is Nothing Then Exit Sub

Select Case Button
Case vbLeftButton
    If Not MultiSel Then

        'положить кликнутую ячейку в текстбокс
        TextItemHid.Text = vbNullString
        lvwMsg.X = X / Screen.TwipsPerPixelX
        lvwMsg.Y = Y / Screen.TwipsPerPixelY
        DoEvents
        SendMessage ListView.hwnd, 4153, 0, lvwMsg
        If lvwMsg.Flgs And 14 Then
            If lvwMsg.SubItm Then
                TextItemHid.Text = ListView.ListItems(lvwMsg.Itm + 1).SubItems(lvwMsg.SubItm)
            Else
                TextItemHid.Text = ListView.ListItems(lvwMsg.Itm + 1).Text
            End If
        End If
    End If    'not multi

Case vbRightButton
    'LockWindowUpdate ListView.hwnd
    'FrameView.Enabled = False: ListView.Enabled = False: ListView.Enabled = True: FrameView.Enabled = True
    'ListView.Enabled = True
    'FrameView.Visible = True
    'LockWindowUpdate 0
    'ReleaseCapture
    'DoEvents

    'финт для снятия клика с листвью
    'WM_RBUTTONUP = &H205

    Call SendMessage(ListView.hwnd, &H205, 0, ByVal 0&)    'это кликает! lvclick. До опроса SelCount


    'LVPopupFlag = True
    'ПОПАП
    'не дает пометить все    If CheckCount > 0 Then mnuLVChecked.Enabled = True Else mnuLVChecked.Enabled = False
    'SelCount заполнен lvclick-ом
    mnuLVChecked.Caption = mnuLVCheckedCaption & vbTab & "(" & CheckCount & ")"
    mnuLVSelected.Caption = mnuLVSelectedCaption & vbTab & "(" & SelCount & ")"

    ''сумма размеров файлов (сделать по запросу?)
    'mSumCh.Caption = mSumChCaption & vbTab & FileSizeSumCh & " Mb"
    'mSumSel.Caption = mSumSelCaption & vbTab & FileSizeSumSel & " Mb"

    If ListView.ListItems.Count > 0 Then
        mnuPlayM.Enabled = True
        If SelCount > 0 Then If Len(CheckNoNull("MovieURL")) <> 0 Then mGotoURL.Enabled = True Else mGotoURL.Enabled = False

        mnuEditMov.Enabled = True
        mSR.Enabled = True

        mnuLVChecked.Enabled = True    'main ch
        mnuCheckAll.Enabled = True
        mnuCheckNone.Enabled = True
        mInvCh.Enabled = True
        '
        mnuLVSelected.Enabled = True    'main sel
        mnuSelectAllLV.Enabled = True
        mInvSel.Enabled = True
    Else
        mnuPlayM.Enabled = False
        mGotoURL.Enabled = False
        mnuEditMov.Enabled = False
        mSR.Enabled = False

        mnuLVChecked.Enabled = False    'main ch
        mnuCheckAll.Enabled = False
        mnuCheckNone.Enabled = False
        mInvCh.Enabled = False
        '
        mnuLVSelected.Enabled = False    'main sel
        mnuSelectAllLV.Enabled = False
        mInvSel.Enabled = False
    End If


    If CheckCount > 0 Then
        mnuCheckNone.Enabled = True
        mnuSortChecked.Enabled = True
        mnuDolgCheck.Enabled = True
        mnuLabelCheck.Enabled = True
        mGetCoverCh.Enabled = True
        mnuExportCheckClip.Enabled = True
        mCh2Excel.Enabled = True
        mnuExportCheckHTML.Enabled = True
        mDelCh.Enabled = True
        If CheckCount > 1 Then mCombine.Enabled = True
        '
        mSelCh.Enabled = True
        mUnSelCh.Enabled = True
    Else
        mnuCheckNone.Enabled = False
        mnuSortChecked.Enabled = False
        mnuDolgCheck.Enabled = False
        mnuLabelCheck.Enabled = False
        mGetCoverCh.Enabled = False
        mnuExportCheckClip.Enabled = False
        mCh2Excel.Enabled = False
        mnuExportCheckHTML.Enabled = False
        mDelCh.Enabled = False
        mCombine.Enabled = False
        '
        mSelCh.Enabled = False
        mUnSelCh.Enabled = False

    End If

    If SelCount > 0 Then
        mChSel.Enabled = True
        mUnChSel.Enabled = True
        '
        mnuDolgSel.Enabled = True
        mnuLabelSel.Enabled = True
        mGetCoverSel.Enabled = True
        mnuCopyLV.Enabled = True
        mSel2Excel.Enabled = True
        mnuHTML.Enabled = True
        mDelSel.Enabled = True
        mnuCopyRow.Enabled = True
    Else
        mChSel.Enabled = False
        mUnChSel.Enabled = False
        '
        mnuDolgSel.Enabled = False
        mnuLabelSel.Enabled = False
        mGetCoverSel.Enabled = False
        mnuCopyLV.Enabled = False
        mSel2Excel.Enabled = False
        mnuHTML.Enabled = False
        mDelSel.Enabled = False
        mnuCopyRow.Enabled = False
    End If

    If rs.RecordCount = 0 Then
        mValid.Enabled = False
    Else
        mValid.Enabled = True
    End If


    'ReleaseCapture
    'LVCLICK
    If addflag Or editFlag Then    'перекрыть кислород если в редакторе
    Else
        Me.PopupMenu Me.popLVHid, vbPopupMenuCenterAlign
    End If
    ', vbPopupMenuLeftButton

    'ListView.Enabled = False: ListView.Enabled = True

Case vbMiddleButton
    If ListView.ListItems.Count > 0 Then
        If ListView.ListItems(FrmMain.ListView.SelectedItem.Index).SmallIcon = 0 Then
            ListView.ListItems(FrmMain.ListView.SelectedItem.Index).SmallIcon = 1
            'ListView.ListItems(FrmMain.ListView.SelectedItem.Index).Ghosted = True

        Else
            ListView.ListItems(FrmMain.ListView.SelectedItem.Index).SmallIcon = 0    'не апдейтится без госта
            ListView.ListItems(FrmMain.ListView.SelectedItem.Index).Ghosted = True
            ListView.ListItems(FrmMain.ListView.SelectedItem.Index).Ghosted = False
            'ListView.ListItems(FrmMain.ListView.SelectedItem.Index).ForeColor = RGB(200, 200, 200)



            '.ListView.refresh

        End If
    End If
End Select
'If Not MultiSel Then Timer2.Enabled = True
End Sub



Private Sub LVActer_ItemClick(ByVal Item As ListItem)
''не кликать программно
'If ActNotManualClick Then
'Else
'    LVActClick
'End If
End Sub

Private Sub LVActer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Or KeyCode = 110 Then ComActDel_Click
End Sub

Private Sub mnuAddNewMov_Click()
VerticalMenu_MenuItemClick 3, 0

End Sub

Private Sub mnuCheckAll_Click()
If ListView.ListItems.Count < 1 Then Exit Sub
Dim Itm As ListItem

LV_AllItemsCheck
For Each Itm In ListView.ListItems
    Itm.Checked = True
Next
LVCLICK
End Sub

Private Sub mnuCheckNone_Click()
If ListView.ListItems.Count < 1 Then Exit Sub
Dim Itm As ListItem

LV_AllItemsUnCheck
For Each Itm In ListView.ListItems
    Itm.Checked = False
Next
LVCLICK
End Sub

Private Sub mnuCopyFace_Click()
Clipboard.Clear
'просмотр
PicFaceV.Picture = PicFaceV.Image
Clipboard.SetData PicFaceV.Picture    ', vbCFBitmap

End Sub

Private Sub mnuCopyFoto_Click()
Clipboard.Clear
'Clipboard.SetData PicActFoto.Picture
Clipboard.SetData PicTempHid(1).Picture
End Sub

Private Sub mnuCopyLV_Click()
'mnuCopyLV_Click
'mnuExportCheckClip_Click
'GetCoverSpisokLabel
'GetCoverSpisok

'выделенное - в клипборд
If SelCount < 1 Then Exit Sub
If rs Is Nothing Then Exit Sub

Dim sOldLang As String
Dim j As Integer, M As Integer
Dim temp As String, temp2 As String
'Dim doflag As Boolean
Dim addLabel As Boolean
Dim tmp As String, tmpm As String
Dim pArr() As Integer
Dim allArr() As String

Timer2.Enabled = False
Screen.MousePointer = vbHourglass
DoEvents

'заполнить массив соответствия   индекс в списке - позиция
'индекс массива - позиция, значение - индекс поля списка

ReDim pArr(ListView.ColumnHeaders.Count)    'As Integer
For j = 0 To LstExport_ListCount     'все 0-24, потом индекс не обработаем (было -1 тут и в ListView.ColumnHeaders.Count , но это не дает переместить поле индекса в другое место )
    pArr(ListView.ColumnHeaders(j + 1).Position) = j
Next j

ReDim allArr(UBound(SelRows))
For M = 1 To UBound(SelRows)

    If MultiSel Then RSGoto SelRowsKey(M)

    'одинаковы ли метки (если да, то прикрутим метку сверху списка)
    If M = 1 Then    'первая запись
        tmpm = CheckNoNullVal(dbLabelInd)
    Else
        If tmpm <> vbNullString Then
            If CheckNoNullVal(dbLabelInd) = tmpm Then
                addLabel = True
            Else
                addLabel = False
                tmpm = vbNullString    ' и не надо боле
            End If
        End If
    End If

    'поля
    For j = 1 To ListView.ColumnHeaders.Count   '1-25 индекс не обработаем
        'Debug.Print ListView.ColumnHeaders(j).Text 'не по реальным позициям

        If LstExport_Arr(pArr(j)) Then

            '            'одинаковы ли метки
            '            If M = 1 And j = 1 Then    'каждая запись каждое начало цикла
            '                tmpm = CheckNoNullVal(dbLabelInd)
            '            Else
            '                If tmpm <> vbNullString Then
            '                    If CheckNoNullVal(dbLabelInd) = tmpm Then
            '                        addLabel = True
            '                    Else
            '                        addLabel = False
            '                        tmpm = vbNullString    ' и не надо боле
            '                    End If
            '                End If
            '            End If

            If pArr(j) <> dbAnnotationInd Then
                If pArr(j) = dbFileNameInd Then
                    tmp = GetFNamesForSpisok(CheckNoNullVal(pArr(j)))
                Else
                    tmp = CheckNoNullVal(pArr(j))
                End If

                'название полей
                If Opt_ShowColNames Then
                    If IsNotEmptyOrZero(tmp) Then
                        If pArr(j) <> dbMovieNameInd Then
                            tmp = TranslatedFieldsNames(pArr(j)) & ": " & tmp
                        End If
                    End If
                End If

                If LenB(temp) = 0 Or j = 1 Then    'j - начало цикла
                    ' пробел после пункта списка
                    If IsNotEmptyOrZero(tmp) Then temp = temp & " " & tmp
                Else
                    'иначе разделитель
                    If IsNotEmptyOrZero(tmp) Then temp = temp & ExportDelim & tmp
                End If
            End If
        End If    'doflag
    Next j

    ' + аннотация в конце
    If LstExport_Arr(dbAnnotationInd) Then
        tmp = CheckNoNullVal(dbAnnotationInd)

        If Opt_ShowColNames Then
            If Len(tmp) <> 0 Then
                'название поля описание
                tmp = TranslatedFieldsNames(dbAnnotationInd) & ": " & tmp
            End If
        End If

        If LenB(temp) = 0 Or j = 0 Then    'j - начало цикла
            'пробел после пункта списка
            If Len(tmp) <> 0 Then temp = temp & " " & tmp
        Else
            'иначе разделитель
            If Len(tmp) <> 0 Then temp = temp & ExportDelim & tmp
        End If
    End If

    If Len(temp) <> 0 Then
        If SelCount > 1 Then
            temp = M & "." & temp    ' пункт списка
        Else
            temp = LTrim$(temp)
        End If

        allArr(M) = temp & vbCrLf

    End If
    temp = vbNullString

Next M
temp2 = Join(allArr, vbNullString)

If M > 1 Then    'm=2 если одна запись
    If addLabel Then
        temp2 = frmEditor.LFilm(1) & ": " & tmpm & " (" & M - 1 & ") " & vbCrLf & temp2
    Else
        'temp2 = temp2 & vbCrLf & "(" & M - 1 & ")" 'total to copy
    End If
End If

'скопировать в буфер
Clipboard.Clear
sOldLang = switchLang("00000419")    'сменить раскладку на рус

On Error Resume Next    'иногда не может?

If Len(temp2) = 0 Then temp2 = NamesStore(6)
Clipboard.SetText temp2

If Len(sOldLang) > 0 Then sOldLang = switchLang(sOldLang)    'вернуть раскладку

'восстановить текущ. поз. в базе
RestoreBasePos

Screen.MousePointer = vbNormal
End Sub

Private Sub mnuDolgCheck_Click()
If CheckCount > 0 Then Call MNU_DOLG(False)    'Check
End Sub

Private Sub mnuDolgSel_Click()
If SelCount < 1 Then Exit Sub
Call MNU_DOLG(True)
End Sub

Private Sub mnuEditMov_Click()

VerticalMenu_MenuItemClick 2, 0
End Sub

Private Sub mnuExportCheckClip_Click()
'аналогичные процедуры в
'mnuCopyLV_Click
'mnuExportCheckClip_Click
'GetCoverSpisokLabel
'GetCoverSpisok

If CheckCount = 0 Then Exit Sub
If rs Is Nothing Then Exit Sub

'помеченное - в клипборд
Dim j As Integer, M As Integer, i As Integer
Dim temp As String, temp2 As String
'Dim doflag As Boolean
Dim sOldLang As String
Dim addLabel As Boolean
Dim tmp As String, tmpm As String
Dim pArr() As Integer
Dim allArr() As String

Timer2.Enabled = False
Screen.MousePointer = vbHourglass
DoEvents

'заполнить массис соответствия   индекс в списке - позиция
'индекс массива - позиция, значение - индекс поля списка

ReDim pArr(ListView.ColumnHeaders.Count)    'As Integer
For j = 0 To LstExport_ListCount    'все 0-24, потом индекс не обработаем
    pArr(ListView.ColumnHeaders(j + 1).Position) = j
Next j

ReDim allArr(UBound(CheckRows))
For M = 1 To UBound(CheckRows)

    RSGoto CheckRowsKey(M)

    'одинаковы ли метки
    If M = 1 Then
        tmpm = CheckNoNullVal(dbLabelInd)
    Else
        If tmpm <> vbNullString Then
            If CheckNoNullVal(dbLabelInd) = tmpm Then
                addLabel = True
            Else
                addLabel = False
                tmpm = vbNullString    ' и не надо боле
            End If
        End If
    End If

    'поля
    For j = 1 To ListView.ColumnHeaders.Count    '1-25 индекс не обработаем

        If LstExport_Arr(pArr(j)) Then

            '            'одинаковы ли метки
            '            If M = 1 And j = 1 Then    'каждая запись каждое начало цикла
            '                tmpm = CheckNoNullVal(dbLabelInd)
            '            Else
            '                If tmpm <> vbNullString Then
            '                    If CheckNoNullVal(dbLabelInd) = tmpm Then
            '                        addLabel = True
            '                    Else
            '                        addLabel = False
            '                        tmpm = vbNullString    ' и не надо боле
            '                    End If
            '                End If
            '            End If

            If pArr(j) <> dbAnnotationInd Then

                If pArr(j) = dbFileNameInd Then
                    tmp = GetFNamesForSpisok(CheckNoNullVal(pArr(j)))
                Else
                    tmp = CheckNoNullVal(pArr(j))
                End If

                'название полей
                If Opt_ShowColNames Then
                    If IsNotEmptyOrZero(tmp) Then
                        If pArr(j) <> dbMovieNameInd Then
                            tmp = TranslatedFieldsNames(pArr(j)) & ": " & tmp
                        End If
                    End If
                End If

                If LenB(temp) = 0 Or j = 1 Then    'j - начало цикла
                    ' пробел после пункта списка
                    If IsNotEmptyOrZero(tmp) Then temp = temp & " " & tmp
                Else
                    'иначе разделитель
                    If IsNotEmptyOrZero(tmp) Then temp = temp & ExportDelim & tmp
                End If

            End If
        End If    'doflag
    Next j

    ' + аннотация
    If LstExport_Arr(dbAnnotationInd) Then
        tmp = CheckNoNullVal(dbAnnotationInd)

        If Opt_ShowColNames Then
            If Len(tmp) <> 0 Then
                'название поля описание
                tmp = TranslatedFieldsNames(dbAnnotationInd) & ": " & tmp
            End If
        End If

        If LenB(temp) = 0 Or j = 0 Then    'j - начало цикла
            'пробел после пункта списка
            If Len(tmp) <> 0 Then temp = temp & " " & tmp
        Else
            'иначе разделитель
            If Len(tmp) <> 0 Then temp = temp & ExportDelim & tmp
        End If
    End If

    If Len(temp) <> 0 Then
        If CheckCount > 1 Then
            temp = M & "." & temp    ' пункт списка
        Else
            temp = LTrim$(temp)
        End If

        allArr(M) = temp & vbCrLf

    End If
    temp = vbNullString
Next M
temp2 = Join(allArr, vbNullString)

If M > 1 Then
    If addLabel Then
        temp2 = frmEditor.LFilm(1) & ": " & tmpm & " (" & M - 1 & ") " & vbCrLf & temp2
    Else
        'temp2 = temp2 & vbCrLf & "(" & M - 1 & ")" 'total to copy
    End If
End If

'восстановить текущ. поз. в базе
RestoreBasePos

'скопировать в буфер
Clipboard.Clear
sOldLang = switchLang("00000419")

On Error Resume Next    'иногда не берет?
Clipboard.SetText temp2

If Len(sOldLang) > 0 Then sOldLang = switchLang(sOldLang)

Screen.MousePointer = vbNormal
End Sub





Private Sub mnuExportCheckHTML_Click()
If CheckCount = 0 Then Exit Sub
'False - помеченные v
Export2HTML False
End Sub

Private Sub mnuHTML_Click()
'true - выделенные (не помеченные)
If SelCount < 1 Then Exit Sub
Export2HTML True
End Sub






Private Sub mnuPlayM_Click()
ComPlay_Click
End Sub
Private Sub mnuAutoSizeLV_Click()
Dim ret As Long
'ToDebug "учитывать заголовки при автосайзе?"
ret = myMsgBox(msgsvc(43), vbYesNoCancel, , Me.hwnd)
If ret = vbYes Then
    lvwAutoSizeColumns ListView, True
ElseIf ret = vbNo Then
    lvwAutoSizeColumns ListView, False
Else         'cancel
    Exit Sub
End If         'ret

End Sub
'Private Sub LV_Line_BackColor(ch As Boolean)
'Dim c As Long
'Dim cd As cCommonDialog
'Dim ret As Long
'Dim tempL As Long
''mzt Dim i As Long
'Dim j As Integer
'Dim itmX As ListItem
'
'Set cd = New cCommonDialog
'cd.CustomColor(0) = tempL
'ret = cd.VBChooseColor(c, True, False, False, Me.hWnd)
'If ret = 0 Then Exit Sub
'tempL = c
''Set cd = Nothing
'
'If ch Then 'помеченные
''For i = 1 To UBound(CheckRows) - 1
''If CheckRows(i) Then
''Next i
'
'For Each itmX In ListView.ListItems
'    If itmX.Checked Then
'        itmX.ForeColor = tempL
'        For j = 1 To itmX.ListSubItems.Count
'            itmX.ListSubItems(j).ForeColor = tempL
'        Next j
'    End If
'Next
'
'Else 'выделенные
'
'For Each itmX In ListView.ListItems
'    If itmX.Selected Then
'        itmX.ForeColor = tempL
'        For j = 1 To itmX.ListSubItems.Count
'            itmX.ListSubItems(j).ForeColor = tempL
'        Next j
'    End If
'Next
'
'End If
'
'ListView.Refresh 'для смены цветов сабов
'
'End Sub


Private Sub mnuSaveFace_Click()
'просмотр, сохранить обложку из базы

Dim tmp As String

On Error Resume Next

If Not (rs Is Nothing) Then
    tmp = CheckNoNullVal(dbMovieNameInd)
    ReplaceFNStr tmp
End If
Call SavePicFromBase(1, "FrontFace", tmp)  'из базы

End Sub

Private Sub mnuSaveFoto_Click()
'сохранить фото актера, не из базы,
'или определить по флагам редактирования?

Dim tmp As String

On Error Resume Next

If Not (ars Is Nothing) Then
    If Not IsNull(ars("Name")) Then tmp = ars("Name") Else tmp = vbNullString
    ReplaceFNStr tmp
End If

'SavePicFromPic PicActFoto, FrmMain.hwnd, tmp
SavePicFromPic PicTempHid(1), FrmMain.hwnd, tmp

End Sub

Private Sub mnuSelectAllLV_Click()
'выделить все
If ListView.ListItems.Count < 1 Then Exit Sub
'Dim SelIt As Boolean
Dim Itm As ListItem

For Each Itm In ListView.ListItems
    'If Itm.Index = CurSearch Then
    '    If Itm.Selected Then
    '        SelIt = False
    '    Else
    '        SelIt = True
    '    End If
    'End If

    Itm.Selected = True
Next

ListView.ListItems(CurSearch).Selected = True

'LVCLICK
End Sub

Public Sub mnuShowThisActer_Click()
'ToActFromLV , был CurAct добыт из sub MenuActSelect
'переход на актера
On Error Resume Next
DoEvents
Set LVActer.SelectedItem = LVActer.ListItems(ToActFromLV)
' нет If LVActer.SelectedItem <> sPerson Then mnuShowThisActer.Enabled = False: Exit Sub 'нету, фильтр например
If InStr(LVActer.SelectedItem, sPerson) = 0 Then
    MenuActSelect sPerson    'поискать еще раз (если только что добавили)
    Set LVActer.SelectedItem = LVActer.ListItems(ToActFromLV)
End If
If InStr(LVActer.SelectedItem, sPerson) = 0 Then
    mnuShowThisActer.Enabled = False
    Exit Sub    'нету, фильтр например
End If

TextSearchLVActTypeHid = sPerson    'LVActer.ListItems(LVActer.SelectedItem.Index)
'Debug.Print sPerson
VerticalMenu_MenuItemClick 5, 0
UCLV.Clear
End Sub

Private Sub mnuSortChecked_Click()
If CheckCount = 0 Then Exit Sub

Dim ColHead As MSComctlLib.ColumnHeader

'sort
LVSortColl = -1
'LVSortChecked = True 'SortByCheck 0
'LVSortOrder

ComNext.Enabled = True

LVSortOrder = ListView.SortOrder

'LockWindowUpdate ListView.hWnd

For Each ColHead In ListView.ColumnHeaders
    If right$(ColHead.Text, 2) = " >" Or right$(ColHead.Text, 2) = " <" Then ColHead.Text = left$(ColHead.Text, Len(ColHead.Text) - 2)
Next

SortByCheck 0

'LockWindowUpdate 0

End Sub


Private Sub PicActFoto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

PicActFoto.Picture = PicActFoto.Image

If Not NoPicActFlag Then
    If Button = 2 Then Me.PopupMenu Me.popActHid
End If

End Sub



Public Sub PicActFotoScroll_Resize()
Dim lHeight As Long
Dim lWidth As Long
Dim lProportion As Long
On Error Resume Next

If Me.Visible Then
    If Not FrameActer.Visible Then Exit Sub


    lHeight = (PicActFoto.Height - PicActFotoScroll.ScaleHeight) \ Screen.TwipsPerPixelY
    If (lHeight > 0) Then
        lProportion = lHeight \ (PicActFotoScroll.ScaleHeight \ Screen.TwipsPerPixelY) + 1
        ma_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
        ma_cScroll.Max(efsVertical) = lHeight
        ma_cScroll.Visible(efsVertical) = True
    Else
        ma_cScroll.Visible(efsVertical) = False
        PicActFoto.top = 0
    End If

    lWidth = (PicActFoto.Width - PicActFotoScroll.ScaleWidth) \ Screen.TwipsPerPixelX
    If (lWidth > 0) Then
        lProportion = lWidth \ (PicActFotoScroll.ScaleWidth \ Screen.TwipsPerPixelX) + 1
        ma_cScroll.LargeChange(efsHorizontal) = lWidth \ lProportion
        ma_cScroll.Max(efsHorizontal) = lWidth
        ma_cScroll.Visible(efsHorizontal) = True
    Else
        ma_cScroll.Visible(efsHorizontal) = False
        PicActFoto.left = 0
    End If

End If
End Sub



Private Sub PicCoverPaper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MeForeGround Then PicCoverPaper.SetFocus

DragCover = False
PicCoverTextWnd.BorderStyle = 0
PicCoverTextWnd.MousePointer = vbNormal

End Sub

Private Sub PicCoverTextWnd_LostFocus()
PicCoverTextWnd.BorderStyle = 0
End Sub

Private Sub PicCoverTextWnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PicCoverTextWnd.MousePointer = 7 Or PicCoverTextWnd.MousePointer = 9 Then
    DragCover = True
Else
    DragCover = False
End If
Exit Sub

'Select Case TabStripCover.SelectedItem.Index
'Case 1, 2, 3
'If (0 <= y) And (y < 3) Then
'    PicCoverTextWnd.MousePointer = 7
'    PicCoverTextWnd.BorderStyle = 1
'    DragCover = True
'Else
'    PicCoverTextWnd.MousePointer = vbNormal
'    PicCoverTextWnd.BorderStyle = 0
'    DragCover = False
'End If
'
'Case 4
If (0 <= Y) And (Y < 3) Then
    PicCoverTextWnd.MousePointer = 7
    PicCoverTextWnd.BorderStyle = 1
    CoverMoveDirection = 1
    DragCover = True
ElseIf (PicCoverTextWnd.Height - 3 <= Y) And (Y <= PicCoverTextWnd.Height) Then
    PicCoverTextWnd.MousePointer = 7
    PicCoverTextWnd.BorderStyle = 1
    CoverMoveDirection = 2
    DragCover = True
ElseIf (0 <= X) And (X < 3) Then
    PicCoverTextWnd.MousePointer = 9
    PicCoverTextWnd.BorderStyle = 1
    CoverMoveDirection = 3
    DragCover = True
ElseIf (PicCoverTextWnd.Width - 3 < X) And (X <= PicCoverTextWnd.Width) Then
    PicCoverTextWnd.MousePointer = 9
    PicCoverTextWnd.BorderStyle = 1
    CoverMoveDirection = 4
    DragCover = True
Else
    PicCoverTextWnd.MousePointer = vbNormal
    PicCoverTextWnd.BorderStyle = 1
    DragCover = False
End If


'End Select
End Sub

Private Sub PicCoverTextWnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PicCoverTextWnd.ToolTipText = Int(PicCoverTextWnd.Width) & " x " & Int(PicCoverTextWnd.Height)

If Not DragCover Then
    If (0 <= Y) And (Y < 5) Then
        PicCoverTextWnd.MousePointer = 7
        PicCoverTextWnd.BorderStyle = 1
        CoverMoveDirection = 1
    ElseIf (PicCoverTextWnd.Height - 5 <= Y) And (Y <= PicCoverTextWnd.Height) Then
        PicCoverTextWnd.MousePointer = 7
        PicCoverTextWnd.BorderStyle = 1
        CoverMoveDirection = 2
    ElseIf (0 <= X) And (X < 5) Then
        PicCoverTextWnd.MousePointer = 9
        PicCoverTextWnd.BorderStyle = 1
        CoverMoveDirection = 3
    ElseIf (PicCoverTextWnd.Width - 5 < X) And (X <= PicCoverTextWnd.Width) Then
        PicCoverTextWnd.MousePointer = 9
        PicCoverTextWnd.BorderStyle = 1
        CoverMoveDirection = 4
    Else
        PicCoverTextWnd.MousePointer = vbNormal
        PicCoverTextWnd.BorderStyle = 1
    End If

End If

On Error Resume Next
'Debug.Print PicCoverTextWnd.Top, PicCoverTextWnd.Height

If DragCover Then
    'wndHPix = ScaleY(PicCoverTextWnd.Height, 6, 3)
    'wndWPix = ScaleX(PicCoverTextWnd.Width, 6, 3)

    Select Case TabStripCover.SelectedItem.Index
    Case 1    'stand
        Select Case CoverMoveDirection
        Case 1    'up
            'Debug.Print PicCoverTextWnd.Top ', PicCoverTextWnd.Height '
            'if up and down
            If (PicCoverTextWnd.top + Y > 145) And (PicCoverTextWnd.Height - Y > 9) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top + Y, PicCoverTextWnd.Width, PicCoverTextWnd.Height - Y
            End If
        Case 2    'down
            'роль высоты бокса тут играет y
            'Debug.Print PicCoverTextWnd.Top + y
            If (PicCoverTextWnd.top + Y < 262) And (PicCoverTextWnd.Height + Y > 19) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top, PicCoverTextWnd.Width, Y
            End If

        Case 3    ' left
            If (PicCoverTextWnd.left + X >= 35.4) And ((PicCoverTextWnd.Width - X) >= 15) Then
                'Debug.Print PicCoverTextWnd.Left + x
                PicCoverTextWnd.Move PicCoverTextWnd.left + X, PicCoverTextWnd.top, PicCoverTextWnd.Width - X, PicCoverTextWnd.Height
            End If

        Case 4    'right
            'Debug.Print PicCoverTextWnd.Left + x
            If (PicCoverTextWnd.left + X < 171.5) And (PicCoverTextWnd.Width + X > 30) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top, X, PicCoverTextWnd.Height
            End If

        End Select

    Case 2    'convert
        Select Case CoverMoveDirection
        Case 1    'up
            'Debug.Print PicCoverTextWnd.Top + y, PicCoverTextWnd.Height - y
            'if up and down
            If (PicCoverTextWnd.top + Y > 25) And (PicCoverTextWnd.Height - Y > 9) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top + Y, PicCoverTextWnd.Width, PicCoverTextWnd.Height - Y
            End If
        Case 2    'down
            'роль высоты бокса тут играет y
            'Debug.Print PicCoverTextWnd.Top + y, PicCoverTextWnd.Height + y
            If (PicCoverTextWnd.top + Y < 144) And (PicCoverTextWnd.Height + Y > 19) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top, PicCoverTextWnd.Width, Y
            End If

        Case 3    ' left
            'Debug.Print PicCoverTextWnd.Left + x
            If (PicCoverTextWnd.left + X > 35.4) And ((PicCoverTextWnd.Width - X) >= 15) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left + X, PicCoverTextWnd.top, PicCoverTextWnd.Width - X, PicCoverTextWnd.Height
            End If

        Case 4    'right
            'Debug.Print PicCoverTextWnd.Left + x
            If (PicCoverTextWnd.left + X < 154.5) And (PicCoverTextWnd.Width + X > 30) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top, X, PicCoverTextWnd.Height
            End If

        End Select

    Case 3, 4    'dvd
        Select Case CoverMoveDirection
        Case 1    'up
            'Debug.Print PicCoverTextWnd.Top + y, PicCoverTextWnd.Height - y
            'if up and down
            If (PicCoverTextWnd.top + Y > 15.5) And (PicCoverTextWnd.Height - Y > 9) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top + Y, PicCoverTextWnd.Width, PicCoverTextWnd.Height - Y
            End If
        Case 2    'down
            'роль высоты бокса тут играет y
            'Debug.Print PicCoverTextWnd.Top + y, PicCoverTextWnd.Height + y
            If (PicCoverTextWnd.top + Y < DVD_BotY) And (PicCoverTextWnd.Height + Y > 19) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top, PicCoverTextWnd.Width, Y
            End If

        Case 3    ' left
            'Debug.Print PicCoverTextWnd.Left + x
            If (PicCoverTextWnd.left + X > 10.5) And ((PicCoverTextWnd.Width - X) >= 15) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left + X, PicCoverTextWnd.top, PicCoverTextWnd.Width - X, PicCoverTextWnd.Height
            End If

        Case 4    'right
            'Debug.Print PicCoverTextWnd.Left + x
            If (PicCoverTextWnd.left + X < 139.5) And (PicCoverTextWnd.Width + X > 30) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top, X, PicCoverTextWnd.Height
            End If

        End Select

    Case 5    'list
        Select Case CoverMoveDirection
        Case 1    'up
            'Debug.Print PicCoverTextWnd.Top + y, PicCoverTextWnd.Height - y
            'if up and down
            If (PicCoverTextWnd.top + Y > 8) And (PicCoverTextWnd.Height - Y > 9) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top + Y, PicCoverTextWnd.Width, PicCoverTextWnd.Height - Y
            End If
        Case 2    'down
            'роль высоты бокса тут играет y
            'Debug.Print PicCoverTextWnd.Top + y, PicCoverTextWnd.Height + y
            If (PicCoverTextWnd.top + Y < 285) And (PicCoverTextWnd.Height + Y > 19) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top, PicCoverTextWnd.Width, Y
            End If

        Case 3    ' left
            'Debug.Print PicCoverTextWnd.Left + x
            If (PicCoverTextWnd.left + X > 5) And ((PicCoverTextWnd.Width - X) >= 15) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left + X, PicCoverTextWnd.top, PicCoverTextWnd.Width - X, PicCoverTextWnd.Height
            End If

        Case 4    'right
            'Debug.Print PicCoverTextWnd.Left + x
            If (PicCoverTextWnd.left + X < 205) And (PicCoverTextWnd.Width + X > 30) Then
                PicCoverTextWnd.Move PicCoverTextWnd.left, PicCoverTextWnd.top, X, PicCoverTextWnd.Height
            End If

        End Select

    End Select
End If

End Sub


Private Sub PicCoverTextWnd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicCoverTextWnd.MousePointer = vbNormal
DragCover = False
'PicCoverTextWnd.Line (0, 0)-(PicCoverTextWnd.Width, PicCoverTextWnd.Height), CoverHorBackColor, BF
Select Case TabStripCover.SelectedItem.Index
Case 1
    'координаты
    cov_stan.l = PicCoverTextWnd.left
    cov_stan.t = PicCoverTextWnd.top
    cov_stan.w = PicCoverTextWnd.Width
    cov_stan.H = PicCoverTextWnd.Height

    Call ShowCoverStandard
Case 2
    'координаты
    cov_conv.l = PicCoverTextWnd.left
    cov_conv.t = PicCoverTextWnd.top
    cov_conv.w = PicCoverTextWnd.Width
    cov_conv.H = PicCoverTextWnd.Height

    Call ShowCoverConvert

Case 3
    'координаты
    cov_dvd.l = PicCoverTextWnd.left
    cov_dvd.t = PicCoverTextWnd.top
    cov_dvd.w = PicCoverTextWnd.Width
    cov_dvd.H = PicCoverTextWnd.Height
    Call ShowCoverDVD(False)
Case 4
    'слим
    cov_dvd.l = PicCoverTextWnd.left
    cov_dvd.t = PicCoverTextWnd.top
    cov_dvd.w = PicCoverTextWnd.Width
    cov_dvd.H = PicCoverTextWnd.Height
    Call ShowCoverDVD(True)

Case 5
    'координаты
    cov_list.l = PicCoverTextWnd.left
    cov_list.t = PicCoverTextWnd.top
    cov_list.w = PicCoverTextWnd.Width
    cov_list.H = PicCoverTextWnd.Height

    Call ShowCoverSpisok
End Select
End Sub


Private Sub PicFaceV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tOld As Boolean

If NoPicFrontFaceFlag Then Exit Sub
If NoDBFlag Then Exit Sub
If Button = 2 Then Me.PopupMenu Me.popFaceHid: Exit Sub

tOld = Timer2.Enabled
Timer2.Enabled = False

If Button = 1 Then
    PicTempHid(1).Picture = PicFaceV.Image
    IsCoverShowFlag = True
    ViewScrShotFlag = False
    FormShowPic.Visible = False

    'не видим скролл
    'FormShowPic.hb_cScroll.Visible(efsHorizontal) = False
    FormShowPic.PicHB.Visible = False
    PicManualFlag = True
    ShowInShowPic 1, FrmMain
End If    'but 1

Timer2.Enabled = tOld
End Sub


Private Function MeForeGround() As Boolean
If GetForegroundWindow = Me.hwnd Then MeForeGround = True
End Function


Public Sub PicPrintScroll_Resize()
Dim lHeight As Long
Dim lWidth As Long
Dim lProportion As Long
On Error Resume Next

If Not FrameCover.Visible Then Exit Sub

lHeight = (PicCoverPaper.Height - PicPrintScroll.ScaleHeight) \ Screen.TwipsPerPixelY
If (lHeight > 0) Then
    lProportion = lHeight \ (PicPrintScroll.ScaleHeight \ Screen.TwipsPerPixelY) + 1
    map_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
    map_cScroll.Max(efsVertical) = lHeight
    'map_cScroll.Value(efsVertical) = 0 'map_cScroll.Value(efsVertical) ' + 1
    map_cScroll.Visible(efsVertical) = True

Else
    map_cScroll.Value(efsVertical) = 0
    map_cScroll.Visible(efsVertical) = False
    PicCoverPaper.top = 0
End If

lWidth = (PicCoverPaper.Width - PicPrintScroll.ScaleWidth) \ Screen.TwipsPerPixelX
If (lWidth > 0) Then
    lProportion = lWidth \ (PicPrintScroll.ScaleWidth \ Screen.TwipsPerPixelX) + 1
    map_cScroll.LargeChange(efsHorizontal) = lWidth \ lProportion
    map_cScroll.Max(efsHorizontal) = lWidth
    ' map_cScroll.Value(efsHorizontal) = 0 'map_cScroll.Value(efsHorizontal) ' + 1
    map_cScroll.Visible(efsHorizontal) = True
Else
    map_cScroll.Value(efsHorizontal) = 0
    map_cScroll.Visible(efsHorizontal) = False
    PicCoverPaper.left = 0
End If

End Sub

Private Sub TabLVHid_Click()
'кнопки выбора базы
Dim reclick As Boolean


'не кликать если видны опции и не нажат там выход
If frmOptFlag Then If FrmOptions.Visible Then Exit Sub

If oldTabLVInd = TabLVHid.SelectedItem.Index Then
    NoSetColorFlag = True
    reclick = True
    If InitFlag Then reclick = False
Else
    'oldTabLVInd = TabLVHid.SelectedItem.Index - присвоим, когда база открылась
    FirstLVFill = True
    NoSetColorFlag = False

    If Not FirstActivateFlag Then    'не сохранять ничего при первом старте
        'сохраняем позицию в списке
        If ListView.ListItems.Count > 0 Then
            LastInd = FrmMain.ListView.SelectedItem.SubItems(lvIndexPole)
            WriteKey "LIST", "LastItem", str$(LastInd), iniFileName
        End If
        If LVActer.ListItems.Count > 0 Then
            WriteKey "LIST", "LastItemAct", CStr(Val(CurActKey)), iniFileName
        End If
        'историю
        SaveHistory GetNameFromPathAndName(bdname)
    End If
End If

bdname = LstBases_List(TabLVHid.SelectedItem.Index)
CurrentBaseIndex = TabLVHid.SelectedItem.Index
LstBases_ListIndex = CurrentBaseIndex - 1

If Not reclick Then

    'открыть новую базу
    OpenNewDataBase

    'восстановить историю новой базы
    RestoreHistory GetNameFromPathAndName(bdname)
End If    ' not reclick


FrameView.Caption = FrameViewCaption & " " & ListView.ListItems.Count & " )"
FrmMain.ComFilter.BackColor = &HFFFFFF
End Sub

Private Sub TabLVHid_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If UBound(SelRows) > 1 Then Effect = vbNoDrop: Exit Sub
If myMsgBox(msgsvc(31) & vbCrLf & "(" & SelCount & ")", vbOKCancel, , Me.hwnd) = vbOK Then
    'MouseOverTabLV сперва возникнет TabLVHid_OLEDragOver
    Call LVDragDropSingle(ListView)    ', x, y)
End If
End Sub



Private Sub TabLVHid_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim i As Integer
For i = 1 To TabLVHid.Tabs.Count
    If X > TabLVHid.Tabs.Item(i).left And X < TabLVHid.Tabs.Item(i).Width + TabLVHid.Tabs.Item(i).left Then
        MouseOverTabLV = i
    End If
Next
If TabLVHid.SelectedItem.Index = MouseOverTabLV Then Effect = vbNoDrop
End Sub

Public Sub TabStripCover_Click()

'рефреш после опций
'If oldTabStripCoverInd = TabStripCover.SelectedItem.Index Then Exit Sub 'повтор

ImBlankHid.BackColor = CoverHorBackColor

PicCoverTextWnd.Visible = False
If CheckCount > 0 Then
    ChPrintChecked.Enabled = True
    chPrnAllOne.Enabled = True
Else
    ChPrintChecked.Enabled = False
    'chPrnAllOne.Enabled = False ': chPrnAllOne.Value = vbUnchecked

End If
ChPrintChecked.Value = ChPrintCheckedFlag

Select Case TabStripCover.SelectedItem.Index

Case 1    'cd
    'начальные для обложки-растяжки
    PicCoverTextWnd.Move cov_stan.l, cov_stan.t, cov_stan.w, cov_stan.H
    'ToDebug "...обложка Standart Case"
    Call ShowCoverStandard

Case 2    'конверт
    PicCoverTextWnd.Move cov_conv.l, cov_conv.t, cov_conv.w, cov_conv.H
    'ToDebug "...обложка Convert"
    Call ShowCoverConvert

Case 3    'dvd
    PicCoverTextWnd.Move cov_dvd.l, cov_dvd.t, cov_dvd.w, cov_dvd.H
    'ToDebug "...обложка DVD"
    Call ShowCoverDVD(False)
Case 4    'dvd slim
    PicCoverTextWnd.Move cov_dvd.l, cov_dvd.t, cov_dvd.w, cov_dvd.H
    'ToDebug "...обложка DVD slim"
    Call ShowCoverDVD(True)

Case 5    'список
    PicCoverTextWnd.Move cov_list.l, cov_list.t, cov_list.w, cov_list.H
    'ToDebug "...обложка List"

    ChPrintCheckedFlag = ChPrintChecked.Value
    ChPrintChecked.Value = 0: ChPrintChecked.Enabled = False

    Call ShowCoverSpisok

End Select

oldTabStripCoverInd = TabStripCover.SelectedItem.Index
End Sub

Private Sub TextActBio_KeyDown(KeyCode As Integer, Shift As Integer)
If Not ActFlag Then KeyCode = 0
End Sub

Private Sub TextActBio_KeyPress(KeyAscii As Integer)
'If Not ActFlag Then KeyAscii = 0
End Sub

Private Sub TextActName_KeyDown(KeyCode As Integer, Shift As Integer)
'If Not ActFlag Then KeyCode = 0
If KeyCode = 13 Then TextActBio.SetFocus    'TextActName_LostFocus
End Sub

Private Sub TextActName_KeyPress(KeyAscii As Integer)
'If Not ActFlag Then KeyAscii = 0
End Sub


Private Sub TextFind_Change()
ComNext.Enabled = True
End Sub

Private Sub TextFind_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then Call ComFind_Click
If CombFind.Text = vbNullString Then CombFind.ListIndex = 0
End Sub


Private Sub TextItemHid_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then    'ctrl+A
    TextItemHid.SelStart = 0: TextItemHid.SelLength = Len(TextItemHid)
End If
End Sub

Private Sub TextItemHid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If addflag Or editFlag Then Exit Sub    'перекрыть кислород если в редакторе - хотя поле уже дизейбл во фрейме PicSplitLVDHid
If Button = vbRightButton And Shift = 0 Then
    DisplayTextPopup TextItemHid

Else
    'ctrl выбор имени актера (не шифт)
    If Shift = 2 Then SelectWordsGroup TextItemHid, TextItemHid.SelStart    'нажать на ктрл чтобы автовыделилось
    'If Shift = 0 Then SelectWordsGroup TextItemHid, TextItemHid.SelStart 'в uclv наоборот
End If
End Sub



Private Sub TextItemHid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Len(TextItemHid.SelText) <> 0 Then sPerson = TextItemHid.SelText

PutCoverUCLV (MenuActSelect(TextItemHid.SelText))    'в UCLV или актера или ковер, если нет фото актера

End Sub


Private Sub TextSearchLVActTypeHid_Change()
'Поиск в актерах
Dim itmX As ListItem

'in view
For Each itmX In LVActer.ListItems
    If InStr(1, itmX.Text, TextSearchLVActTypeHid.Text, vbTextCompare) <> 0 Then
        Set LVActer.SelectedItem = itmX
        'LVActer.ListItems(itmX.Index).EnsureVisible
        LV_EnsureVisible LVActer, LVActer.SelectedItem.Index
        LVActClick    'itmX
        Exit Sub
    End If
Next

End Sub

Private Sub TextSearchLVActTypeHid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nextFlag As Boolean
Dim itmX As ListItem

If KeyCode = 13 Then

    For Each itmX In LVActer.ListItems
        If nextFlag Then
            If InStr(1, itmX.Text, TextSearchLVActTypeHid.Text, vbTextCompare) <> 0 Then
                Set LVActer.SelectedItem = itmX
                'LVActer.ListItems(itmX.Index).EnsureVisible
                LV_EnsureVisible LVActer, LVActer.SelectedItem.Index
                LVActClick    'itmX
                Exit Sub
            End If
        End If
        If LVActer.SelectedItem = itmX Then nextFlag = True
    Next

End If        '13
End Sub

Private Sub TextVAnnot_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then ComNext_Click
End Sub


Private Sub TextVAnnot_KeyPress(KeyAscii As Integer)
Select Case KeyAscii

Case 6    '^F
    ComFind_Click

End Select
End Sub


Public Sub Timer2_Timer()
On Error Resume Next

DoEvents    'для прорисовки lvadd например

'if meResize then
'Timer2.Enabled = False: Exit Sub

If timerflag Then If frmEditorFlag Then Timer2.Enabled = False: Exit Sub
If Opt_NoSlideShow Then Timer2.Enabled = False: Exit Sub

'FrameView.Refresh

If FrameView.Visible Then
    Timer2.Enabled = True
Else
    Timer2.Enabled = False
    Exit Sub
End If


If NoPic1Flag And NoPic2Flag And NoPic3Flag Then
    Timer2.Enabled = False
    Image0.Cls
    Image0.Move 0, 0, FrameImageHid.Width, FrameImageHid.Height
    Image0.PaintPicture ImageList.ListImages(LastImageListInd).Picture, 0, 0, Image0.Width, Image0.Height    'nopic
    Exit Sub
End If


Select Case SlideShowFlag

Case 0
    If Not NoPic2Flag Then
        SlideShowFlag = 1
    ElseIf Not NoPic3Flag Then
        SlideShowFlag = 2
    Else
        Timer2.Enabled = False
    End If

    If Not NoPic1Flag Then
        SlideShowLastFlag = 0
        GetPic Image0, 1, "SnapShot1"
    Else
        Timer2_Timer    'быстрее еще раз, если нет картинки
    End If

Case 1
    If Not NoPic3Flag Then
        SlideShowFlag = 2
    ElseIf Not NoPic1Flag Then
        SlideShowFlag = 0
    Else
        Timer2.Enabled = False
    End If

    If Not NoPic2Flag Then
        SlideShowLastFlag = 1
        GetPic Image0, 1, "SnapShot2"
    Else
        Timer2_Timer
    End If

Case 2
    If Not NoPic1Flag Then
        SlideShowFlag = 0
    ElseIf Not NoPic2Flag Then
        SlideShowFlag = 1
    Else
        Timer2.Enabled = False
    End If

    If Not NoPic3Flag Then
        SlideShowLastFlag = 2
        GetPic Image0, 1, "SnapShot3"
    Else
        Timer2_Timer
    End If

End Select

If Abs(NoPic1Flag) + Abs(NoPic2Flag) + Abs(NoPic3Flag) = 2 Then Timer2.Enabled = False

timerflag = True


'Debug.Print Time & " " & SlideShowLastGoodPic

End Sub


Private Sub tvGroup_Click()
DoEvents
TVCLICK
End Sub

Private Sub tvGroup_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case ColumnHeader.Index
Case 1

    'галочку
    Dim i As Integer
    For i = 0 To mGroup.Count - 1
        mGroup(i).Checked = False
    Next i
    If GroupInd = -1 Then
        mGroup(0).Checked = True
    Else
        mGroup(GroupInd + 1).Checked = True
    End If

    'выпасть
    Me.PopupMenu mPopGroup    ', DefaultMenu:=mGroup(2)

Case 2    'сорт по нумерации

    If tvGroup.ListItems.Count < 1 Then Exit Sub
    'If tvGroup.ListItems.Item(1).ListSubItems.Count < 1 Then 'Заполнить кол-вом
    If Len(tvGroup.ListItems.Item(1).SubItems(1)) = 0 Then


        FillTVNums    'заполнить номерами
        Exit Sub

    End If

    tvGroup.SortOrder = (tvGroup.SortOrder + 1) Mod 2
    SortByNumber 1, tvGroup
End Select

End Sub

Private Sub tvGroup_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 16, 17, 18, 20, 27
Case Else
    DoEvents
    TVCLICK
End Select
End Sub

Private Sub tvGroup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    'tvGroup.SortOrder = (tvGroup.SortOrder + 1) Mod 2
    'tvGroup.Sorted = True 'для нормальной работы сорта по тексту

    tvGroup.Sorted = True    'для нормальной работы сорта по тексту

    SortByNumber 0, tvGroup
    'если там .Sorted = True, то хорошо сортирует и текст, есди нет - текст не сортируется (число сортируется)

    tvGroup.Sorted = False    'для переноса tvGroup.SortOrder = после сортировки
    tvGroup.SortOrder = (tvGroup.SortOrder + 1) Mod 2
End If
End Sub

Private Sub tvGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'поймать фокус
On Error Resume Next
If frmSRFlag Then
    If frmSR.Visible Then Exit Sub
End If
If txtEdit.Visible Then Exit Sub
If LstFiles.Visible Then Exit Sub
If TextItemHid.SelLength > 0 Then Exit Sub

If GetForegroundWindow = Me.hwnd Then
    If ActiveControl.name <> "tvGroup" Then tvGroup.SetFocus
End If
End Sub



Private Sub UCLV_tActMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, s As String)
'с шифтом выпадет системное меню
If addflag Or editFlag Then Exit Sub    'перекрыть кислород если в редакторе

If Button = 2 And Shift = 0 Then

    '   ' Make VB discard the mouse capture.
    '    txt.Enabled = False
    '    txt.Enabled = True
    '    txt.SetFocus

    ' Display the custom menu.
    'может сделать? копи пасты работают с ActiveControl, надо текстбокс, а сейчас UC?
    m7Hid.Visible = False
    mnuUndo.Visible = False
    mnuCut.Visible = False
    mnuCopy.Visible = False
    mnuPaste.Visible = False
    mnuDelete.Visible = False
    m8Hid.Visible = False
    mnuSelectAll.Visible = False

    If Len(sPerson) = 0 Then mFiltAct.Enabled = False Else mFiltAct.Enabled = True


    PopupMenu mnuTextMenuHid
End If
End Sub

Private Sub UCLV_tActMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, s As String)
'поискать персону s в базе актеров
If Len(s) <> 0 Then sPerson = s
PutCoverUCLV (MenuActSelect(s))  'в UCLV или актера или ковер, если нет фото актера

'ImListFoto.ListImages.Add 1, "foto", PicActFoto.Picture
'    cIM.IconIndex("foto") = ImListFoto.ListImages(1).Picture
'ImageList.ListImages("SAVE_ICON").Index - 1


End Sub

Public Sub VerticalMenu_MenuItemClick(MenuItem As Integer, Shift As Integer)
Dim i As Integer
Dim ret As Integer

If NoDBFlag Then
    bdname = vbNullString
Else
    '''''''''не нажимать

    If MenuItem = 1 Then    'нажали мэйн
        If addflag Or editFlag Then Exit Sub    'уже должно быть видно и нажимать нельзя
        If frmOptFlag And InitFlag Then Exit Sub    'не, когда в опциях
    End If

    If MenuItem = 2 Then    'нажали редактор
        If addflag Then Exit Sub    'не редактировать, когда в добавлении нового
        If frmOptFlag Then Exit Sub    'не редактировать, когда в опциях
        If rs Is Nothing Then Exit Sub
    End If

    If MenuItem = 3 Then    'добавл новый
        If frmOptFlag Then Exit Sub    'не добавлять, когда в опциях
        If rs Is Nothing Then Exit Sub
    End If

    If MenuItem = 4 Then    'нажали обложку
        If addflag Or editFlag Then Exit Sub
        If rs Is Nothing Then Exit Sub
    End If

    If MenuItem = 5 Then    'нажали актеров
        If addflag Or editFlag Then Exit Sub
        If ars Is Nothing Then Exit Sub
    End If

    If MenuItem = 6 Then    'opt
        If addflag Or editFlag Then Exit Sub
    End If

    If MenuItem = 7 Then    'нажали статист
        If addflag Or editFlag Then Exit Sub
        If rs Is Nothing Then Exit Sub
    End If

    If MenuItem = 3 Then    'не кликать повторы кроме '3 - добавлять новый
        'If MenuItem = 1 And InitFlag = True Then
    Else
        If MenuItem = LastVMI Then
            Exit Sub
        End If
        'End If
    End If

End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ToDebug "MenuItem: " & MenuItem & "/" & Shift


If NoDBFlag And MenuItem <> 8 And MenuItem <> 5 And MenuItem <> 6 Then
    ToDebug "SVC просит базу фильмов"
    myMsgBox msgsvc(21), vbInformation, , Me.hwnd     'Откройте базу фильмов.
    'ComOpenBD.SetFocus
    GoTo Opt
End If
If NoDBFlag And (MenuItem <> 6) And (MenuItem <> 8) Then MenuItem = 5    'в актеры

''''''''''''''''''''''''''''''''''''''''''''''''авто
If Shift = 1 Then
    If MenuItem = 3 Then
        If BaseReadOnly Or BaseReadOnlyU Then
            'myMsgBox msgsvc(24), vbInformation, , Me.hwnd
            Exit Sub
        End If

        frmAuto.Show 1, Me
        If LastVMI <> 1 Then VerticalMenu_MenuItemClick 1, 0
        Exit Sub
    End If
End If

'''''''''''''''''''''''''''''''''''''''''''''''''Статистика
If MenuItem = 7 Then  '
    ToDebug "Статистика"
    FrmStat.Show 0, Me
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''Настройки
If MenuItem = 6 Then
    'optsaved = True ' нет изменений, false - было изменение
Opt:
    Timer2.Enabled = False
    'ToDebug "Настройки"
    If SplashFlag Then Unload frmSplash: Set frmSplash = Nothing
    'If addFlag Or editFlag Then 'нет опций при редакторе (и наоборот)
    '    Exit Sub
    'Else
    FrmOptions.Show 0, Me
    'End If
    'PrevVMI = LastVMI
    'LastVMI = 6
    Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'если редактируется актер, спросить, что делать
If ars Is Nothing Then
Else
    If ars.EditMode Then
        ComActSave.Enabled = False: ComActPast.Enabled = False: ComActFile.Enabled = False: ComActFotoDel.Enabled = False
        ret = myMsgBox(msgsvc(52), vbYesNoCancel, , Me.hwnd)
        If ret = vbNo Then    'no
            ToDebug "- не сохранять"
            ComCancelAct_Click
        ElseIf ret = vbYes Then
            ToDebug "- сохранить"
            ComActSave_Click
            ActFlag = False
        Else    'cancel
            If ActNewFlag Then ars.CancelUpdate
            ComActEdit_Click
            Exit Sub
        End If    'ret
    End If    'ars.EditMode
End If    'ars is nothing

'If rs Is Nothing Then
'Else
'    'если редактируется База фильмов, спросить, что делать
'    If rs.EditMode Then
'        ret = myMsgBox(msgsvc(6), vbYesNoCancel, , Me.hwnd)
'        If ret = vbYes Then
'            ToDebug "- сохранить"
'            SaveFromEditor
'        ElseIf ret = vbNo Then
'            ToDebug "- не сохранять"
'            rs.CancelUpdate
'        Else    'cancel
'            Exit Sub
'        End If    'ret
'    End If    'rs.EditMode
'End If    'rs nodb:

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'химчистка
ListView.MultiSelect = True
Timer2.Enabled = False
Set map_cScroll = Nothing    'print

If m_cAVI Is Nothing Then
Else
    m_cAVI.filename = vbNullString    'unload
    Set m_cAVI = Nothing
End If
If MpegMediaOpen Then Call MpegMediaClose    'очистка экрана видео

sTimeSum = vbNullString

isMPGflag = False: isAVIflag = False: isDShflag = False    'или ClearVideo()?

FrameActer.Visible = False
FrameCover.Visible = False
If FirstActivateFlag And Not NoDBFlag Then
    FrameView.Visible = True
Else
    If MenuItem = 2 Or MenuItem = 3 Or MenuItem = 1 Then Else FrameView.Visible = False
End If

Mark2SaveFlag = True    '?
ShowCoverFlag = False
AutoNoMessFlag = False

If (MenuItem <> 1) And frmSRFlag Then frmSR.Hide    'показ только в списке
Unload frmActFilt    'убрать фильтр актеров

'With frmEditor
''не тут, после сохранения нового мы чистить будем видео
'    .Position.Enabled = False: .PositionP.Enabled = False
'    .Position.Value = 0: .PositionP.Value = 0
'    .ComKeyAvi(0).Enabled = False: .ComKeyAvi(1).Enabled = False
'    .optAspect(0).Value = False: .optAspect(1).Value = False
'    Set .ImgPrCov = Nothing
'    Set .PicSS1Big = Nothing: Set .PicSS2Big = Nothing: Set .PicSS3Big = Nothing
'    'Set .PicFrontFace = Nothing: Set .picCanvas = Nothing    'в редакторе
'End With
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'If NoDBFlag And MenuItem < 8 Then
'    If optsaved Then
'        FrmOptions.Show vbModal, Me
'        ToDebug "SVC просит таки открыть базу с фильмами"
'        myMsgBox msgsvc(21), vbInformation, , Me.hwnd 'Откройте базу фильмов.
'        'ComOpenBD.SetFocus
'    End If
'    GoTo Opt
'End If

'чистка памяти
'CoFreeUnusedLibraries
'SetWrkSize


If MenuItem = 1 Then    '                                            просмотр
    'химч ListView.MultiSelect = True
    If Me.Visible Then FrameView.Visible = True
    If Not FirstActivateFlag Then Form_Resize    ' мож не надо?

    If InitFlag Then
        ToDebug "Заполнение фильмов"
        If TabLVHid.Tabs.Count >= CurrentBaseIndex Then         'LastBaseInd Then
            TabLVHid.Tabs(CurrentBaseIndex).Selected = True
        Else
            TabLVHid.Tabs(1).Selected = True
        End If

        FrmMain.ComFilter.BackColor = &HFFFFFF
    Else
        'просто кликнуть
        If GroupedFlag Then
            TVCLICK
        Else
            LVCLICK
        End If
    End If

    'ListView.MultiSelect = True
    InitFlag = False: addflag = False: OpenAddmovFlag = False: editFlag = False
    ' Unload frmEditor

    If Not (ListView.SelectedItem Is Nothing) Then
        If Not MultiSel Then
            If Not FirstActivateFlag Then LV_EnsureVisible ListView, ListView.SelectedItem.Index
        End If
    End If

    FrameView.Caption = FrameViewCaption & " " & ListView.ListItems.Count & " )"


    CheckSameDisk = True

    ' вернуть если уберем из MpegMediaClose
    Clear_mobjManager

    ToDebug "Список фильмов"
    Screen.MousePointer = vbNormal
    PrevVMI = LastVMI
    LastVMI = 1
    Exit Sub
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' новый
If MenuItem = 3 Then    '
    If rs Is Nothing Then VerticalMenu_MenuItemClick 1, 0: Exit Sub

    If BaseReadOnly Then
        LastVMI = 3
        VerticalMenu_MenuItemClick 1, 0
        myMsgBox msgsvc(24), vbInformation, , Me.hwnd
        Exit Sub
    End If

    FrameView.Visible = True


    'With frmEditor
    '    Set .PicFrontFace = Nothing: Set .picCanvas = Nothing    'в новом
    'End With
    With frmEditor
        .Position.Enabled = False: .PositionP.Enabled = False
        .Position.Value = 0: .PositionP.Value = 0
        .ComKeyAvi(0).Enabled = False: .ComKeyAvi(1).Enabled = False
        .optAspect(0).Value = False: .optAspect(1).Value = False
        Set .ImgPrCov = Nothing
        Set .PicSS1Big = Nothing: Set .PicSS2Big = Nothing: Set .PicSS3Big = Nothing
        Set .PicFrontFace = Nothing: Set .picCanvas = Nothing
    End With

    ListView.MultiSelect = False

    ToDebug "Новый фильм"
    Screen.MousePointer = vbHourglass

    addflag = True: editFlag = False



    With frmEditor
        .ComDel.Enabled = False: .ComDel.Visible = True: .ComAdd.Enabled = False
        .TxtIName.Text = vbNullString

        If rs.EditMode Then rs.CancelUpdate
        rs.AddNew

        ClearVideo
        ClearFields

        Set .PicSS1 = Nothing: .PicSS1.Width = ScrShotEd_W: .PicSS1.Height = ScrShotEd_H
        Set .PicSS2 = Nothing: .PicSS2.Width = ScrShotEd_W: .PicSS2.Height = ScrShotEd_H
        Set .PicSS3 = Nothing: .PicSS3.Width = ScrShotEd_W: .PicSS3.Height = ScrShotEd_H

        NoPicFrontFaceFlag = True: NoPic1Flag = True: NoPic2Flag = True: NoPic3Flag = True

        .TabStrAdEd.SelectedItem = .TabStrAdEd.Tabs(.TabStrAdEd.SelectedItem.Index)
        InitFlag = False

        .FrameAddEdit.Caption = VerticalMenu.Controls("LVMB")(2)    'AddEditCapt
        .ComSaveRec.BackColor = &HC0C0E0    'надо

        PicSplitLVDHid.Enabled = False    'не кликать картинки в мэйне

        .Show , FrmMain    'frmeditor.show
        frmEditorFlag = True

        'фокусы текстовых полей после TabStrAdEd.SelectedItem
        If .TextCDN.Visible Then .TextCDN.SetFocus
        If .TextMName.Visible Then .TextMName.SetFocus: .TextMName.SelLength = 0

    End With

    'ListView.MultiSelect = True

    PrevVMI = LastVMI
    LastVMI = 3

    Screen.MousePointer = vbNormal
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''редактор
If MenuItem = 2 Then
    If rs Is Nothing Then VerticalMenu_MenuItemClick 1, 0: Exit Sub
    If rs.RecordCount < 1 Then LastVMI = 2: VerticalMenu_MenuItemClick 1, 0: Exit Sub
    If ListView.ListItems.Count < 1 Then LastVMI = 2: VerticalMenu_MenuItemClick 1, 0: Exit Sub
    If BaseReadOnly Then
        LastVMI = 2
        VerticalMenu_MenuItemClick 1, 0
        myMsgBox msgsvc(24), vbInformation, , Me.hwnd
        Exit Sub
    End If
    If BaseReadOnlyU Then
        LastVMI = 2
        VerticalMenu_MenuItemClick 1, 0
        myMsgBox msgsvc(22), vbInformation, , Me.hwnd
        Exit Sub
    End If

    FrameView.Visible = True
    'ListView.MultiSelect = False

    ToDebug "Edit_Key=" & rs("Key")

    addflag = False: editFlag = True

    frmEditor.EditorNoVideoClear 'дизейблить ползунки
    '    .ComRND(0).Enabled = False: .ComRND(1).Enabled = False: .ComRND(2).Enabled = False
    '    For i = 0 To 2: .optAspect(i).Enabled = False: Next i

    'и еще кое что
    With frmEditor
        .movie.Cls: .movie.Width = MovieEd_W: .movie.Height = MovieEd_H: .movie.Visible = True
        .ComDel.Visible = True: .ComDel.Enabled = True

        If rs.EditMode Then rs.CancelUpdate
        .ComAdd.Enabled = False

        .PicSS1.Width = ScrShotEd_W: .PicSS1.Height = ScrShotEd_H: Set .PicSS1 = Nothing
        .PicSS2.Width = ScrShotEd_W: .PicSS2.Height = ScrShotEd_H: Set .PicSS2 = Nothing
        .PicSS3.Width = ScrShotEd_W: .PicSS3.Height = ScrShotEd_H: Set .PicSS3 = Nothing

        CheckSameDisk = False

        GetEditPix
        Mark2SaveFlag = False
        GetFields
        Mark2SaveFlag = True
        'lbInetMovieList.Clear

        .ComSaveRec.BackColor = &HC0E0C0

        If ListView.ListItems.Count > 0 Then LastInd = ListView.SelectedItem.SubItems(lvIndexPole)

        'кликнуть таб
        .TabStrAdEd.SelectedItem = .TabStrAdEd.Tabs(.TabStrAdEd.SelectedItem.Index)

        PicSplitLVDHid.Enabled = False    'не кликать картинки в мэйне
        .Show , FrmMain
        frmEditorFlag = True

    End With
    InitFlag = False

    PrevVMI = LastVMI
    LastVMI = 2
    Exit Sub
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


If MenuItem = 5 Then    '                                 актеры


    If LenB(abdname) = 0 Then VerticalMenu_MenuItemClick LastVMI, 0: Exit Sub   'goto last menu

    ToDebug "Люди кино"
    Screen.MousePointer = vbHourglass
    ComActSave.Enabled = False
    ComActPast.Enabled = False
    ComActFile.Enabled = False
    ComActFotoDel.Enabled = False
    ComCancelAct.Enabled = False
    addflag = False: editFlag = False

    If Len(bdname) = 0 Then ComSelMovIcon.Enabled = False
    LVActer.ColumnHeaders(1).Width = LVActer.Width - 70
    LVActer.Sorted = True
    If Not LVActerFilled Then
        FillActListView
    Else
        'If Not IsNull(LVActer.SelectedItem.Index) Then LV_EnsureVisible LVActer, LVActer.SelectedItem.Index
        If Not (LVActer.SelectedItem Is Nothing) Then LV_EnsureVisible LVActer, LVActer.SelectedItem.Index
    End If

    ListBActHid.Clear
    LActMarkCount.Caption = LActMarkCountCaption + " " + str$(mcount)

    LVActer.ColumnHeaders(1).Width = LVActer.Width - 380    'нет гор прокрутке

    'проверка пометок красными галками в ListView
    Dim itmX As ListItem
    i = 0
    For Each itmX In ListView.ListItems
        If itmX.SmallIcon = 1 Then i = i + 1
    Next

    fr_acter = True
    Form_Resize
    Timer2.Enabled = False

    fr_acter = False
    FrameActer.Visible = True

    LVActClick    'клик на помеченной записи

    Screen.MousePointer = vbNormal

    LActMarkCount.Caption = LActMarkCountCaption & " " & i
    PrevVMI = LastVMI
    LastVMI = 5

    Exit Sub
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''обложка
If MenuItem = 4 Then

    If rs.RecordCount < 1 Then LastVMI = 4: VerticalMenu_MenuItemClick 1, 0: Exit Sub
    If ListView.ListItems.Count < 1 Then LastVMI = 4: VerticalMenu_MenuItemClick 1, 0: Exit Sub

    Timer2.Enabled = False
    ToDebug "Обложка"

    addflag = False: editFlag = False
    Screen.MousePointer = vbHourglass

    'A4
    PicCoverPaper.ScaleWidth = 210
    PicCoverPaper.ScaleHeight = 296

    ' Set up scroll bars
    Set map_cScroll = New cScrollBars
    map_cScroll.Create PicPrintScroll.hwnd
    PicCoverPaper.Move 0, 0



    NoPic1Flag = False: NoPic2Flag = False: NoPic3Flag = False    ' а то после "нового" неберучка картинок
    PicCoverTextWnd.BorderStyle = 0

    ChPrintChecked.Caption = ChPrintCheckedCaption & ": " & CheckCount
    If CheckCount > 1 Then
        ChPrintChecked.Enabled = True
        chPrnAllOne.Enabled = True
    Else
        ChPrintChecked.Enabled = False
        'chPrnAllOne.Enabled = False 'лишний клик, или флаг ставить chPrnAllOne.Value = vbUnchecked
    End If

    ImBlankHid.BackColor = CoverHorBackColor

    Select Case TabStripCover.SelectedItem.Index
    Case 1
        Call ShowCoverStandard
    Case 2
        Call ShowCoverConvert
    Case 3
        Call ShowCoverDVD(False)
    Case 4
        Call ShowCoverDVD(True)

    Case 5
        ChPrintChecked.Value = 0: ChPrintChecked.Enabled = False     'убрать галку печатать все
        Call ShowCoverSpisok
    End Select



    Screen.MousePointer = vbNormal
    PrevVMI = LastVMI
    LastVMI = 4
    Form_Resize
    Timer2.Enabled = False

    ShowCoverFlag = True
    Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If MenuItem = 8 Then    '                                выход
    ToDebug "Выход."
    Unload Me
End If
End Sub




Public Sub FillActListView()
'Dim liChild As ListItem
Dim i As Integer
'Dim temp As String


If ars.RecordCount = 0 Then Exit Sub

ToDebug "Заполнение актеров"
Screen.MousePointer = vbHourglass

LVActer.Visible = False
LVActer.ListItems.Clear
LVActer.Sorted = False

ars.MoveFirst

On Error Resume Next    'надо для null
i = 0
Do While Not ars.EOF
    i = i + 1

    'temp = vbNullString
    'If ars.Fields("Name").Value <> vbNullString Then
    '    temp = ars.Fields("Name").Value
    'End If

    'If Not IsNull(ars("Name")) Then temp = ars("Name") Else temp = vbNullString

    'ключ
    LVActer.ListItems.Add i, ars("Key") & Kavs, ars("Name")    'temp

    LVActer.ListItems(i).SubItems(1) = i
    ars.MoveNext
Loop
err.Clear

LVActer.Sorted = True: LVActer.Visible = True

'поместить на запомненный ключ
GotoLVAct CurActKey

LV_EnsureVisible LVActer, LVActer.SelectedItem.Index
LastIndAct = LVActer.ListItems(LVActer.SelectedItem.Index).SubItems(1)
CurAct = LVActer.SelectedItem.Index

'Debug.Print "After Fill L=" & LastIndAct & " C=" & CurAct

FrameActer.Caption = FrameActerCaption & LVActer.ListItems.Count & ")"
LVActerFilled = True
Screen.MousePointer = vbNormal
End Sub

'Private Sub clearLVIcon()
'Dim i As Integer
'For i = 1 To ListView.ListItems.Count
'ListView.ListItems.Item(i).SmallIcon = 0 'blank
'Next i
'End Sub

Public Sub FillListView()
Dim j As Integer, i As Integer
Dim ind As Long
'Dim tmpFlag As Boolean
'Dim ColorFlag As Boolean
Dim temp As String
'Dim Itm As ListItem

Dim rsCount As Long


On Error Resume Next
'CheckedInLV = 0
'rs.MoveFirst: rs.MoveLast: rs.MoveFirst
rs.MoveLast: rs.MoveFirst
rsCount = rs.RecordCount
'ModLVSubClass.UnAttach FrameView.hWnd
If rsCount < 1 Or err Then
    FrmMain.FrameView.Caption = FrameViewCaption & " 0 )"
    ListView.ListItems.Clear
    TextVAnnot.Text = vbNullString
    UCLV.Clear
    Set PicFaceV = Nothing
    NoPicFrontFaceFlag = True
    NoPic1Flag = True: NoPic2Flag = True: NoPic3Flag = True
    Timer2_Timer    'убивать image0
    Screen.MousePointer = vbNormal
    Exit Sub
End If
On Error GoTo 0

'Screen.MousePointer = vbHourglass
'If frmFilterFlag Then tmpFlag = True
If frmFilterFlag Then Me.Refresh

'DoEvents нет!

If FormShowPicLoaded Then Unload FormShowPic: Set FormShowPic = Nothing
If ListView.SelectedItem Is Nothing Then: Else CurSearch = ListView.SelectedItem.Index

nScrollPos = GetScrollPos(ListView.hwnd, SB_HORZ)
ListView.ListItems.Clear
ListView.Sorted = False

ReDim CheckRows(0): ReDim SelRows(0): ReDim lvItemLoaded(0)
ReDim CheckRowsKey(0): ReDim SelRowsKey(0)

On Error Resume Next
'прогрессбары
PBar.min = 0: PBar.Max = rsCount: PBar.Value = 0
If SplashFlag Then frmSplash.PBar.min = 0: frmSplash.PBar.Max = PBar.Max
'TextItemHid.ZOrder 0
PBar.ZOrder 0

ind = 1

'LockWindowUpdate ListView.hwnd
'LockWindowUpdate 0

'массив загрузок (полный - названия) - сейчас не используется...
ReDim lvItemLoaded(rsCount)

'Dim t As Long
't = timeGetTime()
'FrameView.Enabled = False
'ListView.Visible = False

'заполнение списка из базы
Do While Not rs.EOF

    'If ind / 10 = Int(ind / 10) Then
    If (ind Mod 10) = 0 Then
        If SplashFlag Then
            frmSplash.PBar.Value = ind
        Else
            PBar.Value = ind
        End If
    End If

    'Title
    If Not IsNull(rs(dbMovieNameInd)) Then temp = rs(dbMovieNameInd) Else temp = vbNullString    'нужно, иначе не создается новое поле, если Null

    ListView.ListItems.Add ind, rs.Fields("Key") & Kavs, temp
    'lvAddItem ListView.hwnd, temp, ind, , 0
    'Index
    ListView.ListItems(ind).SubItems(lvIndexPole) = ind - 1

    'Пометки
    'If Not IsNull(rs(dbCheckedInd)) Then temp = rs(dbCheckedInd) Else temp = vbNullString
    ListView.ListItems(ind).Checked = Val(rs(dbCheckedInd))  'Val(temp)

    'Val(CheckNoNullVal(dbCheckedInd))
    ' CheckedInLV = CheckedInLV + Abs(ListView.ListItems(ind).Checked)

    'плюсовать размер массива загруженных строк
    'До заполнения сабайтемов
    'ReDim lvItemLoaded(ind) '(UBound(lvItemLoaded) + 1) 'всетаки сверху
    'Заполнить сабайтемы
    If Opt_LoadOnlyTitles = False Then
        'не только названия
        FillLvSubs ind
    End If

    ind = ind + 1
    rs.MoveNext
Loop

err.Clear



'LockWindowUpdate 0
'ListView.Visible = True

'ToDebug timeGetTime() - t '6788? 8200
'err.Clear

TextItemHid.ZOrder 0

'FilteredFlag = False 'флаг неполного показа

If ListView.ListItems.Count > 0 Then
    If CurSearch < 1 Then CurSearch = 1
    If CurSearch > ListView.ListItems.Count Then CurSearch = ListView.ListItems.Count

    ListView.MultiSelect = False
    'If Not GroupedFlag Then
    'ListView.SelectedItem.Selected = False 'не помню ваще зачем снимать

    'Select
    If FirstLVFill Then
        If LastInd > ListView.ListItems.Count - 1 Then LastInd = 0
        Set ListView.SelectedItem = ListView.ListItems(LastInd + 1)
    Else
        If GroupedFlag Then
            'переместить гор-скролл, где был
            ListViewScroll ListView, nScrollPos, 0
            'пометки по ключу тут глючат (выделяется несколько, если в группе выделено по шифту)
        Else
            SelectLVItemFromKey LastKey    'пометить по ключу,если поле видно и переместить гор-скролл, где был
        End If
    End If

    ListView.MultiSelect = True

    'поиск фильма при старте
    If Me.Visible Then
    Else
        'If FirstActivateFlag Then
        If Len(Command$) <> 0 Then
            'найти фильм из командной строки
            temp = GetNameFromPathAndName(Command$)
            'temp = "Instinct.avi"""
            temp = Replace(temp, """", vbNullString)
            If Len(temp) <> 0 Then
                i = SearchLV(dbFileNameInd, temp, vbTextCompare)
                If i > 0 Then
                    ListView.ListItems(ListView.SelectedItem.Index).Checked = True
                    ChMarkFindHid.Value = vbChecked
                    j = FindNextLV(dbFileNameInd, temp)    'пока LV невидим - ищет по всему списку
                    ChMarkFindHid.Value = vbUnchecked
                    TextItemHid.Text = NamesStore(3) & "(" & i + j & ")" & Chr$(32) & temp    'найдено
                End If
                'Else
            End If    'Len(temp)

        End If    'command <> ""
    End If    'при старте

    'если была сортировка - произвести ее (полюбому пусть сортирует в начале)
    If LVSortColl > 0 Then LVSOrt (LVSortColl)
    If LVSortColl = -1 Then SortByCheck 0

    'шаманство: нажать кнопку влево, для кликанья на выделенную строку
    'чтобы был правильный селект с шифт
    Call SendMessage(ListView.hwnd, &H100, vbKeyLeft, ByVal 0&)
    Call SendMessage(ListView.hwnd, &H101, vbKeyLeft, ByVal 0&)    'это же и клик
    ' ^ LVCLICK    ' обновляет картинки при смене баз, надо

    '           там LastKey меняется
    'CheckedInLV = CheckCount
End If    'lvcount>0

FrmMain.FrameView.Caption = FrameViewCaption & " " & FrmMain.ListView.ListItems.Count & " )"    'нестыкуется с группами

'delFlag = False
FirstLVFill = False
Screen.MousePointer = vbNormal

End Sub
Private Sub SelectLVItemFromKey(kl As Long)
'Позиционирование в ListView по заданному ключу
Dim itmX As ListItem
On Error Resume Next

For Each itmX In ListView.ListItems
    If Val(itmX.Key) = kl Then
        itmX.Selected = True
        itmX.EnsureVisible
    End If
Next

'переместить гор-скролл, где был
ListViewScroll ListView, nScrollPos, 0
End Sub


'Public Function AndOr(ch As CheckBox) As String
'If ch = vbChecked Then
'    AndOr = "AND"
'ElseIf ch = vbUnchecked Then
''не ставить or на незаполненное
'If Len(FrmFilter.cbs(ch.Index)) <> 0 Then
'    AndOr = "OR"
'Else
'    AndOr = "AND"
'End If
'
'End If
'End Function


Private Sub ma_cScroll_Change(eBar As EFSScrollBarConstants)
ma_cScroll_Scroll eBar
End Sub

Private Sub ma_cScroll_Scroll(eBar As EFSScrollBarConstants)
If (eBar = efsHorizontal) Then
    PicActFoto.left = -Screen.TwipsPerPixelX * ma_cScroll.Value(eBar)
Else
    PicActFoto.top = -Screen.TwipsPerPixelY * ma_cScroll.Value(eBar)
End If
End Sub
Private Sub m_cscroll_Change(eBar As EFSScrollBarConstants)
m_cscroll_Scroll eBar
End Sub

Private Sub m_cscroll_Scroll(eBar As EFSScrollBarConstants)
If (eBar = efsHorizontal) Then
    PicFaceV.left = -Screen.TwipsPerPixelX * m_cScroll.Value(eBar)
Else
    PicFaceV.top = -Screen.TwipsPerPixelY * m_cScroll.Value(eBar)
End If
End Sub
Private Sub map_cScroll_Change(eBar As EFSScrollBarConstants)
map_cScroll_Scroll eBar
End Sub

Private Sub map_cScroll_Scroll(eBar As EFSScrollBarConstants)
If (eBar = efsHorizontal) Then
    PicCoverPaper.left = -Screen.TwipsPerPixelX * map_cScroll.Value(eBar)
Else
    PicCoverPaper.top = -Screen.TwipsPerPixelY * map_cScroll.Value(eBar)
End If
End Sub

Private Sub picScrollBoxV_Resize()
Dim lHeight As Long
Dim lWidth As Long
Dim lProportion As Long
On Error Resume Next

If Me.Visible = False Then Exit Sub
'If picScrollBoxV.Visible = False Then Exit Sub 'надо выполнять для первого запуска
If m_cScroll Is Nothing Then Exit Sub

lHeight = (PicFaceV.Height - picScrollBoxV.ScaleHeight) \ Screen.TwipsPerPixelY
If (lHeight > 0) Then
    lProportion = lHeight \ (picScrollBoxV.ScaleHeight \ Screen.TwipsPerPixelY) + 1
    m_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
    m_cScroll.Max(efsVertical) = lHeight
    m_cScroll.Visible(efsVertical) = True
Else
    m_cScroll.Visible(efsVertical) = False
    PicFaceV.top = 0
End If

lWidth = (PicFaceV.Width - picScrollBoxV.ScaleWidth) \ Screen.TwipsPerPixelX
If (lWidth > 0) Then
    lProportion = lWidth \ (picScrollBoxV.ScaleWidth \ Screen.TwipsPerPixelX) + 1
    m_cScroll.LargeChange(efsHorizontal) = lWidth \ lProportion
    m_cScroll.Max(efsHorizontal) = lWidth
    m_cScroll.Visible(efsHorizontal) = True
Else
    m_cScroll.Visible(efsHorizontal) = False
    PicFaceV.left = 0
End If

End Sub

Private Sub DisplayTextPopup(ByVal txt As TextBox)
'мое меню
' Make VB discard the mouse capture.
txt.Enabled = False
txt.Enabled = True
txt.SetFocus

' Enable appropriate menu items.
' See if anything is selected.
If txt.SelLength > 0 Then
    ' Text is selected.
    mnuCut.Enabled = True
    mnuCopy.Enabled = True
    mnuDelete.Enabled = True
Else
    ' No text is selected.
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    mnuDelete.Enabled = False
End If

If SendMessage(txt.hwnd, EM_CANUNDO, 0&, ByVal 0&) Then mnuUndo.Enabled = True Else mnuUndo.Enabled = False

' Display the custom menu.

m7Hid.Visible = True
mnuUndo.Visible = True
mnuCut.Visible = True
mnuCopy.Visible = True
mnuPaste.Visible = True
mnuDelete.Visible = True
m8Hid.Visible = True
mnuSelectAll.Visible = True

If Len(sPerson) = 0 Then mFiltAct.Enabled = False Else mFiltAct.Enabled = True

PopupMenu mnuTextMenuHid
End Sub

Private Sub mnuCopy_Click()
Dim sOldLang As String
On Error Resume Next

sOldLang = switchLang("00000419")

Clipboard.Clear
Clipboard.SetText ActiveControl.SelText
If Len(sOldLang) > 0 Then sOldLang = switchLang(sOldLang)

End Sub

' Cut the current TextBox's text
' to the clipboard.
Private Sub mnuCut_Click()
Dim sOldLang As String
sOldLang = switchLang("00000419")
Clipboard.Clear
Clipboard.SetText ActiveControl.SelText
ActiveControl.SelText = vbNullString
If Len(sOldLang) > 0 Then sOldLang = switchLang(sOldLang)
End Sub


' Delete the selected text.
Private Sub mnuDelete_Click()
ActiveControl.SelText = vbNullString
End Sub

' Paste into the current TextBox.
Private Sub mnuPaste_Click()
ActiveControl.SelText = Clipboard.GetText
End Sub

' Select the TextBox's text.
Private Sub mnuSelectAll_Click()
ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub mnuSavePic_Click()
'сохранение в файл картинки из 1:1 окна
'не из базы

SavePicFromPic FormShowPic.Picture, FormShowPic.hwnd

'Dim iFile As String
'iFile = pSaveDialog(FormShowPic.hwnd, DTitle:=ComSaveRec.Caption)
'If iFile <> vbNullString Then
'If LCase$(Right$(iFile, 3)) = "bmp" Then
'    SavePicture FormShowPic.Picture, iFile
'End If 'bmp
'If LCase$(Right$(iFile, 3)) = "jpg" Then
'    m_cDib.CreateFromPicture FormShowPic.Picture
'    SaveJPG m_cDib, iFile, QJPG
'End If 'jpg
'Set m_cDib = Nothing
'End If ' iFile <> vbNullString

End Sub
Private Sub mnuCopyPic_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetData FormShowPic.Picture    ', vbCFBitmap
End Sub

Private Sub mnuKillPic_Click()
If IsCoverShowFlag Then
    CoverWindTop = FormShowPic.top
    CoverWindLeft = FormShowPic.left
Else
    ScrShotWindTop = FormShowPic.top
    ScrShotWindLeft = FormShowPic.left
End If
Unload FormShowPic: Set FormShowPic = Nothing
End Sub
Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
m_emr = RHS
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
m_emr = emrConsume
ISubclass_MsgResponse = m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'ограничение ресайза
Dim mmiT As MINMAXINFO
'resize
' Copy parameter to local variable for processing
CopyMemory mmiT, ByVal lParam, LenB(mmiT)

' Minimium width and height for sizing
mmiT.ptMinTrackSize.X = 807
mmiT.ptMinTrackSize.Y = 608

' Copy modified results back to parameter
CopyMemory ByVal lParam, mmiT, LenB(mmiT)

End Function
Private Sub FillTVNums()
'заполнить, в скольких позициях участвует (потрошенный актер например)
Dim itmX As ListItem
On Error Resume Next
Screen.MousePointer = vbHourglass

PBar.min = 0: PBar.Value = 0
If tvGroup.ListItems.Count > 0 Then PBar.Max = tvGroup.ListItems.Count + 1 Else PBar.Max = 1
PBar.ZOrder 0

For Each itmX In tvGroup.ListItems
    'DoEvents
    If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then Exit For
    'Debug.Print GetKeyState(vbKeyEscape)
    If GetKeyState(vbKeyEscape) < 0 Then Exit For
    'End If
    'Exit For
    'End If

    PBar.Value = PBar.Value + 1
    'itmX.ListSubItems.Add Text:=GetGroupNum(itmX.Text)
    'сабы вбили при заполнении пустыми
    itmX.ListSubItems(1) = GetGroupNum(itmX.Text)
    'Debug.Print itmX.ListSubItems(1)
Next

TextItemHid.ZOrder 0
Screen.MousePointer = vbNormal
End Sub


Public Sub SetListboxScrollbar(lB As ListBox, frm As Form)
'добавление горизонтальной прокрутки listbox
Dim i As Integer
Dim new_len As Long
Dim max_len As Long

'у frmmain теперь TextWidth зависит от фонта


For i = 0 To lB.ListCount - 1
    new_len = 10 + ScaleX(frm.TextWidth(lB.List(i)), frm.ScaleMode, vbPixels)
    If max_len < new_len Then max_len = new_len
Next i

SendMessage lB.hwnd, LB_SETHORIZONTALEXTENT, max_len, 0
End Sub

Private Sub m_LV_Vert_MouseLeave()
'End tracking
Screen.MousePointer = vbNormal
End Sub

Private Sub FrLV_Vert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Tracking is initialised by entering the control:
FrLV_Vert.MousePointer = 9
If Not (m_LV_Vert.Tracking) Then m_LV_Vert.StartMouseTracking
End Sub
Private Sub FrTV_Vert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Tracking is initialised by entering the control:
FrTV_Vert.MousePointer = 9
If Not (m_TV_Vert.Tracking) Then m_TV_Vert.StartMouseTracking
End Sub

Private Sub FrLV_Vert_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'двигаем вертик сплиттер
If Opt_Group_Vis Then
    LVWidth = ((X + ListView.Width) * 100) / (MainWidth - tvGroup.Width)
Else
    LVWidth = ((X + ListView.Width) * 100) / MainWidth
End If
'DoEvents
Form_Resize
End Sub
Private Sub FrTV_Vert_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'двигаем вертик сплиттер
TVWidth = X + tvGroup.Width
If TVWidth > MainWidth - 5000 Then
    TVWidth = MainWidth - 5000
End If

If TVWidth < 2100 Then TVWidth = 2100

On Error Resume Next    'если TVWidth меньше(уже не надо)
tvGroup.ColumnHeaders(1).Width = TVWidth - tvGroup.ColumnHeaders(2).Width - 260    '+ setforecolor
'DoEvents
Form_Resize

End Sub




Private Sub SetMenuIcon()
'Dim M As Menu
Set cIM = New cIconMenu

cIM.Attach FrmMain.hwnd, FrmMain.ScaleMode
cIM.OfficeXpStyle = True
cIM.ImageList = ImageList

'On Error Resume Next
'For Each m In mnuEdit
'cIM.IconIndex(m.Caption) = ImageList1.ListImages(m.Caption).Index - 1
'Next
'err.Clear
'On Error GoTo 0

End Sub

Public Sub LVScroll()
'Если находимся в режиме редактирования
If mblnEditing Then
    'то прекращаем его
    txtEdit_LostFocus
End If
End Sub
Private Sub txtEdit_Change()
Dim tmps As Long    'Single 'tb width all
Dim tdw_optim As Long    'Single 'tb width в рамках lv
Dim nRows As Integer

On Error GoTo err
'вызывается из ListViewEdit для коррекции
'Изменяем ширину текстбокса в зависимости от ширины текста
If mblnEditing Then
    tdw_optim = (ListView.left + ListView.Width - txtEdit.left - 20)
    tmps = (Me.TextWidth(txtEdit.Text) + 20) * Screen.TwipsPerPixelX
    'не надо за правый край лезть
    If tmps > tdw_optim Then

        'дать еще строк
        nRows = tmps \ tdw_optim + 1
        'Debug.Print nRows
        txtEdit.Height = nRows * mlngTBoxH + mlngTBoxH / 2

        tmps = tdw_optim
        ' Debug.Print txtEdit.Height / Screen.TwipsPerPixelY, txtEdit.Font.Size

    Else
        'одна строка
        txtEdit.Height = mlngTBoxH
    End If

    txtEdit.Width = tmps

    'Debug.Print txtEdit.Width
End If

Exit Sub
err:
'Debug.Print "err_che"
End Sub

Private Sub txtEdit_GotFocus()
'Выделям текст
txtEdit.SelStart = 0: txtEdit.SelLength = Len(txtEdit.Text)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
'отредактированную ячейку
If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    txtEdit_LostFocus
ElseIf KeyAscii = vbKeyReturn Then
    'записать
    Call PutLVED2Base
    KeyAscii = 0
End If
End Sub

Private Sub txtEdit_LostFocus()
'Скрываем текстбокс редактирования lv
txtEdit.Visible = False
'Убираем флаг
mblnEditing = False
txtEdit.Text = vbNullString
End Sub
Private Sub PutLVED2Base()
'поместить отредактированную ячейку lv в базу

On Error GoTo err
If Not mblnEditing Then Exit Sub

'переписать ячейку в базу
RSGoto ListView.ListItems(mlngIndex).Key
rs.Edit

If mlngSubIndex > 0 Then    'из саба
    rs.Fields(mlngSubIndex) = txtEdit.Text
    'Debug.Print ListView.ColumnHeaders(lvSubIndex + 1)
Else
    'первую
    rs.Fields(dbMovieNameInd) = txtEdit.Text
    'Debug.Print ListView.ColumnHeaders(1)
End If

rs.Update

'считать строку с базы
If mlngSubIndex > 0 Then
    FillLvSubs mlngIndex
Else
    'подгрузить из базы название
    ListView.ListItems(mlngIndex).Text = CheckNoNullVal(dbMovieNameInd)
End If
'    ListView.ListItems(mlngIndex).SubItems(mlngSubIndex) = txtEdit.Text
'    ListView.ListItems(mlngIndex).Text = txtEdit.Text

RestoreBasePos

If Opt_SortLVAfterEdit Then
    'если была сортировка - произвести ее
    If LVSortColl > 0 Then LVSOrt (LVSortColl)
    If LVSortColl = -1 Then SortByCheck 0, True
End If

ListView.SetFocus    'txtEdit_LostFocus


Exit Sub
err:
ToDebug "Err_plved:" & err.Description
'Debug.Print "Err_plved:" & err.Description

End Sub
Public Sub ListViewEdit(lvList As ListView)

Dim ret As Long
Dim hLV As Long
Dim pPos As POINTAPI
Dim hitInfo As LVHITTESTINFO
Dim rectLabel As RECT    ', rectIcon As RECT
'Dim tmps As Single
Dim k As Integer
If lvList.View <> lvwReport Then Exit Sub

On Error GoTo err

' for Me.ScaleMode = 3

hLV = lvList.hwnd

'Получаем координаты курсора
ret = GetMessagePos()
'И пишем их в структуру
pPos.X = LoWord(ret)
pPos.Y = HIWORD(ret)

'Пересчитываемся относительно LV
If ScreenToClient(hLV, pPos) = 0 Then Exit Sub

'Используем в структуре пересчитанные координаты
hitInfo.pt = pPos
'Сотрим куда мы "попали"
ret = SendMessage(hLV, LVM_SUBITEMHITTEST, 0&, hitInfo)
'Если неудачно "посмотрели", то вывалимся
If ret = -1 Then Exit Sub

'iIndex - является 0-based, поэтому +1 чтобы работать с коллекцией ListItems
mlngIndex = hitInfo.iItem + 1
mlngSubIndex = hitInfo.iSubitem

'не редактировать индекс
If ListView.ColumnHeaders(lvIndexPole).Index = mlngSubIndex Then Exit Sub

lvList.ListItems(mlngIndex).EnsureVisible

rectLabel.top = mlngSubIndex
rectLabel.left = LVIR_LABEL
If SendMessage(hLV, LVM_GETSUBITEMRECT, hitInfo.iItem, rectLabel) = 0 Then Exit Sub

'    rectIcon.Top = hitInfo.iSubitem
'    rectIcon.Left = LVIR_ICON
'    If SendMessage(hLV, LVM_GETSUBITEMRECT, hitInfo.iItem, rectIcon) = 0 Then Exit Sub

'Выставляем флаг редактирования
mblnEditing = True
With txtEdit
    '        'Заполняем текстбох в зависимости от того, что выбрано
    '        If mlngSubIndex > 0 Then
    '            .Text = lvList.ListItems(mlngIndex).SubItems(mlngSubIndex)
    '        Else
    '            .Text = lvList.ListItems(mlngIndex).Text
    '        End If

    'rectLabelLeft = rectLabel.Left
    'а не выходит...
    'Изменяем ширину текстбокса в зависимости от ширины текста
    '(Me.TextWidth(txtEdit.Text) + 20) * Screen.TwipsPerPixelX

    '.Width = (Me.TextWidth(.Text) + 20) '* Screen.TwipsPerPixelX



    'If tmps > lvList.Width - rectLabel.Left - 20 Then tmps = lvList.Width - rectLabel.Left - 20
    '.Width = tmps


    'Отображаем текстбокс
    '.Visible = True

    'Перемещаем текстбох в нужное место
    '(rectIcon.Right - rectIcon.Left) - IIf(mlngSubIndex > 0, 5, 10)
    '+ (rectIcon.Right - rectIcon.Left)
    '6 / Me.Font.Size
    Select Case Me.Font.Size
    Case 7.5 To 8.25: k = 1
    Case 6: k = 3
    Case 6.75: k = 2
    Case 9
        If Me.Font.name = "Verdana" Then k = 1
    Case Else: k = 0
    End Select

    If rectLabel.left < 0 Then rectLabel.left = 0    'не надо за левый край лезть

    MoveWindow txtEdit.hwnd, _
               rectLabel.left + IIf(mlngSubIndex > 0, 5, 1), _
               rectLabel.top + k, _
               txtEdit.Width, _
               rectLabel.bottom - rectLabel.top, _
               1&

    mlngTBoxH = .Height
    'rectLabel.Bottom - rectLabel.Top = 17
    '.height=255//screen.TwipsPerPixely = 17
    'Debug.Print .Height

    'коррекция местоположения в зависимости от положения листа на форме
    .left = (.left + (lvList.left))
    .top = (.top + (lvList.top))

    .Width = 0
    'ширину правим
    'Заполняем текстбох в зависимости от того, что выбрано
    If mlngSubIndex > 0 Then
        .Text = lvList.ListItems(mlngIndex).SubItems(mlngSubIndex)
    Else
        .Text = lvList.ListItems(mlngIndex).Text
    End If
    If Len(.Text) = 0 Then Call txtEdit_Change

    'И передаем в него фокус
    .Visible = True
    .SetFocus
End With

Exit Sub
err:
ToDebug "Err_lved:" & err.Description
'Debug.Print "Err_lved:" & err.Description
End Sub

