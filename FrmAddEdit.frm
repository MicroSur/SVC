VERSION 5.00
Begin VB.Form FrmAddEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrImgPrCov 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   8940
      TabIndex        =   1
      Top             =   3120
      Width           =   1275
      Begin VB.CommandButton ComX 
         Caption         =   "X"
         Height          =   255
         Index           =   4
         Left            =   1020
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   60
         Width           =   255
      End
      Begin VB.Image ImgPrCov 
         Height          =   1335
         Left            =   0
         MouseIcon       =   "FrmAddEdit.frx":0000
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1215
      End
   End
   Begin SurVideoCatalog.XpB ComDel 
      Height          =   375
      Left            =   8580
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   688
      Caption         =   "Del"
      ButtonStyle     =   3
      Picture         =   "FrmAddEdit.frx":08CA
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB ComOpen 
      Height          =   375
      Left            =   8580
      TabIndex        =   4
      Top             =   540
      Width           =   1935
      _ExtentX        =   265
      _ExtentY        =   265
      Caption         =   "Open"
      ButtonStyle     =   3
      Picture         =   "FrmAddEdit.frx":12DC
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      MaskColor       =   16711935
   End
   Begin SurVideoCatalog.XpB ComSaveRec 
      Height          =   375
      Left            =   8580
      TabIndex        =   5
      Top             =   2100
      Width           =   1935
      _ExtentX        =   265
      _ExtentY        =   265
      Caption         =   "Save"
      ButtonStyle     =   3
      Picture         =   "FrmAddEdit.frx":1630
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB ComAdd 
      Height          =   375
      Left            =   8580
      TabIndex        =   6
      Top             =   1020
      Width           =   1935
      _ExtentX        =   265
      _ExtentY        =   265
      Caption         =   "Add"
      ButtonStyle     =   3
      Picture         =   "FrmAddEdit.frx":1BCA
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin SurVideoCatalog.XpB ComCancel 
      Height          =   375
      Left            =   8580
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
      _ExtentX        =   265
      _ExtentY        =   265
      Caption         =   "Cancel"
      ButtonStyle     =   3
      PictureWidth    =   16
      PictureHeight   =   16
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin VB.Frame FrAdEdTechHid 
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   120
      TabIndex        =   90
      Top             =   420
      Width           =   10455
      Begin VB.TextBox TextFileName 
         Height          =   315
         Left            =   1380
         ScrollBars      =   2  'Vertical
         TabIndex        =   104
         Top             =   3540
         Width           =   6975
      End
      Begin VB.TextBox TextCDN 
         Height          =   315
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   103
         Top             =   120
         Width           =   2835
      End
      Begin VB.TextBox TextTimeHid 
         Height          =   315
         Left            =   1380
         TabIndex        =   102
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox TextResolHid 
         Height          =   315
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   101
         Top             =   2040
         Width           =   6975
      End
      Begin VB.TextBox TextAudioHid 
         Height          =   315
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   100
         Top             =   3000
         Width           =   6975
      End
      Begin VB.TextBox TextFPSHid 
         Height          =   315
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   99
         Top             =   2460
         Width           =   6975
      End
      Begin VB.TextBox TextFilelenHid 
         Height          =   315
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   98
         Top             =   3960
         Width           =   6975
      End
      Begin VB.TextBox TextVideoHid 
         Height          =   315
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   97
         Top             =   1620
         Width           =   6975
      End
      Begin VB.TextBox CDSerialCur 
         Height          =   315
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   96
         Top             =   540
         Width           =   6975
      End
      Begin VB.ComboBox TextUser 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":25DC
         Left            =   1380
         List            =   "FrmAddEdit.frx":25DE
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   95
         Top             =   5580
         Width           =   6975
      End
      Begin VB.ComboBox ComboNos 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":25E0
         Left            =   4260
         List            =   "FrmAddEdit.frx":25E2
         Sorted          =   -1  'True
         TabIndex        =   94
         Top             =   120
         Width           =   4095
      End
      Begin VB.TextBox TextCoverURL 
         Height          =   315
         Left            =   4320
         MaxLength       =   255
         TabIndex        =   93
         Top             =   4500
         Width           =   4035
      End
      Begin VB.TextBox TextMovURL 
         Height          =   315
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   92
         Top             =   4980
         Width           =   6975
      End
      Begin VB.ComboBox cBasePicURL 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":25E4
         Left            =   1380
         List            =   "FrmAddEdit.frx":25E6
         Sorted          =   -1  'True
         TabIndex        =   91
         Top             =   4500
         Width           =   2895
      End
      Begin SurVideoCatalog.XpB ComInterGoHid 
         Height          =   375
         Index           =   1
         Left            =   8460
         TabIndex        =   105
         Top             =   4920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   ""
         ButtonStyle     =   3
         Picture         =   "FrmAddEdit.frx":25E8
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
         MaskColor       =   16711935
      End
      Begin SurVideoCatalog.XpB ComGetCover 
         Height          =   375
         Left            =   8460
         TabIndex        =   106
         Top             =   4470
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Get"
         ButtonStyle     =   3
         Picture         =   "FrmAddEdit.frx":293C
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
         MaskColor       =   16711935
      End
      Begin SurVideoCatalog.XpB ComPlusHid 
         Height          =   300
         Index           =   2
         Left            =   7980
         TabIndex        =   107
         Top             =   1080
         Width           =   315
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "+"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.Label LTech 
         Caption         =   "File"
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   119
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label LTech 
         Caption         =   "NNCD"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   118
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label LTech 
         Caption         =   "Length"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   117
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label LTech 
         Caption         =   "Resol"
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   116
         Top             =   2100
         Width           =   1275
      End
      Begin VB.Label LTech 
         Caption         =   "Audio"
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   115
         Top             =   3060
         Width           =   1275
      End
      Begin VB.Label LTech 
         Caption         =   "FPS"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   114
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label LTech 
         Caption         =   "Size"
         Height          =   255
         Index           =   8
         Left            =   60
         TabIndex        =   113
         Top             =   4020
         Width           =   1275
      End
      Begin VB.Label LTech 
         Caption         =   "Video"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   112
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label LTech 
         Caption         =   "SN"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   111
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label LTech 
         Caption         =   "Debtor"
         Height          =   255
         Index           =   11
         Left            =   60
         TabIndex        =   110
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label LTech 
         Caption         =   "CoverURL"
         Height          =   255
         Index           =   9
         Left            =   60
         TabIndex        =   109
         Top             =   4560
         Width           =   1275
      End
      Begin VB.Label LTech 
         Caption         =   "MovieURL"
         Height          =   255
         Index           =   10
         Left            =   60
         TabIndex        =   108
         Top             =   5040
         Width           =   1275
      End
   End
   Begin VB.Frame FrAdEdPixHid 
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   120
      TabIndex        =   51
      Top             =   420
      Width           =   10455
      Begin VB.OptionButton optAspect 
         Appearance      =   0  'Flat
         Caption         =   "w:h"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   8700
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox PicSS3Big 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7560
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   4380
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.PictureBox PicSS2Big 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4320
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   4440
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.PictureBox PicSS1Big 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   4440
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.PictureBox PicSS1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   60
         MouseIcon       =   "FrmAddEdit.frx":334E
         MousePointer    =   99  'Custom
         ScaleHeight     =   185
         ScaleMode       =   0  'User
         ScaleWidth      =   224
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   4620
         Width           =   3360
         Begin VB.PictureBox Picture1 
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   0
            TabIndex        =   79
            Top             =   0
            Width           =   0
         End
         Begin VB.CommandButton ComX 
            Caption         =   "X"
            Height          =   315
            Index           =   0
            Left            =   0
            MousePointer    =   1  'Arrow
            TabIndex        =   78
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.CheckBox ChLockFHid 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   76
         Top             =   3720
         Width           =   195
      End
      Begin VB.CheckBox ChLockSHid 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   0
         TabIndex        =   75
         Top             =   4080
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.PictureBox picCanvas 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3525
         Left            =   0
         MouseIcon       =   "FrmAddEdit.frx":3C18
         MousePointer    =   99  'Custom
         ScaleHeight     =   235
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   239
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   60
         Width           =   3585
         Begin VB.CommandButton ComFrontFaceFile 
            Height          =   375
            Index           =   0
            Left            =   2760
            MousePointer    =   1  'Arrow
            Picture         =   "FrmAddEdit.frx":44E2
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   0
            Width           =   435
         End
         Begin VB.CommandButton ComFrontFace 
            Height          =   375
            Index           =   0
            Left            =   3180
            MousePointer    =   1  'Arrow
            Picture         =   "FrmAddEdit.frx":4A6C
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   0
            Width           =   435
         End
         Begin VB.CommandButton ComX 
            Caption         =   "X"
            Height          =   315
            Index           =   3
            Left            =   0
            MousePointer    =   1  'Arrow
            TabIndex        =   72
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.PictureBox PicFrontFace 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   180
         MousePointer    =   99  'Custom
         ScaleHeight     =   44
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   44
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CommandButton ComKeyNext 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10020
         Picture         =   "FrmAddEdit.frx":4FF6
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3660
         Width           =   375
      End
      Begin VB.CommandButton ComKeyPrev 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10020
         Picture         =   "FrmAddEdit.frx":5760
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   4020
         Width           =   375
      End
      Begin VB.CommandButton ComFrontFaceFile 
         Height          =   375
         Index           =   1
         Left            =   1980
         MousePointer    =   1  'Arrow
         Picture         =   "FrmAddEdit.frx":5ECA
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   4500
         Width           =   435
      End
      Begin VB.CommandButton ComFrontFace 
         Height          =   375
         Index           =   1
         Left            =   2400
         MousePointer    =   1  'Arrow
         Picture         =   "FrmAddEdit.frx":6454
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   4500
         Width           =   435
      End
      Begin VB.CommandButton ComCap 
         Height          =   495
         Index           =   0
         Left            =   2820
         MousePointer    =   1  'Arrow
         Picture         =   "FrmAddEdit.frx":69DE
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   4500
         Width           =   615
      End
      Begin VB.PictureBox movie 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3525
         Left            =   3660
         ScaleHeight     =   3525
         ScaleWidth      =   4680
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   60
         Width           =   4680
      End
      Begin VB.PictureBox PicSS2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   3540
         MouseIcon       =   "FrmAddEdit.frx":6CE8
         MousePointer    =   99  'Custom
         ScaleHeight     =   185
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   4620
         Width           =   3360
         Begin VB.CommandButton ComX 
            Caption         =   "X"
            Height          =   315
            Index           =   1
            Left            =   0
            MousePointer    =   1  'Arrow
            TabIndex        =   63
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.CommandButton ComFrontFaceFile 
         Height          =   375
         Index           =   2
         Left            =   5460
         MousePointer    =   1  'Arrow
         Picture         =   "FrmAddEdit.frx":75B2
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   4500
         Width           =   435
      End
      Begin VB.CommandButton ComFrontFace 
         Height          =   375
         Index           =   2
         Left            =   5880
         MousePointer    =   1  'Arrow
         Picture         =   "FrmAddEdit.frx":7B3C
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4500
         Width           =   435
      End
      Begin VB.CommandButton ComCap 
         Height          =   495
         Index           =   1
         Left            =   6300
         MousePointer    =   1  'Arrow
         Picture         =   "FrmAddEdit.frx":80C6
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   4500
         Width           =   615
      End
      Begin VB.PictureBox PicSS3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   7020
         MouseIcon       =   "FrmAddEdit.frx":83D0
         MousePointer    =   99  'Custom
         ScaleHeight     =   185
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   4620
         Width           =   3360
         Begin VB.CommandButton ComX 
            Caption         =   "X"
            Height          =   315
            Index           =   2
            Left            =   0
            MousePointer    =   1  'Arrow
            TabIndex        =   58
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.CommandButton ComFrontFaceFile 
         Height          =   375
         Index           =   3
         Left            =   8940
         MousePointer    =   1  'Arrow
         Picture         =   "FrmAddEdit.frx":8C9A
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   4500
         Width           =   435
      End
      Begin VB.CommandButton ComFrontFace 
         Height          =   375
         Index           =   3
         Left            =   9360
         MousePointer    =   1  'Arrow
         Picture         =   "FrmAddEdit.frx":9224
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   4500
         Width           =   435
      End
      Begin VB.CommandButton ComCap 
         Height          =   495
         Index           =   2
         Left            =   9780
         MousePointer    =   1  'Arrow
         Picture         =   "FrmAddEdit.frx":97AE
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4500
         Width           =   615
      End
      Begin VB.OptionButton optAspect 
         Appearance      =   0  'Flat
         Caption         =   "4:3"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2760
         Width           =   495
      End
      Begin VB.OptionButton optAspect 
         Appearance      =   0  'Flat
         Caption         =   "16:9"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   9660
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2760
         Width           =   495
      End
      Begin MSComctlLib.Slider PositionP 
         Height          =   375
         Left            =   240
         TabIndex        =   84
         Top             =   4080
         Width           =   9735
         _ExtentX        =   17198
         _ExtentY        =   688
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   1
         TickStyle       =   3
         TickFrequency   =   10
         Value           =   1
      End
      Begin MSComctlLib.Slider Position 
         Height          =   375
         Left            =   240
         TabIndex        =   85
         Top             =   3660
         Width           =   9735
         _ExtentX        =   17198
         _ExtentY        =   688
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   1
         TickStyle       =   3
         TickFrequency   =   10
         Value           =   1
      End
      Begin SurVideoCatalog.XpB ComRND 
         Height          =   255
         Index           =   2
         Left            =   7320
         TabIndex        =   86
         Top             =   7440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         Caption         =   "Auto"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComRND 
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   87
         Top             =   7440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         Caption         =   "Auto"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComRND 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   88
         Top             =   7440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         Caption         =   "Auto"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComAutoScrShots 
         Height          =   435
         Left            =   8460
         TabIndex        =   89
         Top             =   3120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   767
         Caption         =   "Random"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
   End
   Begin VB.Frame FrAdEdTextHid 
      BorderStyle     =   0  'None
      Height          =   7755
      Left            =   120
      TabIndex        =   8
      Top             =   420
      Width           =   10455
      Begin VB.ComboBox ComboOther 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9AB8
         Left            =   5400
         List            =   "FrmAddEdit.frx":9ABA
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   3420
         Width           =   2535
      End
      Begin VB.CheckBox ChInFilFl 
         Alignment       =   1  'Right Justify
         Caption         =   "Empty"
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   3840
         Value           =   1  'Checked
         Width           =   4815
      End
      Begin VB.ComboBox TextRate 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9ABC
         Left            =   1380
         List            =   "FrmAddEdit.frx":9ABE
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   30
         Top             =   3060
         Width           =   2775
      End
      Begin VB.ComboBox TextSubt 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9AC0
         Left            =   5400
         List            =   "FrmAddEdit.frx":9AC2
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   29
         Top             =   3060
         Width           =   2955
      End
      Begin VB.ComboBox TextLang 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9AC4
         Left            =   5400
         List            =   "FrmAddEdit.frx":9AC6
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   28
         Top             =   2700
         Width           =   2955
      End
      Begin VB.TextBox TextRole 
         Height          =   735
         Left            =   1380
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   1920
         Width           =   6975
      End
      Begin VB.ComboBox TextMName 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9AC8
         Left            =   1380
         List            =   "FrmAddEdit.frx":9ACA
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   26
         Top             =   120
         Width           =   6975
      End
      Begin VB.ComboBox TextCountry 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9ACC
         Left            =   1380
         List            =   "FrmAddEdit.frx":9ACE
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   25
         Top             =   1200
         Width           =   3975
      End
      Begin VB.ComboBox TextOther 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9AD0
         Left            =   1380
         List            =   "FrmAddEdit.frx":9AD2
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   24
         Top             =   3420
         Width           =   3975
      End
      Begin VB.Frame FrInetInfo 
         Caption         =   "IFind"
         Height          =   3075
         Left            =   5040
         TabIndex        =   17
         Top             =   4140
         Width           =   5415
         Begin VB.CommandButton ComRHid 
            Caption         =   "G"
            Height          =   315
            Index           =   1
            Left            =   4980
            TabIndex        =   21
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox TxtIName 
            Height          =   315
            Left            =   1200
            TabIndex        =   20
            Top             =   240
            Width           =   3735
         End
         Begin VB.ListBox lbInetMovieList 
            Height          =   1815
            ItemData        =   "FrmAddEdit.frx":9AD4
            Left            =   120
            List            =   "FrmAddEdit.frx":9AD6
            TabIndex        =   19
            Top             =   1080
            Width           =   5175
         End
         Begin VB.ComboBox ComboInfoSites 
            Height          =   315
            ItemData        =   "FrmAddEdit.frx":9AD8
            Left            =   105
            List            =   "FrmAddEdit.frx":9ADA
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   660
            Width           =   3615
         End
         Begin SurVideoCatalog.XpB ComInetFind 
            Height          =   315
            Left            =   3840
            TabIndex        =   22
            Top             =   660
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "Find"
            ButtonStyle     =   3
            Picture         =   "FrmAddEdit.frx":9ADC
            PictureWidth    =   16
            PictureHeight   =   16
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
         End
         Begin VB.Label LIName 
            Caption         =   "Title"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ComboBox TextYear 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9C36
         Left            =   1380
         List            =   "FrmAddEdit.frx":9C38
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   16
         Top             =   2700
         Width           =   2775
      End
      Begin VB.ComboBox TextAuthor 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9C3A
         Left            =   1380
         List            =   "FrmAddEdit.frx":9C3C
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   15
         Top             =   1560
         Width           =   6975
      End
      Begin VB.ComboBox TextGenre 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9C3E
         Left            =   1380
         List            =   "FrmAddEdit.frx":9C40
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   14
         Top             =   840
         Width           =   3975
      End
      Begin VB.ComboBox TextLabel 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9C42
         Left            =   1380
         List            =   "FrmAddEdit.frx":9C44
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   13
         Top             =   480
         Width           =   6975
      End
      Begin VB.ComboBox ComboSites 
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   7320
         Width           =   5475
      End
      Begin VB.ComboBox ComboCountry 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9C46
         Left            =   5400
         List            =   "FrmAddEdit.frx":9C48
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox TextAnnotation 
         Height          =   3015
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   4200
         Width           =   4890
      End
      Begin VB.ComboBox ComboGenre 
         Height          =   315
         ItemData        =   "FrmAddEdit.frx":9C4A
         Left            =   5400
         List            =   "FrmAddEdit.frx":9C4C
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   2535
      End
      Begin SurVideoCatalog.XpB ComClsEd 
         Height          =   315
         Left            =   1380
         TabIndex        =   33
         Top             =   3780
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         Caption         =   "Clear"
         PicturePosition =   0
         ButtonStyle     =   3
         PictureWidth    =   0
         PictureHeight   =   0
         ShowFocusRect   =   0   'False
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
         MaskColor       =   16777215
      End
      Begin SurVideoCatalog.XpB ComShowBin 
         Height          =   435
         Left            =   6300
         TabIndex        =   34
         Top             =   7260
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   767
         Caption         =   "drag-and-drop"
         ButtonStyle     =   3
         Picture         =   "FrmAddEdit.frx":9C4E
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComInterGoHid 
         Height          =   435
         Index           =   0
         Left            =   5640
         TabIndex        =   35
         Top             =   7260
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   767
         Caption         =   ""
         ButtonStyle     =   3
         Picture         =   "FrmAddEdit.frx":A1E8
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
         MaskColor       =   16711935
      End
      Begin SurVideoCatalog.XpB ComPlusHid 
         Height          =   300
         Index           =   0
         Left            =   8040
         TabIndex        =   36
         Top             =   840
         Width           =   315
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "+"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComPlusHid 
         Height          =   300
         Index           =   1
         Left            =   8040
         TabIndex        =   37
         Top             =   1200
         Width           =   315
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "+"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComPlusHid 
         Height          =   300
         Index           =   3
         Left            =   8040
         TabIndex        =   38
         Top             =   3420
         Width           =   315
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "+"
         ButtonStyle     =   3
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.Label LSubt 
         Alignment       =   1  'Right Justify
         Caption         =   "Subtitle"
         Height          =   255
         Left            =   4260
         TabIndex        =   50
         Top             =   3120
         Width           =   1035
      End
      Begin VB.Label LLang 
         Alignment       =   1  'Right Justify
         Caption         =   "Language"
         Height          =   255
         Left            =   4260
         TabIndex        =   49
         Top             =   2760
         Width           =   1035
      End
      Begin VB.Label LRate 
         Caption         =   "Rating"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label LOther 
         Caption         =   "Comments"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label LYear 
         Caption         =   "Year"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   1155
      End
      Begin VB.Label LAct 
         Caption         =   "Actors"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label LCountry 
         Caption         =   "Country"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label LRes 
         Caption         =   "Director"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label LAnnot 
         Caption         =   "Descr"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label LLabel 
         Caption         =   "Label"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label LMName 
         Caption         =   "Title"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label LGenre 
         Caption         =   "Genre"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   900
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip TabStrAdEd 
      Height          =   8235
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   14526
      TabWidthStyle   =   2
      TabFixedWidth   =   5997
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Video"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tech"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Set the list box's horizontal extent
'добавление горизонтальной прокрутки listbox, вещь местная для формы
Public Sub SetListboxScrollbar(lb As ListBox)
Dim i As Integer
Dim new_len As Long
Dim max_len As Long

For i = 0 To lb.ListCount - 1
 new_len = 10 + ScaleX(TextWidth(lb.List(i)), ScaleMode, vbPixels)
 If max_len < new_len Then max_len = new_len
Next i

SendMessage lb.hWnd, LB_SETHORIZONTALEXTENT, max_len, 0
End Sub
Private Sub EdAspect2Buttons()
'не должно помечать , если иной формат
Dim i As Integer

If Len(MMI_Format_str) = 0 Then
    If MMI_Format = 1.333 Then
        MMI_Format_str = "4/3"
    ElseIf MMI_Format = 1.777 Then
        MMI_Format_str = "16/9"
        '    ElseIf MMI_Format = 1 Then
        '        MMI_Format_str = "1/1"
    Else
        MMI_Format_str = "4/3"
    End If
End If

'подгонять под mmi_format

For i = 0 To 2: optAspect(i).Value = False: Next i

ChangeFromCode_optAspect = True
    Select Case Trim$(MMI_Format_str)
    Case "4/3", "4:3"
        optAspect(0).Value = True
    Case "16/9", "16:9"
        optAspect(1).Value = True
        'Case "1/1", "1:1"
        '    optAspect(2).Value = True
    End Select
    
    If Not Opt_UseAspect Then    'если не юзать аспект, то не показывать кнопки
        'optAspect(2).Value = True 'w:h
        For i = 0 To 2: optAspect(i).Enabled = False: Next i
    Else
        For i = 0 To 2: optAspect(i).Enabled = True: Next i
    End If
ChangeFromCode_optAspect = False

End Sub

Public Sub Mark2Save()
If rs.RecordCount < 1 Then Exit Sub
If BaseReadOnly Or BaseReadOnlyU Then Exit Sub

If Mark2SaveFlag Then
'    If FrameAddEdit.Visible Then
        If rs.EditMode = 0 Then
        rs.Edit
        ComSaveRec.BackColor = &HC0C0E0 'покраснить
        End If
'    End If
End If
End Sub
Private Function RenderMPV2() As Boolean
'вставка в граф
'"Fraunhofer Video Decoder" (dvdvideo.ax из фри кодека) + audio
'"MPEG-2 Splitter" (mpg2splt.ax)

Dim objRegFilterInfo As IRegFilterInfo
Dim objFilterInfo As IFilterInfo
Dim GraphFilter As IFilterInfo

Dim objPin As IPinInfo
Dim SourceFilter As Boolean    'удачно ли прошел первый фильтр Universal Open Source MPEG Source"


On Error Resume Next    'надо

'If False Then

If MpegMediaOpen Then Call MpegMediaClose
Set mobjManager = New FilgraphManager



'впихнуть в граф наш декодер
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrVDName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    ToDebug "Error: Необходимый фильтр (" & cstrVDName & ") не зарегистрирован"
Else
    objRegFilterInfo.Filter objFilterInfo
End If

'ТУТ не надо - ошибка при проигрывании cstrMPG2DemName демультиплексер mpg2splt.ax
''' ^ впихнули

''впихнуть в граф наш декодер cstrSVDName декодер
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrSVDName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    ToDebug "Error: Необходимый фильтр (" & cstrSVDName & ") не зарегистрирован"
    SourceFilter = False
Else
    SourceFilter = True
End If

If SourceFilter Then
    objRegFilterInfo.Filter objFilterInfo
    If err.Number <> 0 Then
        ToDebug "Error: переустановите Universal Open Source MPEG Source"
        SourceFilter = False
    End If

    If SourceFilter Then
        'для source фильтра "Universal Open Source MPEG Source"
        'Рендерим
        If Not objRegFilterInfo Is Nothing Then
            objFilterInfo.filename = mpgName
            For Each objPin In objFilterInfo.Pins
                objPin.Render 'не все, некоторые отсутствуют (on error)
            Next
        End If
        err.Clear     'если нет субтитров у "Universal Open Source MPEG Source"
        If objFilterInfo.filename = mpgName Then
            'вышло с соурсом
            SourceFilter = True
        Else
            SourceFilter = False
        End If
    End If
End If

'End If 'надо sourse

If Not SourceFilter Then
    'проба с другим нашим фильтром (не соурс)
    'все с начала
    MpegMediaClose
    Set mobjManager = New FilgraphManager

    'впихнуть в граф наш декодер
    For Each objRegFilterInfo In mobjManager.RegFilterCollection
        If (cstrVDName = objRegFilterInfo.name) Then
            Exit For
        End If
        Set objRegFilterInfo = Nothing
    Next
    If (objRegFilterInfo Is Nothing) Then
        ToDebug "Error: Необходимый фильтр (" & cstrVDName & ") не зарегистрирован"
        MpegMediaClose    'Set mobjManager = Nothing
        RenderMPV2 = False
        Exit Function
    End If

    err.Clear
    objRegFilterInfo.Filter objFilterInfo

    If err.Number <> 0 Then
        ToDebug "Error: наш фильтр не смог"
        MpegMediaClose    'Set mobjManager = Nothing
        RenderMPV2 = False
        Exit Function
    End If

    ''впихнуть в граф наш декодер cstrMPG2DemName демультиплексер mpg2splt.ax
    For Each objRegFilterInfo In mobjManager.RegFilterCollection
        If (cstrMPG2DemName = objRegFilterInfo.name) Then
            Exit For
        End If
        Set objRegFilterInfo = Nothing
    Next
    If (objRegFilterInfo Is Nothing) Then
        ToDebug "Error: Необходимый фильтр (" & cstrMPG2DemName & ") не зарегистрирован"
    Else
        objRegFilterInfo.Filter objFilterInfo
    End If



    'срендерить
    Call mobjManager.RenderFile(mpgName)

End If    'Not SourceFilter

Set objRegFilterInfo = Nothing

If (0& = err.Number) Then
    MpegMediaOpen = True

    Call MpegSizeAdjust(False)

Else
    ToDebug "Ошибка: Не могу обработать файл: " & mpgName
    err.Clear
    '        Set mobjSampleGrabber = Nothing
    MpegMediaClose     'Set mobjManager = Nothing
    RenderMPV2 = False
    Exit Function
End If

Set objPosition = mobjManager
Set objVideo = mobjManager
Set objAudio = mobjManager
On Error Resume Next: objAudio.Volume = -10000: On Error GoTo 0
'Set objAudio = Nothing

'список фильтров

'Debug.Print "MPV2"
ToDebug "Render MPV2 Filters Begin"
For Each GraphFilter In mobjManager.FilterCollection
    'Debug.Print GraphFilter.Name
    ToDebug vbTab & GraphFilter.name
Next GraphFilter
ToDebug "Render MPV2 Filters End"

If AutoShots Then
    'проба захвата кадра
    DoEvents
    MPGCaptureBasicVideo FrmMain.PicTempHid(0)
    If MPGCaptured = False Then
        'ToDebug "Ошибка: файл загружен, но есть ошибка захвата кадра: " & mpgName
        If Not AutoNoMessFlag Then myMsgBox msgsvc(38) & vbCrLf & mpgName
    End If
End If

RenderMPV2 = True
End Function


Private Function RenderMPV1() As Boolean
'вставка в граф
'"MPEG-I Stream Splitter" 'quartz.dll
'"MPEG Video Decoder" 'quartz.dll

Dim objRegFilterInfo As IRegFilterInfo
Dim objFilterInfo As IFilterInfo
'Dim udtMediaType As TAMMediaType
Dim GraphFilter As IFilterInfo

If MpegMediaOpen Then Call MpegMediaClose

Set mobjManager = New FilgraphManager
'впихнуть в граф наш декодер cstrFVDName
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrM1SSName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    ToDebug "Error: Необходимый фильтр (" & cstrM1SSName & ") не зарегистрирован"
    MpegMediaClose     'Set mobjManager = Nothing
    RenderMPV1 = False
    Exit Function
End If
'    On Error Resume Next
Call objRegFilterInfo.Filter(objFilterInfo)
Set objRegFilterInfo = Nothing

'впихнуть в граф наш декодер cstrM1SSName
For Each objRegFilterInfo In mobjManager.RegFilterCollection
    If (cstrM1VDName = objRegFilterInfo.name) Then
        Exit For
    End If
    Set objRegFilterInfo = Nothing
Next
If (objRegFilterInfo Is Nothing) Then
    ToDebug "Error: Необходимый фильтр (" & cstrM1VDName & ") не зарегистрирован"
    MpegMediaClose     'Set mobjManager = Nothing
    RenderMPV1 = False
    Exit Function
End If
'    On Error Resume Next
Call objRegFilterInfo.Filter(objFilterInfo)
Set objRegFilterInfo = Nothing


'With udtMediaType
'    .MajorType = UUIDFromString(amIDMediaTypeVideo)
'    .SubType = UUIDFromString(amIDMediaTypeVideoRGB24)
'    .FormatType = UUIDFromString(amIDFormatVideoInfo)
'End With

err.Clear
mobjManager.RenderFile mpgName
If (0& = err.Number) Then

    MpegMediaOpen = True
    Call MpegSizeAdjust(False)

    '    With mudtBitmapInfo
    '        .Size = Len(mudtBitmapInfo)
    '        .Planes = 1
    '        .BitCount = 24
    '    End With

    ' что-то не соединилось
    '   For Each objPinInfo In objFilterInfo.Pins
    '       If (objPinInfo.ConnectedTo Is Nothing) Then
    '           Set mobjSampleGrabber = Nothing
    '           Set objPinInfo = Nothing
    '           Exit For
    '       End If
    '       Set objPinInfo = Nothing
    '   Next

    err.Clear

Else
    ToDebug "Ошибка: не могу обработать файл: " & mpgName
    RenderMPV1 = False
    MpegMediaClose     'Set mobjManager = Nothing
    Exit Function


End If

Set objPosition = mobjManager
Set objVideo = mobjManager
Set objAudio = mobjManager
On Error Resume Next: objAudio.Volume = -10000: On Error GoTo 0
'Set objAudio = Nothing

'список фильтров
'Debug.Print "MPV1"
ToDebug "Render MPV1 Filters Begin"
For Each GraphFilter In mobjManager.FilterCollection
    'Debug.Print GraphFilter.Name
    ToDebug " " & GraphFilter.name
Next GraphFilter
ToDebug "Render MPV1 Filters End"

If AutoShots Then
'проба захвата кадра
DoEvents
MPGCaptureBasicVideo FrmMain.PicTempHid(0)
If MPGCaptured = False Then
    'ToDebug "Ошибка: файл загружен, но есть ошибка захвата кадра: " & mpgName
    If Not AutoNoMessFlag Then myMsgBox msgsvc(38) & vbCrLf & mpgName
End If
End If

RenderMPV1 = True
End Function


Private Function RenderAuto() As Boolean
'Рендер как есть автоматом
Dim GraphFilter As IFilterInfo
'Dim imEvent As IMediaEvent

If MpegMediaOpen Then Call MpegMediaClose

Set mobjManager = New FilgraphManager

On Error Resume Next 'RenderAuto()

mobjManager.RenderFile mpgName
If (0& = err.Number) Then
    MpegMediaOpen = True

'подогнать и показать видео окно
    MpegSizeAdjust False


Else
    ToDebug "Ошибка фильтра по умолчанию: не могу обработать файл: " & mpgName
    RenderAuto = False
    MpegMediaClose  'Set mobjManager = Nothing
    Exit Function
End If

Set objPosition = mobjManager
Set objVideo = mobjManager
Set objAudio = mobjManager

On Error Resume Next: objAudio.Volume = -10000: On Error GoTo 0
'Set objAudio = Nothing

'список фильтров
'Debug.Print "Auto"
'If mobjManager.FilterCollection Is Nothing Then
'Debug.Print "f"
'End If

On Error Resume Next
ToDebug "--- Render Auto Filters Begin"
For Each GraphFilter In mobjManager.FilterCollection
    ToDebug " " & GraphFilter.name
    'Debug.Print GraphFilter.Name
Next GraphFilter
ToDebug "--- Render Auto Filters End"
On Error GoTo 0

If AutoShots Then
    'проба захвата кадра
    DoEvents
    MPGCaptureBasicVideo FrmMain.PicTempHid(0)
    If MPGCaptured = False Then
        If Not AutoNoMessFlag Then
            myMsgBox msgsvc(38) & vbCrLf & mpgName
            ToDebug "Err RA: " & msgsvc(38) & vbCrLf & mpgName
        End If
    End If
End If

RenderAuto = True
End Function

Private Sub AutoFillStore()
' в текущем сеансе запоминать вписываемое в редакторе
Dim i As Integer
Dim StoreFlag As Boolean

ToDebug "SaveEditorStore."

'название
If Len(TextMName.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextMName.ListCount - 1
        If TextMName.List(i) = TextMName.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextMName.AddItem TextMName.Text
End If

'метка
If Len(TextLabel.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextLabel.ListCount - 1
        If TextLabel.List(i) = TextLabel.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextLabel.AddItem TextLabel.Text
End If

'жанр
If Len(TextGenre.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextGenre.ListCount - 1
        If TextGenre.List(i) = TextGenre.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextGenre.AddItem TextGenre.Text
End If

'произв
If Len(TextCountry.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextCountry.ListCount - 1
        If TextCountry.List(i) = TextCountry.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextCountry.AddItem TextCountry.Text
End If

'год
If Len(TextYear.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextYear.ListCount - 1
        If TextYear.List(i) = TextYear.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextYear.AddItem TextYear.Text
End If

'реж
If Len(TextAuthor.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextAuthor.ListCount - 1
        If TextAuthor.List(i) = TextAuthor.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextAuthor.AddItem TextAuthor.Text
End If

'в ролях
'StoreFlag = True
'For i = 0 To TextRole.ListCount
'If TextRole.List(i) = TextRole.Text Then StoreFlag = False: Exit For
'Next i
'If StoreFlag Then TextRole.AddItem TextRole.Text

'должник
If Len(TextUser.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextUser.ListCount - 1
        If TextUser.List(i) = TextUser.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextUser.AddItem TextUser.Text
End If

'Примеч TextOther
If Len(TextOther.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextOther.ListCount - 1
        If TextOther.List(i) = TextOther.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextOther.AddItem TextOther.Text
End If

'Rating
If Len(TextRate.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextRate.ListCount - 1
        If TextRate.List(i) = TextRate.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextRate.AddItem TextRate.Text
End If

'Lang
If Len(TextLang.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextLang.ListCount - 1
        If TextLang.List(i) = TextLang.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextLang.AddItem TextLang.Text
End If

'Subt
If Len(TextSubt.Text) <> 0 Then
    StoreFlag = True
    For i = 0 To TextSubt.ListCount - 1
        If TextSubt.List(i) = TextSubt.Text Then StoreFlag = False: Exit For
    Next i
    If StoreFlag Then TextSubt.AddItem TextSubt.Text
End If

End Sub

'по кнопке auto1
Private Sub AutoScrShotsN(F As Long, ss As Integer)
Select Case ss
Case 1

Set PicSS1 = Nothing
PicSS1.Height = ScrShotEd_W * movie.Height / movie.Width
PicSS1.Width = ScrShotEd_W
pos1 = 1 + (Rnd() * F)
pos1 = m_cAVI.AVIStreamNearestNextKeyFrame(pos1)

    If Opt_PicRealRes Then 'большие
PicSS1Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
PicSS1Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
m_cAVI.DrawFrame PicSS1Big.hdc, pos1, 0, 0, Transparent:=False
PicSS1Big.Picture = PicSS1Big.Image
    End If
    
m_cAVI.DrawFrame PicSS1.hdc, pos1, lWidth:=PicSS1.ScaleWidth, lHeight:=PicSS1.ScaleHeight, Transparent:=False
PicSS1.Picture = PicSS1.Image

Case 2

Set PicSS2 = Nothing
PicSS2.Height = ScrShotEd_W * movie.Height / movie.Width
PicSS2.Width = ScrShotEd_W

pos2 = 1 + (Rnd() * F)
pos2 = m_cAVI.AVIStreamNearestNextKeyFrame(pos2)

    If Opt_PicRealRes Then 'большие
PicSS2Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
PicSS2Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
m_cAVI.DrawFrame PicSS2Big.hdc, pos2, 0, 0, Transparent:=False
PicSS2Big.Picture = PicSS2Big.Image
    End If
    
m_cAVI.DrawFrame PicSS2.hdc, pos2, lWidth:=PicSS1.ScaleWidth, lHeight:=PicSS1.ScaleHeight, Transparent:=False
PicSS2.Picture = PicSS2.Image
'PicSS1.Picture = movie.Image

Case 3

Set PicSS3 = Nothing
PicSS3.Height = ScrShotEd_W * movie.Height / movie.Width
PicSS3.Width = ScrShotEd_W

pos3 = 1 + (Rnd() * F)
pos3 = m_cAVI.AVIStreamNearestNextKeyFrame(pos3)

    If Opt_PicRealRes Then 'большие
PicSS3Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
PicSS3Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
m_cAVI.DrawFrame PicSS3Big.hdc, pos3, 0, 0, Transparent:=False
PicSS3Big.Picture = PicSS3Big.Image
    End If
    
m_cAVI.DrawFrame PicSS3.hdc, pos3, lWidth:=PicSS1.ScaleWidth, lHeight:=PicSS1.ScaleHeight, Transparent:=False
PicSS3.Picture = PicSS3.Image

End Select

End Sub



Private Sub AutoScrShots(F As Long)
'автоскриншоты для ави
'F - всего фреймов
Dim tmp As Long
Set PicSS1 = Nothing ': Set PicSS1Big = Nothing
Set PicSS2 = Nothing
Set PicSS3 = Nothing


PicSS1.Height = ScrShotEd_W * movie.Height / movie.Width
PicSS1.Width = ScrShotEd_W
 PicSS2.Height = PicSS1.Height
 PicSS2.Width = PicSS1.Width
  PicSS3.Height = PicSS1.Height
  PicSS3.Width = PicSS1.Width

tmp = (F - F * 0.05) / 3 '4
pos1 = 1 + (Rnd() * tmp) '1-4
pos2 = tmp + (Rnd() * tmp) + 1 '4-8
pos3 = tmp * 2 + (Rnd() * tmp) + 1 '8-12

pos1 = m_cAVI.AVIStreamNearestNextKeyFrame(pos1)
pos2 = m_cAVI.AVIStreamNearestNextKeyFrame(pos2)
pos3 = m_cAVI.AVIStreamNearestNextKeyFrame(pos3)

    If Opt_PicRealRes Then 'большие
PicSS1Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
PicSS1Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
PicSS2Big.Width = PicSS1Big.Width
PicSS2Big.Height = PicSS1Big.Height
PicSS3Big.Width = PicSS1Big.Width
PicSS3Big.Height = PicSS1Big.Height

'тестовый
'tmp = m_cAVI.AVIStreamNearestNextKeyFrame(pos1)
'm_cAVI.DrawFrame PicSS1Big.hdc, tmp, 0, 0, Transparent:=False

m_cAVI.DrawFrame PicSS1Big.hdc, pos1, 0, 0, Transparent:=False
m_cAVI.DrawFrame PicSS2Big.hdc, pos2, 0, 0, Transparent:=False
m_cAVI.DrawFrame PicSS3Big.hdc, pos3, 0, 0, Transparent:=False

PicSS2Big.Picture = PicSS2Big.Image
PicSS1Big.Picture = PicSS1Big.Image
PicSS3Big.Picture = PicSS3Big.Image

'small
PicSS1.PaintPicture PicSS1Big, 0, 0, PicSS1.ScaleWidth, PicSS1.ScaleHeight
PicSS1.Picture = PicSS1.Image
PicSS2.PaintPicture PicSS2Big, 0, 0, PicSS2.ScaleWidth, PicSS2.ScaleHeight
PicSS2.Picture = PicSS2.Image
PicSS3.PaintPicture PicSS3Big, 0, 0, PicSS3.ScaleWidth, PicSS3.ScaleHeight
PicSS3.Picture = PicSS3.Image

    Else
'only small
m_cAVI.DrawFrame PicSS1.hdc, pos1, lWidth:=PicSS1.ScaleWidth, lHeight:=PicSS1.ScaleHeight, Transparent:=False
PicSS1.Picture = PicSS1.Image
m_cAVI.DrawFrame PicSS2.hdc, pos2, lWidth:=PicSS2.ScaleWidth, lHeight:=PicSS2.ScaleHeight, Transparent:=False
PicSS2.Picture = PicSS2.Image
m_cAVI.DrawFrame PicSS3.hdc, pos3, lWidth:=PicSS3.ScaleWidth, lHeight:=PicSS3.ScaleHeight, Transparent:=False
PicSS3.Picture = PicSS3.Image

    End If
'Position.Value = lastRendedAVI
ToDebug "ScrShotsPos Avi: " & pos1 & " " & pos2 & " " & pos3
End Sub


Public Sub DrawCoverEdit()
'BitBlt picCanvas.hdc, 0, 0, PicFrontFace.Width, PicFrontFace.Height, _
  PicFrontFace.hdc, 0, 0, SRCCOPY
'picCanvas.Refresh

Dim PRatio As Double

Dim CanvasW As Single 'long
Dim CanvasH As Single
Dim CanvasHalfW As Single
Dim CanvasHalfH As Single
Dim k As Single
Dim chOr As Boolean

CanvasW = picCanvas.Width / Screen.TwipsPerPixelX
CanvasH = picCanvas.Height / Screen.TwipsPerPixelX
CanvasHalfW = CanvasW / 2
CanvasHalfH = CanvasH / 2

If PicFrontFace.Picture <> 0 Then

PRatio = PicFrontFace.Height / PicFrontFace.Width
k = CanvasH / CanvasW

If k < PRatio Then chOr = True

If chOr Then If PRatio > 1 Then PRatio = 1 / PRatio


    If chOr Then
        'centre hor
            picCanvas.PaintPicture PicFrontFace.Picture, CanvasHalfW - (CanvasW * PRatio * k) / 2, 0, CanvasW * PRatio * k, CanvasH
    Else
        'centre VERT
            picCanvas.PaintPicture PicFrontFace.Picture, 0, CanvasHalfH - (CanvasH * PRatio / k) / 2, CanvasW, CanvasH * PRatio / k
    End If


End If
End Sub



Private Function PutFields() As Long
'текстовые поля в базу, поля в базу
'возвратить rs("Key") до апдейта
'Dim tmp As String

On Error GoTo err

'replace2regional тут нужен? нет/ не всегда

'комбосы
'Select Case rs.Fields("Label").Type
'Case 2, 3, 4, 6, 7         'метка -числовые
'    tmp = Replace(TextLabel.Text, ",", ".")
'    rs.Fields("Label") = Val(tmp)
'Case Else
'    'как текст
If Len(TextLabel.Text) > 255 Then rs.Fields("Label") = Left$(TextLabel.Text, 255) Else rs.Fields("Label") = TextLabel.Text
'End Select

If Len(TextMName.Text) > 255 Then rs.Fields("MovieName") = Left$(TextMName.Text, 255) Else rs.Fields("MovieName") = TextMName.Text
If Len(TextGenre.Text) > 255 Then rs.Fields("Genre") = Left$(TextGenre.Text, 255) Else rs.Fields("Genre") = TextGenre.Text
If Len(TextYear.Text) > 255 Then rs.Fields("Year") = Left$(TextYear.Text, 255) Else rs.Fields("Year") = TextYear.Text
If Len(TextCountry.Text) > 255 Then rs.Fields("Country") = Left$(TextCountry.Text, 255) Else rs.Fields("Country") = TextCountry.Text
If Len(TextAuthor.Text) > 255 Then rs.Fields("Director") = Left$(TextAuthor.Text, 255) Else rs.Fields("Director") = TextAuthor.Text
If Len(TextSubt.Text) > 255 Then rs.Fields("SubTitle") = Left$(TextSubt.Text, 255) Else rs.Fields("SubTitle") = TextSubt.Text
If Len(TextLang.Text) > 255 Then rs.Fields("Language") = Left$(TextLang.Text, 255) Else rs.Fields("Language") = TextLang.Text
If Len(TextRate.Text) > 255 Then rs.Fields("Rating") = Left$(TextRate.Text, 255) Else rs.Fields("Rating") = TextRate.Text
rs.Fields("Other") = TextOther.Text    'memo
If Len(TextUser.Text) > 255 Then rs.Fields("Debtor") = Left$(TextUser.Text, 255) Else rs.Fields("Debtor") = TextUser.Text

'текст
rs.Fields("Acter") = TextRole.Text
rs.Fields("Time") = TextTimeHid.Text
rs.Fields("Resolution") = TextResolHid.Text
rs.Fields("Audio") = TextAudioHid.Text
rs.Fields("FPS") = TextFPSHid.Text
rs.Fields("FileLen") = Val(TextFilelenHid.Text)
rs.Fields("CDN") = TextCDN.Text
rs.Fields("MediaType") = ComboNos.Text
rs.Fields("Video") = TextVideoHid.Text
rs.Fields("FileName") = TextFileName.Text

If CDSerialCur <> vbNullString Then rs.Fields("snDisk") = CDSerialCur.Text

rs.Fields("CoverPath") = TextCoverURL.Text
rs.Fields("MovieURL") = TextMovURL.Text
rs.Fields("Annotation") = TextAnnotation.Text

PutFields = rs("Key")
rs.Update

ToDebug "Текст - в базе"
Exit Function

err:
If err <> 0 Then MsgBox err.Description
End Function





Public Sub ClearFields()
'очистка полей в редакторе

TextFilelenHid = "0"
TextTimeHid = vbNullString
'TextCompanyHid = rs.Fields("")
TextVideoHid = vbNullString
TextFPSHid = vbNullString
TextResolHid = vbNullString
TextAudioHid = vbNullString
'TextFramesHid = rs.Fields("")
TextTimeMSHid = "0"
TextMName.Text = vbNullString
TextLabel.Text = vbNullString
TextGenre.Text = vbNullString
TextCountry.Text = vbNullString
TextYear.Text = vbNullString
TextAuthor.Text = vbNullString
TextRole.Text = vbNullString
TextUser.Text = vbNullString
TextCDN.Text = "0"
TextFileName.Text = vbNullString
TextAnnotation.Text = vbNullString
TextOther.Text = vbNullString
CDSerialCur.Text = vbNullString

'lbInetMovieList.Clear
ComboSites.Text = vbNullString
ComboGenre.Text = vbNullString
ComboCountry.Text = vbNullString

TextRate = vbNullString
TextLang = vbNullString
TextSubt = vbNullString
TextCoverURL = vbNullString
TextMovURL = vbNullString

ComboNos.Text = vbNullString

End Sub
Public Sub GetFields()
'ToDebug "Записи - из базы"
Dim tmp As String

Mark2SaveFlag = False 'не делать Mark2Save

TextMName.Text = CheckNoNull("MovieName"):    FrameAddEdit.Caption = AddEditCapt & " > " & TextMName.Text

TextLabel.Text = CheckNoNull("Label")
TextGenre.Text = CheckNoNull("Genre")
TextYear.Text = CheckNoNull("Year")
TextCountry.Text = CheckNoNull("Country")
TextAuthor.Text = CheckNoNull("Director")
TextRole.Text = CheckNoNull("Acter")
TextTimeHid.Text = CheckNoNull("Time")
TextResolHid.Text = CheckNoNull("Resolution")
TextAudioHid.Text = CheckNoNull("Audio")
TextFPSHid.Text = CheckNoNull("FPS")
TextFilelenHid.Text = CheckNoNull("FileLen")
TextCDN.Text = CheckNoNull("CDN")
ComboNos.Text = CheckNoNull("MediaType")
TextVideoHid.Text = CheckNoNull("Video")
TextSubt.Text = CheckNoNull("SubTitle")
TextLang.Text = CheckNoNull("Language")
TextRate.Text = CheckNoNull("Rating")
TextFileName.Text = CheckNoNull("FileName")
TextUser.Text = CheckNoNull("Debtor")
CDSerialCur.Text = CheckNoNull("SNDisk")
TextOther.Text = CheckNoNull("Other")
TextCoverURL.Text = CheckNoNull("CoverPath")
TextMovURL.Text = CheckNoNull("MovieURL")
TextAnnotation.Text = CheckNoNull("Annotation")

'для инет поиска
tmp = Replace(TextMName.Text, "(", vbNullString)
tmp = Replace(tmp, ")", vbNullString)
tmp = Replace(tmp, "/", vbNullString)
If Len(TextMName.Text) <> 0 Then TxtIName.Text = LCase$(tmp)

'почистить комбо со списками
ComboGenre.Text = vbNullString
ComboCountry.Text = vbNullString
ComboOther.Text = vbNullString


Mark2SaveFlag = True 'вернуть Mark2Save
End Sub

Public Sub GetEditPix()
'ToDebug "Картинки - из базы"


Set PicSS1 = Nothing: NoPic1Flag = False
Set PicSS1Big = Nothing
If GetPic(PicSS1, 1, "SnapShot1") Then
    PicSS1.Picture = PicSS1.Image
End If

Set PicSS2 = Nothing: NoPic2Flag = False
Set PicSS2Big = Nothing
If GetPic(PicSS2, 1, "SnapShot2") Then PicSS2.Picture = PicSS2.Image

Set PicSS3 = Nothing: NoPic3Flag = False
Set PicSS3Big = Nothing
If GetPic(PicSS3, 1, "SnapShot3") Then PicSS3.Picture = PicSS3.Image

Set PicFrontFace = Nothing
Set picCanvas = Nothing
If GetPic(PicFrontFace, 1, "FrontFace") Then
    PicFrontFace.Picture = PicFrontFace.Image
    DrawCoverEdit
End If

End Sub
Private Sub PosScroll()
Dim temp As Long

On Error Resume Next

temp = Position.Value + 1000

If temp <= PPMin Then
    PPMin = Position.Value - 1000
    If PPMin < 0 Then PPMin = 0
    PositionP.Min = PPMin
    PPMax = Position.Value + 1000
    If PPMax > Frames Then PPMax = Frames
    PositionP.Max = PPMax
Else
    PPMax = Position.Value + 1000
    If PPMax > Frames Then PPMax = Frames - 1
    PositionP.Max = PPMax
    PPMin = Position.Value - 1000
    If PPMin < 0 Then PPMin = 0
    PositionP.Min = PPMin
End If

    'If Position.Value > 1000 Then
    ' PositionP.Value = Position.Value 'PositionP.Min + 1000
    'Else
PositionP.Value = Position.Value  'PositionP.Min
    'End If

    Screen.MousePointer = vbHourglass
    If Position.Value = 0 Then
        pRenderFrame 1
    Else
       pRenderFrame CDbl(Position.Value)
    End If
    
Screen.MousePointer = vbNormal
End Sub


Private Sub MPGPosScroll()

'Dim objPosition As IMediaPosition
'Dim objBasicVideo As IBasicVideo
'objBasicVideo.AvgTimePerFrame

'   On Error Resume Next
'    Set objPosition = mobjManager
'    Set objBasicVideo = mobjManager
'mobjManager.Stop
'Debug.Print (Position.Value - 1) / 1000
'mobjManager.StopWhenReady
'mobjManager.Pause

Dim temp As Long
'MPGPosScroll = True
'MPGPosScrollFlag = True

On Error Resume Next
temp = Position.Value + Range

If temp <= PPMin Then
PPMin = Position.Value - Range
If PPMin < 0 Then PPMin = 0
    PositionP.Min = PPMin
PPMax = Position.Value + Range
If PPMax > TimesX100 Then PPMax = TimesX100
    PositionP.Max = PPMax
Else
PPMax = Position.Value + Range
If PPMax > TimesX100 Then PPMax = TimesX100
    PositionP.Max = PPMax
PPMin = Position.Value - Range
If PPMin < 0 Then PPMin = 0
    PositionP.Min = PPMin
End If

PositionP.Value = Position.Value  'PositionP.Min

If Position.Value = Position.Max Then PositionP.Value = PositionP.Max

'Screen.MousePointer = vbHourglass

    If Position.Value = 1 Then
        objPosition.CurrentPosition = 0 '(Position.Value - 1) / 1000 + 0.04
        'mobjManager.Pause
    Else
        objPosition.CurrentPosition = Position.Value / 100 '(Position.Value - 1) / 1000 + 0.04
        'Debug.Print objPosition.CurrentPosition
        
        'movie.Refresh
        
        'objPosition.CurrentPosition = CLng(Position.Value / 100)
        
        
        
        'mobjManager.Pause
    End If
    
Screen.MousePointer = vbNormal


End Sub

Private Function filltrue(o As Object) As Boolean
'заполнять или нельзя
filltrueAdd = False
filltrue = False
If ChInFilFl.Value = vbUnchecked Then 'разрешили все менять
    filltrue = True
    Exit Function
ElseIf ChInFilFl.Value = vbChecked Then  'только если пусто
    If Len(o.Text) = 0 Then filltrue = True
ElseIf ChInFilFl.Value = vbGrayed Then 'добавлять
    If Len(o.Text) = 0 Then filltrue = True
    filltrueAdd = True
    'там проверять на filltrueAdd только если filltrue = False
End If
End Function
Private Sub mnuMovieCopyClip_Click()
Clipboard.Clear
Clipboard.SetData m_cAVI.FramePicture(Position.Value, AviWidth, AviHeight)
End Sub
Private Sub mnuMovieSaveFrame_Click()
'сохранить в файл bmp кадр из окна movie avi
Dim dib As cDIB
Dim pDIB As Long 'pointer to packed DIB in memory

'create a DIB class to load the frames into
Set dib = New cDIB
pDIB = AVIStreamGetFrame(From_m_pGF, Position.Value)  'returns "packed DIB"
If dib.CreateFromPackedDIBPointer(pDIB) Then
    'Call dib.WriteToFile(App.Path & "\" & Position.Value& ".bmp")
    Call dib.WriteToFile(pSaveDialogBMP(DTitle:=ComSaveRec.Caption))
End If
    
Set dib = Nothing
End Sub
Public Sub SetFromScript()

On Error Resume Next

With SC.CodeObject

    If Len(.MTitle) <> 0 Then
        If filltrue(TextMName) Then
            TextMName = .MTitle
        ElseIf filltrueAdd Then
            TextMName = TextMName & " / " & .MTitle
        End If
    End If
    If Len(.MYear) <> 0 Then
        If filltrue(TextYear) Then
            TextYear = .MYear
        ElseIf filltrueAdd Then
            TextYear = TextYear & ", " & .MYear
        End If
    End If
    If Len(.MGenre) <> 0 Then
        If filltrue(TextGenre) Then
            TextGenre = .MGenre
        ElseIf filltrueAdd Then
            TextGenre = TextGenre & ", " & .MGenre
        End If
    End If
    If Len(.MDirector) <> 0 Then
        If filltrue(TextAuthor) Then
            TextAuthor = .MDirector
        ElseIf filltrueAdd Then
            TextAuthor = TextAuthor & ", " & .MDirector
        End If
    End If
    If Len(.MActors) <> 0 Then
        If filltrue(TextRole) Then
            TextRole = .MActors
        ElseIf filltrueAdd Then
            TextRole = TextRole & ", " & .MActors
        End If
    End If
    If Len(.MDescription) <> 0 Then
        If filltrue(TextAnnotation) Then
            TextAnnotation = .MDescription
        ElseIf filltrueAdd Then
            TextAnnotation = TextAnnotation & vbCrLf & .MDescription
        End If
    End If
    If Len(.MCountry) <> 0 Then
        If filltrue(TextCountry) Then
            TextCountry = .MCountry
        ElseIf filltrueAdd Then
            TextCountry = TextCountry & ", " & .MCountry
        End If
    End If
    If Len(.MRating) <> 0 Then
        If filltrue(TextRate) Then
            TextRate = .MRating
        ElseIf filltrueAdd Then
            'TextRate = (Val(TextRate) + Val(.MRating)) / 2
            TextRate = (Str2Val(TextRate) + Str2Val(.MRating)) / 2
        End If
    End If
    If Len(.MLang) <> 0 Then
        If filltrue(TextLang) Then
            TextLang = .MLang
        ElseIf filltrueAdd Then
            TextLang = TextLang & ", " & .MLang
        End If
    End If
    If Len(.MSubt) <> 0 Then
        If filltrue(TextSubt) Then
            TextSubt = .MSubt
        ElseIf filltrueAdd Then
            TextSubt = TextSubt & ", " & .MSubt
        End If
    End If
    If Len(.MPicURL) <> 0 Then 'замещать
            TextCoverURL = .MPicURL
    End If
    If Len(.MOther) <> 0 Then
        If filltrue(TextOther) Then
            TextOther = .MOther
        ElseIf filltrueAdd Then
            TextOther = TextOther & vbCrLf & .MOther
        End If
    End If

    'сайт , замещать
    If Len(ComboSites.Text) <> 0 Then TextMovURL = ComboSites.Text


    'pix
    If (PicFrontFace.Picture = 0) Or (ChInFilFl.Value <> vbChecked) Then
        If Len(.MPicURL) <> 0 Then
            OpenURLProxy .MPicURL, "pic"
            'Debug.Print SC.CodeObject.MPicURL
        Else
            Set ImgPrCov = Nothing: Set PicFrontFace = Nothing: Set picCanvas = Nothing
            NoPicFrontFaceFlag = True
        End If
    End If

End With

ToDebug "инфо - в полях редактора. Ошибки - " & CStr(err.Number <> 0)        'да нет

End Sub

Private Sub ChangeComboHeights()
'высота кобмиков
'паразитно подставляет подходящие значения из списка вместо взятых из базы.
'не вызывать при ресайзе

Dim X As Long, Y As Long, w As Long
'SWP_NOZORDER Or SWP_NOMOVE Or SWP_DRAWFRAME
Const CB_SETDROPPEDWIDTH = &H160

X = ScaleX(ComboGenre.Left, vbTwips, vbPixels)
Y = ScaleY(ComboGenre.Top, vbTwips, vbPixels)
w = ScaleY(ComboGenre.Width, vbTwips, vbPixels)
SetWindowPos ComboGenre.hWnd, 0, X, Y, w, 500, 0

X = ScaleX(ComboCountry.Left, vbTwips, vbPixels)
Y = ScaleY(ComboCountry.Top, vbTwips, vbPixels)
w = ScaleY(ComboCountry.Width, vbTwips, vbPixels)
SetWindowPos ComboCountry.hWnd, 0, X, Y, w, 500, 0

'список аля ComboSites в тех вкладке
X = ScaleX(cBasePicURL.Left, vbTwips, vbPixels)
Y = ScaleY(cBasePicURL.Top, vbTwips, vbPixels)
w = ScaleY(cBasePicURL.Width, vbTwips, vbPixels)
SetWindowPos cBasePicURL.hWnd, 0, X, Y, w, 500, 0
'ширина выпадающего списка
Call SendMessage(cBasePicURL.hWnd, CB_SETDROPPEDWIDTH, 400, ByVal 0&)

'список сайтов в иннфо
X = ScaleX(ComboSites.Left, vbTwips, vbPixels)
Y = ScaleY(ComboSites.Top, vbTwips, vbPixels)
w = ScaleY(ComboSites.Width, vbTwips, vbPixels)
SetWindowPos ComboSites.hWnd, 0, X, Y, w, 500, 0

'список скриптов
X = ScaleX(ComboInfoSites.Left, vbTwips, vbPixels)
Y = ScaleY(ComboInfoSites.Top, vbTwips, vbPixels)
w = ScaleY(ComboInfoSites.Width, vbTwips, vbPixels)
SetWindowPos ComboInfoSites.hWnd, 0, X, Y, w, 500, 0

'примечания, качество
X = ScaleX(ComboOther.Left, vbTwips, vbPixels)
Y = ScaleY(ComboOther.Top, vbTwips, vbPixels)
w = ScaleY(ComboOther.Width, vbTwips, vbPixels)
SetWindowPos ComboOther.hWnd, 0, X, Y, w, 500, 0

''не пустые
X = ScaleX(ComboNos.Left, vbTwips, vbPixels)
Y = ScaleY(ComboNos.Top, vbTwips, vbPixels)
w = ScaleY(ComboNos.Width, vbTwips, vbPixels)
SetWindowPos ComboNos.hWnd, 0, X, Y, w, 500, 0

'x = ScaleX(TextLang.Left, vbTwips, vbPixels)
'y = ScaleY(TextLang.Top, vbTwips, vbPixels)
'w = ScaleY(TextLang.Width, vbTwips, vbPixels)
''SetWindowPos TextLang.hwnd, 0, x, y, w, 500, SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_NOSENDCHANGING ' Or SWP_NOSIZE
'
'x = ScaleX(TextSubt.Left, vbTwips, vbPixels)
'y = ScaleY(TextSubt.Top, vbTwips, vbPixels)
'w = ScaleY(TextSubt.Width, vbTwips, vbPixels)
'SetWindowPos TextSubt.hwnd, 0, x, y, w, 500, 0
'
'x = ScaleX(TextOther.Left, vbTwips, vbPixels)
'y = ScaleY(TextOther.Top, vbTwips, vbPixels)
'w = ScaleY(TextOther.Width, vbTwips, vbPixels)
'SetWindowPos TextOther.hwnd, 0, x, y, w, 500, 0

      'фокусы
TextMName.SelLength = 0
ComboNos.SelLength = 0
cBasePicURL.SelLength = 0
'If TextCDN.Visible Then TextCDN.SetFocus
ComboSites.SelLength = 0
ComboInfoSites.SelLength = 0
ComboGenre.SelLength = 0
ComboCountry.SelLength = 0
ComboOther.SelLength = 0
'TextLang.SelLength = 0
'TextSubt.SelLength = 0
'TextOther.SelLength = 0
End Sub

Public Sub ClearVideo()
Dim i As Integer
'очистка (видео, обложка...)

Set movie = Nothing
movie.Width = MovieEd_W
movie.Height = MovieEd_H


isMPGflag = False: isAVIflag = False: isDShflag = False
pos1 = 0: pos2 = 0: pos3 = 0
PositionP.Value = 0: Position.Value = 0

MPGCodec = vbNullString
PixelRatio = 1: PixelRatioSS = 1

ComRND(0).Enabled = False: ComRND(1).Enabled = False: ComRND(2).Enabled = False
ComAutoScrShots.Enabled = False
For i = 0 To 2: optAspect(i).Enabled = False: Next i

        
'Set PicFrontFace = Nothing
'Set picCanvas = Nothing
Set PicFaceV = Nothing
Set Image0 = Nothing

Set ImgPrCov = Nothing

'Set AVIInf = Nothing
Timer2.Enabled = False

On Error Resume Next

If Not (m_cAVI Is Nothing) Then m_cAVI.filename = vbNullString 'unload
If MpegMediaOpen Then Call MpegMediaClose
'Debug.Print Time, "clear video"
End Sub


Public Sub SaveAutoAdd()
Dim curKey As String

FirstLVFill = False

ToDebug "Save AutoAdd"
'ToDebug "EditMode: " & rs.EditMode

curKey = rs("Key") 'ключ добавляемого поля. Важно переместится на него после апдейта

ToDebug "SaveAutKey=" & curKey

If SavePic1Flag Then
    If NoPic1Flag Then
        rs.Fields("SnapShot1") = vbNullString
ToDebug "Pic1 - no"
    Else
        If Opt_PicRealRes Then    'большую
            Pic2JPG PicSS1Big, 1, "SnapShot1"
ToDebug "Pic1 - big"
        Else    'мелкую
            Pic2JPG PicSS1, 1, "SnapShot1"
ToDebug "Pic1 - small"
        End If
    End If
End If
If SavePic2Flag Then
    If NoPic2Flag Then
        rs.Fields("SnapShot2") = vbNullString
ToDebug "Pic2 - no"
    Else
        If Opt_PicRealRes Then    'большую
            Pic2JPG PicSS2Big, 1, "SnapShot2"
ToDebug "Pic2 - big"
        Else    'мелкую
            Pic2JPG PicSS2, 1, "SnapShot2"
ToDebug "Pic2 - small"
        End If
    End If
End If
If SavePic3Flag Then
    If NoPic3Flag Then
        rs.Fields("SnapShot3") = vbNullString
ToDebug "Pic3 - no"
    Else
        If Opt_PicRealRes Then    'большую
            Pic2JPG PicSS3Big, 1, "SnapShot3"
ToDebug "Pic3 - big"
        Else    'мелкую
            Pic2JPG PicSS3, 1, "SnapShot3"
ToDebug "Pic3 - small"
        End If
    End If
End If
If SaveCoverFlag Then
    If NoPicFrontFaceFlag Then
        rs.Fields("FrontFace") = vbNullString
ToDebug "Cover - no"
    Else
        Pic2JPG PicFrontFace, 1, "FrontFace"
ToDebug "Cover - yes"
    End If
End If
SavePic1Flag = False: SavePic2Flag = False: SavePic3Flag = False: SaveCoverFlag = False

'положить в базу поля
PutFields    'там апдейт с возвратом позиции на до добавления


    RSGoto curKey 'встать на добавленный

    ListView.Sorted = False

    ReDim Preserve lvItemLoaded(ListView.ListItems.Count + 1)    ' 1
    Add2LV ListView.ListItems.Count, ListView.ListItems.Count + 1    '2

    CurLVKey = rs("Key") & """"
    
    CurSearch = GotoLV(CurLVKey)
    'пометить все автодобавляемые
    If ListView.ListItems.Count > 0 Then
        Set ListView.SelectedItem = ListView.ListItems(CurSearch)
    End If


'если была сортировка - произвести ее       - путь будут в конце
'If LVSortColl > 0 Then LVSOrt (LVSortColl)
'If LVSortColl = -1 Then SortByCheck 0, True

If FrameView.Visible Then ListView.SelectedItem.EnsureVisible    ': LVCLICK

ToDebug "...saved"
End Sub


Public Function OpenNewMovie(Optional mn As String) As Boolean
'еще OpenMovieForCapture
Dim tmp As String    'полученное имя файла
Dim AVIInf As New clsAVIInfo
Dim Handle As Long
Dim i As Integer
Dim tmpdrive As String
Dim tmpSerial As String
Dim sVolumeName As String
Dim temp As String
Dim MMI_Flag As Boolean    'узнал ли файл мминфо модуль

OpenNewMovie = True    'если нет, укажем это потом
'Mark2SaveFlag = False 'не краснить save
NewDiskAddFlag = False

TxtIName.Text = vbNullString

If Len(mn) <> 0 Then
    tmp = mn
Else
    If Not AppendMovieFlag Then
    'Открыть новый фильм
    
    FrameAddEdit.Enabled = False 'чтоб не нажимались батоны
        tmp = pLoadDialog    '(ComOpen.Caption)
    FrameAddEdit.Enabled = True
    
        DoEvents
        
        If LenB(tmp) = 0 Then
            Screen.MousePointer = vbNormal
            OpenNewMovie = False
            'Mark2SaveFlag = False
            Exit Function
        End If
    Else
        'добавление
        'взять текущее имя ави для добавления
        tmp = aviName
    End If
End If

If MpegMediaOpen Then MpegMediaClose

'ClearVideo 'nah
'If addFlag Then ClearFields - а если сначала вводили тексты
Set movie = Nothing

'Auto name in inet name
temp = GetNameFromPathAndName(tmp)
GetExtensionFromFileName temp, temp
If Len(TxtIName.Text) = 0 Then TxtIName.Text = LCase$(temp)

If Not AppendMovieFlag Then    'если добавление - то в add
    'Серийник. CDSerialCur-поле
    CDSerialCur = vbNullString    ': SameCDLabel = vbNullString
    tmpdrive = Left$(LCase$(tmp), 3)
    If DriveType(tmpdrive) = "CD-ROM" Then
        tmpSerial = Hex$(GetSerialNumber(tmpdrive, sVolumeName))
        If tmpSerial <> "0" Then    'надо
        
        If MediaSN = tmpSerial Then
            IsSameCdFlag = True 'это тот же носитель
        Else
            IsSameCdFlag = False
            MediaSN = tmpSerial 'запомнить серийник cd
        End If
        
            If CheckSameDisk Then
                'If SearchLVSimple(dbsnDiskInd, tmpSerial) Then    'есть ли уже в базе
                If SearchSNinbase(tmpSerial) Then    'есть ли уже в базе, там меняется SameCDLabel

                    If Not AutoNoMessFlag Then
                        If myMsgBox(msgsvc(28), vbYesNo, App.title, Me.hWnd) = vbNo Then    'продолжить?
                            Set AVIInf = Nothing
                            OpenNewMovie = False
                            Screen.MousePointer = vbNormal
                            Exit Function
                        Else
                            CheckSameDisk = False    ' AutoNoMessFlag = True 'не спрашивать боле
                            CDSerialCur = tmpSerial
                            TextLabel = SameCDLabel
                        End If
                    Else    'без вопросов
                        CDSerialCur = tmpSerial
                        TextLabel = SameCDLabel
                    End If
                Else
                    'CheckSameDisk = False 'не нашли и не искать боле
                    SameCDLabel = vbNullString

                    CDSerialCur = tmpSerial
                    If Len(TextLabel) = 0 Then TextLabel = sVolumeName
                End If
            Else
                CDSerialCur = tmpSerial
                If Len(TextLabel) = 0 Then
                    If Len(SameCDLabel) = 0 Then
                        TextLabel = sVolumeName
                    Else
                        TextLabel = SameCDLabel
                    End If
                End If
            End If
        End If
        
        'Носитель
        If IsSameCdFlag Then 'если тот же носитель
        ToDebug "Носитель тот-же."
            ComboNos.Text = MediaType
        Else
        'определить тип носителя
            ComboNos.Text = GetOptoInfo(Left$(tmp, 2))
            MediaType = ComboNos.Text
        End If
        
    Else
        ComboNos.Text = DriveType(tmpdrive)
        tmpSerial = Hex$(GetSerialNumber(tmpdrive))
        If tmpSerial <> "0" Then CDSerialCur = tmpSerial
    End If    'drivetype
End If    'append


Screen.MousePointer = vbHourglass

'определить что за файл
isMPGflag = False: isAVIflag = False: isDShflag = False

OpenAddmovFlag = True

If Not Opt_AviDirectShow Then
    '                                                           AVI
    aviName = tmp

    'если размер не 0
    Dim fSize As Long
    If isWindowsNt Then
        Dim Pointer As Long, lpFSHigh As Currency
        Pointer = lopen(tmp, OF_READ)
        GetFileSizeEx Pointer, lpFSHigh
        fSize = Int(lpFSHigh * 10000 / 1024)
        lclose Pointer
    Else
        fSize = Int(FileLen(aviName) / 1024)
    End If

    If fSize = 0 Then
        OpenNewMovie = False
        ToDebug "error: нулевая длина: " & aviName
        'Mark2SaveFlag = False
        Screen.MousePointer = vbNormal
        Exit Function
    End If

    'раcширения avi
    Dim FileExt As String
    Dim isAVIext As Boolean
    FileExt = LCase$(getExtFromFile(aviName))
    Select Case FileExt
    Case "avi"
        isAVIext = True
    Case "vid"
        isAVIext = True
    Case "divx"
        isAVIext = True
    Case "xvid"
        isAVIext = True
    End Select

    ToDebug "Open FileExt=" & FileExt
    '                                                               GetAviInfo
    If isAVIext Then GetAviInfo    'если разширения подходят для avi

    If isAVIflag Then
        If aferror Then
            ToDebug "Обработан как AVI, Ошибка захвата кадра"
        Else
            ToDebug "Обработан как AVI"
        End If
        Set AVIInf = Nothing
        Screen.MousePointer = vbNormal
        'Mark2SaveFlag = False
        Exit Function
    End If

End If    'not Opt_AviDirectShow

'                                                                       mpeg1/2?
On Error Resume Next
Handle = MediaInfo_New()
If err Then
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(51) & App.Path & "\MediaInfo.dll", vbCritical
            End If
'    MsgBox "MediaInfo.dll not found. Reinstall SurVideoCatalog!", vbCritical
ToDebug "Err_NoDLL: MediaInfo.dll"
    OpenNewMovie = False
    Screen.MousePointer = vbNormal
    Exit Function
End If

Call MediaInfo_Open(Handle, StrPtr(tmp))

If err = 0 Then
    i = MediaInfo_Count_Get(Handle, 1, -1)
    If i > 0 Then    'есть видео
        MMI_Flag = True
        MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))
        Select Case strMPEG
            Case "MPV1", "MPV2"
                isMPGflag = True
        End Select
    End If
    MediaInfo_Close Handle
Else
    err.Clear    ':On Error GoTo 0
End If

On Error GoTo 0

If isMPGflag Then
Else    '                                                          не MPV

    ''                                                               Обработка DS
    ''If isDShflag Then
    ToDebug "Info DShow файла: " & tmp
    DShName = tmp
    mpgName = tmp
    If DShGetInfo Then
        isDShflag = True
    Else
        OpenNewMovie = False
        isDShflag = False
    End If
    ''End If

    '    'срендерить автоматом
    '    mpgName = tmp
    ''исправить : DShGetInfo должна итти до RenderAuto!!! иначе не видит аспект
    '    If RenderAuto Then
    '        isDShflag = True
    '    End If

    mpgName = vbNullString
End If

If isMPGflag Or isDShflag Then
Else
    ToDebug "error: Не поддерживается или уже открыт: " & tmp
    OpenNewMovie = False
    If Not AutoNoMessFlag Then
        myMsgBox msgsvc(13) & tmp, vbInformation, , Me.hWnd    'Этот файл не поддерживается или уже открыт
    End If
End If


'                                                               Обработка MPEG
If isMPGflag Then
    ToDebug "Обработка файла MPEG: " & tmp
    mpgName = tmp
    If Not MpgGetInfo Then OpenNewMovie = False
End If
'If MPGCaptured = False Then ToDebug "Ошибка: файл загружен, но есть ошибка захвата кадра: " & mpgName

''                                                               Обработка DS
'If isDShflag Then
'    ToDebug "Обработка DShow файла: " & tmp
'    DShName = tmp
'    If Not DShGetInfo Then OpenNewMovie = False
'End If

Screen.MousePointer = vbNormal

On Error Resume Next: If Not frmAutoFlag Then Me.SetFocus
End Function

Private Sub OpenMovieForCapture(fName As String)
'открыть файл только для возможности снятия скриншотов
'fName - полный путь
Dim FileExt As String
Dim isAVIext As Boolean
Dim fSize As Long
Dim Handle As Long
Dim i As Integer
'Dim tmpdrive As String
'Dim tmpSerial As String
'Dim sVolumeName As String
'Dim temp As String
Dim MMI_Flag As Boolean    'узнал ли файл мминфо модуль

Set movie = Nothing

If Not Opt_AviDirectShow Then
    '                                                           AVI
    'если размер не 0
    If isWindowsNt Then
        Dim Pointer As Long, lpFSHigh As Currency
        Pointer = lopen(fName, OF_READ)
        GetFileSizeEx Pointer, lpFSHigh
        fSize = Int(lpFSHigh * 10000 / 1024)
        lclose Pointer
    Else
        fSize = Int(FileLen(fName) / 1024)
    End If
    If fSize = 0 Then
        ToDebug "error: нулевая длина: " & fName
        'Mark2SaveFlag = False
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    'раcширения avi
    FileExt = LCase$(getExtFromFile(fName))
    Select Case FileExt
    Case "avi"
        isAVIext = True
    Case "vid"
        isAVIext = True
    Case "divx"
        isAVIext = True
    Case "xvid"
        isAVIext = True
    End Select

    ToDebug "Open FileExt=" & FileExt

    If isAVIext Then
        'если разширения подходят для avi
        Call PrepareAviForCupture(fName)
    End If

    If isAVIflag Then
        If aferror Then
            ToDebug "ssAVI, cupture error"
        Else
            OpenAddmovFlag = True
            ToDebug "AVI ready for cupture"
        End If
        Screen.MousePointer = vbNormal
        Exit Sub
    End If

End If    'not Opt_AviDirectShow

'
'                                                                       mpeg?
On Error Resume Next
Handle = MediaInfo_New()
If err Then
'            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(51) & App.Path & "\MediaInfo.dll", vbCritical
'            End If
'    MsgBox "MediaInfo.dll not found. Reinstall SurVideoCatalog!", vbCritical
ToDebug "Err_NoDLL: MediaInfo.dll"
    Screen.MousePointer = vbNormal
    Exit Sub
End If

Call MediaInfo_Open(Handle, StrPtr(fName))

If err = 0 Then
    i = MediaInfo_Count_Get(Handle, 1, -1)
    If i > 0 Then    'есть видео
        MMI_Flag = True
        If UCase$(Left$(bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0)), 3)) = "MPV" Then
            isMPGflag = True
        End If
    End If
    MediaInfo_Close Handle
Else
    err.Clear
End If
On Error GoTo 0

If isMPGflag Then
Else    '                                                          не MPV
    ''                                                             Обработка DS

    ToDebug "DShow: " & fName
    DShName = fName
    mpgName = fName
    '                                          DS INFO
    If DShGetInfoForCapture Then
        isDShflag = True
        OpenAddmovFlag = True
    Else
        isDShflag = False
    End If

    mpgName = vbNullString
End If

If isMPGflag Or isDShflag Then
Else
    ToDebug "error: Не поддерживается или уже открыт: " & fName
    If Not AutoNoMessFlag Then
        myMsgBox msgsvc(13) & fName, vbInformation, , Me.hWnd
    End If
End If

'                                                               Обработка MPEG
If isMPGflag Then
    ToDebug "MPEG: " & fName
    mpgName = fName
    '                                               MPG INFO
    If MpgGetInfoForCapture Then
        isMPGflag = True
        OpenAddmovFlag = True
    Else
        isMPGflag = False
    End If
End If

Screen.MousePointer = vbNormal

On Error Resume Next: If Not frmAutoFlag Then Me.SetFocus

End Sub

Private Sub pRenderFrame(pos As Long)
'movie.Cls
m_cAVI.DrawFrame movie.hdc, pos, lWidth:=MovieWidth, lHeight:=MovieHeight, Transparent:=False
'm_cAVI.DrawFrame movie.hDC, pos, lWidth:=ScaleX(MovieWidth, vbPixels, vbTwips), lHeight:=ScaleY(MovieHeight, vbPixels, vbTwips), Transparent:=False
movie.Refresh
End Sub
Private Sub PrepareAviForCupture(fName As String)
'берет инфо, готовит видео окно и слайдеры
'генерит флаг ошибки aferror, если была

Dim AVIInf As New clsAVIInfo
'Dim temp As String

AVIInf.ReadFile (fName)

If AVIInf.NumStreams <> 0 Then        'avi?
    isAVIflag = True
Else
    isAVIflag = False
    Set AVIInf = Nothing
    Exit Sub
End If

Frames = AVIInf.numFrames
If Frames < 1 Then
    m_cAVI.filename = vbNullString
    Set AVIInf = Nothing
    Exit Sub    'no frames
End If

'resolution
AviWidth = AVIInf.Width
AviHeight = AVIInf.Height
Ratio = AviWidth / AviHeight

'Позиции
Position.Min = 0
If Frames > 1 Then Position.Max = Frames - 1 Else Position.Max = 1

Position.TickFrequency = Position.Max / 100
Position.SmallChange = Position.TickFrequency / 5
Position.LargeChange = Position.Max / 100

PositionP.Min = 0         'PPMin
PPMax = Position.Value + 1000
If PPMax > Frames Then PPMax = Frames
PositionP.Max = PPMax
PositionP.SmallChange = 1         'PositionP.TickFrequency / 1
PositionP.LargeChange = 50         'PositionP.Max / 100

MovieWidth = ScaleX(movie.Width, vbTwips, vbPixels)
If Ratio < 1 Then Ratio = 1.333
MovieHeight = MovieWidth / Ratio
movie.Height = ScaleY(MovieHeight, vbPixels, vbTwips)

Set AVIInf = Nothing

If Not NoVideoProcess Then
    'видео окно
    Set m_cAVI = New cAVIFrameExtract
    m_cAVI.filename = fName

    If Not aferror Then
        pRenderFrame 0
        lastRendedAVI = 0
        Position.Value = 0
        PositionP.Value = 0
        Position.Enabled = True
        PositionP.Enabled = True
        ComKeyNext.Enabled = True: ComKeyPrev.Enabled = True
        ComRND(0).Enabled = True: ComRND(1).Enabled = True: ComRND(2).Enabled = True
        ComAutoScrShots.Enabled = True
        'For i = 0 To 2: optAspect(i).Enabled = True: Next i
    Else
        m_cAVI.filename = vbNullString        'unload
    End If
End If        'видео процесс

End Sub


Private Sub GetAviInfo()

Dim AVIInf As New clsAVIInfo
Dim TimeS As String
Dim i As Integer
Dim temp As String

AVIInf.ReadFile (aviName)

If AVIInf.NumStreams <> 0 Then        'avi?
    isAVIflag = True
Else
    isAVIflag = False
    Set AVIInf = Nothing
    Exit Sub
End If

If Not AppendMovieFlag Then        'Новый файл
    '
    Frames = AVIInf.numFrames
    If Frames < 1 Then m_cAVI.filename = vbNullString: Exit Sub        'no frames
ToDebug "AVI_Frames=" & Frames

    'имя файла
    TextFileName.Text = GetNameFromPathAndName(aviName)

    'Time
    TimeS = FormatTime(AVIInf.PlayLength)
    TextTimeMSHid = AVIInf.PlayLength
    TextTimeHid = TimeS

    'fps
    TextFPSHid = Round(AVIInf.FrameRate, 3)
ToDebug "AVI_FPS=" & TextFPSHid

    'file size
    If isWindowsNt Then
        Dim Pointer As Long, lpFSHigh As Currency
        Pointer = lopen(aviName, OF_READ)
        'size of the file
        GetFileSizeEx Pointer, lpFSHigh
        TextFilelenHid = Int(lpFSHigh * 10000 / 1024)
        lclose Pointer
    Else
        TextFilelenHid = Int(FileLen(aviName) / 1024)
    End If
ToDebug "AVI_fSize=" & TextFilelenHid

    'аудио
    'TextAudioHid = AVIInf.SamplesPerSec & " " & AVIInf.Channels & " " & AVIInf.AudioFormat
    TextAudioHid = AVIInf.AllAudio
    If Trim$(TextAudioHid) = "0" Then TextAudioHid = "-"

    'видео
    temp = AVIInf.VideoCodec & " - " & AVIInf.CodecToName(AVIInf.VideoCodec) & AVIInf.VideoCodec2 & AVIInf.VideoBitrate
    '    tmp = AVIInf.VideoCodec
    '    If Len(tmp) <> 0 Then
    '        temp = tmp & " - " & AVIInf.CodecToName(tmp) & AVIInf.VideoCodec2 & AVIInf.VideoBitrate
    '    Else
    '        temp = "- " & AVIInf.CodecToName(AVIInf.VideoCodec22) & AVIInf.VideoCodec2 & AVIInf.VideoBitrate
    '    End If

    Do While InStr(1, temp, "  ")
        temp = Replace(temp, "  ", " ")
    Loop
    TextVideoHid = temp

    'resolution
    AviWidth = AVIInf.Width
    AviHeight = AVIInf.Height
    Ratio = AviWidth / AviHeight

    'FrSize
    TextResolHid = Trim$(str$(AviWidth)) & " x " & Trim$(str$(AviHeight))
    '
    TextCDN.Text = 1

    'Мета данные (Info)
    Dim mtag() As Long
    Dim InfoChar As String
    Dim InfoText As String

    'для суммирующихся
    Dim tmpLang As String
    Dim tmpCountry As String
    Dim tmpGenre As String
    '
    If AVIInf.GetInfoList(mtag) > 0 Then

        For i = 0 To UBound(mtag)
            If Len(mtag(i)) = 4 Then
                InfoChar = AVIInf.LongToFourCC(mtag(i))
                InfoText = AVIInf.QueryInfo(mtag(i))
                If Len(InfoText) <> 0 Then

                    Select Case InfoChar
                    Case "INAM"        'Name/Title
                        If filltrue(TextMName) Then
                            TextMName = InfoText
                        ElseIf filltrueAdd Then
                            TextMName = TextMName & ", " & InfoText
                        End If
                    Case "IART"        'Artist , Director
                        If filltrue(TextAuthor) Then
                            TextAuthor = InfoText
                        ElseIf filltrueAdd Then
                            TextAuthor = TextAuthor & ", " & InfoText
                        End If
                    Case "ISTR"        'Starring
                        If filltrue(TextRole) Then TextRole = InfoText
                        If filltrue(TextRole) Then
                            TextRole = InfoText
                        ElseIf filltrueAdd Then
                            TextRole = TextRole & ", " & InfoText
                        End If
                    Case "IAS1", "IAS2", "IAS3", "IAS4", "IAS5", "IAS6", "IAS7", "IAS8", "IAS9"        'First -9 Language
                        tmpLang = InfoText & ", " & tmpLang
                        'Case "ILNG" 'Language ?
                        '    If filltrue(TextLang) Then TextLang = TextLang & ", " & InfoText
                    Case "ICNT", "ISTD"        'Страна, 'Production studio
                        tmpCountry = InfoText & ", " & tmpCountry
                    Case "IGNR", "ISGN"        'Genre, 'Secondary genre
                        tmpGenre = InfoText & ", " & tmpGenre
                    Case "IWEB"        'Internet address
                        If filltrue(TextMovURL) Then
                            TextMovURL = InfoText
                        ElseIf filltrueAdd Then
                            TextMovURL = TextMovURL & ", " & InfoText
                        End If
                    Case "ICRD"        '  Creation Date (YYYYMMDD) наверно это год фильма
                        If filltrue(TextYear) Then
                            TextYear = Left$(InfoText, 4)
                        ElseIf filltrueAdd Then
                            TextYear = TextYear & ", " & Left$(InfoText, 4)
                        End If
                    Case "ISFT"      'Software
                        ToDebug AVIInf.QueryInfo(mtag(i))
                    Case "ICMT"    'comments
                        If filltrue(TextOther) Then
                            TextOther = InfoText
                        ElseIf filltrueAdd Then
                            TextOther = TextOther & ", " & InfoText
                        End If

                    End Select

                    'Debug.Print AVIInf.QueryInfo(mtag(i))
                End If        '0 text
            End If        '4
        Next i

        'поместить суммы, кроме пустых
        If Len(tmpLang) <> 0 Then
        If filltrue(TextLang) Then
            TextLang = tmpLang
        ElseIf filltrueAdd Then
            TextLang = TextLang & ", " & tmpLang
        End If
        End If
        
        If Len(tmpCountry) <> 0 Then
        If filltrue(TextCountry) Then
            TextCountry = tmpCountry
        ElseIf filltrueAdd Then
            TextCountry = TextCountry & ", " & tmpCountry
        End If
        End If
        
        If Len(tmpGenre) <> 0 Then
        If filltrue(TextGenre) Then
            TextGenre = tmpGenre
        ElseIf filltrueAdd Then
            TextGenre = TextGenre & ", " & tmpGenre
        End If
        End If
        
        If Right$(TextLang, 2) = ", " Then TextLang = Left$(TextLang, Len(TextLang) - 2)
        If Right$(TextCountry, 2) = ", " Then TextCountry = Left$(TextCountry, Len(TextCountry) - 2)
        If Right$(TextGenre, 2) = ", " Then TextGenre = Left$(TextGenre, Len(TextGenre) - 2)
    End If        ' AVIInf.GetInfoList(mtag) > 0 есть метаданные

Else            '                                      Append
    'Debug.Print "Append AVI"

    Frames = AVIInf.numFrames
    If Frames < 1 Then Exit Sub            'no frames

    TextTimeMSHid = Int(TextTimeMSHid) + AVIInf.PlayLength
    TextTimeHid = FormatTime(TextTimeMSHid)

    'size
    If isWindowsNt Then
        Pointer = lopen(aviName, OF_READ)
        GetFileSizeEx Pointer, lpFSHigh
        temp = Int(lpFSHigh * 10000 / 1024)
        lclose Pointer
        TextFilelenHid = Int(Val(TextFilelenHid)) + Int(Val(temp))
    Else
        TextFilelenHid = Int(Val(TextFilelenHid)) + Int(FileLen(aviName) / 1024)
    End If

    If NewDiskAddFlag Then TextCDN.Text = Replace(TextCDN, Val(TextCDN), Val(TextCDN) + 1)
    TextFileName.Text = TextFileName.Text & " | " & GetNameFromPathAndName(aviName)

End If

''Отразить аспекты на кнопках     ави - 1:1                                               4:3 16:9
For i = 0 To 2
optAspect(i).Value = False
optAspect(i).Enabled = False
Next i

'Позиции
Position.Min = 0
If Frames > 1 Then Position.Max = Frames - 1 Else Position.Max = 1

Position.TickFrequency = Position.Max / 100
Position.SmallChange = Position.TickFrequency / 5
Position.LargeChange = Position.Max / 100

PositionP.Min = 0         'PPMin
PPMax = Position.Value + 1000
If PPMax > Frames Then PPMax = Frames
PositionP.Max = PPMax
PositionP.SmallChange = 1         'PositionP.TickFrequency / 1
PositionP.LargeChange = 50         'PositionP.Max / 100

MovieWidth = ScaleX(movie.Width, vbTwips, vbPixels)
If Ratio < 1 Then Ratio = 1.333
MovieHeight = MovieWidth / Ratio
movie.Height = ScaleY(MovieHeight, vbPixels, vbTwips)

Set AVIInf = Nothing

If Not NoVideoProcess Then
    'видео окно
    Set m_cAVI = New cAVIFrameExtract
    m_cAVI.filename = aviName

    If Not aferror Then
        pRenderFrame 0
        lastRendedAVI = 0
        Position.Value = 0
        PositionP.Value = 0
        Position.Enabled = True
        PositionP.Enabled = True
        ComKeyNext.Enabled = True: ComKeyPrev.Enabled = True
        For i = 0 To 2: ComCap(i).Enabled = True: Next i

        
        ComAdd.Enabled = True        ': ComAddHid.Enabled = True
        ComRND(0).Enabled = True
        ComRND(1).Enabled = True
        ComRND(2).Enabled = True
        ComAutoScrShots.Enabled = True
        'For i = 0 To 2: optAspect(i).Enabled = True: Next i
        
    Else
        m_cAVI.filename = vbNullString        'unload

    End If

End If        'видео процесс

End Sub

Private Function DShGetInfoForCapture() As Boolean

Dim i As Integer
Dim temp As Currency
Dim Handle As Long

DShGetInfoForCapture = True    'если что, false и выход

'инфа из файла
On Error Resume Next
Handle = MediaInfo_New()
If err Then
            'If Not AutoNoMessFlag Then
                myMsgBox msgsvc(51) & App.Path & "\MediaInfo.dll", vbCritical
            'End If
    'MsgBox "MediaInfo.dll not found. Reinstall SurVideoCatalog!", vbCritical
    ToDebug "Err_NoDLL: MediaInfo.dll"
    DShGetInfoForCapture = False
    Exit Function
End If

Call MediaInfo_Open(Handle, StrPtr(DShName))
'ToDebug "MediaInfo: " & bstr(MediaInfo_Option(0, StrPtr("Info_Version"), StrPtr("")))
err.Clear: On Error GoTo 0

'видео
If MediaInfo_Count_Get(Handle, MediaInfo_Stream_Video, -1) > 0 Then
    MMI_Format_str = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
    If Len(MMI_Format_str) = 0 Then
        MMI_Format_str = "4/3"
        ToDebug "MMInfo_DS no format. = " & MMI_Format_str
    Else
        ToDebug "MMInfo_DS Format = " & MMI_Format_str
    End If
    MMI_Format = CalcFormat(MMI_Format_str)

Else    'нет видео
    MMI_Format_str = "4/3"
    MMI_Format = CalcFormat(MMI_Format_str)
    ToDebug "MMInfo не опознал файл. Format=4/3"
End If

If Handle <> 0 Then MediaInfo_Close Handle

'****************************************************DIRECT X *********************

If Not RenderAuto Then
    DShGetInfoForCapture = False
    Exit Function
End If

'тестовый запуск
mobjManager.Run
'objPosition.CurrentPosition = 0
Sleep 300    ' а то не видно после первого скролла
mobjManager.Pause

On Error Resume Next    'automation error на нулевых вобах и др

'Time
TimeL = objPosition.Duration

If err Then
    ToDebug err.Description
    DShGetInfoForCapture = False
    Exit Function
End If


If (TimeL = 0) Or (TimeL > 10000000) Then
    ToDebug "Error: неприемлемая длительность."
    'ClearVideo
    DShGetInfoForCapture = False
    Exit Function
End If

' aspect
'    If Format(MMI_Format, "0.000") = 1.333 Then 'тут может ошибиться в расчетах ()и затем
PixelRatio = objVideo.SourceHeight * MMI_Format / objVideo.SourceWidth
PixelRatioSS = ScrShotEd_W / MMI_Format

TimesX100 = TimeL * 100
temp = objVideo.AvgTimePerFrame * 100

Position.Min = 0
Position.Max = TimesX100     '- temp
Position.Value = 0
Position.TickFrequency = Position.Max / 100
Position.SmallChange = temp * 100      '0.04*100 *1000
Position.LargeChange = temp * 1000

PPMax = Position.Value + 10000     ' Const Range As Integer = 10000 в MpegPosScroll
If PPMax > TimesX100 Then PPMax = Position.Max     'TimesX100

PositionP.Min = 0     'PPMin
PositionP.TickFrequency = temp
PositionP.Max = PPMax
PositionP.SmallChange = temp / 2
PositionP.LargeChange = temp * 10
PositionP.Value = 0

Position.Enabled = True
PositionP.Enabled = True

ComKeyNext.Enabled = False: ComKeyPrev.Enabled = False
If MPGCaptured Then
    For i = 0 To 2
    ': ComCap(i).Enabled = True:
        ComRND(i).Enabled = True
    Next i
    ComAutoScrShots.Enabled = True
    DShGetInfoForCapture = True
    'Отразить аспекты на кнопках                                                    4:3 16:9
Call EdAspect2Buttons

Else
    For i = 0 To 2
        'ComCap(i).Enabled = False
        ComRND(i).Enabled = False
    Next i
    ComAutoScrShots.Enabled = False
    For i = 0 To 2: optAspect(i).Enabled = False: Next i
        DShGetInfoForCapture = False
End If

End Function


Private Function DShGetInfo() As Boolean
'еще DShGetInfoForCapture
Dim i As Integer
Dim temp As Currency    'Long
Dim tmps As String, tmp2s As String
Dim Handle As Long
Dim Pointer As Long, lpFSHigh As Currency
Dim tmp As String
Dim tmpL As Long

DShGetInfo = True    'если что, false и выход

'инфа из файла
On Error Resume Next
Handle = MediaInfo_New()
If err Then
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(51) & App.Path & "\MediaInfo.dll", vbCritical
            End If
    'MsgBox "MediaInfo.dll not found. Reinstall SurVideoCatalog!", vbCritical
    ToDebug "Err_NoDLL: MediaInfo.dll"
    DShGetInfo = False
    Exit Function
End If

Call MediaInfo_Open(Handle, StrPtr(DShName))
ToDebug "MediaInfo: " & bstr(MediaInfo_Option(0, StrPtr("Info_Version"), StrPtr("")))

'звук
TextAudioHid = vbNullString
tmpL = MediaInfo_Count_Get(Handle, MediaInfo_Stream_Audio, -1)
If tmpL > 0 Then
    For i = 0 To tmpL - 1
        tmps = vbNullString
        tmps = bstr(MediaInfo_Get(Handle, 2, i, StrPtr("SamplingRate"), 1, 0)) & " "
        tmps = tmps & Chnnls(bstr(MediaInfo_Get(Handle, 2, i, StrPtr("Channels"), 1, 0))) & " "
        tmps = tmps & bstr(MediaInfo_Get(Handle, 2, i, StrPtr("Codec"), 1, 0))

        tmp = bstr(MediaInfo_Get(Handle, 2, i, StrPtr("BitRate"), 1, 0))
        tmp = Replace(tmp, ".", ",")
        If IsNumeric(tmp) Then
            If Val(tmp) > 0 Then
                tmp = tmp / 1000 & "kbps"
                tmps = tmps & " (" & tmp & ")"
            End If
        End If

        TextAudioHid = tmps & ", " & TextAudioHid
    Next i
    TextAudioHid = Left$(TextAudioHid, Len(TextAudioHid) - 2)
ToDebug "DS_Audio>" & TextAudioHid

End If
err.Clear: On Error GoTo 0

'видео
If MediaInfo_Count_Get(Handle, MediaInfo_Stream_Video, -1) > 0 Then
    '?last14.12.2006 MMI_Ratio = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("PixelRatio"), 1, 0))
    '    MMI_Ratio = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio"), 1, 0))
    '    If Len(MMI_Ratio) = 0 Then
    '        '1выбор юзером аспект
    '        MMI_Ratio = "1.333"
    '        'NoAspectFlag = True
    '        ToDebug "MMI_Ratio_DS не найден. = " & MMI_Ratio
    '    Else
    '        ToDebug "MMI_Ratio_DS=" & MMI_Ratio
    '    End If

    'Debug.Print "MMI_Ratio_DS=" & MMI_Ratio
    'строка 4/3
    MMI_Format_str = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
    If Len(MMI_Format_str) = 0 Then
        MMI_Format_str = "4/3"
        ToDebug "MMInfo_DS не нашел Format. = " & MMI_Format_str
    Else
        ToDebug "MMInfo_DS Format = " & MMI_Format_str
        
    End If
    MMI_Format = CalcFormat(MMI_Format_str)

    tmps = vbNullString
    tmps = MyCodec(bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))) & " "
    MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))
ToDebug "DS_Video>" & MPGCodec

    tmp2s = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("FrameRate"), 1, 0))
    tmp2s = Replace(tmp2s, ".", ",")
    If IsNumeric(tmp2s) Then
        TextFPSHid = Val(tmp2s)
        Select Case Int(Val(tmp2s))
        Case "25"
            tmps = tmps & "PAL" & " "
        Case "29"
            tmps = tmps & "NTSC" & " "
        Case "23"    '23.976
            tmps = tmps & "FILM" & " "
        Case Else
            'tmps = tmps & TextFPSHid & " "
        End Select
    Else
        'mmi не нашли видео
        TextFPSHid = vbNullString
    End If
    
    'добавить аспект
    tmps = tmps & MMI_Format_str & " "
    tmps = Replace(tmps, "2.35", "2.35:1")
    
    'и битрейт
    tmp = vbNullString
    tmp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("BitRate"), 1, 0))
    tmp = Replace(tmp, ".", ",")
    If IsNumeric(tmp) Then
        If Val(tmp) > 0 Then
            tmp = tmp / 1000 & "kbps"
            tmps = tmps & "(" & tmp & ")"
        End If
    End If
    tmps = Replace(tmps, "16/9", "16:9")
    tmps = Replace(tmps, "4/3", "4:3")

    TextVideoHid = RTrim$(tmps)

Else
    '    MMI_Ratio = "1.333"
    MMI_Format_str = "4/3"
    MMI_Format = CalcFormat(MMI_Format_str)
    ToDebug "MMInfo не опознал файл. Format=4/3"
End If

If Handle <> 0 Then MediaInfo_Close Handle

'****************************************************DIRECT X *********************

If Not RenderAuto Then
    DShGetInfo = False
    Exit Function
End If

'тестовый запуск
mobjManager.Run
'objPosition.CurrentPosition = 0
Sleep 300    ' а то не видно после первого скролла

mobjManager.Pause

If Not AppendMovieFlag Then
    TextFileName.Text = GetNameFromPathAndName(DShName)
Else
    TextFileName.Text = TextFileName.Text & " | " & GetNameFromPathAndName(DShName)
End If

On Error Resume Next    'automation error на нулевых вобах и др

'Time
If Not AppendMovieFlag Then
    TextTimeMSHid = objPosition.Duration
Else
    TextTimeMSHid = TextTimeMSHid + objPosition.Duration
End If
TextTimeHid = FormatTime(TextTimeMSHid)
TimeL = objPosition.Duration

If err Then
    ToDebug err.Description
    DShGetInfo = False
    Exit Function
End If


If (TimeL = 0) Or (TimeL > 10000000) Then
    ToDebug "Error: получен неприемлемый размер видео."
    'ClearVideo
    DShGetInfo = False
    Exit Function
End If

'fps
If objVideo.AvgTimePerFrame > 0 Then
    TextFPSHid = Round(1 / objVideo.AvgTimePerFrame, 3)
End If

'file size

If isWindowsNt Then
    Pointer = lopen(DShName, OF_READ)
    GetFileSizeEx Pointer, lpFSHigh
    If Not AppendMovieFlag Then
        TextFilelenHid = Int(lpFSHigh * 10000 / 1024)
    Else
        TextFilelenHid = Val(TextFilelenHid) + Int(lpFSHigh * 10000 / 1024)
    End If
    lclose Pointer
Else
    If Not AppendMovieFlag Then
        TextFilelenHid = Int(FileLen(DShName) / 1024)
    Else
        TextFilelenHid = Val(TextFilelenHid) + Int(FileLen(DShName) / 1024)
    End If
End If

'FrSize
TextResolHid = Trim$(str$(objVideo.SourceWidth)) & " x " & Trim$(str$(objVideo.SourceHeight))

' aspect
'    If Format(MMI_Format, "0.000") = 1.333 Then
'тут может ошибиться в расчетах ()и затем
PixelRatio = objVideo.SourceHeight * MMI_Format / objVideo.SourceWidth
PixelRatioSS = ScrShotEd_W / MMI_Format


'cds
If Not AppendMovieFlag Then
    TextCDN.Text = 1
Else
    If NewDiskAddFlag Then
        TextCDN.Text = Replace(TextCDN, Val(TextCDN), Val(TextCDN) + 1)
    End If
End If

TimesX100 = TimeL * 100

temp = objVideo.AvgTimePerFrame * 100

Position.Min = 0
Position.Max = TimesX100     '- temp
Position.Value = 0
Position.TickFrequency = Position.Max / 100
Position.SmallChange = temp * 100      '0.04*100 *1000
Position.LargeChange = temp * 1000

PPMax = Position.Value + 10000     ' Const Range As Integer = 10000 в MpegPosScroll
If PPMax > TimesX100 Then PPMax = Position.Max     'TimesX100

PositionP.Min = 0     'PPMin
PositionP.TickFrequency = temp
PositionP.Max = PPMax
PositionP.SmallChange = temp / 2
PositionP.LargeChange = temp * 10
PositionP.Value = 0


Position.Enabled = True
PositionP.Enabled = True

ComKeyNext.Enabled = False: ComKeyPrev.Enabled = False
If MPGCaptured Then
    For i = 0 To 2: ComCap(i).Enabled = True: ComRND(i).Enabled = True: Next i
    ComAutoScrShots.Enabled = True
    
    'Отразить аспекты на кнопках                                                    4:3 16:9
Call EdAspect2Buttons

Else
    For i = 0 To 2: ComCap(i).Enabled = False: ComRND(i).Enabled = False: Next i
    ComAutoScrShots.Enabled = False
    For i = 0 To 2: optAspect(i).Enabled = False: Next i
End If

ComAdd.Enabled = True

End Function

Private Function MpgGetInfoForCapture() As Boolean
'это выполняется, когда mminfo опознал файл (нашел там видеопоток) и файл MPV

Dim i As Integer    ', j As Integer
Dim temp As Currency    'Long
Dim Handle As Long
'Dim tmps As String    ', tmp2s As String
Dim MMI_Height As Integer    'из MMInfo
Dim MMI_Width As Integer
Dim objv_Height As Integer    'из objVideo.Source
Dim objv_Width As Integer

'Dim ret As Long
'Dim WFD As WIN32_FIND_DATA
Dim ifo_handle As Long
Dim rendMPV2 As Boolean    'срендерили как мпег2 с нашими фильтрами
Dim IsVob As Boolean


MpgGetInfoForCapture = True

Handle = MediaInfo_New()
Call MediaInfo_Open(Handle, StrPtr(mpgName))
'ToDebug "MediaInfo: " & bstr(MediaInfo_Option(0, StrPtr("Info_Version"), StrPtr("")))
'инфа из файла
On Error Resume Next

'видео mminfo
'                           аспект
If MediaInfo_Count_Get(Handle, MediaInfo_Stream_Video, -1) > 0 Then

    'строка 4/3  mminfo
    MMI_Format_str = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
    If Len(MMI_Format_str) = 0 Then
        MMI_Format_str = "4/3"
        ToDebug "MMInfo no Format. = " & MMI_Format_str
    Else
        ToDebug "MMInfo Format = " & MMI_Format_str
    End If
    MMI_Format = CalcFormat(MMI_Format_str)

    MMI_Height = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Height"), 1, 0))
    MMI_Width = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Width"), 1, 0))
    'AspectRatio = GetAspectRatio(AspectRatioS, HeightS)
    MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))

End If

If Handle <> 0 Then MediaInfo_Close Handle

ToDebug "MMI_Format=" & MMI_Format

''оставить текущий!
'MMI_Format = GetAspectFromTextVideo
ToDebug "Ifo Format=" & MMI_Format

'****************************************************DIRECT X *********************
SVCDflag = False
If Not IsVob Then
    If MMI_Width = 480 And MMI_Height = 576 Then SVCDflag = True
    If MMI_Width = 480 And MMI_Height = 480 Then SVCDflag = True
End If
ToDebug "SVCD = " & SVCDflag

Select Case strMPEG
Case "MPV2"
    If Opt_UseOurMpegFilters And Not SVCDflag Then        'нашим декодером
        ToDebug "попытка RenderMPV2..."
        If RenderMPV2 Then
            rendMPV2 = True
        Else
            ToDebug "попытка RenderAuto..."
            If Not RenderAuto Then
                'все плохо
                If Not AutoNoMessFlag Then
                    myMsgBox msgsvc(10) & mpgName        'Ошибка работы с файлом
                End If
                MpgGetInfoForCapture = False
                Exit Function
            End If
        End If
    Else        'сразу авто
        ToDebug "сразу RenderAuto..."
        If Not RenderAuto Then
            'все плохо
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(10) & mpgName
            End If
            MpgGetInfoForCapture = False
            Exit Function
        End If
    End If

Case "MPV1"
    If RenderMPV1 Then
        'rendMPV1 = True
    Else
        If Not RenderAuto Then
            'все плохо
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(10) & mpgName
            End If
            MpgGetInfoForCapture = False
            Exit Function
        End If
    End If

End Select

'тестовый запуск
FrAdEdPixHid.Visible = True
'movie.Visible = False

ToDebug "Run..."
On Error Resume Next
mobjManager.Run
Sleep 300
mobjManager.Pause
ToDebug "Pause. Err=" & err.Number
err.Clear: On Error GoTo 0
'MPGPosScroll

ToDebug "Editor/FileName=" & mpgName

On Error Resume Next        'automation error на нулевых вобах

TimeL = objPosition.Duration

'Debug.Print "TimeL = " & TimeL
If (TimeL = 0) Or (TimeL > 10000000) Then
    'ClearVideo
    ToDebug msgsvc(40) & " = " & TimeL
    If Not AutoNoMessFlag Then myMsgBox msgsvc(40) & " = " & TimeL, vbCritical, , FrmMain.hWnd
    MpgGetInfoForCapture = False
    Exit Function
End If

'                                                                 Frame Size
objv_Width = objVideo.SourceWidth
objv_Height = objVideo.SourceHeight

'                                                                   aspect
PixelRatio = 1.333: PixelRatioSS = ScrShotEd_W / (4 / 3)        'дефолт 4:3
'MMI_Format не менять без MMI_Format_str

'то же MpgGetInfo
PixelRatio = (objv_Height * MMI_Format) / objv_Width
PixelRatioSS = ScrShotEd_W / MMI_Format

'Select Case VideoStand
'Case "PAL"
'    Select Case strMPEG
'
'    Case "MPV2"
'        If SVCDflag Then
'            'svcd
'            PixelRatio = objv_Height * MMI_Format / objv_Width
'            PixelRatioSS = ScrShotEd_W / MMI_Format
'        Else
'            'dvd
'            PixelRatio = objv_Height * MMI_Format / objv_Width
'            PixelRatioSS = ScrShotEd_W / MMI_Format
'        End If
'
'    Case "MPV1"
'        'vcd
'        MMI_Format = 4 / 3: MMI_Format_str = "4/3"
'        PixelRatio = (objv_Height * MMI_Format) / objv_Width
'        'дефолт PixelRatioSS = ScrShotEd_W / (4 / 3)
'    End Select
'
'Case "NTSC"
'    Select Case strMPEG
'    Case "MPV2"
'        If SVCDflag Then
'            'svcd
'            PixelRatio = objv_Height * MMI_Format / objv_Width
'            PixelRatioSS = ScrShotEd_W / MMI_Format
'        Else
'            'dvd
'            PixelRatio = objv_Height * MMI_Format / objv_Width
'            PixelRatioSS = ScrShotEd_W / MMI_Format
'        End If
'    Case "MPV1"
'        'vcd
'        MMI_Format = 4 / 3: MMI_Format_str = "4/3"
'        PixelRatio = (objv_Height * MMI_Format) / objv_Width
'        'PixelRatio = (objv_Height * 4 / 3) / objv_Width
'        'дефолт PixelRatioSS = ScrShotEd_W / (4 / 3)
'    End Select
'End Select


TimesX100 = TimeL * 100
temp = objVideo.AvgTimePerFrame * 100

Position.Min = 0
Position.Max = TimesX100        '- temp
Position.Value = 0

Position.TickFrequency = Position.Max / 100
Position.SmallChange = temp * 100        '0.04*100 *1000
Position.LargeChange = temp * 1000

PPMax = Position.Value + 10000        ' Const Range As Integer = 10000 в MpegPosScroll
If PPMax > TimesX100 Then PPMax = Position.Max        ' TimesX100

PositionP.Min = 0        'PPMin
PositionP.TickFrequency = temp
PositionP.Max = PPMax
PositionP.SmallChange = temp
PositionP.LargeChange = temp * 10
PositionP.Value = 0
Position.Enabled = True: PositionP.Enabled = True

ComKeyNext.Enabled = False: ComKeyPrev.Enabled = False

If MPGCaptured Then
    For i = 0 To 2: ComCap(i).Enabled = True: ComRND(i).Enabled = True: Next i
    ComAutoScrShots.Enabled = True
    MpgGetInfoForCapture = True

    'Отразить аспекты на кнопках                                                    4:3 16:9
    Call EdAspect2Buttons

Else
    For i = 0 To 2
        ': ComCap(i).Enabled = False:
        ComRND(i).Enabled = False
    Next i
    ComAutoScrShots.Enabled = False
    For i = 0 To 2: optAspect(i).Enabled = False: Next i
    MpgGetInfoForCapture = False
End If


End Function


Private Function MpgGetInfo() As Boolean
'MpgGetInfoForCapture
'это выполняется, когда mminfo опознал файл (нашел там видеопоток) и файл MPV

Dim WeGetTimeFromIfo As Boolean

Dim i As Integer, j As Integer
Dim temp As Currency    'Long
Dim Handle As Long
Dim TimeS As String
Dim tmps As String, tmp2s As String
Dim tmp As String, tmp2 As String
Dim tmpL As Long
Dim tmpa As String    'битр аудио
Dim MMI_Height As Integer    'из MMInfo
Dim MMI_Width As Integer
Dim objv_Height As Integer    'из objVideo.Source
Dim objv_Width As Integer

Dim ret As Long
Dim WFD As WIN32_FIND_DATA

'Dim ifo_fname As String
Dim ifo_handle As Long
Dim length_pgc(3) As Byte
Dim NumIFOChains As Long
Dim NumCells As Long

Dim AllVobsSize As String
Dim Pointer As Long, lpFSHigh As Currency
Dim rendMPV2 As Boolean    'срендерили как мпег2 с нашими фильтрами

Dim vstr_NumSubPic As Long
Dim IsVob As Boolean


MpgGetInfo = True

If Not AppendMovieFlag Then
    TextTimeHid = vbNullString
    TextFPSHid = vbNullString
End If

'поиск файла ifo для данного воба
If StrComp(Right$(mpgName, 3), "vob", vbTextCompare) = 0 Then   'vob
    IsVob = True
    tmps = Left$(mpgName, Len(mpgName) - 5) & "0.ifo"
    ret = FindFirstFile(tmps, WFD)
    If ret > 0 Then
        'open
        On Error Resume Next
        ifo_handle = ifoOpen(tmps, fio_USE_ASPI)
        If ifo_handle = 0 Then
            ToDebug ("ERR_vstrip.dll, не открыть: " & tmps)
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(51) & App.Path & "\vstrip.dll", vbCritical
            End If
        End If
        On Error GoTo 0
    End If
End If

Handle = MediaInfo_New()
Call MediaInfo_Open(Handle, StrPtr(mpgName))
ToDebug "MediaInfo: " & bstr(MediaInfo_Option(0, StrPtr("Info_Version"), StrPtr("")))

If ifo_handle = 0 Then

    If Not AppendMovieFlag Then
        'инфа из файла

        On Error Resume Next

        'звук
        TextAudioHid = vbNullString
        tmpL = MediaInfo_Count_Get(Handle, MediaInfo_Stream_Audio, -1)
        If tmpL > 0 Then
            For i = 0 To tmpL - 1
                tmps = vbNullString
                tmps = bstr(MediaInfo_Get(Handle, 2, i, StrPtr("SamplingRate"), 1, 0)) & " "
                tmps = tmps & Chnnls(bstr(MediaInfo_Get(Handle, 2, i, StrPtr("Channels"), 1, 0))) & " "
                tmps = tmps & bstr(MediaInfo_Get(Handle, 2, i, StrPtr("Codec"), 1, 0)) & " "

                tmp = bstr(MediaInfo_Get(Handle, 2, i, StrPtr("BitRate"), 1, 0))
                tmp = Replace(tmp, ".", ",")
                If IsNumeric(tmp) Then
                    If Val(tmp) > 0 Then
                        tmp = tmp / 1000 & "kbps"
                        tmps = tmps & "(" & tmp & ")"
                    End If
                End If

                TextAudioHid = tmps & ", " & TextAudioHid
            Next i
            TextAudioHid = Left$(TextAudioHid, Len(TextAudioHid) - 2)
            TextAudioHid = Trim$(TextAudioHid)
            
        Else
            ToDebug "MMInfo не нашел звук"
        End If
        err.Clear: On Error GoTo 0

    End If    'append

    'видео mminfo
    '                           аспект
    If MediaInfo_Count_Get(Handle, MediaInfo_Stream_Video, -1) > 0 Then

        'строка 4/3  mminfo
        MMI_Format_str = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
        If Len(MMI_Format_str) = 0 Then
            MMI_Format_str = "4/3"
            ToDebug "MMInfo не нашел Format. = " & MMI_Format_str
        Else
            ToDebug "MMInfo Format = " & MMI_Format_str
        End If
        'MMI_Format = CalcFormat(bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0)))
        MMI_Format = CalcFormat(MMI_Format_str)

        MMI_Height = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Height"), 1, 0))
        MMI_Width = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Width"), 1, 0))
        'AspectRatio = GetAspectRatio(AspectRatioS, HeightS)
        tmps = vbNullString
        'Call MediaInfo_Option(Handle, StrPtr("Complete"), StrPtr("1"))
        MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))
        tmps = MyCodec(MPGCodec) & " "
        

        If Not AppendMovieFlag Then
            'fps  mminfo
            'tmps = tmps & AspectRatioS & " "
            tmp2s = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("FrameRate"), 1, 0))
            TextFPSHid = Val(tmp2s)

            Select Case Int(Val(tmp2s))
            Case "25"
                tmps = tmps & "PAL" & " "
            Case "29"        '29.97
                tmps = tmps & "NTSC" & " "
            Case "23"        '23.976
                tmps = tmps & "FILM" & " "
            Case Else
                tmps = tmps & tmp2s & " "
            End Select

    'добавить аспект
            tmps = tmps & MMI_Format_str & " "
    'и битрейт
            tmp = vbNullString
            tmp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("BitRate"), 1, 0))
            tmp = Replace(tmp, ".", ",")
            If IsNumeric(tmp) Then
                If Val(tmp) > 0 Then
                    tmp = tmp / 1000 & "kbps"
                    tmps = tmps & "(" & tmp & ")"
                End If
            End If
            tmps = Replace(tmps, "16/9", "16:9")
            tmps = Replace(tmps, "4/3", "4:3")
            tmps = Replace(tmps, " 0.000", vbNullString)
            TextVideoHid = RTrim$(tmps)
        End If

    End If    'append

    'MediaInfo_Close Handle

Else                                                       'инфу с ifo
    ToDebug "vstrip, информация из " & tmps

    If Not AppendMovieFlag Then
        'video ifo
        TextVideoHid = Trim$(IFObstr(ifoGetVideoDesc(ifo_handle)))

        'bitrate ifo
        '        Handle = MediaInfo_New(): Call MediaInfo_Open(Handle, StrPtr(mpgName))
        tmp = vbNullString
        tmp = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("BitRate"), 1, 0))
        tmp = Replace(tmp, ".", ",")
        If IsNumeric(tmp) Then tmp = tmp / 1000 & "kbps"
        If Len(tmp) > 0 Then TextVideoHid = TextVideoHid & " (" & tmp & ")"

        'fps ifo
        TextFPSHid = vbNullString
        If InStr(1, TextVideoHid, "PAL", vbTextCompare) > 0 Then TextFPSHid = "25"
        If InStr(1, TextVideoHid, "NTSC", vbTextCompare) > 0 Then TextFPSHid = "29.97"

        'audio & Language ifo
        j = ifoGetNumAudio(ifo_handle)
        If j > 0 Then
            tmps = vbNullString: tmp2s = vbNullString
            For i = 0 To j - 1
                tmp = vbNullString
                tmp = IFObstr(ifoGetAudioDesc(ifo_handle, i))
                'bitrate audio
                tmpa = vbNullString
                tmpa = bstr(MediaInfo_Get(Handle, MediaInfo_Stream_Audio, i, StrPtr("BitRate"), 1, 0))
                tmpa = Replace(tmpa, ".", ",")
                If IsNumeric(tmpa) Then tmpa = tmpa / 1000 & "kbps"
                If Len(tmpa) > 0 Then tmp = tmp & " (" & tmpa & ")"

                'Debug.Print "VStrip Audio: " & tmp
                'Debug.Print "Lang: " & IFObstr(ifoGetLangDesc(ifo_handle, i))
                tmp2 = IFObstr(ifoGetLangDesc(ifo_handle, i))
                'tmp = Replace(tmp, "[0]", vbNullString)
                'tmp = Left$(tmp, InStr(1, tmp, "ch,", vbTextCompare) + 1) & ")"
                tmps = tmps & tmp & ", "
                tmp2s = tmp2s & tmp2 & ", "
            Next i
            'Lang
            If Len(tmps) > 2 Then TextAudioHid = Trim$(Left$(tmps, Len(tmps) - 2))
            If Len(tmp2s) > 2 Then tmp2s = Trim$(Left$(tmp2s, Len(tmp2s) - 2))


            If filltrue(TextLang) Then
                TextLang = CountryLocal(tmp2s)
            ElseIf filltrueAdd Then
                TextLang = TextLang & ", " & CountryLocal(tmp2s)
            End If
        End If

        'SubTitle ifo
        tmp = vbNullString
        vstr_NumSubPic = ifoGetNumSubPic(ifo_handle)
        For i = 0 To vstr_NumSubPic - 1
            If i > 0 Then
                tmp = tmp & ", " & IFObstr(ifoGetSubPicDesc(ifo_handle, i))
            Else
                tmp = IFObstr(ifoGetSubPicDesc(ifo_handle, i))
            End If
        Next i

        If filltrue(TextSubt) Then
            TextSubt = CountryLocal(tmp)
        ElseIf filltrueAdd Then
            TextSubt = TextSubt & ", " & CountryLocal(tmp)
        End If

    End If    'not append

    '                                                           время фильма с ifo
    NumIFOChains = ifoGetNumPGCI(ifo_handle)

    If AppendMovieFlag Then TextTimeHid = TextTimeHid & ", " 'потом, если возможно, суммируется, если нет  (1 ифо на несколько фильмов ?) то добавиться

    For i = 0 To NumIFOChains - 1
        NumCells = ifoGetPGCIInfo(ifo_handle, i, length_pgc(0))    ' нужен length_pgc(0)
        TextTimeHid = TextTimeHid & Format$(length_pgc(0), "0#") & ":" & Format$(length_pgc(1), "0#") & ":" & Format$(length_pgc(2), "0#") & ", "    ' & "," & Format$(length_pgc(3), "0#")
        'Debug.Print FormatTime(TextTimeMSHid)
    Next i
    
    'заменим TextTimeHid, если можно суммировать
    If i > 1 Then    '1 ифо на несколько фильмов
        TextTimeMSHid = 0 'чтобы не суммировать потом DS
        WeGetTimeFromIfo = True
        'нажать плюсик
        Call ComPlusHid_Click(2)
    Else
        If Not AppendMovieFlag Then
            TextTimeHid = Left$(TextTimeHid, Len(TextTimeHid) - 2)
            TextTimeMSHid = length_pgc(0) * 3600 + length_pgc(1) * 60 + length_pgc(2)
            WeGetTimeFromIfo = True
        Else
            'суммировать
            TextTimeMSHid = TextTimeMSHid + length_pgc(0) * 3600 + length_pgc(1) * 60 + length_pgc(2)
            TextTimeHid = FormatTime(TextTimeMSHid)
            WeGetTimeFromIfo = True
        End If
    End If
    'убрать последние ", "
    If Right$(TextTimeHid, 2) = ", " Then TextTimeHid = Left$(TextTimeHid, Len(TextTimeHid) - 2)

    'close ifo
    If ifoClose(ifo_handle) Then ToDebug "ifo закрыт."

    If Not AppendMovieFlag Then
        If Len(TextFPSHid) = 0 Then    'если попытка с ifo не прошла
            'fps mminfo
            TextFPSHid = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("FrameRate"), 1, 0))
            'HeightS = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Height"), 1, 0))
            'WidthS = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Width"), 1, 0))
        End If
    End If

    MPGCodec = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Codec"), 1, 0))

'    MMI_Ratio = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio"), 1, 0))
'    If Len(MMI_Ratio) = 0 Then MMI_Ratio = "1.333" 'было 1
    
    'Debug.Print "MMI_Format=" & bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
    '4/3=?
    'MMI_Format = 1
    
        MMI_Format_str = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("AspectRatio/String"), 1, 0))
        If Len(MMI_Format_str) = 0 Then
            MMI_Format_str = "4/3"
            ToDebug "MMInfo не нашел Format. = " & MMI_Format_str
        Else
            ToDebug "MMInfo Format = " & MMI_Format_str
        End If
    MMI_Format = CalcFormat(MMI_Format_str)
    
    MMI_Height = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Height"), 1, 0))
    MMI_Width = bstr(MediaInfo_Get(Handle, 1, 0, StrPtr("Width"), 1, 0))
        
End If    ' ifo

If Handle <> 0 Then MediaInfo_Close Handle

'Debug.Print "MMI_Ratio=" & MMI_Ratio
'ToDebug "MMI_Ratio=" & MMI_Ratio
'Debug.Print "MMI_Width=" & MMI_Width & " MMI_Height=" & MMI_Height

ToDebug "MMI_Format=" & MMI_Format
'взять MMI_Format из поля видео в редакторе, если нет - оставить текущий
MMI_Format = GetAspectFromTextVideo
ToDebug "Ifo Format=" & MMI_Format

'****************************************************DIRECT X *********************
SVCDflag = False
If Not IsVob Then
    If MMI_Width = 480 And MMI_Height = 576 Then SVCDflag = True
    If MMI_Width = 480 And MMI_Height = 480 Then SVCDflag = True
End If
ToDebug "SVCD = " & SVCDflag
ToDebug "MMI_Codec: " & MPGCodec

Select Case strMPEG
Case "MPV2"
    If Opt_UseOurMpegFilters And Not SVCDflag Then        'нашим декодером
        ToDebug "попытка RenderMPV2..."
        If RenderMPV2 Then
            rendMPV2 = True
        Else
            ToDebug "попытка RenderAuto..."
            If Not RenderAuto Then
                'все плохо

                If Not AutoNoMessFlag Then
                    myMsgBox msgsvc(10) & mpgName        'Ошибка работы с файлом
                End If
                MpgGetInfo = False
                Exit Function
            End If
        End If
    Else        'сразу авто
        ToDebug "сразу RenderAuto..."
        If Not RenderAuto Then
            'все плохо
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(10) & mpgName
            End If
            MpgGetInfo = False
            Exit Function
        End If
    End If

Case "MPV1"
    If RenderMPV1 Then
        'rendMPV1 = True
    Else
        If Not RenderAuto Then
            'все плохо
            If Not AutoNoMessFlag Then
                myMsgBox msgsvc(10) & mpgName
            End If
            MpgGetInfo = False
            Exit Function
        End If
    End If

Case Else
    MpgGetInfo = False

Exit Function
            
End Select

'тестовый запуск
FrAdEdPixHid.Visible = True
'movie.Visible = False

ToDebug "Run..."
On Error Resume Next
mobjManager.Run
Sleep 300
mobjManager.Pause

ToDebug "Pause. Err=" & err.Number

err.Clear: On Error GoTo 0
'MPGPosScroll
'________________________________


If StrComp(Right$(mpgName, 3), "vob", vbTextCompare) = 0 Then  'сменить vob на ifo
    tmps = Left$(mpgName, Len(mpgName) - 5) & "0.ifo"
    ret = FindFirstFile(tmps, WFD)
    If ret > 0 Then
        ToDebug "Структура DVD."
        'имена файлов
        If Not AppendMovieFlag Then
            TextFileName.Text = GetNameFromPathAndName(tmps)    'ifo
        Else
            TextFileName.Text = TextFileName.Text & " | " & GetNameFromPathAndName(tmps)    'ifo
        End If
        'найти все вобы фильма и их размер
        AllVobsSize = GetAllVobsSize(GetPathFromPathAndName(tmps), GetNameFromPathAndName(tmps)) 'нет! TextFileName.Text)
        ToDebug "VobsSize=" & AllVobsSize
    Else
        If Not AppendMovieFlag Then
            TextFileName.Text = GetNameFromPathAndName(mpgName)
        Else
            TextFileName.Text = TextFileName.Text & " | " & GetNameFromPathAndName(mpgName)
        End If
    End If
Else
    If Not AppendMovieFlag Then
        TextFileName.Text = GetNameFromPathAndName(mpgName)
    Else
        TextFileName.Text = TextFileName.Text & " | " & GetNameFromPathAndName(mpgName)
    End If
End If
ToDebug "Editor/FileName=" & TextFileName.Text

'показывает из настроек декодера, те может быть неверно (50 вместо 25...)
'TextFPSHid = Round(1 / objVideo.AvgTimePerFrame, 3)
'FrameS = 1 / objVideo.AvgTimePerFrame * objPosition.Duration
'If FrameS < 1 Then Exit Sub  'no frames

On Error Resume Next    'automation error на нулевых вобах

'Time dx - 1 файл
If Not WeGetTimeFromIfo Then
If Not AppendMovieFlag Then
'    If Len(TextTimeHid) = 0 Then    'не узнали из ifo
        TextTimeMSHid = objPosition.Duration
        TimeS = FormatTime(TextTimeMSHid)    'mminfo.bas
        TextTimeHid = TimeS
'    End If
Else
    'суммировать при добавлении если было к чему
    If TextTimeMSHid <> 0 Then
        TextTimeMSHid = TextTimeMSHid + objPosition.Duration
        TextTimeHid = FormatTime(TextTimeMSHid)
    End If
End If
End If

TimeL = objPosition.Duration

'Debug.Print "TimeL = " & TimeL
If (TimeL = 0) Or (TimeL > 10000000) Then
    'ClearVideo
    ToDebug msgsvc(40) & " = " & TimeL
    If Not AutoNoMessFlag Then myMsgBox msgsvc(40) & " = " & TimeL, vbCritical, , FrmMain.hWnd
    MpgGetInfo = False
    Exit Function
End If

'file size
If Len(AllVobsSize) = 0 Then
    'посчитать размер файла
    If isWindowsNt Then
        Pointer = lopen(mpgName, OF_READ)
        GetFileSizeEx Pointer, lpFSHigh
        If Not AppendMovieFlag Then
            TextFilelenHid = Int(lpFSHigh * 10000 / 1024)
        Else
            TextFilelenHid = Val(TextFilelenHid) + Int(lpFSHigh * 10000 / 1024)
        End If
        lclose Pointer
    Else
        If Not AppendMovieFlag Then
            TextFilelenHid = Int(FileLen(mpgName) / 1024)
        Else
            TextFilelenHid = Val(TextFilelenHid) + Int(FileLen(mpgName) / 1024)
        End If
    End If
Else
    'или подставить размер всех вобов
    If Not AppendMovieFlag Then
        TextFilelenHid = AllVobsSize
    Else
        TextFilelenHid = Val(TextFilelenHid) + Val(AllVobsSize)
    End If
End If

'                                                                 Frame Size
objv_Width = objVideo.SourceWidth
objv_Height = objVideo.SourceHeight

If Not AppendMovieFlag Then
    TextResolHid = Trim$(str$(objv_Width)) & " x " & Trim$(str$(objv_Height))
End If

'                                                                   aspect
'PixelRatio = 1.333: PixelRatioSS = ScrShotEd_W / (4 / 3)    'дефолт 4:3
'PixelRatioSS = ScrShotEd_W / (4 / 3)    'дефолт 4:3

'то же MpgGetInfoForCapture
PixelRatio = (objv_Height * MMI_Format) / objv_Width
PixelRatioSS = ScrShotEd_W / MMI_Format


'MMI_Format не менять без MMI_Format_str
'Select Case VideoStand
'Case "PAL"
'    Select Case strMPEG
'
'    Case "MPV2"
'        If SVCDflag Then
'            'svcd
'            PixelRatio = objv_Height * MMI_Format / objv_Width
'            PixelRatioSS = ScrShotEd_W / MMI_Format
'        Else
'            'dvd
'            PixelRatio = objv_Height * MMI_Format / objv_Width
'            PixelRatioSS = ScrShotEd_W / MMI_Format
'        End If
'
'    Case "MPV1"
'        'vcd
'        MMI_Format = 4 / 3: MMI_Format_str = "4/3"
'        PixelRatio = (objv_Height * MMI_Format) / objv_Width
'        'дефолт PixelRatioSS = ScrShotEd_W / (4 / 3)
'    End Select
'
'Case "NTSC"
'    Select Case strMPEG
'    Case "MPV2"
'        If SVCDflag Then
'            'svcd
'            PixelRatio = objv_Height * MMI_Format / objv_Width
'            PixelRatioSS = ScrShotEd_W / MMI_Format
'        Else
'            'dvd
'            PixelRatio = objv_Height * MMI_Format / objv_Width
'            PixelRatioSS = ScrShotEd_W / MMI_Format
'        End If
'    Case "MPV1"
'        'vcd
'        MMI_Format = 4 / 3: MMI_Format_str = "4/3"
'        PixelRatio = (objv_Height * MMI_Format) / objv_Width
'        'PixelRatio = (objv_Height * 4 / 3) / objv_Width
'        'дефолт PixelRatioSS = ScrShotEd_W / (4 / 3)
'    End Select
'End Select

'Debug.Print "PixelRatio=" & PixelRatio, "PixelRatioSS=" & PixelRatioSS
'ToDebug "AspectFull=" & PixelRatio & " AspectMini=" & PixelRatioSS
'TextResolHid = WidthS & " x " & HeightS

'cds
If Not AppendMovieFlag Then
    TextCDN.Text = 1
Else
    If NewDiskAddFlag Then
        TextCDN.Text = Replace(TextCDN, Val(TextCDN), Val(TextCDN) + 1)
    End If
End If

TimesX100 = TimeL * 100

temp = objVideo.AvgTimePerFrame * 100

Position.Min = 0
Position.Max = TimesX100     '- temp
Position.Value = 0

Position.TickFrequency = Position.Max / 100
Position.SmallChange = temp * 100     '0.04*100 *1000
Position.LargeChange = temp * 1000

PPMax = Position.Value + 10000     ' Const Range As Integer = 10000 в MpegPosScroll
If PPMax > TimesX100 Then PPMax = Position.Max     ' TimesX100

PositionP.Min = 0     'PPMin
PositionP.TickFrequency = temp
PositionP.Max = PPMax
PositionP.SmallChange = temp
PositionP.LargeChange = temp * 10
PositionP.Value = 0

Position.Enabled = True: PositionP.Enabled = True

ComKeyNext.Enabled = False: ComKeyPrev.Enabled = False

If MPGCaptured Then
    For i = 0 To 2: ComCap(i).Enabled = True: ComRND(i).Enabled = True: Next i
    ComAutoScrShots.Enabled = True
    
    'Отразить аспекты на кнопках                                                    4:3 16:9
    Call EdAspect2Buttons

Else
    For i = 0 To 2: ComCap(i).Enabled = False: ComRND(i).Enabled = False: Next i
    ComAutoScrShots.Enabled = False
    For i = 0 To 2: optAspect(i).Enabled = False: Next i
End If

'ComAdd.Enabled = False
ComAdd.Enabled = True

End Function

Private Sub KeyNext()
lastRendedAVI = m_cAVI.AVIStreamNearestNextKeyFrame(lastRendedAVI)
Position.Value = lastRendedAVI
PosScroll
End Sub
Private Sub KeyPrev()
lastRendedAVI = m_cAVI.AVIStreamNearestPrevKeyFrame(lastRendedAVI)
'pRenderFrame lastRendedAVI
Position.Value = lastRendedAVI
PosScroll
End Sub

Private Sub CDSerialCur_Change()
Mark2Save
End Sub

Private Sub ChInFilFl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static working As Boolean
If working Then Exit Sub
working = True
Select Case ChInFilFl.Value
Case Unchecked
    ChInFilFl.Value = Checked
Case Gray
    ChInFilFl.Value = Unchecked
Case Checked
    ChInFilFl.Value = Gray
End Select
working = False
ReleaseCapture
End Sub

Private Sub ComAdd_Click()
'Append new file
Dim tmpdrive As String
Dim tmpSerial As String

aviName = pLoadDialog(ComAdd.Caption)
DoEvents

If LenB(aviName) = 0 Then Exit Sub

ToDebug "Добавление."

'Серийник
NewDiskAddFlag = False    'не новый носитель
tmpdrive = Left$(LCase$(aviName), 3)
If DriveType(tmpdrive) = "CD-ROM" Then
    tmpSerial = Hex$(GetSerialNumber(tmpdrive))
    If tmpSerial <> "0" Then
    
        If MediaSN = tmpSerial Then
            IsSameCdFlag = True 'это тот же носитель
        Else
            IsSameCdFlag = False
            MediaSN = tmpSerial 'запомнить серийник cd
        End If

        If InStr(1, CDSerialCur, tmpSerial, vbTextCompare) = 0 Then
            'если серийник другой
            If Len(CDSerialCur) = 0 Then
                'установить серийник
                CDSerialCur = tmpSerial
            Else
                'иначе добавить следуюший
                CDSerialCur = CDSerialCur & ", " & tmpSerial
                NewDiskAddFlag = True
            End If
        End If    'instr
    End If
    
        'Носитель
        If IsSameCdFlag Then 'если тот же носитель
            ComboNos.Text = MediaType
        Else
            ComboNos.Text = ComboNos.Text & ", " & GetOptoInfo(Left$(aviName, 2)) 'добавить тип
            MediaType = ComboNos.Text
        End If
        
Else
    'ComboNos.Text = vbNullString 'стирает то что было
End If    'drivetype


OpenAddmovFlag = True

AppendMovieFlag = True
ComOpen_Click
AppendMovieFlag = False
'Call SubAdd

End Sub

Public Sub ComAutoScrShots_Click()
Dim tmp As Currency
'Dim tmpsi As Single

Mark2Save
Screen.MousePointer = vbHourglass

'                                                               AVI
If isAVIflag And (Not aferror) Then
    AutoScrShots (Frames)
End If

'                                                               MPEG
If isMPGflag Or isDShflag Then

    'objPosition.CurrentPosition = Position.Value / 100
    tmp = (TimesX100 - TimesX100 * 0.05) / 3    '4
    pos1 = 1 + (Rnd() * tmp)    '1-4
    pos2 = tmp + (Rnd() * tmp)    '4-8
    pos3 = tmp * 2 + (Rnd() * tmp)    '8-12

    'Position.Value = pos3
    MPGPosScroll

    On Error Resume Next    'если ошибается objPosition.CurrentPosition, в меню например

    Position.Value = pos1
    objPosition.CurrentPosition = Position.Value / 100
    Sleep 500
    ComCap_Click 0
    Screen.MousePointer = vbHourglass

    PicSS1.Refresh: DoEvents
    Sleep 500
    Position.Value = pos2
    objPosition.CurrentPosition = Position.Value / 100
    Sleep 500
    ComCap_Click 1
    Screen.MousePointer = vbHourglass

    PicSS2.Refresh: DoEvents
    Sleep 500
    Position.Value = pos3
    objPosition.CurrentPosition = Position.Value / 100
    Sleep 500
    ComCap_Click 2

End If    'mpeg

SavePic1Flag = True: SavePic2Flag = True: SavePic3Flag = True
NoPic1Flag = False: NoPic2Flag = False: NoPic3Flag = False

If Not frmAutoFlag Then Me.SetFocus

Screen.MousePointer = vbNormal
End Sub

Private Sub ComboNos_Change()
Mark2Save
End Sub

Private Sub ComboNos_Click()
Mark2Save
End Sub

Private Sub ComboNos_LostFocus()
ComboNos.Text = sTrimChars(ComboNos.Text, vbNewLine)
End Sub

Private Sub ComClsEd_Click()
'очистка полей вкладки фильмы, запись в автофил

AutoFillStore
DoEvents

TextMName.SetFocus: SendKeys "{Del}", True
TextLabel.SetFocus: SendKeys "{Del}", True
TextGenre.SetFocus: SendKeys "{Del}", True
TextCountry.SetFocus: SendKeys "{Del}", True
TextYear.SetFocus: SendKeys "{Del}", True
TextAuthor.SetFocus: SendKeys "{Del}", True
TextOther.SetFocus: SendKeys "{Del}", True

TextLang.SetFocus: SendKeys "{Del}", True
TextSubt.SetFocus: SendKeys "{Del}", True
TextRate.SetFocus: SendKeys "{Del}", True

TextRole.SetFocus: TextRole.SelStart = 0: TextRole.SelLength = Len(TextRole.Text)
SendKeys "{Del}", True ''SendKeys "^A": win2000?

TextAnnotation.SetFocus:  TextAnnotation.SelStart = 0: TextAnnotation.SelLength = Len(TextAnnotation.Text)
SendKeys "{Del}", True ''SendKeys "^A": win2000?

TextMName.SetFocus
End Sub

Private Sub ComGetCover_Click()
'получить картинку из инета по абсолютному URL
Dim tmp As String

'pix
If Len(TextCoverURL) <> 0 Then
    '        tmp = UrlEncode(cBasePicURL & TextCoverURL)
    tmp = cBasePicURL & TextCoverURL
    'Debug.Print tmp
    OpenURLProxy tmp, "pic"
Else
    'Set ImgPrCov = Nothing: Set PicFrontFace = Nothing: Set picCanvas = Nothing
    'NoPicFrontFaceFlag = True
End If
End Sub

Private Sub ComInterGoHid_Click(Index As Integer)

Dim temp As Long
Dim strPath As String
Dim site As String

strPath = Space$(255)

Select Case Index
Case 0
    site = ComboSites.Text
Case 1
    site = TextMovURL.Text
End Select

temp = FindExecutable(site, "", strPath)
Select Case temp
Case 31
    myMsgBox msgsvc(17), vbInformation, , Me.hWnd
    Exit Sub
Case 2
End Select

temp = ShellExecute(GetDesktopWindow(), "open", site, vbNull, vbNull, 1)
ToDebug "InterGo_ret=" & temp
End Sub


Private Sub ComPlusHid_Click(Index As Integer)
Select Case Index

Case 0
    If LenB(Trim$(ComboGenre.Text)) = 0 Then Exit Sub    ''не добавлять пустоту с запятыми
    If LenB(Trim$(TextGenre.Text)) = 0 Then
        TextGenre.Text = ComboGenre.Text
    Else
        TextGenre.Text = TextGenre.Text & ", " & ComboGenre.Text
    End If

Case 1
    If LenB(Trim$(ComboCountry.Text)) = 0 Then Exit Sub
    If LenB(Trim$(TextCountry.Text)) = 0 Then
        TextCountry.Text = ComboCountry.Text
    Else
        TextCountry.Text = TextCountry.Text & ", " & ComboCountry.Text
    End If

Case 3
    'к примечаниям
    If LenB(Trim$(ComboOther.Text)) = 0 Then Exit Sub
    If LenB(Trim$(TextOther.Text)) = 0 Then
        TextOther.Text = ComboOther.Text
    Else
        TextOther.Text = TextOther.Text & ", " & ComboOther.Text
    End If

Case 2

    'суммировать время
    Dim arr() As String
    Dim i As Integer
    Dim tmpL As Long, tmpL1 As Long

    'убрать последние ", "
    If Right$(TextTimeHid.Text, 2) = ", " Then TextTimeHid.Text = Left$(TextTimeHid.Text, Len(TextTimeHid.Text) - 2)

    If Len(TextTimeHid.Text) <> 0 Then
        arr = Split(TextTimeHid.Text, ",")
        If UBound(arr) > 0 Then
            TextTimeHidText = TextTimeHid.Text
            For i = 0 To UBound(arr)
                tmpL1 = Time2sec(arr(i))
                If tmpL1 > -1 Then tmpL = tmpL + tmpL1
            Next i
            If tmpL > 0 Then TextTimeHid.Text = FormatTime(tmpL)
        Else
            'вернуть
            If Len(TextTimeHidText) <> 0 Then
                'TextTimeHidText
                TextTimeHid.Text = TextTimeHidText
                TextTimeHidText = vbNullString
            End If
        End If
    End If

End Select

End Sub

Private Sub ComRHid_Click(Index As Integer)

Dim tmp As String
Dim temp As Long
Dim strPath As String
Dim site As String

On Error GoTo err

strPath = Space$(255)

Select Case Index
Case 0 'картинку для актера
    tmp = Replace(TextActName.Text, "(", vbNullString)
    tmp = Replace(tmp, ")", vbNullString)
    tmp = Replace(tmp, "/", vbNullString)

    'site = "http://images.yandex.ru/yandsearch?stype=image&text=" & tmp
    site = "http://images.google.com/images?q=" & tmp
    
Case 1 'поиск названия фильма
    tmp = Replace(TxtIName.Text, "(", vbNullString)
    tmp = Replace(tmp, ")", vbNullString)
    tmp = Replace(tmp, "/", vbNullString)

    site = "http://www.google.com/search?q=" & tmp
End Select


temp = FindExecutable(site, "", strPath)
Select Case temp
Case 31
myMsgBox msgsvc(17), vbInformation, , Me.hWnd
Exit Sub
Case 2
End Select

temp = ShellExecute(GetDesktopWindow(), "open", site, vbNull, vbNull, 1)
ToDebug "crh_ret=" & temp

Exit Sub
err:
ToDebug "Err_RHid: " & err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next    'надо для аудио
Timer2.Enabled = False

'Debug.Print KeyCode, Shift

Select Case KeyCode
    
  
Case 45 'Insert auto add
    If Shift = 0 Then FrmAuto.Show 1, Me

Case 27 'esc
    FrmMain.VerticalMenu_MenuItemClick 1, 0
    
Case 123 'f12
    FormDebug.Show , FrmMain    ': FrmMain.SetFocus ' F12
    
Case 122 'f11
    FrmPeople.Show , FrmMain   ': FrmMain.SetFocus ' F11
    
Case 121 'F10 options
  If Not frmOptFlag Then FrmMain.VerticalMenu_MenuItemClick 6, 0

Case 112 'F1
    If ChBTT.Value = 0 Then ChBTT.Value = 1 Else ChBTT.Value = 0

Case 80   'P play DXShow
    If MpegMediaOpen Then
        If TabStrAdEd.SelectedItem.Index = 1 Then
            If MediaState = 2 Then    'играет, на паузу
                mobjManager.Pause
                tPlay.Enabled = False
            Else
                mobjManager.Run
                tPlay.Enabled = True
            End If
        End If
    End If

Case 83   'S sound DXShow
    If MpegMediaOpen Then
        If TabStrAdEd.SelectedItem.Index = 1 Then
            Set objAudio = mobjManager
            If objAudio.Volume = 0 Then objAudio.Volume = -10000 Else objAudio.Volume = 0
            Set objAudio = Nothing
        End If
    End If


End Select
End Sub




Private Sub Form_Load()
'нет тултипов слайдеров
Const TBM_SETTOOLTIPS = &H41D
SendMessage Position.hWnd, TBM_SETTOOLTIPS, 0, 0
SendMessage PositionP.hWnd, TBM_SETTOOLTIPS, 0, 0

'Список скриптов
FillTemplateCombo App.Path & "\Scripts\", ComboInfoSites

Call ChangeComboHeights
End Sub


Private Sub ImgPrCov_Click()
Call picCanvas_Click
End Sub

Private Sub optAspect_Click(Index As Integer)
On Error Resume Next
If ChangeFromCode_optAspect Then Exit Sub

DoEvents

Select Case Index
Case 0    '4:3
    MMI_Format_str = "4/3"
    MMI_Format = 1.333
    ToDebug "UserAspect - 4:3"
Case 1    '16:9
    MMI_Format_str = "16/9"
    MMI_Format = 1.777
    ToDebug "UserAspect - 16:9"
Case 2    'w:h
    'MMI_Format_str = "1/1"
    MMI_Format = 1.333 'если ошибка...
    MMI_Format = objVideo.SourceWidth / objVideo.SourceHeight
    ToDebug "UserAspect - w:h"
    If err Then ToDebug "Error: Video object"

End Select

If Not (objVideo Is Nothing) Then
PixelRatio = objVideo.SourceHeight * MMI_Format / objVideo.SourceWidth
PixelRatioSS = ScrShotEd_W / MMI_Format
End If

movie.SetFocus

End Sub

Private Sub picCanvas_Click()

If NoPicFrontFaceFlag Then Exit Sub
If NoDBFlag Then Exit Sub

FrmMain.PicTempHid(1).Picture = PicFrontFace.Image
IsCoverShowFlag = True
ViewScrShotFlag = False
FormShowPic.Visible = False

'не видим скролл
'FormShowPic.hb_cScroll.Visible(efsHorizontal) = False
FormShowPic.PicHB.Visible = False

PicManualFlag = True

ShowInShowPic 1

FormShowPic.Visible = True

End Sub

Private Sub TextAudioHid_Change()
Mark2Save
End Sub

Private Sub TextCountry_Change()
Mark2Save
If Len(TextCountry.Text) > 255 Then TextCountry.Text = Left$(TextCountry.Text, 255)
'SendMessage TextCountry.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0
End Sub

Private Sub ComboInfoSites_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub ComCancel_Click()
OpenAddmovFlag = False
addflag = False
If rs.EditMode Then rs.CancelUpdate
NoPicFrontFaceFlag = False: NoPic1Flag = False: NoPic2Flag = False: NoPic3Flag = False
FrmMain.VerticalMenu_MenuItemClick 1, 0
End Sub

Private Sub ComCap_AVI(Index As Integer)
On Error GoTo ex
Screen.MousePointer = vbHourglass

Mark2Save
'If rs.EditMode = 0 Then rs.Edit

Select Case Index
''''''''''''''''''''''''''''''''''''0
Case 0
 Set PicSS1Big = Nothing
 NoPic1Flag = False
 PicSS1.Cls
 pos1 = Position.Value
 SavePic1Flag = True 'картинка изменилась, сохранять

If Opt_PicRealRes Then 'большую
   
   PicSS1Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
   PicSS1Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
   m_cAVI.DrawFrame PicSS1Big.hdc, Position.Value, 0, 0, Transparent:=False
   PicSS1Big.Picture = PicSS1Big.Image
End If
   'мал картинка для показа
   PicSS1.Height = PicSS1.Width * movie.Height / movie.Width
   PicSS1.PaintPicture movie.Image, 0, 0, PicSS1.ScaleWidth, PicSS1.ScaleHeight
   PicSS1.Picture = PicSS1.Image
   
   
''''''''''''''''''''''''''''''''''''1
Case 1
 Set PicSS2Big = Nothing
 NoPic2Flag = False
 PicSS2.Cls
 pos2 = Position.Value
 SavePic2Flag = True 'картинка изменилась, сохранять

If Opt_PicRealRes Then 'большую
    
  PicSS2Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
  PicSS2Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
  m_cAVI.DrawFrame PicSS2Big.hdc, Position.Value, 0, 0, Transparent:=False
  PicSS2Big.Picture = PicSS2Big.Image
End If
   'мал картинка для показа
   PicSS2.Height = PicSS2.Width * movie.Height / movie.Width
   PicSS2.PaintPicture movie.Image, 0, 0, PicSS2.ScaleWidth, PicSS2.ScaleHeight
   PicSS2.Picture = PicSS2.Image
    
''''''''''''''''''''''''''''''''''''2
Case 2
 Set PicSS3Big = Nothing
 NoPic3Flag = False
 PicSS3.Cls
 pos3 = Position.Value
 SavePic3Flag = True 'картинка изменилась, сохранять

 If Opt_PicRealRes Then 'большую
    
   PicSS3Big.Width = ScaleX(AviWidth, vbPixels, vbTwips)
   PicSS3Big.Height = ScaleY(AviHeight, vbPixels, vbTwips)
   m_cAVI.DrawFrame PicSS3Big.hdc, Position.Value, 0, 0, Transparent:=False
   PicSS3Big.Picture = PicSS3Big.Image
 End If
   'мал картинка для показа
   PicSS3.Height = PicSS3.Width * movie.Height / movie.Width
   PicSS3.PaintPicture movie.Image, 0, 0, PicSS3.ScaleWidth, PicSS3.ScaleHeight
   PicSS3.Picture = PicSS3.Image
   
End Select

ToDebug "ScrShotsPos Avi: " & pos1 & " " & pos2 & " " & pos3

ex:
If err.Number <> 0 Then Debug.Print "ComCap_AVI: " & err.Description

If Not frmAutoFlag Then Me.SetFocus
Screen.MousePointer = vbNormal

End Sub

Private Sub ComCap_DSh(Index As Integer)
' без учета аспекта
On Error GoTo ex
Screen.MousePointer = vbHourglass

Mark2Save
'If rs.EditMode = 0 Then rs.Edit
ToDebug "Cap_DSh: " & Round(objVideo.SourceWidth, 0) & "x" & objVideo.SourceHeight

Select Case Index
''''''''''''''''''''''''''''''''''''0
Case 0
 Set PicSS1Big = Nothing
 NoPic1Flag = False
 PicSS1.Cls
 pos1 = Position.Value
 SavePic1Flag = True 'картинка изменилась, сохранять

 If Opt_PicRealRes Then 'большую
    
   MPGCaptureBasicVideo PicSS1Big
   If MPGCaptured = False Then SavePic1Flag = False: GoTo ex
   PicSS1Big.Width = ScaleX(objVideo.SourceWidth, vbPixels, vbTwips)
   PicSS1Big.Height = ScaleY(objVideo.SourceHeight, vbPixels, vbTwips)
   PicSS1Big.PaintPicture PicSS1Big.Picture, 0, 0, PicSS1Big.ScaleWidth, PicSS1Big.ScaleHeight
   PicSS1Big.Picture = PicSS1Big.Image
   PicSS1.Cls
   PicSS1.PaintPicture PicSS1Big.Picture, 0, 0, PicSS1.ScaleWidth, PicSS1.ScaleHeight
   PicSS1.Picture = PicSS1.Image
   
 Else ' маленькую
    MPGCaptureBasicVideo PicSS1Big
    If MPGCaptured = False Then SavePic1Flag = False: GoTo ex
    PicSS1Big.Width = ScaleX(objVideo.SourceWidth, vbPixels, vbTwips)
    PicSS1Big.Height = ScaleY(objVideo.SourceHeight, vbPixels, vbTwips)
    PicSS1Big.PaintPicture PicSS1Big.Picture, 0, 0, PicSS1Big.ScaleWidth, PicSS1Big.ScaleHeight
    PicSS1Big.Picture = PicSS1Big.Image
    PicSS1.Cls
    PicSS1.PaintPicture PicSS1Big.Picture, 0, 0, PicSS1.ScaleWidth, PicSS1.ScaleHeight
    PicSS1.Picture = PicSS1.Image
   
End If

''''''''''''''''''''''''''''''''''''1
Case 1
 Set PicSS2Big = Nothing
 NoPic2Flag = False
 PicSS2.Cls
 pos2 = Position.Value
 SavePic2Flag = True 'картинка изменилась, сохранять

 If Opt_PicRealRes Then 'большую
    
  MPGCaptureBasicVideo PicSS2Big
  If MPGCaptured = False Then SavePic2Flag = False: GoTo ex
  PicSS2Big.Width = ScaleX(objVideo.SourceWidth, vbPixels, vbTwips)
  PicSS2Big.Height = ScaleY(objVideo.SourceHeight, vbPixels, vbTwips)
  PicSS2Big.PaintPicture PicSS2Big.Picture, 0, 0, PicSS2Big.ScaleWidth, PicSS2Big.ScaleHeight
  PicSS2Big.Picture = PicSS2Big.Image
  PicSS2.Cls
  PicSS2.PaintPicture PicSS2Big.Picture, 0, 0, PicSS2.ScaleWidth, PicSS2.ScaleHeight
  PicSS2.Picture = PicSS2.Image

 Else ' маленькую

    MPGCaptureBasicVideo PicSS2Big
    If MPGCaptured = False Then SavePic2Flag = False: GoTo ex
    PicSS2Big.Width = ScaleX(objVideo.SourceWidth, vbPixels, vbTwips)
    PicSS2Big.Height = ScaleY(objVideo.SourceHeight, vbPixels, vbTwips)
    PicSS2Big.PaintPicture PicSS2Big.Picture, 0, 0, PicSS2Big.ScaleWidth, PicSS2Big.ScaleHeight
    PicSS2Big.Picture = PicSS2Big.Image
    PicSS2.Cls
    PicSS2.PaintPicture PicSS2Big.Picture, 0, 0, PicSS2.ScaleWidth, PicSS2.ScaleHeight
    PicSS2.Picture = PicSS2.Image

End If

''''''''''''''''''''''''''''''''''''2
Case 2
 Set PicSS3Big = Nothing
 NoPic3Flag = False
 PicSS3.Cls
 pos3 = Position.Value
 SavePic3Flag = True 'картинка изменилась, сохранять

 If Opt_PicRealRes Then 'большую
    
   MPGCaptureBasicVideo PicSS3Big
   If MPGCaptured = False Then SavePic3Flag = False: GoTo ex
   PicSS3Big.Width = ScaleX(objVideo.SourceWidth, vbPixels, vbTwips)
   PicSS3Big.Height = ScaleY(objVideo.SourceHeight, vbPixels, vbTwips)
   PicSS3Big.PaintPicture PicSS3Big.Picture, 0, 0, PicSS3Big.ScaleWidth, PicSS3Big.ScaleHeight
   PicSS3Big.Picture = PicSS3Big.Image
   PicSS3.Cls
   PicSS3.PaintPicture PicSS3Big.Picture, 0, 0, PicSS3.ScaleWidth, PicSS3.ScaleHeight
   PicSS3.Picture = PicSS3.Image

 Else

   MPGCaptureBasicVideo PicSS3Big
   If MPGCaptured = False Then SavePic3Flag = False: GoTo ex
   PicSS3Big.Width = ScaleX(objVideo.SourceWidth, vbPixels, vbTwips)
   PicSS3Big.Height = ScaleY(objVideo.SourceHeight, vbPixels, vbTwips)
   PicSS3Big.PaintPicture PicSS3Big.Picture, 0, 0, PicSS3Big.ScaleWidth, PicSS3Big.ScaleHeight
   PicSS3Big.Picture = PicSS3Big.Image
   PicSS3.Cls
   PicSS3.PaintPicture PicSS3Big.Picture, 0, 0, PicSS3.ScaleWidth, PicSS3.ScaleHeight
   PicSS3.Picture = PicSS3.Image
 End If

End Select

ex:
If err.Number <> 0 Then Debug.Print err.Description

If Not frmAutoFlag Then Me.SetFocus
Screen.MousePointer = vbNormal

End Sub

Private Sub ComCap_DSh_A(Index As Integer)
'c аспектом
On Error GoTo ex

Screen.MousePointer = vbHourglass

Mark2Save

ToDebug "PixelRatio_DS_A = " & PixelRatio
ToDebug "Cap_DSh_A: " & Round(objVideo.SourceWidth * PixelRatio, 0) & "x" & objVideo.SourceHeight

Select Case Index
    ''''''''''''''''''''''''''''''''''''0
Case 0
    Set PicSS1Big = Nothing
    NoPic1Flag = False
    PicSS1.Cls
    pos1 = Position.Value    'objPosition.CurrentPosition * 100
    SavePic1Flag = True    'картинка изменилась, сохранять

    If Opt_PicRealRes Then    'большую

        MPGCaptureBasicVideo PicSS1Big
        If MPGCaptured = False Then SavePic1Flag = False: GoTo ex

        ResizeWIA PicSS1Big, _
                  objVideo.SourceWidth * PixelRatio, _
                  objVideo.SourceHeight

        PicSS1.Cls
        PicSS1.Height = PixelRatioSS
        PicSS1.PaintPicture PicSS1Big.Picture, 0, 0, PicSS1.ScaleWidth, PicSS1.ScaleHeight
        PicSS1.Picture = PicSS1.Image

    Else    ' маленькую
        MPGCaptureBasicVideo PicSS1Big
        If MPGCaptured = False Then SavePic1Flag = False: GoTo ex

        ResizeWIA PicSS1Big, _
                  objVideo.SourceWidth * PixelRatio, _
                  objVideo.SourceHeight

        PicSS1.Cls
        PicSS1.Height = PixelRatioSS
        PicSS1.PaintPicture PicSS1Big.Picture, 0, 0, PicSS1.ScaleWidth, PicSS1.ScaleHeight
        PicSS1.Picture = PicSS1.Image

    End If

    ''''''''''''''''''''''''''''''''''''1
Case 1
    Set PicSS2Big = Nothing
    NoPic2Flag = False
    PicSS2.Cls
    pos2 = Position.Value
    SavePic2Flag = True    'картинка изменилась, сохранять

    If Opt_PicRealRes Then    'большую

        MPGCaptureBasicVideo PicSS2Big
        If MPGCaptured = False Then SavePic2Flag = False: GoTo ex

        ResizeWIA PicSS2Big, _
                  objVideo.SourceWidth * PixelRatio, _
                  objVideo.SourceHeight

        PicSS2.Cls
        PicSS2.Height = PixelRatioSS
        PicSS2.PaintPicture PicSS2Big.Picture, 0, 0, PicSS2.ScaleWidth, PicSS2.ScaleHeight
        PicSS2.Picture = PicSS2.Image

    Else    ' маленькую

        MPGCaptureBasicVideo PicSS2Big
        If MPGCaptured = False Then SavePic2Flag = False: GoTo ex

        ResizeWIA PicSS2Big, _
                  objVideo.SourceWidth * PixelRatio, _
                  objVideo.SourceHeight

        PicSS2.Cls
        PicSS2.Height = PixelRatioSS
        PicSS2.PaintPicture PicSS2Big.Picture, 0, 0, PicSS2.ScaleWidth, PicSS2.ScaleHeight
        PicSS2.Picture = PicSS2.Image

    End If

    ''''''''''''''''''''''''''''''''''''2
Case 2
    Set PicSS3Big = Nothing
    NoPic3Flag = False
    PicSS3.Cls
    pos3 = Position.Value
    SavePic3Flag = True    'картинка изменилась, сохранять

    If Opt_PicRealRes Then    'большую

        MPGCaptureBasicVideo PicSS3Big
        If MPGCaptured = False Then SavePic3Flag = False: GoTo ex

        ResizeWIA PicSS3Big, _
                  objVideo.SourceWidth * PixelRatio, _
                  objVideo.SourceHeight

        PicSS3.Cls
        PicSS3.Height = PixelRatioSS
        PicSS3.PaintPicture PicSS3Big.Picture, 0, 0, PicSS3.ScaleWidth, PicSS3.ScaleHeight
        PicSS3.Picture = PicSS3.Image

    Else

        MPGCaptureBasicVideo PicSS3Big
        If MPGCaptured = False Then SavePic3Flag = False: GoTo ex

        ResizeWIA PicSS3Big, _
                  objVideo.SourceWidth * PixelRatio, _
                  objVideo.SourceHeight

        PicSS3.Cls
        PicSS3.Height = PixelRatioSS
        PicSS3.PaintPicture PicSS3Big.Picture, 0, 0, PicSS3.ScaleWidth, PicSS3.ScaleHeight
        PicSS3.Picture = PicSS3.Image
    End If

End Select

ex:
If err.Number <> 0 Then Debug.Print err.Description

If Not frmAutoFlag Then Me.SetFocus
Screen.MousePointer = vbNormal

End Sub



Private Sub ComCap_Click(Index As Integer)

If isAVIflag Then

    ComCap_AVI Index
    Exit Sub

ElseIf isMPGflag Or isDShflag Then

    If Opt_UseAspect Then    'в масштабе
        ComCap_DSh_A Index
    Else                     'не в масштабе
        ComCap_DSh Index
    End If
    Exit Sub

Else
    'если флагов нет, открыть файл для захвата кадра
    Dim tmp As String
    tmp = pLoadDialog
    ToDebug "Open file for capture"
    If Len(tmp) <> 0 Then OpenMovieForCapture (tmp)
End If

End Sub
Private Sub ComRND_Click(Index As Integer)
Dim tmp As Long
'Dim tmpsi As Single

Mark2Save

Screen.MousePointer = vbHourglass
'                                                                AVI
If isAVIflag Then
    Select Case Index
        Case 0: AutoScrShotsN Frames, 1: SavePic1Flag = True: NoPic1Flag = False
        Case 1: AutoScrShotsN Frames, 2: SavePic2Flag = True: NoPic2Flag = False
        Case 2: AutoScrShotsN Frames, 3: SavePic3Flag = True: NoPic3Flag = False
    End Select
ToDebug "ScrShotsPos Avi: " & pos1 & " " & pos2 & " " & pos3
End If


'                                                               MPEG
If isMPGflag Or isDShflag Then
tmp = TimesX100 / 3 '4

On Error Resume Next 'objPosition.CurrentPosition

    Select Case Index
        Case 0
            pos1 = 1 + (Rnd() * tmp)
            Position.Value = pos1
            objPosition.CurrentPosition = Position.Value / 100
            Sleep 500
            ComCap_Click 0
        
        Case 1
            pos2 = tmp + (Rnd() * tmp)
            Position.Value = pos2
            objPosition.CurrentPosition = Position.Value / 100
            Sleep 500
            ComCap_Click 1

        Case 2: pos3 = tmp * 2 + (Rnd() * tmp)
            Position.Value = pos3
            objPosition.CurrentPosition = Position.Value / 100
            Sleep 500
            ComCap_Click 2
    End Select

MPGPosScroll
End If

Screen.MousePointer = vbNormal


End Sub


Private Sub ComDel_Click()
'If BaseReadOnly Then myMsgBox msgsvc(24), vbInformation, , Me.hwnd: LastVMI = 3: VerticalMenu_MenuItemClick 1, 1: Exit Sub
'If BaseReadOnlyU Then myMsgBox msgsvc(22), vbInformation, , Me.hwnd: LastVMI = 3: VerticalMenu_MenuItemClick 1, 1: Exit Sub
'Dim tmpL As Long
Dim okk As Integer
Dim ComDelEnabled As Boolean
Dim itmX As MSComctlLib.ListItem
Dim i As Long
Dim temp As Long

okk = myMsgBox(msgsvc(15), vbOKCancel, , Me.hWnd)
If okk = 2 Then Exit Sub

ComDel.Enabled = False: ComDelEnabled = True    'не нажимать пока не закончено
ToDebug "DelKey=" & rs("Key")
ClearVideo

If rs.RecordCount > 0 Then

    '?Удалить запись в LV в соответствии ключу базы
    For Each itmX In ListView.ListItems
        If Val(itmX.Key) = rs("Key") Then ListView.ListItems.Remove itmX.Index: Exit For
    Next
    '?Пометить как удаленное - убрать ключ
    'For Each itmX In ListView.ListItems
    '    If Val(itmX.Key) = rs("Key") Then itmX.Key = "": Exit For
    'Next

    'переписать все индекс поля в LV ? а если сортировано...
    ListView.Sorted = False
    For i = 1 To ListView.ListItems.Count
        ListView.ListItems(i).SubItems(lvIndexPole) = i - 1
    Next i
    'если была сортировка - произвести ее
    If LVSortColl > 0 Then LVSOrt (LVSortColl)
    If LVSortColl = -1 Then SortByCheck 0


    rs.Delete    'удалить в базе

    'InitFlag = True    'апдейтить листвью
    'delFlag = True
    If rs.RecordCount < 1 Then
        ComDelEnabled = False: VerticalMenu.SetFocus
        'VerticalMenu_MenuItemClick 3, 0 'добавить новый
        NoListClear
        FrmMain.VerticalMenu_MenuItemClick 1, 0
        'rs.AddNew
        Exit Sub
    End If


Else    'rs.RecordCount <= 0
    Exit Sub
End If

'на след запись или на пред.
rs.MoveNext
If rs.EOF Then
    rs.MovePrevious
    If rs.BOF Then
        ComDelEnabled = False: VerticalMenu.SetFocus
        '        delFlag = True
    End If
End If
'If rs.RecordCount < 1 Then ComDelEnabled = False: VerticalMenu.SetFocus

If rs.EditMode Then
    editFlag = True    ' думаем, что продолжаем редактирование
Else
    'позеленить
    ComSaveRec.BackColor = &HC0E0C0
End If

ToDebug "...удалили"

'загрузить след.
ToDebug "LoadKey=" & rs("Key")
GetEditPix

Mark2SaveFlag = False
GetFields
Mark2SaveFlag = True

'пометить
temp = GotoLVLong(rs("Key"))
If temp > -1 Then
    Set ListView.SelectedItem = ListView.ListItems(temp)
End If

If ComDelEnabled Then ComDel.Enabled = True: ComDel.SetFocus    'вернуть

End Sub


Private Sub ComFrontFaceFile_Click(Index As Integer)
Dim iFile As String
Dim temp As Single
Dim TifPngFlag As Boolean

iFile = pLoadPixDialog
If LCase$(getExtFromFile(iFile)) = "png" Then TifPngFlag = True
If Left$(LCase$(getExtFromFile(iFile)), 3) = "tif" Then TifPngFlag = True

If iFile <> vbNullString Then

Mark2Save
'If rs.EditMode = 0 Then rs.Edit

    On Error GoTo err 'инвалид пикча

    Select Case Index
    Case 0
        Set PicFrontFace = Nothing
        Set picCanvas = Nothing

        If TifPngFlag Then
            PicFrontFace.Picture = LoadPictureWIA(iFile)
        Else
            PicFrontFace.Picture = LoadPicture(iFile)
        End If

        NoPicFrontFaceFlag = False
        SaveCoverFlag = True
        DrawCoverEdit


    Case 1
        If TifPngFlag Then
            PicSS1Big.Picture = LoadPictureWIA(iFile)
        Else
            PicSS1Big.Picture = LoadPicture(iFile)
        End If

        NoPic1Flag = False
        SavePic1Flag = True    'картинка изменилась, сохранять
        PicSS1.Cls
        temp = PicSS1Big.Width / ScrShotEd_W '3360
        PicSS1.Width = ScrShotEd_W '3360
        If PicSS1Big.Height / PicSS1Big.Width < 1 Then PicSS1.Height = PicSS1Big.Height / temp 'Else PicSS1.Height = ScrShotEd_W '3360

        PicSS1.PaintPicture PicSS1Big.Picture, 0, 0, PicSS1.ScaleWidth, PicSS1.ScaleHeight
        PicSS1.Picture = PicSS1.Image

    Case 2
        If TifPngFlag Then
            PicSS2Big.Picture = LoadPictureWIA(iFile)
        Else
            PicSS2Big.Picture = LoadPicture(iFile)
        End If

        NoPic2Flag = False
        SavePic2Flag = True    'картинка изменилась, сохранять
        PicSS2.Cls
        temp = PicSS2Big.Width / ScrShotEd_W
        PicSS2.Width = ScrShotEd_W
        If PicSS2Big.Height / PicSS2Big.Width < 1 Then PicSS2.Height = PicSS2Big.Height / temp 'Else PicSS2.Height = ScrShotEd_W

        PicSS2.PaintPicture PicSS2Big.Picture, 0, 0, PicSS2.ScaleWidth, PicSS2.ScaleHeight
        PicSS2.Picture = PicSS2.Image

    Case 3
        If TifPngFlag Then
            PicSS3Big.Picture = LoadPictureWIA(iFile)
        Else
            PicSS3Big.Picture = LoadPicture(iFile)
        End If

        NoPic3Flag = False
        SavePic3Flag = True    'картинка изменилась, сохранять
        PicSS3.Cls
        temp = PicSS3Big.Width / ScrShotEd_W
        PicSS3.Width = ScrShotEd_W
        If PicSS3Big.Height / PicSS3Big.Width < 1 Then PicSS3.Height = PicSS3Big.Height / temp 'Else PicSS3.Height = ScrShotEd_W

        PicSS3.PaintPicture PicSS3Big.Picture, 0, 0, PicSS3.ScaleWidth, PicSS3.ScaleHeight
        PicSS3.Picture = PicSS3.Image

    End Select

End If    'iFile <> vbNullString

Exit Sub
err:
'Call err.Clear
MsgBox err.Description, vbCritical

End Sub

Private Sub ComInetFind_Click()
Dim SearchString As String 'lcase
Dim i As Integer

Dim iFile As Long
Dim s As String

On Error Resume Next

If Len(ComboInfoSites.Text) = 0 Then Exit Sub    'не выбран скрипт
SearchString = LCase$(TxtIName.Text)
If Len(SearchString) = 0 Then
    'myMsgBox msgsvc(29)
'   Exit Sub
End If

'Set ImgPrCov = Nothing

'прочитать скрипт
iFile = FreeFile
'ComboInfoSites.Text
ToDebug "Скрипт: " & ComboInfoSites

Open App.Path & "\Scripts\" & ComboInfoSites.Text For Binary As #iFile
s = Space$(LOF(iFile))
Get #iFile, , s
Close #iFile

'задать инстанс
Set SC = Nothing
Set SC = CreateObject("ScriptControl")
SC.Language = "VBScript"
SC.Timeout = 35000
SC.AddCode s
'UI окно
'SC.SitehWnd = TextUI.hWnd
'SC.AllowUI = True 'False
'всунуть класс
SC.AddObject "SVC", objScript

ToDebug "Инет поиск: " & SearchString
ComInetFind.Enabled = False
lbInetMovieList.Clear

'Запросить скрипт о url = "http://www.videoguide.ru/find.asp?Search=Simple&types=film&titles="
url = SC.CodeObject.url & SearchString
'url = SC.CodeObject.url & UrlEncode(SearchString)

ComboSites.Text = url
'url = "file://" & App.Path & "\Scripts\inet_dvdempire1.htm"

'ToDebug "запрос: " & url
PageText = OpenURLProxy(url, "txt")
'по строкам
PageArray() = Split(PageText, vbLf)

'Egg
If frmPeopleFlag Then
    FrmPeople.List1.Visible = False
    FrmPeople.List1.Clear
    For i = LBound(PageArray) To UBound(PageArray)
        FrmPeople.List1.AddItem i & " |" & PageArray(i)
    Next i
    FrmPeople.SetListboxScrollbar FrmPeople.List1
    FrmPeople.List1.Visible = True
End If

'скрипт main
SC.Run "AnalyzePage"
If err.Number <> 0 Or SC.Error.Number <> 0 Then GoTo ErrorHandler

'заполнить названиями лист
For i = 0 To UBound(SC.CodeObject.MTitles)
    ' нет мб ошибки в Дата If err.Number <> 0 Or SC.Error.Number <> 0 Then GoTo ErrorHandler
    If Len(SC.CodeObject.MTitles(i)) <> 0 Then
        lbInetMovieList.AddItem SC.CodeObject.MTitles(i)
        lbInetMovieList.ItemData(i) = SC.CodeObject.MData(i)
    End If
Next i
SetListboxScrollbar lbInetMovieList

err.Clear

If (Len(SC.CodeObject.MTitles(0)) <> 0) Or (UBound(SC.CodeObject.MTitles) > 0) Then
    ToDebug "список: " & i & " (" & err.Description & ")"
Else
    ToDebug "список: " & 0
End If

ComInetFind.FontBold = False
ComInetFind.Enabled = True
Erase PageArray
'нужен далее Set SC = Nothing

Exit Sub
ErrorHandler:

If SC.Error.Number <> 0 Then
    MsgBox "Script Error : " & SC.Error.Number _
           & ": " & SC.Error.Description & " строка " & SC.Error.Line _
           & " колонка " & SC.Error.Column, vbCritical

    Set SC = Nothing
Else
    If err.Number <> 0 Then
        'Select Case err.Number
        'Case 9
        '    MsgBox "Nothing found or may be script error"
        'Resume Next
        'Case Else
        MsgBox err.Description, vbCritical
        ToDebug "Err_IFind: " & err.Description
        'End Select
    End If
End If

ComInetFind.FontBold = False
ComInetFind.Enabled = True
End Sub

Private Sub ComKeyNext_Click()
KeyNext
End Sub

Private Sub ComKeyPrev_Click()
KeyPrev
End Sub


Private Sub ComFrontFace_Click(Index As Integer)
'скриншот из буффера
Dim temp As Single

If Clipboard.GetFormat(vbCFDIB) Then

    Mark2Save
    'If rs.EditMode = 0 Then rs.Edit

    Select Case Index

    Case 0
        Set PicFrontFace = Nothing
        Set picCanvas = Nothing

        NoPicFrontFaceFlag = False
        PicFrontFace.Picture = Clipboard.GetData
        SaveCoverFlag = True
        DrawCoverEdit

    Case 1
        PicSS1Big.Picture = Clipboard.GetData
        'Debug.Print "1b", "h:" & PicSS1Big.Height, "sh:" & PicSS1Big.ScaleHeight
        NoPic1Flag = False
        SavePic1Flag = True    'картинка изменилась, сохранять
        PicSS1.Cls
        temp = PicSS1Big.Width / ScrShotEd_W    '3360
        PicSS1.Width = ScrShotEd_W    '3360
        If PicSS1Big.Height / PicSS1Big.Width < 1 Then PicSS1.Height = PicSS1Big.Height / temp Else PicSS1.Height = ScrShotEd_W    '3360
        PicSS1.PaintPicture PicSS1Big.Picture, 0, 0, PicSS1.ScaleWidth, PicSS1.ScaleHeight
        PicSS1.Picture = PicSS1.Image

    Case 2
        PicSS2Big.Picture = Clipboard.GetData
        NoPic2Flag = False
        SavePic2Flag = True    'картинка изменилась, сохранять
        PicSS2.Cls
        temp = PicSS2Big.Width / ScrShotEd_W
        PicSS2.Width = ScrShotEd_W
        If PicSS2Big.Height / PicSS2Big.Width < 1 Then PicSS2.Height = PicSS2Big.Height / temp Else PicSS2.Height = ScrShotEd_W
        PicSS2.PaintPicture PicSS2Big.Picture, 0, 0, PicSS2.ScaleWidth, PicSS2.ScaleHeight
        PicSS2.Picture = PicSS2.Image

    Case 3
        PicSS3Big.Picture = Clipboard.GetData
        NoPic3Flag = False
        SavePic3Flag = True    'картинка изменилась, сохранять
        PicSS3.Cls
        temp = PicSS3Big.Width / ScrShotEd_W
        PicSS3.Width = ScrShotEd_W
        If PicSS3Big.Height / PicSS3Big.Width < 1 Then PicSS3.Height = PicSS3Big.Height / temp Else PicSS3.Height = ScrShotEd_W
        PicSS3.PaintPicture PicSS3Big.Picture, 0, 0, PicSS3.ScaleWidth, PicSS3.ScaleHeight
        PicSS3.Picture = PicSS3.Image
    End Select

End If    'Clipboard.GetFormat(vbCFDIB)
End Sub
Private Sub ComOpen_Click()
'редактирование
'If rs.EditMode = 0 Then rs.Edit
OpenNewMovie
End Sub

Public Sub ComSaveRec_Click()
Dim curKey As String

FirstLVFill = False

ToDebug "SaveClick..."
ToDebug " EditMode: " & rs.EditMode

If rs.EditMode Then
    If Not addflag Then editFlag = True    ' думаем, что продолжаем редактирование
    'editFlag = True ' думаем, что продолжаем редактирование
Else
    If rs.RecordCount < 1 Then Exit Sub
    rs.Edit
    editFlag = True
End If

ToDebug " AddFlag: " & addflag
ToDebug " EditFlag: " & editFlag

curKey = rs("Key") 'ключ добавляемого поля. Важно переместится на него после апдейта

ToDebug " SaveRecKey=" & curKey

If SavePic1Flag Then
    If NoPic1Flag Then
        rs.Fields("SnapShot1") = vbNullString
ToDebug " Pic1 - no"
    Else
        If Opt_PicRealRes Then    'большую
            Pic2JPG PicSS1Big, 1, "SnapShot1"
ToDebug " Pic1 - big"
        Else    'мелкую
            Pic2JPG PicSS1, 1, "SnapShot1"
ToDebug " Pic1 - small"
        End If
    End If
End If
If SavePic2Flag Then
    If NoPic2Flag Then
        rs.Fields("SnapShot2") = vbNullString
ToDebug " Pic2 - no"
    Else
        If Opt_PicRealRes Then    'большую
            Pic2JPG PicSS2Big, 1, "SnapShot2"
ToDebug " Pic2 - big"
        Else    'мелкую
            Pic2JPG PicSS2, 1, "SnapShot2"
ToDebug " Pic2 - small"
        End If
    End If
End If
If SavePic3Flag Then
    If NoPic3Flag Then
        rs.Fields("SnapShot3") = vbNullString
ToDebug " Pic3 - no"
    Else
        If Opt_PicRealRes Then    'большую
            Pic2JPG PicSS3Big, 1, "SnapShot3"
ToDebug " Pic3 - big"
        Else    'мелкую
            Pic2JPG PicSS3, 1, "SnapShot3"
ToDebug " Pic3 - small"
        End If
    End If
End If
If SaveCoverFlag Then
    If NoPicFrontFaceFlag Then
        rs.Fields("FrontFace") = vbNullString
ToDebug " Cover - no"
    Else
        Pic2JPG PicFrontFace, 1, "FrontFace"
ToDebug " Cover - yes"
    End If
End If
SavePic1Flag = False: SavePic2Flag = False: SavePic3Flag = False: SaveCoverFlag = False

'положить в базу поля
PutFields    'там апдейт с возвратом позиции на до добавления

FrameAddEdit.Caption = AddEditCapt & " > " & TextMName.Text


If addflag Then    '

    RSGoto curKey

    ListView.Sorted = False

    ReDim Preserve lvItemLoaded(ListView.ListItems.Count + 1)    ' 1
    Add2LV ListView.ListItems.Count, ListView.ListItems.Count + 1    '2

    CurLVKey = rs("Key") & """"
    CurSearch = GotoLV(CurLVKey)
    'пометить
    'If Not frmAutoFlag Then
    If ListView.ListItems.Count > 0 Then
        Set ListView.SelectedItem = ListView.ListItems(CurSearch)
    End If
    'End If

    addflag = False
End If

If editFlag Then
    ' Dim lvSelInd As Long

    ListView.Sorted = False
    'редактировать lv

    CurLVKey = rs("Key") & """"
    CurSearch = GotoLV(CurLVKey)

    EditLV CurSearch

    'сказать, что еще не показаны сабы
    If Opt_LoadOnlyTitles = True Then    ' только названия
        lvItemLoaded(CurSearch) = False
    End If

    'пометить
    'нужно тк мб тут же редактирование и запись в CurSearch
    Set ListView.SelectedItem = ListView.ListItems(CurSearch)
    'GotoLV CurLVKey

End If    'editflag


If delFlag Then
    'апдейтить если были удаления
    InitFlag = True
Else
    InitFlag = False
End If


'если была сортировка - произвести ее
If LVSortColl > 0 Then LVSOrt (LVSortColl)
If LVSortColl = -1 Then SortByCheck 0, True

'last DoEvents

'Set ListView.SelectedItem = ListView.ListItems(CurSearch)
If FrameView.Visible Then ListView.SelectedItem.EnsureVisible    ': LVCLICK

'On Error GoTo 0

AutoFillStore    'запомнить поля в списки для последующего авто ввода

ComDel.Visible = True: ComDel.Enabled = True: ComSaveRec.BackColor = &HC0E0C0
ToDebug "...сохранили"

End Sub

Private Sub ComShowBin_Click()

FrmBin.Show
'FrmMainState = FrmMain.WindowState
'FrmMain.WindowState = vbMinimized

End Sub


Private Sub ComX_Click(Index As Integer)

Select Case Index
 Case 0
  If PicSS1.Picture = 0 Then Exit Sub
 Case 1
  If PicSS2.Picture = 0 Then Exit Sub
 Case 2
  If PicSS3.Picture = 0 Then Exit Sub
 Case 3, 4
  If PicFrontFace.Picture = 0 Then Exit Sub
End Select

If myMsgBox(msgsvc(4), vbOKCancel, , Me.hWnd) = vbCancel Then Exit Sub 'удалять?

Mark2Save

Select Case Index

Case 0
Set PicSS1 = Nothing: Set PicSS1Big = Nothing: NoPic1Flag = True
SavePic1Flag = True

Case 1
Set PicSS2 = Nothing: Set PicSS2Big = Nothing: NoPic2Flag = True
SavePic2Flag = True

Case 2
Set PicSS3 = Nothing: Set PicSS3Big = Nothing: NoPic3Flag = True
SavePic3Flag = True

Case 3, 4
Set PicFrontFace = Nothing: NoPicFrontFaceFlag = True
Set picCanvas = Nothing
Set ImgPrCov = Nothing
SaveCoverFlag = True

End Select


End Sub



Private Sub ComX_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ComX(Index).BackColor = &HC0C0FF
End Sub

Private Sub ComX_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ComX(Index).BackColor = &HFFFFFF
End Sub


Private Sub lbInetMovieList_Click()
Dim temp As String
Dim i As Integer

On Error Resume Next

If SC Is Nothing Then Exit Sub
AutoFillStore 'запомнить поля в списки для последующего авто ввода

With SC.CodeObject
    'ClearTextFields 'не надо
    'If UBound(.MTitlesURL) = LBound(.MTitlesURL) Then Exit Sub  а один?
    temp = .MTitlesURL(lbInetMovieList.ListIndex) '+ 1)
End With

If Len(temp) = 0 Then Exit Sub
'Debug.Print temp, URLTitleArr(lbInetMovieList.ListIndex + 1)
ComboSites.Text = temp 'url текущей в список урл-ов

'DoEvents
'ToDebug "поиск инфо на " & temp

PageText = OpenURLProxy(temp, "txt")

'Debug.Print PageText

'Open "g:\t" For Output As #2
'Print #2, PageText ' Input(LOF(1), 1)
'Close #2
PageArray() = Split(PageText, vbLf)

    'Egg
    If frmPeopleFlag Then
    FrmPeople.List1.Clear
    FrmPeople.List1.Visible = False
    For i = LBound(PageArray) To UBound(PageArray)
    FrmPeople.List1.AddItem i & " |" & PageArray(i)
    Next i
    FrmPeople.SetListboxScrollbar FrmPeople.List1
    FrmPeople.List1.Visible = True
    End If


'почистить
With SC.CodeObject
.MTitle = ""
.MYear = ""
.MGenre = ""
.MDirector = ""
.MActors = ""
.MDescription = ""
.MCountry = ""
.MPicURL = ""
.MRating = ""
.MLang = ""
.MSubt = ""
.MOther = ""
End With

'скрипт , передать индекс фильма
SC.Run "AnalyzeMoviePage", CVar(lbInetMovieList.ListIndex)
If SC.Error.Number <> 0 Or err.Number <> 0 Then GoTo ErrorHandler

'что нашли положить в поля редактора
SetFromScript 'todebug там

AutoFillStore 'запомнить поля после
Erase PageArray

Exit Sub
ErrorHandler:

With SC.Error
If .Number <> 0 Then
    MsgBox "Script Error : " & .Number _
        & ": " & .Description & " строка " & .Line _
        & " колонка " & .Column, vbCritical
Set SC = Nothing
Else
MsgBox err.Description, vbCritical
End If
End With
End Sub

Private Sub movie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'не задействуется, если не ави
If m_cAVI Is Nothing Then Exit Sub 'на всякий

'pRenderFrame 1
'PosScroll

'pRenderFrame Position.Value 'astRendedAVI 'типа напомнить, что в муви окне
'Position.Refresh

If OpenAddmovFlag Then
    If Button = 2 Then
    MsgBox ""
'    Me.PopupMenu Me.popMovieHid
    Else

    Set FrmMain.PicTempHid(0) = Nothing: Set FrmMain.PicTempHid(1) = Nothing
    FrmMain.PicTempHid(0).Width = ScaleX(AviWidth, vbPixels, vbTwips)
    FrmMain.PicTempHid(0).Height = ScaleY(AviHeight, vbPixels, vbTwips)

        If Position.Value = 0 Then Position.Value = 1
        m_cAVI.DrawFrame FrmMain.PicTempHid(0).hdc, Position.Value  ', 0, 0, Transparent:=False
        Call ShowInShowPic(0)
    End If ' buttons
End If 'ActFlag нет видео нет попапа

If Not frmAutoFlag Then Me.SetFocus
End Sub

Private Sub PicSS1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set FrmMain.PicTempHid(0) = Nothing
If Button = 1 Then
    If PicSS1.Picture <> 0 Then

        IsCoverShowFlag = False
        ViewScrShotFlag = True

        If PicSS1Big.Picture <> 0 Then
            FrmMain.PicTempHid(0).Picture = PicSS1Big.Picture
        Else
            If GetPic(FrmMain.PicTempHid(0), 1, "SnapShot1") Then FrmMain.PicTempHid(0).Picture = FrmMain.PicTempHid(0).Image: PicSS1Big.Picture = FrmMain.PicTempHid(0).Image
        End If
    End If
    If FrmMain.PicTempHid(0).Picture <> 0 Then

        FormShowPic.PicHB.Visible = False    'убрать скролл
        ShowInShowPic 0
        FormShowPic.Visible = True
    End If

End If

If Button = 2 Then

    If OpenAddmovFlag = True Then
        If isAVIflag Then
            Position.Value = pos1: lastRendedAVI = pos1: PosScroll
        End If
        If isMPGflag Or isDShflag Then
            Position.Value = pos1: MPGPosScroll
        End If
            ToDebug "SShotGoto: " & pos1
    End If
End If


End Sub


Private Sub PicSS2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set FrmMain.PicTempHid(0) = Nothing
If Button = 1 Then
    If PicSS2.Picture <> 0 Then
        IsCoverShowFlag = False

        If PicSS2Big.Picture <> 0 Then
            FrmMain.PicTempHid(0).Picture = PicSS2Big.Picture
        Else
            If GetPic(FrmMain.PicTempHid(0), 1, "SnapShot2") Then FrmMain.PicTempHid(0).Picture = FrmMain.PicTempHid(0).Image: PicSS2Big.Picture = FrmMain.PicTempHid(0).Image
        End If
    End If
    If FrmMain.PicTempHid(0).Picture <> 0 Then

        FormShowPic.PicHB.Visible = False    'убрать скролл
        ShowInShowPic 0
        FormShowPic.Visible = True
    End If

End If

If Button = 2 Then
    If OpenAddmovFlag = True Then
        If isAVIflag Then
            Position.Value = pos2: lastRendedAVI = pos2: PosScroll
        End If
        If isMPGflag Or isDShflag Then
            Position.Value = pos2: MPGPosScroll
        End If
            ToDebug "SShotGoto: " & pos2
    End If
End If

End Sub


Private Sub PicSS3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set FrmMain.PicTempHid(0) = Nothing
If Button = 1 Then
    If PicSS3.Picture <> 0 Then
        IsCoverShowFlag = False

        If PicSS3Big.Picture <> 0 Then
            FrmMain.PicTempHid(0).Picture = PicSS3Big.Picture
        Else
            If GetPic(FrmMain.PicTempHid(0), 1, "SnapShot3") Then FrmMain.PicTempHid(0).Picture = FrmMain.PicTempHid(0).Image: PicSS3Big.Picture = FrmMain.PicTempHid(0).Image
        End If
    End If
    If FrmMain.PicTempHid(0).Picture <> 0 Then

        FormShowPic.PicHB.Visible = False    'убрать скролл
        ShowInShowPic 0
        FormShowPic.Visible = True
    End If

End If

If Button = 2 Then
    If OpenAddmovFlag = True Then
        If isAVIflag Then
            Position.Value = pos3: lastRendedAVI = pos3: PosScroll
        End If
        If isMPGflag Or isDShflag Then
            Position.Value = pos3: MPGPosScroll
        End If
            ToDebug "SShotGoto: " & pos3
    End If
End If

End Sub


Private Sub Position_KeyUp(KeyCode As Integer, Shift As Integer)
'Debug.Print KeyCode
Select Case KeyCode
Case 37, 38, 39, 40
If isAVIflag Then PosScroll: lastRendedAVI = Position.Value
If isMPGflag Or isDShflag Then MPGPosScroll: lastRendedMPG = Position.Value
End Select
End Sub

Private Sub Position_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If isAVIflag Then PosScroll: lastRendedAVI = Position.Value
If isMPGflag Or isDShflag Then MPGPosScroll: lastRendedMPG = Position.Value
End Sub

Private Sub Position_Scroll()
'Position.Value = Position.Value
'Position.Text = "-"
If isAVIflag Then
If lastRendedAVI = Position.Value Then Exit Sub
If ChLockFHid Then pRenderFrame Position.Value
lastRendedAVI = Position.Value
End If

If isMPGflag Or isDShflag Then
If lastRendedMPG = Position.Value Then Exit Sub
    If ChLockFHid Then
    
    MPGPosScroll
    '    objPosition.CurrentPosition = Position.Value / 100
    lastRendedMPG = Position.Value
    mobjManager.Pause
    End If
End If

'Debug.Print "scroll"
End Sub

Private Sub PositionP_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case 37, 38, 39, 40 'стрелки

Position.Value = PositionP.Value

If isAVIflag Then
pRenderFrame Position.Value 'PosScroll
lastRendedAVI = Position.Value
End If

If isMPGflag Or isDShflag Then 'MPGPosScroll
        objPosition.CurrentPosition = Position.Value / 100
        lastRendedMPG = Position.Value
'Debug.Print objPosition.CurrentPosition
        'mobjManager.Pause
End If

End Select
End Sub

Private Sub PositionP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next 'VTS_01_0.VOB
Position.Value = PositionP.Value

If isAVIflag Then pRenderFrame Position.Value: lastRendedAVI = Position.Value

If isMPGflag Or isDShflag Then 'MPGPosScroll
Screen.MousePointer = vbHourglass
        objPosition.CurrentPosition = Position.Value / 100
Screen.MousePointer = vbNormal
lastRendedMPG = Position.Value
'        mobjManager.Pause
'Debug.Print objPosition.CurrentPosition
End If

'pRenderFrame Position.Value
End Sub

Private Sub PositionP_Scroll()
On Error Resume Next 'маленькие или нелепые файлы (VTS_01_0.VOB)
Position.Value = PositionP.Value

If isAVIflag Then
If lastRendedAVI = Position.Value Then Exit Sub
If ChLockSHid Then pRenderFrame Position.Value
End If

If isMPGflag Or isDShflag Then 'MPGPosScroll
If lastRendedMPG = Position.Value Then Exit Sub
    If ChLockSHid Then
    Screen.MousePointer = vbHourglass
      objPosition.CurrentPosition = Position.Value / 100
        mobjManager.Pause
    Screen.MousePointer = vbNormal
'Debug.Print temp, objPosition.CurrentPosition
    End If
End If

End Sub

Private Sub TabStrAdEd_Click()
'On Error Resume Next

Select Case TabStrAdEd.SelectedItem.Index

Case 2  'текстовая инфа
    FrAdEdTextHid.Visible = True
    FrAdEdTextHid.ZOrder 0
    FrAdEdPixHid.Visible = False
    FrAdEdTechHid.Visible = False

    ImgPrCov.Picture = PicFrontFace.Picture
    FrImgPrCov.ZOrder 0

Case 1  'картинки
    FrAdEdPixHid.Visible = True
    FrAdEdPixHid.ZOrder 0
    FrAdEdTextHid.Visible = False
    FrAdEdTechHid.Visible = False

    FrImgPrCov.ZOrder 1
    
Case 3  'техинфо
    FrAdEdTechHid.Visible = True
    FrAdEdTechHid.ZOrder 0
    FrAdEdTextHid.Visible = False
    FrAdEdPixHid.Visible = False

    ImgPrCov.Picture = PicFrontFace.Picture
    FrImgPrCov.ZOrder 0

End Select

'фокусы текстовых полей
If addflag Then
    If TextCDN.Visible Then TextCDN.SetFocus
    If TextMName.Visible Then TextMName.SetFocus: TextMName.SelLength = 0
End If

ComOpen.ZOrder 0
ComAdd.ZOrder 0
ComDel.ZOrder 0
ComSaveRec.ZOrder 0
ComCancel.ZOrder 0
'movie.ZOrder 0 тогда и кнопки, что сверху
End Sub

Private Sub TextAnnotation_Change()
Mark2Save
End Sub

Private Sub TextAnnotation_KeyPress(KeyAscii As Integer)
Const ASC_CTRL_A As Integer = 1

    If KeyAscii = ASC_CTRL_A Then
        TextAnnotation.SelStart = 0
        TextAnnotation.SelLength = Len(TextAnnotation.Text)
    End If

End Sub

Private Sub TextAuthor_Change()
Mark2Save
End Sub

Private Sub TextAuthor_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextAuthor) Then
Mark2Save
End If
End Sub

Private Sub TextAuthor_LostFocus()
TextAuthor.Text = sTrimChars(TextAuthor.Text, vbNewLine)
End Sub

Private Sub tPlay_Timer()
On Error Resume Next
Position.Value = objPosition.CurrentPosition * 100
End Sub
Private Sub TextCDN_Change()
Mark2Save
End Sub

Private Sub TextCountry_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextCountry) Then
Mark2Save
End If
End Sub

Private Sub TextCountry_LostFocus()
TextCountry.Text = sTrimChars(TextCountry.Text, vbNewLine)
End Sub


Private Sub TextCoverURL_Change()
Mark2Save
End Sub

Private Sub TextFilelenHid_Change()
Mark2Save
End Sub

Private Sub TextFileName_Change()
Mark2Save
End Sub

Private Sub TextFPSHid_Change()
Mark2Save
End Sub

Private Sub TextGenre_Change()
Mark2Save
End Sub

Private Sub TextGenre_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextGenre) Then
Mark2Save
End If
End Sub

Private Sub TextGenre_KeyPress(KeyAscii As Integer)
'ComboKeyAscii KeyAscii
'If AutoFill(TextGenre) Then
'Mark2Save
'End If
End Sub

Private Sub TextGenre_LostFocus()
TextGenre.Text = sTrimChars(TextGenre.Text, vbNewLine)

End Sub
Private Sub TextLabel_Change()
Mark2Save
End Sub

Private Sub TextLabel_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextLabel) Then
Mark2Save
End If
End Sub

Private Sub TextLabel_LostFocus()
TextLabel.Text = sTrimChars(TextLabel.Text, vbNewLine)
End Sub

Private Sub TextLang_Change()
Mark2Save
End Sub

Private Sub TextLang_Click()
'Mark2Save
End Sub

Private Sub TextLang_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextLang) Then
Mark2Save
End If
End Sub

Private Sub TextLang_LostFocus()
TextLang.Text = sTrimChars(TextLang.Text, vbNewLine)
End Sub

Private Sub TextMName_Change()
Mark2Save
ComInetFind.Enabled = True
End Sub

Private Sub TextMName_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextMName) Then
Mark2Save
End If
End Sub

Private Sub TextMName_LostFocus()
Mark2SaveFlag = False
TextMName.Text = sTrimChars(TextMName.Text, vbNewLine)
Mark2SaveFlag = True
End Sub

Private Sub TextMovURL_Change()
Mark2Save
End Sub

Private Sub TextOther_Change()
Mark2Save
End Sub

Private Sub TextOther_Click()
'Mark2Save
End Sub

Private Sub TextOther_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextOther) Then
Mark2Save
End If
End Sub


Private Sub TextRate_Change()
Mark2Save
End Sub

Private Sub TextRate_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextRate) Then
Mark2Save
End If
End Sub

Private Sub TextRate_LostFocus()
TextRate.Text = sTrimChars(TextRate.Text, vbNewLine)
End Sub

Private Sub TextResolHid_Change()
Mark2Save
End Sub

Private Sub TextRole_Change()
Mark2Save
End Sub

Private Sub TextRole_LostFocus()
TextRole.Text = sTrimChars(TextRole.Text, vbNewLine)
End Sub

Private Sub TextSubt_Change()
Mark2Save
End Sub

Private Sub TextSubt_Click()
'Mark2Save
End Sub

Private Sub TextSubt_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextSubt) Then
Mark2Save
End If
End Sub

Private Sub TextSubt_LostFocus()
TextSubt.Text = sTrimChars(TextSubt.Text, vbNewLine)
End Sub

Private Sub TextTimeHid_Change()
Mark2Save
End Sub

Private Sub TextUser_Change()
Mark2Save
End Sub

Private Sub TextUser_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextUser) Then
Mark2Save
End If
End Sub

Private Sub TextVideoHid_Change()
Mark2Save
End Sub

Private Sub TextYear_Change()
Mark2Save
End Sub

Private Sub TextYear_KeyDown(KeyCode As Integer, Shift As Integer)
ComboKey KeyCode, Shift
If AutoFill(TextYear) Then
Mark2Save
End If
End Sub

Private Sub TextYear_LostFocus()
TextYear.Text = sTrimChars(TextYear.Text, vbNewLine)
End Sub

Private Sub TxtIName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then ComInetFind_Click
End Sub

Public Function MpegSizeAdjust(cl As Boolean) As Boolean
'MMI_Format переменная тут заменена на локальную movieAspect
Dim movieAspect As Currency 'Single

Dim objVideoW As IVideoWindow
Dim lngWidth As Long
Dim lngHeight As Long

On Error Resume Next

MpegSizeAdjust = True
Set objVideoW = mobjManager

If cl Then 'очистка
 If (True <> (objVideoW Is Nothing)) Then objVideoW.Visible = False ': objVideoW.Owner = 0
 Set objVideoW = Nothing
 Exit Function
End If
    
If (True <> (objVideoW Is Nothing)) Then
'Debug.Print "objVideoW.Width.Orig=" & objVideoW.Width & " objVideoW.Height.Orig=" & objVideoW.Height

VideoStand = "PAL"
If InStr(1, TextVideoHid, "PAL", vbTextCompare) > 0 Then VideoStand = "PAL"
If InStr(1, TextVideoHid, "NTSC", vbTextCompare) > 0 Then VideoStand = "NTSC"
If InStr(1, TextVideoHid, "FILM", vbTextCompare) > 0 Then VideoStand = "FILM"

movieAspect = MMI_Format 'по умолчанию

Select Case VideoStand
Case "PAL"
    Select Case strMPEG
     Case "MPV2"
      If SVCDflag Then
       'svcd
       movieAspect = 1.333
       'movieAspect 'objVideo.SourceHeight * movieAspect / objVideo.SourceWidth
       'PixelRatioSS = ScrShotEd_W / movieAspect
      Else
       'dvd
      End If
     Case "MPV1"
     'vcd
      movieAspect = 1.333
    End Select
Case "NTSC"
    Select Case strMPEG
     Case "MPV2"
      If SVCDflag Then
       'svcd
       movieAspect = 1.333
      Else
       'dvd
      End If
     Case "MPV1"
     'vcd
     movieAspect = 1.333
    End Select
Case "FILM"
    movieAspect = MMI_Format
End Select

If movieAspect = 0 Then movieAspect = 1.333 'mp4

 With objVideoW
 
 On Error GoTo err
 
  .Owner = movie.hWnd
'  .WindowStyle = enWindowStyles.WS_VISIBLE '&H80000000    'WS_CHILD
  .WindowStyle = WS_CHILD 'так надо
  lngWidth = .Width: lngHeight = .Height
  '.AutoShow = True
  '.FullScreenMode = True
  MovieWidth = ScaleX(movie.Width, vbTwips, vbPixels)
  MovieHeight = MovieWidth / movieAspect
  movie.Height = ScaleY(MovieHeight, vbPixels, vbTwips)
  Call .SetWindowPosition(0&, 0&, MovieWidth, MovieHeight)
  'Call .SetWindowPosition(0&, 0&, lngWidth, lngHeight)
  
  .Visible = True
 End With
End If

'Debug.Print "objVideoW.Width.movie=" & objVideoW.Width & " objVideoW.Height.movie=" & objVideoW.Height

ToDebug "видео-окно: " & movieAspect

err:
Set objVideoW = Nothing
If err.Number <> 0 Then
ToDebug "Error in MSAcl: " & err.Description
MpegSizeAdjust = False
End If
End Function




