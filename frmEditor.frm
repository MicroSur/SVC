VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor"
   ClientHeight    =   8565
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   571
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   717
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tPlay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   8100
   End
   Begin VB.Frame FrameAddEdit 
      Caption         =   "Edit"
      Height          =   8535
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10665
      Begin VB.Frame FrImgPrCov 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   8940
         TabIndex        =   113
         Top             =   3300
         Width           =   1275
         Begin VB.CommandButton ComX 
            Caption         =   "X"
            Height          =   255
            Index           =   4
            Left            =   1020
            MousePointer    =   1  'Arrow
            TabIndex        =   114
            Top             =   60
            Width           =   255
         End
         Begin VB.Image ImgPrCov 
            Height          =   1335
            Left            =   0
            MouseIcon       =   "frmEditor.frx":0000
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   60
            Width           =   1215
         End
      End
      Begin SurVideoCatalog.XpB ComDel 
         Height          =   375
         Left            =   8580
         TabIndex        =   115
         Top             =   1740
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   688
         Caption         =   "Del"
         ButtonStyle     =   3
         Picture         =   "frmEditor.frx":08CA
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComOpen 
         Height          =   375
         Left            =   8580
         TabIndex        =   116
         Top             =   720
         Width           =   1935
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "Open"
         ButtonStyle     =   3
         Picture         =   "frmEditor.frx":12DC
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
         MaskColor       =   16711935
      End
      Begin SurVideoCatalog.XpB ComSaveRec 
         Height          =   375
         Left            =   8580
         TabIndex        =   117
         Top             =   2280
         Width           =   1935
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "Save"
         ButtonStyle     =   3
         Picture         =   "frmEditor.frx":1630
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComAdd 
         Height          =   375
         Left            =   8580
         TabIndex        =   118
         Top             =   1200
         Width           =   1935
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "Add"
         ButtonStyle     =   3
         Picture         =   "frmEditor.frx":1BCA
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin SurVideoCatalog.XpB ComCancel 
         Height          =   375
         Left            =   8580
         TabIndex        =   119
         Top             =   2820
         Width           =   1935
         _ExtentX        =   265
         _ExtentY        =   265
         Caption         =   "Cancel"
         ButtonStyle     =   3
         Picture         =   "frmEditor.frx":25DC
         PictureWidth    =   16
         PictureHeight   =   16
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
      End
      Begin VB.Frame FrAdEdPixHid 
         BorderStyle     =   0  'None
         Height          =   7695
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   10455
         Begin VB.CommandButton ComX 
            Caption         =   "X"
            Height          =   315
            Index           =   5
            Left            =   3660
            MousePointer    =   1  'Arrow
            TabIndex        =   121
            Top             =   60
            Width           =   315
         End
         Begin MSComctlLib.Slider PositionP 
            Height          =   375
            Left            =   240
            TabIndex        =   64
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
            TabIndex        =   65
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
            TabIndex        =   66
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
            TabIndex        =   67
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
            TabIndex        =   68
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
            TabIndex        =   69
            Top             =   3120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   767
            Caption         =   "Random"
            ButtonStyle     =   3
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
         End
         Begin VB.OptionButton optAspect 
            Appearance      =   0  'Flat
            Caption         =   "16:9"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   9660
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   2760
            Width           =   495
         End
         Begin VB.OptionButton optAspect 
            Appearance      =   0  'Flat
            Caption         =   "4:3"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   9180
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   2760
            Width           =   495
         End
         Begin VB.CommandButton ComCap 
            Height          =   495
            Index           =   2
            Left            =   9780
            MousePointer    =   1  'Arrow
            Picture         =   "frmEditor.frx":2FEE
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   4500
            Width           =   615
         End
         Begin VB.CommandButton ComFrontFace 
            Height          =   375
            Index           =   3
            Left            =   9360
            MousePointer    =   1  'Arrow
            Picture         =   "frmEditor.frx":32F8
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   4500
            Width           =   435
         End
         Begin VB.CommandButton ComFrontFaceFile 
            Height          =   375
            Index           =   3
            Left            =   8940
            MousePointer    =   1  'Arrow
            Picture         =   "frmEditor.frx":3882
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   4500
            Width           =   435
         End
         Begin VB.PictureBox PicSS3 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2775
            Left            =   7020
            MouseIcon       =   "frmEditor.frx":3E0C
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
         Begin VB.CommandButton ComCap 
            Height          =   495
            Index           =   1
            Left            =   6300
            MousePointer    =   1  'Arrow
            Picture         =   "frmEditor.frx":46D6
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   4500
            Width           =   615
         End
         Begin VB.CommandButton ComFrontFace 
            Height          =   375
            Index           =   2
            Left            =   5880
            MousePointer    =   1  'Arrow
            Picture         =   "frmEditor.frx":49E0
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   4500
            Width           =   435
         End
         Begin VB.CommandButton ComFrontFaceFile 
            Height          =   375
            Index           =   2
            Left            =   5460
            MousePointer    =   1  'Arrow
            Picture         =   "frmEditor.frx":4F6A
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   4500
            Width           =   435
         End
         Begin VB.PictureBox PicSS2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2775
            Left            =   3540
            MouseIcon       =   "frmEditor.frx":54F4
            MousePointer    =   99  'Custom
            ScaleHeight     =   185
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   224
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   4620
            Width           =   3360
            Begin VB.CommandButton ComX 
               Caption         =   "X"
               Height          =   315
               Index           =   1
               Left            =   0
               MousePointer    =   1  'Arrow
               TabIndex        =   53
               Top             =   0
               Width           =   315
            End
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
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   60
            Width           =   4680
         End
         Begin VB.CommandButton ComCap 
            Height          =   495
            Index           =   0
            Left            =   2820
            MousePointer    =   1  'Arrow
            Picture         =   "frmEditor.frx":5DBE
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   4500
            Width           =   615
         End
         Begin VB.CommandButton ComFrontFace 
            Height          =   375
            Index           =   1
            Left            =   2400
            MousePointer    =   1  'Arrow
            Picture         =   "frmEditor.frx":60C8
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   4500
            Width           =   435
         End
         Begin VB.CommandButton ComFrontFaceFile 
            Height          =   375
            Index           =   1
            Left            =   1980
            MousePointer    =   1  'Arrow
            Picture         =   "frmEditor.frx":6652
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   4500
            Width           =   435
         End
         Begin VB.CommandButton ComKeyAvi 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   10020
            Picture         =   "frmEditor.frx":6BDC
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   4020
            Width           =   375
         End
         Begin VB.CommandButton ComKeyAvi 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   10020
            Picture         =   "frmEditor.frx":7346
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   3660
            Width           =   375
         End
         Begin VB.PictureBox PicFrontFace 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   660
            Left            =   120
            MousePointer    =   99  'Custom
            ScaleHeight     =   44
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   44
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.PictureBox picCanvas 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3525
            Left            =   0
            MouseIcon       =   "frmEditor.frx":7AB0
            MousePointer    =   99  'Custom
            ScaleHeight     =   235
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   239
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   60
            Width           =   3585
            Begin VB.CommandButton ComX 
               Caption         =   "X"
               Height          =   315
               Index           =   3
               Left            =   0
               MousePointer    =   1  'Arrow
               TabIndex        =   44
               Top             =   0
               Width           =   315
            End
            Begin VB.CommandButton ComFrontFace 
               Height          =   375
               Index           =   0
               Left            =   3180
               MousePointer    =   1  'Arrow
               Picture         =   "frmEditor.frx":837A
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   0
               Width           =   435
            End
            Begin VB.CommandButton ComFrontFaceFile 
               Height          =   375
               Index           =   0
               Left            =   2760
               MousePointer    =   1  'Arrow
               Picture         =   "frmEditor.frx":8904
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   0
               Width           =   435
            End
         End
         Begin VB.CheckBox ChLockFSHid 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   40
            Top             =   4080
            Value           =   1  'Checked
            Width           =   195
         End
         Begin VB.CheckBox ChLockFSHid 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   39
            Top             =   3720
            Width           =   195
         End
         Begin VB.PictureBox PicSS1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2775
            Left            =   60
            MouseIcon       =   "frmEditor.frx":8E8E
            MousePointer    =   99  'Custom
            ScaleHeight     =   185
            ScaleMode       =   0  'User
            ScaleWidth      =   224
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   4620
            Width           =   3360
            Begin VB.CommandButton ComX 
               Caption         =   "X"
               Height          =   315
               Index           =   0
               Left            =   0
               MousePointer    =   1  'Arrow
               TabIndex        =   38
               Top             =   0
               Width           =   315
            End
            Begin VB.PictureBox Picture1 
               Height          =   0
               Left            =   0
               ScaleHeight     =   0
               ScaleWidth      =   0
               TabIndex        =   37
               Top             =   0
               Width           =   0
            End
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
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   4440
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
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   4440
            Visible         =   0   'False
            Width           =   555
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
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   4380
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.OptionButton optAspect 
            Appearance      =   0  'Flat
            Caption         =   "w:h"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   8700
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2760
            Width           =   495
         End
      End
      Begin VB.Frame FrAdEdTechHid 
         BorderStyle     =   0  'None
         Height          =   7695
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   10455
         Begin VB.ComboBox cBasePicURL 
            Height          =   315
            ItemData        =   "frmEditor.frx":9758
            Left            =   1260
            List            =   "frmEditor.frx":975A
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   4500
            Width           =   3015
         End
         Begin VB.TextBox TextMovURL 
            Height          =   315
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   14
            Top             =   4980
            Width           =   7095
         End
         Begin VB.TextBox TextCoverURL 
            Height          =   315
            Left            =   4320
            MaxLength       =   255
            TabIndex        =   13
            Top             =   4500
            Width           =   4035
         End
         Begin VB.ComboBox ComboNos 
            Height          =   315
            ItemData        =   "frmEditor.frx":975C
            Left            =   4140
            List            =   "frmEditor.frx":975E
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   120
            Width           =   4215
         End
         Begin VB.ComboBox TextUser 
            Height          =   315
            ItemData        =   "frmEditor.frx":9760
            Left            =   1260
            List            =   "frmEditor.frx":9762
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   11
            Top             =   5580
            Width           =   7095
         End
         Begin VB.TextBox CDSerialCur 
            Height          =   315
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   10
            Top             =   540
            Width           =   7095
         End
         Begin VB.TextBox TextVideoHid 
            Height          =   315
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   9
            Top             =   1620
            Width           =   7095
         End
         Begin VB.TextBox TextFilelenHid 
            Height          =   315
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   8
            Top             =   3960
            Width           =   7095
         End
         Begin VB.TextBox TextFPSHid 
            Height          =   315
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   7
            Top             =   2460
            Width           =   7095
         End
         Begin VB.TextBox TextAudioHid 
            Height          =   315
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   6
            Top             =   3000
            Width           =   7095
         End
         Begin VB.TextBox TextResolHid 
            Height          =   315
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   5
            Top             =   2040
            Width           =   7095
         End
         Begin VB.TextBox TextTimeHid 
            Height          =   315
            Left            =   1260
            TabIndex        =   4
            Top             =   1080
            Width           =   6615
         End
         Begin VB.TextBox TextFileName 
            Height          =   315
            Left            =   1260
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   3540
            Width           =   7095
         End
         Begin VB.TextBox TextCDN 
            Height          =   315
            Left            =   1260
            MaxLength       =   255
            TabIndex        =   2
            Top             =   120
            Width           =   2835
         End
         Begin SurVideoCatalog.XpB ComInterGoHid 
            Height          =   375
            Index           =   1
            Left            =   8460
            TabIndex        =   16
            Top             =   4920
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            Caption         =   ""
            ButtonStyle     =   3
            Picture         =   "frmEditor.frx":9764
            PictureWidth    =   16
            PictureHeight   =   16
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
            MaskColor       =   16711935
         End
         Begin SurVideoCatalog.XpB ComGetCover 
            Height          =   375
            Left            =   8460
            TabIndex        =   17
            Top             =   4470
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            Caption         =   "Get"
            ButtonStyle     =   3
            Picture         =   "frmEditor.frx":9AB8
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
            TabIndex        =   18
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
            Caption         =   "MovieURL"
            Height          =   255
            Index           =   10
            Left            =   60
            TabIndex        =   30
            Top             =   5040
            Width           =   1275
         End
         Begin VB.Label LTech 
            Caption         =   "CoverURL"
            Height          =   255
            Index           =   9
            Left            =   60
            TabIndex        =   29
            Top             =   4560
            Width           =   1275
         End
         Begin VB.Label LTech 
            Caption         =   "Debtor"
            Height          =   255
            Index           =   11
            Left            =   60
            TabIndex        =   28
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label LTech 
            Caption         =   "SN"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   27
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label LTech 
            Caption         =   "Video"
            Height          =   255
            Index           =   3
            Left            =   60
            TabIndex        =   26
            Top             =   1680
            Width           =   1275
         End
         Begin VB.Label LTech 
            Caption         =   "Size"
            Height          =   255
            Index           =   8
            Left            =   60
            TabIndex        =   25
            Top             =   4020
            Width           =   1275
         End
         Begin VB.Label LTech 
            Caption         =   "FPS"
            Height          =   255
            Index           =   5
            Left            =   60
            TabIndex        =   24
            Top             =   2520
            Width           =   1275
         End
         Begin VB.Label LTech 
            Caption         =   "Audio"
            Height          =   255
            Index           =   6
            Left            =   60
            TabIndex        =   23
            Top             =   3060
            Width           =   1275
         End
         Begin VB.Label LTech 
            Caption         =   "Resol"
            Height          =   255
            Index           =   4
            Left            =   60
            TabIndex        =   22
            Top             =   2100
            Width           =   1275
         End
         Begin VB.Label LTech 
            Caption         =   "Length"
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   21
            Top             =   1140
            Width           =   1275
         End
         Begin VB.Label LTech 
            Caption         =   "NNCD"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label LTech 
            Caption         =   "File"
            Height          =   255
            Index           =   7
            Left            =   60
            TabIndex        =   19
            Top             =   3600
            Width           =   1215
         End
      End
      Begin MSComctlLib.TabStrip TabStrAdEd 
         Height          =   8235
         Left            =   60
         TabIndex        =   120
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   14526
         TabWidthStyle   =   2
         TabFixedWidth   =   5995
         HotTracking     =   -1  'True
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
      Begin VB.Frame FrAdEdTextHid 
         BorderStyle     =   0  'None
         Height          =   7755
         Left            =   120
         TabIndex        =   70
         Top             =   600
         Width           =   10455
         Begin VB.ComboBox ComboOther 
            Height          =   315
            ItemData        =   "frmEditor.frx":A4CA
            Left            =   5400
            List            =   "frmEditor.frx":A4CC
            Sorted          =   -1  'True
            TabIndex        =   94
            Top             =   3420
            Width           =   2535
         End
         Begin VB.CheckBox ChInFilFl 
            Alignment       =   1  'Right Justify
            Caption         =   "Empty"
            Height          =   255
            Left            =   3720
            TabIndex        =   93
            Top             =   3840
            Value           =   1  'Checked
            Width           =   4815
         End
         Begin VB.ComboBox TextRate 
            Height          =   315
            ItemData        =   "frmEditor.frx":A4CE
            Left            =   1260
            List            =   "frmEditor.frx":A4D0
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   92
            Top             =   3060
            Width           =   2895
         End
         Begin VB.ComboBox TextSubt 
            Height          =   315
            ItemData        =   "frmEditor.frx":A4D2
            Left            =   5400
            List            =   "frmEditor.frx":A4D4
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   91
            Top             =   3060
            Width           =   2955
         End
         Begin VB.ComboBox TextLang 
            Height          =   315
            ItemData        =   "frmEditor.frx":A4D6
            Left            =   5400
            List            =   "frmEditor.frx":A4D8
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   90
            Top             =   2700
            Width           =   2955
         End
         Begin VB.TextBox TextRole 
            Height          =   735
            Left            =   1260
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   89
            Top             =   1920
            Width           =   7095
         End
         Begin VB.ComboBox TextMName 
            Height          =   315
            ItemData        =   "frmEditor.frx":A4DA
            Left            =   1260
            List            =   "frmEditor.frx":A4DC
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   88
            Top             =   120
            Width           =   7095
         End
         Begin VB.ComboBox TextCountry 
            Height          =   315
            ItemData        =   "frmEditor.frx":A4DE
            Left            =   1260
            List            =   "frmEditor.frx":A4E0
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   87
            Top             =   1200
            Width           =   4095
         End
         Begin VB.ComboBox TextOther 
            Height          =   315
            ItemData        =   "frmEditor.frx":A4E2
            Left            =   1260
            List            =   "frmEditor.frx":A4E4
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   86
            Top             =   3420
            Width           =   4095
         End
         Begin VB.Frame FrInetInfo 
            Caption         =   "IFind"
            Height          =   3195
            Left            =   5040
            TabIndex        =   79
            Top             =   4140
            Width           =   5415
            Begin VB.CommandButton ComRHid 
               Caption         =   "G"
               Height          =   315
               Index           =   0
               Left            =   4980
               TabIndex        =   83
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox TxtIName 
               Height          =   315
               Left            =   1200
               TabIndex        =   82
               Top             =   240
               Width           =   3735
            End
            Begin VB.ListBox lbInetMovieList 
               Height          =   2010
               ItemData        =   "frmEditor.frx":A4E6
               Left            =   120
               List            =   "frmEditor.frx":A4E8
               TabIndex        =   81
               Top             =   1080
               Width           =   5175
            End
            Begin VB.ComboBox ComboInfoSites 
               Height          =   315
               ItemData        =   "frmEditor.frx":A4EA
               Left            =   105
               List            =   "frmEditor.frx":A4EC
               Sorted          =   -1  'True
               TabIndex        =   80
               Top             =   660
               Width           =   3615
            End
            Begin SurVideoCatalog.XpB ComInetFind 
               Height          =   315
               Left            =   3840
               TabIndex        =   84
               Top             =   660
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               Caption         =   "Find"
               ButtonStyle     =   3
               Picture         =   "frmEditor.frx":A4EE
               PictureWidth    =   16
               PictureHeight   =   16
               XPColor_Pressed =   15116940
               XPColor_Hover   =   4692449
            End
            Begin VB.Label LIName 
               Caption         =   "Title"
               Height          =   375
               Left            =   120
               TabIndex        =   85
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.ComboBox TextYear 
            Height          =   315
            ItemData        =   "frmEditor.frx":AF00
            Left            =   1260
            List            =   "frmEditor.frx":AF02
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   78
            Top             =   2700
            Width           =   2895
         End
         Begin VB.ComboBox TextAuthor 
            Height          =   315
            ItemData        =   "frmEditor.frx":AF04
            Left            =   1260
            List            =   "frmEditor.frx":AF06
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   77
            Top             =   1560
            Width           =   7095
         End
         Begin VB.ComboBox TextGenre 
            Height          =   315
            ItemData        =   "frmEditor.frx":AF08
            Left            =   1260
            List            =   "frmEditor.frx":AF0A
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   76
            Top             =   840
            Width           =   4095
         End
         Begin VB.ComboBox TextLabel 
            Height          =   315
            ItemData        =   "frmEditor.frx":AF0C
            Left            =   1260
            List            =   "frmEditor.frx":AF0E
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   75
            Top             =   480
            Width           =   7095
         End
         Begin VB.ComboBox ComboSites 
            Height          =   315
            Left            =   60
            TabIndex        =   74
            Top             =   7440
            Width           =   5475
         End
         Begin VB.ComboBox ComboCountry 
            Height          =   315
            ItemData        =   "frmEditor.frx":AF10
            Left            =   5400
            List            =   "frmEditor.frx":AF12
            Sorted          =   -1  'True
            TabIndex        =   73
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox TextAnnotation 
            Height          =   3135
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            Top             =   4200
            Width           =   4890
         End
         Begin VB.ComboBox ComboGenre 
            Height          =   315
            ItemData        =   "frmEditor.frx":AF14
            Left            =   5400
            List            =   "frmEditor.frx":AF16
            Sorted          =   -1  'True
            TabIndex        =   71
            Top             =   840
            Width           =   2535
         End
         Begin SurVideoCatalog.XpB ComClsEd 
            Height          =   315
            Left            =   1260
            TabIndex        =   95
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
            Height          =   375
            Left            =   6300
            TabIndex        =   96
            Top             =   7380
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   661
            Caption         =   "drag-and-drop"
            ButtonStyle     =   3
            Picture         =   "frmEditor.frx":AF18
            PictureWidth    =   16
            PictureHeight   =   16
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
         End
         Begin SurVideoCatalog.XpB ComInterGoHid 
            Height          =   375
            Index           =   0
            Left            =   5640
            TabIndex        =   97
            Top             =   7380
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   661
            Caption         =   ""
            ButtonStyle     =   3
            Picture         =   "frmEditor.frx":B4B2
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
            TabIndex        =   98
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
            TabIndex        =   99
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
            TabIndex        =   100
            Top             =   3420
            Width           =   315
            _ExtentX        =   265
            _ExtentY        =   265
            Caption         =   "+"
            ButtonStyle     =   3
            XPColor_Pressed =   15116940
            XPColor_Hover   =   4692449
         End
         Begin VB.Label LFilm 
            Alignment       =   1  'Right Justify
            Caption         =   "Subtitle"
            Height          =   255
            Index           =   11
            Left            =   4260
            TabIndex        =   112
            Top             =   3120
            Width           =   1035
         End
         Begin VB.Label LFilm 
            Alignment       =   1  'Right Justify
            Caption         =   "Language"
            Height          =   255
            Index           =   10
            Left            =   4260
            TabIndex        =   111
            Top             =   2760
            Width           =   1035
         End
         Begin VB.Label LFilm 
            Caption         =   "Rating"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   110
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label LFilm 
            Caption         =   "Comments"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   109
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label LFilm 
            Caption         =   "Year"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   108
            Top             =   2760
            Width           =   1155
         End
         Begin VB.Label LFilm 
            Caption         =   "Actors"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   107
            Top             =   1980
            Width           =   1215
         End
         Begin VB.Label LFilm 
            Caption         =   "Country"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   106
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label LFilm 
            Caption         =   "Director"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   105
            Top             =   1620
            Width           =   1215
         End
         Begin VB.Label LFilm 
            Caption         =   "Descr"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   104
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label LFilm 
            Caption         =   "Label"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   103
            Top             =   540
            Width           =   1155
         End
         Begin VB.Label LFilm 
            Caption         =   "Title"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   102
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label LFilm 
            Caption         =   "Genre"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   101
            Top             =   900
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private unloadEditorFlag As Boolean 'не активировать в коде выхода


Private Sub CDSerialCur_Change()
Mark2Save
End Sub

Private Sub ChInFilFl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static working As Boolean
If working Then Exit Sub
working = True
Select Case ChInFilFl.Value
Case Unchecked
    ChInFilFl.Value = Gray
Case Gray
    ChInFilFl.Value = Checked
Case Checked
    ChInFilFl.Value = Unchecked
End Select
working = False
ReleaseCapture
End Sub

Private Sub ComAdd_Click()
'Append new file
Dim tmpdrive As String
Dim tmpSerial As String

ToDebug "Add2New..."

FrameAddEdit.Enabled = False    'чтоб не нажимались батоны
aviName = pLoadDialog(ComAdd.Caption)
FrameAddEdit.Enabled = True
DoEvents

If LenB(aviName) = 0 Then
ToDebug "...Cancel"
Exit Sub
End If

ToDebug "Добавление:" & aviName

'Серийник
NewDiskAddFlag = False    'не новый носитель
tmpdrive = left$(LCase$(aviName), 3)
If DriveType(tmpdrive) = "CD-ROM" Then
    tmpSerial = Hex$(GetSerialNumber(tmpdrive))
    If tmpSerial <> "0" Then

        If MediaSN = tmpSerial Then
            IsSameCdFlag = True    'это тот же носитель
        Else
            IsSameCdFlag = False
            MediaSN = tmpSerial    'запомнить серийник cd
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
    If IsSameCdFlag Then    'если тот же носитель
        ComboNos.Text = MediaType
    Else
        ComboNos.Text = ComboNos.Text & ", " & GetOptoInfo(left$(aviName, 2))    'добавить тип
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
    Sleep 300
    ComCap_Click 0
    Screen.MousePointer = vbHourglass

    PicSS1.Refresh: DoEvents
    'Sleep 500
    Position.Value = pos2
    objPosition.CurrentPosition = Position.Value / 100
    Sleep 300
    ComCap_Click 1
    Screen.MousePointer = vbHourglass

    PicSS2.Refresh: DoEvents
    'Sleep 500
    Position.Value = pos3
    objPosition.CurrentPosition = Position.Value / 100
    Sleep 300
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
    myMsgBox msgsvc(17), vbInformation, , Me.hwnd
    Exit Sub
Case 2
End Select

temp = ShellExecute(GetDesktopWindow(), "open", site, vbNull, vbNull, 1)
ToDebug "InetGo_ret=" & temp
End Sub




Public Sub ComPlusHid_Click(Index As Integer)
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
    If right$(TextTimeHid.Text, 2) = ", " Then TextTimeHid.Text = left$(TextTimeHid.Text, Len(TextTimeHid.Text) - 2)

    If Len(TextTimeHid.Text) <> 0 Then
        arr = Split(TextTimeHid.Text, ",")
        If UBound(arr) > 0 Then
            sTimeSum = TextTimeHid.Text
            For i = 0 To UBound(arr)
                tmpL1 = Time2sec(arr(i))
                If tmpL1 > -1 Then tmpL = tmpL + tmpL1
            Next i
            If tmpL > 0 Then TextTimeHid.Text = FormatTime(tmpL)
        Else
            'вернуть
            If Len(sTimeSum) <> 0 Then
                'sTimeSum
                TextTimeHid.Text = sTimeSum
                sTimeSum = vbNullString
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

'Select Case Index
'Case 1 'поиск названия фильма
    tmp = Replace(TxtIName.Text, "(", vbNullString)
    tmp = Replace(tmp, ")", vbNullString)
    tmp = Replace(tmp, "/", vbNullString)
    tmp = Replace(tmp, ".", vbNullString)
    tmp = Replace(tmp, "[", vbNullString)
    tmp = Replace(tmp, "]", vbNullString)

    site = "http://www.google.com/search?q=" & tmp
'End Select

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

Private Sub Form_Activate()
'если из кода выхода, то не активить - а все равно глюк
'If Not unloadEditorFlag Then frmEditorFlag = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next    'надо для аудио

Select Case KeyCode
Case 112 'F1
    If FrmMain.ChBTT.Value = 0 Then FrmMain.ChBTT.Value = 1 Else FrmMain.ChBTT.Value = 0
    Me.SetFocus
    
Case 45 'Insert auto add
    If Shift = 0 Then frmAuto.Show 1, Me

Case 27 'esc
    Unload Me
    
Case 123 'f12
    FormDebug.Show , FrmMain
    
Case 80   'P play DXShow
    If MpegMediaOpen Then
        If TabStrAdEd.SelectedItem.Index = 1 Then
            If MediaState = 2 Then    'играет, на паузу
                mobjManager.Pause
                tPlay.Enabled = False
'                'позиционировать слайдеры'                If isMPGflag Or isDShflag Then MPGPosScroll: lastRendedMPG = Position.Value
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

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 15 '^O
        ComOpen_Click
Case 19 '^S
        SaveFromEditor
Case 14 'новый
    Call FrmMain.VerticalMenu_MenuItemClick(3, False)
End Select
End Sub

Private Sub Form_Load()
Me.top = (Screen.Height - Height) / 2
Me.left = (Screen.Width - Width) / 2

'нет тултипов слайдеров
Const TBM_SETTOOLTIPS = &H41D
SendMessage Position.hwnd, TBM_SETTOOLTIPS, 0, 0
SendMessage PositionP.hwnd, TBM_SETTOOLTIPS, 0, 0

'Список скриптов
FillTemplateCombo App.Path & "\Scripts\", ComboInfoSites



Call ChangeComboHeights

'хук - no auto CD
If Not DebugMode Then
    If Opt_QueryCancelAutoPlay Then
       Call HookWindowEditor(Me.hwnd, Me) '+ Friend Function WindowProc
    End If
End If

'нет frmEditorFlag = True
End Sub
Friend Function WindowProc(hwnd As Long, Msg As Long, wp As Long, lp As Long) As Long
Dim result As Long
Select Case Msg
Case m_RegMsg       ' QueryCancelAutoPlay
    ' TRUE: cancel AutoRun
    ' *must* be 1, not -1!
    ' FALSE: allow AutoRun
    result = 1
ToDebug "CancelAutoPlayEditor"
Case Else
    ' Pass along to default window procedure.
    result = InvokeWindowProcEditor(hwnd, Msg, wp, lp)
End Select
' Return desired result code to Windows.
WindowProc = result
End Function
Private Sub ChangeComboHeights()
'высота кобмиков
'паразитно подставляет подходящие значения из списка вместо взятых из базы.
'не вызывать при ресайзе

Dim X As Long, Y As Long, w As Long
'SWP_NOZORDER Or SWP_NOMOVE Or SWP_DRAWFRAME
Const CB_SETDROPPEDWIDTH = &H160

X = ScaleX(ComboGenre.left, vbTwips, vbPixels)
Y = ScaleY(ComboGenre.top, vbTwips, vbPixels)
w = ScaleY(ComboGenre.Width, vbTwips, vbPixels)
SetWindowPos ComboGenre.hwnd, 0, X, Y, w, 500, 0

X = ScaleX(ComboCountry.left, vbTwips, vbPixels)
Y = ScaleY(ComboCountry.top, vbTwips, vbPixels)
w = ScaleY(ComboCountry.Width, vbTwips, vbPixels)
SetWindowPos ComboCountry.hwnd, 0, X, Y, w, 500, 0

'список аля ComboSites в тех вкладке
X = ScaleX(cBasePicURL.left, vbTwips, vbPixels)
Y = ScaleY(cBasePicURL.top, vbTwips, vbPixels)
w = ScaleY(cBasePicURL.Width, vbTwips, vbPixels)
SetWindowPos cBasePicURL.hwnd, 0, X, Y, w, 500, 0
'ширина выпадающего списка
Call SendMessage(cBasePicURL.hwnd, CB_SETDROPPEDWIDTH, 400, ByVal 0&)

'список сайтов в иннфо
X = ScaleX(ComboSites.left, vbTwips, vbPixels)
Y = ScaleY(ComboSites.top, vbTwips, vbPixels)
w = ScaleY(ComboSites.Width, vbTwips, vbPixels)
SetWindowPos ComboSites.hwnd, 0, X, Y, w, 500, 0

'список скриптов
X = ScaleX(ComboInfoSites.left, vbTwips, vbPixels)
Y = ScaleY(ComboInfoSites.top, vbTwips, vbPixels)
w = ScaleY(ComboInfoSites.Width, vbTwips, vbPixels)
SetWindowPos ComboInfoSites.hwnd, 0, X, Y, w, 500, 0

'примечания, качество
X = ScaleX(ComboOther.left, vbTwips, vbPixels)
Y = ScaleY(ComboOther.top, vbTwips, vbPixels)
w = ScaleY(ComboOther.Width, vbTwips, vbPixels)
SetWindowPos ComboOther.hwnd, 0, X, Y, w, 500, 0

''не пустые
X = ScaleX(ComboNos.left, vbTwips, vbPixels)
Y = ScaleY(ComboNos.top, vbTwips, vbPixels)
w = ScaleY(ComboNos.Width, vbTwips, vbPixels)
SetWindowPos ComboNos.hwnd, 0, X, Y, w, 500, 0

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
'ComboInfoSites.SelLength = 0
'ComboGenre.SelLength = 0
'ComboCountry.SelLength = 0
'ComboOther.SelLength = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'Dim ret As Long
'unloadEditorFlag = True
frmEditorFlag = False

If Not ExitSVC Then

    'в квери? If rs Is Nothing Then
    'Else
    '    'если редактируется База фильмов, спросить, что делать
    '    If rs.EditMode Then
    '        ret = myMsgBox(msgsvc(6), vbYesNoCancel, , Me.hwnd)
    '        If ret = vbYes Then
    '            ToDebug "- сохранить"
    '            ComSaveRec_Click
    '        ElseIf ret = vbNo Then
    '            ToDebug "- не сохранять"
    '            rs.CancelUpdate
    '        Else    'cancel
    '            Exit Sub
    '        End If    'ret
    '    End If    'rs.EditMode
    'End If    'rs nodb:


    'FrmMain.VerticalMenu_MenuItemClick 1, 0

    Cancel = True
    Me.Hide
    
    LastVMI = 1 'типа в просмотре
    FrmMain.ListView.MultiSelect = True 'вырубали для нового
    If rs.EditMode Then rs.CancelUpdate
    addflag = False: editFlag = False
    'вырубали для редактора
    FrmMain.PicSplitLVDHid.Enabled = True
    If frmOptFlag Then FrmOptions.FrameBD.Enabled = True
    
    POST_flag = False 'отменить пост, еще раз на всякий случай
    
    MakeNormal FrmMain.hwnd 'это выведет окно в фокус
End If
'unloadEditorFlag = False
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
    ToDebug "UserAspect-4:3"
Case 1    '16:9
    MMI_Format_str = "16/9"
    MMI_Format = 1.777
    ToDebug "UserAspect-16:9"
Case 2    'w:h
    'MMI_Format_str = "1/1"
    MMI_Format = 1.333 'если ошибка...
    MMI_Format = objVideo.SourceWidth / objVideo.SourceHeight
    ToDebug "UserAspect-w:h"
    If err Then ToDebug "Err_VideoObj"

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

ShowInShowPic 1, frmEditor

FormShowPic.Visible = True

End Sub



Private Sub TextAudioHid_Change()
Mark2Save
End Sub

Private Sub TextCountry_Change()
Mark2Save
If Len(TextCountry.Text) > 255 Then TextCountry.Text = left$(TextCountry.Text, 255)
'SendMessage TextCountry.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0
End Sub

Private Sub ComboInfoSites_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub ComCancel_Click()
OpenAddmovFlag = False: addflag = False
If rs.EditMode Then rs.CancelUpdate
NoPicFrontFaceFlag = False: NoPic1Flag = False: NoPic2Flag = False: NoPic3Flag = False

frmEditorFlag = False
Unload Me
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

ToDebug "ScrShotsPosAvi: " & pos1 & " " & pos2 & " " & pos3

ex:
If err.Number <> 0 Then ToDebug "ComCap_AVI: " & err.Description

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

DoEvents

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
    
    'Debug.Print Position.Value, lastRendedMPG
    
    
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
'If err.Number <> 0 Then Debug.Print err.Description

If Not frmAutoFlag Then Me.SetFocus
Screen.MousePointer = vbNormal

End Sub



Private Sub ComCap_Click(Index As Integer)
Dim tmp As String
    
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
    FrameAddEdit.Enabled = False    'чтоб не нажимались батоны
    tmp = pLoadDialog
    FrameAddEdit.Enabled = True
    DoEvents: DoEvents
    
    ToDebug "Open for capture"
    If Len(tmp) <> 0 Then OpenMovieForCapture (tmp)
End If

End Sub

Private Sub ComDel_Click()
DelFromEditor
FrmMain.LVCLICK
End Sub


Private Sub ComFrontFaceFile_Click(Index As Integer)
Dim iFile As String
Dim temp As Single
Dim TifPngFlag As Boolean

iFile = pLoadPixDialog
If LCase$(getExtFromFile(iFile)) = "png" Then TifPngFlag = True
If left$(LCase$(getExtFromFile(iFile)), 3) = "tif" Then TifPngFlag = True

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
'If Len(SearchString) = 0 Then
'    'myMsgBox msgsvc(29)
''   Exit Sub
'End If

'Set ImgPrCov = Nothing

'прочитать скрипт
iFile = FreeFile
'ComboInfoSites.Text
ToDebug "Скрипт: " & ComboInfoSites
Open App.Path & "\Scripts\" & ComboInfoSites.Text For Binary As #iFile
's = Space$(LOF(iFile))
s = AllocString_ADV(LOF(iFile))
'Debug.Print Len(s)

Get #iFile, , s
Close #iFile

'задать инстанс
Set SC = Nothing
Set SC = CreateObject("ScriptControl")
SC.language = "VBScript"
SC.Timeout = 20000
SC.AddCode s
'UI окно
'SC.SitehWnd = TextUI.hWnd
'SC.AllowUI = True 'False
'всунуть класс
SC.AddObject "SVC", objScript

'заменить нек символы кодами
SearchString = Replace(SearchString, "&", "%26") 'http://www.kinopoisk.ru/index.php?kp_query=Love+%26+Sex

ToDebug "Инет поиск: " & SearchString
ComInetFind.Enabled = False
lbInetMovieList.Clear

'сгенерить реферер
sReferer = GetReferer(SC.CodeObject.BaseAddress)
'Запросить скрипт о url = "http://www.videoguide.ru/find.asp?Search=Simple&types=film&titles="
url = SC.CodeObject.url & SearchString
'url = SC.CodeObject.url & UrlEncode(SearchString)
'Запросить скрипт о GET POST
POST_flag = CBool(SC.CodeObject.PostMethod)

ComboSites.Text = url
'url = "file://" & App.Path & "\Scripts\inet_dvdempire1.htm"

'ToDebug "запрос: " & url
PageText = OpenURLProxy(url, "txt")
'по строкам
PageArray() = Split(PageText, vbLf)

''Egg
'If frmPeopleFlag Then
'    FrmPeople.List1.Visible = False
'    FrmPeople.List1.Clear
'    For i = LBound(PageArray) To UBound(PageArray)
'        FrmPeople.List1.AddItem i & " |" & PageArray(i)
'    Next i
'    FrmPeople.SetListboxScrollbar FrmPeople.List1
'    FrmPeople.List1.Visible = True
'End If

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
    ToDebug " " & SC.CodeObject.MTitles(0)
Else
    ToDebug "список: " & 0
End If

ComInetFind.FontBold = False
ComInetFind.Enabled = True
Erase PageArray
'нужен далее Set SC = Nothing
POST_flag = False 'отменить пост

Exit Sub


ErrorHandler:
POST_flag = False 'отменить пост
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

Private Sub ComKeyAvi_Click(Index As Integer)
Select Case Index
Case 0
KeyNext
Case 1
KeyPrev
End Select

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
Dim tmp As String
tmp = GetFileNameFromEditor 'попробовать найти в поле редактора
If Len(tmp) = 0 Then
    OpenNewMovie
Else
    OpenNewMovie tmp
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

Private Sub ComSaveRec_Click()
SaveFromEditor
'флаги редакт и добавл = false
ToDebug "&"
'не FrmMain.VerticalMenu_MenuItemClick 1, 0
'лучше встать на созданную иначе в SaveFromEditor 'FrmMain.LVCLICK


'в сейве If editFlag Then FrmMain.LVCLICK
End Sub

Private Sub ComShowBin_Click()
FrmBin.Show
'FrmMainState = FrmMain.WindowState
'FrmMain.WindowState = vbMinimized
End Sub


Private Sub ComX_Click(Index As Integer)
Dim i As Integer
Select Case Index
Case 0
    If PicSS1.Picture = 0 Then Exit Sub
Case 1
    If PicSS2.Picture = 0 Then Exit Sub
Case 2
    If PicSS3.Picture = 0 Then Exit Sub
Case 3, 4
    If PicFrontFace.Picture = 0 Then Exit Sub

Case 5    'очистка видео

    'аля химчистка
    EditorNoVideoClear    'из в.меню тоже вызывается

    '    Position.Value = 0: PositionP.Value = 0
    '    Position.Enabled = True: PositionP.Enabled = True
    '    ComKeyAvi(0).Enabled = False: ComKeyAvi(1).Enabled = False
    '    ComAutoScrShots.Enabled = False
    '    For i = 0 To 2
    '        optAspect(i).Enabled = False
    '        optAspect(i).Value = vbUnchecked
    '        ComRND(i).Enabled = False
    '    Next i

    If m_cAVI Is Nothing Then
    Else
        m_cAVI.filename = vbNullString    'unload
        Set m_cAVI = Nothing
    End If
    If MpegMediaOpen Then Call MpegMediaClose    'очистка экрана видео
    sTimeSum = vbNullString
    isMPGflag = False: isAVIflag = False: isDShflag = False    'или ClearVideo()?
    Set movie = Nothing
    movie.Width = MovieEd_W: movie.Height = MovieEd_H    ': movie.Visible = True
    Exit Sub

End Select

If myMsgBox(msgsvc(4), vbOKCancel, , Me.hwnd) = vbCancel Then Exit Sub    'удалять?

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


Private Sub lbInetMovieList_Click()
Dim temp As String
Dim i As Integer
'клик на списке найденных фильмов
On Error Resume Next

If SC Is Nothing Then Exit Sub
AutoFillStore    'запомнить поля в списки для последующего авто ввода


With SC.CodeObject
    'ClearTextFields 'не надо
    'If UBound(.MTitlesURL) = LBound(.MTitlesURL) Then Exit Sub  а один?
    temp = .MTitlesURL(lbInetMovieList.ListIndex)    '+ 1)
    
    'POST_flag = CBool(.PostMethod)
    POST_flag = False
    
End With

If Len(temp) = 0 Then Exit Sub
'Debug.Print temp, URLTitleArr(lbInetMovieList.ListIndex + 1)
ComboSites.Text = temp    'url текущей в список урл-ов

'DoEvents
'ToDebug "поиск инфо на " & temp

PageText = OpenURLProxy(temp, "txt")

'Debug.Print PageText

'Open "g:\t" For Output As #2
'Print #2, PageText ' Input(LOF(1), 1)
'Close #2
PageArray() = Split(PageText, vbLf)

'Egg
'If frmPeopleFlag Then
'    FrmPeople.List1.Clear
'    FrmPeople.List1.Visible = False
'    For i = LBound(PageArray) To UBound(PageArray)
'        FrmPeople.List1.AddItem i & " |" & PageArray(i)
'    Next i
'    FrmPeople.SetListboxScrollbar FrmPeople.List1
'    FrmPeople.List1.Visible = True
'End If


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
SetFromScript    'todebug там

AutoFillStore    'запомнить поля после
Erase PageArray
POST_flag = False 'отменить пост

Exit Sub

ErrorHandler:
POST_flag = False 'отменить пост
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

Private Sub SetListboxScrollbar(lB As ListBox) ', frm As Form)
'добавление горизонтальной прокрутки listbox
Dim i As Integer
Dim new_len As Long
Dim max_len As Long

'у frmmain теперь TextWidth зависит от фонта
For i = 0 To lB.ListCount - 1
 new_len = 10 + ScaleX(frmEditor.TextWidth(lB.List(i)), frmEditor.ScaleMode, vbPixels)
 If max_len < new_len Then max_len = new_len
Next i
SendMessage lB.hwnd, LB_SETHORIZONTALEXTENT, max_len, 0
End Sub

Private Sub movie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'не задействуется, если не ави
If m_cAVI Is Nothing Then Exit Sub    'на всякий

If OpenAddmovFlag Then
    If Button = 2 Then
        FrmMain.PopupMenu FrmMain.popMovieHid

    Else
        With FrmMain
            Set .PicTempHid(0) = Nothing: Set .PicTempHid(1) = Nothing
            .PicTempHid(0).Width = ScaleX(AviWidth, vbPixels, vbTwips)
            .PicTempHid(0).Height = ScaleY(AviHeight, vbPixels, vbTwips)

            If Position.Value = 0 Then Position.Value = 1
            m_cAVI.DrawFrame .PicTempHid(0).hdc, Position.Value  ', 0, 0, Transparent:=False
            ShowInShowPic 0, frmEditor
        End With
    End If    ' buttons

End If    'ActFlag нет видео нет попапа

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
        ShowInShowPic 0, frmEditor
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
            ToDebug "SShot1Goto: " & pos1
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
        ShowInShowPic 0, frmEditor
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
            ToDebug "SShot2Goto: " & pos2
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
        ShowInShowPic 0, frmEditor
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
            ToDebug "SShot3Goto: " & pos3
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
    If ChLockFSHid(0) Then pRenderFrame Position.Value
    lastRendedAVI = Position.Value
End If

If isMPGflag Or isDShflag Then
    If lastRendedMPG = Position.Value Then Exit Sub
    If ChLockFSHid(0) Then

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
If ChLockFSHid(1) Then pRenderFrame Position.Value
End If

If isMPGflag Or isDShflag Then 'MPGPosScroll
If lastRendedMPG = Position.Value Then Exit Sub
    If ChLockFSHid(1) Then
    Screen.MousePointer = vbHourglass
      objPosition.CurrentPosition = Position.Value / 100
        mobjManager.Pause
    Screen.MousePointer = vbNormal
'Debug.Print temp, objPosition.CurrentPosition
    End If
End If

End Sub

Private Sub TabStrAdEd_Click()
On Error Resume Next

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
    
    'включить DS видео окно
    If Not (objVideoW Is Nothing) Then objVideoW.Visible = True
    
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

Private Sub tPlay_Timer()
On Error Resume Next
Position.Value = objPosition.CurrentPosition * 100

If isAVIflag Then
'    PosScroll
    lastRendedAVI = Position.Value
ElseIf isMPGflag Or isDShflag Then
'    MPGPosScroll
    lastRendedMPG = Position.Value
    
'On Error Resume Next
'Dim temp As Long
'temp = Position.Value + cMPGRange
'If temp <= PPMin Then
'    PPMin = Position.Value - cMPGRange
'    If PPMin < 0 Then PPMin = 0
'    PositionP.min = PPMin
'    PPMax = Position.Value + cMPGRange
'    If PPMax > TimesX100 Then PPMax = TimesX100
'    PositionP.Max = PPMax
'Else
'    PPMax = Position.Value + cMPGRange
'    If PPMax > TimesX100 Then PPMax = TimesX100
'    PositionP.Max = PPMax
'    PPMin = Position.Value - cMPGRange
'    If PPMin < 0 Then PPMin = 0
'    PositionP.min = PPMin
'End If
'PositionP.Value = Position.Value  'PositionP.Min
'If Position.Value = Position.Max Then PositionP.Value = PositionP.Max



End If
End Sub

Private Sub TxtIName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then ComInetFind_Click
End Sub

Public Sub EditorNoVideoClear()
 Dim i As Integer
 
Position.Enabled = False: PositionP.Enabled = False
Position.Value = 0: PositionP.Value = 0
ComKeyAvi(0).Enabled = False: ComKeyAvi(1).Enabled = False
optAspect(0).Value = False: optAspect(1).Value = False

ComAutoScrShots.Enabled = False

    For i = 0 To 2
        optAspect(i).Enabled = False
        optAspect(i).Value = vbUnchecked
        ComRND(i).Enabled = False
    Next i
    
'        Set .ImgPrCov = Nothing
'        Set .PicSS1Big = Nothing: Set .PicSS2Big = Nothing: Set .PicSS3Big = Nothing
'        Set .PicFrontFace = Nothing: Set .picCanvas = Nothing

End Sub
