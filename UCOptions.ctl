VERSION 5.00
Begin VB.UserControl UCOptions 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CheckBox CheckLoadLastBD 
      Appearance      =   0  'Flat
      Caption         =   "ToCatalog"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Value           =   1  'Checked
      Width           =   10080
   End
   Begin VB.CheckBox CheckSavBigPix 
      Appearance      =   0  'Flat
      Caption         =   "RealPix"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   300
      Value           =   1  'Checked
      Width           =   10095
   End
   Begin VB.CheckBox ChDSFilt 
      Appearance      =   0  'Flat
      Caption         =   "DirectShow"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   10035
   End
   Begin VB.CheckBox ChOnlyTitle 
      Appearance      =   0  'Flat
      Caption         =   "Titles"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   900
      Width           =   10035
   End
   Begin VB.CheckBox ChLVGrid 
      Appearance      =   0  'Flat
      Caption         =   "Grid"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Value           =   1  'Checked
      Width           =   10095
   End
End
Attribute VB_Name = "UCOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

