VERSION 5.00
Begin VB.Form Form_test 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin SurVideoCatalog.UCLVaddon UCLVaddon1 
      Height          =   5475
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   9657
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
End
Attribute VB_Name = "Form_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
UCLVaddon1.BackColor = &HC0FFFF
'UCLVaddon1.ForeColor = vbWhite

UCLVaddon1.tAudio = "safdasd"

End Sub
