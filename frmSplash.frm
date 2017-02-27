VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2475
   ClientLeft      =   7305
   ClientTop       =   2100
   ClientWidth     =   5970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   2220
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Min             =   1e-4
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   239
      X2              =   494
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      Top             =   30
      Width           =   885
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sur Video Catalog"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   180
      TabIndex        =   0
      Top             =   1500
      Width           =   3750
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
On Error Resume Next
DoEvents
Load FormDebug
Load FrmMain

NoResizePlease = True
FrmMain.Move FrmMain.Left, FrmMain.Top, MeWidth, MeHeight 'נאח נוסאיח
NoResizePlease = False
'FrmMain.Width = MeWidth: FrmMain.Height = MeHeight

    FrmMain.Show 'הגא נוסאיח

End Sub


Private Sub Form_Load()
lblVersion.Caption = "V " & App.Major & "." & App.Minor & "." & App.Revision
lblProductName.Caption = "Sur Video Catalog"
    
Me.Top = (Screen.Height - Height) / 2
Me.Left = (Screen.Width - Width) / 2

'SendMessage PBar.hwnd, &H2001, 0, ByVal RGB(255, 255, 100) 'PBar Forecolor
'SendMessage PBar.hwnd, &H409, 0, ByVal RGB(50, 150, 0) 'PBar Backcolor
SendMessage PBar.hwnd, &H2001, 0, ByVal RGB(0, 0, 100) 'PBar Forecolor

SplashFlag = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
SplashFlag = False
End Sub

