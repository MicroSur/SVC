VERSION 5.00
Begin VB.Form FormDebug 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "SVC_Debug"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TextDebug 
      Height          =   5595
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "FormDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Me.ScaleMode = 1 'vbTwips

End Sub

Private Sub Form_Resize()
TextDebug.Width = Me.Width - 100
TextDebug.Height = Me.Height - 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not ExitSVC Then
 Me.Hide
 Cancel = True
End If
End Sub
