VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptoMSF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************
'*** Flamed MSF class
'****************************************

Private lngMinutes As Long
Private lngSeconds As Long
Private lngFrames As Long

Private LBAConv As New cOptoMISC

Public Property Get MSF() As String
MSF = Format(M, "00") & ":" & _
        Format(s, "00") & ":" & _
        Format(F, "00")
End Property

Public Property Let MSF(aval As String)
With LBAConv.MSF2Parts(aval)
    M = .M
    s = .s
    F = .F
End With
End Property

Public Property Let LBA(ByVal LBA As Long)
With LBAConv.MSF2Parts(LBAConv.LBA2MSF(LBA))
    lngMinutes = .M
    lngSeconds = .s
    lngFrames = .F
End With
End Property

Public Property Get LBA() As Long
LBA = LBAConv.MSF2LBA(lngMinutes, _
        lngSeconds, _
        lngFrames, True)
End Property

Public Property Get M() As Long
M = lngMinutes
End Property

Public Property Let M(lng As Long)
lngMinutes = lng
End Property

Public Property Get s() As Long
s = lngSeconds
End Property

Public Property Let s(lng As Long)
lngSeconds = lng
End Property

Public Property Get F() As Long
F = lngFrames
End Property

Public Property Let F(lng As Long)
lngFrames = lng
End Property
