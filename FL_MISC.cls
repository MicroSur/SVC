VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptoMISC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************
'*** Flamed Miscellanious Class (private)
'***
'*** Functions for MSF <-> LBA conversion
'*** and bit shifting
'***
'*** for public MSF conversion functions
'*** use FL_MSF
'****************************************

Private lngPower2(31) As Long

'' by Microsoft
'Function UnsignedToLong(Value As Double) As Long
'    If Value < 0 Or Value >= 4294967296# Then Error 6 ' Overflow
'    If Value <= 2147483647 Then
'        UnsignedToLong = Value
'    Else
'        UnsignedToLong = Value - 4294967296#
'    End If
'End Function
'
'' by Microsoft
'Function LongToUnsigned(Value As Long) As Double
'    If Value < 0 Then
'        LongToUnsigned = Value + 4294967296#
'    Else
'        LongToUnsigned = Value
'  End If
'End Function

'is a bit in a byte set?
Public Function IsBitSet(ByVal InByte As Byte, ByVal Bit As Byte) As Boolean
IsBitSet = ((InByte And (2 ^ Bit)) > 0)
End Function

' strips MSF from format: "MM:SS:FF"
Public Function MSF2Parts(ByVal MSF As String) As cOptoMSF

Set MSF2Parts = New cOptoMSF

With MSF2Parts

    ' minutes
    .M = Val(Left$(MSF, InStr(MSF, ":") - 1))
    MSF = Mid$(MSF, 3 + 1)
    ' seconds
    .s = Val(Left$(MSF, InStr(MSF, ":") - 1))
    MSF = Mid$(MSF, 3 + 1)
    ' frames
    .F = Val(Left$(MSF, Len(MSF)))

End With
End Function

'MSF to LBA
Public Function MSF2LBA(ByVal mins As Long, ByVal secs As Long, _
        ByVal Frames As Long, _
        Optional positive As Boolean) As Long

MSF2LBA = CLng(mins) * 60 * 75 + (secs * 75) + Frames

If mins < 90 Or positive Then
    MSF2LBA = MSF2LBA - 150
Else
    MSF2LBA = MSF2LBA - 450150
End If

End Function

'LBA to MSF
Public Function LBA2MSF(ByVal LBA As Long) As String

Dim M As Long, s As Long, F As Long, start As Long

start = Choose(Abs(CBool(LBA >= -150)) + 1, 450150, 150)

M = Fix((LBA + start) / (60 * 75))
s = Fix((LBA + start - M * 60 * 75) / 75)
F = Fix(LBA + start - M * 60 * 75 - s * 75)

LBA2MSF = Format(M, "00") & ":" & Format(s, "00") & ":" & Format(F, "00")

End Function

'>> Operator for VB
'by VB-Accelerator
Public Function RShift(ByVal lThis As Long, ByVal lBits As Long) As Long

Static Init As Boolean

If Not Init Then InitShifting: Init = True

If (lBits <= 0) Then
    RShift = lThis
ElseIf (lBits > 63) Then
    Exit Function
ElseIf (lBits > 31) Then
    RShift = 0
Else
    If (lThis And lngPower2(31)) = lngPower2(31) Then
        RShift = (lThis And &H7FFFFFFF) \ lngPower2(lBits) Or lngPower2(31 - lBits)
    Else
        RShift = lThis \ lngPower2(lBits)
    End If
End If

End Function

'<< Operator for VB
'by VB-Accelerator
Public Function LShift(ByVal lThis As Long, ByVal lBits As Long) As Long

Static Init As Boolean

If Not Init Then InitShifting: Init = True

If (lBits <= 0) Then
    LShift = lThis
ElseIf (lBits > 63) Then
    Exit Function
ElseIf (lBits > 31) Then
    LShift = 0
Else
    If (lThis And lngPower2(31 - lBits)) = lngPower2(31 - lBits) Then
        LShift = (lThis And (lngPower2(31 - lBits) - 1)) * lngPower2(lBits) Or lngPower2(31)
    Else
        LShift = (lThis And (lngPower2(31 - lBits) - 1)) * lngPower2(lBits)
    End If
End If

End Function

'Shifting values
'by VB-Accelerator
Private Sub InitShifting()
lngPower2(0) = &H1&: lngPower2(1) = &H2&: lngPower2(2) = &H4&
lngPower2(3) = &H8&: lngPower2(4) = &H10&: lngPower2(5) = &H20&
lngPower2(6) = &H40&: lngPower2(7) = &H80&: lngPower2(8) = &H100&
lngPower2(9) = &H200&: lngPower2(10) = &H400&: lngPower2(11) = &H800&
lngPower2(12) = &H1000&: lngPower2(13) = &H2000&: lngPower2(14) = &H4000&
lngPower2(15) = &H8000&: lngPower2(16) = &H10000: lngPower2(17) = &H20000
lngPower2(18) = &H40000: lngPower2(19) = &H80000: lngPower2(20) = &H100000
lngPower2(21) = &H200000: lngPower2(22) = &H400000: lngPower2(23) = &H800000
lngPower2(24) = &H1000000: lngPower2(25) = &H2000000: lngPower2(26) = &H4000000
lngPower2(27) = &H8000000: lngPower2(28) = &H10000000: lngPower2(29) = &H20000000
lngPower2(30) = &H40000000: lngPower2(31) = &H80000000
End Sub
